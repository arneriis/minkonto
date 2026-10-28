// "Spør Arne" – Cloudflare Worker that keeps the Anthropic API key secret
// and answers questions from minkonto.riis.cc using the FAQ published on the site.
import Anthropic from "@anthropic-ai/sdk";

const LANG_NAMES = { nb: "norsk (bokmål)", sv: "svenska", da: "dansk", en: "English" };
const MAX_MESSAGES = 12;
const MAX_CHARS = 2000;
const FAQ_TTL_MS = 5 * 60 * 1000;

let faqCache = { text: "", fetchedAt: 0 };

async function loadFaq(env) {
  if (faqCache.text && Date.now() - faqCache.fetchedAt < FAQ_TTL_MS) return faqCache.text;
  const res = await fetch(env.FAQ_URL, { cf: { cacheTtl: 300 } });
  if (!res.ok) throw new Error(`FAQ fetch failed: ${res.status}`);
  faqCache = { text: await res.text(), fetchedAt: Date.now() };
  return faqCache.text;
}

function systemPrompt(faq, lang) {
  return `Du er «Arne», en vennlig IT-assistent for ansatte i Recover. Du svarer på spørsmål fra nettsiden minkonto.riis.cc, som oftest på mobil.

Regler:
- Svar kun ut fra FAQ-en og snarveiene nedenfor. Ikke finn på steg, lenker eller kontaktinformasjon.
- Finner du ikke svaret i FAQ-en, si det ærlig og be brukeren kontakte Iver support.
- Hold svarene korte og lette å lese på mobil: én kort innledning, deretter nummererte steg når det passer.
- Ta med relevant lenke fra FAQ-en som full URL.
- Svar på samme språk som brukeren skriver på. Er det uklart, svar på ${LANG_NAMES[lang] || LANG_NAMES.nb}. Navn på knapper i skjermbilder (som «Next» og «Finish») skal stå som de er.
- Be aldri om passord, koder eller annen sensitiv informasjon.

Bestilling av skriver:
- Når noen vil bestille skriver/printer, tilby å fylle ut bestillingen for dem («Skal jeg gjøre dette for deg?»).
- Samle inn det som trengs (se FAQ-en), noen få spørsmål om gangen. Start med modell og antall – hjelp brukeren å velge ut fra behov (A3 eller bare A4, volum, kortleser) og nevn månedsprisen.
- Ikke gjett verdier. Er noe uklart, spør. Fakturametode er PDF eller EHF; e-post for faktura trengs bare ved PDF.
- Når alt er samlet inn, vis en kort oppsummering og spør om den stemmer. Kall verktøyet lag_printerbestilling først når brukeren har bekreftet.
- Etter verktøykallet: si kort at e-posten er klar, og at brukeren må trykke på knappen og sende den fra sin egen e-post.

<faq>
${faq}
</faq>`;
}

const nullableString = { type: ["string", "null"] };

const PRINTER_ORDER_TOOL = {
  name: "lag_printerbestilling",
  description:
    "Lager en ferdig utfylt printerbestilling (Canon-bestillingsskjema for Recover) som brukeren kan sende på e-post til Iver servicedesk. Kall kun når all informasjon er samlet inn og brukeren har bekreftet oppsummeringen.",
  strict: true,
  input_schema: {
    type: "object",
    additionalProperties: false,
    required: [
      "produkter", "firmanavn", "orgnr", "fakturametode", "faktura_epost", "fakturareferanse",
      "adresse", "postnummer", "poststed", "etasje", "kontaktperson", "kontakt_mobil_epost",
      "bestiller_navn", "sted", "kommentar",
    ],
    properties: {
      produkter: {
        type: "array",
        description: "Modellene som bestilles, med antall.",
        items: {
          type: "object",
          additionalProperties: false,
          required: ["modell", "antall"],
          properties: {
            modell: { type: "string", enum: ["bordmodell", "mfp_liten_a4", "mfp_medium_a4a3", "mfp_stor_a4a3"] },
            antall: { type: "integer" },
          },
        },
      },
      firmanavn: { type: "string", description: "Juridisk enhet som tiltrer avtalen." },
      orgnr: { type: "string" },
      fakturametode: { type: "string", enum: ["PDF", "EHF"] },
      faktura_epost: { ...nullableString, description: "Kun ved PDF-faktura, ellers null." },
      fakturareferanse: { type: "string" },
      adresse: { type: "string" },
      postnummer: { type: "string" },
      poststed: { type: "string" },
      etasje: nullableString,
      kontaktperson: { type: "string" },
      kontakt_mobil_epost: { type: "string" },
      bestiller_navn: { type: "string", description: "Navnet til den som bestiller, med blokkbokstaver." },
      sted: { type: "string", description: "Sted for underskrift." },
      kommentar: nullableString,
    },
  },
};

function corsHeaders(origin, env) {
  const allowed = env.ALLOWED_ORIGINS.split(",").map((s) => s.trim());
  return {
    "Access-Control-Allow-Origin": allowed.includes(origin) ? origin : allowed[0],
    "Access-Control-Allow-Methods": "POST, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type",
    "Vary": "Origin",
  };
}

function json(body, status, headers) {
  return new Response(JSON.stringify(body), {
    status,
    headers: { ...headers, "Content-Type": "application/json; charset=utf-8" },
  });
}

function cleanMessages(input) {
  if (!Array.isArray(input)) return null;
  const messages = input
    .filter((m) => m && (m.role === "user" || m.role === "assistant") && typeof m.content === "string")
    .map((m) => ({ role: m.role, content: m.content.slice(0, MAX_CHARS) }))
    .slice(-MAX_MESSAGES);
  while (messages.length && messages[0].role !== "user") messages.shift();
  if (!messages.length || messages[messages.length - 1].role !== "user") return null;
  return messages;
}

export default {
  async fetch(request, env) {
    const origin = request.headers.get("Origin") || "";
    const cors = corsHeaders(origin, env);

    if (request.method === "OPTIONS") return new Response(null, { status: 204, headers: cors });
    if (request.method !== "POST") return json({ error: "Method not allowed" }, 405, cors);

    const allowed = env.ALLOWED_ORIGINS.split(",").map((s) => s.trim());
    if (!allowed.includes(origin)) return json({ error: "Forbidden" }, 403, cors);

    let body;
    try {
      body = await request.json();
    } catch {
      return json({ error: "Invalid JSON" }, 400, cors);
    }

    const messages = cleanMessages(body.messages);
    if (!messages) return json({ error: "Invalid messages" }, 400, cors);
    const lang = LANG_NAMES[body.lang] ? body.lang : "nb";

    try {
      const faq = await loadFaq(env);
      const client = new Anthropic({ apiKey: env.ANTHROPIC_API_KEY });
      const response = await client.beta.messages.create({
        model: "claude-opus-5-5",
        // Short chat answers on a public endpoint – cap output to keep cost bounded.
        max_tokens: 2000,
        output_config: { effort: "low" },
        betas: ["server-side-fallback-2026-07-01"],
        fallbacks: "default",
        system: systemPrompt(faq, lang),
        tools: [PRINTER_ORDER_TOOL],
        messages,
      });

      if (response.stop_reason === "refusal") {
        return json({ reply: null, error: "refusal" }, 200, cors);
      }

      const reply = response.content
        .filter((block) => block.type === "text")
        .map((block) => block.text)
        .join("\n")
        .trim();

      // The page turns the order into a ready-to-send e-mail; no tool_result round trip is needed.
      const toolUse = response.content.find(
        (block) => block.type === "tool_use" && block.name === PRINTER_ORDER_TOOL.name,
      );

      return json({ reply, order: toolUse ? toolUse.input : null }, 200, cors);
    } catch (err) {
      if (err instanceof Anthropic.RateLimitError) {
        return json({ error: "busy" }, 429, cors);
      }
      if (err instanceof Anthropic.APIError) {
        console.error("Anthropic API error", err.status, err.message);
        return json({ error: "upstream" }, 502, cors);
      }
      console.error("Worker error", err);
      return json({ error: "internal" }, 500, cors);
    }
  },
};
