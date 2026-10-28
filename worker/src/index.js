// "Spør Arne" – Cloudflare Worker that keeps the Anthropic API key secret
// and answers questions from minkonto.riis.cc using Recover's internal FAQ.
import Anthropic from "@anthropic-ai/sdk";
import { createRemoteJWKSet, jwtVerify } from "jose";
// Internal knowledge – bundled into the Worker, never published (knowledge/ is git-ignored).
import FAQ from "../knowledge/faq.md";

const LANG_NAMES = { nb: "norsk (bokmål)", sv: "svenska", da: "dansk", en: "English" };
const MAX_MESSAGES = 12;
const MAX_CHARS = 2000;

let jwks = null;

// Sign-in with Recover's Microsoft Entra ID. Active once AZURE_TENANT_ID and AZURE_CLIENT_ID are set.
function authEnabled(env) {
  return Boolean(env.AZURE_TENANT_ID && env.AZURE_CLIENT_ID);
}

async function verifyUser(request, env) {
  const header = request.headers.get("Authorization") || "";
  const token = header.startsWith("Bearer ") ? header.slice(7) : "";
  if (!token) return null;
  jwks ??= createRemoteJWKSet(
    new URL(`https://login.microsoftonline.com/${env.AZURE_TENANT_ID}/discovery/v2.0/keys`),
  );
  try {
    const { payload } = await jwtVerify(token, jwks, {
      issuer: `https://login.microsoftonline.com/${env.AZURE_TENANT_ID}/v2.0`,
      audience: env.AZURE_CLIENT_ID,
    });
    return payload;
  } catch {
    return null;
  }
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

Reise og utlegg (Expense fra SD Worx – Norge):
- Veiledningene for Expense (Daily- og Manager-appen) gjelder ansatte i Norge. Er det uklart hvilket land brukeren jobber i, spør. Gjelder det ikke Norge, si at veiledning for andre land kommer senere, og henvis til Iver support.
- Svar med de konkrete stegene for det brukeren spør om (utlegg, reise på regning, diett innland/utland, kontering, godkjenning, innlogging/passord).

Tjenestebil, bilforsikring og bilordninger (Norge):
- Vil noen låne tjenestebilen privat (tilfeldig transport), forklar kort reglene og tilby skjemaet: ta med markøren [[form:laan-bil]] alene på egen linje. Skjemaet fyller de ut og signerer med fingeren, og det sendes fra deres egen e-post (før turen til leder, etter turen til bil og lønn).
- Alt om tjenestebil, skade, forsikring, lading, kjørebok, privat bil i tjeneste og bilrabatter gjelder Recover Norge. Er det uklart hvilket land brukeren jobber i, spør. Gjelder det ikke Norge, si at dette kommer senere og henvis til bil@recover.no eller Iver support.
- Ved skade: si hva brukeren skal gjøre nå (sikkerhet, veihjelp, skademelding, hvem som skal ha kopi), og skill mellom forsikringsskade og mindre skade under kr 23 000. Spør om det er person-/varebil eller ansattbil hvis egenandelen er relevant.

Videoer:
- Noen steg i FAQ-en har en videomarkør som [[video:utlegg-app]]. Når du svarer på noe som har en slik markør, ta med markøren nøyaktig slik den står, alene på egen linje, rett etter stegene den hører til. Ikke lag nye markører og ikke skriv lenker til videoene selv.

Bestilling av skriver (kun Norge):
- Printerbestillingen gjelder foreløpig bare Norge. Skal skriveren til et annet land, eller skriver brukeren på et annet språk enn norsk og det er uklart hvilket land det gjelder, spør om landet. Gjelder det ikke Norge, si at bestilling for andre land kommer senere, og henvis til Iver support.
- Når noen vil bestille skriver/printer, tilby å fylle ut bestillingen for dem («Skal jeg gjøre dette for deg?»).
- Samle inn det som trengs (se FAQ-en), noen få spørsmål om gangen. Start med modell og antall – hjelp brukeren å velge ut fra behov (A3 eller bare A4, volum, kortleser) og nevn månedsprisen.
- Ikke gjett verdier. Er noe uklart, spør. Ikke spør om fakturainformasjon – den håndteres av Iver.
- Når alt er samlet inn, vis en kort oppsummering og spør om den stemmer. Kall verktøyet lag_printerbestilling først når brukeren har bekreftet.
- Etter verktøykallet: si kort at e-posten er klar, og at brukeren må trykke på knappen og sende den fra sin egen e-post.

Overføre mobilnummer / overdragelse av abonnement (kun Norge):
- Gjelder kun norske mobilnumre. Spør om nummeret er norsk hvis det er uklart; er det ikke norsk, si at dette bare gjelder Norge og henvis til Iver support.
- Tilby å fylle ut overdragelsen for brukeren og lage e-posten til Mobility («Skal jeg gjøre dette for deg?»).
- Samle inn feltene i skjemaet (se FAQ-en), noen få spørsmål om gangen. Ikke gjett verdier.
- Spør IKKE om fødselsdato – verken brukerens eller tidligere eiers. Brukeren fyller den inn selv i e-posten. Er tidligere eier et firma, spør om org.nr.
- Når alt er samlet inn, vis en kort oppsummering og spør om den stemmer. Kall verktøyet lag_nummeroverforing først når brukeren har bekreftet.
- Etter verktøykallet: si kort at e-posten er klar, at brukeren må fylle inn fødselsdato der det står ________, og at signert overdragelsesskjema må legges ved hvis brukeren ikke selv eier abonnementet i dag.

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
      "produkter", "adresse", "postnummer", "poststed", "etasje", "kontaktperson", "kontakt_mobil_epost",
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

const nullableInt = { type: ["integer", "null"] };

const NUMBER_TRANSFER_TOOL = {
  name: "lag_nummeroverforing",
  description:
    "Lager en ferdig utfylt overdragelse av mobilabonnement fra tidligere eier til Recover (Mobility, kun Norge) som brukeren kan sende på e-post til ordre@mobility.no. Kall kun når all informasjon er samlet inn og brukeren har bekreftet oppsummeringen.",
  strict: true,
  input_schema: {
    type: "object",
    additionalProperties: false,
    required: [
      "abonnement", "tvillingsim_antall", "datasim_antall", "fri_tale_eu", "fri_tale_skandinavia",
      "simkort", "hovedsim_nummer", "navn", "mobilnummer", "ansattnummer", "fakturakonto",
      "tidligere_eier_type", "tidligere_eier_navn", "tidligere_eier_orgnr", "dato_gjennomforing", "kommentar",
    ],
    properties: {
      abonnement: { type: "string", description: "Abonnementet som ønskes, f.eks. «Bedrift Total A (15GB)»." },
      tvillingsim_antall: { type: "integer" },
      datasim_antall: { type: "integer", description: "Datasim trengs for bruk i nettbrett." },
      fri_tale_eu: { type: "boolean" },
      fri_tale_skandinavia: { type: "boolean" },
      simkort: {
        type: "string",
        enum: ["behold_eksisterende", "nytt_i_posten", "ny_esim", "har_nytt_simkort"],
        description: "har_nytt_simkort = brukeren har allerede et nytt SIM-kort og oppgir hovedsim-nummeret.",
      },
      hovedsim_nummer: { ...nullableString, description: "SIM-kortnummer (20 siffer) ved har_nytt_simkort, ellers null." },
      navn: { type: "string" },
      mobilnummer: { type: "string" },
      ansattnummer: { type: "string", description: "Feltet «Avd.nr/7000/Ansattnummer» – ett felt; bruk det nummeret brukeren oppgir." },
      fakturakonto: { type: "string", description: "Fakturakonto hos Telenor. «Standard» hvis brukeren ikke vet." },
      tidligere_eier_type: { type: "string", enum: ["meg_selv", "privatperson", "firma"] },
      tidligere_eier_navn: { ...nullableString, description: "Fullt navn eller firmanavn til tidligere eier. Null hvis meg_selv." },
      tidligere_eier_orgnr: { ...nullableString, description: "Org.nr (9 siffer) hvis tidligere eier er firma, ellers null." },
      dato_gjennomforing: { type: "string", description: "Ønsket dato for gjennomføring, format dd.mm.åååå." },
      kommentar: nullableString,
    },
  },
};

function corsHeaders(origin, env) {
  const allowed = env.ALLOWED_ORIGINS.split(",").map((s) => s.trim());
  return {
    "Access-Control-Allow-Origin": allowed.includes(origin) ? origin : allowed[0],
    "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type, Authorization",
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

    const allowed = env.ALLOWED_ORIGINS.split(",").map((s) => s.trim());
    if (!allowed.includes(origin)) return json({ error: "Forbidden" }, 403, cors);

    if (authEnabled(env) && !(await verifyUser(request, env))) {
      return json({ error: "unauthorized" }, 401, cors);
    }

    // The page's offline FAQ search, for signed-in users only.
    const path = new URL(request.url).pathname;
    if (request.method === "GET" && path === "/faq") {
      return new Response(FAQ, {
        headers: { ...cors, "Content-Type": "text/markdown; charset=utf-8", "Cache-Control": "private, no-store" },
      });
    }
    if (request.method !== "POST") return json({ error: "Method not allowed" }, 405, cors);

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
      const client = new Anthropic({ apiKey: env.ANTHROPIC_API_KEY });
      const response = await client.beta.messages.create({
        model: "claude-opus-5-5",
        // Short chat answers on a public endpoint – cap output to keep cost bounded.
        max_tokens: 2000,
        output_config: { effort: "low" },
        betas: ["server-side-fallback-2026-07-01"],
        fallbacks: "default",
        system: systemPrompt(FAQ, lang),
        tools: [PRINTER_ORDER_TOOL, NUMBER_TRANSFER_TOOL],
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
      const toolInput = (name) => {
        const block = response.content.find((b) => b.type === "tool_use" && b.name === name);
        return block ? block.input : null;
      };

      return json({
        reply,
        order: toolInput(PRINTER_ORDER_TOOL.name),
        transfer: toolInput(NUMBER_TRANSFER_TOOL.name),
      }, 200, cors);
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
