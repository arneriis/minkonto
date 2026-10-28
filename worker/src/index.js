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

<faq>
${faq}
</faq>`;
}

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

      return json({ reply }, 200, cors);
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
