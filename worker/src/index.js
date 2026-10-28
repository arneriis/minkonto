// "Spør Arne" – Cloudflare Worker that keeps the Anthropic API key secret
// and answers questions from minkonto.riis.cc using Recover's internal FAQ.
import Anthropic from "@anthropic-ai/sdk";
import { createRemoteJWKSet, jwtVerify } from "jose";
// Internal knowledge – bundled into the Worker, never published (knowledge/ is git-ignored).
import FAQ from "../knowledge/faq.md";
// Price snapshot from Recover's Mobility shop (sku → price); refreshed in KV whenever Mobility shows a price.
import PRICE_SNAPSHOT from "../knowledge/prices.json";

const LANG_NAMES = { nb: "norsk (bokmål)", sv: "svenska", da: "dansk", en: "English" };
const MAX_MESSAGES = 12;
const MAX_CHARS = 4000;
const MAX_TOOL_ROUNDS = 5;

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
- Finner du ikke svaret i FAQ-en, si det ærlig og be brukeren kontakte Iver support – men gjelder det kjøp av utstyr (mobil, nettbrett, PC, skjermer, tilbehør, bestillingsstatus, levering, retur), er kontakten ordre@mobility.no, ikke Iver. Dette gjelder alle land.
- Hold svarene korte og lette å lese på mobil: én kort innledning, deretter nummererte steg når det passer.
- Ta med relevant lenke fra FAQ-en som full URL.
- Svar på samme språk som brukeren skriver på. Er det uklart, svar på ${LANG_NAMES[lang] || LANG_NAMES.nb}. Navn på knapper i skjermbilder (som «Next» og «Finish») skal stå som de er.
- Be aldri om passord, koder eller annen sensitiv informasjon.

Reise og utlegg (Expense fra SD Worx – Norge):
- Veiledningene for Expense (Daily- og Manager-appen) gjelder ansatte i Norge. Er det uklart hvilket land brukeren jobber i, spør. Gjelder det ikke Norge, si at veiledning for andre land kommer senere, og henvis til Iver support.
- Svar med de konkrete stegene for det brukeren spør om (utlegg, reise på regning, diett innland/utland, kontering, godkjenning, innlogging/passord).

Bestille mobil, nettbrett og tilbehør (Mobility):
- Spør noen hva som er standard eller anbefalt mobil, svar med «Anbefalt standardtelefon» fra FAQ-en (iPhone 16e 128GB og Google Pixel 10a 128GB, ca. kr 6 000, og Samsung Galaxy A57 som alternativ), og tilby å bestille. Vil de bestille, bruk sku-en fra FAQ-en og shop_id 24 i mobility_order (spør om mottaker først). Prisen kan ha endret seg – Mobility viser gjeldende pris.
- Spør noen hva som er mest kjøpt eller populært, svar ut fra tallene systemet gir deg (bestillinger gjort via Arne).
- Kjøp av utstyr i alle land: henvis aldri til Iver for kjøp. Kan noe ikke bestilles her i chatten, eller gjelder det status, levering, retur eller bytte av en bestilling, henvis til ordre@mobility.no.
- PC, skjermer, tastatur, mus, dokkingstasjon og annet PC-utstyr bestilles også med mobility_order (butikken «Standard PC & Utstyr»).
- Bestilling av mobil, nettbrett, tilbehør, smartklokke og SIM skal ALLTID gjøres med mobility_order her i chatten – aldri ved å henvise til Mitt Mobility-lenken, e-post, Iver eller et skjema. Er Mobility ikke koblet til, vis [[action:mobility-connect]].
- Når brukeren vil bestille mobil, nettbrett eller tilbehør, bruk mobility_order. Ledere som vil se eller godkjenne bestillinger bruker mobility_approve. Følg verktøyenes instruksjoner for rekkefølgen på spørsmålene.
- Når Mobility-svaret har «choice», vises valgene automatisk som knapper under svaret ditt. Still spørsmålet kort og ikke list opp alternativene selv.
- Brukerens knappetrykk kommer som tekst med «[valgt nøkkel=verdi]». Bruk verdien (for eksempel sku) i neste kall.
- Tidligere Mobility-steg står i «[Mobility: …]» i dine egne meldinger. Send alle verdier som allerede er valgt med på nytt i hvert kall.
- Send confirmation (JA/YES), eller decision approve/reject, BARE når brukerens siste melding er et trykk på bekreftelsesknappen. Systemet blokkerer det ellers.
- Bestillingen går til godkjenneren (som regel lederen) før Mobility behandler den. Spør om godkjenner når Mobility ber om det.
- Når brukeren har valgt produkt, spør først: er bestillingen til deg selv eller til en annen ansatt (da navn og mobilnummer)? Man kan bestille for andre ansatte – det avgjør Mobility ut fra brukerens tilgang.
- Kall så mobility_order med sku og mottaker. Mobility foreslår leveringsadressen der den ansatte jobber. Vis den adressen tydelig og spør: «Skal den leveres hit, eller til en annen adresse?» Oppgir brukeren en annen adresse, send den som delivery_name, delivery_address, delivery_postal_code og delivery_city i neste kall. Gjenta mottaker og adresse i oppsummeringen før «Send bestilling».
- «[Mobility: …]» er et internt notat som legges til automatisk. Skriv det aldri selv i svaret.

Tjenestebil, bilforsikring og bilordninger (Norge):
- Vil noen låne tjenestebilen privat (tilfeldig transport), forklar kort reglene og tilby skjemaet: ta med markøren [[form:laan-bil]] alene på egen linje. Skjemaet fyller de ut og signerer med fingeren; PDF-en åpnes i Outlook-appen som vedlegg (før turen til leder, etter turen til bil og lønn).
- Alt om tjenestebil, skade, forsikring, lading, kjørebok, privat bil i tjeneste og bilrabatter gjelder Recover Norge. Er det uklart hvilket land brukeren jobber i, spør. Gjelder det ikke Norge, si at dette kommer senere og henvis til bil@recover.no eller Iver support.
- Ved skade: si hva brukeren skal gjøre nå (sikkerhet, veihjelp, skademelding, hvem som skal ha kopi), og skill mellom forsikringsskade og mindre skade under kr 23 000. Spør om det er person-/varebil eller ansattbil hvis egenandelen er relevant.

Videoer:
- Noen steg i FAQ-en har en videomarkør som [[video:utlegg-app]]. Når du svarer på noe som har en slik markør, ta med markøren nøyaktig slik den står, alene på egen linje, rett etter stegene den hører til. Ikke lag nye markører og ikke skriv lenker til videoene selv.

Bestilling av skriver (kun Norge):
- Printerbestillingen gjelder foreløpig bare Norge. Skal skriveren til et annet land, eller skriver brukeren på et annet språk enn norsk og det er uklart hvilket land det gjelder, spør om landet. Gjelder det ikke Norge, si at bestilling for andre land kommer senere, og henvis til Iver support.
- Når noen vil bestille skriver/printer, gi en kort oversikt over modellene og prisene og tilby skjemaet med markøren [[form:printer]] alene på egen linje. Vil brukeren heller at du fyller ut i chatten, gjør det.
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


// ---- Mobility (portal.mobility.no) – ordering with the user's own Mobility login ----
// The page holds each user's Mobility token; the Worker calls Mobility's MCP server itself so it can
// block sending/approving until the user has pressed the confirm button in the chat.
const MOBILITY_BASE = "https://portal.mobility.no";
let mobilityTools = null;

async function mobilityRpc(token, message, sessionId) {
  const headers = {
    "Content-Type": "application/json",
    Accept: "application/json, text/event-stream",
    Authorization: `Bearer ${token}`,
    "MCP-Protocol-Version": "2025-06-18",
  };
  if (sessionId) headers["Mcp-Session-Id"] = sessionId;
  const res = await fetch(`${MOBILITY_BASE}/mcp`, { method: "POST", headers, body: JSON.stringify(message) });
  if (res.status === 401) throw Object.assign(new Error("mobility unauthorized"), { mobilityAuth: true });
  const text = await res.text();
  const payload = text.includes("data:")
    ? text.split("\n").filter((l) => l.startsWith("data:")).map((l) => l.slice(5).trim()).join("\n")
    : text;
  return { sessionId: res.headers.get("Mcp-Session-Id") || sessionId, body: payload ? JSON.parse(payload) : null };
}

async function mobilitySession(token) {
  const init = await mobilityRpc(token, {
    jsonrpc: "2.0", id: 1, method: "initialize",
    params: { protocolVersion: "2025-06-18", capabilities: {}, clientInfo: { name: "spor-arne", version: "1.0" } },
  });
  await mobilityRpc(token, { jsonrpc: "2.0", method: "notifications/initialized", params: {} }, init.sessionId);
  return init.sessionId;
}

async function loadMobilityTools(token) {
  if (mobilityTools) return mobilityTools;
  const sid = await mobilitySession(token);
  const { body } = await mobilityRpc(token, { jsonrpc: "2.0", id: 2, method: "tools/list", params: {} }, sid);
  mobilityTools = (body?.result?.tools || []).map((t) => ({
    name: t.name,
    description: t.description || t.title || t.name,
    input_schema: t.inputSchema || { type: "object", properties: {} },
  }));
  return mobilityTools;
}

// Sending an order or approving/rejecting one needs a real button press from the user.
const CANCEL_VALUES = ["", "NO", "NEI", "CANCEL", "AVBRYT"];

function blockedMobilityCall(name, args, confirmed) {
  // Mobility's own buttons use JA/NEI (the tool text says YES/NO) – treat any non-cancel value as sending.
  const sends = name === "mobility_order" && !CANCEL_VALUES.includes(String(args.confirmation ?? "").trim().toUpperCase());
  const decides = name === "mobility_approve" && ["approve", "reject"].includes(String(args.decision || "").toLowerCase());
  if ((sends || decides) && !confirmed) {
    return "Blokkert: brukeren har ikke trykket på bekreftelsesknappen i chatten ennå. Vis bestillingen/ordren og la brukeren velge med knappene.";
  }
  return null;
}

// "Bestill en iPhone", "trenger nytt nettbrett", "order a phone" … → must go through the Mobility tools.
const ORDER_WORDS = /\b(bestill\w*|kjøp\w*|beställ\w*|bestil\w*|order\w*|buy|trenger|behøver|ny|nytt|nye)\b/i;
const DEVICE_WORDS = /\b(pc\w*|laptop\w*|bærbar\w*|datamaskin\w*|computer|skjerm|skjermer|monitor\w*|tastatur|mus|dokking\w*|docking\w*|headset|hodetelefon\w*|mobil\w*|telefon\w*|phone|iphone\w*|ipad\w*|nettbrett\w*|surfplatta|tablet\w*|samsung|galaxy|pixel|tilbehør|tillbehör|accessor\w*|deksel|skjermbeskytter|lader|smartklokke|watch|airpods|headset|multisim|sim-kort|simkort|esim)\b/i;

// Questions about setting up or fixing a device are not orders.
const NOT_ORDER_WORDS = /\b(sett\w*\s+(?:\w+\s+)?opp|oppsett\w*|konfigur\w*|installer\w*|set up|setup|authenticator|mfa|firmaportal|company portal|intune|overføre|overfør\w*|flytte nummer\w*|virker ikke|fungerer ikke|feil)\b/i;

function wantsDeviceOrder(messages) {
  const last = messages[messages.length - 1];
  const text = typeof last?.content === "string" ? last.content : "";
  return ORDER_WORDS.test(text) && DEVICE_WORDS.test(text) && !NOT_ORDER_WORDS.test(text);
}

// Before Mobility prepares an order, Arne must have asked who it is for. Mobility then proposes the delivery
// address where that employee works; Arne shows it and lets the user keep it or give another one.
function missingOrderDetails(name, args) {
  if (name !== "mobility_order" || !args.sku) return null;
  const hasRecipient = typeof args.ordered_for_name === "string" || args.ordered_for || args.recipient;
  if (hasRecipient) return null;
  return "Ikke sendt til Mobility ennå. Spør brukeren først: Er bestillingen til deg selv eller til en annen ansatt? " +
    "Er den til en annen, trengs navn og mobilnummer. Kall deretter mobility_order igjen med sku og alle tidligere valg, " +
    "pluss ordered_for_name + ordered_for_phone (ordered_for_name tom streng \"\" hvis den er til brukeren selv). " +
    "Ikke spør om leveringsadresse før Mobility har foreslått en.";
}

async function callMobilityTool(token, name, args) {
  const sid = await mobilitySession(token);
  const { body } = await mobilityRpc(token, { jsonrpc: "2.0", id: 3, method: "tools/call", params: { name, arguments: args } }, sid);
  if (body?.error) return { isError: true, text: body.error.message || "Mobility-feil", structured: null };
  const result = body?.result || {};
  const text = (result.content || []).filter((c) => c.type === "text").map((c) => c.text).join("\n");
  return { isError: Boolean(result.isError), text, structured: result.structuredContent || null };
}

async function exchangeMobilityToken(env, body) {
  const form = new URLSearchParams({ client_id: env.MOBILITY_CLIENT_ID, resource: MOBILITY_BASE });
  if (body.refresh_token) {
    form.set("grant_type", "refresh_token");
    form.set("refresh_token", body.refresh_token);
  } else {
    form.set("grant_type", "authorization_code");
    form.set("code", body.code || "");
    form.set("code_verifier", body.code_verifier || "");
    form.set("redirect_uri", body.redirect_uri || "");
  }
  const res = await fetch(`${MOBILITY_BASE}/oauth/token`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded", Accept: "application/json" },
    body: form,
  });
  const data = await res.json().catch(() => ({}));
  if (!res.ok || !data.access_token) return null;
  return { access_token: data.access_token, refresh_token: data.refresh_token || null, expires_in: data.expires_in || 0 };
}

// ---- Suggestion usage (anonymous counts per suggestion id, shared by everyone) ----
const CHIP_IDS = ["travel", "mobile", "loan", "car-damage", "expense", "vacation", "printer", "transfer", "password",
  "iphone", "android", "mfa", "meps", "network"];

// Orders actually sent through Arne, counted per product (anonymous) – answers "what is bought most?".
async function countSubmittedOrder(env, result, args) {
  const lines = Array.isArray(result?.lines) && result.lines.length ? result.lines : [{ name: args.product || args.sku }];
  const counts = (await env.STATS.get("orders", "json")) || {};
  for (const line of lines) {
    const name = String(line.name || "").trim();
    if (!name || /ordrenummer|frakt|shipping/i.test(name)) continue;
    counts[name] = (counts[name] || 0) + (Number(line.quantity) || 1);
  }
  await env.STATS.put("orders", JSON.stringify(counts));
}

async function topOrdered(env, n = 8) {
  const counts = (await env.STATS.get("orders", "json")) || {};
  return Object.entries(counts).sort((a, b) => b[1] - a[1]).slice(0, n);
}

const POPULAR_WORDS = /\b(mest\s+(kjøpt|bestilt|brukt|populær\w*)|populær\w*|mest\s+köpt|most\s+(bought|ordered|popular)|bestselger\w*|mest solgt)\b/i;

// Mobility only shows prices once a product is chosen, so Arne remembers them to show in the picker.
async function learnPrices(env, result) {
  const lines = Array.isArray(result?.lines) ? result.lines : [];
  const fresh = lines.filter((l) => l.sku && Number(l.unit_price) > 0);
  if (!fresh.length) return;
  const known = (await env.STATS.get("prices", "json")) || {};
  let changed = false;
  for (const l of fresh) {
    const price = Math.round(Number(l.unit_price) * 100) / 100;
    if (known[l.sku] !== price) { known[l.sku] = price; changed = true; }
  }
  if (changed) await env.STATS.put("prices", JSON.stringify(known));
}

async function withPrices(env, choice) {
  if (!choice || choice.key !== "sku" || !Array.isArray(choice.options)) return choice;
  const known = (await env.STATS.get("prices", "json").catch(() => null)) || {};
  return {
    ...choice,
    options: choice.options.map((o) => {
      const price = known[o.value] ?? PRICE_SNAPSHOT[o.value];
      return price ? { ...o, price } : o;
    }),
  };
}

async function readChipCounts(env) {
  return (await env.STATS.get("chips", "json")) || {};
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
    if (request.method === "GET" && path === "/stats/chips") {
      return json({ counts: await readChipCounts(env) }, 200, { ...cors, "Cache-Control": "private, max-age=60" });
    }
    if (request.method !== "POST") return json({ error: "Method not allowed" }, 405, cors);

    let body;
    try {
      body = await request.json();
    } catch {
      return json({ error: "Invalid JSON" }, 400, cors);
    }

    if (path === "/stats/chip") {
      if (!CHIP_IDS.includes(body.id)) return json({ error: "unknown" }, 400, cors);
      const counts = await readChipCounts(env);
      counts[body.id] = (counts[body.id] || 0) + 1;
      await env.STATS.put("chips", JSON.stringify(counts));
      return json({ ok: true }, 200, cors);
    }

    // Mobility login: exchange the code (or refresh) for the user's own Mobility token.
    if (path === "/mobility/token") {
      const token = await exchangeMobilityToken(env, body);
      return token ? json(token, 200, cors) : json({ error: "mobility_token_failed" }, 400, cors);
    }

    const messages = cleanMessages(body.messages);
    if (!messages) return json({ error: "Invalid messages" }, 400, cors);
    const lang = LANG_NAMES[body.lang] ? body.lang : "nb";

    const mobilityToken = typeof body.mobilityToken === "string" && body.mobilityToken ? body.mobilityToken : null;
    const confirmed = body.confirmed === true;

    try {
      const client = new Anthropic({ apiKey: env.ANTHROPIC_API_KEY });
      let tools = [PRINTER_ORDER_TOOL, NUMBER_TRANSFER_TOOL];
      let mobilityNote = "Mobility er IKKE koblet til for denne brukeren. Vil brukeren bestille mobil, nettbrett eller tilbehør, forklar kort at de må koble til Mobility først, og ta med markøren [[action:mobility-connect]] alene på egen linje.";
      if (mobilityToken) {
        try {
          tools = tools.concat(await loadMobilityTools(mobilityToken));
          mobilityNote = "Mobility er koblet til for denne brukeren (verktøyene mobility_order og mobility_approve).";
        } catch (err) {
          if (err.mobilityAuth) return json({ error: "mobility_unauthorized" }, 401, cors);
          console.error("Mobility tools unavailable", err);
          mobilityNote = "Mobility svarer ikke akkurat nå. Be brukeren prøve igjen senere.";
        }
      }

      const convo = messages.slice();
      let note = mobilityNote + (confirmed ? " Brukerens siste melding er et trykk på bekreftelsesknappen." : "");
      if (wantsDeviceOrder(messages)) {
        note += mobilityToken
          ? " Brukeren vil bestille utstyr: kall mobility_order nå. Ikke henvis til Mitt Mobility-lenken, Iver eller skjema."
          : " Brukeren vil bestille utstyr: forklar kort at bestillingen gjøres her i chatten via Mobility, og vis [[action:mobility-connect]].";
      }
      const lastText = typeof messages[messages.length - 1]?.content === "string" ? messages[messages.length - 1].content : "";
      if (POPULAR_WORDS.test(lastText)) {
        const top = await topOrdered(env).catch(() => []);
        note += top.length
          ? " Mest bestilt via Arne (antall): " + top.map(([name, c]) => `${name} (${c})`).join(", ") + ". Si at tallene gjelder bestillinger gjort via Arne."
          : " Det er ennå ikke registrert bestillinger via Arne, så du kan ikke si hva som er mest kjøpt. Si det, og nevn heller anbefalt standardtelefon.";
      }
      convo.push({ role: "system", content: note });

      let response;
      let lastMobility = null;
      for (let round = 0; ; round++) {
        response = await client.messages.create({
          model: "claude-haiku-5-5",
          // Short chat answers on a public endpoint – cap output to keep cost bounded.
          max_tokens: 2000,
          output_config: { effort: "low" },
          // The FAQ-heavy system prompt is identical across requests – cache it to cut cost per question.
          system: [{ type: "text", text: systemPrompt(FAQ, lang), cache_control: { type: "ephemeral" } }],
          tools,
          messages: convo,
        });
        const mobilityCalls = response.content.filter((b) => b.type === "tool_use" && b.name.startsWith("mobility_"));
        if (response.stop_reason !== "tool_use" || !mobilityCalls.length || round >= MAX_TOOL_ROUNDS) break;

        // Run Mobility tools ourselves, then let Arne continue with the results.
        convo.push({ role: "assistant", content: response.content });
        const results = [];
        for (const call of response.content.filter((b) => b.type === "tool_use")) {
          if (!call.name.startsWith("mobility_")) {
            results.push({ type: "tool_result", tool_use_id: call.id, content: "OK" });
            continue;
          }
          const blocked = blockedMobilityCall(call.name, call.input || {}, confirmed) || missingOrderDetails(call.name, call.input || {});
          if (blocked) {
            results.push({ type: "tool_result", tool_use_id: call.id, content: blocked, is_error: true });
            continue;
          }
          try {
            const r = await callMobilityTool(mobilityToken, call.name, call.input || {});
            lastMobility = { tool: call.name, args: call.input || {}, result: r.structured };
            if (r.structured) await learnPrices(env, r.structured).catch((e) => console.error("prices", e));
            if (call.name === "mobility_order" && r.structured?.submitted === true) {
              await countSubmittedOrder(env, r.structured, call.input || {}).catch((e) => console.error("order count", e));
            }
            results.push({ type: "tool_result", tool_use_id: call.id, content: r.structured ? JSON.stringify(r.structured) : r.text, is_error: r.isError });
          } catch (err) {
            if (err.mobilityAuth) return json({ error: "mobility_unauthorized" }, 401, cors);
            results.push({ type: "tool_result", tool_use_id: call.id, content: "Mobility svarte ikke. Prøv igjen.", is_error: true });
          }
        }
        convo.push({ role: "user", content: results });
      }

      if (response.stop_reason === "refusal") {
        return json({ reply: null, error: "refusal" }, 200, cors);
      }

      const reply = response.content
        .filter((block) => block.type === "text")
        .map((block) => block.text)
        .join("\n")
        .replace(/\s*\[Mobility:\s*\{[\s\S]*?\}\]\s*/g, "\n")  // internal state note, never shown
        .trim();

      // The page turns these orders into a ready-to-send e-mail; no tool_result round trip is needed.
      const toolInput = (name) => {
        const block = response.content.find((b) => b.type === "tool_use" && b.name === name);
        return block ? block.input : null;
      };

      // Mobility choices become buttons in the chat; the memo keeps the order state for the next turn.
      const res = lastMobility?.result || null;
      const choice = res?.choice && Array.isArray(res.choice.options) ? await withPrices(env, res.choice) : null;
      const memo = lastMobility
        ? JSON.stringify({ tool: lastMobility.tool, args: lastMobility.args, message: res?.message, submitted: res?.submitted, decided: res?.decided })
        : null;

      return json({
        reply,
        order: toolInput(PRINTER_ORDER_TOOL.name),
        transfer: toolInput(NUMBER_TRANSFER_TOOL.name),
        choice,
        memo,
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
