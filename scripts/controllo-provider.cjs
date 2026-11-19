// Controllo periodico dei provider (eseguito da .github/workflows/controllo-provider.yml).
// 1) Senza chiavi: ogni indirizzo usato dal pannello risponde alla richiesta preliminare CORS del sito,
//    con origine, metodo POST e TUTTE le intestazioni che il pannello invia davvero.
// 2) Con le chiavi nei segreti del repository (facoltative, KEY_GEMINI, KEY_QWEN, ...): la chiave e' valida
//    e i modelli del catalogo esistono ancora presso il provider.
// Esce con codice 1 se qualcosa non va; il riepilogo (testo ripulito, chiavi oscurate) va su stdout.
const fs = require('node:fs');
const vm = require('node:vm');

const ORIGIN = 'https://pai-office-automation.github.io';
const TIMEOUT_MS = 20000;

function loadPanel(fetchImpl = fetch) {
  const context = vm.createContext({
    localStorage: { getItem: () => null, setItem: () => {}, removeItem: () => {} },
    fetch: fetchImpl, setTimeout, clearTimeout, AbortController, URL
  });
  vm.runInContext(fs.readFileSync('public/ai.js', 'utf8'), context);
  return s => vm.runInContext(s, context);
}

/** Indirizzi e intestazioni da provare, uguali a quelli delle chiamate del pannello. */
function endpoints(run) {
  const compat = run('AI_COMPAT');
  const eps = {
    gemini: { url: 'https://generativelanguage.googleapis.com/v1beta/models/x:generateContent', headers: ['content-type', 'x-goog-api-key'] },
    openai: { url: 'https://api.openai.com/v1/chat/completions', headers: ['authorization', 'content-type'] },
    claude: { url: 'https://api.anthropic.com/v1/messages',
              headers: ['anthropic-beta', 'anthropic-dangerous-direct-browser-access', 'anthropic-version', 'content-type', 'x-api-key'] }
  };
  for (const [p, cfg] of Object.entries(compat)) {
    eps[p] = { url: `${cfg.base}/chat/completions`,
               headers: ['authorization', 'content-type', ...Object.keys(cfg.headers || {}).map(h => h.toLowerCase())] };
  }
  return eps;
}

function listHas(value, item) {
  if (!value) return false;
  const parts = value.split(',').map(x => x.trim().toLowerCase());
  return parts.includes('*') || parts.includes(item.toLowerCase());
}

/** null se il pannello puo' chiamare l'indirizzo, altrimenti il motivo. */
async function checkCors({ url, headers }, fetchImpl = fetch) {
  const r = await fetchImpl(url, { method: 'OPTIONS', signal: AbortSignal.timeout(TIMEOUT_MS), headers: {
    Origin: ORIGIN, 'Access-Control-Request-Method': 'POST', 'Access-Control-Request-Headers': headers.join(',') } });
  if (r.status >= 400) return `richiesta preliminare rifiutata (HTTP ${r.status})`;
  const allowOrigin = r.headers.get('access-control-allow-origin');
  if (allowOrigin !== '*' && allowOrigin !== ORIGIN) return `origine del sito non accettata (Allow-Origin: ${allowOrigin || 'assente'})`;
  const methods = r.headers.get('access-control-allow-methods');
  if (methods && !listHas(methods, 'POST')) return `metodo POST non accettato (Allow-Methods: ${methods})`;
  const allowHeaders = r.headers.get('access-control-allow-headers');
  const missing = headers.filter(h => !listHas(allowHeaders, h));
  if (missing.length) return `intestazioni non accettate: ${missing.join(', ')}`;
  return null;
}

/** Testo sicuro per il riepilogo pubblico: chiavi oscurate, niente Markdown/HTML attivo, lunghezza limitata. */
function sanitize(text, secrets = []) {
  let s = String(text);
  for (const k of secrets) if (k && k.length >= 6) s = s.split(k).join('[chiave oscurata]');
  s = s.replace(/\b(sk|AIza|key)[-_A-Za-z0-9]{12,}/g, '[chiave oscurata]');
  return s.replace(/[\r\n]+/g, ' ')
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')   // niente HTML attivo
    .replace(/[`*_[\]|#!\\]/g, c => '\\' + c)                             // niente Markdown attivo
    .slice(0, 300);
}

async function main() {
  const run = loadPanel();
  const NAMES = run('AI_PROVIDER_NAMES');
  const CATALOG = run('AI_CATALOG');
  const secrets = Object.entries(process.env).filter(([k]) => k.startsWith('KEY_')).map(([, v]) => v);
  const problems = [], notes = [];
  for (const [p, ep] of Object.entries(endpoints(run))) {
    try {
      const err = await checkCors(ep);
      if (err) problems.push(`- **${NAMES[p]}** (${ep.url}): ${sanitize(err, secrets)}`);
    } catch (e) { problems.push(`- **${NAMES[p]}** (${ep.url}): non raggiungibile (${sanitize(e.message, secrets)})`); }

    const key = process.env[`KEY_${p.toUpperCase()}`];
    if (!key) { notes.push(`${NAMES[p]}: modelli non verificati (nessuna chiave nei segreti)`); continue; }
    try {
      const ids = await run(`aiListModels(${JSON.stringify(p)}, ${JSON.stringify(key)})`);
      const missing = CATALOG.filter(m => m.provider === p && !ids.includes(m.id)).map(m => m.id);
      if (missing.length) problems.push(`- **${NAMES[p]}**: modelli del catalogo non piu' disponibili: ${sanitize(missing.join(', '), secrets)}`);
    } catch (e) { problems.push(`- **${NAMES[p]}**: verifica della chiave non riuscita (${sanitize(String(e.message).split('\n')[0], secrets)})`); }
  }
  console.log(problems.length
    ? `Il controllo dei provider ha trovato problemi:\n\n${problems.join('\n')}\n\nCorreggere \`AI_COMPAT\`/\`AI_CATALOG\` in public/ai.js (vale subito per tutti) e il catalogo di Outlook in ModelCatalog.cs.`
    : 'Indirizzi raggiungibili; richiesta preliminare CORS accettata con origine, metodo e intestazioni del pannello.');
  if (notes.length) console.log(`\n<details><summary>Note</summary>\n\n${notes.map(n => '- ' + n).join('\n')}\n</details>`);
  return problems.length ? 1 : 0;
}

module.exports = { checkCors, sanitize, endpoints, loadPanel, ORIGIN };
if (require.main === module) main().then(code => process.exit(code));
