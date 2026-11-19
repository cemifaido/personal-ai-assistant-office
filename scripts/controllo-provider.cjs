// Controllo periodico dei provider (eseguito da .github/workflows/controllo-provider.yml).
// 1) Senza chiavi: ogni indirizzo usato dal pannello risponde e accetta le chiamate dal sito (CORS).
// 2) Con le chiavi nei segreti del repository (facoltative, KEY_GEMINI, KEY_QWEN, ...): la chiave e' valida
//    e i modelli del catalogo esistono ancora presso il provider.
// Esce con codice 1 se qualcosa non va; il riepilogo in Markdown va su stdout (diventa la segnalazione).
const fs = require('node:fs');
const vm = require('node:vm');

const ORIGIN = 'https://pai-office-automation.github.io';
const context = vm.createContext({
  localStorage: { getItem: () => null, setItem: () => {}, removeItem: () => {} },
  fetch, setTimeout, clearTimeout, AbortController, URL
});
vm.runInContext(fs.readFileSync('public/ai.js', 'utf8'), context);
const run = s => vm.runInContext(s, context);
const NAMES = run('AI_PROVIDER_NAMES');
const COMPAT = run('AI_COMPAT');
const CATALOG = run('AI_CATALOG');

// Indirizzo da provare con la richiesta preliminare CORS, come la farebbe il pannello.
const ENDPOINTS = {
  gemini: { url: 'https://generativelanguage.googleapis.com/v1beta/models/x:generateContent', headers: 'content-type,x-goog-api-key' },
  openai: { url: 'https://api.openai.com/v1/chat/completions', headers: 'authorization,content-type' },
  claude: { url: 'https://api.anthropic.com/v1/messages', headers: 'anthropic-dangerous-direct-browser-access,anthropic-version,content-type,x-api-key' }
};
for (const [p, cfg] of Object.entries(COMPAT)) {
  ENDPOINTS[p] = { url: `${cfg.base}/chat/completions`,
                   headers: ['authorization', 'content-type', ...Object.keys(cfg.headers || {})].join(',').toLowerCase() };
}

async function checkCors(p, { url, headers }) {
  const r = await fetch(url, { method: 'OPTIONS', headers: {
    Origin: ORIGIN, 'Access-Control-Request-Method': 'POST', 'Access-Control-Request-Headers': headers } });
  const allow = r.headers.get('access-control-allow-origin');
  if (r.status >= 400) return `richiesta preliminare rifiutata (HTTP ${r.status})`;
  if (allow !== '*' && allow !== ORIGIN) return `chiamate dal pannello non accettate (Access-Control-Allow-Origin: ${allow || 'assente'})`;
  return null;
}

(async () => {
  const problems = [], notes = [];
  for (const [p, ep] of Object.entries(ENDPOINTS)) {
    try {
      const err = await checkCors(p, ep);
      if (err) problems.push(`- **${NAMES[p]}** (${ep.url}): ${err}`);
    } catch (e) { problems.push(`- **${NAMES[p]}** (${ep.url}): non raggiungibile (${e.message})`); }

    const key = process.env[`KEY_${p.toUpperCase()}`];
    if (!key) { notes.push(`${NAMES[p]}: modelli non verificati (nessuna chiave nei segreti)`); continue; }
    try {
      const ids = await run(`aiListModels(${JSON.stringify(p)}, ${JSON.stringify(key)})`);
      const missing = CATALOG.filter(m => m.provider === p && !ids.includes(m.id)).map(m => m.id);
      if (missing.length) problems.push(`- **${NAMES[p]}**: modelli del catalogo non piu' disponibili: ${missing.join(', ')}`);
    } catch (e) { problems.push(`- **${NAMES[p]}**: verifica della chiave non riuscita (${String(e.message).split('\n')[0]})`); }
  }
  console.log(problems.length
    ? `Il controllo settimanale dei provider ha trovato problemi:\n\n${problems.join('\n')}\n\nCorreggere \`AI_COMPAT\`/\`AI_CATALOG\` in public/ai.js (vale subito per tutti) e il catalogo di Outlook in ModelCatalog.cs.`
    : 'Tutti i provider rispondono e accettano le chiamate dal pannello.');
  if (notes.length) console.log(`\n<details><summary>Note</summary>\n\n${notes.map(n => '- ' + n).join('\n')}\n</details>`);
  process.exit(problems.length ? 1 : 0);
})();
