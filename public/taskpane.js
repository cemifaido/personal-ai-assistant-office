/**
 * Personal AI Assistant — controller del pannello (Word, Excel, Outlook).
 * Le chiamate ai modelli sono in ai.js (dirette, nessun server locale).
 */

let currentHost = 'browser';
let lastGeneratedText = '';
let currentAbort = null;
let excelOverwriteArmed = false;

Office.onReady((info) => {
  const hostBadge = document.getElementById('host-badge');
  const browserTabs = document.getElementById('browser-tabs');

  buildModelSelect();

  if (info && info.host === Office.HostType.Word) {
    setHost('word', 'Word', hostBadge);
  } else if (info && info.host === Office.HostType.Excel) {
    setHost('excel', 'Excel', hostBadge);
  } else if (info && info.host === Office.HostType.Outlook) {
    setHost('outlook', 'Outlook', hostBadge);
    loadOutlookData();
  } else {
    // Modalità browser / test (pagina aperta fuori da Office)
    currentHost = 'word';
    hostBadge.textContent = 'Browser / Test';
    browserTabs.classList.remove('hidden');
    showPanel('word');
  }

  initEventListeners();
  loadSavedKeys();
  updateKeyBadges();
});

function setHost(host, label, badge) {
  currentHost = host;
  badge.textContent = label;
  badge.className = `host-badge ${host}`;
  showPanel(host);
}

function showPanel(panelName) {
  ['word', 'excel', 'outlook'].forEach(p => document.getElementById(`panel-${p}`).classList.add('hidden'));
  const target = document.getElementById(`panel-${panelName}`);
  if (target) target.classList.remove('hidden');
  document.querySelectorAll('.tab-btn').forEach((btn) => {
    btn.classList.toggle('active', btn.dataset.tab === panelName);
  });
  if (document.getElementById('browser-tabs').classList.contains('hidden') === false) {
    currentHost = panelName;
  }
}

/* ===================== Storage sicuro ===================== */

function storeGet(k) { try { return localStorage.getItem(k) || ''; } catch (e) { return ''; } }
function storeSet(k, v) { try { v ? localStorage.setItem(k, v) : localStorage.removeItem(k); } catch (e) {} }

/* ===================== Modelli ===================== */

function buildModelSelect() {
  const select = document.getElementById('select-model');
  select.innerHTML = '';
  for (const provider of Object.keys(AI_PROVIDER_NAMES)) {
    const group = document.createElement('optgroup');
    group.label = AI_PROVIDER_NAMES[provider];
    AI_CATALOG.filter(m => m.provider === provider).forEach(m => {
      const opt = document.createElement('option');
      opt.value = `${m.provider}:${m.id}`;
      opt.textContent = m.label;
      group.appendChild(opt);
    });
    select.appendChild(group);
  }
  const saved = storeGet('active_model_spec');
  if (saved && [...select.options].some(o => o.value === saved)) select.value = saved;
}

function getSelectedModelInfo() {
  const select = document.getElementById('select-model');
  const [provider, model] = select.value.split(':');
  return { provider, model, label: select.options[select.selectedIndex].text };
}

/* ===================== Chiavi ===================== */

const KEY_INPUTS = { gemini: 'input-gemini-key', openai: 'input-openai-key', claude: 'input-claude-key' };

function loadSavedKeys() {
  for (const [p, id] of Object.entries(KEY_INPUTS)) {
    document.getElementById(id).value = storeGet(`${p}_api_key`);
  }
}

function updateKeyBadges() {
  for (const p of Object.keys(KEY_INPUTS)) {
    const el = document.getElementById(`badge-${p}-status`);
    const has = Boolean(storeGet(`${p}_api_key`));
    el.textContent = has ? 'Salvata' : 'Mancante';
    el.className = `key-badge ${has ? 'ready' : 'missing'}`;
  }
}

async function verifyKeys() {
  const lines = [];
  for (const p of Object.keys(KEY_INPUTS)) {
    const key = storeGet(`${p}_api_key`).trim();
    const el = document.getElementById(`badge-${p}-status`);
    if (!key) { lines.push(`${AI_PROVIDER_NAMES[p]}: nessuna chiave`); continue; }
    el.textContent = 'Verifica…';
    try {
      const available = await aiListModels(p, key);
      const missing = AI_CATALOG.filter(m => m.provider === p && !available.includes(m.id)).map(m => m.id);
      el.textContent = 'Valida';
      el.className = 'key-badge ready';
      lines.push(`${AI_PROVIDER_NAMES[p]}: chiave valida` +
        (missing.length ? ` — modelli non trovati: ${missing.join(', ')}` : ' — tutti i modelli disponibili'));
    } catch (err) {
      el.textContent = 'Errore';
      el.className = 'key-badge missing';
      lines.push(`${AI_PROVIDER_NAMES[p]}: ${err.message.split('\n')[0]}`);
    }
  }
  notify(lines.join('\n'), 'info');
}

/* ===================== UI helpers ===================== */

function notify(msg, type = 'info') {
  const box = document.getElementById('notice');
  document.getElementById('notice-text').textContent = msg;
  box.className = `notice ${type}`;
  box.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}

function setLoader(visible, text = 'Elaborazione in corso...') {
  document.getElementById('loader-text').textContent = text;
  document.getElementById('global-loader').classList.toggle('hidden', !visible);
}

function initEventListeners() {
  document.getElementById('select-model').addEventListener('change', (e) => {
    storeSet('active_model_spec', e.target.value);
    notify(`Modello impostato: ${getSelectedModelInfo().label}`);
  });

  document.getElementById('btn-toggle-settings').addEventListener('click', () => {
    document.getElementById('settings-card').classList.toggle('hidden');
  });

  document.getElementById('btn-save-all-keys').addEventListener('click', () => {
    for (const [p, id] of Object.entries(KEY_INPUTS)) {
      storeSet(`${p}_api_key`, document.getElementById(id).value.trim());
    }
    updateKeyBadges();
    notify('Chiavi salvate in questo pannello.', 'success');
  });
  document.getElementById('btn-verify-keys').addEventListener('click', verifyKeys);

  document.getElementById('notice-close').addEventListener('click', () => {
    document.getElementById('notice').classList.add('hidden');
  });
  document.getElementById('btn-cancel').addEventListener('click', () => {
    if (currentAbort) currentAbort.abort();
  });

  document.querySelectorAll('.tab-btn').forEach((btn) => {
    btn.addEventListener('click', () => showPanel(btn.dataset.tab));
  });

  document.getElementById('btn-close-result').addEventListener('click', () => {
    document.getElementById('result-container').classList.add('hidden');
  });
  document.getElementById('btn-copy-result').addEventListener('click', async () => {
    if (!lastGeneratedText) return;
    try { await navigator.clipboard.writeText(lastGeneratedText); notify('Copiato negli appunti.', 'success'); }
    catch (e) { notify('Copia non consentita: seleziona il testo nel riquadro e usa Ctrl+C.', 'warn'); }
  });

  document.getElementById('btn-apply-replace').addEventListener('click', () => applyTextToDocument(lastGeneratedText, 'replace'));
  document.getElementById('btn-apply-insert').addEventListener('click', () => applyTextToDocument(lastGeneratedText, 'insert'));

  // Word
  document.getElementById('btn-word-rewrite').addEventListener('click', handleWordRewrite);
  document.getElementById('btn-word-grammar').addEventListener('click', handleWordGrammar);
  document.getElementById('btn-word-summarize').addEventListener('click', handleWordSummarize);
  document.getElementById('btn-word-translate').addEventListener('click', handleWordTranslate);
  document.getElementById('btn-word-generate').addEventListener('click', handleWordGenerate);
  document.getElementById('btn-word-ask-doc').addEventListener('click', handleWordAskFullDoc);

  // Excel
  document.getElementById('btn-excel-gen-formula').addEventListener('click', handleExcelFormula);
  document.getElementById('btn-excel-analyze-range').addEventListener('click', handleExcelAnalyzeRange);
  document.getElementById('btn-excel-gen-table').addEventListener('click', handleExcelGenTable);

  // Outlook (solo account che supportano gli add-in web; per POP usare la macro VBA)
  document.getElementById('btn-outlook-read').addEventListener('click', loadOutlookData);
  document.querySelectorAll('.btn-quick-reply').forEach((btn) => {
    btn.addEventListener('click', () => handleOutlookQuickReply(btn.dataset.intent));
  });
  document.getElementById('btn-outlook-summarize').addEventListener('click', handleOutlookSummarize);
  document.getElementById('btn-outlook-gen-custom').addEventListener('click', handleOutlookCustomReply);
}

/* ===================== Risultato ===================== */

// Rimuove prefissi tipici ("Ecco il testo:", "**Riscrittura:**") e virgolette avvolgenti.
function cleanForDocument(rawText) {
  let text = aiMarkdownToPlain(rawText || '').trim();
  text = text.replace(/^(Ecco\s+(la\s+riscrittura|il\s+testo\s+corretto|la\s+traduzione|la\s+bozza|il\s+testo|la\s+versione)[^:\n]*|Riscrittura|Testo\s+riscritto|Testo\s+corretto|Bozza|Versione\s+migliorata|Traduzione|Risposta|Output)\s*:\s*/i, '');
  if ((text.startsWith('"') && text.endsWith('"')) || (text.startsWith('«') && text.endsWith('»'))) {
    text = text.slice(1, -1).trim();
  }
  return text;
}

let typingIntervalId = null;

function displayResult(result) {
  const cleaned = cleanForDocument(result.text);
  lastGeneratedText = cleaned;
  excelOverwriteArmed = false;

  const container = document.getElementById('result-container');
  const content = document.getElementById('result-content');
  const badge = document.getElementById('result-model-badge');

  const changed = result.actualModel && !result.actualModel.startsWith(result.requestedModel);
  badge.textContent = `Risultato — ${result.actualModel || result.requestedModel}` +
    (changed ? ` (richiesto: ${result.requestedModel})` : '');

  container.classList.remove('hidden');
  container.scrollIntoView({ behavior: 'smooth' });

  if (typingIntervalId) { clearInterval(typingIntervalId); typingIntervalId = null; }

  // Effetto digitazione solo nel pannello (nel documento si inserisce in un'unica operazione).
  content.innerHTML = '<span id="typewriter-text"></span><span class="typing-cursor"></span>';
  const textSpan = document.getElementById('typewriter-text');
  const tokens = cleaned.split(/(\s+)/);
  let i = 0;
  typingIntervalId = setInterval(() => {
    // più token per tick sui testi lunghi, così non si aspetta troppo
    const step = Math.max(1, Math.ceil(tokens.length / 400));
    for (let s = 0; s < step && i < tokens.length; s++) textSpan.textContent += tokens[i++];
    content.scrollTop = content.scrollHeight;
    if (i >= tokens.length) {
      clearInterval(typingIntervalId);
      typingIntervalId = null;
      content.querySelector('.typing-cursor')?.remove();
    }
  }, 18);
}

/** Esegue una richiesta al modello con loader, annullamento ed errori uniformi. */
async function runAI(loaderText, { prompt, context, system }) {
  const { provider, model } = getSelectedModelInfo();
  currentAbort = new AbortController();
  setLoader(true, `${loaderText} (${AI_PROVIDER_NAMES[provider]})`);
  try {
    const result = await aiCall({ provider, model, prompt, context, system, signal: currentAbort.signal });
    displayResult(result);
  } catch (err) {
    notify(err.message, 'error');
  } finally {
    currentAbort = null;
    setLoader(false);
  }
}

/* ===================== Word ===================== */

async function getWordSelection() {
  if (typeof Word === 'undefined') {
    return 'Questo è un testo di prova simulato nel browser per testare le funzionalità di riscrittura multi-modello.';
  }
  return Word.run(async (context) => {
    const sel = context.document.getSelection();
    sel.load('text');
    await context.sync();
    return sel.text.trim();
  });
}

async function getWordFullBody() {
  if (typeof Word === 'undefined') {
    return 'Documento di prova per il browser. Contiene specifiche, note e requisiti per il test dei modelli AI.';
  }
  return Word.run(async (context) => {
    const body = context.document.body;
    body.load('text');
    await context.sync();
    return body.text.trim();
  });
}

async function withSelection(fn) {
  try {
    const text = await getWordSelection();
    if (!text) { notify('Seleziona prima del testo nel documento Word.', 'warn'); return; }
    await fn(text);
  } catch (err) {
    notify('Errore Word: ' + err.message, 'error');
  }
}

const WORD_ONLY_TEXT = 'Restituisci esclusivamente il testo finale, senza prefissi, commenti, virgolette o formattazione markdown.';

function handleWordRewrite() {
  return withSelection((text) => {
    const tone = document.getElementById('word-tone-select').value;
    return runAI(`Riscrittura (${tone})`, {
      prompt: `Riscrivi il testo migliorandone fluidità e leggibilità, con tono ${tone}. Mantieni la lingua originale e la suddivisione in paragrafi.`,
      context: text,
      system: `Sei un editor per Microsoft Word. ${WORD_ONLY_TEXT}`
    });
  });
}

function handleWordGrammar() {
  return withSelection((text) => runAI('Correzione bozze', {
    prompt: 'Correggi errori grammaticali, ortografici e di punteggiatura. Se il testo è già corretto, restituiscilo identico. Non cambiare stile né contenuto.',
    context: text,
    system: `Sei un correttore di bozze per Microsoft Word. ${WORD_ONLY_TEXT}`
  }));
}

function handleWordSummarize() {
  return withSelection((text) => runAI('Riassunto', {
    prompt: 'Fornisci un riassunto esecutivo in 3-5 punti elenco sintetici.',
    context: text
  }));
}

function handleWordTranslate() {
  return withSelection((text) => runAI('Traduzione', {
    prompt: 'Traduci il testo in inglese in modo professionale e accurato.',
    context: text,
    system: `Sei un traduttore per Microsoft Word. ${WORD_ONLY_TEXT}`
  }));
}

function handleWordGenerate() {
  const prompt = document.getElementById('word-custom-prompt').value.trim();
  if (!prompt) { notify('Scrivi una richiesta nel campo di testo.', 'warn'); return; }
  return runAI('Generazione testo', {
    prompt,
    system: `Sei un redattore esperto. Genera contenuti ben strutturati adatti a un documento professionale. ${WORD_ONLY_TEXT}`
  });
}

async function handleWordAskFullDoc() {
  const question = document.getElementById('word-custom-prompt').value.trim();
  if (!question) { notify('Scrivi la domanda da porre sul documento.', 'warn'); return; }
  try {
    setLoader(true, 'Lettura del documento...');
    const body = await getWordFullBody();
    setLoader(false);
    await runAI(`Domanda sul documento (${Math.round(body.length / 1000)}k caratteri)`, {
      prompt: `Rispondi alla domanda basandoti solo sul documento. Se l'informazione non c'è, dillo.\nDOMANDA: ${question}`,
      context: body,
      system: 'Rispondi in italiano citando le parti pertinenti del documento fornito.'
    });
  } catch (err) {
    setLoader(false);
    notify('Errore Word: ' + err.message, 'error');
  }
}

/* ===================== Excel ===================== */

const EXCEL_MAX_CELLS = 40000; // oltre, si inviano solo le prime righe (dichiarandolo)

function colLetter(index) {
  let s = '', n = index + 1;
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

function excelErrorMessage(err) {
  if (err && (err.code === 'InvalidSelection' || /selection/i.test(err.message || ''))) {
    return 'Seleziona un solo intervallo continuo (non più aree con Ctrl).';
  }
  return 'Errore Excel: ' + (err && err.message ? err.message : err);
}

/** Struttura del foglio: nome, area dati, intestazioni con lettere di colonna, cella attiva. */
async function readExcelOverview() {
  return Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getActiveWorksheet();
    const used = sheet.getUsedRangeOrNullObject(true);
    const active = ctx.workbook.getActiveCell();
    sheet.load('name');
    active.load('address');
    used.load(['address', 'rowIndex', 'columnIndex', 'columnCount', 'rowCount']);
    await ctx.sync();
    if (used.isNullObject) return `Foglio "${sheet.name}" vuoto. Cella attiva: ${active.address}.`;
    const cols = Math.min(used.columnCount, 60);
    const sampleRows = Math.min(used.rowCount, 4); // intestazione + 3 righe di esempio
    const sample = sheet.getRangeByIndexes(used.rowIndex, used.columnIndex, sampleRows, cols);
    sample.load(['text', 'values', 'numberFormat']);
    await ctx.sync();

    // Per ogni colonna: intestazione, tipo REALE del contenuto ed esempi come visualizzati.
    const headers = sample.text[0].map((h, i) => {
      let type = 'vuota';
      const examples = [];
      for (let r = 1; r < sampleRows; r++) {
        const v = sample.values[r][i];
        const fmt = String(sample.numberFormat[r][i] || '');
        if (v === '' || v === null) continue;
        examples.push(sample.text[r][i]);
        if (typeof v === 'number') {
          type = /[dgmyaDGMYA]/.test(fmt) && !/^General|^Generale/i.test(fmt)
            ? `data vera (numero seriale, formato "${fmt}")` : 'numero';
        } else if (typeof v === 'boolean') type = 'logico';
        else type = 'testo';
      }
      return `${colLetter(used.columnIndex + i)} = "${h || '(senza intestazione)'}" — ${type}${examples.length ? ' — es.: ' + examples.join(' | ') : ''}`;
    }).join('\n');

    const lang = (Office.context && Office.context.displayLanguage) || 'it-IT';
    return `Foglio: "${sheet.name}"\nArea dati: ${used.address} (${used.rowCount} righe, la prima è l'intestazione)\n` +
           `Cella attiva (dove verrà inserita la formula): ${active.address}\n` +
           `Lingua di Excel: ${lang}\nColonne:\n${headers}`;
  });
}

/** Legge la selezione limitata all'area con dati. Restituisce testo per il modello. */
async function readExcelSelection() {
  return Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getActiveWorksheet();
    const sel = ctx.workbook.getSelectedRange();
    const used = sheet.getUsedRangeOrNullObject(true);
    sheet.load('name');
    used.load(['rowIndex']);
    await ctx.sync();
    if (used.isNullObject) throw new Error('Il foglio non contiene dati.');

    // Colonne intere / selezioni enormi: si tiene solo la parte con dati.
    const range = sel.getIntersectionOrNullObject(sheet.getUsedRange(true));
    range.load(['address', 'rowCount', 'columnCount', 'rowIndex', 'columnIndex']);
    await ctx.sync();
    if (range.isNullObject) throw new Error('La selezione non contiene dati.');

    const maxRows = Math.max(1, Math.floor(EXCEL_MAX_CELLS / range.columnCount));
    const take = Math.min(range.rowCount, maxRows);
    const part = take < range.rowCount ? range.getResizedRange(take - range.rowCount, 0) : range;
    part.load(['values', 'text']);

    // Se la selezione non include la riga di intestazione, la aggiungiamo.
    let header = null;
    if (range.rowIndex > used.rowIndex) {
      header = sheet.getRangeByIndexes(used.rowIndex, range.columnIndex, 1, range.columnCount);
      header.load('text');
    }
    await ctx.sync();

    const headerRow = header ? header.text[0] : part.text[0];
    const dataValues = header ? part.values : part.values.slice(1);
    const dataText = header ? part.text : part.text.slice(1);

    // Statistiche esatte calcolate qui: i modelli sbagliano spesso i conti su molte righe.
    const stats = [];
    for (let c = 0; c < range.columnCount; c++) {
      const nums = [];
      let looksDate = false;
      dataValues.forEach((row, r) => {
        if (typeof row[c] === 'number') {
          nums.push(row[c]);
          if (/\d{1,2}[\/.-]\d{1,2}[\/.-]\d{2,4}/.test(String(dataText[r][c]))) looksDate = true;
        }
      });
      if (nums.length >= 2 && !looksDate) {
        const sum = nums.reduce((a, b) => a + b, 0);
        const fmt = (x) => Number.isInteger(x) ? String(x) : x.toFixed(2);
        stats.push(`${colLetter(range.columnIndex + c)} "${headerRow[c] || ''}": n=${nums.length}, somma=${fmt(sum)}, media=${fmt(sum / nums.length)}, min=${fmt(Math.min(...nums))}, max=${fmt(Math.max(...nums))}`);
      }
    }

    const lines = (header ? [headerRow].concat(part.text) : part.text).map(r => r.join('\t')).join('\n');
    const truncated = take < range.rowCount ? `\nATTENZIONE: inviate solo le prime ${take} righe su ${range.rowCount}.` : '';
    return `Foglio: "${sheet.name}" — intervallo con dati: ${range.address} (${range.rowCount} righe × ${range.columnCount} colonne)${truncated}\n\n` +
           `STATISTICHE CALCOLATE (esatte, su tutte le righe inviate):\n${stats.join('\n') || '(nessuna colonna numerica)'}\n\n` +
           `DATI (come visualizzati, separati da TAB):\n${lines}`;
  });
}

async function handleExcelFormula() {
  const prompt = document.getElementById('excel-formula-prompt').value.trim();
  if (!prompt) { notify('Descrivi cosa vuoi calcolare.', 'warn'); return; }
  let overview = '';
  try {
    overview = typeof Excel !== 'undefined'
      ? await readExcelOverview()
      : 'Foglio: "Prova"\nArea dati: A1:C4\nCella attiva: E2\nColonne:\nA = Mese\nB = Vendite\nC = Costi';
  } catch (err) {
    notify(excelErrorMessage(err), 'error');
    return;
  }
  return runAI('Generazione formula', {
    prompt: `L'utente vuole una formula Excel per: "${prompt}".
Regole:
- Usa le colonne reali indicate nel contesto. Se la formula va ripetuta riga per riga, riferisciti alla STESSA RIGA della cella attiva (es. cella attiva C2 → B2), così si può trascinare in giù.
- Rispetta il tipo indicato: una "data vera" è un numero seriale, NON va scomposta con funzioni di testo (LEFT/MID/RIGHT); un "testo" sì.
- Nomi delle funzioni in inglese e separatore argomenti virgola (verranno tradotti automaticamente).
- ATTENZIONE: le stringhe di formato dentro TEXT() NON vengono tradotte: scrivile nella lingua di Excel indicata (it-IT: g=giorno, m=mese, a=anno, es. "mm/gg/aaaa").
- Se il risultato si ottiene meglio senza formula (es. cambiare il formato delle celle), dillo nella spiegazione.
Prima riga: "FORMULA: =...". Poi una spiegazione di massimo 3 righe, citando le colonne usate.`,
    context: overview,
    system: 'Sei un esperto di Microsoft Excel e formule avanzate.'
  });
}

async function handleExcelAnalyzeRange() {
  let rangeData;
  try {
    setLoader(true, 'Lettura delle celle selezionate...');
    rangeData = typeof Excel !== 'undefined'
      ? await readExcelSelection()
      : 'Foglio: "Prova" — intervallo con dati: A1:C4\n\nSTATISTICHE CALCOLATE:\nB "Vendite": n=3, somma=38000, media=12666.67, min=11000, max=15000\n\nDATI:\nMese\tVendite\tCosti\nGennaio\t12000\t8000\nFebbraio\t15000\t8500\nMarzo\t11000\t9200';
  } catch (err) {
    setLoader(false);
    notify(excelErrorMessage(err), 'error');
    return;
  }
  setLoader(false);
  return runAI('Analisi dati', {
    prompt: `Analizza questi dati Excel. Per i valori numerici usa le STATISTICHE CALCOLATE (sono esatte): non ricalcolarle.
Fornisci:
- Sintesi (cosa contengono i dati, valori medi e intervalli principali)
- Punti di attenzione o anomalie (indica righe/valori specifici)
- Un suggerimento pratico.`,
    context: rangeData
  });
}

function handleExcelGenTable() {
  const prompt = document.getElementById('excel-data-prompt').value.trim();
  if (!prompt) { notify('Descrivi che tabella vuoi generare.', 'warn'); return; }
  return runAI('Generazione tabella', {
    prompt: `Genera una tabella con: "${prompt}".
Formato OBBLIGATORIO: una riga per record, colonne separate da un carattere TAB, prima riga di intestazione.
Numeri senza separatore delle migliaia e con il punto come separatore decimale. Nessun altro testo.`,
    system: 'Restituisci solo dati tabellari separati da TAB.'
  });
}

/** Converte il testo del modello in una matrice rettangolare. Restituisce null se non è una tabella. */
function parseTable(text) {
  const lines = text.replace(/\r/g, '').split('\n').map(l => l.trimEnd()).filter(l => l.trim() !== '');
  let sep = null;
  if (lines.some(l => l.includes('\t'))) sep = '\t';
  else if (lines.length > 1 && lines.every(l => l.includes(';'))) sep = ';';
  else if (lines.length > 1 && lines.every(l => l.trim().startsWith('|'))) sep = '|';
  if (!sep) return null;

  let rows = lines.map(l => {
    let cells = l.split(sep).map(c => c.trim());
    if (sep === '|') cells = cells.slice(1, cells[cells.length - 1] === '' ? -1 : undefined);
    return cells;
  });
  if (sep === '|') rows = rows.filter(r => !r.every(c => /^:?-{2,}:?$/.test(c))); // riga separatrice markdown
  const width = Math.max(...rows.map(r => r.length));
  return rows.map(r => {
    const padded = r.concat(Array(width - r.length).fill(''));
    // Un testo che inizia con "=" deve restare testo, non diventare una formula.
    return padded.map(c => (c.startsWith('=') ? "'" + c : c));
  });
}

async function applyToExcel(text, rawText) {
  await Excel.run(async (context) => {
    const start = context.workbook.getSelectedRange().getCell(0, 0);

    const formulaMatch = rawText.match(/FORMULA:\s*`*(=[^\r\n`]+)/i) || text.match(/^(=[^\r\n]+)/);
    if (formulaMatch) {
      start.formulas = [[formulaMatch[1].trim()]];
      await context.sync();
      notify('Formula inserita nella cella attiva.', 'success');
      return;
    }

    const matrix = parseTable(text) || [[text]];
    const target = start.getResizedRange(matrix.length - 1, matrix[0].length - 1);
    target.load(['address', 'values']);
    await context.sync();

    const occupied = target.values.some(r => r.some(v => v !== '' && v !== null));
    if (occupied && !excelOverwriteArmed) {
      excelOverwriteArmed = true;
      notify(`L'area ${target.address} contiene già dati. Premi di nuovo il pulsante per sovrascriverla, oppure seleziona un'altra cella.`, 'warn');
      return;
    }
    target.values = matrix;
    await context.sync();
    excelOverwriteArmed = false;
    notify(`Inseriti ${matrix.length}×${matrix[0].length} valori in ${target.address}.`, 'success');
  });
}

/* ===================== Outlook (add-in web) ===================== */

let currentEmailBody = '';

function loadOutlookData() {
  const item = (typeof Office !== 'undefined' && Office.context && Office.context.mailbox) ? Office.context.mailbox.item : null;
  const sender = document.getElementById('email-sender');
  const subject = document.getElementById('email-subject');
  if (item) {
    let fromText = 'Sconosciuto';
    if (item.from) fromText = item.from.displayName ? `${item.from.displayName} (${item.from.emailAddress})` : item.from.emailAddress;
    sender.innerHTML = `<strong>Da:</strong> ${escapeHtml(fromText)}`;
    subject.innerHTML = `<strong>Oggetto:</strong> ${escapeHtml(item.subject || 'Nessun oggetto')}`;
    if (item.body && item.body.getAsync) {
      item.body.getAsync(Office.CoercionType.Text, (r) => {
        if (r.status === Office.AsyncResultStatus.Succeeded) currentEmailBody = r.value;
      });
    }
  } else {
    sender.innerHTML = '<strong>Da:</strong> Mario Rossi (mario.rossi@azienda.it)';
    subject.innerHTML = '<strong>Oggetto:</strong> Richiesta preventivo e disponibilità progetto Q1';
    currentEmailBody = 'Gentile team, siamo interessati ad approfondire la proposta tecnica. Potete inviarci una stima economica e le tempistiche per la prima fase entro venerdì? Cordiali saluti, Mario Rossi.';
  }
}

function handleOutlookQuickReply(intent) {
  return runAI('Risposta email', {
    prompt: `Scrivi una risposta all'email con questo intento: "${intent}". Tono cordiale e professionale, in italiano, pronta all'invio.`,
    context: currentEmailBody || 'Nessun corpo email fornito.'
  });
}

function handleOutlookSummarize() {
  return runAI('Analisi email', {
    prompt: 'Estrai: 1. obiettivo del messaggio; 2. azioni richieste e scadenze esplicite (non inventarne); 3. decisioni da prendere.',
    context: currentEmailBody || 'Nessun corpo email fornito.'
  });
}

function handleOutlookCustomReply() {
  const prompt = document.getElementById('outlook-custom-reply').value.trim();
  if (!prompt) { notify('Scrivi le istruzioni per la risposta.', 'warn'); return; }
  return runAI('Risposta su misura', {
    prompt: `Scrivi una risposta all'email seguendo queste istruzioni dell'utente: "${prompt}"`,
    context: currentEmailBody || 'Nessun corpo email fornito.'
  });
}

/* ===================== Applicazione nel documento ===================== */

async function applyTextToDocument(text, mode = 'replace') {
  if (!text) return;

  if (currentHost === 'word' && typeof Word !== 'undefined') {
    try {
      await Word.run(async (context) => {
        const selection = context.document.getSelection();
        selection.load(['font/name', 'font/size', 'font/color', 'style']);
        await context.sync();

        const fmt = { name: selection.font.name, size: selection.font.size, color: selection.font.color, style: selection.style };

        // Un'unica scrittura: un solo passo di Annulla (Ctrl+Z) ripristina l'originale.
        const range = mode === 'replace'
          ? selection.insertText(text, Word.InsertLocation.replace)
          : selection.insertText('\n' + text, Word.InsertLocation.after);

        if (fmt.name) range.font.name = fmt.name;
        if (fmt.size) range.font.size = fmt.size;
        if (fmt.color) range.font.color = fmt.color;
        if (fmt.style) { try { range.style = fmt.style; } catch (e) {} }
        await context.sync();
      });
      notify(mode === 'replace' ? 'Testo sostituito (Ctrl+Z per annullare).' : 'Testo inserito (Ctrl+Z per annullare).', 'success');
    } catch (err) {
      notify('Errore inserimento Word: ' + err.message, 'error');
    }
  } else if (currentHost === 'excel' && typeof Excel !== 'undefined') {
    try {
      await applyToExcel(text, text);
    } catch (err) {
      notify('Errore inserimento Excel: ' + err.message, 'error');
    }
  } else if (currentHost === 'outlook' && typeof Office !== 'undefined' && Office.context && Office.context.mailbox) {
    const item = Office.context.mailbox.item;
    if (item && item.body && item.body.setSelectedDataAsync) {
      item.body.setSelectedDataAsync(text, { coercionType: Office.CoercionType.Text }, (res) => {
        notify(res.status === Office.AsyncResultStatus.Succeeded
          ? 'Testo inserito nella bozza.'
          : 'Impossibile inserire: apri la risposta/bozza e riprova, oppure usa "Copia testo".', res.status === Office.AsyncResultStatus.Succeeded ? 'success' : 'warn');
      });
    } else {
      notify('Apri una risposta o una bozza per inserire il testo, oppure usa "Copia testo".', 'warn');
    }
  } else {
    // Modalità browser: mostra cosa succederebbe
    const preview = currentHost === 'excel' && parseTable(text)
      ? `Tabella riconosciuta: ${parseTable(text).length} righe × ${parseTable(text)[0].length} colonne.`
      : `Azione "${mode}" simulata.`;
    notify(`[Modalità browser] ${preview}`, 'info');
  }
}

function escapeHtml(str) {
  return String(str).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}
