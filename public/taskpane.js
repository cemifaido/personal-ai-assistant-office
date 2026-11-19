/**
 * Personal AI Assistant — controller del pannello (Word, Excel, Outlook).
 * Le chiamate ai modelli sono in ai.js (dirette, nessun server locale).
 */

let currentHost = 'browser';
let lastGeneratedText = '';
// Tipo di risultato deciso dal PULSANTE premuto, non dal testo del modello:
// 'formula' | 'table' | 'text'. Determina come viene applicato in Excel.
let lastResultKind = 'text';
let currentAbort = null;
// Conferma di sovrascrittura legata a destinazione + contenuto: se cambiano, si richiede di nuovo.
let pendingOverwrite = null;

Office.onReady((info) => {
  const hostBadge = document.getElementById('host-badge');
  const browserTabs = document.getElementById('browser-tabs');

  buildModelSelect();

  if (info && info.host === Office.HostType.Word) {
    setHost('word', 'Word', hostBadge);
  } else if (info && info.host === Office.HostType.Excel) {
    setHost('excel', 'Excel', hostBadge);
  } else if (info && info.host === Office.HostType.PowerPoint) {
    setHost('powerpoint', 'PowerPoint', hostBadge);
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
  document.getElementById('app-version').textContent = `v${PAA_VERSION}`;
  showOnboardingIfNeeded();
});

/** Primo avvio: nessuna chiave salvata → impostazioni aperte con istruzioni. */
function showOnboardingIfNeeded() {
  const hasAnyKey = Object.keys(KEY_INPUTS).some(p => aiGetKey(p));
  if (hasAnyKey) return;
  setSettingsOpen(true);
  notify('Benvenuto! Personal AI Assistant funziona solo con una TUA chiave API personale, non inclusa: ' +
    'è legata al tuo account e al tuo contratto con il provider. Senza chiave non può elaborare nulla.\n' +
    'La più semplice: Google Gemini → "Ottieni chiave", crea la chiave, copiala qui sotto e premi "Salva chiavi".', 'info');
}

function openExternal(url) {
  url = new URL(url, window.location.href).href; // link relativi al sito del pannello
  try {
    if (typeof Office !== 'undefined' && Office.context && Office.context.ui && Office.context.ui.openBrowserWindow) {
      Office.context.ui.openBrowserWindow(url);
      return;
    }
  } catch (e) { /* ripiego sotto */ }
  window.open(url, '_blank', 'noopener');
}

function setHost(host, label, badge) {
  currentHost = host;
  badge.textContent = label;
  badge.className = `host-badge ${host}`;
  showPanel(host);
}

function showPanel(panelName) {
  ['word', 'excel', 'powerpoint', 'outlook'].forEach(p => document.getElementById(`panel-${p}`).classList.add('hidden'));
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
const MODEL_STORE_KEY = 'paa.model';

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
  const saved = storeGet(MODEL_STORE_KEY);
  if (saved && [...select.options].some(o => o.value === saved)) select.value = saved;
}

/** Se il modello scelto è di un provider senza chiave, passa al primo provider che ne ha una. */
function selectModelWithKey() {
  const { provider } = getSelectedModelInfo();
  if (aiGetKey(provider)) return;
  const withKey = AI_CATALOG.find(m => aiGetKey(m.provider));
  if (!withKey) return;
  const select = document.getElementById('select-model');
  select.value = `${withKey.provider}:${withKey.id}`;
  storeSet(MODEL_STORE_KEY, select.value);
  notify(`Modello impostato su ${withKey.label}: è il primo provider per cui hai una chiave.`, 'info');
}

function getSelectedModelInfo() {
  const select = document.getElementById('select-model');
  const [provider, model] = select.value.split(':');
  return { provider, model, label: select.options[select.selectedIndex].text };
}

/* ===================== Chiavi ===================== */

const KEY_INPUTS = { gemini: 'input-gemini-key', openai: 'input-openai-key', claude: 'input-claude-key' };

/**
 * I campi restano sempre vuoti: la chiave salvata non viene ricopiata nella pagina.
 * Il segnaposto indica solo che esiste, con le ultime 4 cifre per riconoscerla.
 */
function loadSavedKeys() {
  const PLACEHOLDERS = { gemini: 'AIzaSy...', openai: 'sk-proj-...', claude: 'sk-ant-...' };
  for (const [p, id] of Object.entries(KEY_INPUTS)) {
    const input = document.getElementById(id);
    const key = aiGetKey(p);
    input.value = '';
    input.placeholder = key ? `Chiave ${aiKeyStorage(p)} (…${key.slice(-4)}): lascia vuoto per mantenerla` : PLACEHOLDERS[p];
  }
  const persist = document.getElementById('chk-persist-keys');
  if (persist) persist.checked = aiPersistKeys();
}

function updateKeyBadges() {
  for (const p of Object.keys(KEY_INPUTS)) {
    const el = document.getElementById(`badge-${p}-status`);
    const where = aiKeyStorage(p);
    const has = Boolean(where);
    el.textContent = where === 'salvata' ? 'Salvata' : where === 'sessione' ? 'Solo sessione' : 'Mancante';
    el.className = `key-badge ${has ? 'ready' : 'missing'}`;
  }
}

async function verifyKeys() {
  const lines = [];
  for (const p of Object.keys(KEY_INPUTS)) {
    const key = aiGetKey(p).trim();
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
  const reduce = window.matchMedia && window.matchMedia('(prefers-reduced-motion: reduce)').matches;
  box.scrollIntoView({ behavior: reduce ? 'auto' : 'smooth', block: 'nearest' });
}

/** Apre o chiude il riquadro delle chiavi (ingranaggio, pulsante chiudi, dopo il salvataggio). */
function setSettingsOpen(open) {
  document.getElementById('settings-card').classList.toggle('hidden', !open);
  document.getElementById('btn-toggle-settings').setAttribute('aria-expanded', String(open));
}

/** Messaggio breve per i lettori di schermo. */
function announce(text) {
  const el = document.getElementById('sr-status');
  if (!el) return;
  el.textContent = '';
  setTimeout(() => { el.textContent = text; }, 50);
}

function setLoader(visible, text = 'Elaborazione in corso...') {
  document.getElementById('loader-text').textContent = text;
  document.getElementById('global-loader').classList.toggle('hidden', !visible);
}

function initEventListeners() {
  document.getElementById('select-model').addEventListener('change', (e) => {
    storeSet(MODEL_STORE_KEY, e.target.value);
    notify(`Modello impostato: ${getSelectedModelInfo().label}`);
  });

  document.querySelectorAll('[data-key-link]').forEach((a) => {
    a.addEventListener('click', (e) => { e.preventDefault(); openExternal(a.dataset.keyLink); });
  });
  document.getElementById('btn-clear-keys').addEventListener('click', () => {
    const failed = Object.keys(KEY_INPUTS).filter(p => !aiSetKey(p, ''));
    loadSavedKeys();
    updateKeyBadges();
    if (failed.length) {
      notify(`Non sono riuscito a cancellare dal disco la chiave di: ${failed.map(p => AI_PROVIDER_NAMES[p]).join(', ')}. ` +
             'Riprova; se il problema resta, revoca la chiave dal sito del provider.', 'error');
    } else {
      notify('Chiavi cancellate da questo pannello. Personal AI in Outlook ha impostazioni separate: cancellale anche lì. ' +
             'Cancellare una chiave non la disattiva: se temi che sia stata copiata, revocala dal sito del provider.', 'success');
    }
  });

  document.getElementById('btn-toggle-settings').addEventListener('click', () => {
    setSettingsOpen(document.getElementById('settings-card').classList.contains('hidden'));
  });

  document.getElementById('btn-save-all-keys').addEventListener('click', () => {
    // Si aggiornano solo i campi compilati: un campo vuoto mantiene la chiave gia' presente.
    let ok = aiSetPersist(document.getElementById('chk-persist-keys').checked);
    for (const [p, id] of Object.entries(KEY_INPUTS)) {
      const value = document.getElementById(id).value.trim();
      if (value && !aiSetKey(p, value)) ok = false;
    }
    loadSavedKeys();
    updateKeyBadges();
    const hasKey = Object.keys(KEY_INPUTS).some(p => aiGetKey(p));
    if (!ok) {
      notify('Attenzione: una copia precedente della chiave non è stata rimossa dal disco. Riprova o usa "Cancella".', 'error');
    } else {
      notify(!hasKey ? 'Nessuna chiave salvata: senza una chiave il pannello non può elaborare nulla.'
        : aiPersistKeys() ? "Chiavi salvate su questo PC. Per modificarle usa l'ingranaggio in alto."
        : 'Chiavi attive solo finché il pannello resta aperto: non sono state scritte su disco.', hasKey ? 'success' : 'warn');
    }
    if (hasKey && ok) setSettingsOpen(false);
    selectModelWithKey();
  });
  document.getElementById('btn-verify-keys').addEventListener('click', verifyKeys);
  document.getElementById('btn-close-settings').addEventListener('click', () => {
    setSettingsOpen(false);
    document.getElementById('btn-toggle-settings').focus();
  });

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
  document.getElementById('btn-excel-chart').addEventListener('click', handleExcelChart);

  // PowerPoint
  document.querySelectorAll('.btn-ppt-text').forEach((btn) => {
    btn.addEventListener('click', () => handlePptText(btn.dataset.action));
  });
  document.getElementById('btn-ppt-outline').addEventListener('click', handlePptOutline);
  document.getElementById('btn-ppt-improve').addEventListener('click', handlePptImprove);

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

function displayResult(result, kind) {
  const cleaned = cleanForDocument(result.text);
  lastGeneratedText = cleaned;
  lastResultKind = kind || 'text';
  pendingOverwrite = null;

  const container = document.getElementById('result-container');
  const content = document.getElementById('result-content');
  const badge = document.getElementById('result-model-badge');

  const changed = result.actualModel && !result.actualModel.startsWith(result.requestedModel);
  badge.textContent = `Risultato — ${result.actualModel || result.requestedModel}` +
    (changed ? ` (richiesto: ${result.requestedModel})` : '');

  const reduceMotion = window.matchMedia && window.matchMedia('(prefers-reduced-motion: reduce)').matches;
  container.classList.remove('hidden');
  container.scrollIntoView({ behavior: reduceMotion ? 'auto' : 'smooth' });
  // Annuncio breve per i lettori di schermo (non l'intero testo che si aggiorna).
  announce('Risultato pronto.');

  if (typingIntervalId) { clearInterval(typingIntervalId); typingIntervalId = null; }

  if (reduceMotion) {
    content.textContent = cleaned;
    return;
  }

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
async function runAI(loaderText, { prompt, context, system, kind = 'text' }) {
  const { provider, model } = getSelectedModelInfo();
  currentAbort = new AbortController();
  setLoader(true, `${loaderText} (${AI_PROVIDER_NAMES[provider]})`);
  try {
    let result = await aiCall({ provider, model, prompt, context, system, signal: currentAbort.signal });
    // Risultati strutturati: validati PRIMA di mostrarli; nel riquadro compare una descrizione leggibile.
    if (kind === 'chart') {
      lastChartSpec = parseChartSpec(result.text);
      result = { ...result, text: describeChartSpec(lastChartSpec) };
    } else if (kind === 'ppt-slides') {
      lastPptSlides = parsePptSlides(result.text);
      result = { ...result, text: describePptSlides(lastPptSlides) };
    } else if (kind === 'ppt-improve') {
      lastPptImprove = parsePptImprove(result.text, pendingPptSlide);
      result = { ...result, text: describePptImprove(lastPptImprove) };
    }
    displayResult(result, kind);
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

/** Excel in italiano? (fuori da Office, in modalità test, si assume di sì) */
function excelIsItalian() {
  const lang = (typeof Office !== 'undefined' && Office.context && Office.context.displayLanguage) || 'it-IT';
  return /^it/i.test(lang);
}

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

/**
 * Dove va la formula, in base alla selezione:
 * - se la selezione parte dall'intestazione (o è la colonna intera) si salta l'intestazione;
 * - se copre più righe, si estende fino all'ultima riga di dati (non oltre).
 */
async function getFormulaTarget(ctx) {
  const sheet = ctx.workbook.worksheets.getActiveWorksheet();
  const sel = ctx.workbook.getSelectedRange();
  const used = sheet.getUsedRangeOrNullObject(true);
  sel.load(['rowIndex', 'columnIndex', 'rowCount']);
  used.load(['rowIndex', 'rowCount']);
  await ctx.sync();

  let start = sel.rowIndex;
  let end = sel.rowIndex + sel.rowCount - 1;
  let skippedHeader = false;
  if (!used.isNullObject) {
    const usedLast = used.rowIndex + used.rowCount - 1;
    if (sel.rowIndex <= used.rowIndex) { start = used.rowIndex + 1; skippedHeader = true; }
    if (sel.rowCount > 1) end = Math.min(end, usedLast);
  }
  if (sel.rowCount === 1 || end < start) end = start;

  const first = sheet.getRangeByIndexes(start, sel.columnIndex, 1, 1);
  const full = sheet.getRangeByIndexes(start, sel.columnIndex, end - start + 1, 1);
  first.load('address');
  full.load('address');
  await ctx.sync();
  const strip = (a) => a.replace(/^.*!/, '');
  return { first, full, rows: end - start + 1, firstAddress: strip(first.address), fullAddress: strip(full.address), skippedHeader };
}

/** Struttura del foglio: nome, area dati, intestazioni con lettere di colonna, destinazione della formula. */
async function readExcelOverview() {
  return Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getActiveWorksheet();
    const used = sheet.getUsedRangeOrNullObject(true);
    sheet.load('name');
    used.load(['address', 'rowIndex', 'columnIndex', 'columnCount', 'rowCount']);
    await ctx.sync();
    const target = await getFormulaTarget(ctx);
    const dest = target.rows > 1
      ? `${target.firstAddress} (prima riga di dati); se la formula è per riga verrà estesa a ${target.fullAddress}`
      : target.firstAddress;
    if (used.isNullObject) return `Foglio "${sheet.name}" vuoto. Cella di destinazione: ${dest}.`;
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
           `Cella di destinazione della formula: ${dest}\n` +
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
      : 'Foglio: "Prova"\nArea dati: A1:C4\nCella di destinazione della formula: D2 (prima riga di dati); se la formula è per riga verrà estesa a D2:D4\nColonne:\nA = "Mese" — testo\nB = "Vendite" — numero\nC = "Costi" — numero';
  } catch (err) {
    notify(excelErrorMessage(err), 'error');
    return;
  }
  return runAI('Generazione formula', {
    prompt: `L'utente vuole una formula Excel per: "${prompt}".
Regole:
- Usa le colonne reali indicate nel contesto. La formula è scritta nella cella di destinazione indicata: se va ripetuta riga per riga, riferisciti alla STESSA RIGA di quella cella (es. destinazione C2 → B2); verrà estesa alle righe sotto adattando i riferimenti.
- Per un calcolo unico su tutta la colonna (media, totale, conteggio) usa l'intervallo dei dati senza intestazione (es. E2:E120).
- Rispetta il tipo indicato: una "data vera" è un numero seriale, NON va scomposta con funzioni di testo (LEFT/MID/RIGHT); un "testo" sì.
${excelIsItalian()
  ? `- Scrivi la formula in ITALIANO, esattamente come la si digita in Excel italiano: nomi delle funzioni italiani (es. SOMMA, MEDIA, SE, CERCA.X, TESTO, DATA, STRINGA.ESTRAI) e separatore argomenti punto e virgola ";". Separatore decimale virgola.
- Nei codici formato di TESTO usa i codici italiani: g=giorno, m=mese, a=anno (es. "mm/gg/aaaa").`
  : `- Write the formula in English with comma as argument separator.`}
- Se il risultato si ottiene meglio senza formula (es. cambiare il formato delle celle), dillo nella spiegazione.
Formato della risposta:
riga 1: "FORMULA: =..."
riga 2: "TIPO: RIGA" se la formula va ripetuta per ogni riga di dati, oppure "TIPO: UNICA" se è un solo calcolo (totale, media, conteggio...)
poi una spiegazione di massimo 3 righe, citando le colonne usate.`,
    context: overview,
    system: 'Sei un esperto di Microsoft Excel e formule avanzate.',
    kind: 'formula'
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
    system: 'Restituisci solo dati tabellari separati da TAB.',
    kind: 'table'
  });
}

/* ---------- Grafici ---------- */

const CHART_TYPES = {
  colonne: 'ColumnClustered', barre: 'BarClustered', linee: 'Line', torta: 'Pie',
  anello: 'Doughnut', area: 'Area', dispersione: 'XYScatter'
};
const CHART_MODES = ['intervalli', 'conteggio', 'somma', 'media', 'istogramma'];
const CHART_MAX_BINS = 50;
const CHART_DATA_SHEET = 'Dati grafici PAI';
const CHART_MAX_POINTS = 2000;
const CHART_MAX_GROUPS = 20; // oltre, le categorie minori finiscono in "Altri"
let lastChartSpec = null;

async function handleExcelChart() {
  const request = document.getElementById('excel-chart-prompt').value.trim() ||
    'Crea il grafico più adatto ai dati selezionati.';
  let context;
  try {
    if (typeof Excel !== 'undefined') {
      const overview = await readExcelOverview();
      const selection = await Excel.run(async (ctx) => {
        const sel = ctx.workbook.getSelectedRange();
        sel.load('address');
        await ctx.sync();
        return sel.address.replace(/^.*!/, '');
      });
      context = `${overview}\nSelezione attuale: ${selection}`;
    } else {
      context = 'Foglio: "Prova"\nArea dati: A1:C4\nColonne:\nA = "Mese" — testo\nB = "Vendite" — numero\nC = "Costi" — numero\nSelezione attuale: A1:C4';
    }
  } catch (err) {
    notify(excelErrorMessage(err), 'error');
    return;
  }
  return runAI('Proposta di grafico', {
    prompt: `L'utente vuole un grafico Excel: "${request}".
Rispondi SOLO con un oggetto JSON valido, senza altro testo, con questi campi:
{"tipo": uno tra ${Object.keys(CHART_TYPES).map(t => `"${t}"`).join(', ')},
 "titolo": "titolo breve",
 "modo": "intervalli" se i valori da disegnare sono già nelle celle; "conteggio" per contare quante righe ci sono per ogni categoria (testo); "somma" o "media" per aggregare una colonna numerica per categoria; "istogramma" per raggruppare UNA colonna numerica in fasce (es. punteggi 0-4, 5-9...) e contare quante righe cadono in ogni fascia,
 "categorie": intervallo della colonna delle etichette SENZA intestazione (es. "A2:A41"), oppure "" se non servono,
 "valori": elenco di intervalli numerici SENZA intestazione (es. ["F2:F41"]); per "conteggio" usa []; per "istogramma" una sola colonna,
 "ampiezza": solo per "istogramma", larghezza di ogni fascia (numero, es. 5); se l'utente non la indica, ometti il campo,
 "inizio": solo per "istogramma", valore da cui parte la prima fascia (es. 0); facoltativo,
 "nomiSerie": elenco dei nomi delle serie, uno per ogni intervallo di "valori" (per conteggio/somma/media un solo nome),
 "spiegazione": "una frase su cosa mostra il grafico"}
Regole: ogni intervallo è UNA sola colonna, con le stesse righe delle categorie; usa le colonne reali del contesto e i dati della selezione se pertinenti; "torta" e "anello" con una sola serie.`,
    context,
    system: 'Sei un esperto di Microsoft Excel e di visualizzazione dei dati. Rispondi solo con JSON.',
    kind: 'chart'
  });
}

/** Valida la proposta del modello. Restituisce lo spec normalizzato oppure lancia un errore leggibile. */
function parseChartSpec(text) {
  const m = String(text).match(/\{[\s\S]*\}/);
  if (!m) throw new Error('Il modello non ha proposto un grafico valido: riprova descrivendolo meglio.');
  let s;
  try { s = JSON.parse(m[0]); } catch (e) { throw new Error('Proposta di grafico non leggibile: riprova.'); }
  const cell = '\\$?[A-Z]{1,3}\\$?\\d{1,7}';
  const rangeRe = new RegExp(`^${cell}:${cell}$`, 'i');
  const tipo = String(s.tipo || '').toLowerCase();
  const modo = String(s.modo || 'intervalli').toLowerCase();
  if (!CHART_TYPES[tipo]) throw new Error(`Tipo di grafico non supportato: ${s.tipo}`);
  if (!CHART_MODES.includes(modo)) throw new Error(`Modo non supportato: ${s.modo}`);
  const categorie = String(s.categorie || '').trim();
  const valori = (Array.isArray(s.valori) ? s.valori : [s.valori]).filter(Boolean).map(v => String(v).trim());
  if (categorie && !rangeRe.test(categorie)) throw new Error(`Intervallo categorie non valido: ${categorie}`);
  for (const v of valori) if (!rangeRe.test(v)) throw new Error(`Intervallo valori non valido: ${v}`);
  if (modo === 'intervalli' && valori.length === 0) throw new Error('La proposta non indica i valori da disegnare.');
  if (modo === 'istogramma') {
    if (valori.length !== 1) throw new Error('Per un istogramma serve una sola colonna di numeri: indica quale (es. "istogramma del punteggio finale").');
  } else if (modo !== 'intervalli' && !categorie) {
    throw new Error('Per contare o aggregare per categoria serve una colonna di etichette (es. i luoghi di nascita). ' +
                    'Se vuoi raggruppare dei numeri in fasce, chiedi un istogramma (es. "istogramma del punteggio finale a fasce di 5").');
  }
  const ampiezza = Number(s.ampiezza);
  const inizio = Number(s.inizio);
  if ((modo === 'somma' || modo === 'media') && valori.length !== 1) throw new Error('Per somma o media serve una sola colonna di valori.');
  if ((tipo === 'torta' || tipo === 'anello') && valori.length > 1) throw new Error('Torta e anello ammettono una sola serie.');
  const nomi = Array.isArray(s.nomiSerie) ? s.nomiSerie.map(String) : [];
  return {
    tipo: modo === 'istogramma' && (tipo === 'torta' || tipo === 'anello' || tipo === 'dispersione') ? 'colonne' : tipo,
    modo, categorie: modo === 'istogramma' ? '' : categorie, valori: modo === 'conteggio' ? [] : valori,
    ampiezza: modo === 'istogramma' && ampiezza > 0 ? ampiezza : null,
    inizio: modo === 'istogramma' && Number.isFinite(inizio) && s.inizio !== undefined && s.inizio !== '' ? inizio : null,
    titolo: String(s.titolo || '').slice(0, 120),
    nomiSerie: nomi,
    spiegazione: String(s.spiegazione || '')
  };
}

function describeChartSpec(s) {
  const modi = { intervalli: 'valori presi dalle celle', conteggio: 'conteggio delle righe per categoria',
                 somma: 'somma per categoria', media: 'media per categoria',
                 istogramma: `conteggio per fasce${s.ampiezza ? ` di ampiezza ${s.ampiezza}` : ' (ampiezza automatica)'}` };
  return `Grafico proposto: ${s.tipo} — "${s.titolo}"\n` +
    `Dati: ${modi[s.modo]}` + (s.categorie ? `, categorie ${s.categorie}` : '') +
    (s.valori.length ? `, valori ${s.valori.join(', ')}` : '') + '\n' +
    (s.modo !== 'intervalli' ? `I numeri vengono calcolati dal pannello e scritti nel foglio "${CHART_DATA_SHEET}".\n` : '') +
    (s.spiegazione ? `\n${s.spiegazione}\n` : '') +
    '\nPremi "Inserisci" per creare il grafico.';
}

/**
 * Istogramma esatto: fasce di ampiezza fissa (scelta dal modello o automatica), fasce vuote incluse.
 * Numeri interi: fasce chiuse "0–4", "5–9"; decimali: fasce "da a fino a <b".
 */
function buildHistogram(values, width, start) {
  const nums = values.filter(v => typeof v === 'number' && Number.isFinite(v));
  if (!nums.length) throw new Error("La colonna scelta per l'istogramma non contiene numeri.");
  const min = Math.min(...nums), max = Math.max(...nums);
  if (!(width > 0)) {
    // ampiezza automatica: circa radice(n) fasce, arrotondata a un numero "comodo"
    const raw = (max - min || 1) / Math.max(1, Math.ceil(Math.sqrt(nums.length)));
    const pow = Math.pow(10, Math.floor(Math.log10(raw)));
    width = [1, 2, 5, 10].map(m => m * pow).find(c => c >= raw) || 10 * pow;
  }
  if (!Number.isFinite(start)) start = Math.floor(min / width) * width;
  start = Math.min(start, min);
  const bins = Math.floor((max - start) / width) + 1;
  if (bins > CHART_MAX_BINS) throw new Error(`Con questa ampiezza servirebbero ${bins} fasce: chiedi fasce più larghe.`);
  const integers = nums.every(Number.isInteger) && Number.isInteger(width) && Number.isInteger(start);
  const counts = new Array(bins).fill(0);
  nums.forEach(v => { counts[Math.min(bins - 1, Math.floor((v - start) / width))]++; });
  const fmt = (x) => Number.isInteger(x) ? String(x) : x.toFixed(2).replace(/\.?0+$/, '');
  const rows = counts.map((c, i) => {
    const a = start + i * width;
    const label = integers ? `${fmt(a)}–${fmt(a + width - 1)}` : `${fmt(a)}–<${fmt(a + width)}`;
    return [label, c];
  });
  // Le fasce vuote prima del primo dato non aggiungono informazione: si saltano.
  const first = rows.findIndex(r => r[1] > 0);
  return rows.slice(first);
}

async function insertChart(spec) {
  await Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getActiveWorksheet();
    const used = sheet.getUsedRange(true);
    used.load(['rowIndex', 'columnIndex', 'columnCount']);
    sheet.load('name');

    let catRange = spec.categorie ? sheet.getRange(spec.categorie) : null;
    let valRanges = spec.valori.map(a => sheet.getRange(a));
    if (catRange) catRange.load(['rowCount', 'columnCount', 'values']);
    valRanges.forEach(r => r.load(['rowCount', 'columnCount', 'values']));
    await ctx.sync();

    const all = (catRange ? [catRange] : []).concat(valRanges);
    if (all.some(r => r.columnCount !== 1)) throw new Error('Ogni intervallo del grafico deve essere una sola colonna.');
    const rows = all[0].rowCount;
    if (all.some(r => r.rowCount !== rows)) throw new Error('Categorie e valori hanno un numero di righe diverso.');
    if (rows > CHART_MAX_POINTS) throw new Error(`Troppi punti (${rows}): seleziona meno righe o chiedi un aggregato.`);

    let note = '';
    if (spec.modo !== 'intervalli') {
      // Aggregazione esatta fatta qui, non dal modello.
      let table;
      if (spec.modo === 'istogramma') {
        table = buildHistogram(valRanges[0].values.map(r => r[0]), spec.ampiezza, spec.inizio);
      } else {
      const groups = new Map();
      catRange.values.forEach((r, i) => {
        const k = String(r[0] ?? '').trim() || '(vuoto)';
        const v = spec.modo === 'conteggio' ? 1 : valRanges[0].values[i][0];
        if (spec.modo !== 'conteggio' && typeof v !== 'number') return;
        const g = groups.get(k) || { sum: 0, n: 0 };
        g.sum += v; g.n += 1; groups.set(k, g);
      });
      table = [...groups.entries()].map(([k, g]) => [k, spec.modo === 'media' ? g.sum / g.n : g.sum]);
      table.sort((a, b) => b[1] - a[1]);
      if (table.length > CHART_MAX_GROUPS && spec.modo !== 'media') {
        const rest = table.slice(CHART_MAX_GROUPS - 1).reduce((s, r) => s + r[1], 0);
        table = table.slice(0, CHART_MAX_GROUPS - 1).concat([['Altri', rest]]);
        note = ` Le categorie oltre le prime ${CHART_MAX_GROUPS - 1} sono raggruppate in "Altri".`;
      }
      }
      if (!table.length) throw new Error('Nessun dato numerico da aggregare.');

      let dataSheet = ctx.workbook.worksheets.getItemOrNullObject(CHART_DATA_SHEET);
      await ctx.sync();
      let start = 0;
      if (dataSheet.isNullObject) {
        dataSheet = ctx.workbook.worksheets.add(CHART_DATA_SHEET); // foglio nuovo: si parte dalla riga 1
      } else {
        // Prima si verifica se c'e' un'area usata, poi se ne leggono le proprieta'.
        const dUsed = dataSheet.getUsedRangeOrNullObject(true);
        await ctx.sync();
        if (!dUsed.isNullObject) {
          dUsed.load(['rowIndex', 'rowCount']);
          await ctx.sync();
          start = dUsed.rowIndex + dUsed.rowCount + 1;
        }
      }
      const label = spec.nomiSerie[0] || (spec.modo === 'conteggio' || spec.modo === 'istogramma' ? 'Conteggio' : spec.modo === 'somma' ? 'Somma' : 'Media');
      dataSheet.getRangeByIndexes(start, 0, 1, 2).values = [[spec.titolo || 'Grafico', label]];
      dataSheet.getRangeByIndexes(start + 1, 0, table.length, 2).values = table;
      catRange = dataSheet.getRangeByIndexes(start + 1, 0, table.length, 1);
      valRanges = [dataSheet.getRangeByIndexes(start + 1, 1, table.length, 1)];
      spec = { ...spec, nomiSerie: [label] };
      note = ` Dati calcolati nel foglio "${CHART_DATA_SHEET}".` + note;
    }

    const chart = sheet.charts.add(CHART_TYPES[spec.tipo], valRanges[0], Excel.ChartSeriesBy.columns);
    chart.title.text = spec.titolo || '';
    const first = chart.series.getItemAt(0);
    first.name = spec.nomiSerie[0] || '';
    if (catRange) first.setXAxisValues(catRange);
    for (let i = 1; i < valRanges.length; i++) {
      const s = chart.series.add(spec.nomiSerie[i] || `Serie ${i + 1}`);
      s.setValues(valRanges[i]);
      if (catRange) s.setXAxisValues(catRange);
    }
    if (spec.tipo === 'torta' || spec.tipo === 'anello') chart.dataLabels.showPercentage = true;
    // A destra dei dati, senza coprirli.
    chart.setPosition(`${colLetter(used.columnIndex + used.columnCount + 1)}${used.rowIndex + 2}`);
    await ctx.sync();
    notify(`Grafico inserito nel foglio "${sheet.name}".${note}`, 'success');
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

/**
 * Conferma di sovrascrittura: vale solo per la STESSA destinazione e lo STESSO contenuto
 * per cui è stata chiesta. Se cambia la selezione o il risultato, si chiede di nuovo.
 * Restituisce true se si può scrivere.
 */
function confirmOverwrite(dest, content) {
  if (!dest.occupied) { pendingOverwrite = null; return true; }
  // Il token lega la conferma a: foglio (id), indirizzo completo, contenuto attuale della
  // destinazione, tipo di risultato e testo da scrivere. Qualunque cambiamento richiede un nuovo clic.
  const token = [dest.sheetId, dest.address, dest.snapshot, lastResultKind, content].join('\u0000');
  if (pendingOverwrite === token) { pendingOverwrite = null; return true; }
  pendingOverwrite = token;
  notify(`${dest.label} contiene già dati. Premi di nuovo lo stesso pulsante per sovrascrivere proprio quest'area, oppure cambia selezione.`, 'warn');
  return false;
}

/** Legge lo stato di un intervallo di destinazione: occupato anche se contiene solo formule (es. =""). */
async function describeDestination(context, range) {
  range.load(['address', 'values', 'formulas']);
  range.worksheet.load(['id', 'name']);
  await context.sync();
  const occupied = range.values.some(r => r.some(v => v !== '' && v !== null)) ||
                   range.formulas.some(r => r.some(f => typeof f === 'string' && f.startsWith('=')));
  return {
    occupied,
    sheetId: range.worksheet.id,
    address: range.address,                       // qualificato con il nome del foglio
    label: `${range.worksheet.name}!${range.address.replace(/^.*!/, '')}`,
    snapshot: JSON.stringify(range.formulas)
  };
}

/** Applica il risultato in Excel secondo il tipo deciso dal pulsante che l'ha generato. */
async function applyToExcel(text) {
  if (lastResultKind === 'chart') {
    if (!lastChartSpec) { notify('Nessun grafico proposto: premi "Proponi Grafico".', 'warn'); return; }
    await insertChart(lastChartSpec);
    return;
  }
  await Excel.run(async (context) => {
    if (lastResultKind === 'formula') {
      const formulaMatch = text.match(/FORMULA:\s*`*(=[^\r\n`]+)/i);
      if (!formulaMatch) {
        notify('Il risultato non contiene una riga "FORMULA: =...". Rigenera la formula.', 'warn');
        return;
      }
      const f = formulaMatch[1].trim();
      const perRow = /TIPO:\s*RIGA/i.test(text);
      const t = await getFormulaTarget(context);
      const dest = perRow && t.rows > 1 ? t.full : t.first;

      if (!confirmOverwrite(await describeDestination(context, dest), f)) return;

      // Formula italiana → formulasLocal (come digitata in Excel); altrimenti formulas (inglese).
      if (excelIsItalian()) t.first.formulasLocal = [[f]]; else t.first.formulas = [[f]];
      if (dest === t.full) {
        // Estensione come il trascinamento: in R1C1 i riferimenti relativi si adattano da soli.
        t.first.load('formulasR1C1');
        await context.sync();
        const r1c1 = t.first.formulasR1C1[0][0];
        t.full.formulasR1C1 = Array.from({ length: t.rows }, () => [r1c1]);
      }
      await context.sync();

      let msg = dest === t.full
        ? `Formula inserita in ${t.fullAddress} (${t.rows} righe).`
        : `Formula inserita in ${t.firstAddress}.`;
      if (t.skippedHeader) msg += ' Intestazione saltata.';
      if (!perRow && t.rows > 1) msg += ' È un calcolo unico: non è stata ripetuta sulle altre righe.';
      msg += '\nSe compare #NOME?, un nome di funzione non è stato riconosciuto: rigenera o correggilo.';
      notify(msg, 'success');
      return;
    }

    const start = context.workbook.getSelectedRange().getCell(0, 0);

    let matrix;
    if (lastResultKind === 'table') {
      matrix = parseTable(text);
      if (!matrix) { notify('Il risultato non è una tabella separata da TAB: rigenera i dati.', 'warn'); return; }
    } else {
      // Testo libero (es. un'analisi): una sola cella, mai interpretato come formula.
      matrix = [[text.startsWith('=') ? "'" + text : text]];
    }
    const target = start.getResizedRange(matrix.length - 1, matrix[0].length - 1);
    if (!confirmOverwrite(await describeDestination(context, target), text)) return;
    target.values = matrix;
    await context.sync();
    notify(`Inseriti ${matrix.length}×${matrix[0].length} valori in ${target.address.replace(/^.*!/, '')}.`, 'success');
  });
}

/* ===================== PowerPoint ===================== */

const PPT_MAX_SLIDES = 15;
const PPT_MAX_POINTS = 8;
let lastPptSlides = null;     // proposta di nuove slide validata
let pendingPptSlide = null;   // testi letti dalla slide selezionata, in attesa della proposta
let lastPptImprove = null;    // proposta di miglioramento validata

function pptSupports(version) {
  try { return Office.context.requirements.isSetSupported('PowerPointApi', version); } catch (e) { return false; }
}

function pptGetSelectedText() {
  if (typeof PowerPoint === 'undefined') {
    return Promise.resolve('Il nostro servizio offre soluzioni innovative per migliorare la produttività dei team commerciali.');
  }
  return new Promise((resolve, reject) => {
    Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, (r) => {
      if (r.status === Office.AsyncResultStatus.Succeeded) resolve(String(r.value || '').trim());
      else reject(new Error(r.error ? r.error.message : 'selezione non leggibile'));
    });
  });
}

function pptSetSelectedText(text) {
  return new Promise((resolve, reject) => {
    Office.context.document.setSelectedDataAsync(text, { coercionType: Office.CoercionType.Text }, (r) => {
      if (r.status === Office.AsyncResultStatus.Succeeded) resolve();
      else reject(new Error(r.error ? r.error.message : 'impossibile scrivere nella selezione'));
    });
  });
}

const PPT_TEXT_ACTIONS = {
  riscrivi: { label: 'Riscrittura', prompt: 'Riscrivi il testo per una slide: chiaro, incisivo e professionale. Mantieni la lingua e la struttura a punti se presente.' },
  correggi: { label: 'Correzione', prompt: 'Correggi errori grammaticali, ortografici e di punteggiatura senza cambiare stile né contenuto. Se è corretto, restituiscilo identico.' },
  accorcia: { label: 'Sintesi', prompt: 'Accorcia il testo di circa la metà, adatto a una slide: frasi brevi, niente ripetizioni, mantieni i punti essenziali.' },
  traduci: { label: 'Traduzione', prompt: 'Traduci il testo in inglese in modo professionale, adatto a una slide.' }
};

async function handlePptText(action) {
  const a = PPT_TEXT_ACTIONS[action];
  let text;
  try { text = await pptGetSelectedText(); } catch (err) { notify('Errore PowerPoint: ' + err.message, 'error'); return; }
  if (!text) { notify('Seleziona prima del testo dentro una casella della slide.', 'warn'); return; }
  return runAI(a.label, {
    prompt: a.prompt, context: text,
    system: `Sei un esperto di presentazioni PowerPoint. ${WORD_ONLY_TEXT}`
  });
}

/** Elenco dei titoli già presenti, per dare coerenza alle nuove slide (al massimo 30). */
async function pptReadOutline() {
  if (typeof PowerPoint === 'undefined' || !pptSupports('1.3')) return '';
  return PowerPoint.run(async (ctx) => {
    const slides = ctx.presentation.slides;
    slides.load('items');
    await ctx.sync();
    const first = slides.items.slice(0, 30).map(sl => { const sh = sl.shapes; sh.load('items/type'); return sh; });
    await ctx.sync();
    const titles = first.map(sh => {
      const t = sh.items.find(x => ['Placeholder', 'TextBox', 'GeometricShape'].includes(x.type));
      if (t) t.textFrame.textRange.load('text');
      return t;
    });
    await ctx.sync().catch(() => {});
    const lines = titles.map((t, i) => `${i + 1}. ${t ? (t.textFrame.textRange.text || '').split('\n')[0].slice(0, 80) : '(senza testo)'}`);
    return `Presentazione attuale: ${slides.items.length} slide.\n${lines.join('\n')}`;
  });
}

const PPT_LAYOUTS = ['titolo', 'sezione', 'contenuto', 'confronto', 'numeri', 'passi', 'tabella'];

async function handlePptOutline() {
  const request = document.getElementById('ppt-outline-prompt').value.trim();
  if (!request) { notify('Descrivi le slide da creare.', 'warn'); return; }
  if (typeof PowerPoint !== 'undefined' && !pptSupports('1.4')) {
    notify('Questa versione di PowerPoint non permette di creare slide dai componenti aggiuntivi (serve PowerPointApi 1.4).', 'error');
    return;
  }
  let outline = '';
  try { outline = await pptReadOutline(); } catch (e) { outline = ''; }
  return runAI('Proposta di slide', {
    prompt: `L'utente vuole aggiungere slide a una presentazione: "${request}".
Rispondi SOLO con un oggetto JSON valido, senza altro testo:
{"slides": [ { "layout": "...", "titolo": "...", ...campi del layout... } ]}
Layout disponibili (scegli il piu' adatto al contenuto di ogni slide, variando quando ha senso):
- "titolo": copertina. Campi: "titolo", "sottotitolo". Solo come prima slide e solo se la presentazione e' vuota o l'utente chiede una presentazione nuova.
- "sezione": separatore di capitolo. Campi: "titolo", "sottotitolo" (facoltativo).
- "contenuto": elenco puntato. Campi: "titolo", "punti" (2-${PPT_MAX_POINTS} punti brevi).
- "confronto": due colonne affiancate (pro/contro, prima/dopo, opzione A/B). Campi: "titolo", "colonne": [{"titolo": "...", "punti": [...]}, {"titolo": "...", "punti": [...]}].
- "numeri": 2-4 dati chiave in evidenza. Campi: "titolo", "numeri": [{"valore": "+18%", "etichetta": "crescita vendite"}]. USALO SOLO con numeri presenti nella richiesta dell'utente o nel contesto: non inventare dati.
- "passi": processo o fasi in sequenza (3-6 passi). Campi: "titolo", "passi": ["fase 1", "fase 2", ...] (massimo 6 parole per passo).
- "tabella": confronto strutturato. Campi: "titolo", "tabella": {"intestazioni": [...], "righe": [[...], ...]} (massimo 5 colonne e 6 righe, testi brevi).
Regole: al massimo ${PPT_MAX_SLIDES} slide; testi brevi, senza simboli di elenco iniziali; lingua della richiesta;
coerenza con le slide gia' presenti se indicate nel contesto, senza ripeterle.`,
    context: outline,
    system: 'Sei un esperto di presentazioni aziendali. Rispondi solo con JSON.',
    kind: 'ppt-slides'
  });
}

function extractJson(text) {
  const m = String(text).match(/\{[\s\S]*\}/);
  if (!m) throw new Error('Il modello non ha restituito una proposta valida: riprova.');
  try { return JSON.parse(m[0]); } catch (e) { throw new Error('Proposta non leggibile: riprova.'); }
}

const pptText = (v, max) => String(v == null ? '' : v).replace(/^\s*[-•*]\s*/, '').trim().slice(0, max);
const pptList = (a, n, max) => (Array.isArray(a) ? a : []).map(x => pptText(x, max)).filter(Boolean).slice(0, n);

/** Valida e normalizza la proposta; una slide grafica incompleta ricade su "contenuto". */
function parsePptSlides(text) {
  const data = extractJson(text);
  const slides = Array.isArray(data.slides) ? data.slides : [];
  if (!slides.length) throw new Error('La proposta non contiene slide.');
  return slides.slice(0, PPT_MAX_SLIDES).map((sl) => {
    let layout = PPT_LAYOUTS.includes(sl.layout) ? sl.layout : 'contenuto';
    const out = { titolo: pptText(sl.titolo, 150), sottotitolo: pptText(sl.sottotitolo, 200),
                  punti: pptList(sl.punti, PPT_MAX_POINTS, 200) };
    if (layout === 'confronto') {
      out.colonne = (Array.isArray(sl.colonne) ? sl.colonne : []).slice(0, 2)
        .map(c => ({ titolo: pptText(c && c.titolo, 80), punti: pptList(c && c.punti, 6, 150) }));
      if (out.colonne.length !== 2) layout = 'contenuto';
    } else if (layout === 'numeri') {
      out.numeri = (Array.isArray(sl.numeri) ? sl.numeri : []).slice(0, 4)
        .map(n => ({ valore: pptText(n && n.valore, 16), etichetta: pptText(n && n.etichetta, 60) })).filter(n => n.valore);
      if (out.numeri.length < 2) layout = 'contenuto';
    } else if (layout === 'passi') {
      out.passi = pptList(sl.passi, 6, 60);
      if (out.passi.length < 2) layout = 'contenuto';
    } else if (layout === 'tabella') {
      const t = sl.tabella || {};
      const head = pptList(t.intestazioni, 5, 40);
      const rows = (Array.isArray(t.righe) ? t.righe : []).slice(0, 6)
        .map(r => (Array.isArray(r) ? r : []).slice(0, head.length).map(c => pptText(c, 60)));
      out.tabella = { intestazioni: head, righe: rows.map(r => r.concat(Array(Math.max(0, head.length - r.length)).fill(''))) };
      if (head.length < 2 || !rows.length) layout = 'contenuto';
    }
    out.layout = layout;
    if (layout === 'contenuto' && !out.punti.length) {
      // contenuto di riserva dai campi grafici non validi
      out.punti = [].concat(out.passi || [], (out.numeri || []).map(n => `${n.valore} ${n.etichetta}`),
        ...(out.colonne || []).map(c => c.punti)).slice(0, PPT_MAX_POINTS);
    }
    return out;
  }).filter(sl => sl.titolo || sl.punti.length || sl.layout !== 'contenuto');
}

const PPT_LAYOUT_NAMES = { titolo: 'copertina', sezione: 'sezione', contenuto: 'elenco', confronto: 'confronto',
                           numeri: 'numeri chiave', passi: 'processo a passi', tabella: 'tabella' };

function describePptSlides(slides) {
  const body = (sl) => {
    switch (sl.layout) {
      case 'titolo': case 'sezione': return sl.sottotitolo ? `   ${sl.sottotitolo}` : '';
      case 'confronto': return sl.colonne.map(c => `   [${c.titolo}] ${c.punti.join('; ')}`).join('\n');
      case 'numeri': return sl.numeri.map(n => `   ${n.valore} — ${n.etichetta}`).join('\n');
      case 'passi': return `   ${sl.passi.join(' → ')}`;
      case 'tabella': return `   ${sl.tabella.intestazioni.join(' | ')} (${sl.tabella.righe.length} righe)`;
      default: return sl.punti.map(p => `   • ${p}`).join('\n');
    }
  };
  return `Slide proposte (${slides.length}), da aggiungere in fondo alla presentazione:\n\n` +
    slides.map((sl, i) => `${i + 1}. [${PPT_LAYOUT_NAMES[sl.layout]}] ${sl.titolo}\n${body(sl)}`).join('\n\n') +
    '\n\nPremi "Inserisci" per creare le slide.';
}

/** Legge i testi della prima slide selezionata (forme con testo), con gli identificativi per riscriverli. */
async function pptReadSelectedSlide() {
  if (typeof PowerPoint === 'undefined') {
    return { slideId: 'prova', shapes: [{ id: '1', text: 'Risultati del trimestre' }, { id: '2', text: 'Vendite in crescita\nNuovi clienti nel nord Italia' }] };
  }
  return PowerPoint.run(async (ctx) => {
    const sel = ctx.presentation.getSelectedSlides();
    sel.load('items/id');
    await ctx.sync();
    if (!sel.items.length) throw new Error('Seleziona una slide nel riquadro delle miniature.');
    const slide = sel.items[0];
    const shapes = slide.shapes;
    shapes.load('items/id,items/type,items/name');
    await ctx.sync();
    const withText = shapes.items.filter(s => ['Placeholder', 'TextBox', 'GeometricShape'].includes(s.type));
    withText.forEach(s => s.textFrame.textRange.load('text'));
    await ctx.sync();
    const list = withText.map(s => ({ id: s.id, text: s.textFrame.textRange.text || '' })).filter(s => s.text.trim());
    if (!list.length) throw new Error('La slide selezionata non contiene testi modificabili.');
    return { slideId: slide.id, shapes: list };
  });
}

async function handlePptImprove() {
  if (typeof PowerPoint !== 'undefined' && !pptSupports('1.5')) {
    notify('Questa versione di PowerPoint non permette di leggere la slide selezionata (serve PowerPointApi 1.5).', 'error');
    return;
  }
  try { pendingPptSlide = await pptReadSelectedSlide(); }
  catch (err) { notify('Errore PowerPoint: ' + err.message, 'error'); return; }
  const wish = document.getElementById('ppt-improve-prompt').value.trim();
  const numbered = pendingPptSlide.shapes.map((s, i) => `[${i + 1}]\n${s.text}`).join('\n\n');
  return runAI('Miglioramento della slide', {
    prompt: `Migliora i testi di questa slide${wish ? ` con questa indicazione: "${wish}"` : ': più chiari, sintetici e incisivi'}.
I testi sono numerati tra parentesi quadre, uno per casella. Rispondi SOLO con JSON valido:
{"testi": ["nuovo testo casella 1", "nuovo testo casella 2", "..."]}
Regole: esattamente ${pendingPptSlide.shapes.length} testi, nello stesso ordine; mantieni gli a capo dove servono gli elenchi; non aggiungere numerazioni o simboli.`,
    context: numbered,
    system: 'Sei un esperto di presentazioni aziendali. Rispondi solo con JSON.',
    kind: 'ppt-improve'
  });
}

function parsePptImprove(text, slide) {
  if (!slide) throw new Error('Slide non disponibile: riprova.');
  const data = extractJson(text);
  const testi = Array.isArray(data.testi) ? data.testi.map(String) : [];
  if (testi.length !== slide.shapes.length) {
    throw new Error(`La proposta ha ${testi.length} testi invece di ${slide.shapes.length}: riprova.`);
  }
  return { slideId: slide.slideId, shapes: slide.shapes.map((s, i) => ({ id: s.id, before: s.text, after: testi[i].trim() })) };
}

function describePptImprove(p) {
  return 'Testi proposti per la slide selezionata:\n\n' +
    p.shapes.map((s, i) => `Casella ${i + 1}:\n${s.after}`).join('\n\n') +
    '\n\nPremi "Sostituisci" per applicarli (Ctrl+Z per annullare).';
}

// Colori degli elementi grafici disegnati dal pannello (sobri, leggibili sui temi chiari comuni).
const PPT_COLORS = { accent: '#0369A1', accentText: '#FFFFFF', soft: '#E0F2FE', softText: '#0C4A6E',
                     band: '#F1F5F9', bandText: '#0F172A' };

/** Layout del tema per nome (italiano o inglese); null se il tema non lo ha. */
function pptFindLayout(layouts, kind) {
  const names = {
    titolo: /^(diapositiva titolo|title slide)$/i,
    sezione: /^(intestazione sezione|section header)$/i,
    contenuto: /^(titolo e contenuto|title and content)$/i,
    confronto: /^(due contenuti|two content)$/i,
    soloTitolo: /^(solo titolo|title only)$/i
  };
  const l = layouts.find(x => names[kind].test((x.name || '').trim()));
  return l ? l.id : null;
}

function pptStyleShape(shape, fill, color, size, bold, align) {
  shape.fill.setSolidColor(fill);
  shape.lineFormat.visible = false;
  const tr = shape.textFrame.textRange;
  tr.font.color = color;
  tr.font.size = size;
  tr.font.bold = !!bold;
  tr.paragraphFormat.horizontalAlignment = align || 'Center';
  shape.textFrame.verticalAlignment = 'Middle';
  shape.textFrame.wordWrap = true;
}

/** Disegna gli elementi grafici nell'area sotto il titolo. */
function pptDrawGraphics(slide, spec, area) {
  const G = PowerPoint.GeometricShapeType;
  if (spec.layout === 'numeri') {
    const n = spec.numeri.length, gap = 20, w = (area.width - gap * (n - 1)) / n, h = Math.min(170, area.height);
    spec.numeri.forEach((it, i) => {
      const left = area.left + i * (w + gap), top = area.top + (area.height - h) / 2;
      const card = slide.shapes.addGeometricShape(G.roundRectangle, { left, top, width: w, height: h });
      pptStyleShape(card, PPT_COLORS.soft, PPT_COLORS.softText, 12);
      const val = slide.shapes.addTextBox(it.valore, { left, top: top + 18, width: w, height: 70 });
      val.textFrame.textRange.font.size = 40;
      val.textFrame.textRange.font.bold = true;
      val.textFrame.textRange.font.color = PPT_COLORS.accent;
      val.textFrame.textRange.paragraphFormat.horizontalAlignment = 'Center';
      const lab = slide.shapes.addTextBox(it.etichetta, { left: left + 10, top: top + 92, width: w - 20, height: h - 100 });
      lab.textFrame.textRange.font.size = 16;
      lab.textFrame.textRange.font.color = PPT_COLORS.bandText;
      lab.textFrame.textRange.paragraphFormat.horizontalAlignment = 'Center';
      lab.textFrame.wordWrap = true;
    });
  } else if (spec.layout === 'passi') {
    const n = spec.passi.length, overlap = 8, w = (area.width + overlap * (n - 1)) / n, h = 90;
    const top = area.top + (area.height - h) / 2;
    spec.passi.forEach((p, i) => {
      const shape = slide.shapes.addGeometricShape(i === 0 ? G.homePlate : G.chevron,
        { left: area.left + i * (w - overlap), top, width: w, height: h });
      shape.textFrame.textRange.text = p;
      pptStyleShape(shape, i % 2 ? PPT_COLORS.soft : PPT_COLORS.accent, i % 2 ? PPT_COLORS.softText : PPT_COLORS.accentText,
                    n > 4 ? 12 : 14, true);
    });
  } else if (spec.layout === 'tabella') {
    const head = spec.tabella.intestazioni, rows = spec.tabella.righe;
    const cols = head.length, rh = Math.min(42, area.height / (rows.length + 1)), cw = area.width / cols;
    [head].concat(rows).forEach((r, ri) => {
      r.forEach((cell, ci) => {
        const box = slide.shapes.addGeometricShape(G.rectangle,
          { left: area.left + ci * cw, top: area.top + ri * rh, width: cw - 2, height: rh - 2 });
        box.textFrame.textRange.text = cell || ' ';
        pptStyleShape(box, ri === 0 ? PPT_COLORS.accent : (ri % 2 ? PPT_COLORS.band : '#FFFFFF'),
                      ri === 0 ? PPT_COLORS.accentText : PPT_COLORS.bandText, rows.length > 4 ? 12 : 14, ri === 0, 'Left');
      });
    });
  }
}

async function insertPptSlides(slides) {
  const graphics = pptSupports('1.4');
  await PowerPoint.run(async (ctx) => {
    const pres = ctx.presentation;
    let layouts = [];
    try {
      const coll = pres.slideMasters.getItemAt(0).layouts;
      coll.load('items/id,items/name');
      await ctx.sync();
      layouts = coll.items;
    } catch (e) { /* layout predefinito */ }
    const ids = {
      titolo: pptFindLayout(layouts, 'titolo'), sezione: pptFindLayout(layouts, 'sezione'),
      contenuto: pptFindLayout(layouts, 'contenuto'), confronto: pptFindLayout(layouts, 'confronto'),
      soloTitolo: pptFindLayout(layouts, 'soloTitolo')
    };
    // Un layout grafico senza "Solo titolo" nel tema, o senza API per le forme, ricade sull'elenco puntato.
    const plan = slides.map(sl => {
      let layout = sl.layout;
      if (['numeri', 'passi', 'tabella'].includes(layout) && (!graphics || !ids.soloTitolo)) layout = 'contenuto';
      if (layout === 'confronto' && !ids.confronto) layout = 'contenuto';
      const layoutId = { titolo: ids.titolo, sezione: ids.sezione, contenuto: ids.contenuto, confronto: ids.confronto,
                         numeri: ids.soloTitolo, passi: ids.soloTitolo, tabella: ids.soloTitolo }[layout] || ids.contenuto;
      return { spec: sl, layout, layoutId };
    });

    const before = pres.slides.getCount();
    await ctx.sync();
    plan.forEach(p => pres.slides.add(p.layoutId ? { layoutId: p.layoutId } : undefined));
    await ctx.sync();
    pres.slides.load('items');
    await ctx.sync();
    const added = pres.slides.items.slice(before.value);
    const shapeLists = added.map(sl => { const sh = sl.shapes; sh.load('items/type,items/left,items/top,items/width,items/height'); return sh; });
    await ctx.sync();

    plan.forEach((p, i) => {
      const slide = added[i], spec = p.spec;
      const boxes = shapeLists[i].items.filter(x => ['Placeholder', 'TextBox', 'GeometricShape'].includes(x.type));
      const title = boxes[0];
      if (title) title.textFrame.textRange.text = spec.titolo;
      else if (spec.titolo) slide.shapes.addTextBox(spec.titolo, { left: 60, top: 30, width: 840, height: 70 });

      // Se il layout grafico non e' disponibile, i suoi contenuti diventano punti dell'elenco.
      const bulletText = (p.layout === 'contenuto' && spec.layout !== 'contenuto'
        ? [].concat(spec.punti, spec.passi || [], (spec.numeri || []).map(n => `${n.valore} ${n.etichetta}`),
                    ...((spec.colonne || []).map(c => [c.titolo].concat(c.punti))),
                    ((spec.tabella && spec.tabella.righe) || []).map(r => r.join(' – ')))
        : spec.punti).filter(Boolean).join('\n');

      if (p.layout === 'titolo' || p.layout === 'sezione') {
        if (boxes[1]) boxes[1].textFrame.textRange.text = spec.sottotitolo || '';
      } else if (p.layout === 'confronto') {
        spec.colonne.forEach((c, k) => {
          const target = boxes[k + 1];
          if (!target) return;
          target.textFrame.textRange.text = [c.titolo].concat(c.punti).join('\n');
          if (c.titolo) target.textFrame.textRange.getSubstring(0, c.titolo.length).font.bold = true;
        });
      } else if (p.layout === 'contenuto') {
        if (boxes[1]) boxes[1].textFrame.textRange.text = bulletText;
        else if (bulletText) slide.shapes.addTextBox(bulletText, { left: 60, top: 140, width: 840, height: 340 });
      } else {
        // Area di disegno: sotto il titolo, con gli stessi margini laterali del titolo.
        const t = title || { left: 48, top: 30, width: 864, height: 70 };
        const slideW = t.left * 2 + t.width;
        const top = t.top + t.height + 24;
        const area = { left: t.left, top, width: t.width, height: Math.max(160, Math.round(slideW * 9 / 16) - top - 40) };
        pptDrawGraphics(slide, spec, area);
      }
    });
    await ctx.sync();
    const simplified = plan.filter(p => p.layout !== p.spec.layout).length;
    notify(`Aggiunte ${slides.length} slide in fondo alla presentazione.` +
      (simplified ? ` ${simplified} con layout semplificato (il tema o la versione di PowerPoint non ha il layout richiesto).` : '') +
      '\nPer cambiare colori e caratteri: Progettazione → Temi.', 'success');
  });
}

async function applyPptImprove(p) {
  await PowerPoint.run(async (ctx) => {
    const slide = ctx.presentation.slides.getItem(p.slideId);
    // Prima di scrivere si rileggono i testi: se la slide e' cambiata dopo la proposta, non si tocca nulla.
    const ranges = p.shapes.map(s => {
      const r = slide.shapes.getItem(s.id).textFrame.textRange;
      r.load('text');
      return r;
    });
    try { await ctx.sync(); }
    catch (e) { throw new Error('La slide o una delle sue caselle non esiste più: proponi di nuovo il miglioramento.'); }
    const changed = ranges.some((r, i) => (r.text || '') !== p.shapes[i].before);
    if (changed) {
      lastPptImprove = null;
      notify('La slide è stata modificata dopo la proposta: per non perdere le tue modifiche non applico nulla. Premi di nuovo "Proponi Miglioramento".', 'warn');
      return;
    }
    ranges.forEach((r, i) => { r.text = p.shapes[i].after; });
    await ctx.sync();
    lastPptImprove = null;
    notify('Testi della slide aggiornati (Ctrl+Z per annullare).', 'success');
  });
}

async function applyToPowerPoint(text, mode) {
  if (lastResultKind === 'ppt-slides') {
    if (!lastPptSlides) { notify('Nessuna slide proposta.', 'warn'); return; }
    return insertPptSlides(lastPptSlides);
  }
  if (lastResultKind === 'ppt-improve') {
    if (!lastPptImprove) { notify('Nessun miglioramento proposto.', 'warn'); return; }
    return applyPptImprove(lastPptImprove);
  }
  // Testo: sostituisce la selezione, oppure lo aggiunge come nuova casella sulla slide selezionata.
  if (mode === 'replace') {
    await pptSetSelectedText(text);
    notify('Testo sostituito (Ctrl+Z per annullare).', 'success');
  } else {
    await PowerPoint.run(async (ctx) => {
      const sel = ctx.presentation.getSelectedSlides();
      sel.load('items');
      await ctx.sync();
      if (!sel.items.length) throw new Error('Seleziona una slide.');
      sel.items[0].shapes.addTextBox(text, { left: 60, top: 140, width: 600, height: 200 });
      await ctx.sync();
    });
    notify('Testo aggiunto come nuova casella sulla slide.', 'success');
  }
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
      await applyToExcel(text);
    } catch (err) {
      notify('Errore inserimento Excel: ' + err.message, 'error');
    }
  } else if (currentHost === 'powerpoint' && typeof PowerPoint !== 'undefined') {
    try {
      await applyToPowerPoint(text, mode);
    } catch (err) {
      notify('Errore PowerPoint: ' + err.message, 'error');
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
