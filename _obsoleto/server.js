const https = require('https');
const fs = require('fs');
const path = require('path');
const url = require('url');

// Lettura dinamica delle chiavi da .env senza riavviare il server
function getEnvKeys() {
  const keys = {
    gemini: process.env.GEMINI_API_KEY || '',
    openai: process.env.OPENAI_API_KEY || '',
    claude: process.env.ANTHROPIC_API_KEY || ''
  };

  const envPath = path.join(__dirname, '.env');
  if (fs.existsSync(envPath)) {
    const lines = fs.readFileSync(envPath, 'utf8').split(/\r?\n/);
    for (const line of lines) {
      const trimmed = line.trim();
      if (!trimmed || trimmed.startsWith('#')) continue;
      const eqIdx = trimmed.indexOf('=');
      if (eqIdx !== -1) {
        const key = trimmed.slice(0, eqIdx).trim();
        const val = trimmed.slice(eqIdx + 1).trim();
        if (key === 'GEMINI_API_KEY' && val.length > 5) keys.gemini = val;
        if (key === 'OPENAI_API_KEY' && val.length > 5) keys.openai = val;
        if (key === 'ANTHROPIC_API_KEY' && val.length > 5) keys.claude = val;
      }
    }
  }
  return keys;
}

const PORT = process.env.PORT || 3000;

// Caricamento certificato SSL (PFX)
let sslOptions = {};
const pfxPath = path.join(__dirname, 'certs', 'localhost.pfx');
if (fs.existsSync(pfxPath)) {
  sslOptions = {
    pfx: fs.readFileSync(pfxPath),
    passphrase: 'officeaddin'
  };
} else {
  console.error("ERRORE: Certificato SSL non trovato in certs/localhost.pfx");
}

const MIME_TYPES = {
  '.html': 'text/html; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.js': 'application/javascript; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.svg': 'image/svg+xml',
  '.ico': 'image/x-icon',
  '.xml': 'application/xml; charset=utf-8'
};

// --- HANDLER GOOGLE GEMINI ---
async function callGeminiApi({ apiKey, model, fullContent, systemInstruction, temperature }) {
  const candidateModels = model 
    ? [model, 'gemini-3.1-flash-lite', 'gemini-2.5-flash-lite'] 
    : ['gemini-3.1-flash-lite', 'gemini-2.5-flash-lite', 'gemini-3.8-flash'];

  let lastError = null;
  for (const m of candidateModels) {
    try {
      const geminiUrl = `https://generativelanguage.googleapis.com/v1beta/models/${m}:generateContent?key=${encodeURIComponent(apiKey)}`;
      const geminiBody = {
        systemInstruction: { parts: [{ text: systemInstruction }] },
        contents: [{ role: "user", parts: [{ text: fullContent }] }],
        generationConfig: {
          temperature: temperature !== undefined ? temperature : 0.4
        }
      };

      const response = await fetch(geminiUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(geminiBody)
      });

      const data = await response.json();
      if (response.ok && data?.candidates?.[0]?.content?.parts?.[0]?.text) {
        return data.candidates[0].content.parts[0].text;
      }
      lastError = data?.error?.message || `Errore HTTP ${response.status} su ${m}`;
    } catch (err) {
      lastError = err.message;
    }
  }
  throw new Error("Gemini API Error: " + (lastError || "Nessuna risposta valida"));
}

// --- HANDLER OPENAI / CODEX ---
async function callOpenAiApi({ apiKey, model, fullContent, systemInstruction, temperature }) {
  const chosenModel = model || 'gpt-4o-mini';
  const openAiUrl = 'https://api.openai.com/v1/chat/completions';

  const body = {
    model: chosenModel,
    messages: [
      { role: 'system', content: systemInstruction },
      { role: 'user', content: fullContent }
    ],
    temperature: temperature !== undefined ? temperature : 0.4
  };

  const response = await fetch(openAiUrl, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${apiKey.trim()}`
    },
    body: JSON.stringify(body)
  });

  const data = await response.json();
  if (!response.ok) {
    throw new Error(data?.error?.message || `OpenAI HTTP ${response.status}`);
  }

  return data?.choices?.[0]?.message?.content || '';
}

// --- HANDLER ANTHROPIC CLAUDE ---
async function callClaudeApi({ apiKey, model, fullContent, systemInstruction, temperature }) {
  const candidateModels = model 
    ? [model, 'claude-sonnet-5-5', 'claude-haiku-5-5', 'claude-opus-5-5']
    : ['claude-sonnet-5-5', 'claude-haiku-5-5', 'claude-opus-5-5'];

  const claudeUrl = 'https://api.anthropic.com/v1/messages';
  let lastError = null;

  for (const m of candidateModels) {
    try {
      const body = {
        model: m,
        max_tokens: 4096,
        system: systemInstruction,
        messages: [
          { role: 'user', content: fullContent }
        ],
        temperature: temperature !== undefined ? temperature : 0.4
      };

      const response = await fetch(claudeUrl, {
        method: 'POST',
        headers: {
          'content-type': 'application/json',
          'x-api-key': apiKey.trim(),
          'anthropic-version': '2023-06-01'
        },
        body: JSON.stringify(body)
      });

      const data = await response.json();
      if (response.ok) {
        const textPart = data?.content?.find(c => c.type === 'text');
        if (textPart?.text) {
          return textPart.text;
        }
      }
      lastError = data?.error?.message || `Anthropic HTTP ${response.status} su ${m}`;
    } catch (err) {
      lastError = err.message;
    }
  }

  throw new Error("Anthropic Claude Error: " + (lastError || "Nessuna risposta"));
}

function markdownToHtml(md) {
  if (!md) return '';
  let html = md.trim();

  // Rimuovi backticks `codice`
  html = html.replace(/`([^`]+)`/g, '$1');

  // Converti **grassetto** in <b>grassetto</b>
  html = html.replace(/\*\*([^*]+)\*\*/g, '<b>$1</b>');

  // Converti *corsivo* in <i>corsivo</i>
  html = html.replace(/(^|[^*])\*([^*]+)\*(?!\*)/g, '$1<i>$2</i>');

  // Converti elenchi puntati (* o - a inizio riga) in bullet points puliti
  html = html.replace(/^[*-]\s+(.+)$/gm, '<div style="margin-left: 12px; margin-bottom: 4px;">&bull; $1</div>');

  // Converti a capo
  html = html.replace(/\r?\n\r?\n/g, '<br><br>');
  html = html.replace(/\r?\n/g, '<br>');

  return html;
}

function markdownToPlainText(md) {
  if (!md) return '';
  let txt = md.trim();
  txt = txt.replace(/`([^`]+)`/g, '$1');
  txt = txt.replace(/\*\*([^*]+)\*\*/g, '$1');
  txt = txt.replace(/(^|[^*])\*([^*]+)\*(?!\*)/g, '$1$2');
  return txt;
}

// Router centrale AI
async function handleAICall(req, res) {
  let body = '';
  req.on('data', chunk => { body += chunk; });
  req.on('end', async () => {
    try {
      const payload = JSON.parse(body || '{}');
      const provider = (payload.provider || 'gemini').toLowerCase();
      const model = payload.model || '';
      const prompt = payload.prompt || '';
      const context = payload.context || '';
      const systemInstruction = payload.systemInstruction || 'Sei un assistente AI integrato in Microsoft Office. Rispondi in italiano in modo chiaro, utile e conciso. Non inserire prefissi o etichette come "**Risposta:**", né racchiudere il testo tra asterischi.';
      const temperature = payload.temperature;

      const envKeys = getEnvKeys();

      // Risoluzione chiave: priorità a quella passata dall'interfaccia (localStorage), altrimenti da .env
      let apiKey = payload.apiKey;
      if (!apiKey) {
        if (provider === 'gemini') apiKey = envKeys.gemini;
        else if (provider === 'openai') apiKey = envKeys.openai;
        else if (provider === 'claude') apiKey = envKeys.claude;
      }

      if (!apiKey) {
        res.writeHead(400, { 'Content-Type': 'application/json' });
        return res.end(JSON.stringify({
          error: `Chiave API mancante per il provider ${provider.toUpperCase()}. Inseriscila nelle impostazioni (⚙️) o nel file .env!`
        }));
      }

      // Costruisci il prompt combinato
      let fullContent = '';
      if (context) {
        fullContent += `CONTESTO DAL DOCUMENTO:\n"""\n${context}\n"""\n\n`;
      }
      fullContent += `RICHIESTA:\n${prompt}`;

      let generatedText = '';
      if (provider === 'openai') {
        generatedText = await callOpenAiApi({ apiKey, model, fullContent, systemInstruction, temperature });
      } else if (provider === 'claude') {
        generatedText = await callClaudeApi({ apiKey, model, fullContent, systemInstruction, temperature });
      } else {
        // default: Gemini
        generatedText = await callGeminiApi({ apiKey, model, fullContent, systemInstruction, temperature });
      }

      const generatedHtml = markdownToHtml(generatedText);
      const cleanText = markdownToPlainText(generatedText);

      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ 
        text: cleanText, 
        html: generatedHtml,
        raw: generatedText,
        provider, 
        model 
      }));
    } catch (err) {
      console.error("Errore chiamata AI:", err);
      res.writeHead(500, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ error: err.message }));
    }
  });
}

const server = https.createServer(sslOptions, (req, res) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization');

  if (req.method === 'OPTIONS') {
    res.writeHead(204);
    return res.end();
  }

  const parsedUrl = new URL(req.url, `https://${req.headers.host || 'localhost:3000'}`);
  const pathname = parsedUrl.pathname;

  // Endpoint Stato Chiavi
  if (pathname === '/api/status' && req.method === 'GET') {
    const keys = getEnvKeys();
    res.writeHead(200, { 'Content-Type': 'application/json' });
    return res.end(JSON.stringify({
      ok: true,
      keys: {
        gemini: Boolean(keys.gemini && keys.gemini.length > 5),
        openai: Boolean(keys.openai && keys.openai.length > 5),
        claude: Boolean(keys.claude && keys.claude.length > 5)
      }
    }));
  }

  // Endpoint Unificato Chiamate AI (supporta anche il vecchio /api/gemini per retrocompatibilità)
  if ((pathname === '/api/ai' || pathname === '/api/gemini') && req.method === 'POST') {
    return handleAICall(req, res);
  }

  // Servizio file statici da public/
  let filePath = path.join(__dirname, 'public', pathname === '/' ? 'index.html' : pathname);

  if (!fs.existsSync(filePath)) {
    const rootCandidate = path.join(__dirname, pathname);
    if (fs.existsSync(rootCandidate)) {
      filePath = rootCandidate;
    }
  }

  if (fs.existsSync(filePath) && fs.statSync(filePath).isFile()) {
    const ext = path.extname(filePath).toLowerCase();
    const contentType = MIME_TYPES[ext] || 'application/octet-stream';
    res.writeHead(200, { 'Content-Type': contentType });
    fs.createReadStream(filePath).pipe(res);
  } else {
    res.writeHead(404, { 'Content-Type': 'text/plain; charset=utf-8' });
    res.end('404 Not Found: ' + pathname);
  }
});

server.listen(PORT, () => {
  const keys = getEnvKeys();
  console.log(`=======================================================`);
  console.log(`🚀 Prototipo Office Add-in Multi-Modello (HTTPS)`);
  console.log(`🔗 URL: https://localhost:${PORT}`);
  console.log(`🔑 Chiavi rilevate da .env:`);
  console.log(`   - Gemini:   ${keys.gemini ? '✅ Attiva' : '❌ Non impostata'}`);
  console.log(`   - OpenAI:   ${keys.openai ? '✅ Attiva' : '❌ Non impostata'}`);
  console.log(`   - Claude:   ${keys.claude ? '✅ Attiva' : '❌ Non impostata'}`);
  console.log(`=======================================================`);
});
