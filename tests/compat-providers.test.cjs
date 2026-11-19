// Provider compatibili OpenAI (Qwen, MiniMax, DeepSeek, OpenRouter): rete simulata, chiavi fittizie.
const { test } = require('node:test');
const assert = require('node:assert/strict');
const vm = require('node:vm');
const fs = require('node:fs');
const source = fs.readFileSync('public/ai.js', 'utf8');

function setup(responses) {
  const data = new Map();
  const calls = [];
  const fetch = async (url, init = {}) => {
    calls.push({ url, init, body: init.body ? JSON.parse(init.body) : null });
    const r = responses.shift();
    return { ok: r.status < 300, status: r.status, headers: { get: () => null }, json: async () => r.json };
  };
  const context = vm.createContext({
    localStorage: { getItem: k => data.get(k) ?? null, setItem: (k, v) => data.set(k, v), removeItem: k => data.delete(k) },
    fetch, setTimeout, clearTimeout, AbortController
  });
  vm.runInContext(source, context);
  return { calls, run: s => vm.runInContext(s, context) };
}

const ok = text => ({ status: 200, json: { model: 'm-reale', choices: [{ message: { content: text }, finish_reason: 'stop' }] } });

test('qwen: endpoint internazionale, Bearer, ragionamento tolto', async () => {
  const s = setup([ok('<think>ragiono a lungo</think>\nRisposta finale')]);
  s.run("aiSetKey('qwen','fake-qwen')");
  const r = await s.run("aiCall({ provider: 'qwen', model: 'qwen-plus', prompt: 'ciao' })");
  assert.equal(r.text, 'Risposta finale');
  assert.equal(r.actualModel, 'm-reale');
  assert.equal(s.calls[0].url, 'https://dashscope-intl.aliyuncs.com/compatible-mode/v1/chat/completions');
  assert.equal(s.calls[0].init.headers.authorization, 'Bearer fake-qwen');
  assert.equal(s.calls[0].body.model, 'qwen-plus');
  assert.equal(s.calls[0].body.messages[0].role, 'system');
});

test('minimax: reasoning_split nel corpo; errore 200 con base_resp segnalato', async () => {
  const s = setup([{ status: 200, json: { choices: [], base_resp: { status_code: 1004, status_msg: 'login fail' } } }]);
  s.run("aiSetKey('minimax','fake-mm')");
  await assert.rejects(s.run("aiCall({ provider: 'minimax', model: 'MiniMax-M3', prompt: 'x' })"), /MiniMax non ha restituito testo \(login fail\)/);
  assert.equal(s.calls[0].url, 'https://api.minimax.io/v1/chat/completions');
  assert.equal(s.calls[0].body.reasoning_split, true);
});

test('deepseek: chiave non valida classificata', async () => {
  const s = setup([{ status: 401, json: { error: { message: 'Authentication Fails' } } }]);
  s.run("aiSetKey('deepseek','fake-ds')");
  await assert.rejects(s.run("aiCall({ provider: 'deepseek', model: 'deepseek-v4-flash', prompt: 'x' })"), /Chiave DeepSeek non valida/);
  assert.equal(s.calls[0].url, 'https://api.deepseek.com/chat/completions');
});

test('openrouter: intestazioni di attribuzione; verifica chiave su /key poi elenco modelli', async () => {
  const s = setup([{ status: 200, json: { data: { label: 'k' } } }, { status: 200, json: { data: [{ id: 'a/b:free' }, { id: 'c/d' }] } }]);
  const ids = await s.run("aiListModels('openrouter', 'fake-or')");
  assert.deepEqual([...ids], ['a/b:free', 'c/d']);
  assert.equal(s.calls[0].url, 'https://openrouter.ai/api/v1/key');
  assert.equal(s.calls[1].url, 'https://openrouter.ai/api/v1/models');
  assert.equal(s.calls[0].init.headers['X-Title'], 'Personal AI Assistant');
  assert.equal(s.calls[0].init.headers.authorization, 'Bearer fake-or');
});

test('risposta troncata segnalata', async () => {
  const s = setup([{ status: 200, json: { choices: [{ message: { content: 'parziale' }, finish_reason: 'length' }] } }]);
  s.run("aiSetKey('qwen','fake')");
  const r = await s.run("aiCall({ provider: 'qwen', model: 'qwen-plus', prompt: 'x' })");
  assert.match(r.text, /^parziale\n\n\[Risposta troncata/);
});

test('nuovi provider nella gestione chiavi e conferma per quelli con sede in Cina', () => {
  const s = setup([]);
  assert.deepEqual([...s.run('AI_PROVIDERS')], ['gemini', 'openai', 'claude', 'qwen', 'minimax', 'deepseek', 'openrouter']);
  assert.deepEqual(Object.keys(s.run('AI_CHINA_PROVIDERS')), ['qwen', 'minimax', 'deepseek']);
  assert.equal(s.run("aiSetKey('openrouter','fake'); aiKeyStorage('openrouter')"), 'salvata');
  for (const p of ['qwen', 'minimax', 'deepseek', 'openrouter'])
    assert.ok(s.run(`AI_CATALOG.some(m => m.provider === '${p}')`), p);
});

test('pannello: campi chiave, link e conferma presenti in index.html', () => {
  const html = fs.readFileSync('public/index.html', 'utf8');
  for (const p of ['qwen', 'minimax', 'deepseek', 'openrouter']) {
    assert.match(html, new RegExp(`id="input-${p}-key"`));
    assert.match(html, new RegExp(`id="badge-${p}-status"`));
  }
  assert.match(html, /id="china-confirm"/);
  assert.match(html, /id="custom-model-row"/);
});
