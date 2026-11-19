// Controllo settimanale: casi negativi della verifica CORS e pulizia del riepilogo (rete simulata).
const { test } = require('node:test');
const assert = require('node:assert/strict');
const { checkCors, sanitize, endpoints, loadPanel, ORIGIN } = require('../scripts/controllo-provider.cjs');

const reply = (status, h) => async () => ({ status, headers: { get: k => h[k.toLowerCase()] ?? null } });
const ep = { url: 'https://api.example/v1/chat/completions', headers: ['authorization', 'content-type', 'x-title'] };

test('accettato con origine, metodo e intestazioni', async () => {
  assert.equal(await checkCors(ep, reply(204, { 'access-control-allow-origin': ORIGIN,
    'access-control-allow-methods': 'GET, POST', 'access-control-allow-headers': 'Authorization, Content-Type, X-Title' })), null);
  assert.equal(await checkCors(ep, reply(200, { 'access-control-allow-origin': '*', 'access-control-allow-headers': '*' })), null);
});

test('origine mancante o diversa', async () => {
  assert.match(await checkCors(ep, reply(200, { 'access-control-allow-headers': '*' })), /origine del sito non accettata/);
  assert.match(await checkCors(ep, reply(200, { 'access-control-allow-origin': 'https://altro.example', 'access-control-allow-headers': '*' })), /non accettata/);
});

test('intestazione o metodo non accettati', async () => {
  assert.match(await checkCors(ep, reply(200, { 'access-control-allow-origin': '*', 'access-control-allow-headers': 'authorization,content-type' })), /intestazioni non accettate: x-title/);
  assert.match(await checkCors(ep, reply(200, { 'access-control-allow-origin': '*', 'access-control-allow-methods': 'GET', 'access-control-allow-headers': '*' })), /POST non accettato/);
});

test('preliminare rifiutata', async () => {
  assert.match(await checkCors(ep, reply(403, {})), /HTTP 403/);
});

test('Claude: verificata anche anthropic-beta; OpenRouter: intestazioni di attribuzione', () => {
  const eps = endpoints(loadPanel(async () => ({})));
  assert.ok(eps.claude.headers.includes('anthropic-beta'));
  assert.ok(eps.openrouter.headers.includes('x-title') && eps.openrouter.headers.includes('http-referer'));
  assert.equal(eps.qwen.url, 'https://dashscope-intl.aliyuncs.com/compatible-mode/v1/chat/completions');
});

test('riepilogo: chiavi oscurate e Markdown/HTML neutralizzati', () => {
  const s = sanitize('errore con chiave segreta-123456 e sk-abcdefghijklmnop <img src=x> **grassetto**\nriga', ['segreta-123456']);
  assert.ok(!s.includes('segreta-123456') && !s.includes('sk-abcdefghijklmnop'));
  assert.ok(!s.includes('<img') && s.includes('&lt;img'));
  assert.ok(!s.includes('**grassetto**') && s.includes('\\*\\*grassetto'));
  assert.ok(!s.includes('\n'));
});
