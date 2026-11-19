const { test } = require('node:test');
const assert = require('node:assert/strict');
const vm = require('node:vm');
const fs = require('node:fs');
const source = fs.readFileSync('public/ai.js', 'utf8');
function setup(initial = {}, blocked = {}) {
  const data = new Map(Object.entries(initial));
  const context = vm.createContext({ localStorage: {
    getItem(k) { if (blocked.read?.(k)) throw Error('read denied'); return data.get(k) ?? null; },
    setItem(k,v) { if (blocked.write?.(k)) throw Error('write denied'); data.set(k,v); },
    removeItem(k) { if (blocked.remove?.(k)) throw Error('remove denied'); data.delete(k); }
  }, setTimeout, clearTimeout });
  vm.runInContext(source, context);
  return { data, run: s => vm.runInContext(s, context) };
}
test('persistent key survives reopening', () => {
  const s = setup(); assert.equal(s.run("aiSetKey('openai','fake-key')"), true);
  assert.equal(setup(Object.fromEntries(s.data)).run("aiGetKey('openai')"), 'fake-key');
});
test('session choice removes persistent key and does not survive reopening', () => {
  const s = setup({'paa.key.openai':'fake-key'});
  assert.equal(s.run('aiSetPersist(false)'), true);
  assert.equal(s.run("aiGetKey('openai')"), 'fake-key');
  assert.equal(s.run("aiKeyStorage('openai')"), 'sessione');
  assert.equal(setup(Object.fromEntries(s.data)).run("aiGetKey('openai')"), '');
});
test('legacy migration respects session preference', () => {
  const s = setup({'paa.persist':'0',openai_api_key:'fake-key'});
  assert.equal(s.run("aiGetKey('openai')"), 'fake-key');
  assert.equal(s.data.has('openai_api_key'), false);
  assert.equal(s.data.has('paa.key.openai'), false);
});
test('failed legacy removal is reported at startup', () => {
  const s = setup({'paa.persist':'0',openai_api_key:'fake-key'}, {remove:k=>k==='openai_api_key'});
  assert.equal(s.run("aiKeyStorage('openai')"), 'salvata');
  assert.equal(s.run("aiLegacyNotRemoved().includes('openai')"), true);
});
test('failed current removal must not label disk key as session only', () => {
  const s = setup({'paa.key.openai':'fake-key'}, {remove:k=>k==='paa.key.openai'});
  assert.equal(s.run('aiSetPersist(false)'), false);
  assert.equal(s.run("aiKeyStorage('openai')"), 'salvata');
});
test('failed storage write must be reported to caller', () => {
  const s = setup({}, {write:k=>k==='paa.key.openai'});
  assert.equal(s.run("aiSetKey('openai','fake-key')"), false);
  assert.equal(s.run("aiGetKey('openai')"), 'fake-key');
});
test('legacy removal failure must not be hidden by successful current save', () => {
  const s = setup({openai_api_key:'old-fake'}, {remove:k=>k==='openai_api_key'});
  assert.equal(s.run("aiSetKey('openai','new-fake')"), false);
});
test('clear removes both storage names and memory', () => {
  const s = setup({openai_api_key:'fake-key'});
  assert.equal(s.run("aiSetKey('openai','')"), true);
  assert.equal(s.run("aiGetKey('openai')"), '');
  assert.equal(s.run("aiKeyStorage('openai')"), '');
});
test('failed preference write still keeps new keys in memory', () => {
  const s = setup({}, {write:k=>k==='paa.persist'});
  assert.equal(s.run('aiSetPersist(false)'), false);
  s.run("aiSetKey('openai','fake-key')");
  assert.equal(s.data.has('paa.key.openai'), false);
});
test('successful retry clears uncertain current storage status', () => {
  let deny = true;
  const s = setup({'paa.key.openai':'fake-key'}, {remove:k=>deny && k==='paa.key.openai'});
  assert.equal(s.run('aiSetPersist(false)'), false);
  deny = false;
  assert.equal(s.run('aiSetPersist(false)'), true);
  assert.equal(s.run("aiKeyStorage('openai')"), 'sessione');
});
test('unreadable storage cannot be reported as successfully cleared', () => {
  const s = setup({}, {read:()=>true});
  assert.equal(s.run("aiSetKey('openai','')"), false);
});
test('installed guide hides download and renumbers remaining steps', () => {
  const html = fs.readFileSync('public/installa.html', 'utf8');
  const code = html.slice(html.indexOf('// Aperta dall\'installatore:'), html.lastIndexOf('</script>'));
  const download = {hidden:false};
  const cards = [download, ...Array.from({length:3}, () => ({hidden:false, number:{}, querySelector() { return this.number; }}))];
  const ids = {installato:{hidden:true}, 'titolo-installazione':{}};
  vm.runInNewContext(code, {location:{hash:'#installato'}, document:{
    getElementById:id=>ids[id],
    querySelectorAll:q=>q==='.solo-download'?[download]:cards
  }});
  assert.equal(download.hidden,true);
  assert.equal(ids.installato.hidden,false);
  assert.deepEqual(cards.slice(1).map(c=>c.number.textContent),['1','2','3']);
});
