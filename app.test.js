import test from 'node:test';
import assert from 'node:assert/strict';
import { documentHtml, escapeHtml, validateLaw, applicableLaws, baseLaws, resolveRole, typeNames, uniqueId } from './app.js';
import { checkSources, createServer, listenAddress } from './server.mjs';
import { buildDocx } from './docx.js';

test('renders three types and never cites unverified defaults as applicable', () => {
  const data = { org: 'Тестовый ДК', role: 'sound', location: 'Сцена' };
  const job = documentHtml({ ...data, type: 'job' }, baseLaws);
  const safety = documentHtml({ ...data, type: 'safety' }, baseLaws);
  const fire = documentHtml({ ...data, type: 'fire' }, baseLaws);
  assert.match(job, /ДОЛЖНОСТНАЯ ИНСТРУКЦИЯ/);
  assert.doesNotMatch(job, /3612-1/);
  assert.doesNotMatch(job, /1479/);
  assert.doesNotMatch(safety, /2464/);
  assert.match(safety, /Перед началом работы/);
  assert.doesNotMatch(fire, /1479/);
  assert.match(fire, /101 или 112/);
  assert.doesNotMatch(job, /299-I/);
});

test('escapes user text and validates imported acts', () => {
  assert.equal(escapeHtml('<script>'), '&lt;script&gt;');
  assert.doesNotMatch(documentHtml({ type: 'fire', org: '<img src=x>', role: 'other', customRole: '<script>' }, baseLaws), /<script>|<img/);
  const valid = { id: 'custom-1', region: 'ЛНР', kind: 'Закон', title: 'О культуре', number: '№ 1', url: 'https://nslnr.su/', scope: 'all', verified: '2026-09-25' };
  assert.equal(validateLaw(valid).status, 'review');
  assert.equal(validateLaw({ ...valid, url: 'javascript:alert(1)' }), null);
  assert.equal(validateLaw({ ...valid, title: '<'.repeat(300) }), null);
  assert.equal(validateLaw({ ...valid, effectiveFrom: '2026-02-30' }), null);
});

test('requires dated confirmation and selects laws by territory and document year', () => {
  const today = new Date().toISOString().slice(0, 10);
  const federal = { id: 'federal-check', region: 'РФ', kind: 'Закон', title: 'Проверенный пример', number: '№ 1', url: 'https://publication.pravo.gov.ru/document/123', scope: 'job', verified: today, status: 'active', effectiveFrom: '2020-01-01', effectiveTo: '', appliesLnr: false };
  assert.equal(validateLaw({ ...federal, url: 'https://publication.pravo.gov.ru/' }), null);
  assert.equal(validateLaw({ ...federal, url: 'https://example.com/document/123' }), null);
  assert.equal(validateLaw({ ...federal, verified: '' }), null);
  assert.equal(applicableLaws([federal], 'job', 'ЛНР', '2025-01-01').length, 0);
  assert.equal(applicableLaws([federal], 'job', 'РФ', '2019-01-01').length, 0);
  assert.equal(applicableLaws([federal], 'job', 'РФ', '2022-01-01').length, 1);
  assert.equal(applicableLaws([{ ...federal, appliesLnr: true }], 'job', 'ЛНР', '2022-01-01').length, 1);
  assert.equal(applicableLaws([{ ...federal, status: 'repealed' }], 'job', 'РФ', '2022-01-01').length, 0);
  assert.equal(applicableLaws([{ ...federal, effectiveTo: '2021-12-31' }], 'job', 'РФ', '2022-01-01').length, 0);
});

test('creates a real DOCX ZIP with WordprocessingML and valid escaped document text', () => {
  const bytes = buildDocx([{ text: 'Инструкция 2018 года', style: 'Title' }, { text: 'Зал & <сцена>\u000b', style: 'Normal' }]);
  assert.equal(Buffer.from(bytes.subarray(0, 4)).toString('hex'), '504b0304');
  const content = Buffer.from(bytes).toString('utf8');
  assert.match(content, /word\/document\.xml/);
  assert.match(content, /Инструкция 2018 года/);
  assert.match(content, /Зал &amp; &lt;сцена&gt;/);
  assert.match(content, /\[Content_Types\]\.xml/);
  const view = new DataView(bytes.buffer);
  let local = 0;
  let documentXml = '';
  while (view.getUint32(local, true) === 0x04034b50) {
    const nameLength = view.getUint16(local + 26, true);
    const name = new TextDecoder().decode(bytes.subarray(local + 30, local + 30 + nameLength));
    const dataOffset = local + 30 + nameLength;
    const dataSize = view.getUint32(local + 18, true);
    if (name === 'word/document.xml') documentXml = new TextDecoder().decode(bytes.subarray(dataOffset, dataOffset + dataSize));
    local = dataOffset + dataSize;
  }
  assert.match(documentXml, /Зал &amp; &lt;сцена&gt;/);
  assert.doesNotMatch(documentXml, /\u000b/);
  const trailer = bytes.length - 22;
  assert.equal(view.getUint32(trailer, true), 0x06054b50);
  assert.equal(view.getUint16(trailer + 10, true), 5);
  let position = view.getUint32(trailer + 16, true);
  const names = [];
  for (let index = 0; index < 5; index++) {
    assert.equal(view.getUint32(position, true), 0x02014b50);
    const nameLength = view.getUint16(position + 28, true);
    const name = new TextDecoder().decode(bytes.subarray(position + 46, position + 46 + nameLength));
    const localOffset = view.getUint32(position + 42, true);
    assert.equal(view.getUint32(localOffset, true), 0x04034b50);
    assert.equal(view.getUint32(localOffset + 18, true), view.getUint32(position + 20, true));
    names.push(name);
    position += 46 + nameLength;
  }
  assert.deepEqual(names, ['[Content_Types].xml', '_rels/.rels', 'word/_rels/document.xml.rels', 'word/document.xml', 'word/styles.xml']);
});

test('accepts both suggested and manually entered roles, including previous saved drafts', () => {
  assert.deepEqual(resolveRole({ role: 'Звукорежиссёр' }), { name: 'Звукорежиссёр', roleId: 'sound' });
  assert.deepEqual(resolveRole({ role: 'Режиссёр массовых мероприятий' }), { name: 'Режиссёр массовых мероприятий', roleId: 'other' });
  assert.deepEqual(resolveRole({ role: 'other', customRole: 'Методист' }), { name: 'Методист', roleId: 'other' });
  assert.match(documentHtml({ type: 'job', role: 'Звукорежиссёр' }, baseLaws), /Подготавливает, настраивает и эксплуатирует/);
  assert.match(documentHtml({ type: 'job', role: 'Методист' }, baseLaws), /для Методист/);
  assert.equal(Object.keys(typeNames).length, 3);
});

test('supports local-only defaults, hosted binding and HTTP LAN IDs', () => {
  assert.deepEqual(listenAddress({}), { host: '127.0.0.1', port: 3210 });
  assert.deepEqual(listenAddress({ RENDER: 'true', PORT: '10000' }), { host: '0.0.0.0', port: 10000 });
  assert.deepEqual(listenAddress({ HOST: '0.0.0.0', PORT: '3211' }), { host: '0.0.0.0', port: 3211 });
  const first = uniqueId({ getRandomValues: bytes => bytes.fill(12) });
  assert.match(first, /^[a-f0-9]{32}$/);
  assert.equal(first, '0c'.repeat(16));
});

test('source monitor reports errors independently and fingerprints content', async () => {
  const result = await checkSources(async url => {
    if (url.includes('nslnr')) throw Error('offline');
    return { ok: true, text: async () => '<html><script>random</script><p>Новый акт</p></html>' };
  });
  assert.match(result[0].fingerprint, /^[a-f0-9]{64}$/);
  assert.match(result[1].error, /offline/);
});

test('serves the form, styles and script over HTTP', async () => {
  let checks = 0;
  const server = createServer({ sourceChecker: async () => { checks++; return [{ id: 'rf', error: 'offline' }, { id: 'lnr', error: 'offline' }]; } });
  await new Promise(resolve => server.listen(0, '127.0.0.1', resolve));
  try {
    const base = `http://127.0.0.1:${server.address().port}`;
    const [page, style, script, docx, health, unknown] = await Promise.all(['/', '/form.css', '/app.js', '/docx.js', '/health', '/missing'].map(uri => fetch(base + uri)));
    assert.equal(page.status, 200);
    assert.match(await page.text(), /Создать инструкцию/);
    assert.equal(style.status, 200);
    assert.match(await style.text(), /create-button/);
    assert.equal(script.status, 200);
    assert.match(await script.text(), /selectType/);
    assert.equal(docx.status, 200);
    assert.deepEqual(await health.json(), { status: 'ok' });
    assert.equal(unknown.status, 404);
    const results = await Promise.all(Array.from({ length: 5 }, () => fetch(base + '/api/check')));
    assert.ok(results.every(result => result.status === 200));
    assert.equal(checks, 1);
  } finally { await new Promise((resolve, reject) => server.close(error => error ? reject(error) : resolve())); }
});
