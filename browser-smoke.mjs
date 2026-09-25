// Optional smoke test against an already running Chromium/Edge with --remote-debugging-port=9238.
import assert from 'node:assert/strict';

const targets = await (await fetch('http://127.0.0.1:9238/json')).json();
const target = targets.find(item => item.type === 'page' && item.url.startsWith('http://localhost:3210/'));
assert.ok(target, 'Откройте http://localhost:3210/ в браузере с портом отладки 9238');
const ws = new WebSocket(target.webSocketDebuggerUrl);
await new Promise((resolve, reject) => { ws.addEventListener('open', resolve, { once: true }); ws.addEventListener('error', reject, { once: true }); });
let counter = 0;
const pending = new Map();
const exceptions = [];
ws.addEventListener('message', event => {
  const response = JSON.parse(event.data);
  if (response.method === 'Runtime.exceptionThrown') exceptions.push(response.params.exceptionDetails.text);
  if (!response.id) return;
  const handler = pending.get(response.id);
  if (!handler) return;
  pending.delete(response.id);
  if (response.error || response.result?.exceptionDetails) handler.reject(Error(JSON.stringify(response.error || response.result.exceptionDetails)));
  else handler.resolve(response.result);
});
function call(method, params = {}) {
  return new Promise((resolve, reject) => {
    const id = ++counter;
    pending.set(id, { resolve, reject });
    ws.send(JSON.stringify({ id, method, params }));
  });
}
async function evaluate(expression) {
  const response = await call('Runtime.evaluate', { expression, returnByValue: true, awaitPromise: true });
  return response.result.value;
}
async function ready() {
  for (let i = 0; i < 30; i++) {
    if (await evaluate('Boolean(document.querySelector("#document-preview .doc-title"))')) return;
    await new Promise(resolve => setTimeout(resolve, 100));
  }
  throw Error('Приложение не загрузилось');
}
async function waitFor(expression) {
  for (let i = 0; i < 40; i++) {
    if (await evaluate(expression)) return;
    await new Promise(resolve => setTimeout(resolve, 100));
  }
  throw Error(`Не дождались результата: ${expression}`);
}

try {
  await call('Runtime.enable');
  await ready();
  await evaluate('document.querySelector("[data-type=job]").click()');
  assert.match(await evaluate('document.querySelector("#document-preview .doc-title").textContent'), /ДОЛЖНОСТНАЯ/);
  await evaluate('document.querySelector("[data-type=safety]").click()');
  assert.match(await evaluate('document.querySelector("#document-preview .doc-title").textContent'), /ОХРАНЕ ТРУДА/);
  assert.equal(await evaluate('document.querySelector("[data-type=safety]").getAttribute("aria-pressed")'), 'true');
  await evaluate('document.querySelector("[data-type=fire]").click()');
  assert.match(await evaluate('document.querySelector("#document-preview .doc-title").textContent'), /ПОЖАРНОЙ БЕЗОПАСНОСТИ/);
  assert.equal(await evaluate('document.querySelector("#fire-fields").hidden'), false);
  const before = await evaluate('JSON.parse(localStorage.getItem("instruction-studio-v1-archive") || "[]").length');
  await evaluate('(() => { for (const [name, value] of [["org", "   "], ["role", "Методист"]]) { const input = document.querySelector(`[name=${name}]`); input.value = value; input.dispatchEvent(new Event("input", { bubbles: true })); } document.querySelector(".create-button").click(); })()');
  assert.equal(await evaluate('JSON.parse(localStorage.getItem("instruction-studio-v1-archive") || "[]").length'), before);
  assert.match(await evaluate('document.querySelector("[name=org]").validationMessage'), /пробел/);
  await evaluate('(() => { for (const [name, value] of [["org", "Тестовый дом культуры"], ["role", "Методист"], ["date", "2018-10-10"]]) { const input = document.querySelector(`[name=${name}]`); input.value = value; input.dispatchEvent(new Event("input", { bubbles: true })); } document.querySelector(".create-button").click(); })()');
  assert.match(await evaluate('document.querySelector("#preview-status").textContent'), /Инструкция создана/);
  assert.match(await evaluate('document.querySelector("#document-preview .doc-title").textContent'), /Методист/);
  assert.match(await evaluate('document.querySelector("#document-preview .doc-head").textContent'), /Тестовый дом культуры/);
  assert.match(await evaluate('document.querySelector("#archive-list").textContent'), /Методист/);
  assert.match(await evaluate('document.querySelector("#archive-list").textContent'), /2018-10-10/);
  assert.equal(await evaluate('JSON.parse(localStorage.getItem("instruction-studio-v1-archive")).length'), before + 1);
  const word = await evaluate('import("/docx.js").then(({ buildDocx, extractParagraphs }) => { const saved = JSON.parse(localStorage.getItem("instruction-studio-v1-archive"))[0]; const bytes = buildDocx(extractParagraphs(saved.html)); return { signature: [...bytes.slice(0, 4)].join(","), hasText: new TextDecoder().decode(bytes).includes("Тестовый дом культуры") }; })');
  assert.deepEqual(word, { signature: '80,75,3,4', hasText: true });
  await call('Page.reload', { ignoreCache: true });
  await ready();
  assert.equal(await evaluate('document.querySelector("#role").value'), 'Методист');
  assert.match(await evaluate('document.querySelector("#document-preview .doc-title").textContent'), /ПОЖАРНОЙ БЕЗОПАСНОСТИ/);
  await evaluate('document.querySelector("[data-page=laws]").click(); document.querySelector("#add-law").click()');
  assert.equal(await evaluate('document.querySelector("#law-dialog").open'), true);
  await evaluate('(() => { const form = document.querySelector("#law-form"); for (const [name, value] of [["region", "ЛНР"], ["kind", "Закон"], ["title", "Проверочный акт"], ["number", "№ 42"], ["url", "https://nslnr.su/"]]) form.elements[name].value = value; form.querySelector("[type=submit]").click(); })()');
  assert.equal(await evaluate('document.querySelector("#law-dialog").open'), false);
  assert.match(await evaluate('document.querySelector("#law-list").textContent'), /Проверочный акт/);
  assert.doesNotMatch(await evaluate('document.querySelector("#document-preview").textContent'), /Проверочный акт/);
  await evaluate('(() => { const row = [...document.querySelectorAll(".law-row")].find(node => node.textContent.includes("Проверочный акт")); row.querySelector(".edit-law").click(); const f = document.querySelector("#law-form"); f.elements.status.value = "active"; f.elements.effectiveFrom.value = "2017-01-01"; f.elements.verified.value = new Date().toISOString().slice(0, 10); f.elements.url.value = "https://nslnr.su/zakonodatelstvo/act/42"; f.querySelector("[type=submit]").click(); })()');
  assert.equal(await evaluate('document.querySelector("#law-dialog").open'), false);
  assert.match(await evaluate('document.querySelector("#document-preview").textContent'), /Проверочный акт/);
  await evaluate('(() => { const row = [...document.querySelectorAll(".law-row")].find(node => node.textContent.includes("Проверочный акт")); row.querySelector(".edit-law").click(); const f = document.querySelector("#law-form"); f.elements.status.value = "repealed"; f.querySelector("[type=submit]").click(); })()');
  assert.doesNotMatch(await evaluate('document.querySelector("#document-preview").textContent'), /Проверочный акт/);
  await evaluate('(() => { window.confirm = () => true; const row = [...document.querySelectorAll(".law-row")].find(node => node.textContent.includes("Проверочный акт")); row.querySelector(".delete-law").click(); })()');
  assert.doesNotMatch(await evaluate('document.querySelector("#law-list").textContent'), /Проверочный акт/);
  await evaluate('(() => { const record = { id: "import-check", region: "ЛНР", kind: "Закон", title: "Импортированный акт", number: "№ 7", url: "https://nslnr.su/", scope: "fire", verified: "" }; const file = new File([JSON.stringify({ format: "instruction-studio-laws", version: 1, laws: [record] })], "laws.json", { type: "application/json" }); const dt = new DataTransfer(); dt.items.add(file); const input = document.querySelector("#import-laws"); input.files = dt.files; input.dispatchEvent(new Event("change", { bubbles: true })); })()');
  await waitFor('document.querySelector("#law-list").textContent.includes("Импортированный акт")');
  await evaluate('(() => { window.__capturedBlob = null; URL.createObjectURL = blob => { window.__capturedBlob = blob; return "blob:http://localhost:3210/smoke" }; HTMLAnchorElement.prototype.click = function () { window.__downloadName = this.download; }; document.querySelector("#export-laws").click(); })()');
  const exported = await evaluate('window.__capturedBlob.text().then(text => JSON.parse(text))');
  assert.equal(exported.version, 2);
  assert.ok(exported.laws.some(law => law.title === 'Импортированный акт'));
  await evaluate('document.querySelector("[data-page=create]").click(); document.querySelector("#word-btn").click()');
  const downloaded = await evaluate('window.__capturedBlob.arrayBuffer().then(buffer => ({ mime: window.__capturedBlob.type, signature: [...new Uint8Array(buffer).slice(0, 4)].join(",") }))');
  assert.equal(downloaded.signature, '80,75,3,4');
  assert.match(downloaded.mime, /wordprocessingml/);
  assert.match(await evaluate('window.__downloadName'), /2018-10-10-Методист\.docx$/);
  await evaluate('document.querySelector("[data-page=archive]").click(); document.querySelector(".archive-word").click()');
  assert.equal(await evaluate('window.__capturedBlob.arrayBuffer().then(buffer => [...new Uint8Array(buffer).slice(0, 4)].join(","))'), '80,75,3,4');
  const count = await evaluate('JSON.parse(localStorage.getItem("instruction-studio-v1-archive")).length');
  await evaluate('(() => { document.querySelector("[data-page=create]").click(); const original = Storage.prototype.setItem; Storage.prototype.setItem = function (key, value) { if (key === "instruction-studio-v1-archive") throw new DOMException("Quota exceeded", "QuotaExceededError"); return original.call(this, key, value); }; try { document.querySelector(".create-button").click(); } finally { Storage.prototype.setItem = original; } })()');
  assert.equal(await evaluate('JSON.parse(localStorage.getItem("instruction-studio-v1-archive")).length'), count);
  assert.match(await evaluate('document.querySelector("#toast").textContent'), /архив не сохранился/);
  assert.deepEqual(exceptions, []);
  console.log('Браузерная проверка пройдена: формы, нормативная база, импорт/экспорт, архив и скачивание DOCX.');
} finally { ws.close(); }
