import { buildDocx, extractParagraphs } from './docx.js';

const STORAGE = 'instruction-studio-v1';
const MONITOR = 'instruction-studio-monitor-v1';
const RF = 'https://publication.pravo.gov.ru/';
const LNR = 'https://nslnr.su/zakonodatelstvo/normativno-pravovaya-baza/';

export const baseLaws = [
  { id: 'tk', region: 'РФ', kind: 'Трудовой кодекс', title: 'Трудовой кодекс Российской Федерации', number: '№ 197-ФЗ от 30.12.2001', url: RF, scope: 'all', verified: '' },
  { id: 'culture', region: 'РФ', kind: 'Закон РФ', title: 'Основы законодательства Российской Федерации о культуре', number: '№ 3612-1 от 09.10.1992', url: RF, scope: 'job', verified: '' },
  { id: 'sout', region: 'РФ', kind: 'Федеральный закон', title: 'О специальной оценке условий труда', number: '№ 426-ФЗ от 28.12.2013', url: RF, scope: 'safety', verified: '' },
  { id: 'ot-training', region: 'РФ', kind: 'Постановление Правительства', title: 'О порядке обучения по охране труда и проверки знания требований охраны труда', number: '№ 2464 от 24.12.2021', url: RF, scope: 'safety', verified: '' },
  { id: 'ot-instructions', region: 'РФ', kind: 'Приказ Минтруда', title: 'Об утверждении основных требований к порядку разработки и содержанию правил и инструкций по охране труда', number: '№ 772н от 29.10.2021', url: RF, scope: 'safety', verified: '' },
  { id: 'fire-law', region: 'РФ', kind: 'Федеральный закон', title: 'О пожарной безопасности', number: '№ 69-ФЗ от 21.12.1994', url: RF, scope: 'fire', verified: '' },
  { id: 'fire-reg', region: 'РФ', kind: 'Федеральный закон', title: 'Технический регламент о требованиях пожарной безопасности', number: '№ 123-ФЗ от 22.07.2008', url: RF, scope: 'fire', verified: '' },
  { id: 'fire-rules', region: 'РФ', kind: 'Постановление Правительства', title: 'Правила противопожарного режима в Российской Федерации', number: '№ 1479 от 16.09.2020', url: RF, scope: 'fire', verified: '' }
];

export const roles = [
  ['director', 'Директор учреждения культуры'], ['art', 'Художественный руководитель'],
  ['events', 'Организатор культурно-массовых мероприятий'], ['sound', 'Звукорежиссёр'],
  ['light', 'Художник по свету'], ['library', 'Библиотекарь'],
  ['museum', 'Музейный смотритель'], ['cleaner', 'Уборщик служебных помещений']
];

export const typeNames = { job: 'Должностная инструкция', safety: 'Инструкция по охране труда', fire: 'Инструкция о мерах пожарной безопасности' };

export function uniqueId(source = globalThis.crypto) {
  if (typeof source?.randomUUID === 'function') return source.randomUUID();
  const bytes = new Uint8Array(16);
  if (typeof source?.getRandomValues === 'function') source.getRandomValues(bytes);
  else for (let i = 0; i < bytes.length; i++) bytes[i] = Math.floor(Math.random() * 256);
  return [...bytes].map(byte => byte.toString(16).padStart(2, '0')).join('');
}

export function resolveRole(data) {
  const input = String(data.role ?? '').trim();
  const legacy = input === 'other' ? String(data.customRole ?? '').trim() : roles.find(([id]) => id === input)?.[1];
  const name = legacy || input || 'укажите должность';
  const roleId = roles.find(([, title]) => title.toLocaleLowerCase('ru') === name.toLocaleLowerCase('ru'))?.[0] || 'other';
  return { name, roleId };
}

const duties = {
  director: ['Организует деятельность учреждения в соответствии с уставом и утверждёнными планами.', 'Обеспечивает организацию труда, сохранность имущества и доступность услуг культуры.', 'Организует выполнение требований охраны труда и пожарной безопасности, назначает ответственных лиц в пределах полномочий.'],
  art: ['Разрабатывает художественную концепцию программ и репертуарный план.', 'Координирует работу творческих коллективов, репетиций и выступлений.', 'Согласовывает проведение мероприятий с ответственными за площадку и техническое обеспечение.'],
  events: ['Планирует и готовит культурно-массовые мероприятия, взаимодействует с участниками и подрядчиками.', 'Контролирует готовность площадки, доступность проходов и порядок допуска посетителей.', 'Согласовывает технический план мероприятия с ответственными службами.'],
  sound: ['Подготавливает, настраивает и эксплуатирует звуковое оборудование по утверждённому плану мероприятия.', 'Проверяет исправность соединений, кабелей и аппаратуры перед началом работы.', 'Сообщает о неисправностях и прекращает использование неисправного оборудования.'],
  light: ['Обеспечивает световое оформление мероприятий и настройку осветительного оборудования.', 'Проверяет крепления приборов, кабельные линии и исправность пультов в пределах своей компетенции.', 'Согласовывает монтаж и демонтаж оборудования с ответственным за площадку.'],
  library: ['Организует обслуживание читателей и учёт библиотечного фонда.', 'Обеспечивает сохранность документов и соблюдение правил пользования библиотекой.', 'Готовит выставки, встречи и культурно-просветительские мероприятия.'],
  museum: ['Осуществляет наблюдение за экспозицией и соблюдением правил посещения.', 'Информирует ответственное лицо о повреждении экспонатов и нарушениях условий хранения.', 'Помогает посетителям ориентироваться в залах и на маршрутах движения.'],
  cleaner: ['Выполняет уборку закреплённых помещений по установленному графику.', 'Использует инвентарь и моющие средства согласно инструкциям изготовителя.', 'Сообщает о повреждениях оборудования, разливах и иных опасных условиях.'],
  other: ['Выполняет трудовые функции, закреплённые трудовым договором и локальными актами учреждения.', 'Соблюдает утверждённые процедуры работы и сообщает руководителю о препятствиях для выполнения задач.']
};

export function escapeHtml(value) {
  return String(value ?? '').replace(/[&<>"']/g, char => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[char]);
}

export function validateLaw(value) {
  if (!value || typeof value !== 'object') return null;
  const fields = ['id', 'region', 'kind', 'title', 'number', 'url', 'scope', 'verified'];
  if (fields.some(key => typeof value[key] !== 'string')) return null;
  if (!/^[\w-]{1,100}$/.test(value.id) || !['РФ', 'ЛНР'].includes(value.region) || !['all', 'job', 'safety', 'fire'].includes(value.scope)) return null;
  if (!value.kind.trim() || !value.title.trim() || !value.number.trim() || value.kind.length > 80 || value.title.length > 250 || value.number.length > 100) return null;
  const status = value.status || 'review';
  const effectiveFrom = value.effectiveFrom || '';
  const effectiveTo = value.effectiveTo || '';
  if (!['review', 'active', 'repealed'].includes(status)) return null;
  if ([value.verified, effectiveFrom, effectiveTo].some(date => date && (!/^\d{4}-\d{2}-\d{2}$/.test(date) || Number.isNaN(Date.parse(date)) || new Date(date).toISOString().slice(0, 10) !== date))) return null;
  if (effectiveFrom && effectiveTo && effectiveFrom > effectiveTo) return null;
  if (status === 'active' && (!value.verified || !effectiveFrom || value.verified > new Date().toISOString().slice(0, 10))) return null;
  try { const url = new URL(value.url); if (url.protocol !== 'https:' || url.username || url.password || !url.hostname.includes('.')) return null; }
  catch { return null; }
  if (status === 'active') {
    const url = new URL(value.url);
    const host = url.hostname.toLowerCase();
    const official = host === 'pravo.gov.ru' || host.endsWith('.pravo.gov.ru') || host === 'government.ru'
      || host.endsWith('.gov.ru') || host === 'nslnr.su' || host.endsWith('.nslnr.su') || host === 'government.lnr.ru';
    if (!official || url.pathname === '/' || (host === 'nslnr.su' && url.pathname.replace(/\/$/, '') === new URL(LNR).pathname.replace(/\/$/, ''))) return null;
  }
  return { ...Object.fromEntries(fields.map(key => [key, value[key]])), status, effectiveFrom, effectiveTo, appliesLnr: value.appliesLnr === true || value.appliesLnr === 'on' };
}

export function applicableLaws(laws, type, jurisdiction, date) {
  const at = /^\d{4}-\d{2}-\d{2}$/.test(date || '') ? date : new Date().toISOString().slice(0, 10);
  return laws.map(validateLaw).filter(law => law && law.status === 'active' && (law.scope === 'all' || law.scope === type)
    && law.effectiveFrom <= at && (!law.effectiveTo || at <= law.effectiveTo)
    && (jurisdiction === 'ЛНР' ? (law.region === 'ЛНР' || law.appliesLnr) : law.region === 'РФ'));
}

export function documentHtml(data, laws) {
  const e = escapeHtml;
  const type = ['job', 'safety', 'fire'].includes(data.type) ? data.type : 'job';
  const { name: role, roleId } = resolveRole(data);
  const org = data.org?.trim() || 'Наименование учреждения';
  const location = data.location?.trim() || 'указать рабочее место / объект';
  const extra = data.details?.trim();
  const title = { job: 'ДОЛЖНОСТНАЯ ИНСТРУКЦИЯ', safety: 'ИНСТРУКЦИЯ ПО ОХРАНЕ ТРУДА', fire: 'ИНСТРУКЦИЯ О МЕРАХ ПОЖАРНОЙ БЕЗОПАСНОСТИ' }[type];
  const sections = [];
  if (type === 'job') {
    sections.push(['1. Общие положения', [`Настоящая должностная инструкция определяет обязанности, права и ответственность работника по должности «${role}» в ${org}.`, 'Работник назначается и освобождается от должности в установленном работодателем порядке; подчинённость, требования к квалификации и порядок замещения определяются штатным расписанием, трудовым договором и локальными актами.', 'Работник руководствуется законодательством РФ и применимыми актами ЛНР, уставом учреждения, правилами внутреннего трудового распорядка и настоящей инструкцией.']]);
    sections.push(['2. Должностные обязанности', [...(duties[roleId] || duties.other), `Выполняет задачи на объекте: ${location}.`, 'Соблюдает требования охраны труда, пожарной безопасности и правила обращения с персональными данными в пределах своих обязанностей.']]);
    sections.push(['3. Права', ['Получать информацию и необходимые ресурсы для выполнения трудовой функции.', 'Вносить предложения по улучшению работы и сообщать руководителю о выявленных недостатках.', 'На безопасные условия труда и обучение в установленном порядке.']]);
    sections.push(['4. Ответственность и взаимодействие', ['Несёт ответственность в пределах, установленных трудовым законодательством и локальными актами, за ненадлежащее выполнение обязанностей.', 'Взаимодействует с руководителем, коллегами и ответственными за проведение мероприятий по вопросам своей компетенции.']]);
  } else if (type === 'safety') {
    sections.push(['1. Общие требования охраны труда', [`Инструкция предназначена для должности «${role}»; рабочее место: ${location}.`, 'К работе допускаются лица после необходимых инструктажей, обучения и проверки знания требований охраны труда в случаях, предусмотренных законодательством и локальными актами.', 'Работник соблюдает требования безопасности, применяет выданные средства индивидуальной защиты по назначению и сообщает руководителю об опасностях и происшествиях.', 'Опасные и вредные факторы, СИЗ и меры защиты уточняются по результатам СОУТ, оценке профессиональных рисков и документации на оборудование.']]);
    sections.push(['2. Перед началом работы', ['Осмотреть рабочее место, проходы, освещение и исправность используемого оборудования.', 'Проверить доступность средств связи и свободный доступ к выходам; при обнаружении неисправностей сообщить руководителю и не приступать к опасной работе.']]);
    sections.push(['3. Во время работы', ['Применять оборудование по назначению и инструкциям изготовителя; не выполнять работы, к которым работник не допущен.', 'Не загромождать проходы и выходы, учитывать присутствие посетителей и участников мероприятий.', 'При возникновении угрозы жизни и здоровью прекратить работу, предупредить находящихся рядом людей и сообщить руководителю.']]);
    sections.push(['4. В аварийных ситуациях', ['При пожаре или иной опасности сообщить по телефону 112, действовать по плану эвакуации и указаниям ответственных лиц.', 'О несчастном случае немедленно уведомить руководителя, организовать вызов помощи; первую помощь оказывать при наличии навыков, не подвергая себя опасности.']]);
    sections.push(['5. По окончании работы', ['Выключить используемое оборудование в соответствии с инструкциями изготовителя и локальным порядком.', 'Привести рабочее место в порядок и сообщить ответственному о неисправностях и происшествиях.']]);
  } else {
    sections.push(['1. Общие положения и объект', [`Инструкция устанавливает организационные меры пожарной безопасности для объекта: ${location}${data.address?.trim() ? `, адрес: ${data.address.trim()}` : ''}.`, `Ответственный за пожарную безопасность: ${data.firePerson?.trim() || 'назначается распорядительным документом учреждения'}.`, 'До утверждения необходимо уточнить категорию помещений (если требуется), противопожарное оборудование, план эвакуации, число людей и особенности эксплуатации здания.']]);
    sections.push(['2. Порядок содержания помещений', ['Содержать эвакуационные пути и выходы свободными; не блокировать двери и доступ к средствам пожаротушения.', 'Использовать электрооборудование в соответствии с документацией изготовителя; не эксплуатировать повреждённые приборы и проводку.', 'Порядок осмотра, закрытия помещений и проведения мероприятий устанавливается локальными актами учреждения с учётом фактического объекта.']]);
    sections.push(['3. Мероприятия с посетителями', ['Перед мероприятием ответственным лицам проверить доступность выходов, средств связи и оповещения.', 'Соблюдать допустимое число посетителей и порядок размещения оборудования по проектной документации и локальным правилам.', 'Обеспечить участие назначенных работников в организации выхода посетителей, включая маломобильных граждан.']]);
    sections.push(['4. Действия при пожаре', ['Немедленно сообщить о пожаре по телефону 101 или 112, указать адрес, место возгорания и свои данные.', 'Включить оповещение (если предусмотрено), организовать эвакуацию по утверждённому плану, сообщить ответственному и встретить пожарные подразделения.', 'При отсутствии угрозы жизни использовать первичные средства пожаротушения, если работник обучен и это безопасно.', `Пути эвакуации и место сбора: ${data.evacuation?.trim() || 'указать по действующему плану эвакуации объекта'}.`]]);
  }
  if (extra) sections.push(['Особенности объекта и работы (проверить и конкретизировать)', [extra]]);
  const jurisdiction = data.jurisdiction === 'РФ' ? 'РФ' : 'ЛНР';
  const relevant = applicableLaws(laws, type, jurisdiction, data.date);
  const regional = relevant.some(law => law.region === 'ЛНР');
  const legalList = relevant.map(law => `<li>${e(law.kind)} ${e(law.number)} «${e(law.title)}» — <a href="${e(law.url)}" target="_blank" rel="noopener noreferrer">официальный текст</a> (проверено пользователем ${e(law.verified)})</li>`).join('');
  return `<div class="doc-head"><div>${e(org)}</div><div class="approval">УТВЕРЖДАЮ<br>Руководитель: ${e(data.head?.trim() || '________________')}<br>Подпись: ____________<br>Дата: ${e(data.date || '____________')}</div></div><div class="doc-title">${title}<small>для ${e(role)}${data.number?.trim() ? ` · № ${e(data.number.trim())}` : ''}</small></div><div class="doc-place">${e(data.city?.trim() || 'Населённый пункт не указан')} · ${e(data.date?.slice(0, 4) || '20__')} г.</div>${sections.map(([heading, items]) => `<section class="doc-section"><h3>${e(heading)}</h3>${items.map((text, index) => `<p>${index + 1}. ${e(text)}</p>`).join('')}</section>`).join('')}<section class="doc-section doc-sources"><h3>Нормативная основа — ${e(jurisdiction)} · проверить перед утверждением</h3><p>Включены только акты, отмеченные пользователем как проверенные и применимые на дату документа. Это не автоматическое подтверждение актуальности редакции${data.date ? ` на ${e(data.date)}` : '; дата документа не указана, отбор сделан на текущую дату'}. Уточните локальные акты и особенности переходного регулирования в ЛНР.</p><ul>${legalList || '<li>Проверенных применимых актов на указанную дату нет. Добавьте их через нормативную базу после изучения официального текста.</li>'}</ul>${jurisdiction === 'ЛНР' && !regional ? '<p>Проверенных региональных актов ЛНР в базе нет.</p>' : ''}</section><div class="doc-sign">С инструкцией ознакомлен(а): __________________ / __________________<br>Дата ознакомления: «___» __________ 20___ г.</div><div class="doc-draft">ПРОЕКТ · требует проверки и утверждения ответственным специалистом</div>`;
}

function stored(key, fallback) { try { return JSON.parse(localStorage.getItem(key)) ?? fallback; } catch { return fallback; } }
function save(key, value) { try { localStorage.setItem(key, JSON.stringify(value)); return true; } catch { toast('Не удалось сохранить данные в браузере'); return false; } }
let customLaws = [];
let archive = [];
let form;
let currentType = 'job';
let toastTimer;
function toast(message) { const node = document.getElementById('toast'); node.textContent = message; node.classList.add('visible'); clearTimeout(toastTimer); toastTimer = setTimeout(() => node.classList.remove('visible'), 3500); }
function allLaws() { const map = new Map(baseLaws.map(law => [law.id, law])); customLaws.forEach(law => map.set(law.id, law)); return [...map.values()]; }
function fields() { return Object.fromEntries(new FormData(form).entries()); }
function documentFormValid() {
  for (const name of ['org', 'role']) {
    const input = form.elements[name];
    input.setCustomValidity(input.value.trim() ? '' : 'Введите значение, состоящее не только из пробелов');
  }
  return form.reportValidity();
}
function updatePreview() {
  const data = { ...fields(), type: currentType };
  document.getElementById('document-preview').innerHTML = documentHtml(data, allLaws());
  document.getElementById('preview-heading').textContent = typeNames[currentType];
  document.getElementById('preview-status').textContent = 'Черновик · обновляется автоматически';
  save(STORAGE + '-draft', data);
}
function syncFields() { document.getElementById('fire-fields').hidden = currentType !== 'fire'; }
function selectType(type) {
  currentType = type;
  document.querySelectorAll('.type').forEach(node => {
    const selected = node.dataset.type === type;
    node.classList.toggle('selected', selected);
    node.setAttribute('aria-pressed', String(selected));
  });
  syncFields(); updatePreview();
}
function renderLaws() {
  const search = document.getElementById('law-search').value.toLocaleLowerCase('ru');
  const region = document.getElementById('law-filter').value;
  const filtered = allLaws().filter(law => (region === 'all' || region === law.region) && (law.title + law.number + law.kind).toLocaleLowerCase('ru').includes(search));
  document.getElementById('law-count').textContent = allLaws().length;
  document.getElementById('law-list').innerHTML = filtered.length ? filtered.map(law => {
    const status = law.status === 'active' ? `Проверен пользователем ${escapeHtml(law.verified)} · действует с ${escapeHtml(law.effectiveFrom)}` : law.status === 'repealed' ? 'Утратил силу / не применять' : 'Не проверен · в документы не включается';
    return `<div class="law-row"><div class="law-symbol">§</div><div class="law-info"><div class="law-meta"><span class="region">${escapeHtml(law.region)}</span> ${escapeHtml(law.kind)} · ${escapeHtml(law.number)}</div><h3>${escapeHtml(law.title)}</h3><div class="law-foot">${status} · <a target="_blank" rel="noopener noreferrer" href="${escapeHtml(law.url)}">Открыть источник ↗</a></div></div><button class="button small secondary edit-law" data-id="${escapeHtml(law.id)}">Изменить</button>${baseLaws.some(item => item.id === law.id) ? '' : `<button class="button small danger delete-law" data-id="${escapeHtml(law.id)}">Удалить</button>`}</div>`;
  }).join('') : '<div class="empty">Акты по заданным условиям не найдены.</div>';
}
function openLaw(law) {
  const dialog = document.getElementById('law-dialog');
  const editor = document.getElementById('law-form');
  editor.reset();
  editor.elements.id.value = law?.id || '';
  if (law) {
    for (const key of ['region', 'kind', 'title', 'number', 'url', 'scope', 'verified', 'status', 'effectiveFrom', 'effectiveTo']) editor.elements[key].value = law[key] || '';
    editor.elements.appliesLnr.checked = law.appliesLnr;
  }
  document.getElementById('dialog-title').textContent = law ? 'Изменить акт' : 'Добавить акт';
  dialog.showModal();
}
function download(name, content, mime) { const url = URL.createObjectURL(new Blob([content], { type: mime })); const link = document.createElement('a'); link.href = url; link.download = name; link.click(); setTimeout(() => URL.revokeObjectURL(url), 5000); }
function wordFile(html, date, type, role) {
  const documentDate = /^\d{4}-\d{2}-\d{2}$/.test(date || '') ? date : 'bez-daty';
  const position = String(role || '').trim().slice(0, 48).replace(/[\\/:*?"<>|\u0000-\u001f]/g, '').replace(/\s+/g, '-');
  download(`instrukciya-${type}-${documentDate}${position ? `-${position}` : ''}.docx`, buildDocx(extractParagraphs(html)), 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
}
function renderArchive() {
  document.getElementById('archive-list').innerHTML = archive.length ? archive.map(item => `<div class="law-row"><div class="law-symbol">▤</div><div class="law-info"><div class="law-meta">${escapeHtml(item.date || 'Без даты')} · создано ${escapeHtml(new Date(item.createdAt).toLocaleString('ru-RU'))}</div><h3>${escapeHtml(typeNames[item.type])} — ${escapeHtml(item.role)}</h3><div class="law-foot">${escapeHtml(item.org)} · ${escapeHtml(item.jurisdiction)}</div></div><button type="button" class="button small primary archive-word" data-id="${escapeHtml(item.id)}">↓ Word</button><button type="button" class="button small danger archive-delete" data-id="${escapeHtml(item.id)}">Удалить</button></div>`).join('') : '<div class="empty">Пока нет сохранённых документов. Заполните форму и нажмите «Создать инструкцию».</div>';
}
function renderSources(statuses = []) {
  const previous = stored(MONITOR, {});
  const sources = [{ id: 'rf', name: 'Портал правовой информации РФ', url: RF }, { id: 'lnr', name: 'Нормативно-правовая база Народного Совета ЛНР', url: LNR }];
  document.getElementById('source-list').innerHTML = sources.map(source => {
    const status = statuses.find(item => item.id === source.id);
    const seen = previous[source.id];
    let message = seen ? `Последняя успешная проверка: ${escapeHtml(new Date(seen.checkedAt).toLocaleString('ru-RU'))}` : 'Ещё не проверялся';
    let badge = 'Ожидает проверки';
    if (status?.error) { badge = 'Недоступен'; message = `Ошибка доступа к источнику: ${escapeHtml(status.error)}. Попробуйте позже или откройте сайт вручную.`; }
    else if (status?.fingerprint) { badge = !seen ? 'Первый снимок' : seen.fingerprint === status.fingerprint ? 'Без изменений' : 'Страница изменилась'; message = `Проверено: ${escapeHtml(new Date(status.checkedAt).toLocaleString('ru-RU'))}`; }
    return `<div class="source-row"><span class="source-icon">↗</span><div><h3>${escapeHtml(source.name)}</h3><p>${message}</p><a href="${escapeHtml(source.url)}" target="_blank" rel="noopener noreferrer">Открыть официальный источник ↗</a></div><span class="status-pill ${badge === 'Страница изменилась' ? 'changed' : ''}">${badge}</span></div>`;
  }).join('');
}

function init() {
  form = document.getElementById('document-form');
  const saved = stored(STORAGE + '-laws', []);
  customLaws = Array.isArray(saved) ? saved.map(validateLaw).filter(Boolean).slice(0, 1000) : [];
  const savedArchive = stored(STORAGE + '-archive', []);
  archive = Array.isArray(savedArchive) ? savedArchive.filter(item => item && typeof item.html === 'string' && typeof item.id === 'string' && typeNames[item.type] && typeof item.createdAt === 'string' && typeof item.org === 'string' && typeof item.role === 'string') : [];
  document.getElementById('role-suggestions').innerHTML = roles.map(([, name]) => `<option value="${name}"></option>`).join('');
  const draft = stored(STORAGE + '-draft', {});
  for (const [key, value] of Object.entries(draft)) if (key !== 'type' && form.elements[key] && typeof value === 'string') form.elements[key].value = value;
  form.elements.role.value = resolveRole(draft).name === 'укажите должность' ? '' : resolveRole(draft).name;
  currentType = ['job', 'safety', 'fire'].includes(draft.type) ? draft.type : 'job';
  selectType(currentType); renderLaws(); renderArchive(); renderSources();
  document.querySelectorAll('.nav').forEach(node => node.addEventListener('click', () => { document.querySelectorAll('.nav, .page').forEach(item => item.classList.remove('active')); node.classList.add('active'); document.getElementById(node.dataset.page).classList.add('active'); }));
  document.querySelectorAll('.type').forEach(node => node.addEventListener('click', () => selectType(node.dataset.type)));
  form.addEventListener('input', event => {
    if (['org', 'role'].includes(event.target.name)) event.target.setCustomValidity('');
    syncFields(); updatePreview();
  });
  form.addEventListener('change', () => { syncFields(); updatePreview(); });
  form.addEventListener('submit', event => {
    event.preventDefault();
    if (!documentFormValid()) return;
    updatePreview();
    const data = { ...fields(), type: currentType };
    const snapshot = { id: uniqueId(), createdAt: new Date().toISOString(), html: document.getElementById('document-preview').innerHTML, type: currentType, date: data.date, role: resolveRole(data).name, org: data.org, jurisdiction: data.jurisdiction || 'ЛНР' };
    const next = [snapshot, ...archive];
    const archived = save(STORAGE + '-archive', next);
    if (archived) { archive = next; renderArchive(); }
    document.getElementById('preview-status').textContent = 'Инструкция создана · проверьте перед утверждением';
    document.getElementById('preview-column').scrollIntoView({ behavior: 'smooth', block: 'start' });
    toast(archived ? 'Инструкция создана и сохранена в архиве. Доступна загрузка Word.' : 'Инструкция создана, но архив не сохранился. Скачайте Word.');
  });
  document.getElementById('word-btn').addEventListener('click', () => {
    if (!documentFormValid()) return;
    wordFile(document.getElementById('document-preview').innerHTML, form.elements.date.value, currentType, form.elements.role.value);
  });
  document.getElementById('archive-list').addEventListener('click', event => {
    const id = event.target.closest('[data-id]')?.dataset.id;
    const item = archive.find(record => record.id === id);
    if (!item) return;
    if (event.target.closest('.archive-word')) wordFile(item.html, item.date, item.type, item.role);
    if (event.target.closest('.archive-delete') && confirm('Удалить сохранённую инструкцию?')) {
      const next = archive.filter(record => record.id !== id);
      if (save(STORAGE + '-archive', next)) { archive = next; renderArchive(); }
    }
  });
  document.getElementById('print-btn').addEventListener('click', () => { if (!documentFormValid()) return; window.print(); });
  document.getElementById('download-btn').addEventListener('click', () => {
    if (!documentFormValid()) return;
    const page = document.getElementById('document-preview').innerHTML;
    const css = '.paper{max-width:780px;margin:auto;font:14px/1.6 Georgia,serif;color:#20252a}.doc-head{min-height:115px;display:flex;justify-content:space-between;gap:24px}.approval{text-align:left;min-width:230px}.doc-title{text-align:center;font-weight:bold;font-size:19px;margin:38px 0 10px}.doc-title small{display:block;font-size:14px;text-transform:none}.doc-place{text-align:center;color:#555}.doc-section{margin-top:26px}.doc-section h3{font-size:15px}.doc-section p{text-align:justify}.doc-sign{margin-top:40px}.doc-draft{margin-top:35px;text-align:center;color:#777;font-size:11px}';
    download(`инструкция-${currentType}.html`, `<!doctype html><html lang="ru"><meta charset="utf-8"><title>Проект инструкции</title><style>${css}</style><article class="paper">${page}</article></html>`, 'text/html;charset=utf-8');
  });
  document.getElementById('law-search').addEventListener('input', renderLaws);
  document.getElementById('law-filter').addEventListener('change', renderLaws);
  document.getElementById('add-law').addEventListener('click', () => openLaw());
  document.getElementById('close-dialog').addEventListener('click', () => document.getElementById('law-dialog').close());
  document.getElementById('cancel-dialog').addEventListener('click', () => document.getElementById('law-dialog').close());
  document.getElementById('law-list').addEventListener('click', event => {
    const edit = event.target.closest('.edit-law');
    if (edit) openLaw(allLaws().find(law => law.id === edit.dataset.id));
    const remove = event.target.closest('.delete-law');
    if (remove && confirm('Удалить этот акт из нормативной базы?')) {
      const next = customLaws.filter(law => law.id !== remove.dataset.id);
      if (save(STORAGE + '-laws', next)) { customLaws = next; renderLaws(); updatePreview(); }
    }
  });
  document.getElementById('law-form').addEventListener('submit', event => {
    event.preventDefault();
    const values = Object.fromEntries(new FormData(event.currentTarget).entries());
    if (!values.id) values.id = `custom-${uniqueId()}`;
    const law = validateLaw(values);
    if (!law) { toast('Для действующего акта нужны дата проверки, начало действия и ссылка на конкретный акт'); return; }
    const next = [...customLaws.filter(item => item.id !== law.id), law];
    if (!save(STORAGE + '-laws', next)) return;
    customLaws = next;
    document.getElementById('law-dialog').close(); renderLaws(); updatePreview(); toast('Акт сохранён');
  });
  document.getElementById('export-laws').addEventListener('click', () => download('normativnaya-baza.json', JSON.stringify({ format: 'instruction-studio-laws', version: 2, laws: customLaws }, null, 2), 'application/json;charset=utf-8'));
  document.getElementById('import-laws').addEventListener('change', async event => {
    const file = event.target.files[0]; event.target.value = '';
    if (!file) return;
    if (file.size > 2_000_000) { toast('Файл слишком большой (максимум 2 МБ)'); return; }
    try {
      const payload = JSON.parse(await file.text());
      if (payload.format !== 'instruction-studio-laws' || ![1, 2].includes(payload.version) || !Array.isArray(payload.laws) || payload.laws.length > 1000) throw Error();
      const parsed = payload.laws.map(validateLaw);
      if (parsed.some(item => !item)) throw Error();
      const map = new Map(customLaws.map(law => [law.id, law]));
      parsed.forEach(law => map.set(law.id, law));
      const merged = [...map.values()];
      if (merged.length > 1000) throw Error();
      if (!save(STORAGE + '-laws', merged)) return;
      customLaws = merged; renderLaws(); updatePreview(); toast(`Импортировано актов: ${parsed.length}`);
    } catch { toast('Неверный формат файла или данных актов'); }
  });
  document.getElementById('check-updates').addEventListener('click', async event => {
    const button = event.currentTarget; button.disabled = true; button.textContent = 'Проверяем…';
    try {
      const response = await fetch('/api/check', { cache: 'no-store' });
      if (!response.ok) throw Error('Сервис недоступен');
      const result = await response.json();
      if (!Array.isArray(result) || result.length !== 2) throw Error('Некорректный ответ');
      renderSources(result);
      const latest = stored(MONITOR, {});
      result.forEach(item => { if (item.fingerprint && ['rf', 'lnr'].includes(item.id)) latest[item.id] = { fingerprint: item.fingerprint, checkedAt: item.checkedAt }; });
      save(MONITOR, latest);
      toast('Проверка завершена. Изучите результаты и официальные тексты.');
    } catch { toast('Запустите приложение через npm start для проверки источников'); }
    finally { button.disabled = false; button.textContent = '↻ Проверить сейчас'; }
  });
}

if (typeof document !== 'undefined') init();
