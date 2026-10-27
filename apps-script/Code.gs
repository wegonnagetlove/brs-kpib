// ===== Скрипт сайта БРС 104 КПИБ =====
// 1) загрузка подтверждений с сайта на Google Диск (веб-приложение, doPost)
// 2) лист «Заявки»: редактируемая копия ответов формы, новые ответы дописываются сами
// 3) правка заявки с сайта: предмет, название, доп. информация, скрыть
// 4) лист «Расписание»: обновляется из выгрузки портала, положенной в папку на Диске

const SPREADSHEET_ID = '1Ibgbp67ott3PQ05LWFKuNsUUFS-343ctYclRpxP0UzI'; // таблица «БРС-для-сайта»
const SOURCE_SHEET = 'Ответы';   // лист с IMPORTRANGE из таблицы формы
const EDIT_SHEET = 'Заявки';     // лист, который можно править руками
const HEAD = ['Отметка времени','ФИО','Дата','Мероприятие','Роль','Доп. информация','Подтверждение','Предмет','Зачтено (да/нет)','Скрыть (да)'];

/* ---------- загрузка файлов ---------- */
const FOLDER_NAME = 'Подтверждения БРС 104 КПИБ';
const MAX_MB = 15;
function getFolder_() { const it = DriveApp.getFoldersByName(FOLDER_NAME); return it.hasNext() ? it.next() : DriveApp.createFolder(FOLDER_NAME); }
function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    if (d.action === 'edit') return json_(editApp_(d));
    const bytes = Utilities.base64Decode(d.file);
    if (bytes.length > MAX_MB * 1024 * 1024) throw new Error('Файл больше ' + MAX_MB + ' МБ');
    const type = String(d.type || '');
    if (!/^image\//.test(type) && type !== 'application/pdf') throw new Error('Можно только фото или PDF');
    const safe = s => String(s || '').replace(/[\\/:*?"<>|]+/g, ' ').slice(0, 80).trim();
    const name = [safe(d.who), safe(d.date), safe(d.title)].filter(String).join('_') + '_' + safe(d.name);
    const file = getFolder_().createFile(Utilities.newBlob(bytes, type, name));
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return json_({ ok: true, url: file.getUrl() });
  } catch (err) { return json_({ ok: false, error: String(err.message || err) }); }
}
function doGet() { return json_({ ok: true, status: 'работает' }); }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

/* ---------- лист «Заявки» ---------- */
// Запустите ОДИН раз вручную: создаст лист, скопирует все ответы и включит обновление каждые 5 минут.
function setup() {
  syncApps();
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'syncApps').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('syncApps').timeBased().everyMinutes(5).create();
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'importSchedule').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('importSchedule').timeBased().everyHours(1).create();
  getSchedFolder_();
}

function key_(ts, name) {
  const t = ts instanceof Date ? ts.getTime() : String(ts || '').trim();
  return t + '|' + String(name || '').trim().toLowerCase();
}

function syncApps() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const src = ss.getSheetByName(SOURCE_SHEET);
  let dst = ss.getSheetByName(EDIT_SHEET);
  if (!dst) {
    dst = ss.insertSheet(EDIT_SHEET);
    dst.getRange(1, 1, 1, HEAD.length).setValues([HEAD]).setFontWeight('bold').setBackground('#F1F3F4');
    dst.setFrozenRows(1);
    dst.setColumnWidth(2, 230); dst.setColumnWidth(4, 320); dst.setColumnWidth(6, 220); dst.setColumnWidth(8, 220);
  }
  const srcVals = src.getDataRange().getValues().slice(1).filter(r => String(r[1] || '').trim() !== '');
  const last = dst.getLastRow();
  const have = new Set();
  if (last > 1) dst.getRange(2, 1, last - 1, 2).getValues().forEach(r => have.add(key_(r[0], r[1])));
  // столбцы ответов формы: A отметка, B ФИО, C название, D дата, E роль, F доп., G подтверждение, H предмет, I зачтено (если есть)
  const add = srcVals.filter(r => !have.has(key_(r[0], r[1])))
    .map(r => [r[0], r[1], r[3], r[2], r[4], r[5], r[6], r[7], r[8] || '', '']);
  if (add.length) {
    dst.getRange(dst.getLastRow() + 1, 1, add.length, HEAD.length).setValues(add);
    dst.getRange(2, 3, dst.getLastRow() - 1, 1).setNumberFormat('dd.mm.yyyy');
  }
  return add.length;
}

/* ---------- правка заявки с сайта ---------- */
// строка ищется по отметке времени и ФИО; зачтённые (столбец I = да) не меняются
function editApp_(d) {
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const sh = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(EDIT_SHEET);
    if (!sh) throw new Error('Нет листа «Заявки». Запустите setup');
    const n = sh.getLastRow() - 1; if (n < 1) throw new Error('Заявок нет');
    const disp = sh.getRange(2, 1, n, HEAD.length).getDisplayValues();
    const nm = String(d.name || '').trim().toLowerCase(), ts = String(d.ts || '').trim(), old = String(d.old || '').trim();
    let i = disp.findIndex(r => r[0].trim() === ts && r[1].trim().toLowerCase() === nm);
    if (i < 0) { const c = disp.map((r, k) => k).filter(k => disp[k][1].trim().toLowerCase() === nm && disp[k][3].trim() === old); if (c.length === 1) i = c[0]; }
    if (i < 0) throw new Error('Заявка не найдена, обновите страницу');
    if (/^да/i.test(disp[i][8].trim())) throw new Error('Заявка уже зачтена, правка через старосту');
    const row = i + 2, s = d.set || {}, cut = v => String(v == null ? '' : v).slice(0, 1000);
    if (s.title != null && String(s.title).trim()) sh.getRange(row, 4).setValue(cut(s.title));
    if (s.info != null) sh.getRange(row, 6).setValue(cut(s.info));
    if (s.subj) sh.getRange(row, 8).setValue(cut(s.subj));
    if (s.proof) { const was = String(disp[i][6] || '').trim(); const keep = was && !/^нет$/i.test(was) ? was + ' ; ' : ''; sh.getRange(row, 7).setValue(cut(keep + s.proof)); }
    if (s.hide) sh.getRange(row, 10).setValue('да');
    return { ok: true, row: row };
  } finally { lock.releaseLock(); }
}

/* ---------- расписание из выгрузки портала ---------- */
// Выгрузку с портала (.xls или .xlsx) кладёте в папку «Расписание БРС 104 КПИБ» на Диске.
// Раз в час скрипт берёт самый свежий файл и пересобирает лист «Расписание».
// Запустить сразу: выбрать importSchedule и нажать «Выполнить».
// Нужна служба Drive API: слева «Сервисы» (+) → Drive API → Добавить.
const SCHED_FOLDER = 'Расписание БРС 104 КПИБ';
const SCHED_SHEET = 'Расписание';
const SCHED_SUBJ = { 'ПРОИ': 'ПРОИ', 'ПВДИ': 'ПВДИ', 'АПИП': 'АПИП', 'ФИЛПР': 'ФилПр' };

function getSchedFolder_() { const it = DriveApp.getFoldersByName(SCHED_FOLDER); return it.hasNext() ? it.next() : DriveApp.createFolder(SCHED_FOLDER); }

function importSchedule(e) {
  const files = getSchedFolder_().getFiles(); let f = null;
  while (files.hasNext()) { const x = files.next(); if (!f || x.getLastUpdated() > f.getLastUpdated()) f = x; }
  if (!f) return 'В папке нет файла';
  const props = PropertiesService.getScriptProperties(), stamp = f.getId() + '|' + f.getLastUpdated().getTime();
  // ручной запуск из редактора обновляет всегда, по расписанию только если файл новый
  if (e && e.triggerUid && props.getProperty('schedStamp') === stamp) return 'Новых выгрузок нет';

  let tmp = null, grid;
  if (f.getMimeType() === MimeType.GOOGLE_SHEETS) grid = SpreadsheetApp.openById(f.getId()).getSheets()[0].getDataRange().getDisplayValues();
  else {
    try { tmp = Drive.Files.copy({ name: 'tmp-расписание', mimeType: MimeType.GOOGLE_SHEETS }, f.getId()); }
    catch (e) { tmp = Drive.Files.copy({ title: 'tmp-расписание', mimeType: MimeType.GOOGLE_SHEETS }, f.getId()); }
    try { grid = SpreadsheetApp.openById(tmp.id).getSheets()[0].getDataRange().getDisplayValues(); }
    finally { DriveApp.getFileById(tmp.id).setTrashed(true); }
  }

  const P = s => { const m = /(\d{1,2})\.(\d{1,2})\.(\d{4})/.exec(String(s || '')); return m ? new Date(+m[3], +m[2] - 1, +m[1]) : null; };
  const DAYS = ['ПН', 'ВТ', 'СР', 'ЧТ', 'ПТ', 'СБ', 'ВС'];
  let dates = null, slot = '', hidden = null; const seen = {}, found = [];
  grid.forEach(row => {
    const a = String(row[0] || '').trim();
    if (DAYS.indexOf(a.toUpperCase()) >= 0) { dates = row.map(P); return; }
    if (/пара/i.test(a)) slot = (a.split('\n')[1] || '').trim();
    if (!dates) return;
    for (let c = 1; c < row.length; c++) {
      const v = String(row[c] || '').trim(), d = dates[c]; if (!v || !d) continue;
      const h = /Скрыто с\s*(\d{1,2}\.\d{1,2}\.\d{4})/.exec(v); if (h) { const hd = P(h[1]); if (!hidden || hd < hidden) hidden = hd; continue; }
      const m = /^(\d\d:\d\d-\d\d:\d\d)\s*\n\s*([\s\S]*)$/.exec(v), time = m ? m[1] : slot, body = (m ? m[2] : v).split('\n').map(x => x.trim()).filter(String);
      const code = (body[0] || '').split('[')[0].trim().toUpperCase(), subj = SCHED_SUBJ[code]; if (!subj) continue;
      const kind = ((/\[(.*?)\]/.exec(body[0]) || [])[1] || ''), k = subj + '|' + d.getTime() + '|' + time;
      if (seen[k]) continue; seen[k] = 1;
      found.push([subj, d, 1, time + ' ' + kind + ', ' + (body[body.length - 1] || ''), 'портал']);
    }
  });
  if (!found.length) return 'Не нашёл пар ПРОИ, ПВДИ, АПИП, ФилПр в файле';

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID); let sh = ss.getSheetByName(SCHED_SHEET);
  if (!sh) sh = ss.insertSheet(SCHED_SHEET);
  const old = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues() : [];
  const dayKey = (s, d) => s + '|' + (d instanceof Date ? Utilities.formatDate(d, ss.getSpreadsheetTimeZone(), 'yyyyMMdd') : '');
  const had = {}; old.forEach(r => { if (r[1] instanceof Date) had[dayKey(r[0], r[1])] = 1; });
  // новые даты из портала, которых раньше не было, занимают место строк «дата уточняется»
  const fresh = {}; found.forEach(r => { if (!had[dayKey(r[0], r[1])]) fresh[r[0]] = (fresh[r[0]] || 0) + 1; });
  let keep = [];
  if (hidden) {
    const und = {};
    old.forEach(r => {
      if (r[1] instanceof Date) { if (r[1] >= hidden) keep.push(r); }
      else if (String(r[0]).trim()) { und[r[0]] = (und[r[0]] || 0) + 1; if (und[r[0]] > (fresh[r[0]] || 0)) keep.push(r); }
    });
  }
  const order = Object.keys(SCHED_SUBJ).map(k => SCHED_SUBJ[k]);
  const all = found.concat(keep).sort((x, y) => order.indexOf(x[0]) - order.indexOf(y[0]) ||
    (x[1] instanceof Date ? x[1].getTime() : 9e15) - (y[1] instanceof Date ? y[1].getTime() : 9e15) || String(x[3]).localeCompare(String(y[3])));

  sh.clearContents(); sh.getRange(1, 1, sh.getMaxRows(), 5).setBackground(null);
  sh.getRange(1, 1, 1, 5).setValues([['Предмет', 'Дата', 'Пар', 'Занятие', 'Источник']]).setFontWeight('bold').setBackground('#F1F3F4');
  sh.getRange(2, 1, all.length, 5).setValues(all); sh.getRange(2, 2, all.length, 1).setNumberFormat('dd.mm.yyyy');
  all.forEach((r, i) => { if (!(r[1] instanceof Date)) sh.getRange(i + 2, 1, 1, 5).setBackground('#FEF7E0'); });
  sh.setFrozenRows(1);
  props.setProperty('schedStamp', stamp);
  const cnt = {}; all.forEach(r => cnt[r[0]] = (cnt[r[0]] || 0) + (+r[2] || 1));
  const msg = 'Обновлено из «' + f.getName() + '»: ' + order.map(s => s + ' ' + (cnt[s] || 0)).join(', ') + (hidden ? '. На портале скрыто с ' + Utilities.formatDate(hidden, ss.getSpreadsheetTimeZone(), 'dd.MM.yyyy') + ', дальше оставлены прежние строки' : '');
  Logger.log(msg); return msg;
}
