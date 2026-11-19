/* ============================================================
   ЖУРНАЛ ДВИЖЕНИЯ СИ ПРИ ПОВЕРКЕ В ЦСМ - серверный скрипт v1.0
   АО "Новосибирская птицефабрика", метрология

   Установка: вставить этот файл в Apps Script и запустить setup -
   он сам создаст таблицу со всеми листами (см. README.md).
   Листы таблицы:
     Приборы        - сам журнал (одна строка = один прибор)
     Пользователи   - ФИО, пароль, роль (метролог / администратор)
     Места          - список цехов и мест для выпадающего меню
     Журнал правок  - кто, когда и что менял
     Настройки      - дни до блокировки, порог подсветки, название

   Правила:
     - Статус и «Дней всего» считаются так же, как формулы в Excel
     - Даты этапов должны идти по порядку, этап 1 обязателен
     - Через N дней (Настройки) после выдачи в цех метролог больше
       не может менять запись. Администратор может - и в колонку
       «Изменения» пишется пометка о корректировке
     - Удалять записи и смотреть журнал правок - только администратор
   ============================================================ */

var SH_DATA = 'Приборы';
var SH_USERS = 'Пользователи';
var SH_PLACES = 'Места';
var SH_LOG = 'Журнал правок';
var SH_SET = 'Настройки';

var FIELDS = ['id', 'num', 'name', 'serial', 'place', 'd1', 'd2', 'd3', 'd4',
              'days', 'status', 'result', 'note', 'changes',
              'createdBy', 'createdAt', 'updatedBy', 'updatedAt', 'clientId'];
var NCOL = FIELDS.length;
var COL = {};
FIELDS.forEach(function (f, i) { COL[f] = i + 1; });

var DATES = ['d1', 'd2', 'd3', 'd4'];
var LABEL = {
  name: 'Наименование СИ', serial: 'Зав. №', place: 'Цех / место',
  d1: '1. Цех→Метрология', d2: '2. Метрология→ЦСМ',
  d3: '3. ЦСМ→Метрология', d4: '4. Метрология→Цех',
  result: 'Примечание', note: 'Дополнение'
};
var EDITABLE = ['name', 'serial', 'place', 'd1', 'd2', 'd3', 'd4', 'result', 'note'];
var RESULTS = ['', 'Поверен', 'Не прошёл проверку'];

// ============================================================
// ▶ ЗАПУСК В РЕДАКТОРЕ: выберите в списке функций вверху
//     setup  - создать таблицу (первый раз) и загрузить данные
//     check  - проверить, всё ли в порядке
//   и нажмите «Выполнить». Остальные функции служебные
//   (имя с «_» на конце) - в списке их не видно.
// ============================================================

function setup() { runSetup_(); }

function check() { runCheck_(); }

// ============================================================
// ВХОД
// ============================================================

function doGet(e) {
  var p = (e && e.parameter) || {};
  return output_(route_(p), p.callback);
}

function output_(obj, callback) {
  var json = JSON.stringify(obj);
  if (callback && /^[A-Za-z_$][\w$]*$/.test(callback)) {
    return ContentService.createTextOutput(callback + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function route_(p) {
  try {
    var ctx = context_();
    var user = findUser_(ctx, p.password);
    if (!user) return { success: false, error: 'Неверный пароль' };

    switch (p.action) {
      case 'init':   return init_(ctx, user);
      case 'list':   return { success: true, records: readAll_(ctx).records, today: ctx.today };
      case 'save':   return withLock_(function () { return saveRecord_(context_(), user, parseData_(p.data)); });
      case 'bulk':   return withLock_(function () { return bulkStage_(context_(), user, parseData_(p.data)); });
      case 'delete': return withLock_(function () { return deleteRecord_(context_(), user, p.id); });
      case 'log':    return getLog_(ctx, user, p.id);
    }
    return { success: false, error: 'Неизвестное действие' };
  } catch (err) {
    return { success: false, error: 'Ошибка сервера: ' + (err && err.message ? err.message : err) };
  }
}

function parseData_(text) {
  try { return JSON.parse(text || '{}'); } catch (e) { return {}; }
}

function withLock_(fn) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(25000); } catch (e) {
    return { success: false, error: 'Сервер занят, повторите через несколько секунд' };
  }
  try { return fn(); } finally { lock.releaseLock(); }
}

// ============================================================
// ОКРУЖЕНИЕ, НАСТРОЙКИ, ПОЛЬЗОВАТЕЛИ
// ============================================================

// Таблица: та, к которой привязан скрипт, или созданная функцией setup()
function openBook_() {
  var ss = null;
  try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) { ss = null; }
  if (ss) return ss;
  var id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  if (!id) throw new Error('Таблица не создана. Запустите функцию setup в редакторе скриптов.');
  return SpreadsheetApp.openById(id);
}

function context_() {
  var ss = openBook_();
  var tz = ss.getSpreadsheetTimeZone() || 'Asia/Novosibirsk';
  var now = new Date();
  var ctx = {
    ss: ss, tz: tz, now: now,
    today: Utilities.formatDate(now, tz, 'yyyy-MM-dd'),
    settings: { lockDays: 10, warnDays: 20, org: 'АО "Новосибирская птицефабрика"' }
  };
  var sh = ss.getSheetByName(SH_SET);
  if (sh && sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(function (r) {
      var k = String(r[0]).trim().toLowerCase();
      if (k.indexOf('блокир') >= 0 && Number(r[1]) >= 0) ctx.settings.lockDays = Number(r[1]);
      else if (k.indexOf('порог') >= 0 && Number(r[1]) > 0) ctx.settings.warnDays = Number(r[1]);
      else if (k.indexOf('назван') >= 0 && String(r[1]).trim()) ctx.settings.org = String(r[1]).trim();
    });
  }
  return ctx;
}

function readUsers_(ctx) {
  var sh = ctx.ss.getSheetByName(SH_USERS);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues()
    .filter(function (r) {
      return String(r[0]).trim() && String(r[1]).trim() &&
             String(r[3]).trim().toLowerCase() !== 'нет';
    })
    .map(function (r) {
      return {
        name: String(r[0]).trim(),
        password: String(r[1]).trim(),
        role: String(r[2]).toLowerCase().indexOf('админ') >= 0 ? 'admin' : 'metrolog'
      };
    });
}

function findUser_(ctx, password) {
  password = String(password || '').trim();
  if (!password) return null;
  var users = readUsers_(ctx);
  for (var i = 0; i < users.length; i++) {
    if (users[i].password === password) return users[i];
  }
  return null;
}

function readPlaces_(ctx) {
  var sh = ctx.ss.getSheetByName(SH_PLACES);
  if (!sh || sh.getLastRow() < 2) return [];
  var seen = {};
  return sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues()
    .map(function (r) { return String(r[0]).trim(); })
    .filter(function (v) { if (!v || seen[v]) return false; seen[v] = 1; return true; });
}

function init_(ctx, user) {
  var people = readUsers_(ctx)
    .filter(function (u) { return u.role === 'metrolog'; })
    .map(function (u) { return u.name; });
  return {
    success: true,
    user: { name: user.name, role: user.role },
    people: people,
    places: readPlaces_(ctx),
    settings: ctx.settings,
    today: ctx.today,
    records: readAll_(ctx).records
  };
}

// Кто вносит: метролог выбирает из списка метрологов, по умолчанию - он сам
function resolveBy_(ctx, user, by) {
  by = String(by || '').trim();
  if (!by) return user.name;
  var ok = readUsers_(ctx).some(function (u) { return u.name === by; });
  return ok ? by : user.name;
}

// ============================================================
// ДАТЫ И РАСЧЁТЫ
// ============================================================

function toIso_(v, tz) {
  if (v === null || v === undefined || v === '') return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return isNaN(v.getTime()) ? '' : Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  }
  var s = String(v).trim(), m;
  if ((m = s.match(/^(\d{4})-(\d{2})-(\d{2})/))) return m[1] + '-' + m[2] + '-' + m[3];
  if ((m = s.match(/(\d{1,2})\.(\d{1,2})\.(\d{2,4})/))) {
    var y = Number(m[3]); if (y < 100) y += 2000;
    return y + '-' + pad2_(m[2]) + '-' + pad2_(m[1]);
  }
  return '';
}

function pad2_(n) { n = String(n); return n.length < 2 ? '0' + n : n; }

function validIso_(s) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(s)) return false;
  var p = s.split('-').map(Number);
  var d = new Date(Date.UTC(p[0], p[1] - 1, p[2]));
  return d.getUTCFullYear() === p[0] && d.getUTCMonth() === p[1] - 1 && d.getUTCDate() === p[2] &&
         p[0] >= 2000 && p[0] <= 2100;
}

function dayNum_(iso) {
  var p = iso.split('-').map(Number);
  return Math.round(Date.UTC(p[0], p[1] - 1, p[2]) / 86400000);
}

function fmtStamp_(v, tz, pattern) {
  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v.getTime())) {
    return Utilities.formatDate(v, tz, pattern || 'dd.MM.yyyy HH:mm');
  }
  return v ? String(v) : '';
}

function calcStatus_(r) {
  if ((r.d2 && r.d1 && r.d2 < r.d1) || (r.d3 && r.d2 && r.d3 < r.d2) || (r.d4 && r.d3 && r.d4 < r.d3)) {
    return 'Некорректная дата';
  }
  if (r.d4) return 'Выдано в цех';
  if (r.d3) return 'Вернулось из ЦСМ';
  if (r.d2) return 'В ЦСМ на поверке';
  if (r.d1) return 'Поступил в метрологию';
  return '-';
}

function calcDays_(r) {
  if (!r.d1) return '';
  var max = r.d1;
  DATES.forEach(function (k) { if (r[k] && r[k] > max) max = r[k]; });
  return max === r.d1 ? '' : dayNum_(max) - dayNum_(r.d1);
}

function isLocked_(ctx, r) {
  return !!r.d4 && (dayNum_(ctx.today) - dayNum_(r.d4)) > ctx.settings.lockDays;
}

function validate_(ctx, r) {
  if (!r.name) return 'Укажите наименование СИ';
  for (var i = 0; i < DATES.length; i++) {
    var k = DATES[i];
    if (r[k] && !validIso_(r[k])) return 'Неверная дата в поле «' + LABEL[k] + '»';
  }
  if (!r.d1 && (r.d2 || r.d3 || r.d4)) return 'Сначала заполните дату этапа «' + LABEL.d1 + '»';
  var prev = null;
  var tomorrow = dayNum_(ctx.today) + 1;
  for (var j = 0; j < DATES.length; j++) {
    var key = DATES[j];
    if (!r[key]) continue;
    if (dayNum_(r[key]) > tomorrow) return 'Дата «' + LABEL[key] + '» ещё не наступила';
    if (prev && r[key] < r[prev]) {
      return 'Дата «' + LABEL[key] + '» раньше даты «' + LABEL[prev] + '»';
    }
    prev = key;
  }
  return '';
}

// ============================================================
// ЧТЕНИЕ И ЗАПИСЬ СТРОК
// ============================================================

function dataSheet_(ctx) {
  var sh = ctx.ss.getSheetByName(SH_DATA);
  if (!sh) throw new Error('Нет листа «' + SH_DATA + '»');
  return sh;
}

function readAll_(ctx) {
  var sh = dataSheet_(ctx);
  var last = sh.getLastRow();
  var list = [], rows = {};
  if (last < 2) return { records: list, raw: list, rows: rows };
  var values = sh.getRange(2, 1, last - 1, NCOL).getValues();
  var raw = [];
  for (var i = 0; i < values.length; i++) {
    var v = values[i];
    if (!String(v[0]).trim()) continue;
    var r = {};
    FIELDS.forEach(function (f, c) { r[f] = v[c]; });
    r.id = String(r.id).trim();
    r.num = Number(r.num) || 0;
    ['name', 'serial', 'place', 'result', 'note', 'changes', 'createdBy', 'updatedBy', 'clientId']
      .forEach(function (f) { r[f] = r[f] === null || r[f] === undefined ? '' : String(r[f]); });
    DATES.forEach(function (k) { r[k] = toIso_(r[k], ctx.tz); });
    r._row = i + 2;
    raw.push(r);
    rows[r.id] = r;
  }
  return { records: raw.map(function (r) { return toClient_(ctx, r); }), raw: raw, rows: rows };
}

// Отметка версии записи: меняется при любом изменении содержимого.
// Если запись успели поменять с другого компьютера - сохранение не затрёт чужую правку
function stampOf_(ctx, r) {
  var s = EDITABLE.map(function (f) { return String(r[f] || ''); }).join('\u0001') + '\u0001' +
          String(r.updatedBy || '') + '\u0001' + fmtStamp_(r.updatedAt || r.createdAt, ctx.tz, 'yyyy-MM-dd HH:mm:ss');
  var h1 = 5381, h2 = 52711;
  for (var i = 0; i < s.length; i++) {
    var c = s.charCodeAt(i);
    h1 = ((h1 << 5) + h1 + c) | 0;
    h2 = ((h2 << 5) + h2 ^ c) | 0;
  }
  return (h1 >>> 0).toString(36) + (h2 >>> 0).toString(36);
}

function toClient_(ctx, r) {
  return {
    id: r.id, num: r.num, name: r.name, serial: r.serial, place: r.place,
    d1: r.d1, d2: r.d2, d3: r.d3, d4: r.d4,
    days: calcDays_(r), status: calcStatus_(r),
    result: r.result, note: r.note, changes: r.changes,
    createdBy: r.createdBy, createdAt: fmtStamp_(r.createdAt, ctx.tz),
    updatedBy: r.updatedBy, updatedAt: fmtStamp_(r.updatedAt, ctx.tz),
    stamp: stampOf_(ctx, r),
    locked: isLocked_(ctx, r)
  };
}

function safeText_(s) {
  s = String(s === null || s === undefined ? '' : s);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function rowValues_(ctx, r) {
  return FIELDS.map(function (f) {
    if (DATES.indexOf(f) >= 0) {
      return r[f] ? Utilities.parseDate(r[f], ctx.tz, 'yyyy-MM-dd') : '';
    }
    if (f === 'days') return calcDays_(r);
    if (f === 'status') return calcStatus_(r);
    if (f === 'num') return r.num;
    if (f === 'createdAt' || f === 'updatedAt') return r[f] || '';
    return safeText_(r[f]);
  });
}

// Лист Google создаётся с ограниченным числом строк - добавляем, когда кончаются
function ensureRows_(sh, lastRowNeeded) {
  var max = sh.getMaxRows();
  if (lastRowNeeded > max) sh.insertRowsAfter(max, Math.max(lastRowNeeded - max, 200));
}

function writeRow_(ctx, sh, rowIndex, r) {
  ensureRows_(sh, rowIndex);
  var range = sh.getRange(rowIndex, 1, 1, NCOL);
  sh.getRange(rowIndex, COL.serial).setNumberFormat('@');
  range.setValues([rowValues_(ctx, r)]);
  sh.getRange(rowIndex, COL.d1, 1, 4).setNumberFormat('dd.mm.yyyy');
  sh.getRange(rowIndex, COL.createdAt).setNumberFormat('dd.mm.yyyy hh:mm');
  sh.getRange(rowIndex, COL.updatedAt).setNumberFormat('dd.mm.yyyy hh:mm');
}

function writeLog_(ctx, entries) {
  if (!entries.length) return;
  var sh = ctx.ss.getSheetByName(SH_LOG);
  if (!sh) {
    sh = ctx.ss.insertSheet(SH_LOG);
    sh.getRange(1, 1, 1, 9).setValues([['Когда', 'Кто', 'Роль', 'ID', '№', 'Действие', 'Поле', 'Было', 'Стало']]);
  }
  var start = sh.getLastRow() + 1;
  ensureRows_(sh, start + entries.length - 1);
  sh.getRange(start, 1, entries.length, 9).setValues(entries.map(function (e) {
    return [ctx.now, e.who, e.role, e.id, e.num, e.action, e.field || '',
            safeText_(e.from), safeText_(e.to)];
  }));
  sh.getRange(start, 1, entries.length, 1).setNumberFormat('dd.mm.yyyy hh:mm:ss');
}

function shownValue_(field, v) {
  if (DATES.indexOf(field) >= 0 && v) {
    var p = v.split('-');
    return p[2] + '.' + p[1] + '.' + p[0];
  }
  return v || '';
}

function cleanInput_(d) {
  var out = {};
  out.name = String(d.name || '').replace(/\s+/g, ' ').trim().slice(0, 300);
  out.serial = String(d.serial || '').trim().slice(0, 100);
  out.place = String(d.place || '').replace(/\s+/g, ' ').trim().slice(0, 200);
  out.note = String(d.note || '').replace(/\r/g, '').trim().slice(0, 3000);
  out.result = RESULTS.indexOf(String(d.result || '')) >= 0 ? String(d.result || '') : '';
  DATES.forEach(function (k) { out[k] = String(d[k] || '').trim(); });
  return out;
}

function copyRec_(r) {
  var c = {};
  for (var k in r) c[k] = r[k];
  return c;
}

function roleName_(user) { return user.role === 'admin' ? 'администратор' : 'метролог'; }

// ============================================================
// ДЕЙСТВИЯ
// ============================================================

function saveRecord_(ctx, user, d) {
  var all = readAll_(ctx);
  var sh = dataSheet_(ctx);
  var by = resolveBy_(ctx, user, d.by);
  var input = cleanInput_(d);
  var isNew = !d.id;
  var rec, old = null;

  if (isNew) {
    if (d.clientId) {
      for (var i = 0; i < all.raw.length; i++) {
        if (all.raw[i].clientId === String(d.clientId)) {
          return { success: true, duplicate: true, record: toClient_(ctx, all.raw[i]) };
        }
      }
    }
    rec = {};
  } else {
    old = all.rows[String(d.id)];
    if (!old) return { success: false, gone: true, error: 'Запись не найдена - возможно, её удалили. Обновите журнал.' };
    if (d.stamp && d.stamp !== stampOf_(ctx, old)) {
      // Повтор из очереди после обрыва связи: правка уже записана - это не конфликт
      var same = EDITABLE.every(function (f) { return String(old[f] || '') === String(input[f] || ''); });
      if (same) return { success: true, unchanged: true, record: toClient_(ctx, old) };
      return {
        success: false, conflict: true, record: toClient_(ctx, old),
        error: 'Эту запись уже изменил(а) ' + (old.updatedBy || old.createdBy) + ' (' +
               fmtStamp_(old.updatedAt || old.createdAt, ctx.tz) + '). Данные обновлены - проверьте и сохраните ещё раз.'
      };
    }
    if (isLocked_(ctx, old) && user.role !== 'admin') {
      return {
        success: false, locked: true, record: toClient_(ctx, old),
        error: 'Запись закрыта: прибор выдан в цех больше ' + ctx.settings.lockDays +
               ' дней назад. Исправить может только администратор.'
      };
    }
    rec = copyRec_(old);
  }

  EDITABLE.forEach(function (f) { rec[f] = input[f]; });
  var err = validate_(ctx, rec);
  if (err) return { success: false, error: err };

  var entries = [];
  if (isNew) {
    var maxNum = 0, maxId = 0;
    all.raw.forEach(function (r) {
      if (r.num > maxNum) maxNum = r.num;
      var m = /^R(\d+)$/.exec(r.id);
      if (m && Number(m[1]) > maxId) maxId = Number(m[1]);
    });
    rec.id = 'R' + ('000' + (maxId + 1)).slice(-4);
    rec.num = maxNum + 1;
    rec.changes = '';
    rec.createdBy = by;
    rec.createdAt = ctx.now;
    rec.updatedBy = '';
    rec.updatedAt = '';
    rec.clientId = String(d.clientId || '');
    writeRow_(ctx, sh, Math.max(sh.getLastRow() + 1, 2), rec);
    entries.push({ who: by, role: roleName_(user), id: rec.id, num: rec.num, action: 'Добавлен',
                   to: rec.name + (rec.serial ? ', зав.№ ' + rec.serial : '') });
  } else {
    EDITABLE.forEach(function (f) {
      if (String(old[f] || '') !== String(rec[f] || '')) {
        entries.push({ who: by, role: roleName_(user), id: rec.id, num: rec.num, action: 'Изменён',
                       field: LABEL[f], from: shownValue_(f, old[f]), to: shownValue_(f, rec[f]) });
      }
    });
    if (!entries.length) return { success: true, unchanged: true, record: toClient_(ctx, old) };
    if (isLocked_(ctx, old)) {
      var mark = 'Корректировка администратором ' + fmtStamp_(ctx.now, ctx.tz);
      rec.changes = rec.changes ? rec.changes + '\n' + mark : mark;
    }
    rec.updatedBy = by;
    rec.updatedAt = ctx.now;
    writeRow_(ctx, sh, old._row, rec);
  }
  writeLog_(ctx, entries);
  return { success: true, record: toClient_(ctx, rec) };
}

// Массовая отметка этапа: несколько приборов, один этап, одна дата
function bulkStage_(ctx, user, d) {
  var stage = Number(d.stage);
  if (!(stage >= 1 && stage <= 4)) return { success: false, error: 'Не выбран этап' };
  var key = 'd' + stage;
  var date = String(d.date || '').trim();
  if (date && !validIso_(date)) return { success: false, error: 'Неверная дата' };
  var setResult = d.result === undefined || d.result === null || d.result === 'keep'
    ? null : (RESULTS.indexOf(String(d.result)) >= 0 ? String(d.result) : null);
  var ids = Array.isArray(d.ids) ? d.ids.map(String) : [];
  if (!ids.length) return { success: false, error: 'Не выбраны приборы' };
  if (ids.length > 500) return { success: false, error: 'Не больше 500 приборов за раз' };

  var all = readAll_(ctx);
  var sh = dataSheet_(ctx);
  var by = resolveBy_(ctx, user, d.by);
  var updated = [], errors = [], entries = [];

  ids.forEach(function (id) {
    var old = all.rows[id];
    if (!old) { errors.push({ id: id, num: '', error: 'запись не найдена' }); return; }
    if (isLocked_(ctx, old) && user.role !== 'admin') {
      errors.push({ id: id, num: old.num, error: 'запись закрыта' }); return;
    }
    var rec = copyRec_(old);
    rec[key] = date;
    if (setResult !== null) rec.result = setResult;
    var err = validate_(ctx, rec);
    if (err) { errors.push({ id: id, num: old.num, error: err }); return; }

    var changed = [];
    if (old[key] !== rec[key]) changed.push(key);
    if (old.result !== rec.result) changed.push('result');
    if (!changed.length) { updated.push(toClient_(ctx, old)); return; }

    changed.forEach(function (f) {
      entries.push({ who: by, role: roleName_(user), id: rec.id, num: rec.num, action: 'Изменён (группой)',
                     field: LABEL[f], from: shownValue_(f, old[f]), to: shownValue_(f, rec[f]) });
    });
    if (isLocked_(ctx, old)) {
      var mark = 'Корректировка администратором ' + fmtStamp_(ctx.now, ctx.tz);
      rec.changes = rec.changes ? rec.changes + '\n' + mark : mark;
    }
    rec.updatedBy = by;
    rec.updatedAt = ctx.now;
    writeRow_(ctx, sh, old._row, rec);
    updated.push(toClient_(ctx, rec));
  });

  writeLog_(ctx, entries);
  return { success: true, updated: updated, errors: errors };
}

function deleteRecord_(ctx, user, id) {
  if (user.role !== 'admin') return { success: false, error: 'Удалять записи может только администратор' };
  var all = readAll_(ctx);
  var old = all.rows[String(id || '')];
  if (!old) return { success: false, gone: true, error: 'Запись уже удалена' };
  dataSheet_(ctx).deleteRow(old._row);
  writeLog_(ctx, [{
    who: user.name, role: roleName_(user), id: old.id, num: old.num, action: 'Удалён',
    from: old.name + (old.serial ? ', зав.№ ' + old.serial : '') + (old.place ? ', ' + old.place : '')
  }]);
  return { success: true, id: old.id };
}

function getLog_(ctx, user, id) {
  if (user.role !== 'admin') return { success: false, error: 'Журнал правок доступен только администратору' };
  var sh = ctx.ss.getSheetByName(SH_LOG);
  if (!sh || sh.getLastRow() < 2) return { success: true, entries: [] };
  var values = sh.getRange(2, 1, sh.getLastRow() - 1, 9).getValues();
  var out = [];
  for (var i = values.length - 1; i >= 0 && out.length < 400; i--) {
    var v = values[i];
    if (id && String(v[3]) !== String(id)) continue;
    out.push({
      when: fmtStamp_(v[0], ctx.tz), who: String(v[1]), role: String(v[2]), id: String(v[3]),
      num: v[4], action: String(v[5]), field: String(v[6]), from: String(v[7]), to: String(v[8])
    });
  }
  return { success: true, entries: out };
}

// ============================================================
// УСТАНОВКА: выберите функцию setup и нажмите «Выполнить».
//   - скрипт привязан к пустой таблице (Расширения -> Apps Script):
//     листы создаются в ней;
//   - отдельный проект на script.google.com: создаётся новая
//     таблица «Журнал движения СИ», ссылка - в журнале выполнения.
// Повторный запуск безопасен: существующие данные не трогаются,
// недостающие листы и оформление добавляются.
// Если в проекте есть файл Import.gs с данными из Excel и лист
// «Приборы» пуст - записи загружаются.
// ============================================================

var HEADERS = {};
HEADERS[SH_DATA] = ['ID', '№', 'Наименование СИ', 'Зав. №', 'Цех / место',
  '1. Цех→Метрология', '2. Метрология→ЦСМ', '3. ЦСМ→Метрология', '4. Метрология→Цех',
  'Дней всего', 'Статус', 'Примечание', 'Дополнение', 'Изменения',
  'Кто внёс', 'Создано', 'Кто изменил', 'Изменено', 'clientId'];
HEADERS[SH_USERS] = ['ФИО', 'Пароль', 'Роль', 'Активен'];
HEADERS[SH_PLACES] = ['Цех / место'];
HEADERS[SH_LOG] = ['Когда', 'Кто', 'Роль', 'ID', '№', 'Действие', 'Поле', 'Было', 'Стало'];
HEADERS[SH_SET] = ['Параметр', 'Значение', 'Пояснение'];

var WIDTHS = {};
WIDTHS[SH_DATA] = [60, 45, 280, 120, 210, 105, 105, 105, 105, 65, 160, 140, 320, 220, 130, 120, 130, 120, 100];
WIDTHS[SH_USERS] = [230, 120, 130, 80];
WIDTHS[SH_PLACES] = [300];
WIDTHS[SH_LOG] = [130, 150, 110, 60, 45, 130, 160, 220, 220];
WIDTHS[SH_SET] = [180, 260, 520];

function runSetup_() {
  var ss = null, created = false;
  try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) { ss = null; }
  if (!ss) {
    var id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
    if (id) { try { ss = SpreadsheetApp.openById(id); } catch (e) { ss = null; } }
  }
  if (!ss) {
    ss = SpreadsheetApp.create('Журнал движения СИ - метрология');
    PropertiesService.getScriptProperties().setProperty('SHEET_ID', ss.getId());
    created = true;
  }
  ss.setSpreadsheetTimeZone('Asia/Novosibirsk');

  var report = [];
  [SH_DATA, SH_USERS, SH_PLACES, SH_LOG, SH_SET].forEach(function (name) {
    var sh = ss.getSheetByName(name);
    if (!sh) { sh = ss.insertSheet(name); report.push('создан лист «' + name + '»'); }
    prepareSheet_(sh, name);
  });

  // пустой лист по умолчанию больше не нужен
  ss.getSheets().forEach(function (sh) {
    var n = sh.getName();
    if (!HEADERS[n] && sh.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(sh);
  });

  // настройки по умолчанию
  var set = ss.getSheetByName(SH_SET);
  if (set.getLastRow() < 2) {
    set.getRange(2, 1, 3, 3).setValues([
      ['Дней до блокировки', 10, 'Через сколько дней после выдачи в цех метролог больше не может менять запись'],
      ['Порог дней', 20, 'Срок «Дней всего», с которого статус подсвечивается красным'],
      ['Название', 'АО "Новосибирская птицефабрика"', 'Печатается в шапке журнала']
    ]);
  }

  // пользователи: случайные пароли, если лист пуст
  var users = ss.getSheetByName(SH_USERS);
  var passwords = [];
  if (users.getLastRow() < 2) {
    var used = {};
    var rows = [['ФИО-1', 'метролог'], ['ФИО-2', 'метролог'], ['ФИО-3', 'метролог'], ['Администратор', 'администратор']]
      .map(function (u) {
        var p;
        do { p = String(100000 + Math.floor(Math.random() * 900000)); } while (used[p]);
        used[p] = true;
        passwords.push(u[0] + ': ' + p);
        return [u[0], p, u[1], 'да'];
      });
    users.getRange(2, 2, rows.length, 1).setNumberFormat('@');
    users.getRange(2, 1, rows.length, 4).setValues(rows);
    report.push('созданы пользователи с паролями');
  }

  // перенос записей из Excel (файл Import.gs)
  var data = ss.getSheetByName(SH_DATA);
  var imported = 0;
  if (data.getLastRow() < 2 && typeof IMPORT_DATA !== 'undefined' && IMPORT_DATA.length) {
    imported = importRows_(ss, data, IMPORT_DATA);
    report.push('перенесено записей из Excel: ' + imported);
  }

  // места: из перенесённых записей, если список пуст
  var pl = ss.getSheetByName(SH_PLACES);
  if (pl.getLastRow() < 2 && data.getLastRow() > 1) {
    var seen = {}, list = [];
    data.getRange(2, COL.place, data.getLastRow() - 1, 1).getValues().forEach(function (r) {
      var v = String(r[0]).trim();
      if (v && !seen[v]) { seen[v] = 1; list.push([v]); }
    });
    list.sort(function (a, b) { return a[0].localeCompare(b[0], 'ru'); });
    if (list.length) pl.getRange(2, 1, list.length, 1).setValues(list);
  }

  ss.setActiveSheet(data);
  Logger.log('ТАБЛИЦА: ' + ss.getUrl() + (created ? ' (создана заново)' : ''));
  Logger.log(report.length ? 'Сделано: ' + report.join('; ') : 'Всё уже было на месте, данные не тронуты');
  if (passwords.length) {
    Logger.log('ПАРОЛИ (запишите и раздайте; ФИО замените на листе «Пользователи»):');
    passwords.forEach(function (p) { Logger.log('   ' + p); });
  }
  Logger.log('Дальше: Начать развертывание -> Новое развертывание -> Веб-приложение');
  check();
}

function prepareSheet_(sh, name) {
  var head = HEADERS[name];
  var current = sh.getRange(1, 1, 1, head.length).getValues()[0];
  if (current.join('') === '') sh.getRange(1, 1, 1, head.length).setValues([head]);
  sh.getRange(1, 1, 1, head.length)
    .setFontWeight('bold').setFontColor('#ffffff').setBackground('#16213e')
    .setWrap(true).setVerticalAlignment('middle').setHorizontalAlignment('center');
  sh.setFrozenRows(1);
  sh.setRowHeight(1, 36);
  WIDTHS[name].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });

  var rows = Math.max(sh.getMaxRows() - 1, 1);
  if (name === SH_DATA) {
    sh.setFrozenColumns(3);
    sh.getRange(2, COL.serial, rows, 1).setNumberFormat('@');
    sh.getRange(2, COL.d1, rows, 4).setNumberFormat('dd.mm.yyyy');
    sh.getRange(2, COL.createdAt, rows, 1).setNumberFormat('dd.mm.yyyy hh:mm');
    sh.getRange(2, COL.updatedAt, rows, 1).setNumberFormat('dd.mm.yyyy hh:mm');
    sh.getRange(2, 1, rows, NCOL).setVerticalAlignment('top').setWrap(true);
    sh.getRange(2, COL.result, rows, 1).setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList(['Поверен', 'Не прошёл проверку'], true).setAllowInvalid(true).build());
    var st = sh.getRange(2, COL.status, rows, 1), rs = sh.getRange(2, COL.result, rows, 1);
    sh.setConditionalFormatRules([
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Некорректная дата')
        .setBackground('#ffc000').setRanges([st]).build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=AND(ISNUMBER($J2),$J2>=INDIRECT("Настройки!B3"))')
        .setBackground('#ff7c80').setRanges([st]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=ISNUMBER($J2)')
        .setBackground('#c6efce').setRanges([st]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Не прошёл проверку')
        .setBackground('#ff7c80').setRanges([rs]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Поверен')
        .setBackground('#c6efce').setRanges([rs]).build()
    ]);
  } else if (name === SH_USERS) {
    sh.getRange(2, 2, rows, 1).setNumberFormat('@');
    sh.getRange(2, 3, rows, 1).setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList(['метролог', 'администратор'], true).setAllowInvalid(false).build());
    sh.getRange(2, 4, rows, 1).setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList(['да', 'нет'], true).setAllowInvalid(false).build());
    sh.getRange('F1').setValue('Роль: «метролог» или «администратор». Активен: «да» / «нет». Пароли у всех разные, не короче 6 знаков.')
      .setFontColor('#666666').setFontStyle('italic');
  } else if (name === SH_LOG) {
    sh.getRange(2, 1, rows, 1).setNumberFormat('dd.mm.yyyy hh:mm:ss');
  } else if (name === SH_PLACES) {
    sh.getRange('C1').setValue('Список для выпадающего меню «Цех / место». Можно дописывать и исправлять.')
      .setFontColor('#666666').setFontStyle('italic');
  }
}

// Строка импорта: [№, наименование, зав.№, место, д1, д2, д3, д4, примечание, дополнение, изменения]
function importRows_(ss, sh, list) {
  var tz = ss.getSpreadsheetTimeZone();
  var now = new Date();
  var values = list.map(function (x, i) {
    var r = { id: 'R' + ('000' + (i + 1)).slice(-4), num: Number(x[0]) || i + 1,
              name: x[1], serial: x[2], place: x[3], d1: x[4], d2: x[5], d3: x[6], d4: x[7],
              result: x[8], note: x[9], changes: x[10],
              createdBy: 'Перенос из Excel', createdAt: now, updatedBy: '', updatedAt: '', clientId: '' };
    return rowValues_({ tz: tz }, r);
  });
  ensureRows_(sh, values.length + 1);
  sh.getRange(2, COL.serial, values.length, 1).setNumberFormat('@');
  sh.getRange(2, 1, values.length, NCOL).setValues(values);
  return values.length;
}

// ============================================================
// ПРОВЕРКА: выберите функцию check и нажмите «Выполнить»
// ============================================================

function runCheck_() {
  var ctx = context_();
  var problems = [];
  [SH_DATA, SH_USERS, SH_PLACES, SH_LOG, SH_SET].forEach(function (n) {
    if (!ctx.ss.getSheetByName(n)) problems.push('нет листа «' + n + '»');
  });
  var users = readUsers_(ctx);
  var seen = {};
  users.forEach(function (u) {
    if (seen[u.password]) problems.push('одинаковый пароль у «' + seen[u.password] + '» и «' + u.name + '»');
    seen[u.password] = u.name;
  });
  if (!users.some(function (u) { return u.role === 'admin'; })) problems.push('нет администратора');
  var n = ctx.ss.getSheetByName(SH_DATA) ? readAll_(ctx).records.length : 0;
  Logger.log('Часовой пояс таблицы: ' + ctx.tz + ', сегодня: ' + ctx.today);
  Logger.log('Записей: ' + n + ', пользователей: ' + users.length + ', мест: ' + readPlaces_(ctx).length);
  Logger.log(problems.length ? 'ЕСТЬ ЗАМЕЧАНИЯ: ' + problems.join('; ') : 'Всё в порядке');
}
