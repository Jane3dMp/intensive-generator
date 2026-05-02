/**
 * Code.gs — серверная логика для генератора интенсивов «Прознание + Coddy»
 *
 * Этот скрипт может работать одновременно с лагерями и интенсивами.
 * Все запросы помечены параметром `project`:
 *   - 'camp'      → лагерь (старый код-генератор)
 *   - 'intensive' → интенсив (новый код-генератор)
 *
 * Если у вас уже работает лагерный Code.gs — этот файл можно
 * добавить как отдельный скрипт в той же таблице (но с другим
 * SHEET_ID или с другими именами листов), либо смёржить.
 *
 * НАСТРОЙКА:
 * 1. В Google Sheets создайте листы (если их нет):
 *    Для интенсивов: «Активности_интенсивы», «Реквизиты_интенсивы»,
 *                   «Опубликованные_интенсивы», «Черновики_интенсивы», «Сессии_интенсивы»
 * 2. Скопируйте SHEET_ID из URL Google Sheets и подставьте ниже
 * 3. (опционально) Project Settings → Script Properties → ANTHROPIC_API_KEY
 * 4. Deploy → New deployment → Web app → Execute as: Me, Who: Anyone
 */

// ============================================================
// КОНФИГ
// ============================================================
const SHEET_ID = 'PUT_YOUR_GOOGLE_SHEETS_ID_HERE';

// Имена листов для интенсивов
const SHEET_INT_ACTIVITIES = 'Активности_интенсивы';
const SHEET_INT_DETAILS    = 'Реквизиты_интенсивы';
const SHEET_INT_PUBLISHED  = 'Опубликованные_интенсивы';
const SHEET_INT_DRAFTS     = 'Черновики_интенсивы';
const SHEET_INT_SESSIONS   = 'Сессии_интенсивы';

// Имена листов для лагерей (если используется тот же SHEET_ID)
const SHEET_CAMP_ACTIVITIES = 'Активности';
const SHEET_CAMP_DETAILS    = 'Реквизиты';
const SHEET_CAMP_PUBLISHED  = 'Опубликованные версии';
const SHEET_CAMP_DRAFTS     = 'Черновики';
const SHEET_CAMP_SESSIONS   = 'Сессии';

// Anthropic
const ANTHROPIC_API_URL = 'https://api.anthropic.com/v1/messages';
const ANTHROPIC_MODEL = 'claude-sonnet-4-20250514';
const ANTHROPIC_MAX_TOKENS = 2000;

// Сессии: считаем активной если heartbeat был не более 5 минут назад
const SESSION_TIMEOUT_MS = 5 * 60 * 1000;

// ============================================================
// УТИЛИТЫ
// ============================================================
function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet_(name) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh = ss.getSheetByName(name);
  if (!sh) {
    // Авто-создание с дефолтными заголовками для известных листов
    sh = ss.insertSheet(name);
    initSheetHeaders_(sh, name);
  }
  return sh;
}

function initSheetHeaders_(sh, name) {
  const headers = {
    [SHEET_INT_ACTIVITIES]: ['Статус','Шаблон занятия','Примеры тем','Педагог','Направление','Локация / кабинет','Возраст','Длительность (мин)','Формат','Описание'],
    [SHEET_INT_DETAILS]: ['Ключ','Значение'],
    [SHEET_INT_PUBLISHED]: ['id','name','title','updated','stateJson'],
    [SHEET_INT_DRAFTS]: ['id','name','updated','stateJson'],
    [SHEET_INT_SESSIONS]: ['sessionId','publishId','editorName','userAgent','updated']
  };
  if (headers[name]) {
    sh.getRange(1, 1, 1, headers[name].length).setValues([headers[name]]);
    sh.setFrozenRows(1);
  }
}

function getProjectSheets_(project) {
  if (project === 'intensive') {
    return {
      activities: SHEET_INT_ACTIVITIES,
      details: SHEET_INT_DETAILS,
      published: SHEET_INT_PUBLISHED,
      drafts: SHEET_INT_DRAFTS,
      sessions: SHEET_INT_SESSIONS
    };
  }
  return {
    activities: SHEET_CAMP_ACTIVITIES,
    details: SHEET_CAMP_DETAILS,
    published: SHEET_CAMP_PUBLISHED,
    drafts: SHEET_CAMP_DRAFTS,
    sessions: SHEET_CAMP_SESSIONS
  };
}

function readSheetAsObjects_(sh) {
  const range = sh.getDataRange();
  const values = range.getValues();
  if (values.length < 2) return [];
  const headers = values[0].map(h => String(h).trim());
  const out = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (row.every(c => c === '' || c == null)) continue;
    const obj = {};
    headers.forEach((h, j) => { obj[h] = row[j]; });
    obj._rowIdx = i + 1; // 1-based
    out.push(obj);
  }
  return out;
}

function findRowById_(sh, id) {
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) return i + 1; // 1-based
  }
  return -1;
}

function nowIso_() {
  return new Date().toISOString();
}

// ============================================================
// ENTRY POINTS
// ============================================================
function doGet(e) {
  try {
    const params = e.parameter || {};
    const action = params.action || 'all';
    const project = params.project || 'camp';

    switch(action) {
      case 'all':           return jsonResponse(getAll_(project));
      case 'library':       return jsonResponse(getLibrary_(project));
      case 'details':       return jsonResponse(getDetails_(project));
      case 'listPublished': return jsonResponse(listPublished_(project));
      case 'getPublished':  return jsonResponse(getPublished_(project, params.id));
      case 'listDrafts':    return jsonResponse(listDrafts_(project));
      case 'loadDraft':     return jsonResponse(loadDraft_(project, params.id));
      case 'listSessions':  return jsonResponse(listSessions_(project, params.publishId));
      default: return jsonResponse({ ok: false, error: 'Unknown action: ' + action });
    }
  } catch(err) {
    return jsonResponse({ ok: false, error: String(err.message || err) });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;
    const project = body.project || 'camp';

    switch(action) {
      case 'claude':       return jsonResponse(callClaude_(body.prompt, body.system));
      case 'publish':      return jsonResponse(publishVersion_(project, body));
      case 'unpublish':    return jsonResponse(unpublishVersion_(project, body.id));
      case 'saveDraft':    return jsonResponse(saveDraft_(project, body));
      case 'deleteDraft':  return jsonResponse(deleteDraft_(project, body.id));
      case 'heartbeat':    return jsonResponse(heartbeat_(project, body));
      case 'endSession':   return jsonResponse(endSession_(project, body.sessionId));
      default: return jsonResponse({ ok: false, error: 'Unknown action: ' + action });
    }
  } catch(err) {
    return jsonResponse({ ok: false, error: String(err.message || err) });
  }
}

// ============================================================
// ACTIVITIES / DETAILS
// ============================================================
function getAll_(project) {
  return {
    ok: true,
    activities: getActivities_(project),
    details: getDetailsRaw_(project)
  };
}

function getLibrary_(project) {
  return { ok: true, activities: getActivities_(project) };
}

function getActivities_(project) {
  const sheets = getProjectSheets_(project);
  const sh = getSheet_(sheets.activities);
  const rows = readSheetAsObjects_(sh);
  const out = [];
  rows.forEach(r => {
    const status = String(r['Статус'] || '').toLowerCase().trim();
    if (status !== 'утверждено') return; // Только утверждённые
    out.push({
      template:    String(r['Шаблон занятия'] || '').trim(),
      examples:    String(r['Примеры тем'] || '').trim(),
      teacher:     String(r['Педагог'] || '').trim(),
      direction:   String(r['Направление'] || '').trim(),
      location:    String(r['Локация / кабинет'] || '').trim(),
      age:         String(r['Возраст'] || '').trim(),
      duration:    Number(r['Длительность (мин)']) || 0,
      format:      String(r['Формат'] || '').trim(),
      description: String(r['Описание'] || '').trim()
    });
  });
  return out;
}

function getDetails_(project) {
  return { ok: true, details: getDetailsRaw_(project) };
}

function getDetailsRaw_(project) {
  const sheets = getProjectSheets_(project);
  let sh;
  try { sh = getSheet_(sheets.details); } catch(e) { return {}; }
  const data = sh.getDataRange().getValues();
  const out = {};
  for (let i = 1; i < data.length; i++) {
    const k = String(data[i][0] || '').trim();
    const v = data[i][1];
    if (k) out[k] = v;
  }
  return out;
}

// ============================================================
// PUBLISH
// ============================================================
function publishVersion_(project, body) {
  const id = String(body.id || '').trim();
  const name = String(body.name || '').trim();
  const stateJson = String(body.stateJson || '');
  if (!id) return { ok: false, error: 'нет id' };
  if (!stateJson) return { ok: false, error: 'нет stateJson' };

  const sheets = getProjectSheets_(project);
  const sh = getSheet_(sheets.published);
  const existingRow = findRowById_(sh, id);

  // Парсим title из state для удобной фильтрации
  let title = name;
  try {
    const obj = JSON.parse(stateJson);
    title = String(obj.title || obj.theme || name);
  } catch(e) {}

  const rowData = [id, name, title, nowIso_(), stateJson];

  if (existingRow > 0) {
    // ВАЖНО: setValues обрезает строки длиннее 50000 символов.
    // Поэтому удаляем строку и добавляем заново.
    sh.deleteRow(existingRow);
    sh.appendRow(rowData);
  } else {
    sh.appendRow(rowData);
  }

  return { ok: true, id, updated: nowIso_() };
}

function listPublished_(project) {
  const sheets = getProjectSheets_(project);
  const sh = getSheet_(sheets.published);
  const data = sh.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    out.push({
      id: String(row[0]),
      name: String(row[1] || ''),
      title: String(row[2] || ''),
      updated: row[3] ? String(row[3]) : ''
      // stateJson намеренно не возвращаем — слишком жирно для списка
    });
  }
  // Сортируем по убыванию updated
  out.sort((a, b) => (b.updated || '').localeCompare(a.updated || ''));
  return { ok: true, versions: out };
}

function getPublished_(project, id) {
  if (!id) return { ok: false, error: 'нет id' };
  const sheets = getProjectSheets_(project);
  const sh = getSheet_(sheets.published);
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      return {
        ok: true,
        version: {
          id: String(data[i][0]),
          name: String(data[i][1] || ''),
          title: String(data[i][2] || ''),
          updated: data[i][3] ? String(data[i][3]) : '',
          stateJson: String(data[i][4] || '')
        }
      };
    }
  }
  return { ok: false, error: 'не найдено' };
}

function unpublishVersion_(project, id) {
  if (!id) return { ok: false, error: 'нет id' };
  const sheets = getProjectSheets_(project);
  const sh = getSheet_(sheets.published);
  const row = findRowById_(sh, id);
  if (row > 0) sh.deleteRow(row);
  return { ok: true };
}

// ============================================================
// DRAFTS
// ============================================================
function saveDraft_(project, body) {
  const sheets = getProjectSheets_(project);
  const sh = getSheet_(sheets.drafts);
  let id = String(body.id || '').trim();
  const name = String(body.name || 'Без названия').trim();
  const stateObj = body.state || {};
  const stateJson = JSON.stringify(stateObj);

  if (!id) {
    id = 'd_' + Math.random().toString(36).slice(2, 10) + Date.now().toString(36).slice(-4);
  }

  const rowData = [id, name, nowIso_(), stateJson];
  const existing = findRowById_(sh, id);
  if (existing > 0) {
    sh.deleteRow(existing);
  }
  sh.appendRow(rowData);

  return { ok: true, draft: { id, name, updated: nowIso_() } };
}

function listDrafts_(project) {
  const sheets = getProjectSheets_(project);
  const sh = getSheet_(sheets.drafts);
  const data = sh.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    out.push({
      id: String(data[i][0]),
      name: String(data[i][1] || ''),
      updated: data[i][2] ? String(data[i][2]) : ''
    });
  }
  out.sort((a, b) => (b.updated || '').localeCompare(a.updated || ''));
  return { ok: true, drafts: out };
}

function loadDraft_(project, id) {
  if (!id) return { ok: false, error: 'нет id' };
  const sheets = getProjectSheets_(project);
  const sh = getSheet_(sheets.drafts);
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      let stateObj = {};
      try { stateObj = JSON.parse(String(data[i][3] || '{}')); } catch(e) {}
      return {
        ok: true,
        draft: {
          id: String(data[i][0]),
          name: String(data[i][1] || ''),
          updated: data[i][2] ? String(data[i][2]) : '',
          state: stateObj
        }
      };
    }
  }
  return { ok: false, error: 'не найдено' };
}

function deleteDraft_(project, id) {
  if (!id) return { ok: false, error: 'нет id' };
  const sheets = getProjectSheets_(project);
  const sh = getSheet_(sheets.drafts);
  const row = findRowById_(sh, id);
  if (row > 0) sh.deleteRow(row);
  return { ok: true };
}

// ============================================================
// SESSIONS / HEARTBEAT
// ============================================================
function heartbeat_(project, body) {
  const sessionId = String(body.sessionId || '').trim();
  const publishId = String(body.publishId || '').trim();
  if (!sessionId || !publishId) return { ok: false, error: 'нет sessionId или publishId' };

  const sheets = getProjectSheets_(project);
  const sh = getSheet_(sheets.sessions);
  const data = sh.getDataRange().getValues();

  // Чистим протухшие
  const now = Date.now();
  const liveRows = [];
  for (let i = data.length - 1; i >= 1; i--) {
    const updated = data[i][4] ? new Date(data[i][4]).getTime() : 0;
    if (now - updated > SESSION_TIMEOUT_MS) {
      sh.deleteRow(i + 1);
    }
  }

  // Обновляем или вставляем
  const freshData = sh.getDataRange().getValues();
  let foundRow = -1;
  for (let i = 1; i < freshData.length; i++) {
    if (String(freshData[i][0]) === sessionId) { foundRow = i + 1; break; }
  }

  const rowData = [sessionId, publishId, body.editorName || '', body.userAgent || '', nowIso_()];
  if (foundRow > 0) {
    sh.getRange(foundRow, 1, 1, rowData.length).setValues([rowData]);
  } else {
    sh.appendRow(rowData);
  }
  return { ok: true };
}

function endSession_(project, sessionId) {
  if (!sessionId) return { ok: false, error: 'нет sessionId' };
  const sheets = getProjectSheets_(project);
  const sh = getSheet_(sheets.sessions);
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === sessionId) {
      sh.deleteRow(i + 1);
      break;
    }
  }
  return { ok: true };
}

function listSessions_(project, publishId) {
  if (!publishId) return { ok: true, sessions: [] };
  const sheets = getProjectSheets_(project);
  const sh = getSheet_(sheets.sessions);
  const data = sh.getDataRange().getValues();
  const now = Date.now();
  const out = [];
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]) !== String(publishId)) continue;
    const updated = data[i][4] ? new Date(data[i][4]).getTime() : 0;
    const ageMs = now - updated;
    if (ageMs > SESSION_TIMEOUT_MS) continue;
    out.push({
      sessionId: String(data[i][0]),
      publishId: String(data[i][1]),
      editorName: String(data[i][2] || ''),
      userAgent: String(data[i][3] || ''),
      updated: data[i][4] ? String(data[i][4]) : '',
      ageSec: Math.round(ageMs / 1000)
    });
  }
  return { ok: true, sessions: out };
}

// ============================================================
// CLAUDE PROXY
// ============================================================
function callClaude_(prompt, system) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!apiKey) {
    return { ok: false, error: 'ANTHROPIC_API_KEY не настроен в Script Properties' };
  }
  if (!prompt) return { ok: false, error: 'нет prompt' };

  const messages = [{ role: 'user', content: prompt }];
  const payload = {
    model: ANTHROPIC_MODEL,
    max_tokens: ANTHROPIC_MAX_TOKENS,
    messages: messages
  };
  if (system) payload.system = system;

  try {
    const response = UrlFetchApp.fetch(ANTHROPIC_API_URL, {
      method: 'post',
      contentType: 'application/json',
      headers: {
        'x-api-key': apiKey,
        'anthropic-version': '2023-06-01'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    const code = response.getResponseCode();
    const text = response.getContentText();
    if (code < 200 || code >= 300) {
      return { ok: false, error: 'Anthropic ' + code + ': ' + text.slice(0, 500) };
    }
    const data = JSON.parse(text);
    let outText = '';
    if (Array.isArray(data.content)) {
      data.content.forEach(c => { if (c.type === 'text') outText += c.text; });
    }
    return { ok: true, text: outText };
  } catch(e) {
    return { ok: false, error: String(e.message || e) };
  }
}

// ============================================================
// УТИЛИТЫ ДЛЯ ОТЛАДКИ (запускайте вручную из редактора Apps Script)
// ============================================================
function _testLoadAll_intensive() {
  const result = getAll_('intensive');
  Logger.log('Активностей: ' + result.activities.length);
  Logger.log('Реквизитов: ' + Object.keys(result.details).length);
  result.activities.slice(0, 3).forEach(a => Logger.log(JSON.stringify(a)));
}

function _testInitSheets_intensive() {
  // Создаёт все нужные листы для интенсивов с правильными заголовками
  ['Активности_интенсивы','Реквизиты_интенсивы','Опубликованные_интенсивы','Черновики_интенсивы','Сессии_интенсивы'].forEach(name => {
    const sh = getSheet_(name);
    Logger.log('OK: ' + name + ' (' + sh.getLastRow() + ' rows)');
  });
}
