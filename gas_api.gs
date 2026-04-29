/**
 * GitHub Pages 用 読み取りAPI。
 * 既存の onEdit / 進捗保存ロジックは変更せず、doGet のAPIだけを追加・更新する想定です。
 * GitHub Pagesから読みやすいように JSONP(callback=...) に対応しています。
 */

function doGet(e) {
  try {
    const p = e && e.parameter ? e.parameter : {};
    const action = p.action || '';

    const actions = {
      getHomeTopInfo: () => api_getHomeTopInfo_(),
      getHomeSummary: () => api_getHomeSummary_(Number(p.days || 7)),
      getTasksDueWithinDays: () => api_getTasksDueWithinDays_(Number(p.days || 7)),
      getUnresolvedQuestions: () => api_getUnresolvedQuestions_(),
      getLineShareText: () => api_getLineShareText_(),
      getTasks: () => api_getTasks_(),
      getTaskDetail: () => api_getTaskDetail_(p.no),
      getTasksByAssignee: () => api_getTasksByAssignee_(p.name),
      getWbsTree: () => api_getWbsTree_(),
      getAssignees: () => api_getAssignees_(),
      getMilestones: () => api_getMilestones_(),
      getSchedule: () => api_getSchedule_(),
      getQuestions: () => api_getQuestions_(p.status),
      getMemos: () => api_getMemos_(),
      getMemoDetail: () => api_getMemoDetail_(p.id),
      getGuests: () => api_getGuests_(),
    };

    if (!actions[action]) throw new Error('Unknown action: ' + action);
    return api_output_(p, { ok: true, data: actions[action]() });
  } catch (err) {
    const p = e && e.parameter ? e.parameter : {};
    return api_output_(p, { ok: false, error: err.message });
  }
}

function api_output_(params, obj) {
  const json = JSON.stringify(obj);
  const callback = params && params.callback ? String(params.callback) : '';
  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

function api_ss_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function api_sheet_(sheetName) {
  return api_ss_().getSheetByName(sheetName);
}

function api_value_(sheetName, a1) {
  const sh = api_sheet_(sheetName);
  if (!sh) return '';
  return sh.getRange(a1).getDisplayValue();
}

function api_rawValue_(sheetName, a1) {
  const sh = api_sheet_(sheetName);
  if (!sh) return '';
  return sh.getRange(a1).getValue();
}

function api_fmtDate_(v) {
  if (!v) return '';
  if (v instanceof Date) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(v);
}

function api_fmtJpDate_(v) {
  if (!v) return '';
  if (v instanceof Date) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'M月d日(E)');
  }
  return String(v);
}

function api_getRows_(sheetName, startRow, startCol, numCols) {
  const sh = api_sheet_(sheetName);
  if (!sh) return [];
  const lastRow = sh.getLastRow();
  if (lastRow < startRow) return [];
  return sh.getRange(startRow, startCol, lastRow - startRow + 1, numCols).getValues();
}

function api_isFalseLike_(value) {
  if (value === false) return true;
  const text = String(value).trim().toLowerCase();
  return text === 'false' || text === '未解決';
}

function api_getHomeTopInfo_() {
  const sheetName = 'ホーム画面🍎';
  const today = new Date();
  const todayLabel = Utilities.formatDate(today, Session.getScriptTimeZone(), 'M月d日(E)');
  const daysUntilEvent = api_value_(sheetName, 'D5');
  const currentPhase = api_value_(sheetName, 'G5');

  return {
    todayLabel,
    daysUntilEvent,
    currentPhase,
    todaySchedule: todayLabel,
  };
}


function api_getTasks_() {
  const sh = api_sheet_('タスクテーブル');
  if (!sh) return [];

  const startRow = 2;
  const lastRow = sh.getLastRow();
  if (lastRow < startRow) return [];

  const values = sh.getRange(startRow, 1, lastRow - startRow + 1, 11).getValues();

  return values.map((r, i) => {
    const no = r[1];
    const title = r[2];
    if (!title) return null;

    return {
      id: no || i + startRow,
      no,
      title,
      taskName: title,
      parentTask: r[3] || '',
      assignee: r[4] || '未定',
      dueDate: api_fmtDate_(r[5]),
      targetDate: api_fmtDate_(r[6]),
      startPlan: api_fmtDate_(r[7]),
      status: r[8] || '',
      progress: r[9] === '' || r[9] === null ? '' : r[9],
      memo: r[10] || '',
    };
  }).filter(Boolean);
}


function api_getTaskDetail_(no) {
  return api_getTasks_().find(t => String(t.no) === String(no)) || null;
}


function api_getTasksDueWithinDays_(days) {
  const today = new Date();
  const start = new Date(today.getFullYear(), today.getMonth(), today.getDate());
  const end = new Date(today.getFullYear(), today.getMonth(), today.getDate() + days);

  return api_getTasks_().filter(t => {
    if (!t.dueDate) return false;
    const d = new Date(t.dueDate);
    if (isNaN(d.getTime())) return false;
    return d >= start && d <= end;
  });
}


function api_getTasksByAssignee_(name) {
  if (!name) return [];
  return api_getTasks_().filter(t => String(t.assignee || '').includes(String(name)));
}


function api_getAssignees_() {
  const names = [];
  api_getTasks_().forEach(t => {
    String(t.assignee || '未定').split(',').forEach(name => {
      const trimmed = name.trim();
      if (trimmed) names.push(trimmed);
    });
  });
  if (!names.includes('未定')) names.push('未定');
  return Array.from(new Set(names)).sort();
}


function api_getWbsTree_() {
  const groups = {};
  api_getTasks_().forEach(t => {
    const parent = String(t.parentTask || '親未設定').trim();
    if (!groups[parent]) groups[parent] = [];
    groups[parent].push(t);
  });
  return Object.keys(groups).map(parentTask => ({ parentTask, children: groups[parentTask] }));
}


function api_getQuestions_(status) {
  const rows = api_getRows_('疑問集約', 3, 2, 6);
  const data = rows.map((r, i) => ({
    id: r[0] || i + 3,
    no: r[0],
    question: r[1] || '',
    owner: r[2] || '',
    due: api_fmtDate_(r[3]),
    answer: r[4] || '',
    resolved: r[5],
  })).filter(q => q.question);

  if (!status) return data;
  if (status === '未解決') return data.filter(q => api_isFalseLike_(q.resolved));
  if (status === '解決済み') return data.filter(q => q.resolved === true || String(q.resolved).toLowerCase() === 'true');
  return data;
}

function api_getUnresolvedQuestions_() {
  return api_getQuestions_('未解決');
}

function api_getMemos_() {
  return api_getRows_('メモ、残したいこと', 2, 2, 3).map((r, i) => ({
    id: r[0] || i + 2,
    no: r[0],
    title: r[1] || '',
    body: r[2] || '',
  })).filter(m => m.title || m.body);
}

function api_getMemoDetail_(id) {
  return api_getMemos_().find(m => String(m.id) === String(id)) || null;
}

function api_getMilestones_() {
  return api_getRows_('マイルストーン、スケジュール', 2, 1, 6).map((r, i) => ({
    id: i + 2,
    date: api_fmtDate_(r[0]),
    title: r[1] || '',
    detail: r[2] || '',
    type: r[3] || '',
  })).filter(r => r.title && String(r.type).toUpperCase() === 'MILESTONE');
}

function api_getSchedule_() {
  return api_getRows_('マイルストーン、スケジュール', 2, 1, 6).map((r, i) => {
    const date = api_fmtDate_(r[0]);
    return {
      id: i + 2,
      date,
      title: r[1] || '',
      detail: r[2] || '',
      type: r[3] || '',
      month: String(date).slice(0, 7),
    };
  }).filter(r => r.title && String(r.type).toUpperCase() !== 'MILESTONE');
}

function api_getGuests_() {
  return api_getRows_('来る人リスト', 2, 2, 5).map((r, i) => ({
    id: r[0] || i + 2,
    no: r[0],
    name: r[1] || '',
    invitedBy: r[2] || '',
    status: r[3] || '',
    probability: r[4] || '',
  })).filter(g => g.name);
}

function api_getHomeSummary_(days) {
  const tasks = api_getTasks_();
  const nearDue = api_getTasksDueWithinDays_(days);
  const unresolved = api_getUnresolvedQuestions_();
  return {
    incompleteTaskCount: tasks.filter(t => {
      const s = String(t.status || '');
      return s !== '完了' && s !== '済' && s !== '完了済み';
    }).length,
    nearDueTaskCount: nearDue.length,
    unresolvedQuestionCount: unresolved.length,
  };
}

function api_getLineShareText_() {
  return { text: api_value_('ホーム画面🍎', 'I5') };
}
