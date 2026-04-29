/**
 * 既存 onEdit ロジックは変更せず、読み取り専用 API を追加するファイル。
 * No.(N列)をキーとして扱う前提を崩さない。
 * O列（関数列）は読み取りのみ。
 */

function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) || '';
    const p = (e && e.parameter) || {};
    const map = {
      getHomeSummary: () => getHomeSummary_(Number(p.days || 7)),
      getTasksDueWithinDays: () => getTasksDueWithinDays_(Number(p.days || 7)),
      getUnresolvedQuestions: () => getUnresolvedQuestions_(),
      getLineShareText: () => getLineShareText_(Number(p.days || 7)),
      getTasks: () => getTasks_(),
      getTaskDetail: () => getTaskDetail_(p.no),
      getTasksByAssignee: () => getTasksByAssignee_(p.name),
      getWbsTree: () => getWbsTree_(),
      getAssignees: () => getAssignees_(),
      getMilestones: () => getMilestones_(),
      getSchedule: () => getSchedule_(),
      getQuestions: () => getQuestions_(p.status),
      getMemos: () => getMemos_(),
      getMemoDetail: () => getMemoDetail_(p.id),
      getGuests: () => getGuests_(),
    };
    if (!map[action]) throw new Error('Unknown action: ' + action);
    return jsonOut_({ ok: true, data: map[action]() });
  } catch (err) {
    return jsonOut_({ ok: false, error: err.message });
  }
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function getSheetValues_(name) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh) throw new Error('Sheet not found: ' + name);
  const values = sh.getDataRange().getValues();
  return values.length > 1 ? values.slice(1) : [];
}

function fmtDate_(v) {
  if (!(v instanceof Date)) return v || '';
  return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function getTasks_() {
  return getSheetValues_('タスク書き出し、進捗').map(function (r) {
    return { no: r[13], taskName: r[14], parentTask: r[3], assignee: r[15], dueDate: fmtDate_(r[16]), targetDate: fmtDate_(r[17]), startPlan: r[18], status: r[21], progress: r[22], memo: r[24] };
  }).filter(function (t) { return t.no; });
}

function getTaskDetail_(no) {
  return getTasks_().filter(function (t) { return String(t.no) === String(no); })[0] || null;
}
function getTasksDueWithinDays_(days) {
  const now = new Date();
  const end = new Date(now.getFullYear(), now.getMonth(), now.getDate() + days);
  return getTasks_().filter(function (t) {
    if (!t.dueDate) return false;
    const d = new Date(t.dueDate);
    return d >= new Date(now.getFullYear(), now.getMonth(), now.getDate()) && d <= end;
  });
}
function getTasksByAssignee_(name) { return getTasks_().filter(function (t) { return t.assignee === name; }); }
function getAssignees_() {
  return Array.from(new Set(getTasks_().map(function (t) { return t.assignee || '未定'; }))).sort();
}
function getWbsTree_() {
  const tasks = getTasks_();
  const groups = {};
  tasks.forEach(function (t) {
    const parent = t.parentTask || '親未設定';
    (groups[parent] ||= []).push(t);
  });
  return Object.keys(groups).map(function (k) { return { parentTask: k, children: groups[k] }; });
}
function getQuestions_(status) {
  const rows = getSheetValues_('疑問集約').map(function (r, i) { return { id: i + 2, question: r[0], category: r[1], assignee: r[2], priority: r[3], status: r[4], memo: r[5] }; });
  return status ? rows.filter(function (q) { return q.status === status; }) : rows;
}
function getUnresolvedQuestions_() { return getQuestions_('未解決'); }
function getMemos_() { return getSheetValues_('メモ、残したいこと').map(function (r, i) { return { id: i + 2, title: r[0], body: r[1] }; }); }
function getMemoDetail_(id) { return getMemos_().filter(function (m) { return String(m.id) === String(id); })[0] || null; }
function getMilestones_() {
  return getSheetValues_('マイルストーン、スケジュール').map(function (r) { return { date: fmtDate_(r[0]), title: r[1], detail: r[2], type: r[3] }; }).filter(function (r) { return r.type === 'MILESTONE'; });
}
function getSchedule_() {
  const days = getSheetValues_('マイルストーン、スケジュール').map(function (r) {
    const date = fmtDate_(r[0]);
    return { date: date, title: r[1], detail: r[2], type: r[3], month: String(date).slice(0, 7) };
  }).filter(function (r) { return r.type !== 'MILESTONE'; });
  return days;
}
function getGuests_() {
  return getSheetValues_('来る人リスト').map(function (r, i) {
    return { id: i + 2, name: r[0], relation: r[1], invitedBy: r[2], status: r[3], contactStatus: r[4], memo: r[5] };
  });
}
function getHomeSummary_(days) {
  const tasks = getTasks_();
  const unresolved = getUnresolvedQuestions_();
  const nearDue = getTasksDueWithinDays_(days);
  return {
    incompleteTaskCount: tasks.filter(function (t) { return t.status !== '完了'; }).length,
    nearDueTaskCount: nearDue.length,
    unresolvedQuestionCount: unresolved.length,
  };
}
function getLineShareText_(days) {
  const s = getHomeSummary_(days);
  return { text: '【林檎の唄。現状】\n未完了タスク: ' + s.incompleteTaskCount + '\n直近タスク: ' + s.nearDueTaskCount + '\n未解決の疑問: ' + s.unresolvedQuestionCount };
}
