const CONFIG = {
  mainSheetName: 'タスク書き出し、進捗',
  storageSheetName: '進捗保存',
  startRow: 5,
  storageStartRow: 2,

  // メイン → 保存
  mainNoCol: 14, // N
  mainTitleCol: 15, // O

  // 編集監視列：P,Q,R,S,V,W,Y
  progressCols: [16, 17, 18, 19, 22, 23, 25],

  // D,E列
  triggerCols: [4, 5],

  // N,O,P,Q,R,S,V,W,Y → A〜I
  // ※O列は「読み取り」はするが、メインシートO列への書き込みはしない
  mainToStorageMap: [
    { main: 14, storage: 1 }, // N -> A
    { main: 15, storage: 2 }, // O -> B
    { main: 16, storage: 3 }, // P -> C
    { main: 17, storage: 4 }, // Q -> D
    { main: 18, storage: 5 }, // R -> E
    { main: 19, storage: 6 }, // S -> F
    { main: 22, storage: 7 }, // V -> G
    { main: 23, storage: 8 }, // W -> H
    { main: 25, storage: 9 }, // Y -> I
  ],

  // 保存 C,D,E,F,G,H,I → メイン P,Q,R,S,V,W,Y
  // ※O列には書き込まない
  restoreMap: [
    { storage: 3, main: 16 }, // C -> P
    { storage: 4, main: 17 }, // D -> Q
    { storage: 5, main: 18 }, // E -> R
    { storage: 6, main: 19 }, // F -> S
    { storage: 7, main: 22 }, // G -> V
    { storage: 8, main: 23 }, // H -> W
    { storage: 9, main: 25 }, // I -> Y
  ],

  // ※O列は除外
  clearColsWhenNoBlank: [16, 17, 18, 19, 22, 23, 25],
  cleanupBlankNoRowsCount: 20,
};

function onEdit(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();

  // 個人タスク確認 B3 の変更は、最優先で処理
  if (
    sheet.getName() === PERSONAL_TASK_CONFIG.personalSheetName &&
    e.range.getA1Notation() === PERSONAL_TASK_CONFIG.userCell
  ) {
    updatePersonalTaskCheck();
    return;
  }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return;

  try {
    const ss = e.source;

    // 個人タスク確認の進捗編集 → 元シート＆進捗保存へ反映
    if (sheet.getName() === PERSONAL_TASK_CONFIG.personalSheetName) {
      syncPersonalTaskEditToMainAndStorage_(ss, sheet, e.range);
      return;
    }

    // ここから下は既存の「タスク書き出し、進捗」用処理
    if (sheet.getName() !== CONFIG.mainSheetName) return;

    const editedRange = e.range;
    const row = editedRange.getRow();
    const col = editedRange.getColumn();
    const numCols = editedRange.getNumColumns();

    if (row < CONFIG.startRow) return;

    const mainSheet = sheet;
    const storageSheet = ss.getSheetByName(CONFIG.storageSheetName);
    if (!storageSheet) throw new Error('進捗保存シートが見つかりません。');

    SpreadsheetApp.flush();

    const editedCols = getEditedCols_(col, numCols);

    const touchedProgress = editedCols.some(c => CONFIG.progressCols.includes(c));
    const touchedDE = editedCols.some(c => CONFIG.triggerCols.includes(c));

    if (touchedProgress) {
      syncEditedProgressToStorage_(mainSheet, storageSheet, editedRange);
    }

    if (touchedDE) {
      handleDEEdit_(mainSheet, storageSheet, editedRange);
    }

  } finally {
    lock.releaseLock();
  }
}

function syncEditedProgressToStorage_(mainSheet, storageSheet, editedRange) {
  const rowStart = editedRange.getRow();
  const rowEnd = rowStart + editedRange.getNumRows() - 1;
  const colStart = editedRange.getColumn();
  const colEnd = colStart + editedRange.getNumColumns() - 1;

  const storageMap = getStorageRowMap_(storageSheet);

  for (let r = rowStart; r <= rowEnd; r++) {
    if (r < CONFIG.startRow) continue;

    const no = mainSheet.getRange(r, CONFIG.mainNoCol).getValue();
    if (no === '' || no === null) continue;

    let storageRow = storageMap[String(no)];

    if (!storageRow) {
      storageRow = getFirstBlankRowInColumn_(storageSheet, 1, CONFIG.storageStartRow);

      const rowValues = CONFIG.mainToStorageMap.map(m => {
        return mainSheet.getRange(r, m.main).getValue();
      });

      storageSheet.getRange(storageRow, 1, 1, 9).setValues([rowValues]);
      storageSheet.getRange(storageRow, 10).setValue(new Date());

      storageMap[String(no)] = storageRow;
      continue;
    }

    CONFIG.mainToStorageMap.forEach(m => {
      if (m.main >= colStart && m.main <= colEnd) {
        const value = mainSheet.getRange(r, m.main).getValue();
        storageSheet.getRange(storageRow, m.storage).setValue(value);
      }
    });

    storageSheet.getRange(storageRow, 10).setValue(new Date());
  }
}

function handleDEEdit_(mainSheet, storageSheet, editedRange) {
  SpreadsheetApp.flush();

  const editedValues = editedRange.getValues();
  const hasNewInput = editedValues.flat().some(v => v !== '' && v !== null);
  const isDeletion = editedValues.flat().every(v => v === '' || v === null);

  restoreAllProgressFromStorage_(mainSheet, storageSheet);

  if (hasNewInput) {
    addNewTaskAsUndecided_(mainSheet, storageSheet, editedRange);
  }

  if (isDeletion) {
    deleteStorageRowsNotInMainN_(mainSheet, storageSheet);
    clearTopBlankNoRows_(mainSheet);
  }
}

function restoreAllProgressFromStorage_(mainSheet, storageSheet) {
  const mainLastRow = mainSheet.getLastRow();
  if (mainLastRow < CONFIG.startRow) return;

  const numRows = mainLastRow - CONFIG.startRow + 1;

  const mainNoValues = mainSheet
    .getRange(CONFIG.startRow, CONFIG.mainNoCol, numRows, 1)
    .getValues()
    .flat();

  const storageData = getStorageData_(storageSheet);
  const storageByNo = {};

  storageData.forEach(item => {
    const no = item.values[0];
    if (no !== '' && no !== null) {
      storageByNo[String(no)] = item.values;
    }
  });

  // O列は関数列なので取得・上書きしない
  const blockPS = mainSheet.getRange(CONFIG.startRow, 16, numRows, 4).getValues(); // P:S
  const blockVW = mainSheet.getRange(CONFIG.startRow, 22, numRows, 2).getValues(); // V:W
  const blockY = mainSheet.getRange(CONFIG.startRow, 25, numRows, 1).getValues(); // Y

  for (let i = 0; i < numRows; i++) {
    const no = mainNoValues[i];
    if (no === '' || no === null) continue;

    const storageRowValues = storageByNo[String(no)];
    if (!storageRowValues) continue;

    // C,D,E,F -> P,Q,R,S
    blockPS[i] = [
      storageRowValues[2],
      storageRowValues[3],
      storageRowValues[4],
      storageRowValues[5],
    ];

    // G,H -> V,W
    blockVW[i] = [
      storageRowValues[6],
      storageRowValues[7],
    ];

    // I -> Y
    blockY[i] = [
      storageRowValues[8],
    ];
  }

  mainSheet.getRange(CONFIG.startRow, 16, numRows, 4).setValues(blockPS);
  mainSheet.getRange(CONFIG.startRow, 22, numRows, 2).setValues(blockVW);
  mainSheet.getRange(CONFIG.startRow, 25, numRows, 1).setValues(blockY);
}

function addNewTaskAsUndecided_(mainSheet, storageSheet, editedRange) {
  const rowStart = editedRange.getRow();
  const rowEnd = rowStart + editedRange.getNumRows() - 1;

  const storageMap = getStorageRowMap_(storageSheet);

  for (let r = rowStart; r <= rowEnd; r++) {
    if (r < CONFIG.startRow) continue;

    const sourceNo = mainSheet.getRange(r, 3).getValue();
    if (sourceNo === '' || sourceNo === null) continue;

    const targetRow = findRowByValue_(mainSheet, CONFIG.mainNoCol, CONFIG.startRow, sourceNo);
    if (!targetRow) continue;

    mainSheet.getRange(targetRow, 16).setValue('未定');

    let storageRow = storageMap[String(sourceNo)];

    if (!storageRow) {
      storageRow = getFirstBlankRowInColumn_(storageSheet, 1, CONFIG.storageStartRow);
      storageSheet.getRange(storageRow, 1).setValue(sourceNo);
      storageSheet.getRange(storageRow, 3).setValue('未定');
      storageSheet.getRange(storageRow, 10).setValue(new Date());
      storageMap[String(sourceNo)] = storageRow;
    } else {
      storageSheet.getRange(storageRow, 3).setValue('未定');
      storageSheet.getRange(storageRow, 10).setValue(new Date());
    }
  }
}

function deleteStorageRowsNotInMainN_(mainSheet, storageSheet) {
  const mainLastRow = mainSheet.getLastRow();
  const mainNos = new Set();

  if (mainLastRow >= CONFIG.startRow) {
    mainSheet
      .getRange(CONFIG.startRow, CONFIG.mainNoCol, mainLastRow - CONFIG.startRow + 1, 1)
      .getValues()
      .flat()
      .forEach(v => {
        if (v !== '' && v !== null) mainNos.add(String(v));
      });
  }

  const storageLastRow = storageSheet.getLastRow();
  if (storageLastRow < CONFIG.storageStartRow) return;

  const storageNos = storageSheet
    .getRange(CONFIG.storageStartRow, 1, storageLastRow - CONFIG.storageStartRow + 1, 1)
    .getValues()
    .flat();

  for (let i = storageNos.length - 1; i >= 0; i--) {
    const no = storageNos[i];
    if (no === '' || no === null) continue;

    if (!mainNos.has(String(no))) {
      storageSheet.deleteRow(CONFIG.storageStartRow + i);
    }
  }
}

function clearTopBlankNoRows_(mainSheet) {
  const lastRow = mainSheet.getLastRow();
  if (lastRow < CONFIG.startRow) return;

  const numRows = lastRow - CONFIG.startRow + 1;
  const noValues = mainSheet
    .getRange(CONFIG.startRow, CONFIG.mainNoCol, numRows, 1)
    .getValues()
    .flat();

  let clearedCount = 0;

  for (let i = 0; i < noValues.length; i++) {
    if (clearedCount >= CONFIG.cleanupBlankNoRowsCount) break;

    const no = noValues[i];
    if (no !== '' && no !== null) continue;

    const row = CONFIG.startRow + i;

    CONFIG.clearColsWhenNoBlank.forEach(col => {
      const cell = mainSheet.getRange(row, col);
      const formula = cell.getFormula();

      if (!formula) {
        cell.clearContent();
      }
    });

    clearedCount++;
  }
}

function getStorageRowMap_(storageSheet) {
  const map = {};
  const lastRow = storageSheet.getLastRow();

  if (lastRow < CONFIG.storageStartRow) return map;

  const values = storageSheet
    .getRange(CONFIG.storageStartRow, 1, lastRow - CONFIG.storageStartRow + 1, 1)
    .getValues();

  values.forEach((row, i) => {
    const no = row[0];
    if (no !== '' && no !== null) {
      map[String(no)] = CONFIG.storageStartRow + i;
    }
  });

  return map;
}

function getStorageData_(storageSheet) {
  const lastRow = storageSheet.getLastRow();
  if (lastRow < CONFIG.storageStartRow) return [];

  const values = storageSheet
    .getRange(CONFIG.storageStartRow, 1, lastRow - CONFIG.storageStartRow + 1, 10)
    .getValues();

  return values.map((row, i) => ({
    rowNumber: CONFIG.storageStartRow + i,
    values: row,
  }));
}

function getFirstBlankRowInColumn_(sheet, col, startRow) {
  const lastRow = Math.max(sheet.getLastRow(), startRow);
  const values = sheet.getRange(startRow, col, lastRow - startRow + 1, 1).getValues();

  for (let i = 0; i < values.length; i++) {
    if (values[i][0] === '' || values[i][0] === null) {
      return startRow + i;
    }
  }

  return lastRow + 1;
}

function findRowByValue_(sheet, col, startRow, value) {
  const lastRow = sheet.getLastRow();
  if (lastRow < startRow) return null;

  const values = sheet.getRange(startRow, col, lastRow - startRow + 1, 1).getValues();

  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]) === String(value)) {
      return startRow + i;
    }
  }

  return null;
}

function getEditedCols_(startCol, numCols) {
  const cols = [];
  for (let c = startCol; c < startCol + numCols; c++) {
    cols.push(c);
  }
  return cols;
}


const PERSONAL_TASK_CONFIG = {
  personalSheetName: '個人タスク確認',
  taskSheetName: 'タスク書き出し、進捗',
  masterSheetName: 'マスタ',

  userCell: 'B3',
  headerRow: 4,
  outputStartRow: 5,

  taskStartRow: 5,

  // 元シート → 転記先
  sourceCols: [14, 15, 16, 17, 18, 19, 22, 23, 25], // N,O,P,Q,R,S,V,W,Y
  outputCols: [4, 5, 6, 7, 8, 9, 12, 13, 15],      // D,E,F,G,H,I,L,M,O ←修正

  taskAssigneeCol: 16, // P列

  masterStartRow: 2,
  masterGroupCol: 14,       // N列
  masterMemberStartCol: 15, // O列
  masterMemberEndCol: 20    // T列
};

function updatePersonalTaskCheck() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const personalSheet = ss.getSheetByName(PERSONAL_TASK_CONFIG.personalSheetName);
  const taskSheet = ss.getSheetByName(PERSONAL_TASK_CONFIG.taskSheetName);
  const masterSheet = ss.getSheetByName(PERSONAL_TASK_CONFIG.masterSheetName);

  if (!personalSheet || !taskSheet || !masterSheet) {
    throw new Error('必要なシートが見つかりません。');
  }

  const userName = String(personalSheet.getRange(PERSONAL_TASK_CONFIG.userCell).getValue()).trim();

  // 先にクリア（指定列のみ）
  clearPersonalTaskOutput_(personalSheet);

  if (!userName) return;

  const targetNames = getTargetNamesFromMaster_(masterSheet, userName);
  targetNames.add(userName);

  const results = collectMatchingTasks_(taskSheet, targetNames);

  if (results.length === 0) return;

  writePersonalTaskOutput_(personalSheet, results);
}

/**
 * 指定列のみクリア（5行目以降）
 */
function clearPersonalTaskOutput_(sheet) {
  const startRow = PERSONAL_TASK_CONFIG.outputStartRow;
  const lastRow = Math.max(sheet.getLastRow(), startRow);

  if (lastRow < startRow) return;

  const numRows = lastRow - startRow + 1;

  PERSONAL_TASK_CONFIG.outputCols.forEach(col => {
    sheet.getRange(startRow, col, numRows, 1).clearContent();
  });
}

/**
 * マスタから所属グループ取得
 */
function getTargetNamesFromMaster_(masterSheet, userName) {
  const groups = new Set();

  const startRow = PERSONAL_TASK_CONFIG.masterStartRow;
  const lastRow = masterSheet.getLastRow();

  if (lastRow < startRow) return groups;

  const numRows = lastRow - startRow + 1;
  const startCol = PERSONAL_TASK_CONFIG.masterGroupCol;
  const numCols =
    PERSONAL_TASK_CONFIG.masterMemberEndCol -
    PERSONAL_TASK_CONFIG.masterGroupCol + 1;

  const values = masterSheet.getRange(startRow, startCol, numRows, numCols).getValues();

  values.forEach(row => {
    const groupName = String(row[0]).trim();
    if (!groupName) return;

    const members = row.slice(1).map(v => String(v).trim()).filter(Boolean);

    const isMember = members.some(member =>
      member.includes(userName) || userName.includes(member)
    );

    if (isMember) {
      groups.add(groupName);
    }
  });

  return groups;
}

/**
 * 該当タスク抽出
 */
function collectMatchingTasks_(taskSheet, targetNames) {
  const startRow = PERSONAL_TASK_CONFIG.taskStartRow;
  const lastRow = taskSheet.getLastRow();

  if (lastRow < startRow) return [];

  const numRows = lastRow - startRow + 1;

  // N〜Y列取得
  const values = taskSheet.getRange(startRow, 14, numRows, 12).getValues();

  const results = [];

  values.forEach(row => {
    const assigneeText = String(row[2]).trim(); // P列

    if (!assigneeText) return;

    const matched = Array.from(targetNames).some(name => {
      const t = String(name).trim();
      return t && assigneeText.includes(t);
    });

    if (!matched) return;

    results.push([
      row[0],  // N → D
      row[1],  // O → E
      row[2],  // P → F
      row[3],  // Q → G
      row[4],  // R → H
      row[5],  // S → I
      row[8],  // V → L ←修正
      row[9],  // W → M ←修正
      row[11]  // Y → O ←修正
    ]);
  });

  return results;
}

/**
 * 転記（列ごとに書き込み）
 */
function writePersonalTaskOutput_(sheet, results) {
  const startRow = PERSONAL_TASK_CONFIG.outputStartRow;
  const outputCols = PERSONAL_TASK_CONFIG.outputCols;

  outputCols.forEach((col, index) => {
    const colValues = results.map(row => [row[index]]);
    sheet.getRange(startRow, col, colValues.length, 1).setValues(colValues);
  });
}

function syncPersonalTaskEditToMainAndStorage_(ss, personalSheet, editedRange) {
  const startRow = PERSONAL_TASK_CONFIG.outputStartRow;

  const personalNoCol = 4; // D列：No.

  const personalToMainAndStorageMap = {
    6:  { main: 16, storage: 3 }, // F → P / C
    7:  { main: 17, storage: 4 }, // G → Q / D
    8:  { main: 18, storage: 5 }, // H → R / E
    9:  { main: 19, storage: 6 }, // I → S / F
    12: { main: 22, storage: 7 }, // L → V / G
    13: { main: 23, storage: 8 }, // M → W / H
    15: { main: 25, storage: 9 }, // O → Y / I
  };

  const editedRowStart = editedRange.getRow();
  const editedColStart = editedRange.getColumn();
  const numRows = editedRange.getNumRows();
  const numCols = editedRange.getNumColumns();

  if (editedRowStart < startRow) return;

  const editedValues = editedRange.getValues();

  const mainSheet = ss.getSheetByName(PERSONAL_TASK_CONFIG.taskSheetName);
  const storageSheet = ss.getSheetByName(CONFIG.storageSheetName);

  if (!mainSheet) throw new Error('タスク書き出し、進捗シートが見つかりません。');
  if (!storageSheet) throw new Error('進捗保存シートが見つかりません。');

  const mainNoRowMap = getMainNoRowMap_(mainSheet);
  const storageNoRowMap = getStorageRowMap_(storageSheet);

  for (let r = 0; r < numRows; r++) {
    const sheetRow = editedRowStart + r;
    if (sheetRow < startRow) continue;

    const no = personalSheet.getRange(sheetRow, personalNoCol).getValue();
    if (no === '' || no === null) continue;

    const mainRow = mainNoRowMap[String(no)];
    const storageRow = storageNoRowMap[String(no)];

    for (let c = 0; c < numCols; c++) {
      const sheetCol = editedColStart + c;
      const map = personalToMainAndStorageMap[sheetCol];

      // F,G,H,I,L,M,O 以外は無視
      if (!map) continue;

      const value = editedValues[r][c];

      // タスク書き出し、進捗へ反映
      if (mainRow) {
        mainSheet.getRange(mainRow, map.main).setValue(value);
      }

      // 進捗保存へ反映
      if (storageRow) {
        storageSheet.getRange(storageRow, map.storage).setValue(value);
        storageSheet.getRange(storageRow, 10).setValue(new Date());
      }
    }
  }
}

function getMainNoRowMap_(mainSheet) {
  const map = {};
  const startRow = CONFIG.startRow;
  const lastRow = mainSheet.getLastRow();

  if (lastRow < startRow) return map;

  const values = mainSheet
    .getRange(startRow, CONFIG.mainNoCol, lastRow - startRow + 1, 1)
    .getValues();

  values.forEach((row, i) => {
    const no = row[0];
    if (no !== '' && no !== null) {
      map[String(no)] = startRow + i;
    }
  });

  return map;
}

function doGet(e) {
  try {
    const p = e && e.parameter ? e.parameter : {};
    const action = p.action || '';

    const actions = {
      getHomeSummary: () => api_getHomeSummary_(Number(p.days || 7)),
      getTasksDueWithinDays: () => api_getTasksDueWithinDays_(Number(p.days || 7)),
      getUnresolvedQuestions: () => api_getUnresolvedQuestions_(),
      getLineShareText: () => api_getLineShareText_(Number(p.days || 7)),
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
      addMemo: () => api_addMemo_(JSON.parse(p.data || '{}')),
      updateMemo: () => api_updateMemo_(p.id, JSON.parse(p.data || '{}')),
      getGuests: () => api_getGuests_(),
      getHomeTopInfo: () => api_getHomeTopInfo_(),
      setHomeDueDays: () => api_setHomeDueDays_(p.days),
      addQuestion: () => api_addQuestion_(JSON.parse(p.data || '{}')),
      updateQuestion: () => api_updateQuestion_(p.id, JSON.parse(p.data || '{}')),
      getQuestionOptions: () => api_getQuestionOptions_(),
    };

    if (!actions[action]) {
      throw new Error('Unknown action: ' + action);
    }

return api_json_({
  ok: true,
  data: actions[action](),
}, p.callback);

  } catch (err) {
const callback = e && e.parameter ? e.parameter.callback : '';

return api_json_({
  ok: false,
  error: err.message,
}, callback);
  }
}

function api_json_(obj, callback) {
  const json = JSON.stringify(obj);

  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

function api_currentCallback_() {
  try {
    const params = PropertiesService.getScriptProperties().getProperty('__dummy__');
  } catch (e) {}

  return '';
}

function api_getHomeTopInfo_() {
  const sh = api_sheet_('ホーム画面🍎');
  if (!sh) {
    return {
      todaySchedule: '',
      daysUntilEvent: '',
      currentPhase: '',
    };
  }

  return {
    todaySchedule: api_getTodayLabel_(),
    daysUntilEvent: sh.getRange('D5').getDisplayValue(),
    currentPhase: sh.getRange('G5').getDisplayValue(),
    dueDays: sh.getRange('D7').getValue(),
  };
}

function api_getTodayLabel_() {
  const days = ['日', '月', '火', '水', '木', '金', '土'];
  const d = new Date();
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'M月d日') + '（' + days[d.getDay()] + '）';
}

function api_ss_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function api_sheet_(sheetName) {
  return api_ss_().getSheetByName(sheetName);
}

function api_fmtDate_(v) {
  if (!v) return '';
  if (v instanceof Date) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(v);
}

function api_getRows_(sheetName, startRow) {
  const sh = api_sheet_(sheetName);
  if (!sh) return [];

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();

  if (lastRow < startRow || lastCol < 1) return [];

  return sh.getRange(startRow, 1, lastRow - startRow + 1, lastCol).getValues();
}

function api_getTasks_() {
  return api_getTasksFromTable_();
}

function api_getTaskDetail_(no) {
  return api_getTasks_().find(t => String(t.no) === String(no)) || null;
}

function api_getTasksDueWithinDays_(days) {
  const limit = Number(days || 7);

  return api_getTasks_().filter(t => {
    const raw = t.daysUntilDue;

    if (!String(raw).trim()) return false;

    const n = Number(raw);
    if (Number.isNaN(n)) return false;

    return n <= limit;
  });
}

function api_getTasksByAssignee_(name) {
  if (!name) return [];

  const masterSheet = api_sheet_(PERSONAL_TASK_CONFIG.masterSheetName);
  const targetNames = masterSheet
    ? getTargetNamesFromMaster_(masterSheet, name)
    : new Set();

  targetNames.add(name);

  return api_getTasks_().filter(t => {
    const assignee = String(t.assignee || '').trim();
    return Array.from(targetNames).some(target => {
      const n = String(target || '').trim();
      return n && assignee.includes(n);
    });
  });
}

function api_getAssignees_() {
  const names = api_getTasks_()
    .map(t => String(t.assignee || '未定').trim())
    .filter(Boolean);

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

  return Object.keys(groups).map(parentTask => ({
    parentTask,
    children: groups[parentTask],
  }));
}

function api_getQuestions_(status) {
  const sh = api_sheet_('疑問集約');
  if (!sh) return [];

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  // B〜G列を取得
  const rows = sh.getRange(2, 2, lastRow - 1, 6).getValues();

  const data = rows.map((r, i) => ({
    id: i + 2,
    no: r[0] || '',
    question: r[1] || '',
    owner: r[2] || '',
    due: api_fmtDate_(r[3]),
    answer: r[4] || '',
    resolved: r[5],
  })).filter(q => q.question);

  if (!status) return data;

  if (status === 'unresolved') {
    return data.filter(q => q.resolved === false || String(q.resolved).toLowerCase() === 'false');
  }

  if (status === 'resolved') {
    return data.filter(q => q.resolved === true || String(q.resolved).toLowerCase() === 'true');
  }

  return data;
}

function api_getUnresolvedQuestions_() {
  return api_getQuestions_('unresolved');
}

function api_getMemos_() {
  const sh = api_sheet_('メモ、残したいこと');
  if (!sh) return [];

  const lastRow = sh.getLastRow();
  if (lastRow < 3) return [];

  // B〜D列：No. / タイトル / 本文
  const rows = sh.getRange(3, 2, lastRow - 2, 3).getValues();

  return rows.map((r, i) => ({
    id: i + 3,
    no: r[0] || '',
    title: r[1] || '',
    body: r[2] || '',
  })).filter(m => m.title || m.body);
}

function api_getMemoDetail_(id) {
  return api_getMemos_().find(m => String(m.id) === String(id)) || null;
}

function api_addMemo_(data) {
  const sh = api_sheet_('メモ、残したいこと');
  if (!sh) throw new Error('メモ、残したいことシートが見つかりません。');

  const startRow = 3;
  const lastRow = Math.max(sh.getLastRow(), startRow);

  // C〜D列だけを見る。B列はNo.が入っているため空白判定に含めない
  const rows = sh.getRange(startRow, 3, lastRow - startRow + 1, 2).getDisplayValues();

  let targetRow = null;

  for (let i = 0; i < rows.length; i++) {
    const isBlank = rows[i].every(v => String(v).trim() === '');
    if (isBlank) {
      targetRow = startRow + i;
      break;
    }
  }

  if (!targetRow) {
    targetRow = lastRow + 1;

    const noValues = sh.getRange(startRow, 2, lastRow - startRow + 1, 1).getValues().flat();
    const maxNo = noValues.reduce((max, v) => {
      const n = Number(v);
      return Number.isFinite(n) ? Math.max(max, n) : max;
    }, 0);

    sh.getRange(targetRow, 2).setValue(maxNo + 1);
  }

  sh.getRange(targetRow, 3, 1, 2).setValues([[
    data.title || '',
    data.body || ''
  ]]);

  return { ok: true, row: targetRow };
}

function api_updateMemo_(id, data) {
  const sh = api_sheet_('メモ、残したいこと');
  if (!sh) throw new Error('メモ、残したいことシートが見つかりません。');

  const row = Number(id);
  if (!Number.isFinite(row) || row < 3) {
    throw new Error('更新対象の行が不正です。');
  }

  sh.getRange(row, 3).setValue(data.title || '');
  sh.getRange(row, 4).setValue(data.body || '');

  return { ok: true };
}

function api_getMilestones_() {
  const rows = api_getRows_('マイルストーン、スケジュール', 2);

  return rows.map((r, i) => ({
    id: i + 2,
    date: api_fmtDate_(r[0]),
    title: r[1] || '',
    detail: r[2] || '',
    type: r[3] || '',
  })).filter(r => {
    return r.title && String(r.type).toUpperCase() === 'MILESTONE';
  });
}

function api_getSchedule_() {
  const rows = api_getRows_('マイルストーン、スケジュール', 2);

  return rows.map((r, i) => {
    const date = api_fmtDate_(r[0]);
    return {
      id: i + 2,
      date,
      title: r[1] || '',
      detail: r[2] || '',
      type: r[3] || '',
      month: String(date).slice(0, 7),
    };
  }).filter(r => {
    return r.title && String(r.type).toUpperCase() !== 'MILESTONE';
  });
}

function api_getGuests_() {
  const rows = api_getRows_('来る人リスト', 2);

  return rows.map((r, i) => ({
    id: i + 2,
    name: r[0] || '',
    relation: r[1] || '',
    invitedBy: r[2] || '',
    status: r[3] || '',
    contactStatus: r[4] || '',
    memo: r[5] || '',
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

function api_getLineShareText_(days) {
  const sh = api_sheet_('ホーム画面🍎');
  if (!sh) return { text: '' };

  return {
    text: sh.getRange('I5').getDisplayValue()
  };
}

function api_getTasksFromTable_() {
  const sh = api_sheet_('タスクテーブル');
  if (!sh) return [];

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  // A〜M列
  const values = sh.getRange(2, 1, lastRow - 1, 13).getValues();

  return values.map(row => ({
    no: row[0],
    taskName: row[1],
    parentTask: row[2],
    assignee: row[3],
    dueDate: api_fmtDate_(row[4]),
    targetDate: api_fmtDate_(row[5]),
    startPlan: api_fmtDate_(row[6]),
    status: row[7],
    progress: row[8],
    memo: row[9],
    daysUntilDue: row[10],
    id: row[11],
    level: row[12],
    parentId: row[11] ? String(row[11]).split('-').slice(0, -1).join('-') : '',
  })).filter(t => t.taskName);
}

function api_setHomeDueDays_(days) {
  const sh = api_sheet_('ホーム画面🍎');
  if (!sh) throw new Error('ホーム画面🍎シートが見つかりません。');

  const n = Number(days);
  if (!Number.isFinite(n) || n < 1) {
    throw new Error('日数は1以上の数字で入力してください。');
  }

  sh.getRange('D7').setValue(n);

  return {
    days: n,
    message: 'ホーム画面🍎!D7を更新しました。'
  };
}

function api_addQuestion_(data) {
  const sh = api_sheet_('疑問集約');
  if (!sh) throw new Error('疑問集約シートが見つかりません。');

  const startRow = 3;
  const lastRow = Math.max(sh.getLastRow(), startRow);

  // C〜F列だけを見る
  // G列はチェックボックスで false が入るため、空白判定に含めない
  const rows = sh.getRange(startRow, 3, lastRow - startRow + 1, 4).getDisplayValues();

  let targetRow = null;

  for (let i = 0; i < rows.length; i++) {
    const isBlank = rows[i].every(v => String(v).trim() === '');
    if (isBlank) {
      targetRow = startRow + i;
      break;
    }
  }

  if (!targetRow) targetRow = lastRow + 1;

  // C〜G列に書き込み
  sh.getRange(targetRow, 3, 1, 5).setValues([[
    data.question || '',
    data.owner || '',
    data.due || '',
    data.answer || '',
    false
  ]]);

  return { ok: true, row: targetRow };
}

function api_updateQuestion_(id, data) {
  const sh = api_sheet_('疑問集約');
  if (!sh) throw new Error('疑問集約シートが見つかりません。');

  const row = Number(id);
  if (!Number.isFinite(row) || row < 2) {
    throw new Error('更新対象の行が不正です。');
  }

  sh.getRange(row, 3).setValue(data.question || '');
  sh.getRange(row, 4).setValue(data.owner || '');
  sh.getRange(row, 5).setValue(data.due || '');
  sh.getRange(row, 6).setValue(data.answer || '');
  sh.getRange(row, 7).setValue(data.resolved === true);

  return { ok: true };
}

function api_getQuestionOptions_() {
  const sh = api_sheet_('マスタ');
  if (!sh) throw new Error('マスタシートが見つかりません。');

  const owners = sh.getRange('O3:O8').getValues()
    .flat()
    .map(v => String(v).trim())
    .filter(Boolean);

  const lastRow = sh.getLastRow();
  const dueOptions = lastRow >= 3
    ? sh.getRange(3, 12, lastRow - 2, 1).getValues().flat()
    : [];

  const dues = dueOptions
    .map(v => String(v).trim())
    .filter(Boolean);

  return {
    owners: Array.from(new Set(owners)),
    dues: Array.from(new Set(dues)),
  };
}