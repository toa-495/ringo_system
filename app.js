const state = {
  currentView: 'home',
  currentDays: 7,
  lineShareText: '',
  questionStatus: 'unresolved',
  questionOptions: { owners: [], dues: [] },
  homeUnresolvedQuestions: [],
  taskStatusTab: 'incomplete',
  taskAssigneeFilter: '',
  taskFilterUsers: [],
  allTasksForWbs: [],
  taskAssigneeCache: {},
allQuestions: null,
allMemos: null,
allGuests: null,
milestoneCache: null,
calendarCache: null,
};

const VIEW_TITLES = {
  home: 'ホーム',
  tasks: 'タスク関連',
  milestones: 'マイルストーン',
  calendar: 'カレンダー',
  questions: '疑問箱',
  memos: 'メモページ',
  expenses: '経費登録',
  guests: '来る人リスト',
};

const el = {
  loading: document.getElementById('loading'),
  error: document.getElementById('error'),
  pageTitle: document.getElementById('page-title'),
  sidebar: document.getElementById('sidebar'),
  sidebarBackdrop: document.getElementById('sidebar-backdrop'),
  sidebarOpen: document.getElementById('sidebar-open'),
  sidebarClose: document.getElementById('sidebar-close'),
  views: {
    home: document.getElementById('view-home'),
    tasks: document.getElementById('view-tasks'),
    milestones: document.getElementById('view-milestones'),
    calendar: document.getElementById('view-calendar'),
    questions: document.getElementById('view-questions'),
    memos: document.getElementById('view-memos'),
    expenses: document.getElementById('view-expenses'),
    guests: document.getElementById('view-guests'),
  },
  navButtons: [...document.querySelectorAll('.nav-btn')],
  modal: document.getElementById('modal'),
  modalBody: document.getElementById('modal-body'),
  modalClose: document.getElementById('modal-close'),
};

function apiGet(action, params = {}) {
  return new Promise((resolve, reject) => {
    const callbackName = `jsonp_${Date.now()}_${Math.floor(Math.random() * 100000)}`;
    const url = new URL(window.APP_CONFIG.GAS_BASE_URL);
    url.searchParams.set('action', action);
    url.searchParams.set('callback', callbackName);

    Object.entries(params).forEach(([key, value]) => {
      if (value !== undefined && value !== null) url.searchParams.set(key, value);
    });

    const script = document.createElement('script');

    window[callbackName] = (json) => {
      delete window[callbackName];
      script.remove();

      if (!json || !json.ok) {
        reject(new Error(json?.error || 'APIエラー'));
        return;
      }

      resolve(json.data);
    };

    script.onerror = () => {
      delete window[callbackName];
      script.remove();
      reject(new Error('GAS APIの読み込みに失敗しました。GAS側がJSONP形式で返しているか確認してください。'));
    };

    script.src = url.toString();
    document.body.appendChild(script);
  });
}

let appleLoadingTimer = null;

function setAppleProgress(percent) {
  const mask = document.querySelector('#loading .apple-red-mask');
  if (!mask) return;

  const safePercent = Math.max(0, Math.min(100, Number(percent) || 0));
  mask.style.height = `${safePercent}%`;
}

function setLoading(show) {
  if (!el.loading) return;

  el.loading.classList.toggle('hidden', !show);

  if (appleLoadingTimer) {
    clearInterval(appleLoadingTimer);
    appleLoadingTimer = null;
  }

  if (show) {
    let progress = 0;
    let dotCount = 1;

    setAppleProgress(0);

    appleLoadingTimer = setInterval(() => {
      // 一気に満ちないように、後半ほどゆっくり進む
      const next = progress + Math.max(1.2, (92 - progress) * 0.08);
      progress = Math.min(next, 92);
      setAppleProgress(progress);

      const dots = document.getElementById('loading-dots');
      if (dots) {
        dots.textContent = '.'.repeat(dotCount);
        dotCount = dotCount >= 3 ? 1 : dotCount + 1;
      }
    }, 220);
  } else {
    setAppleProgress(100);

    const dots = document.getElementById('loading-dots');
    if (dots) dots.textContent = '...';

    setTimeout(() => {
      el.loading.classList.add('hidden');
      setAppleProgress(0);
    }, 350);
  }
}

function setError(message = '') {
  el.error.textContent = message;
  el.error.classList.toggle('hidden', !message);
}

function escapeHtml(value) {
  return String(value ?? '').replace(/[&<>'"]/g, (c) => ({
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    "'": '&#39;',
    '"': '&quot;',
  }[c]));
}

function openSidebar() {
  el.sidebar.classList.add('open');
  el.sidebarBackdrop.classList.remove('hidden');
}

function closeSidebar() {
  el.sidebar.classList.remove('open');
  el.sidebarBackdrop.classList.add('hidden');
}

function openModal(contentHtml) {
  el.modalBody.innerHTML = contentHtml;
  el.modal.classList.remove('hidden');
}

function closeModal() {
  el.modal.classList.add('hidden');
}

function placeholder(view) {
  return `
    <section class="empty-state">
      <p class="eyebrow">Next Step</p>
      <h3>${escapeHtml(VIEW_TITLES[view])}</h3>
      <p>この画面は次の実装フェーズで作ります。まずはホーム画面の接続と表示を固めます。</p>
    </section>
  `;
}

function getTodayJapaneseLabel() {
  const now = new Date();
  const weekdays = ['日', '月', '火', '水', '木', '金', '土'];
  return `${now.getMonth() + 1}月${now.getDate()}日（${weekdays[now.getDay()]}）`;
}

function renderTodaySchedule(text) {
  return `<p class="today-date-text">${escapeHtml(text || getTodayJapaneseLabel())}</p>`;
}


function renderTaskRows(tasks) {
  if (!tasks || tasks.length === 0) {
    return '<p class="meta">指定日数以内のタスクはありません。</p>';
  }

  return tasks.map((task) => {
    const title = escapeHtml(task.title || task.taskName || '無題のタスク');
    const assignee = escapeHtml(task.assignee || '未定');
    const dueDate = escapeHtml(task.dueDate || task.limit || '期限なし');

    return `
      <button class="data-row task-row" type="button" data-modal="task" data-payload='${escapeHtml(JSON.stringify(task))}'>
        <div class="data-main">
          <strong>${title}</strong>
          <span class="meta">担当：${assignee}</span>
        </div>
        <div class="data-sub">
          <span>${dueDate}</span>
        </div>
      </button>
    `;
  }).join('');
}


function renderQuestionRows(questions) {
  if (!questions || questions.length === 0) {
    return '<p class="meta">未解決の疑問はありません。</p>';
  }

  return questions.map(q => `
    <div class="data-row question-manage-row home-question-row" data-home-question-detail="${escapeHtml(q.id)}">
      <div class="data-main">
        <strong>${escapeHtml(q.question || '無題の疑問')}</strong>
        <span class="meta">
          疑問ぬし：${escapeHtml(q.owner || '未定')}
          ${q.due ? ` / ${escapeHtml(q.due)}` : ''}
        </span>
        ${q.answer ? `<span class="meta">回答：${escapeHtml(q.answer)}</span>` : ''}
      </div>

      <button class="question-edit-btn home-question-edit-btn" type="button" data-home-question-edit="${escapeHtml(q.id)}">
        ✒
      </button>
    </div>
  `).join('');
}

function isBlankValue(value) {
  return value === undefined || value === null || String(value).trim() === '';
}

function renderMissingValue() {
  return '<span class="detail-missing">未</span>';
}

function renderDetailValue(value) {
  if (isBlankValue(value)) return renderMissingValue();
  return escapeHtml(value);
}

function formatDetailDate(value) {
  if (isBlankValue(value)) return '';

  const text = String(value).trim();

  // yyyy-mm-dd / yyyy/mm/dd 対応
  const match = text.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
  if (match) {
    return `${Number(match[2])}月${Number(match[3])}日`;
  }

  // すでに 〇月〇日 の場合はそのまま
  if (/\d{1,2}月\d{1,2}日/.test(text)) return text;

  const date = new Date(text);
  if (!Number.isNaN(date.getTime())) {
    return `${date.getMonth() + 1}月${date.getDate()}日`;
  }

  return text;
}

function renderTaskDetail(payload) {
  const no = payload.no || '-';
  const title = payload.title || payload.taskName || '';
  const assignee = payload.assignee || '';
  const parentTask = getParentTaskLabel(payload);
  const dueDate = payload.dueDate || '';
  const targetDate = payload.targetDate || '';
  const startPlan = payload.startPlan || '';
  const daysUntilDue = formatDaysUntilDue(payload.daysUntilDue);
  const status = payload.status || '';
  const progress = normalizeProgress(payload.progress);
  const memo = payload.memo || '';

  return `
    <div class="task-detail-card">
      <div class="task-detail-main">
        <div class="task-detail-no">
          <span>No.</span>
          <strong>${escapeHtml(no)}</strong>
        </div>

        <div class="task-detail-title-block">
          <div class="task-detail-field task-detail-title-field">
            <span class="task-detail-label">タイトル</span>
            <strong>${renderDetailValue(title)}</strong>
          </div>

          <div class="task-detail-field">
            <span class="task-detail-label">担当者</span>
            <strong>${renderDetailValue(assignee)}</strong>
          </div>
        </div>
      </div>

      <div class="task-detail-parent-line">
        <span>親タスク</span>
        <strong>${renderDetailValue(parentTask)}</strong>
      </div>

      <div class="task-detail-mini-grid">
        <div><span>絶対！期日</span><strong>${renderDetailValue(formatDetailDate(dueDate))}</strong></div>
        <div><span>目標期日</span><strong>${renderDetailValue(formatDetailDate(targetDate))}</strong></div>
        <div><span>着手予定時期</span><strong>${renderDetailValue(startPlan)}</strong></div>
        <div><span>作業日数残</span><strong>${renderDetailValue(daysUntilDue)}</strong></div>
      </div>

      <div class="task-detail-status-grid">
        <div><span>進捗状態</span><strong>${renderDetailValue(status)}</strong></div>
        <div><span>進捗%</span><strong>${progress}%</strong></div>
      </div>

      <div class="task-detail-progress">
        ${renderProgressBar(progress)}
      </div>

      <div class="task-detail-memo">
        <span>進捗詳細・メモ</span>
        <p>${renderDetailValue(memo)}</p>
      </div>

      <div class="task-detail-actions">
  <button class="btn-primary task-action-main" type="button" data-task-edit='${escapeHtml(JSON.stringify(payload))}'>
    編集する
  </button>

  <div class="task-action-sub">
    <button class="btn-secondary" type="button" data-task-add-child='${escapeHtml(JSON.stringify(payload))}'>
      傘下にタスク追加
    </button>

    <button class="btn-danger" type="button" data-task-delete='${escapeHtml(JSON.stringify(payload))}'>
      削除
    </button>
  </div>
</div>
    </div>
  `;
}


function bindRowModals() {
  document.querySelectorAll('[data-modal="task"][data-payload]').forEach((button) => {
    button.addEventListener('click', () => {
      const payload = JSON.parse(button.dataset.payload);
      openModal(renderTaskDetail(payload));
      bindTaskDetailActions(payload);
    });
  });
}

function bindTaskDetailActions(task) {
  const editButton = document.querySelector('[data-task-edit]');
  if (editButton) {
    editButton.addEventListener('click', async () => {
      const payload = JSON.parse(editButton.dataset.taskEdit);

      try {
        setLoading(true);
        const [options, parentOptionsRaw] = await Promise.all([
  apiGet('getTaskEditOptions'),
  apiGet('getParentTaskOptions'),
]);

const parentOptions = parentOptionsRaw.filter(item => {
  return String(item.no) !== String(payload.no);
});

openModal(renderTaskEditForm(payload, options, parentOptions));
bindTaskEditForm(payload, options, parentOptions);
      } catch (err) {
        alert(err.message || '編集用の選択肢取得に失敗しました。');
      } finally {
        setLoading(false);
      }
    });
  }

  const deleteButton = document.querySelector('[data-task-delete]');
  if (deleteButton) {
    deleteButton.addEventListener('click', async () => {
      const payload = JSON.parse(deleteButton.dataset.taskDelete);

      const ok = confirm(`「${payload.taskName || 'このタスク'}」を削除しますか？`);
      if (!ok) return;

      await deleteTaskFromUi(payload);
    });
  }
  
  const childAddButton = document.querySelector('[data-task-add-child]');
  if (childAddButton) {
    childAddButton.addEventListener('click', async () => {
      const payload = JSON.parse(childAddButton.dataset.taskAddChild);
      await openTaskAddModal(payload);
    });
  }
}

function bindTaskEditForm(originalTask, options = {}, parentOptions = []) {
  const form = document.getElementById('task-edit-form');
  if (!form) return;

  const cancelButton = document.querySelector('[data-task-edit-cancel]');
  if (cancelButton) {
    cancelButton.addEventListener('click', () => {
      openModal(renderTaskDetail(originalTask));
      bindTaskDetailActions(originalTask);
    });
  }

  form.addEventListener('submit', async (event) => {
    event.preventDefault();

    const formData = new FormData(form);

    const progress = String(formData.get('progress') || '').trim();
    if (progress !== '') {
      const n = Number(progress);
      if (!Number.isFinite(n) || n < 0 || n > 100) {
        alert('進捗(%)は0〜100の半角数字で入力してください。');
        return;
      }
    }

    const data = {
      no: formData.get('no'),
      taskName: formData.get('taskName'),
      parentTask: formData.get('parentTask'),
      assignee: formData.get('assignee'),
      dueDate: formData.get('dueDate'),
      targetDate: formData.get('targetDate'),
      startPlan: formData.get('startPlan'),
      status: formData.get('status'),
      progress,
      memo: formData.get('memo'),
    };

    try {
      const beforeTasks = [...(state.allTasksForWbs || [])];
      const isStructureChange =
        String(originalTask.parentTask || '') !== String(data.parentTask || '');

      closeModal();

      updateOptimisticTaskInState(data);
      rerenderTasksWithoutFetch();

      await apiGet('updateTask', {
        data: JSON.stringify(data),
      });

      if (isStructureChange) {
        state.allTasksForWbs = await apiGet('getTasks');
        state.taskAssigneeCache = {};
        rerenderTasksWithoutFetch();
      }
    } catch (err) {
      alert(err.message || 'タスクの保存に失敗しました。');

      if (typeof beforeTasks !== 'undefined') {
        state.allTasksForWbs = beforeTasks;
        rerenderTasksWithoutFetch();
      }
    }
  });
}

function bindTaskRefreshButton() {
  const button = document.getElementById('task-refresh-btn');
  if (!button) return;

  button.addEventListener('click', async () => {
    state.allTasksForWbs = [];
state.taskAssigneeCache = {};
await loadTasks();
  });
}

function bindTaskAddButton() {
  const addButton = document.getElementById('task-add-btn');
  if (!addButton) return;

  addButton.addEventListener('click', async () => {
    await openTaskAddModal();
  });
}

async function openTaskAddModal(parentTask = null) {
  try {
    setLoading(true);

    const [options, parentOptions] = await Promise.all([
      apiGet('getTaskEditOptions'),
      apiGet('getParentTaskOptions'),
    ]);

    openModal(renderTaskAddModal({
      options,
      parentOptions,
      parentTask,
    }));

    bindTaskAddModalEvents({
      options,
      parentOptions,
      parentTask,
    });
  } catch (err) {
    alert(err.message || 'タスク追加画面の表示に失敗しました。');
  } finally {
    setLoading(false);
  }
}

function renderTaskAddModal({ options = {}, parentOptions = [], parentTask = null }) {
  const forcedParentLabel = parentTask
    ? `${parentTask.no || ''}.${parentTask.taskName || parentTask.title || ''}`.trim()
    : '';

  const forcedParentValue = parentTask
  ? `${parentTask.no || ''}.${parentTask.taskName || parentTask.title || ''}`.trim()
  : '';

  return `
    <p class="eyebrow">Task Add</p>
    <h3>${parentTask ? '傘下にタスク追加' : '新規タスク追加'}</h3>

    <div class="task-add-tabs">
      <button class="task-add-tab active" type="button" data-task-add-mode="single">単発追加</button>
      <button class="task-add-tab" type="button" data-task-add-mode="bulk">複数追加</button>
    </div>

    <div id="task-add-single">
      ${renderTaskAddSingleForm({ options, parentOptions, parentTask, forcedParentLabel, forcedParentValue })}
    </div>

    <div id="task-add-bulk" class="hidden">
      ${renderTaskAddBulkForm({ parentOptions, parentTask, forcedParentLabel, forcedParentValue })}
    </div>
  `;
}

function renderParentSelect(parentOptions, selectedValue = '', disabled = false, name = 'parentTask') {
  return `
    <select name="${escapeHtml(name)}" ${disabled ? 'disabled' : ''}>
      <option value="">親タスク未定</option>
      ${(parentOptions || []).map(item => {
        const value = item.label || '';
        return `
          <option value="${escapeHtml(value)}" ${value === selectedValue ? 'selected' : ''}>
            ${escapeHtml(value)}
          </option>
        `;
      }).join('')}
    </select>
  `;
}

function normalizeParentTaskText(value) {
  return String(value || '')
    .trim()
    .replace(/\s+/g, '')
    .replace('．', '.');
}

function resolveSelectedParentTaskLabel(task, parentOptions = []) {
  const current = normalizeParentTaskText(task.parentTask || '');

  if (!current) return '';

  const matched = parentOptions.find(item => {
    const label = normalizeParentTaskText(item.label || '');
    const noTitle = normalizeParentTaskText(`${item.no || ''}.${item.title || ''}`);
    const titleOnly = normalizeParentTaskText(item.title || '');

    return (
      label === current ||
      noTitle === current ||
      titleOnly === current ||
      current.endsWith(titleOnly)
    );
  });

  return matched ? matched.label : task.parentTask || '';
}

function renderTaskAddSingleForm({ options = {}, parentOptions = [], parentTask = null, forcedParentValue = '' }) {
  const statusOptions = ['まだ💦', '順調！✨', '行き詰ってる…。', '完了！'];
  const assignees = options.assignees || [];
  const startPlans = options.startPlans || [];

  return `
    <form class="form-stack" id="task-add-single-form">
      <label>
        タイトル
        <input name="taskName" placeholder="タスク名を入力">
      </label>

      <label>
        親タスク
        ${renderParentSelect(parentOptions, forcedParentValue, !!parentTask)}
        ${parentTask ? `<input type="hidden" name="parentTask" value="${escapeHtml(forcedParentValue)}">` : ''}
      </label>

      <label>
        担当者
        <select name="assignee">
          <option value="">未設定</option>
          ${assignees.map(name => `
            <option value="${escapeHtml(name)}">${escapeHtml(name)}</option>
          `).join('')}
        </select>
      </label>

      <label>
        絶対！期日
        <input type="date" name="dueDate">
      </label>

      <label>
        目標期日
        <input type="date" name="targetDate">
      </label>

      <label>
        着手予定時期
        <select name="startPlan">
          <option value="">未設定</option>
          ${startPlans.map(plan => `
            <option value="${escapeHtml(plan)}">${escapeHtml(plan)}</option>
          `).join('')}
        </select>
      </label>

      <label>
        進捗状態
        <select name="status">
          <option value="">未設定</option>
          ${statusOptions.map(status => `
            <option value="${escapeHtml(status)}">${escapeHtml(status)}</option>
          `).join('')}
        </select>
      </label>

      <label>
        進捗(%)
        <input name="progress" inputmode="numeric" pattern="[0-9]*" placeholder="0〜100">
      </label>

      <label>
        進捗詳細・メモ
        <textarea name="memo"></textarea>
      </label>

      <div class="modal-actions">
        <button class="btn-primary" type="submit">追加する</button>
      </div>
    </form>
  `;
}

function renderTaskAddBulkForm({ parentOptions = [], parentTask = null, forcedParentValue = '' }) {
  const rows = Array.from({ length: 5 }, (_, index) => index);

  return `
    <form class="form-stack" id="task-add-bulk-form">
      <div id="task-bulk-rows" class="task-bulk-rows">
        ${rows.map(index => renderTaskBulkRow(parentOptions, forcedParentValue, !!parentTask, index)).join('')}
      </div>

      <button class="btn-secondary" type="button" id="task-bulk-add-row">
        ＋ 行を追加
      </button>

      <p class="meta">※タスク名が空の行は登録されません。</p>

      <div class="modal-actions">
        <button class="btn-primary" type="submit">まとめて追加する</button>
      </div>
    </form>
  `;
}

function renderTaskBulkRow(parentOptions = [], selectedParent = '', fixedParent = false, index = 0) {
  return `
    <div class="task-bulk-row" data-task-bulk-row>
      <input
        class="task-bulk-title"
        name="taskName_${index}"
        placeholder="タスクタイトル"
      >

      ${renderParentSelect(parentOptions, selectedParent, fixedParent, `parentTask_${index}`).replace('<select ', '<select class="task-bulk-parent" ')}

      ${fixedParent ? `<input class="task-bulk-parent-hidden" type="hidden" name="parentTask_${index}" value="${escapeHtml(selectedParent)}">` : ''}
    </div>
  `;
}

function bindTaskAddModalEvents(context) {
  document.querySelectorAll('[data-task-add-mode]').forEach(button => {
    button.addEventListener('click', () => {
      const mode = button.dataset.taskAddMode;

      document.querySelectorAll('[data-task-add-mode]').forEach(btn => {
        btn.classList.toggle('active', btn === button);
      });

      document.getElementById('task-add-single').classList.toggle('hidden', mode !== 'single');
      document.getElementById('task-add-bulk').classList.toggle('hidden', mode !== 'bulk');
    });
  });

  const singleForm = document.getElementById('task-add-single-form');
  if (singleForm) {
    singleForm.addEventListener('submit', async event => {
      event.preventDefault();

      const formData = new FormData(singleForm);
      const progress = String(formData.get('progress') || '').trim();

      if (progress !== '') {
        const n = Number(progress);
        if (!Number.isFinite(n) || n < 0 || n > 100) {
          alert('進捗(%)は0〜100の半角数字で入力してください。');
          return;
        }
      }

      const data = {
        taskName: formData.get('taskName'),
        parentTask: formData.get('parentTask'),
        assignee: formData.get('assignee'),
        dueDate: formData.get('dueDate'),
        targetDate: formData.get('targetDate'),
        startPlan: formData.get('startPlan'),
        status: formData.get('status'),
        progress,
        memo: formData.get('memo'),
      };

      if (!String(data.taskName || '').trim()) {
        alert('タイトルを入力してください。');
        return;
      }

      await submitTaskAdd('addTaskSingle', data);
    });
  }

const bulkForm = document.getElementById('task-add-bulk-form');
if (bulkForm) {
  const addRowButton = document.getElementById('task-bulk-add-row');
  const rowsWrap = document.getElementById('task-bulk-rows');

  if (addRowButton && rowsWrap) {
    addRowButton.addEventListener('click', () => {
      const index = rowsWrap.querySelectorAll('[data-task-bulk-row]').length;

      rowsWrap.insertAdjacentHTML(
        'beforeend',
        renderTaskBulkRow(
          context.parentOptions || [],
          context.parentTask
            ? `${context.parentTask.no || ''}.${context.parentTask.taskName || context.parentTask.title || ''}`.trim()
            : '',
          !!context.parentTask,
          index
        )
      );
    });
  }

  bulkForm.addEventListener('submit', async event => {
    event.preventDefault();

    const rows = [...bulkForm.querySelectorAll('[data-task-bulk-row]')].map(row => {
  const taskName = row.querySelector('.task-bulk-title')?.value || '';
  const parentSelect = row.querySelector('.task-bulk-parent');
  const parentHidden = row.querySelector('.task-bulk-parent-hidden');

  return {
    taskName: taskName.trim(),
    parentTask: String(parentHidden?.value || parentSelect?.value || '').trim(),
  };
}).filter(row => row.taskName);

    if (rows.length === 0) {
      alert('追加するタスク名を入力してください。');
      return;
    }

    await submitTaskAdd('addTaskBulk', { rows });
  });
}
}

function extractNoFromParentTaskLabel(parentTask) {
  const text = String(parentTask || '').trim();
  if (!text) return '';
  return String(text.split('.')[0] || '').trim();
}

function findTaskByNo(no) {
  const targetNo = String(no || '').trim();
  if (!targetNo) return null;

  return (state.allTasksForWbs || []).find(task => {
    return String(task.no || '').trim() === targetNo;
  }) || null;
}

function getParentIdFromParentTaskLabel(parentTask) {
  const parentNo = extractNoFromParentTaskLabel(parentTask);
  const parentTaskObj = findTaskByNo(parentNo);

  return parentTaskObj ? String(parentTaskObj.id || '').trim() : '';
}

function countDirectChildrenByParentId(parentId, ignoreTaskNo = '') {
  const targetParentId = String(parentId || '').trim();
  if (!targetParentId) return 0;

  const ignoreNo = String(ignoreTaskNo || '').trim();

  return (state.allTasksForWbs || []).filter(task => {
    const taskNo = String(task.no || '').trim();

    if (ignoreNo && taskNo === ignoreNo) return false;

    return getTaskParentId(task) === targetParentId;
  }).length;
}

function makeOptimisticId(parentTask, ignoreTaskNo = '') {
  const parentId = getParentIdFromParentTaskLabel(parentTask);

  // 親タスク未設定の場合は、IDを空にする
  // buildTaskTree側でID空白タスクは下に回る
  if (!parentId) return '';

  const childCount = countDirectChildrenByParentId(parentId, ignoreTaskNo);
  return `${parentId}-${childCount + 1}`;
}

function buildOptimisticTask(task, index = 0) {
  const optimisticId = makeOptimisticId(task.parentTask || '');

  return {
    no: '',
    taskName: task.taskName || '新規タスク',
    parentTask: task.parentTask || '',
    assignee: task.assignee || '未定',
    dueDate: task.dueDate || '',
    targetDate: task.targetDate || '',
    startPlan: task.startPlan || '',
    status: task.status || '',
    progress: task.progress || '',
    memo: task.memo || '',
    daysUntilDue: '',
    id: optimisticId,
    level: optimisticId ? optimisticId.split('-').length : 1,
    parentId: getTaskParentId({ id: optimisticId }),
    _optimistic: true,
    _optimisticIndex: index,
  };
}

function addOptimisticTaskToState(task) {
  if (!state.allTasksForWbs) state.allTasksForWbs = [];
  state.allTasksForWbs.push(buildOptimisticTask(task));
}

function addOptimisticTasksToState(tasks) {
  if (!state.allTasksForWbs) state.allTasksForWbs = [];

  tasks.forEach((task, index) => {
    state.allTasksForWbs.push(buildOptimisticTask(task, index));
  });
}

function updateOptimisticTaskInState(data) {
  if (!state.allTasksForWbs) state.allTasksForWbs = [];

  const no = String(data.no || '').trim();
  if (!no) return;

  const targetTask = state.allTasksForWbs.find(task => {
    return String(task.no || '').trim() === no;
  });

  if (!targetTask) return;

  const oldId = String(targetTask.id || '').trim();
  const parentTask = 'parentTask' in data ? data.parentTask || '' : targetTask.parentTask || '';
  const newId = makeOptimisticId(parentTask, no);

  state.allTasksForWbs = state.allTasksForWbs.map(task => {
    const taskNo = String(task.no || '').trim();
    const taskId = String(task.id || '').trim();

    // 変更対象タスク
    if (taskNo === no) {
      return {
        ...task,
        taskName: 'taskName' in data ? data.taskName || '' : task.taskName,
        parentTask,
        id: newId,
        level: newId ? newId.split('-').length : 1,
        parentId: getTaskParentId({ id: newId }),
        assignee: 'assignee' in data ? data.assignee || '' : task.assignee,
        dueDate: 'dueDate' in data ? data.dueDate || '' : task.dueDate,
        targetDate: 'targetDate' in data ? data.targetDate || '' : task.targetDate,
        startPlan: 'startPlan' in data ? data.startPlan || '' : task.startPlan,
        status: 'status' in data ? data.status || '' : task.status,
        progress: 'progress' in data ? data.progress || '' : task.progress,
        memo: 'memo' in data ? data.memo || '' : task.memo,
        _optimistic: true,
      };
    }

    // 変更対象の子孫タスクも、仮IDの枝を付け替える
    if (oldId && newId && taskId.startsWith(`${oldId}-`)) {
      const replacedId = taskId.replace(oldId, newId);

      return {
        ...task,
        id: replacedId,
        level: replacedId.split('-').length,
        parentId: getTaskParentId({ id: replacedId }),
        _optimistic: true,
      };
    }

    return task;
  });
}

function rerenderTasksWithoutFetch() {
  const tasks = state.allTasksForWbs || [];

  const safeTasks = addWbsAncestorsForAssigneeFilter(tasks);
  const filteredTasks = filterTasksByStatusKeepAncestors(
    safeTasks,
    state.taskStatusTab || 'incomplete'
  );
  const tree = buildTaskTree(filteredTasks);

  const activeTab =
    TASK_STATUS_TABS.find(tab => tab.key === (state.taskStatusTab || 'incomplete')) ||
    TASK_STATUS_TABS[0];

  const filterLabel = state.taskAssigneeFilter ? ` / ${state.taskAssigneeFilter}` : '';

  el.views.tasks.innerHTML = `
    <section class="card task-manage-card">
      <div class="section-head">
        <div>
          <p class="eyebrow">Task Manage</p>
          <h3>タスク関連</h3>
          <p class="meta">${escapeHtml(activeTab.label)}${escapeHtml(filterLabel)}：${filteredTasks.length}件 / 全${safeTasks.length}件</p>
        </div>
        <button id="task-refresh-btn" class="btn-secondary" type="button">
          最新状態に更新
        </button>
      </div>

      ${renderTaskAssigneeFilter()}
      ${renderTaskStatusTabs(safeTasks)}

      <div id="task-wbs-list" class="wbs-list">
        ${renderTaskTree(tree)}
      </div>

      <button id="task-add-btn" class="floating-add-btn" type="button">＋</button>
    </section>
  `;

  bindTaskAssigneeFilter();
  bindTaskStatusTabs();
  bindWbsToggles();
  bindRowModals();
  bindTaskAddButton();
  bindTaskRefreshButton();
}

async function deleteTaskFromUi(task) {
  const beforeTasks = [...(state.allTasksForWbs || [])];

  try {
    closeModal();

    const deletedNo = String(task.no || '').trim();
    const deletedId = String(task.id || '').trim();
    const newParentTask = String(task.parentTask || '').trim();

    state.allTasksForWbs = (state.allTasksForWbs || [])
      .filter(item => {
        return String(item.no || '').trim() !== deletedNo;
      })
      .map(item => {
        const itemId = String(item.id || '').trim();

        // 削除対象の直下の子だけ、親を付け替える
        if (deletedId && getTaskParentId(item) === deletedId) {
          const newId = makeOptimisticId(newParentTask, String(item.no || '').trim());

          return {
            ...item,
            parentTask: newParentTask,
            id: newId,
            level: newId ? newId.split('-').length : 1,
            parentId: getTaskParentId({ id: newId }),
            _optimistic: true,
          };
        }

        // 孫以下は、直下子のID変更に合わせて枝を付け替える
        if (deletedId && itemId.startsWith(`${deletedId}-`)) {
          const directChildOldId = itemId.split('-').slice(0, deletedId.split('-').length + 1).join('-');

          const directChild = beforeTasks.find(t => {
            return String(t.id || '').trim() === directChildOldId;
          });

          if (!directChild) return item;

          const directChildNo = String(directChild.no || '').trim();
          const directChildNew = state.allTasksForWbs?.find(t => {
            return String(t.no || '').trim() === directChildNo;
          });

          if (!directChildNew || !directChildNew.id) return item;

          const replacedId = itemId.replace(directChildOldId, directChildNew.id);

          return {
            ...item,
            id: replacedId,
            level: replacedId.split('-').length,
            parentId: getTaskParentId({ id: replacedId }),
            _optimistic: true,
          };
        }

        return item;
      });

    rerenderTasksWithoutFetch();

    await apiGet('deleteTask', {
      no: task.no,
    });

    state.allTasksForWbs = await apiGet('getTasks');
    state.taskAssigneeCache = {};
    rerenderTasksWithoutFetch();

  } catch (err) {
    alert(err.message || 'タスク削除に失敗しました。');

    state.allTasksForWbs = beforeTasks;
    rerenderTasksWithoutFetch();
  }
}

async function submitTaskAdd(action, data) {
  const beforeTasks = [...(state.allTasksForWbs || [])];

  try {
    closeModal();

    if (action === 'addTaskSingle') {
      addOptimisticTaskToState(data);
    }

    if (action === 'addTaskBulk') {
      addOptimisticTasksToState(data.rows || []);
    }

    rerenderTasksWithoutFetch();

    await apiGet(action, {
      data: JSON.stringify(data),
    });

    state.allTasksForWbs = await apiGet('getTasks');
    state.taskAssigneeCache = {};
    rerenderTasksWithoutFetch();

  } catch (err) {
    alert(err.message || 'タスク追加に失敗しました。');

    state.allTasksForWbs = beforeTasks;
    rerenderTasksWithoutFetch();
  }
}

function renderTaskEditForm(task, options = {}, parentOptions = []) {
  const no = task.no || '';
  const statusOptions = ['まだ💦', '順調！✨', '行き詰ってる…。', '完了！'];
  const assignees = options.assignees || [];
  const startPlans = options.startPlans || [];

  return `
    <form class="form-stack" id="task-edit-form">
      <input type="hidden" name="no" value="${escapeHtml(no)}">

      <label>
        タイトル
        <input name="taskName" value="${escapeHtml(task.taskName || '')}">
      </label>

      <label>
  親タスク
  ${renderParentSelect(parentOptions, resolveSelectedParentTaskLabel(task, parentOptions))}
</label>

      <label>
        担当者
        <select name="assignee">
          <option value="">未設定</option>
          ${assignees.map(name => `
            <option value="${escapeHtml(name)}" ${task.assignee === name ? 'selected' : ''}>
              ${escapeHtml(name)}
            </option>
          `).join('')}
        </select>
      </label>

      <label>
        絶対！期日
        <input type="date" name="dueDate" value="${escapeHtml(task.dueDate || '')}">
      </label>

      <label>
        目標期日
        <input type="date" name="targetDate" value="${escapeHtml(task.targetDate || '')}">
      </label>

      <label>
        着手予定時期
        <select name="startPlan">
          <option value="">未設定</option>
          ${startPlans.map(plan => `
            <option value="${escapeHtml(plan)}" ${task.startPlan === plan ? 'selected' : ''}>
              ${escapeHtml(plan)}
            </option>
          `).join('')}
        </select>
      </label>

      <label>
        進捗状態
        <select name="status">
          <option value="">未設定</option>
          ${statusOptions.map(status => `
            <option value="${escapeHtml(status)}" ${task.status === status ? 'selected' : ''}>
              ${escapeHtml(status)}
            </option>
          `).join('')}
        </select>
      </label>

      <label>
        進捗(%)
        <input name="progress" inputmode="numeric" pattern="[0-9]*" value="${escapeHtml(normalizeProgress(task.progress))}">
      </label>

      <label>
        進捗詳細・メモ
        <textarea name="memo">${escapeHtml(task.memo || '')}</textarea>
      </label>

      <div class="modal-actions">
        <button class="btn-secondary" type="button" data-task-edit-cancel='${escapeHtml(JSON.stringify(task))}'>
          戻る
        </button>
        <button class="btn-primary" type="submit">
          保存する
        </button>
      </div>
    </form>
  `;
}


function renderTopInfo(topInfo) {
  document.getElementById('home-today-label').textContent = '今日の日程';
  const daysText = topInfo?.daysUntilEvent || topInfo?.daysLeft || '-';
  document.getElementById('home-days-left').textContent = daysText === '-' ? '-' : `${daysText}日`;
  document.getElementById('home-current-phase').textContent = topInfo?.currentPhase || '現在フェイズ未設定';
  document.getElementById('home-today-schedule').innerHTML = renderTodaySchedule(topInfo?.todaySchedule || topInfo?.todayLabel);
}

async function loadView(view) {
  state.currentView = view;

  if (location.hash !== `#${view}`) {
    history.replaceState(null, '', `#${view}`);
  }
  el.pageTitle.textContent = VIEW_TITLES[view] || 'ホーム';

  el.navButtons.forEach((button) => {
    button.classList.toggle('active', button.dataset.view === view);
  });

  Object.entries(el.views).forEach(([key, section]) => {
    if (!section) return;
    section.classList.toggle('active', key === view);
  });

  closeSidebar();
  setError('');

  if (view === 'home') {
  await loadHome();
  return;
}

if (view === 'tasks') {
  await loadTasks();
  return;
}

if (view === 'questions') {
  await loadQuestions();
  return;
}

if (view === 'memos') {
  await loadMemos();
  return;
}

if (view === 'guests') {
  await loadGuests();
  return;
}

if (view === 'milestones') {
  await loadMilestones();
  return;
}

if (view === 'calendar') {
  await loadCalendar();
  return;
}

if (el.views[view]) {
  el.views[view].innerHTML = placeholder(view);
}
}

function getTaskParentId(task) {
  const id = String(task.id || '').trim();
  if (!id.includes('-')) return '';
  return id.split('-').slice(0, -1).join('-');
}

function addWbsAncestorsForAssigneeFilter(tasks) {
  const selectedUser = String(state.taskAssigneeFilter || '').trim();
  const selectedTasks = tasks || [];

  // 全員表示のときは何もしない
  if (!selectedUser) {
    return selectedTasks.map(task => ({
      ...task,
      __assigneeBaseMatched: true,
      __filterAncestorOnly: false,
    }));
  }

  const allTasks = state.allTasksForWbs || [];
  const selectedIds = new Set(
    selectedTasks
      .map(task => String(task.id || '').trim())
      .filter(Boolean)
  );

  const allById = new Map();
  allTasks.forEach(task => {
    const id = String(task.id || '').trim();
    if (id) allById.set(id, task);
  });

  const resultById = new Map();

  // GASから返ってきたタスクは、全部「本来表示対象」なので普通表示
  selectedTasks.forEach(task => {
    const id = String(task.id || '').trim();
    if (!id) return;

    resultById.set(id, {
      ...task,
      __assigneeBaseMatched: true,
      __filterAncestorOnly: false,
    });
  });

  // ID空白タスクも普通表示として残す
  const blankIdTasks = selectedTasks
    .filter(task => !String(task.id || '').trim())
    .map(task => ({
      ...task,
      __assigneeBaseMatched: true,
      __filterAncestorOnly: false,
    }));

  // selectedTasksの親を全体タスクからたどって補完する
  selectedTasks.forEach(task => {
    let parentId = getTaskParentId(task);

    while (parentId) {
      if (!resultById.has(parentId)) {
        const parentTask = allById.get(parentId);

        if (parentTask) {
          resultById.set(parentId, {
            ...parentTask,
            __assigneeBaseMatched: false,
            __filterAncestorOnly: true,
          });
        }
      }

      const currentParent = allById.get(parentId);
      parentId = currentParent ? getTaskParentId(currentParent) : '';
    }
  });

  const merged = [...resultById.values(), ...blankIdTasks];

  // 元の全体順に戻す。ID空白は最後
  const orderMap = new Map();
  allTasks.forEach((task, index) => {
    const id = String(task.id || '').trim();
    if (id) orderMap.set(id, index);
  });

  return merged.sort((a, b) => {
    const aId = String(a.id || '').trim();
    const bId = String(b.id || '').trim();

    if (!aId && !bId) return 0;
    if (!aId) return 1;
    if (!bId) return -1;

    return (orderMap.get(aId) ?? 999999) - (orderMap.get(bId) ?? 999999);
  });
}

function filterTasksByStatusKeepAncestors(tasks, tabKey) {
  const safeTasks = tasks || [];

  if (tabKey === 'all') {
    return safeTasks.map((task) => ({
      ...task,
      __filterAncestorOnly: task.__filterAncestorOnly === true,
    }));
  }

  const byId = new Map();
  safeTasks.forEach((task) => {
    const id = String(task.id || '').trim();
    if (id) byId.set(id, task);
  });

  const shouldShowTask = (task) => {
    // 便宜上追加した親タスクは、状態一致していても「通常表示対象」にしない
    if (task.__filterAncestorOnly === true) return false;

    if (tabKey === 'incomplete') return getTaskStatusKey(task) !== 'done';
    return getTaskStatusKey(task) === tabKey;
  };

  const includeIds = new Set();
  const matchedIds = new Set();

  safeTasks.forEach((task) => {
    const id = String(task.id || '').trim();

    if (shouldShowTask(task)) {
      if (id) {
        includeIds.add(id);
        matchedIds.add(id);
      }

      let parentId = getTaskParentId(task);
      while (parentId) {
        includeIds.add(parentId);
        const parentTask = byId.get(parentId);
        parentId = parentTask ? getTaskParentId(parentTask) : '';
      }
    }
  });

  return safeTasks
    .filter((task) => {
      const id = String(task.id || '').trim();

      // ID空白タスクは「親タスク未定」として残す
      if (!id) {
        if (task.__filterAncestorOnly === true) return false;
        if (tabKey === 'incomplete') return getTaskStatusKey(task) !== 'done';
        return shouldShowTask(task);
      }

      return includeIds.has(id);
    })
    .map((task) => {
      const id = String(task.id || '').trim();
      const alreadyAncestorOnly = task.__filterAncestorOnly === true;

      return {
        ...task,
        __filterAncestorOnly: alreadyAncestorOnly || Boolean(id && includeIds.has(id) && !matchedIds.has(id)),
      };
    });
}

function buildTaskTree(tasks) {
  const safeTasks = (tasks || [])
    .filter((task) => task && task.taskName)
    .map((task, index) => ({ ...task, id: String(task.id || '').trim(), level: Number(task.level || 1), children: [], originalIndex: index }));

  const byId = new Map();
  safeTasks.forEach((task) => { if (task.id) byId.set(task.id, task); });

  const roots = [];
  safeTasks.forEach((task) => {
    const parentId = getTaskParentId(task);
    const parent = parentId ? byId.get(parentId) : null;
    if (parent) parent.children.push(task);
    else roots.push(task);
  });

  const sortById = (a, b) => {
    const aParts = String(a.id || '').split('-').map((n) => Number(n));
    const bParts = String(b.id || '').split('-').map((n) => Number(n));
    const len = Math.max(aParts.length, bParts.length);
    for (let i = 0; i < len; i++) {
      const av = Number.isFinite(aParts[i]) ? aParts[i] : -1;
      const bv = Number.isFinite(bParts[i]) ? bParts[i] : -1;
      if (av !== bv) return av - bv;
    }
    return a.originalIndex - b.originalIndex;
  };

  const sortTree = (items) => {
    items.sort(sortById);
    items.forEach((item) => sortTree(item.children));
  };
  sortTree(roots);
    return roots.sort((a, b) => {
    const aNoId = !String(a.id || '').trim();
    const bNoId = !String(b.id || '').trim();

    if (aNoId && !bNoId) return 1;
    if (!aNoId && bNoId) return -1;

    return (a.originalIndex || 0) - (b.originalIndex || 0);
  });
}

function formatDaysUntilDue(value) {
  if (value === undefined || value === null || String(value).trim() === '') return '';
  const n = Number(value);
  if (Number.isNaN(n)) return '';
  if (n < 0) return `${Math.abs(n)}日超過`;
  if (n === 0) return '今日まで';
  return `あと${n}日`;
}

const TASK_STATUS_TABS = [
  { key: 'incomplete', label: '未完了' },
  { key: 'done', label: '完了！' },
  { key: 'todo', label: 'まだ💦' },
  { key: 'good', label: '順調！✨' },
  { key: 'stuck', label: '行き詰ってる…。' },
  { key: 'other', label: 'その他' },
  { key: 'all', label: '全タスク' },
];

function normalizeTaskStatus(status) {
  return String(status || '').trim();
}

function getTaskStatusKey(task) {
  const status = normalizeTaskStatus(task?.status)
    .replace(/\s/g, '')
    .replace(/。/g, '')
    .replace(/…/g, '');

  if (!status) return 'other';
  if (status.includes('完了') || status.includes('済')) return 'done';
  if (status.includes('まだ')) return 'todo';
  if (status.includes('順調')) return 'good';
  if (status.includes('行き詰')) return 'stuck';

  return 'other';
}

function filterTasksByStatus(tasks, tabKey) {
  const list = tasks || [];
  if (tabKey === 'all') return list;
  if (tabKey === 'incomplete') return list.filter((task) => getTaskStatusKey(task) !== 'done');
  return list.filter((task) => getTaskStatusKey(task) === tabKey);
}

function normalizeProgress(value) {
  const raw = String(value ?? '').replace('%', '').trim();
  const n = Number(raw);
  if (!Number.isFinite(n)) return 0;
  return Math.max(0, Math.min(100, n));
}

function renderProgressBar(progress) {
  const percent = normalizeProgress(progress);
  return `
    <div class="task-progress" aria-label="進捗 ${percent}%">
      <div class="task-progress-track">
        <div class="task-progress-fill" style="--progress:${percent}%"></div>
      </div>
      <span class="task-progress-label">${percent}%</span>
    </div>
  `;
}

function renderTaskAssigneeFilter() {
  const users = state.taskFilterUsers || [];

  return `
    <div class="task-assignee-filter">
      <label for="task-assignee-select">担当者で絞り込み</label>
      <select id="task-assignee-select">
        <option value="">全員のタスク</option>
        ${users.map(user => `
          <option value="${escapeHtml(user)}" ${state.taskAssigneeFilter === user ? 'selected' : ''}>
            ${escapeHtml(user)}
          </option>
        `).join('')}
      </select>
    </div>
  `;
}

function bindTaskAssigneeFilter() {
  const select = document.getElementById('task-assignee-select');
  if (!select) return;

  select.addEventListener('change', () => {
    state.taskAssigneeFilter = select.value || '';
    loadTasks();
  });
}

function renderTaskStatusTabs(tasks) {
  const counts = TASK_STATUS_TABS.reduce((acc, tab) => {
    acc[tab.key] = filterTasksByStatus(tasks, tab.key).length;
    return acc;
  }, {});

  return `
    <div class="task-status-tabs" role="tablist" aria-label="タスク状態フィルター">
      ${TASK_STATUS_TABS.map((tab) => `
        <button class="task-status-tab ${state.taskStatusTab === tab.key ? 'active' : ''}" type="button" data-task-tab="${escapeHtml(tab.key)}">
          <span>${escapeHtml(tab.label)}</span>
          <strong>${counts[tab.key] ?? 0}</strong>
        </button>
      `).join('')}
    </div>
  `;
}

function bindTaskStatusTabs() {
  document.querySelectorAll('[data-task-tab]').forEach((button) => {
    button.addEventListener('click', () => {
      state.taskStatusTab = button.dataset.taskTab || 'incomplete';
      loadTasks();
    });
  });
}

function getParentTaskLabel(task) {
  const parent = String(task?.parentTask || '').trim();
  if (parent) return parent;

  const parentId = getTaskParentId(task);
  if (!parentId) return '親タスク未定';

  return parentId;
}

function renderTaskCard(task) {
  const title = escapeHtml(task.taskName || task.title || '無題のタスク');
  const assignee = escapeHtml(task.assignee || '未定');
  const parentTask = escapeHtml(getParentTaskLabel(task));
  const dueDate = escapeHtml(task.dueDate || '期限なし');
  const targetDate = escapeHtml(task.targetDate || '未設定');
  const startPlan = escapeHtml(task.startPlan || '未設定');
  const daysUntilDue = escapeHtml(formatDaysUntilDue(task.daysUntilDue) || '未計算');
  const status = escapeHtml(task.status || '未設定');
  const progress = normalizeProgress(task.progress);
  const memo = String(task.memo || '').trim();
  const taskJson = escapeHtml(JSON.stringify(task));

  return `
    <button class="task-card" type="button" data-modal="task" data-payload='${taskJson}'>
      <div class="task-card-head">
        <div class="task-card-title-wrap">
          <span class="task-no">No.${escapeHtml(task.no || '-')}</span>
          <strong class="task-card-title">${title}</strong>
        </div>
        <span class="task-status-pill">${status}</span>
      </div>

      <div class="task-card-meta-grid">
        <div><dt>担当者</dt><dd>${assignee}</dd></div>
        <div><dt>親タスク</dt><dd>${parentTask}</dd></div>
        <div><dt>絶対！期日</dt><dd>${dueDate}</dd></div>
        <div><dt>目標期日</dt><dd>${targetDate}</dd></div>
        <div><dt>着手予定時期</dt><dd>${startPlan}</dd></div>
        <div><dt>作業日数残</dt><dd>${daysUntilDue}</dd></div>
      </div>

      <div class="task-card-progress-row">
        <span>進捗バー</span>
        ${renderProgressBar(progress)}
      </div>

      ${memo ? `<p class="task-card-memo">${escapeHtml(memo)}</p>` : ''}
    </button>
  `;
}

function renderTasksScreen(tasks) {
  const safeTasks = tasks || [];
  const filteredTasks = filterTasksByStatus(safeTasks, state.taskStatusTab);
  const activeTab = TASK_STATUS_TABS.find((tab) => tab.key === state.taskStatusTab) || TASK_STATUS_TABS[0];

  el.views.tasks.innerHTML = `
    <section class="card task-manage-card">
      <div class="section-head">
        <div>
          <p class="eyebrow">Task Manage</p>
          <h3>タスク関連</h3>
          <p class="meta">${escapeHtml(activeTab.label)}：${filteredTasks.length}件 / 全${safeTasks.length}件</p>
        </div>
      </div>

      ${renderTaskStatusTabs(safeTasks)}

      <div id="task-card-list" class="task-card-list">
        ${filteredTasks.length ? filteredTasks.map(renderTaskCard).join('') : '<p class="meta">この分類のタスクはありません。</p>'}
      </div>
    </section>
  `;

  bindTaskStatusTabs(safeTasks);
  bindRowModals();
}

function renderWbsProgress(task) {
  const progress = normalizeProgress(task.progress);
  const status = escapeHtml(task.status || '未設定');

  return `
    <div class="wbs-progress">
      <div class="wbs-progress-top">
        <span>${status}</span>
        <strong>${progress}%</strong>
      </div>
      <div class="task-progress-track">
        <div class="task-progress-fill" style="--progress:${progress}%"></div>
      </div>
    </div>
  `;
}

function renderTaskTree(nodes) {
  if (!nodes || nodes.length === 0) return '<p class="meta">タスクはありません。</p>';

  const renderNode = (task) => {
    const hasChildren = task.children && task.children.length > 0;
    const level = Math.max(1, Number(task.level || 1));
    const indent = Math.min((level - 1) * 18, 72);
    const taskJson = escapeHtml(JSON.stringify(task));
    const dueText = escapeHtml(formatDaysUntilDue(task.daysUntilDue));

    const isAncestorOnly = task.__filterAncestorOnly === true;

    if (isAncestorOnly) {
      return `
        <div class="wbs-node wbs-node-ancestor-only" data-task-id="${escapeHtml(task.id || '')}">
          <div class="wbs-row wbs-row-ancestor-only" style="--indent:${indent}px">
            <button class="wbs-toggle ${hasChildren ? '' : 'is-leaf'}" type="button" data-toggle-id="${escapeHtml(task.id || '')}" aria-label="子タスクを開閉" ${hasChildren ? '' : 'disabled'}>${hasChildren ? '▼' : '・'}</button>
            <div class="wbs-ancestor-title">
              ${escapeHtml(task.taskName || '無題のタスク')}
            </div>
          </div>
          ${hasChildren ? `<div class="wbs-children">${task.children.map(renderNode).join('')}</div>` : ''}
        </div>
      `;
    }

    return `
      <div class="wbs-node" data-task-id="${escapeHtml(task.id || '')}">
        <div class="wbs-row" style="--indent:${indent}px">
          <button class="wbs-toggle ${hasChildren ? '' : 'is-leaf'}" type="button" data-toggle-id="${escapeHtml(task.id || '')}" aria-label="子タスクを開閉" ${hasChildren ? '' : 'disabled'}>${hasChildren ? '▼' : '・'}</button>
          <button class="wbs-task-button" type="button" data-modal="task" data-payload='${taskJson}'>
            <span class="wbs-task-title">${escapeHtml(task.taskName || '無題のタスク')}</span>
            <span class="wbs-task-meta">${escapeHtml(task.assignee || '未定')}${task.dueDate ? ` / ${escapeHtml(task.dueDate)}` : ''}</span>
          </button>
          ${dueText ? `<span class="wbs-due-pill">${dueText}</span>` : ''}
        </div>
        ${renderWbsProgress(task)}
        ${hasChildren ? `<div class="wbs-children">${task.children.map(renderNode).join('')}</div>` : ''}
      </div>
    `;
  };

  return nodes.map(renderNode).join('');
}

function bindWbsToggles() {
  document.querySelectorAll('[data-toggle-id]').forEach((button) => {
    button.addEventListener('click', (event) => {
      event.stopPropagation();
      const node = button.closest('.wbs-node');
      if (!node) return;
      const children = node.querySelector(':scope > .wbs-children');
      if (!children) return;
      const isClosed = children.classList.toggle('closed');
      button.textContent = isClosed ? '▶' : '▼';
    });
  });
}

async function loadTasks() {
  setLoading(true);
  setError('');

  try {
    if (!state.taskFilterUsers || state.taskFilterUsers.length === 0) {
      state.taskFilterUsers = await apiGet('getTaskFilterUsers');
    }

    if (!state.allTasksForWbs || state.allTasksForWbs.length === 0) {
  state.allTasksForWbs = await apiGet('getTasks');
}

let tasks = state.allTasksForWbs;

if (state.taskAssigneeFilter) {
  const cacheKey = state.taskAssigneeFilter;

  if (!state.taskAssigneeCache[cacheKey]) {
    state.taskAssigneeCache[cacheKey] = await apiGet('getTasksByAssignee', { name: cacheKey });
  }

  tasks = state.taskAssigneeCache[cacheKey];
}

const safeTasks = addWbsAncestorsForAssigneeFilter(tasks || []);
const filteredTasks = filterTasksByStatusKeepAncestors(safeTasks, state.taskStatusTab || 'incomplete');
    const tree = buildTaskTree(filteredTasks);

    const activeTab = TASK_STATUS_TABS.find((tab) => tab.key === (state.taskStatusTab || 'incomplete')) || TASK_STATUS_TABS[0];
    const filterLabel = state.taskAssigneeFilter ? ` / ${state.taskAssigneeFilter}` : '';

    el.views.tasks.innerHTML = `
      <section class="card task-manage-card">
        <div class="section-head">
  <div>
    <p class="eyebrow">Task Manage</p>
    <h3>タスク関連</h3>
    <p class="meta">${escapeHtml(activeTab.label)}${escapeHtml(filterLabel)}：${filteredTasks.length}件 / 全${safeTasks.length}件</p>
  </div>
  <button id="task-refresh-btn" class="btn-secondary" type="button">
    最新状態に更新
  </button>
</div>

        ${renderTaskAssigneeFilter()}
        ${renderTaskStatusTabs(safeTasks)}

        <div id="task-wbs-list" class="wbs-list">
  ${renderTaskTree(tree)}
</div>

<button id="task-add-btn" class="floating-add-btn" type="button">＋</button>
      </section>
    `;

    bindTaskAssigneeFilter();
    bindTaskStatusTabs();
    bindWbsToggles();
    bindRowModals();
    bindHomeQuestionEvents();
    bindTaskAddButton();
    bindTaskRefreshButton();
  } catch (err) {
    console.error(err);
    setError(err.message || 'タスクの読み込みに失敗しました。');
  } finally {
    setLoading(false);
  }
}

async function loadHome() {
  setLoading(true);
  setError('');

  try {
    const topInfo = await apiGet('getHomeTopInfo');

const sheetDueDays = Number(topInfo?.dueDays);
const days = Number.isFinite(sheetDueDays) && sheetDueDays > 0
  ? sheetDueDays
  : (state.dueDays || window.APP_CONFIG.DEFAULT_DUE_DAYS || 7);

state.dueDays = days;

document.getElementById('due-days').value = days;
document.getElementById('due-task-title').textContent = `${days}日以内のタスク`;

const [summary, dueTasks, unresolvedQuestions, lineShare, questionOptions] = await Promise.all([
  apiGet('getHomeSummary', { days }),
  apiGet('getTasksDueWithinDays', { days }),
  apiGet('getUnresolvedQuestions'),
  apiGet('getLineShareText', { days }),
  apiGet('getQuestionOptions'),
]);

state.questionOptions = questionOptions || { owners: [], dues: [] };
state.homeUnresolvedQuestions = unresolvedQuestions || [];

    renderTopInfo(topInfo || {});

    document.getElementById('home-incomplete-count').textContent = summary?.incompleteTaskCount ?? 0;
    document.getElementById('home-near-due-count').textContent = summary?.nearDueTaskCount ?? 0;
    document.getElementById('home-question-count').textContent = summary?.unresolvedQuestionCount ?? 0;

    document.getElementById('due-task-list').innerHTML = renderTaskRows(dueTasks || []);
    document.getElementById('unresolved-question-list').innerHTML = renderQuestionRows(unresolvedQuestions || []);

    state.lineShareText = lineShare?.text || '';

    bindRowModals();
bindHomeQuestionEvents();
  } catch (err) {
    console.error(err);
    setError(err.message || 'ホーム画面の読み込みに失敗しました。');
  } finally {
    setLoading(false);
  }
}

function setupHomeEvents() {
  document.getElementById('home-refresh').addEventListener('click', () => loadView('home'));
  document.getElementById('home-apply-days').addEventListener('click', async () => {
  const nextDays = Number(document.getElementById('due-days').value || window.APP_CONFIG.DEFAULT_DUE_DAYS || 7);
  const safeDays = Math.max(1, nextDays);

  try {
    setLoading(true);
    setError('');

    await apiGet('setHomeDueDays', { days: safeDays });

    state.dueDays = safeDays;
    await loadView('home');
  } catch (err) {
    console.error(err);
    setError(err.message || '日数の保存に失敗しました。');
  } finally {
    setLoading(false);
  }
});
  document.getElementById('due-days').addEventListener('keydown', (event) => {
    if (event.key === 'Enter') document.getElementById('home-apply-days').click();
  });
  document.getElementById('copy-line').addEventListener('click', async () => {
    await navigator.clipboard.writeText(state.lineShareText || '');
    const result = document.getElementById('copy-result');
    result.classList.remove('hidden');
    setTimeout(() => result.classList.add('hidden'), 1800);
  });
}

el.sidebarOpen.addEventListener('click', openSidebar);
el.sidebarClose.addEventListener('click', closeSidebar);
el.sidebarBackdrop.addEventListener('click', closeSidebar);
el.navButtons.forEach((button) => {
  button.addEventListener('click', () => {
    loadView(button.dataset.view);
  });
});
el.modalClose.addEventListener('click', closeModal);
el.modal.addEventListener('click', (event) => {
  if (event.target === el.modal) closeModal();
});

state.questionStatus = 'unresolved';

async function loadQuestions(status = state.questionStatus || 'unresolved') {
  state.questionStatus = status;
  setLoading(true);
  setError('');

  try {
    if (!state.questionOptions || !(state.questionOptions.owners || []).length) {
      state.questionOptions = await apiGet('getQuestionOptions');
    }

    if (!state.allQuestions) {
      state.allQuestions = await apiGet('getQuestions', { status: 'all' });
    }

    const questions = filterQuestionsByStatus(state.allQuestions || [], status);

    el.views.questions.innerHTML = `
      <section class="card question-page">
        <div class="section-head">
          <div>
            <p class="eyebrow">Questions</p>
            <h3>疑問箱</h3>
          </div>
          <button id="question-refresh-btn" class="btn-secondary" type="button">
            最新状態に更新
          </button>
        </div>

        <div class="question-tabs">
          <button class="question-tab ${status === 'unresolved' ? 'active' : ''}" data-question-status="unresolved">未完了</button>
          <button class="question-tab ${status === 'resolved' ? 'active' : ''}" data-question-status="resolved">完了済</button>
          <button class="question-tab ${status === 'all' ? 'active' : ''}" data-question-status="all">すべて</button>
        </div>

        <div class="question-list">
          ${renderQuestionManageRows(questions || [])}
        </div>

        <button id="question-add-btn" class="floating-add-btn" type="button">＋</button>
      </section>
    `;

    bindQuestionEvents(questions || []);
    bindQuestionRefreshButton();
  } catch (err) {
    console.error(err);
    setError(err.message || '疑問箱の読み込みに失敗しました。');
  } finally {
    setLoading(false);
  }
}

function makeTempId(prefix) {
  return `${prefix}_${Date.now()}_${Math.random().toString(36).slice(2, 7)}`;
}

function rerenderQuestionsWithoutFetch() {
  const questions = filterQuestionsByStatus(state.allQuestions || [], state.questionStatus);

  const list = document.querySelector('.question-list');
  if (list) list.innerHTML = renderQuestionManageRows(questions);

  bindQuestionEvents(questions);
  bindQuestionRefreshButton();
}

function rerenderMemosWithoutFetch() {
  const list = document.querySelector('.memo-list');
  if (list) list.innerHTML = renderMemoRows(state.allMemos || []);

  bindMemoEvents(state.allMemos || []);
  bindMemoRefreshButton();
}

function rerenderGuestsWithoutFetch() {
  const list = document.querySelector('.guest-list');
  if (list) list.innerHTML = renderGuestRows(state.allGuests || []);

  bindGuestEvents(state.allGuests || []);
  bindGuestRefreshButton();
}

function rerenderMilestonesWithoutFetch() {
  if (!state.milestoneCache) return;

  const wrap = document.querySelector('.combined-gantt-wrap');
  if (!wrap) return;

  wrap.innerHTML = renderCombinedMilestoneGantt(
    state.milestoneCache.milestoneData,
    state.milestoneCache.ganttData
  );

  bindMilestoneGridEvents();
  bindMilestoneRefreshButton();
}

function rerenderCalendarWithoutFetch() {
  if (!state.calendarCache) return;

  const content = document.getElementById('calendar-content');
  if (!content) return;

  const activeMode = document.querySelector('[data-calendar-mode].active')?.dataset.calendarMode || 'grid';

  content.innerHTML = activeMode === 'list'
    ? renderCalendarList(state.calendarCache)
    : renderCalendarGrid(state.calendarCache);

  bindCalendarEditEvents();
  bindCalendarRefreshButton();

  if (activeMode === 'grid') {
    scrollCalendarToToday();
    bindCalendarSizeControl();
  }
}

function filterQuestionsByStatus(questions, status) {
  const list = questions || [];

  if (status === 'resolved') {
    return list.filter(q => q.resolved === true || String(q.resolved).toLowerCase() === 'true');
  }

  if (status === 'all') {
    return list;
  }

  return list.filter(q => q.resolved === false || String(q.resolved).toLowerCase() === 'false');
}

function renderQuestionManageRows(questions) {
  if (!questions.length) return '<p class="meta">該当する疑問はありません。</p>';

  return questions.map(q => `
    <div class="data-row question-manage-row">
      <div class="data-main">
        <strong>${escapeHtml(q.question || '無題の疑問')}</strong>
        <span class="meta">
          疑問ぬし：${escapeHtml(q.owner || '未定')}
          ${q.due ? ` / ${escapeHtml(q.due)}` : ''}
        </span>
        ${q.answer ? `<span class="meta">回答：${escapeHtml(q.answer)}</span>` : ''}
      </div>
      <div class="data-sub">
  <button class="question-edit-btn" type="button" data-question-id="${q.id}">✒</button>
  <button class="btn-danger mini-danger-btn question-delete-btn" type="button" data-question-id="${q.id}">削除</button>
</div>
    </div>
  `).join('');
}

function bindHomeQuestionEvents() {
  document.querySelectorAll('[data-home-question-detail]').forEach(row => {
    row.addEventListener('click', () => {
      const q = (state.homeUnresolvedQuestions || []).find(item => {
        return String(item.id) === String(row.dataset.homeQuestionDetail);
      });

      if (!q) return;
      openQuestionDetailModal(q);
    });
  });

  document.querySelectorAll('[data-home-question-edit]').forEach(btn => {
    btn.addEventListener('click', event => {
      event.stopPropagation();

      const q = (state.homeUnresolvedQuestions || []).find(item => {
        return String(item.id) === String(btn.dataset.homeQuestionEdit);
      });

      if (!q) return;
      openQuestionModal(q);
    });
  });
}

function openQuestionDetailModal(question) {
  openModal(`
    <p class="eyebrow">Question Detail</p>
    <h3>${escapeHtml(question.question || '無題の疑問')}</h3>

    <div class="detail-grid">
      <div>
        <span class="detail-label">疑問ぬし</span>
        <strong>${escapeHtml(question.owner || '未定')}</strong>
      </div>
      <div>
        <span class="detail-label">いつごろまでに</span>
        <strong>${escapeHtml(question.due || '未設定')}</strong>
      </div>
      <div>
        <span class="detail-label">状態</span>
        <strong>${question.resolved === true ? '完了済' : '未完了'}</strong>
      </div>
      <div>
        <span class="detail-label">回答</span>
        <p>${escapeHtml(question.answer || '未入力')}</p>
      </div>
    </div>

    <div class="modal-actions">
      <button class="btn-secondary" type="button" id="home-question-edit-from-detail">
        編集する
      </button>
    </div>
  `);

  document.getElementById('home-question-edit-from-detail')?.addEventListener('click', () => {
    openQuestionModal(question);
  });
}

function bindQuestionRefreshButton() {
  const button = document.getElementById('question-refresh-btn');
  if (!button) return;

  button.addEventListener('click', async () => {
    state.allQuestions = null;
    await loadQuestions(state.questionStatus);
  });
}

function bindQuestionEvents(questions) {
  document.querySelectorAll('[data-question-status]').forEach(btn => {
    btn.addEventListener('click', () => {
      loadQuestions(btn.dataset.questionStatus);
    });
  });

  document.getElementById('question-add-btn').addEventListener('click', () => {
    openQuestionModal();
  });

    document.querySelectorAll('.question-delete-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      const q = questions.find(item => String(item.id) === String(btn.dataset.questionId));
      if (q) deleteQuestionFromUi(q);
    });
  });
    btn.addEventListener('click', () => {
      const q = questions.find(item => String(item.id) === String(btn.dataset.questionId));
      openQuestionModal(q);
    });
  });
}

async function deleteQuestionFromUi(question) {
  const beforeQuestions = [...(state.allQuestions || [])];

  try {
    const ok = confirm(`「${question.question || 'この疑問'}」を削除しますか？`);
    if (!ok) return;

    closeModal();

    state.allQuestions = (state.allQuestions || []).filter(item => {
      return String(item.id) !== String(question.id);
    });

    rerenderQuestionsWithoutFetch();

    await apiGet('deleteQuestion', {
      id: question.id,
    });

  } catch (err) {
    alert(err.message || '疑問の削除に失敗しました。');

    state.allQuestions = beforeQuestions;
    rerenderQuestionsWithoutFetch();
  }
}

function openQuestionModal(question = null) {
  const isEdit = !!question;

  openModal(`
    <p class="eyebrow">${isEdit ? '疑問を編集' : '疑問を追加'}</p>
    <h3>${isEdit ? '疑問の編集' : '新しい疑問'}</h3>

    <div class="form-stack">
      <label>疑問内容
        <textarea id="question-form-question">${escapeHtml(question?.question || '')}</textarea>
      </label>

      <label>疑問ぬし
  <select id="question-form-owner">
    <option value="">選択してください</option>
    ${(state.questionOptions.owners || []).map(owner => `
      <option value="${escapeHtml(owner)}" ${owner === question?.owner ? 'selected' : ''}>
        ${escapeHtml(owner)}
      </option>
    `).join('')}
  </select>
</label>

<label>いつごろまでに
  <select id="question-form-due">
    <option value="">選択してください</option>
    ${(state.questionOptions.dues || []).map(due => `
      <option value="${escapeHtml(due)}" ${due === question?.due ? 'selected' : ''}>
        ${escapeHtml(due)}
      </option>
    `).join('')}
  </select>
</label>

      <label>回答
        <textarea id="question-form-answer">${escapeHtml(question?.answer || '')}</textarea>
      </label>

      <label class="check-row">
        <input id="question-form-resolved" type="checkbox" ${question?.resolved === true ? 'checked' : ''}>
        回答済みにする
      </label>

      <button id="question-save-btn" class="primary-btn" type="button">保存</button>
    </div>
  `);

  document.getElementById('question-save-btn').addEventListener('click', async () => {
    const data = {
      question: document.getElementById('question-form-question').value.trim(),
      owner: document.getElementById('question-form-owner').value.trim(),
      due: document.getElementById('question-form-due').value.trim(),
      answer: document.getElementById('question-form-answer').value.trim(),
      resolved: document.getElementById('question-form-resolved').checked,
    };

    if (!data.question) {
      setError('疑問内容を入力してください。');
      return;
    }

try {
      setError('');

      const beforeQuestions = [...(state.allQuestions || [])];

      closeModal();

      if (isEdit) {
        state.allQuestions = (state.allQuestions || []).map(item => {
          if (String(item.id) !== String(question.id)) return item;
          return { ...item, ...data };
        });

        rerenderQuestionsWithoutFetch();

        await apiGet('updateQuestion', {
          id: question.id,
          data: JSON.stringify(data),
        });
      } else {
        const tempId = makeTempId('question');

        state.allQuestions = [
          ...(state.allQuestions || []),
          {
            id: tempId,
            no: '',
            ...data,
            _optimistic: true,
          },
        ];

        rerenderQuestionsWithoutFetch();

        const result = await apiGet('addQuestion', {
          data: JSON.stringify(data),
        });

        if (result?.row) {
          state.allQuestions = (state.allQuestions || []).map(item => {
            if (item.id !== tempId) return item;
            return { ...item, id: result.row, _optimistic: false };
          });
          rerenderQuestionsWithoutFetch();
        }
      }

    } catch (err) {
      console.error(err);
      setError(err.message || '疑問の保存に失敗しました。');

      if (typeof beforeQuestions !== 'undefined') {
        state.allQuestions = beforeQuestions;
        rerenderQuestionsWithoutFetch();
      }
    }
  });
}


async function loadMemos() {
  setLoading(true);
  setError('');

  try {
    if (!state.allMemos) {
  state.allMemos = await apiGet('getMemos');
}

const memos = state.allMemos;

    el.views.memos.innerHTML = `
      <section class="card memo-page">
        <div class="section-head">
  <div>
    <p class="eyebrow">Memos</p>
    <h3>メモページ</h3>
    <p class="meta">DM文面・役割分担・締切メモをスマホで確認して、本文をそのままコピーできます。</p>
  </div>
  <button id="memo-refresh-btn" class="btn-secondary" type="button">
    最新状態に更新
  </button>
</div>

        <div class="memo-list">
          ${renderMemoRows(memos || [])}
        </div>

        <button id="memo-add-btn" class="floating-add-btn" type="button">＋</button>
      </section>
    `;

    bindMemoEvents(memos || []);
    bindMemoRefreshButton();
  } catch (err) {
    console.error(err);
    setError(err.message || 'メモページの読み込みに失敗しました。');
  } finally {
    setLoading(false);
  }
}

function renderMemoRows(memos) {
  if (!memos.length) return '<p class="meta">登録されているメモはありません。</p>';

  return memos.map(memo => {
    const preview = String(memo.body || '').replace(/\s+/g, ' ').slice(0, 70);

    return `
      <div class="data-row memo-row">
        <button class="memo-open-btn" type="button" data-memo-id="${escapeHtml(memo.id)}">
          <div class="data-main">
            <strong>${escapeHtml(memo.title || '無題のメモ')}</strong>
            ${preview ? `<span class="meta">${escapeHtml(preview)}${String(memo.body || '').length > 70 ? '…' : ''}</span>` : ''}
          </div>
          <div class="data-sub">
            <span>No.${escapeHtml(memo.no || '')}</span>
          </div>
        </button>
        <div class="data-sub">
  <button class="memo-edit-btn" type="button" data-memo-id="${escapeHtml(memo.id)}">✒</button>
  <button class="btn-danger mini-danger-btn memo-delete-btn" type="button" data-memo-id="${escapeHtml(memo.id)}">削除</button>
</div>
      </div>
    `;
  }).join('');
}

function bindMemoRefreshButton() {
  const button = document.getElementById('memo-refresh-btn');
  if (!button) return;

  button.addEventListener('click', async () => {
    state.allMemos = null;
    await loadMemos();
  });
}

function bindMemoEvents(memos) {
  const addBtn = document.getElementById('memo-add-btn');
  if (addBtn) {
    addBtn.addEventListener('click', () => openMemoEditModal());
  }

  document.querySelectorAll('.memo-open-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      const memo = memos.find(item => String(item.id) === String(btn.dataset.memoId));
      if (memo) openMemoDetailModal(memo);
    });
  });

  document.querySelectorAll('.memo-edit-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      const memo = memos.find(item => String(item.id) === String(btn.dataset.memoId));
      openMemoEditModal(memo);
    });
  });
    document.querySelectorAll('.memo-delete-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      const memo = memos.find(item => String(item.id) === String(btn.dataset.memoId));
      if (memo) deleteMemoFromUi(memo);
    });
  });
}

function openMemoDetailModal(memo) {
  openModal(`
    <p class="eyebrow">メモ詳細</p>
    <h3>${escapeHtml(memo.title || '無題のメモ')}</h3>
    <div class="memo-body">${escapeHtml(memo.body || '').replace(/\n/g, '<br>')}</div>
    <div class="modal-actions">
      <button id="memo-copy-btn" class="primary-btn" type="button">本文をコピー</button>
      <button id="memo-edit-from-detail-btn" class="btn-secondary" type="button">編集</button>
    </div>
    <p id="memo-copy-result" class="notice hidden">コピーしました。</p>
  `);

  document.getElementById('memo-copy-btn').addEventListener('click', async () => {
    await navigator.clipboard.writeText(memo.body || '');
    const result = document.getElementById('memo-copy-result');
    result.classList.remove('hidden');
    setTimeout(() => result.classList.add('hidden'), 1800);
  });

  document.getElementById('memo-edit-from-detail-btn').addEventListener('click', () => {
    openMemoEditModal(memo);
  });
}

async function deleteMemoFromUi(memo) {
  const beforeMemos = [...(state.allMemos || [])];

  try {
    const ok = confirm(`「${memo.title || 'このメモ'}」を削除しますか？`);
    if (!ok) return;

    closeModal();

    state.allMemos = (state.allMemos || []).filter(item => {
      return String(item.id) !== String(memo.id);
    });

    rerenderMemosWithoutFetch();

    await apiGet('deleteMemo', {
      id: memo.id,
    });

  } catch (err) {
    alert(err.message || 'メモの削除に失敗しました。');

    state.allMemos = beforeMemos;
    rerenderMemosWithoutFetch();
  }
}

function openMemoEditModal(memo = null) {
  const isEdit = !!memo;

  openModal(`
    <p class="eyebrow">${isEdit ? 'メモを編集' : 'メモを追加'}</p>
    <h3>${isEdit ? 'メモの編集' : '新しいメモ'}</h3>

    <div class="form-stack">
      <label>タイトル
        <input id="memo-form-title" type="text" value="${escapeHtml(memo?.title || '')}">
      </label>

      <label>本文
        <textarea id="memo-form-body" class="memo-form-body">${escapeHtml(memo?.body || '')}</textarea>
      </label>

      <button id="memo-save-btn" class="primary-btn" type="button">保存</button>
    </div>
  `);

  document.getElementById('memo-save-btn').addEventListener('click', async () => {
    const data = {
      title: document.getElementById('memo-form-title').value.trim(),
      body: document.getElementById('memo-form-body').value,
    };

    if (!data.title && !data.body.trim()) {
      setError('タイトルか本文のどちらかを入力してください。');
      return;
    }

try {
      setError('');

      const beforeMemos = [...(state.allMemos || [])];

      closeModal();

      if (isEdit) {
        state.allMemos = (state.allMemos || []).map(item => {
          if (String(item.id) !== String(memo.id)) return item;
          return { ...item, ...data };
        });

        rerenderMemosWithoutFetch();

        await apiGet('updateMemo', {
          id: memo.id,
          data: JSON.stringify(data),
        });
      } else {
        const tempId = makeTempId('memo');

        state.allMemos = [
          ...(state.allMemos || []),
          {
            id: tempId,
            no: '',
            ...data,
            _optimistic: true,
          },
        ];

        rerenderMemosWithoutFetch();

        const result = await apiGet('addMemo', {
          data: JSON.stringify(data),
        });

        if (result?.row) {
          state.allMemos = (state.allMemos || []).map(item => {
            if (item.id !== tempId) return item;
            return { ...item, id: result.row, _optimistic: false };
          });
          rerenderMemosWithoutFetch();
        }
      }

    } catch (err) {
      console.error(err);
      setError(err.message || 'メモの保存に失敗しました。');

      if (typeof beforeMemos !== 'undefined') {
        state.allMemos = beforeMemos;
        rerenderMemosWithoutFetch();
      }
    }
  });
}

async function loadMilestones() {
  setLoading(true);
  setError('');

  try {
    if (!state.milestoneCache) {
  const [milestoneData, ganttData] = await Promise.all([
    apiGet('getMilestoneGrid'),
    apiGet('getGanttGrid'),
  ]);

  state.milestoneCache = { milestoneData, ganttData };
}

const { milestoneData, ganttData } = state.milestoneCache;

    el.views.milestones.innerHTML = `
      <section class="card milestone-page">
        <div class="section-head">
  <div>
    <p class="eyebrow">Milestones</p>
    <h3>マイルストーン・ガントチャート</h3>
    <p class="meta">横スクロールで全体を確認できます。</p>
  </div>
  <button id="milestone-refresh-btn" class="btn-secondary" type="button">
    最新状態に更新
  </button>
</div>

        <div class="combined-gantt-wrap">
          ${renderCombinedMilestoneGantt(milestoneData, ganttData)}
        </div>
      </section>
    `;

    bindMilestoneGridEvents();
    bindMilestoneRefreshButton();
  } catch (err) {
    console.error(err);
    setError(err.message || 'マイルストーンの読み込みに失敗しました。');
  } finally {
    setLoading(false);
  }
}

function renderMilestoneRows(milestones) {
  if (!milestones.length) {
    return '<p class="meta">登録されているマイルストーンはありません。</p>';
  }

  return milestones.map(item => `
    <div class="data-row milestone-row">
      <button class="milestone-open-btn" type="button" data-milestone-id="${escapeHtml(item.id)}">
        <div class="data-main">
          <strong>${escapeHtml(item.title || '無題のマイルストーン')}</strong>
          <span class="meta">${escapeHtml(item.date || '日付未設定')}</span>
          ${item.detail ? `<span class="meta">${escapeHtml(String(item.detail).slice(0, 60))}${String(item.detail).length > 60 ? '…' : ''}</span>` : ''}
        </div>
        <div class="data-sub">
          <span>${escapeHtml(item.type || 'MILESTONE')}</span>
        </div>
      </button>
    </div>
  `).join('');
}

function bindMilestoneEvents(milestones) {
  document.querySelectorAll('.milestone-open-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      const item = milestones.find(m => String(m.id) === String(btn.dataset.milestoneId));
      if (item) openMilestoneDetailModal(item);
    });
  });
}

function openMilestoneDetailModal(item) {
  openModal(`
    <p class="eyebrow">マイルストーン詳細</p>
    <h3>${escapeHtml(item.title || '無題のマイルストーン')}</h3>

    <div class="detail-list">
      ${renderDetailRow('日付', item.date || '未設定')}
      ${renderDetailRow('種別', item.type || 'MILESTONE')}
      ${renderDetailRow('詳細', item.detail || '未入力')}
    </div>
  `);
}

async function loadGuests() {
  setLoading(true);
  setError('');

  try {
    if (!state.allGuests) {
  const [guests, guestOptions] = await Promise.all([
    apiGet('getGuests'),
    apiGet('getGuestOptions'),
  ]);

  state.allGuests = guests || [];
  state.guestOptions = guestOptions || { attackers: [], prospects: [] };
}

const guests = state.allGuests;

    el.views.guests.innerHTML = `
      <section class="card guest-page">
        <div class="section-head">
  <div>
    <p class="eyebrow">Guests</p>
    <h3>来る人リスト</h3>
    <p class="meta">来てほしい人・声かけ状況・見込みをスマホで確認できます。</p>
  </div>
  <button id="guest-refresh-btn" class="btn-secondary" type="button">
    最新状態に更新
  </button>
</div>

        <div class="guest-list">
          ${renderGuestRows(guests || [])}
        </div>

        <button id="guest-add-btn" class="floating-add-btn" type="button">＋</button>
      </section>
    `;

    bindGuestEvents(guests || []);
    bindGuestRefreshButton();
  } catch (err) {
    console.error(err);
    setError(err.message || '来る人リストの読み込みに失敗しました。');
  } finally {
    setLoading(false);
  }
}

function renderGuestRows(guests) {
  if (!guests.length) return '<p class="meta">登録されている人はいません。</p>';

  return guests.map(guest => `
    <div class="data-row guest-row">
      <div class="data-main">
        <strong>${escapeHtml(guest.name || '名前未設定')}</strong>
        <span class="meta">
          アタック担当：${escapeHtml(guest.attacker || '未定')}
          ${guest.prospect ? ` / 見込み：${escapeHtml(guest.prospect)}` : ''}
        </span>
        ${guest.status ? `<span class="meta">状況：${escapeHtml(guest.status)}</span>` : ''}
      </div>
      <div class="data-sub">
        <span>No.${escapeHtml(guest.no || '')}</span>
        <button class="guest-edit-btn" type="button" data-guest-id="${escapeHtml(guest.id)}">✒</button>
<button class="btn-danger mini-danger-btn guest-delete-btn" type="button" data-guest-id="${escapeHtml(guest.id)}">削除</button>
      </div>
    </div>
  `).join('');
}

function bindGuestRefreshButton() {
  const button = document.getElementById('guest-refresh-btn');
  if (!button) return;

  button.addEventListener('click', async () => {
    state.allGuests = null;
    await loadGuests();
  });
}

function bindGuestEvents(guests) {
  const addBtn = document.getElementById('guest-add-btn');
  if (addBtn) {
    addBtn.addEventListener('click', () => openGuestEditModal());
  }

  document.querySelectorAll('.guest-edit-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      const guest = guests.find(item => String(item.id) === String(btn.dataset.guestId));
      openGuestEditModal(guest);
    });
  });
    document.querySelectorAll('.guest-delete-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      const guest = guests.find(item => String(item.id) === String(btn.dataset.guestId));
      if (guest) deleteGuestFromUi(guest);
    });
  });
}

async function deleteGuestFromUi(guest) {
  const beforeGuests = [...(state.allGuests || [])];

  try {
    const ok = confirm(`「${guest.name || 'この人'}」を削除しますか？`);
    if (!ok) return;

    closeModal();

    state.allGuests = (state.allGuests || []).filter(item => {
      return String(item.id) !== String(guest.id);
    });

    rerenderGuestsWithoutFetch();

    await apiGet('deleteGuest', {
      id: guest.id,
    });

  } catch (err) {
    alert(err.message || '来る人リストの削除に失敗しました。');

    state.allGuests = beforeGuests;
    rerenderGuestsWithoutFetch();
  }
}

function openGuestEditModal(guest = null) {
  const isEdit = !!guest;

  openModal(`
    <p class="eyebrow">${isEdit ? '来る人を編集' : '来る人を追加'}</p>
    <h3>${isEdit ? '来る人リストの編集' : '新しく追加'}</h3>

    <div class="form-stack">
      <label>名前
        <input id="guest-form-name" type="text" value="${escapeHtml(guest?.name || '')}">
      </label>

      <label>アタック担当
  <select id="guest-form-attacker">
    <option value="">選択してください</option>
    ${(state.guestOptions?.attackers || []).map(attacker => `
      <option value="${escapeHtml(attacker)}" ${attacker === guest?.attacker ? 'selected' : ''}>
        ${escapeHtml(attacker)}
      </option>
    `).join('')}
  </select>
</label>

      <label>状況詳細
        <textarea id="guest-form-status">${escapeHtml(guest?.status || '')}</textarea>
      </label>

      <label>見込み
  <select id="guest-form-prospect">
    <option value="">選択してください</option>
    ${(state.guestOptions?.prospects || []).map(prospect => `
      <option value="${escapeHtml(prospect)}" ${prospect === guest?.prospect ? 'selected' : ''}>
        ${escapeHtml(prospect)}
      </option>
    `).join('')}
  </select>
</label>

      <button id="guest-save-btn" class="primary-btn" type="button">保存</button>
    </div>
  `);

  document.getElementById('guest-save-btn').addEventListener('click', async () => {
    const data = {
      name: document.getElementById('guest-form-name').value.trim(),
      attacker: document.getElementById('guest-form-attacker').value.trim(),
      status: document.getElementById('guest-form-status').value.trim(),
      prospect: document.getElementById('guest-form-prospect').value.trim(),
    };

    if (!data.name && !data.attacker && !data.status && !data.prospect) {
      setError('どれか1つは入力してください。');
      return;
    }

try {
      setError('');

      const beforeGuests = [...(state.allGuests || [])];

      closeModal();

      if (isEdit) {
        state.allGuests = (state.allGuests || []).map(item => {
          if (String(item.id) !== String(guest.id)) return item;
          return { ...item, ...data };
        });

        rerenderGuestsWithoutFetch();

        await apiGet('updateGuest', {
          id: guest.id,
          data: JSON.stringify(data),
        });
      } else {
        const tempId = makeTempId('guest');

        state.allGuests = [
          ...(state.allGuests || []),
          {
            id: tempId,
            no: '',
            ...data,
            _optimistic: true,
          },
        ];

        rerenderGuestsWithoutFetch();

        const result = await apiGet('addGuest', {
          data: JSON.stringify(data),
        });

        if (result?.row) {
          state.allGuests = (state.allGuests || []).map(item => {
            if (item.id !== tempId) return item;
            return { ...item, id: result.row, _optimistic: false };
          });
          rerenderGuestsWithoutFetch();
        }
      }

    } catch (err) {
      console.error(err);
      setError(err.message || '来る人リストの保存に失敗しました。');

      if (typeof beforeGuests !== 'undefined') {
        state.allGuests = beforeGuests;
        rerenderGuestsWithoutFetch();
      }
    }
  });
}

function renderMilestoneGrid(data) {
  const rows = data.rows || [];

  if (!rows.length) {
    return '<p class="meta">マイルストーンが登録されていません。</p>';
  }

  return `
    <table class="milestone-sheet">
      <tbody>
        ${rows.map((row, rowIndex) => `
          <tr class="milestone-sheet-row milestone-sheet-row-${rowIndex + 2}">
            ${row.map((cell, colIndex) => {
              const sheetRow = rowIndex + 2;
              const sheetCol = colIndex + 3;
              const editable = sheetRow === 3;

              return `
                <td
                  class="${editable ? 'milestone-editable-cell' : ''}"
                  data-row="${sheetRow}"
                  data-col="${sheetCol}"
                  data-value="${escapeHtml(cell || '')}"
                >
                  ${escapeHtml(cell || '')}
                </td>
              `;
            }).join('')}
          </tr>
        `).join('')}
      </tbody>
    </table>
  `;
}

function renderGanttGrid(data) {
  const headers = data.headers || [];
  const rows = data.rows || [];

  if (!headers.length || !rows.length) {
    return '<p class="meta">ガントチャートが登録されていません。</p>';
  }

  return `
    <table class="gantt-sheet">
      <thead>
        <tr>
          ${headers.map(header => `
            <th>${escapeHtml(header || '')}</th>
          `).join('')}
        </tr>
      </thead>
      <tbody>
        ${rows.map(row => renderGanttRow(row)).join('')}
      </tbody>
    </table>
  `;
}

function renderCombinedMilestoneGantt(milestoneData, ganttData) {
  const milestoneRows = milestoneData.rows || [];
  const ganttRows = ganttData.rows || [];

  if (!milestoneRows.length) {
    return '<p class="meta">マイルストーンが登録されていません。</p>';
  }

  return `
    <table class="combined-gantt-sheet">
      <tbody>
        ${milestoneRows.map((row, rowIndex) => `
          <tr class="combined-row combined-milestone-row combined-milestone-row-${rowIndex + 2}">
            ${row.map((cell, colIndex) => {
              const sheetRow = rowIndex + 2;
              const sheetCol = colIndex + 3;
              const editable = sheetRow === 3;

              return `
                <td
                  class="${editable ? 'milestone-editable-cell' : ''}"
                  data-row="${sheetRow}"
                  data-col="${sheetCol}"
                  data-value="${escapeHtml(cell || '')}"
                >
                  ${escapeHtml(cell || '')}
                </td>
              `;
            }).join('')}
          </tr>
        `).join('')}

        ${ganttRows.map(row => renderCombinedGanttRow(row)).join('')}
      </tbody>
    </table>
  `;
}

function renderCombinedGanttRow(row) {
  const cells = row.cells || [];
  let html = '';
  let i = 0;

  while (i < cells.length) {
    const isActive = String(cells[i]).trim() === '1';

    if (!isActive) {
      html += '<td class="combined-gantt-cell"></td>';
      i++;
      continue;
    }

    let span = 1;
    while (i + span < cells.length && String(cells[i + span]).trim() === '1') {
      span++;
    }

    html += `
      <td class="combined-gantt-cell gantt-active" colspan="${span}">
        <span class="gantt-task-label">${escapeHtml(row.taskName || '')}</span>
      </td>
    `;

    i += span;
  }

  return `<tr class="combined-gantt-row">${html}</tr>`;
}

function renderGanttRow(row) {
  const cells = row.cells || [];
  let html = '';
  let i = 0;

  while (i < cells.length) {
    const isActive = String(cells[i]).trim() === '1';

    if (!isActive) {
      html += '<td></td>';
      i++;
      continue;
    }

    let span = 1;
    while (i + span < cells.length && String(cells[i + span]).trim() === '1') {
      span++;
    }

    html += `
      <td class="gantt-active" colspan="${span}">
        <span class="gantt-task-label">${escapeHtml(row.taskName || '')}</span>
      </td>
    `;

    i += span;
  }

  return `<tr>${html}</tr>`;
}

function syncMilestoneAndGanttScroll() {
  const milestoneWrap = document.querySelector('.milestone-sheet-wrap');
  const ganttWrap = document.querySelector('.gantt-sheet-wrap');

  if (!milestoneWrap || !ganttWrap) return;

  let syncing = false;

  const sync = (from, to) => {
    if (syncing) return;

    syncing = true;
    to.scrollLeft = from.scrollLeft;

    requestAnimationFrame(() => {
      syncing = false;
    });
  };

  milestoneWrap.addEventListener('scroll', () => sync(milestoneWrap, ganttWrap));
  ganttWrap.addEventListener('scroll', () => sync(ganttWrap, milestoneWrap));
}

function bindMilestoneRefreshButton() {
  const button = document.getElementById('milestone-refresh-btn');
  if (!button) return;

  button.addEventListener('click', async () => {
    state.milestoneCache = null;
    await loadMilestones();
  });
}

function bindMilestoneGridEvents() {
  document.querySelectorAll('.milestone-editable-cell').forEach(cell => {
    cell.addEventListener('click', () => {
      openMilestoneCellEditModal({
        row: cell.dataset.row,
        col: cell.dataset.col,
        value: cell.dataset.value || '',
      });
    });
  });
}

function openMilestoneCellEditModal(cell) {
  openModal(`
    <p class="eyebrow">マイルストーン編集</p>
    <h3>${cell.row === '3' ? 'フェイズ内容' : 'タスク内容'}を編集</h3>

    <div class="form-stack">
      <label>内容
        <textarea id="milestone-cell-value">${escapeHtml(cell.value || '')}</textarea>
      </label>

      <button id="milestone-cell-save-btn" class="primary-btn" type="button">保存</button>
    </div>
  `);

  document.getElementById('milestone-cell-save-btn').addEventListener('click', async () => {
    const value = document.getElementById('milestone-cell-value').value.trim();

try {
      setError('');

      const beforeMilestoneCache = JSON.parse(JSON.stringify(state.milestoneCache || {}));

      closeModal();

      const sheetRow = Number(cell.row);
      const sheetCol = Number(cell.col);
      const rowIndex = sheetRow - 2;
      const colIndex = sheetCol - 3;

      if (state.milestoneCache?.milestoneData?.rows?.[rowIndex]) {
        state.milestoneCache.milestoneData.rows[rowIndex][colIndex] = value;
      }

      rerenderMilestonesWithoutFetch();

      await apiGet('updateMilestoneCell', {
        row: cell.row,
        col: cell.col,
        value,
      });

    } catch (err) {
      console.error(err);
      setError(err.message || 'マイルストーンの保存に失敗しました。');

      if (typeof beforeMilestoneCache !== 'undefined') {
        state.milestoneCache = beforeMilestoneCache;
        rerenderMilestonesWithoutFetch();
      }
    }
  });
}

async function loadCalendar() {
  setLoading(true);
  setError('');

  try {
    if (!state.calendarCache) {
  state.calendarCache = await apiGet('getCalendarGrid');
}

const data = state.calendarCache;

    el.views.calendar.innerHTML = `
      <section class="card calendar-page">
        <div class="section-head">
  <div>
    <p class="eyebrow">Calendar</p>
    <h3>カレンダー</h3>
    <p class="meta">カレンダー型と縦表示を切り替えて確認できます。</p>
  </div>
  <button id="calendar-refresh-btn" class="btn-secondary" type="button">
    最新状態に更新
  </button>
</div>

        <div class="view-switch">
          <button class="view-switch-btn active" type="button" data-calendar-mode="grid">カレンダー</button>
          <button class="view-switch-btn" type="button" data-calendar-mode="list">縦表示</button>
        </div>

        <div class="calendar-month-float" id="calendar-month-float">
  ${escapeHtml(getCalendarCurrentMonth(data, 'grid'))}
</div>

<div class="calendar-size-control">
  <span>表示サイズ</span>
  <input id="calendar-size-range" type="range" min="38" max="150" value="96">
</div>

<div id="calendar-content">
  ${renderCalendarGrid(data)}
</div>
      </section>
    `;

    bindCalendarSwitch(data);
    scrollCalendarToToday();
    bindCalendarSizeControl();
    bindCalendarEditEvents();
    bindCalendarRefreshButton();
  } catch (err) {
    console.error(err);
    setError(err.message || 'カレンダーの読み込みに失敗しました。');
  } finally {
    setLoading(false);
  }
}

function bindCalendarRefreshButton() {
  const button = document.getElementById('calendar-refresh-btn');
  if (!button) return;

  button.addEventListener('click', async () => {
    state.calendarCache = null;
    await loadCalendar();
  });
}

function bindCalendarSwitch(data) {
  document.querySelectorAll('[data-calendar-mode]').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('[data-calendar-mode]').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');

      const mode = btn.dataset.calendarMode;
      const content = document.getElementById('calendar-content');
      const monthFloat = document.getElementById('calendar-month-float');
      const sizeControl = document.querySelector('.calendar-size-control');

      content.innerHTML = mode === 'list'
        ? renderCalendarList(data)
        : renderCalendarGrid(data);
      
      bindCalendarEditEvents();

      if (mode === 'grid') {
  bindCalendarSizeControl();
  bindCalendarEditEvents();
}

      if (monthFloat) {
        monthFloat.textContent = getCalendarCurrentMonth(data, mode);
      }

      if (sizeControl) {
  sizeControl.style.display = mode === 'list' ? 'none' : 'flex';
}

      bindCalendarScrollMonth(data, mode);
      scrollCalendarToToday();
    });
  });

  bindCalendarScrollMonth(data, 'grid');
}

function renderCalendarGrid(data) {
  const weekdays = data.weekdays || [];
  const weeks = data.weeks || [];

  if (!weekdays.length || !weeks.length) {
    return '<p class="meta">カレンダーが登録されていません。</p>';
  }

  return `
    <div class="calendar-grid-wrap" id="calendar-scroll-area">
      <table class="calendar-grid-table">
        <thead>
          <tr>
            ${weekdays.map((day, index) => `
              <th class="${getCalendarDayClass(index)}">${escapeHtml(day || '')}</th>
            `).join('')}
          </tr>
        </thead>
        <tbody>
          ${weeks.map(week => `
            <tr class="calendar-date-row">
              ${week.dateRow.map((dateObj, index) => `
                <td
  class="${getCalendarDayClass(index)} ${dateObj?.isToday ? 'calendar-today' : ''} ${isPastCalendarDate(dateObj?.iso) ? 'calendar-past' : ''}"
  data-calendar-month="${escapeHtml(dateObj?.month || '')}"
  data-calendar-iso="${escapeHtml(dateObj?.iso || '')}"
>
                  <div class="calendar-cell-inner calendar-date-inner">
                    ${escapeHtml(dateObj?.text || '')}
                  </div>
                </td>
              `).join('')}
            </tr>
            <tr class="calendar-event-row">
              ${week.eventRow.map((eventObj, index) => {
                const dateObj = week.dateRow[index] || {};
                return `
                  <td
  class="${getCalendarDayClass(index)} ${dateObj?.isToday ? 'calendar-today' : ''} ${isPastCalendarDate(dateObj?.iso) ? 'calendar-past' : ''} calendar-editable-event"
  data-row="${escapeHtml(eventObj?.row || '')}"
  data-col="${escapeHtml(eventObj?.col || '')}"
  data-value="${escapeHtml(eventObj?.text || '')}"
>
                    <div class="calendar-cell-inner calendar-event-inner">
                      ${escapeHtml(eventObj?.text || '')}
                    </div>
                  </td>
                `;
              }).join('')}
            </tr>
          `).join('')}
        </tbody>
      </table>
    </div>
  `;
}

function renderCalendarList(data) {
  const events = (data.events || []).filter(item => item.date || item.text);

  if (!events.length) {
    return '<p class="meta">予定が登録されていません。</p>';
  }

  return `
    <div class="calendar-list" id="calendar-scroll-area">
      ${events.map(item => {
        const dayClass = getCalendarWeekdayClassByText(item.weekday);

        return `
          <div
  class="data-row calendar-list-row calendar-editable-event ${item.isToday ? 'calendar-today-list' : ''} ${isPastCalendarDate(item.iso) ? 'calendar-past-list' : ''}"
  data-calendar-month="${escapeHtml(item.month || extractMonthFromDateText(item.date))}"
  data-calendar-iso="${escapeHtml(item.iso || '')}"
  data-row="${escapeHtml(item.row || '')}"
  data-col="${escapeHtml(item.col || '')}"
  data-value="${escapeHtml(item.text || '')}"
>
            <div class="data-main">
              <strong class="${dayClass}">
                ${escapeHtml(item.date || '日付未設定')} ${escapeHtml(item.weekday || '')}
              </strong>
              <span class="meta">${escapeHtml(item.text || '予定なし')}</span>
            </div>
          </div>
        `;
      }).join('')}
    </div>
  `;
}

function getCalendarDayClass(index) {
  if (index === 0) return 'calendar-sunday';
  if (index === 6) return 'calendar-saturday';
  return '';
}

function getCalendarWeekdayClassByText(weekday) {
  const text = String(weekday || '');
  if (text.includes('日')) return 'calendar-sunday';
  if (text.includes('土')) return 'calendar-saturday';
  return '';
}

function extractMonthFromDateText(dateText) {
  const text = String(dateText || '').trim();

  const matchMonth = text.match(/(\d{1,2})月/);
  if (matchMonth) return `${Number(matchMonth[1])}月`;

  const matchSlash = text.match(/(\d{1,2})\/\d{1,2}/);
  if (matchSlash) return `${Number(matchSlash[1])}月`;

  return '';
}

function getCalendarCurrentMonth(data, mode) {
  if (mode === 'list') {
    const first = (data.events || []).find(item => item.month || item.date || item.text);
    return first?.month || extractMonthFromDateText(first?.date) || '月未設定';
  }

  for (const week of data.weeks || []) {
    for (const dateObj of week.dateRow || []) {
      if (dateObj?.month) return dateObj.month;
    }
  }

  return '月未設定';
}

function bindCalendarScrollMonth(data, mode) {
  const monthFloat = document.getElementById('calendar-month-float');
  const area = document.getElementById('calendar-scroll-area');

  if (!monthFloat || !area) return;

  area.addEventListener('scroll', () => {
    const month = findVisibleCalendarMonth(area);
    if (month) monthFloat.textContent = month;
  });
}

function findVisibleCalendarMonth(area) {
  const items = area.querySelectorAll('[data-calendar-month]');
  const areaRect = area.getBoundingClientRect();

  for (const item of items) {
    const month = item.dataset.calendarMonth;
    if (!month) continue;

    const rect = item.getBoundingClientRect();

    if (rect.bottom >= areaRect.top + 20 && rect.top <= areaRect.bottom) {
      return month;
    }
  }

  return '';
}

function scrollCalendarToToday() {
  const area = document.getElementById('calendar-scroll-area');
  if (!area) return;

  const todayCell = area.querySelector('.calendar-today, .calendar-today-list');
  if (!todayCell) return;

  setTimeout(() => {
    if (todayCell.classList.contains('calendar-today-list')) {
      todayCell.scrollIntoView({
        behavior: 'auto',
        block: 'start',
        inline: 'nearest',
      });

      area.scrollTop = Math.max(area.scrollTop - 8, 0);
      return;
    }

    area.scrollTop = Math.max(todayCell.offsetTop - 8, 0);
    area.scrollLeft = Math.max(todayCell.offsetLeft - 24, 0);
  }, 120);
}

function isPastCalendarDate(iso) {
  if (!iso) return false;

  const today = new Date();
  const todayIso = [
    today.getFullYear(),
    String(today.getMonth() + 1).padStart(2, '0'),
    String(today.getDate()).padStart(2, '0'),
  ].join('-');

  return iso < todayIso;
}

function bindCalendarSizeControl() {
  const range = document.getElementById('calendar-size-range');
  const table = document.querySelector('.calendar-grid-table');

  if (!range || !table) return;

  const applySize = () => {
    const size = Number(range.value || 96);
    table.style.setProperty('--calendar-cell-width', `${size}px`);
    table.style.setProperty('--calendar-event-height', `${Math.max(54, Math.round(size * 0.95))}px`);
  };

  range.addEventListener('input', applySize);
  applySize();
}

function bindCalendarEditEvents() {
  document.querySelectorAll('.calendar-editable-event').forEach(cell => {
    cell.addEventListener('click', () => {
      openCalendarEventEditModal({
        row: cell.dataset.row,
        col: cell.dataset.col,
        value: cell.dataset.value || '',
      });
    });
  });
}

function openCalendarEventEditModal(item) {
  openModal(`
    <p class="eyebrow">予定編集</p>
    <h3>カレンダー予定を編集</h3>

    <div class="form-stack">
      <label>予定内容
        <textarea id="calendar-event-value">${escapeHtml(item.value || '')}</textarea>
      </label>

      <button id="calendar-event-save-btn" class="primary-btn" type="button">保存</button>
    </div>
  `);

  document.getElementById('calendar-event-save-btn').addEventListener('click', async () => {
    const value = document.getElementById('calendar-event-value').value.trim();

try {
      setError('');

      const beforeCalendarCache = JSON.parse(JSON.stringify(state.calendarCache || {}));

      closeModal();

      const row = Number(item.row);
      const col = Number(item.col);

      (state.calendarCache?.weeks || []).forEach(week => {
        (week.eventRow || []).forEach(eventCell => {
          if (Number(eventCell.row) === row && Number(eventCell.col) === col) {
            eventCell.text = value;
          }
        });
      });

      (state.calendarCache?.events || []).forEach(eventItem => {
        if (Number(eventItem.row) === row && Number(eventItem.col) === col) {
          eventItem.text = value;
        }
      });

      rerenderCalendarWithoutFetch();

      await apiGet('updateCalendarEvent', {
        row: item.row,
        col: item.col,
        value,
      });

    } catch (err) {
      console.error(err);
      setError(err.message || '予定の保存に失敗しました。');

      if (typeof beforeCalendarCache !== 'undefined') {
        state.calendarCache = beforeCalendarCache;
        rerenderCalendarWithoutFetch();
      }
    }
  });
}

setupHomeEvents();

const initialView = location.hash.replace('#', '') || 'home';
loadView(VIEW_TITLES[initialView] ? initialView : 'home');

window.addEventListener('hashchange', () => {
  const view = location.hash.replace('#', '') || 'home';
  loadView(VIEW_TITLES[view] ? view : 'home');
});
