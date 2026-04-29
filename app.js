const state = {
  currentView: 'home',
  dueDays: window.APP_CONFIG.DEFAULT_DUE_DAYS || 7,
  lineShareText: '',
  questionStatus: 'unresolved',
  questionOptions: { owners: [], dues: [] },
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

function setLoading(show) {
  el.loading.classList.toggle('hidden', !show);
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

  return questions.map((question) => {
    const text = escapeHtml(question.question || question.content || '無題の疑問');
    const owner = escapeHtml(question.owner || question.assignee || question.questioner || '未定');
    const due = escapeHtml(question.due || question.priority || question.deadline || '期限未設定');

    return `
      <div class="data-row question-row static-row">
        <div class="data-main">
          <strong>${text}</strong>
          <span class="meta">疑問ぬし：${owner}</span>
        </div>
        <div class="data-sub">
          <span>${due}</span>
        </div>
      </div>
    `;
  }).join('');
}


function renderTaskDetail(payload) {
  const taskTitle = payload.title || payload.taskName || '詳細';
  const items = [
    ['No.', payload.no],
    ['タイトル', payload.title || payload.taskName],
    ['親タスク', payload.parentTask],
    ['担当者', payload.assignee || '未定'],
    ['絶対！期日', payload.dueDate || '期限なし'],
    ['目標期日', payload.targetDate],
    ['着手予定時期', payload.startPlan],
    ['進捗状態', payload.status],
    ['進捗％', payload.progress],
    ['進捗詳細・メモ', payload.memo],
  ].filter(([, value]) => value !== undefined && value !== null && value !== '');

  return `
    <p class="eyebrow">タスク詳細</p>
    <h3>${escapeHtml(taskTitle)}</h3>
    <dl class="detail-list">
      ${items.map(([key, value]) => `
        <div>
          <dt>${escapeHtml(key)}</dt>
          <dd>${escapeHtml(value)}</dd>
        </div>
      `).join('')}
    </dl>
  `;
}


function bindRowModals() {
  document.querySelectorAll('[data-modal="task"][data-payload]').forEach((button) => {
    button.addEventListener('click', () => {
      const payload = JSON.parse(button.dataset.payload);
      openModal(renderTaskDetail(payload));
    });
  });
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

if (el.views[view]) {
  el.views[view].innerHTML = placeholder(view);
}
}

function getTaskParentId(task) {
  const id = String(task.id || '').trim();
  if (!id.includes('-')) return '';
  return id.split('-').slice(0, -1).join('-');
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
  return roots;
}

function formatDaysUntilDue(value) {
  if (value === undefined || value === null || String(value).trim() === '') return '';
  const n = Number(value);
  if (Number.isNaN(n)) return '';
  if (n < 0) return `${Math.abs(n)}日超過`;
  if (n === 0) return '今日まで';
  return `あと${n}日`;
}

function renderTaskTree(nodes) {
  if (!nodes || nodes.length === 0) return '<p class="meta">タスクはありません。</p>';

  const renderNode = (task) => {
    const hasChildren = task.children && task.children.length > 0;
    const level = Math.max(1, Number(task.level || 1));
    const indent = Math.min((level - 1) * 18, 72);
    const taskJson = escapeHtml(JSON.stringify(task));
    const dueText = escapeHtml(formatDaysUntilDue(task.daysUntilDue));

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
    const tasks = await apiGet('getTasks');
    const tree = buildTaskTree(tasks || []);
    el.views.tasks.innerHTML = `
      <section class="card">
        <div class="section-head">
          <div>
            <p class="eyebrow">WBS Tasks</p>
            <h3>タスク関連</h3>
            <p class="meta">親タスクの横のタブで、子タスクを開閉できます。</p>
          </div>
        </div>
        <div id="task-wbs-list" class="wbs-list">${renderTaskTree(tree)}</div>
      </section>
    `;
    bindWbsToggles();
    bindRowModals();
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

const [summary, dueTasks, unresolvedQuestions, lineShare] = await Promise.all([
  apiGet('getHomeSummary', { days }),
  apiGet('getTasksDueWithinDays', { days }),
  apiGet('getUnresolvedQuestions'),
  apiGet('getLineShareText', { days }),
]);

    renderTopInfo(topInfo || {});

    document.getElementById('home-incomplete-count').textContent = summary?.incompleteTaskCount ?? 0;
    document.getElementById('home-near-due-count').textContent = summary?.nearDueTaskCount ?? 0;
    document.getElementById('home-question-count').textContent = summary?.unresolvedQuestionCount ?? 0;

    document.getElementById('due-task-list').innerHTML = renderTaskRows(dueTasks || []);
    document.getElementById('unresolved-question-list').innerHTML = renderQuestionRows(unresolvedQuestions || []);

    state.lineShareText = lineShare?.text || '';

    bindRowModals();
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
el.navButtons.forEach((button) => button.addEventListener('click', () => loadView(button.dataset.view)));
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
    state.questionOptions = await apiGet('getQuestionOptions');
const questions = await apiGet('getQuestions', { status });

    el.views.questions.innerHTML = `
      <section class="card question-page">
        <div class="section-head">
          <div>
            <p class="eyebrow">Questions</p>
            <h3>疑問箱</h3>
          </div>
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
  } catch (err) {
    console.error(err);
    setError(err.message || '疑問箱の読み込みに失敗しました。');
  } finally {
    setLoading(false);
  }
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
      <button class="question-edit-btn" type="button" data-question-id="${q.id}">✒</button>
    </div>
  `).join('');
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

  document.querySelectorAll('.question-edit-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      const q = questions.find(item => String(item.id) === String(btn.dataset.questionId));
      openQuestionModal(q);
    });
  });
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
      setLoading(true);
      setError('');

      if (isEdit) {
        await apiGet('updateQuestion', {
          id: question.id,
          data: JSON.stringify(data),
        });
      } else {
        await apiGet('addQuestion', {
          data: JSON.stringify(data),
        });
      }

      closeModal();
      await loadQuestions(state.questionStatus);
    } catch (err) {
      console.error(err);
      setError(err.message || '疑問の保存に失敗しました。');
    } finally {
      setLoading(false);
    }
  });
}


async function loadMemos() {
  setLoading(true);
  setError('');

  try {
    const memos = await apiGet('getMemos');

    el.views.memos.innerHTML = `
      <section class="card memo-page">
        <div class="section-head">
          <div>
            <p class="eyebrow">Memos</p>
            <h3>メモページ</h3>
            <p class="meta">DM文面・役割分担・締切メモをスマホで確認して、本文をそのままコピーできます。</p>
          </div>
        </div>

        <div class="memo-list">
          ${renderMemoRows(memos || [])}
        </div>

        <button id="memo-add-btn" class="floating-add-btn" type="button">＋</button>
      </section>
    `;

    bindMemoEvents(memos || []);
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
        <button class="memo-edit-btn" type="button" data-memo-id="${escapeHtml(memo.id)}">✒</button>
      </div>
    `;
  }).join('');
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
      setLoading(true);
      setError('');

      if (isEdit) {
        await apiGet('updateMemo', {
          id: memo.id,
          data: JSON.stringify(data),
        });
      } else {
        await apiGet('addMemo', {
          data: JSON.stringify(data),
        });
      }

      closeModal();
      await loadMemos();
    } catch (err) {
      console.error(err);
      setError(err.message || 'メモの保存に失敗しました。');
    } finally {
      setLoading(false);
    }
  });
}

async function loadMilestones() {
  setLoading(true);
  setError('');

  try {
    const [milestoneData, ganttData] = await Promise.all([
      apiGet('getMilestoneGrid'),
      apiGet('getGanttGrid'),
    ]);

    el.views.milestones.innerHTML = `
      <section class="card milestone-page">
        <div class="section-head">
          <div>
            <p class="eyebrow">Milestones</p>
            <h3>マイルストーン・ガント</h3>
            <p class="meta">横スクロールで全体を確認できます。</p>
          </div>
        </div>

        <div class="combined-gantt-wrap">
          ${renderCombinedMilestoneGantt(milestoneData, ganttData)}
        </div>
      </section>
    `;

    bindMilestoneGridEvents();
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
    const [guests, guestOptions] = await Promise.all([
  apiGet('getGuests'),
  apiGet('getGuestOptions'),
]);

state.guestOptions = guestOptions || { attackers: [], prospects: [] };

    el.views.guests.innerHTML = `
      <section class="card guest-page">
        <div class="section-head">
          <div>
            <p class="eyebrow">Guests</p>
            <h3>来る人リスト</h3>
            <p class="meta">来てほしい人・声かけ状況・見込みをスマホで確認できます。</p>
          </div>
        </div>

        <div class="guest-list">
          ${renderGuestRows(guests || [])}
        </div>

        <button id="guest-add-btn" class="floating-add-btn" type="button">＋</button>
      </section>
    `;

    bindGuestEvents(guests || []);
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
      </div>
    </div>
  `).join('');
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
      setLoading(true);
      setError('');

      if (isEdit) {
        await apiGet('updateGuest', {
          id: guest.id,
          data: JSON.stringify(data),
        });
      } else {
        await apiGet('addGuest', {
          data: JSON.stringify(data),
        });
      }

      closeModal();
      await loadGuests();
    } catch (err) {
      console.error(err);
      setError(err.message || '来る人リストの保存に失敗しました。');
    } finally {
      setLoading(false);
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
              const editable = sheetRow === 3 || sheetRow === 4;

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
              const editable = sheetRow === 3 || sheetRow === 4;

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
      setLoading(true);
      setError('');

      await apiGet('updateMilestoneCell', {
        row: cell.row,
        col: cell.col,
        value,
      });

      closeModal();
      await loadMilestones();
    } catch (err) {
      console.error(err);
      setError(err.message || 'マイルストーンの保存に失敗しました。');
    } finally {
      setLoading(false);
    }
  });
}

setupHomeEvents();
loadView('home');
