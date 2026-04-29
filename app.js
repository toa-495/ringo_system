const state = {
  currentView: 'home',
  dueDays: window.APP_CONFIG.DEFAULT_DUE_DAYS || 7,
  lineShareText: '',
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
    const title = escapeHtml(task.taskName || task.title || '無題のタスク');
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
  const items = [
    ['タスクタイトル', payload.taskName],
    ['ID', payload.taskId],
    ['固有No.', payload.no],
    ['親タスクID', payload.parentId],
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
    <h3>${escapeHtml(payload.taskName || '詳細')}</h3>
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
  document.getElementById('home-days-left').textContent = topInfo?.daysUntilEvent ?? '-';
  document.getElementById('home-current-phase').textContent = topInfo?.currentPhase || '現在フェイズ未設定';
  document.getElementById('home-today-schedule').innerHTML = renderTodaySchedule(topInfo?.todaySchedule || topInfo?.todayLabel);
}

async function renderHome() {
  const days = state.dueDays;
  const [topInfo, summary, dueTasks, unresolved, share] = await Promise.all([
    apiGet('getHomeTopInfo'),
    apiGet('getHomeSummary', { days }),
    apiGet('getTasksDueWithinDays', { days }),
    apiGet('getUnresolvedQuestions'),
    apiGet('getLineShareText'),
  ]);

  state.lineShareText = share?.text || '';

  renderTopInfo(topInfo);
  document.getElementById('home-incomplete-count').textContent = summary?.incompleteTaskCount ?? 0;
  document.getElementById('home-near-due-count').textContent = summary?.nearDueTaskCount ?? dueTasks?.length ?? 0;
  document.getElementById('home-question-count').textContent = summary?.unresolvedQuestionCount ?? unresolved?.length ?? 0;
  document.getElementById('due-days').value = days;
  document.getElementById('due-task-title').textContent = `${days}日以内のタスク`;
  document.getElementById('due-task-list').innerHTML = renderTaskRows(dueTasks);
  document.getElementById('unresolved-question-list').innerHTML = renderQuestionRows(unresolved);

  bindRowModals();
}

async function loadView(view) {
  state.currentView = view;
  setError('');
  setLoading(true);

  try {
    Object.entries(el.views).forEach(([key, node]) => {
      node.classList.toggle('active', key === view);
      if (key !== 'home' && key === view && !node.innerHTML.trim()) {
        node.innerHTML = placeholder(key);
      }
    });

    el.navButtons.forEach((button) => button.classList.toggle('active', button.dataset.view === view));
    el.pageTitle.textContent = VIEW_TITLES[view] || 'ホーム';

    if (view === 'home') await renderHome();
  } catch (error) {
    setError(error.message);
  } finally {
    setLoading(false);
    closeSidebar();
  }
}

function setupHomeEvents() {
  document.getElementById('home-refresh').addEventListener('click', () => loadView('home'));
  document.getElementById('home-apply-days').addEventListener('click', () => {
    const nextDays = Number(document.getElementById('due-days').value || window.APP_CONFIG.DEFAULT_DUE_DAYS || 7);
    state.dueDays = Math.max(1, nextDays);
    loadView('home');
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

setupHomeEvents();
loadView('home');
