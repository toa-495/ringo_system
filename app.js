const state = {
  currentView: 'home',
  dueDays: window.APP_CONFIG.DEFAULT_DUE_DAYS,
};

const el = {
  loading: document.getElementById('loading'),
  error: document.getElementById('error'),
  views: {
    home: document.getElementById('view-home'),
    tasks: document.getElementById('view-tasks'),
    schedule: document.getElementById('view-schedule'),
    questions: document.getElementById('view-questions'),
    memos: document.getElementById('view-memos'),
    others: document.getElementById('view-others'),
  },
  navButtons: [...document.querySelectorAll('.nav-btn')],
  modal: document.getElementById('modal'),
  modalBody: document.getElementById('modal-body'),
  modalClose: document.getElementById('modal-close'),
};

async function apiGet(action, params = {}) {
  const url = new URL(window.APP_CONFIG.GAS_BASE_URL);
  url.searchParams.set('action', action);
  Object.entries(params).forEach(([k, v]) => url.searchParams.set(k, v));
  const res = await fetch(url.toString());
  const json = await res.json();
  if (!json.ok) throw new Error(json.error || 'APIエラー');
  return json.data;
}

function setLoading(show) { el.loading.classList.toggle('hidden', !show); }
function setError(msg = '') {
  el.error.textContent = msg;
  el.error.classList.toggle('hidden', !msg);
}

function openModal(data) {
  el.modalBody.innerHTML = `<pre>${escapeHtml(JSON.stringify(data, null, 2))}</pre>`;
  el.modal.classList.remove('hidden');
}
function closeModal() { el.modal.classList.add('hidden'); }
el.modalClose.addEventListener('click', closeModal);
el.modal.addEventListener('click', (e) => { if (e.target === el.modal) closeModal(); });

function escapeHtml(v) {
  return v.replace(/[&<>'"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', "'": '&#39;', '"': '&quot;' }[c]));
}

function card(title, body) { return `<article class="card"><h2>${title}</h2>${body}</article>`; }
function list(items, renderer) {
  if (!items?.length) return '<p class="meta">データがありません。</p>';
  return `<ul class="list">${items.map((it) => `<li class="list-item">${renderer(it)}</li>`).join('')}</ul>`;
}

async function renderHome() {
  const [summary, dueTasks, unresolved, share] = await Promise.all([
    apiGet('getHomeSummary'),
    apiGet('getTasksDueWithinDays', { days: state.dueDays }),
    apiGet('getUnresolvedQuestions'),
    apiGet('getLineShareText', { days: state.dueDays }),
  ]);

  el.views.home.innerHTML = [
    card('概要', `
      <div class="row"><div class="card"><div class="kpi">${summary.incompleteTaskCount ?? 0}</div><div class="meta">未完了タスク</div></div>
      <div class="card"><div class="kpi">${summary.nearDueTaskCount ?? 0}</div><div class="meta">直近タスク</div></div>
      <div class="card"><div class="kpi">${summary.unresolvedQuestionCount ?? 0}</div><div class="meta">未解決疑問</div></div></div>
      <div class="row"><input id="due-days" class="input" type="number" min="1" value="${state.dueDays}"><button id="reload-home">再取得</button><button id="copy-line" class="btn-primary">LINE共有文をコピー</button></div>
    `),
    card(`${state.dueDays}日以内のタスク`, list(dueTasks, (t) => `${t.taskName} <div class="meta">${t.assignee} / ${t.dueDate}</div>`)),
    card('未解決の疑問', list(unresolved, (q) => `${q.question} <div class="meta">${q.assignee || '未定'}</div>`)),
  ].join('');

  document.getElementById('reload-home').onclick = async () => {
    state.dueDays = Number(document.getElementById('due-days').value || 7);
    await loadView('home');
  };
  document.getElementById('copy-line').onclick = async () => {
    await navigator.clipboard.writeText(share.text || '');
    alert('コピーしました');
  };
}

async function renderTasks() {
  const [tasks, assignees, wbs] = await Promise.all([
    apiGet('getTasks'), apiGet('getAssignees'), apiGet('getWbsTree'),
  ]);
  el.views.tasks.innerHTML = [
    card('タスク一覧', list(tasks, (t) => `<button class="link-detail" data-json='${JSON.stringify(t)}'>${t.taskName}<div class="meta">${t.assignee} / ${t.dueDate}</div></button>`)),
    card('担当者別タスク', `<div class="row"><select id="assignee-filter">${assignees.map((a) => `<option>${a}</option>`).join('')}</select><button id="load-by-assignee">表示</button></div><div id="assignee-result" class="meta">担当者を選択してください</div>`),
    card('WBS', list(wbs, (p) => `<details><summary>${p.parentTask}</summary>${list(p.children || [], (c) => `${c.taskName} <div class="meta">${c.assignee} / ${c.dueDate}</div>`)}</details>`)),
  ].join('');

  document.querySelectorAll('.link-detail').forEach((btn) => btn.onclick = () => openModal(JSON.parse(btn.dataset.json)));
  document.getElementById('load-by-assignee').onclick = async () => {
    const name = document.getElementById('assignee-filter').value;
    const result = await apiGet('getTasksByAssignee', { name });
    document.getElementById('assignee-result').innerHTML = list(result, (t) => `${t.taskName}<div class="meta">${t.assignee} / ${t.dueDate}</div>`);
  };
}

async function renderSchedule() {
  const [miles, schedule] = await Promise.all([apiGet('getMilestones'), apiGet('getSchedule')]);
  const byMonth = schedule.reduce((acc, item) => {
    (acc[item.month] ||= []).push(item); return acc;
  }, {});
  el.views.schedule.innerHTML = [
    card('マイルストーン', list(miles, (m) => `<button class="link-detail" data-json='${JSON.stringify(m)}'>${m.date} ${m.title}</button>`)),
    card('カレンダー（縦型タイムライン）', Object.entries(byMonth).map(([m, items]) => `<h3>${m}</h3>${list(items, (i) => `<button class="link-detail" data-json='${JSON.stringify(i)}'>${i.date}<div>${i.title}</div></button>`)}`).join('')),
  ].join('');
  document.querySelectorAll('.link-detail').forEach((btn) => btn.onclick = () => openModal(JSON.parse(btn.dataset.json)));
}

async function renderQuestions() {
  const unresolved = await apiGet('getQuestions', { status: '未解決' });
  el.views.questions.innerHTML = card('疑問一覧', list(unresolved, (q) => `<button class="link-detail" data-json='${JSON.stringify(q)}'>${q.question}<div class="meta">${q.assignee || '未定'} / ${q.status}</div></button>`));
  document.querySelectorAll('.link-detail').forEach((btn) => btn.onclick = () => openModal(JSON.parse(btn.dataset.json)));
}

async function renderMemos() {
  const memos = await apiGet('getMemos');
  el.views.memos.innerHTML = card('メモ一覧', list(memos, (m) => `<button class="link-detail" data-json='${JSON.stringify(m)}'>${m.title}</button>`));
  document.querySelectorAll('.link-detail').forEach((btn) => btn.onclick = () => openModal(JSON.parse(btn.dataset.json)));
}

async function renderOthers() {
  const guests = await apiGet('getGuests');
  el.views.others.innerHTML = card('来る人リスト', list(guests, (g) => `<button class="link-detail" data-json='${JSON.stringify(g)}'>${g.name}<div class="meta">${g.status} / ${g.contactStatus || ''}</div></button>`));
  document.querySelectorAll('.link-detail').forEach((btn) => btn.onclick = () => openModal(JSON.parse(btn.dataset.json)));
}

async function loadView(view) {
  setError(''); setLoading(true);
  try {
    Object.keys(el.views).forEach((v) => el.views[v].classList.toggle('active', v === view));
    el.navButtons.forEach((b) => b.classList.toggle('active', b.dataset.view === view));
    if (view === 'home') await renderHome();
    if (view === 'tasks') await renderTasks();
    if (view === 'schedule') await renderSchedule();
    if (view === 'questions') await renderQuestions();
    if (view === 'memos') await renderMemos();
    if (view === 'others') await renderOthers();
  } catch (e) {
    setError(e.message);
  } finally { setLoading(false); }
}

el.navButtons.forEach((btn) => btn.addEventListener('click', () => loadView(btn.dataset.view)));
loadView('home');
