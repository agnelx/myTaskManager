// ─── STATE ───────────────────────────────────────────────────────────
const todayDate = new Date();
const todayStr = todayDate.toISOString().slice(0, 10);

const SK = 'mytm_clickup_v1';
const API_BASE = window.location.origin;

let state = {
  tasks: [],
  meetings: [],
  reminders: [],
  aiHistory: [],
  outlookClientId: '',
  outlookDays: 14,
  outlookMeetings: [],
  outlookUser: '',
  anthropicKey: ''
};

let currentView = 'dashboard';
let currentFilters = { work: false, personal: false, high: false, overdue: false };
let currentSort = 'due-asc';
let editTaskId = null;
let msalInstance = null;
let outlookSyncing = false;

// ─── INITIALIZATION ───────────────────────────────────────────────────────
async function loadState() {
  try {
    const res = await fetch(`${API_BASE}/api/tasks`);
    if (!res.ok) throw new Error('Could not load tasks from GitHub');
    const data = await res.json();
    state.tasks = data.tasks || [];
    state.meetings = data.meetings || [];
    state.reminders = data.reminders || [];
    state.aiHistory = data.aiHistory || [];
  } catch (e) {
    console.error(e);
    showToast('Using local data — GitHub sync failed');
    const s = localStorage.getItem(SK);
    if (s) state = Object.assign(state, JSON.parse(s));
  }
}

async function saveState() {
  try {
    localStorage.setItem(SK, JSON.stringify(state));
    const res = await fetch(`${API_BASE}/api/tasks`, {
      method: 'PUT',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        tasks: state.tasks,
        meetings: state.meetings,
        reminders: state.reminders,
        aiHistory: state.aiHistory
      })
    });
    if (!res.ok) throw new Error('GitHub save failed');
  } catch (e) {
    console.error(e);
    showToast('Saved locally — GitHub update failed');
  }
}

function seedDefaults() {
  const d = (o = 0) => {
    const x = new Date(todayDate);
    x.setDate(x.getDate() + o);
    return x.toISOString().slice(0, 10);
  };
  state.tasks = [
    { id: 1, type: 'work', title: 'Finalise Appium V3 build-vs-extend recommendation', note: 'CamAPS test automation decision', priority: 'high', category: 'SME', due: d(1), status: 'todo', done: false, progress: 45 },
    { id: 2, type: 'work', title: 'Review Product A OQ protocol — section 4', note: '', priority: 'high', category: 'SME', due: d(0), status: 'in-progress', done: false, progress: 70 },
    { id: 3, type: 'work', title: 'Update FTE capacity model Q3–Q4 2026', note: '', priority: 'high', category: 'Manager', due: d(2), status: 'todo', done: false, progress: 60 },
    { id: 4, type: 'work', title: 'Product B risk assessment — draft agenda', note: 'BCN kickoff', priority: 'med', category: 'SME', due: d(3), status: 'todo', done: false, progress: 20 },
    { id: 5, type: 'work', title: 'Post cross-site written status update', note: '', priority: 'med', category: 'Agile', due: d(0), status: 'todo', done: false, progress: 0 },
    { id: 6, type: 'personal', title: 'Review heat pump contractor offers', note: 'Compare at least 2 before deciding', priority: 'high', category: 'Home', due: d(2), status: 'todo', done: false, progress: 30 },
    { id: 7, type: 'personal', title: 'Book dentist appointment', note: '', priority: 'low', category: 'Health', due: d(7), status: 'todo', done: false, progress: 0 },
    { id: 8, type: 'work', title: 'CamAPS IQ sign-off', note: '', priority: 'high', category: 'SME', due: d(-1), status: 'done', done: true, progress: 100 }
  ];
  saveState();
}

let nextId = 300;
function uid() {
  return ++nextId + (Date.now() % 1000);
}

// ─── OUTLOOK AUTH ────────────────────────────────────────────────────────
const GRAPH_SCOPES = ['Calendars.Read', 'User.Read'];

function initMsal(clientId) {
  try {
    msalInstance = new msal.PublicClientApplication({
      auth: {
        clientId,
        authority: 'https://login.microsoftonline.com/common',
        redirectUri: window.location.href.split('?')[0].split('#')[0]
      },
      cache: { cacheLocation: 'localStorage', storeAuthStateInCookie: true }
    });
    return true;
  } catch (e) {
    console.error('MSAL init error', e);
    return false;
  }
}

async function connectOutlook() {
  const clientId = document.getElementById('ol-client-id').value.trim();
  if (!clientId || clientId.length < 30) {
    showToast('Please enter a valid Client ID');
    return;
  }
  state.outlookClientId = clientId;
  saveState();
  if (!initMsal(clientId)) {
    showToast('Could not initialise Microsoft login. Check Client ID.');
    return;
  }
  try {
    const result = await msalInstance.loginPopup({ scopes: GRAPH_SCOPES });
    state.outlookUser = result.account?.username || result.account?.name || 'Your account';
    saveState();
    await fetchOutlookEvents();
    showOutlookConnected();
    closeModal('outlook-modal');
    showToast('✓ Outlook connected — calendar synced');
  } catch (e) {
    console.error('Login error', e);
    showToast('Sign-in cancelled or failed. Try again.');
  }
}

async function getGraphToken() {
  if (!msalInstance) return null;
  const accounts = msalInstance.getAllAccounts();
  if (!accounts.length) return null;
  try {
    const r = await msalInstance.acquireTokenSilent({ scopes: GRAPH_SCOPES, account: accounts[0] });
    return r.accessToken;
  } catch (e) {
    try {
      const r = await msalInstance.acquireTokenPopup({ scopes: GRAPH_SCOPES });
      return r.accessToken;
    } catch (e2) {
      return null;
    }
  }
}

async function fetchOutlookEvents() {
  const token = await getGraphToken();
  if (!token) {
    showToast('Could not get calendar access token');
    return;
  }
  outlookSyncing = true;
  updateBannerState();
  try {
    const now = new Date();
    const end = new Date();
    end.setDate(end.getDate() + (state.outlookDays || 14));
    const startISO = now.toISOString();
    const endISO = end.toISOString();
    const url = `https://graph.microsoft.com/v1.0/me/calendarView?startDateTime=${startISO}&endDateTime=${endISO}&$select=subject,start,end,location,onlineMeetingUrl,bodyPreview,organizer&$orderby=start/dateTime`;
    const resp = await fetch(url, {
      headers: {
        Authorization: 'Bearer ' + token,
        'Content-Type': 'application/json'
      }
    });
    if (!resp.ok) {
      const err = await resp.json();
      throw new Error(err.error?.message || resp.statusText);
    }
    const data = await resp.json();
    state.outlookMeetings = (data.value || []).map(ev => ({
      id: 'ol_' + (ev.id?.slice(-12) || uid()),
      title: ev.subject || '(No title)',
      date: ev.start?.dateTime?.slice(0, 10) || '',
      time: ev.start?.dateTime?.slice(11, 16) || '',
      location: ev.location?.displayName || '',
      link: ev.onlineMeetingUrl || '',
      source: 'outlook'
    }));
    state.lastOutlookSync = new Date().toISOString();
    saveState();
    outlookSyncing = false;
    updateBannerState();
    render();
  } catch (e) {
    console.error('Graph fetch error', e);
    outlookSyncing = false;
    updateBannerState();
    showToast('Sync error: ' + e.message);
  }
}

function disconnectOutlook() {
  if (!confirm('Disconnect Outlook? Synced meetings will be removed.')) return;
  state.outlookClientId = '';
  state.outlookUser = '';
  state.outlookMeetings = [];
  msalInstance = null;
  saveState();
  updateBannerState();
  closeModal('outlook-modal');
  render();
  showToast('Outlook disconnected');
}

function handleOutlookIconClick() {
  openOutlookSetup();
}

function openOutlookSetup() {
  const connected = !!(state.outlookClientId && msalInstance && state.outlookUser);
  document.getElementById('ol-setup-view').style.display = connected ? 'none' : 'block';
  document.getElementById('ol-connected-view').style.display = connected ? 'block' : 'none';
  if (connected) {
    document.getElementById('ol-user-display').textContent = state.outlookUser;
    document.getElementById('ol-connect-btn').style.display = 'none';
  } else {
    document.getElementById('ol-connect-btn').style.display = 'block';
    if (state.outlookClientId) document.getElementById('ol-client-id').value = state.outlookClientId;
  }
  document.getElementById('outlook-modal').classList.remove('hidden');
}

function updateBannerState() {
  const banner = document.getElementById('outlook-banner');
  const dot = document.getElementById('ob-dot');
  const txt = document.getElementById('ob-text');
  const btn = document.getElementById('ob-action-btn');

  if (state.outlookClientId && state.outlookUser) {
    banner.className = 'outlook-banner connected';
    dot.style.background = 'var(--success)';
    const last = state.lastOutlookSync ? new Date(state.lastOutlookSync).toLocaleTimeString('en-GB', { hour: '2-digit', minute: '2-digit' }) : '—';
    txt.innerHTML = `<b>Outlook connected</b> · ${state.outlookUser} · Last sync: ${last}`;
    btn.textContent = '↻ Sync';
    btn.onclick = () => fetchOutlookEvents();
  } else {
    banner.className = 'outlook-banner setup';
    dot.style.background = 'var(--warning)';
    txt.innerHTML = '<b>Outlook not connected.</b> Sync your calendar to see meetings automatically.';
    btn.textContent = 'Connect';
    btn.onclick = openOutlookSetup;
  }
}

// ─── UI VIEW SWITCHING ────────────────────────────────────────────────────
function switchView(view) {
  currentView = view;
  document.querySelectorAll('.view').forEach(v => v.classList.remove('active'));
  document.getElementById(view + '-view').classList.add('active');
  document.querySelectorAll('.nav-item').forEach(n => n.classList.remove('active'));
  if (event && event.target) event.target.classList.add('active');

  const titles = { dashboard: 'Dashboard', list: 'List View', board: 'Board View', calendar: 'Calendar' };
  document.getElementById('view-title').textContent = titles[view];

  render();
}

function toggleSidebar() {
  document.getElementById('sidebar').classList.toggle('open');
}

// ─── FILTERING & SORTING ──────────────────────────────────────────────────
function applyFilters() {
  currentFilters.work = document.getElementById('filter-work').checked;
  currentFilters.personal = document.getElementById('filter-personal').checked;
  currentFilters.high = document.getElementById('filter-high').checked;
  currentFilters.overdue = document.getElementById('filter-overdue').checked;
  console.log('Filters applied:', currentFilters);
  render();
}

function applySort() {
  currentSort = document.getElementById('sort-by').value;
  console.log('Sort applied:', currentSort);
  render();
}

function filterTasks() {
  const query = document.getElementById('search-input').value.toLowerCase();
  document.querySelectorAll('.task-card').forEach(card => {
    const title = card.querySelector('.task-title').textContent.toLowerCase();
    card.style.display = title.includes(query) ? '' : 'none';
  });
}

function getFilteredTasks() {
  let tasks = [...state.tasks];

  // If no filters are selected, show all tasks
  const hasActiveFilters = currentFilters.work || currentFilters.personal || currentFilters.high || currentFilters.overdue;

  if (hasActiveFilters) {
    tasks = tasks.filter(t => {
      // Type filters (work/personal)
      if (currentFilters.work || currentFilters.personal) {
        if (currentFilters.work && t.type !== 'work') return false;
        if (currentFilters.personal && t.type !== 'personal') return false;
      }

      // Priority filter
      if (currentFilters.high && t.priority !== 'high') return false;

      // Overdue filter
      if (currentFilters.overdue && (!(!t.done && t.due && t.due < todayStr))) return false;

      return true;
    });
  }

  // Apply sort
  if (currentSort === 'due-asc') {
    tasks.sort((a, b) => (a.due || 'z').localeCompare(b.due || 'z'));
  } else if (currentSort === 'due-desc') {
    tasks.sort((a, b) => (b.due || 'z').localeCompare(a.due || 'z'));
  } else if (currentSort === 'priority') {
    const p = { high: 0, med: 1, low: 2 };
    tasks.sort((a, b) => (p[a.priority] || 3) - (p[b.priority] || 3));
  } else if (currentSort === 'created') {
    tasks.sort((a, b) => b.id - a.id);
  }

  return tasks;
}

// ─── RENDERING ────────────────────────────────────────────────────────────
function render() {
  updateStats();
  switch (currentView) {
    case 'dashboard':
      renderDashboard();
      break;
    case 'list':
      renderListView();
      break;
    case 'board':
      renderBoardView();
      break;
    case 'calendar':
      renderCalendarView();
      break;
  }
  scheduleReminders();
}

function updateStats() {
  const open = state.tasks.filter(t => !t.done);
  const today = open.filter(t => t.due === todayStr);
  const overdue = open.filter(t => t.due && t.due < todayStr);
  const done = state.tasks.filter(t => t.done);

  document.getElementById('d-open').textContent = open.length;
  document.getElementById('d-today').textContent = today.length;
  document.getElementById('d-overdue').textContent = overdue.length;
  document.getElementById('d-done').textContent = done.length;
}

function fmtDate(d) {
  if (!d) return '';
  const p = d.split('-');
  const dt = new Date(+p[0], +p[1] - 1, +p[2]);
  if (d === todayStr) return 'Today';
  const tom = new Date(todayDate);
  tom.setDate(tom.getDate() + 1);
  if (d === tom.toISOString().slice(0, 10)) return 'Tomorrow';
  return dt.toLocaleDateString('en-GB', { weekday: 'short', day: 'numeric', month: 'short' });
}

function isPast(d) {
  return d && d < todayStr;
}

function esc(s) {
  return String(s || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function taskCardHtml(t) {
  const isCompleted = t.done || t.status === 'done';
  const over = !isCompleted && isPast(t.due);
  const priorityLabel = { high: '🔴 High', med: '🟡 Medium', low: '🔵 Low' }[t.priority] || 'Medium';
  const typeIcon = t.type === 'work' ? '💼' : '🌿';

  return `
    <div class="task-card ${t.type}${isCompleted ? ' completed' : ''}" data-id="${t.id}">
      <div class="task-card-header">
        <div class="task-checkbox${isCompleted ? ' checked' : ''}" onclick="toggleTask(${t.id})"></div>
        <div class="task-title">${esc(t.title)}</div>
        <div class="task-actions">
          <button class="task-action-btn" onclick="editTask(${t.id})">✏️</button>
          <button class="task-action-btn" onclick="deleteTask(${t.id})">🗑️</button>
        </div>
      </div>
      <div class="task-meta">
        <span class="badge badge-${t.type}">${typeIcon} ${t.type === 'work' ? 'Work' : 'Personal'}</span>
        <span class="badge-priority ${t.priority}">${priorityLabel}</span>
        ${t.due ? `<span class="due-date${over ? ' overdue' : ''}">📅 ${fmtDate(t.due)}${over ? ' · overdue' : ''}</span>` : ''}
        ${t.category ? `<span style="font-size:11px;color:var(--text-tertiary)">${esc(t.category)}</span>` : ''}
      </div>
    </div>
  `;
}

function renderDashboard() {
  const filtered = getFilteredTasks();
  const highPri = filtered.filter(t => !t.done && t.priority === 'high').slice(0, 5);
  const today = filtered.filter(t => !t.done && t.due === todayStr).slice(0, 5);
  const overdue = filtered.filter(t => !t.done && t.due && t.due < todayStr).slice(0, 5);

  document.getElementById('dashboard-high').innerHTML = highPri.map(taskCardHtml).join('') || '<div style="color:var(--text-tertiary);font-size:12px">No high priority tasks</div>';
  document.getElementById('dashboard-today').innerHTML = today.map(taskCardHtml).join('') || '<div style="color:var(--text-tertiary);font-size:12px">No tasks due today</div>';
  document.getElementById('dashboard-overdue').innerHTML = overdue.map(taskCardHtml).join('') || '<div style="color:var(--text-tertiary);font-size:12px">No overdue tasks</div>';
}

function renderListView() {
  const tasks = getFilteredTasks();
  const html = tasks.map(taskCardHtml).join('');
  document.getElementById('list-content').innerHTML = html || '<div style="color:var(--text-tertiary);text-align:center;padding:40px">No tasks found</div>';
}

function renderBoardView() {
  const tasks = getFilteredTasks();
  const statuses = { todo: 'to-do', 'in-progress': 'in-progress', done: 'done' };

  Object.entries(statuses).forEach(([status, id]) => {
    const statusTasks = tasks.filter(t => (t.status || 'todo') === status);
    const html = statusTasks.map(t => `
      <div class="kanban-task" draggable="true" data-id="${t.id}" onclick="event.stopPropagation()">
        <div style="display:flex;align-items:center;gap:8px;margin-bottom:8px">
          <div class="task-checkbox${t.done ? ' checked' : ''}" onclick="toggleTask(${t.id})"></div>
          <span class="task-title" style="flex:1">${esc(t.title)}</span>
        </div>
        <div class="task-meta">
          <span class="badge badge-${t.type}">${t.type === 'work' ? '💼' : '🌿'}</span>
          <span class="badge-priority ${t.priority}">${t.priority === 'high' ? '🔴' : t.priority === 'med' ? '🟡' : '🔵'}</span>
          ${t.due ? `<span class="due-date">${fmtDate(t.due)}</span>` : ''}
        </div>
      </div>
    `).join('');

    document.getElementById('column-' + id).innerHTML = html;
    document.getElementById('count-' + id).textContent = statusTasks.length;
  });
}

function renderCalendarView() {
  document.getElementById('calendar-content').innerHTML = '<p style="color:var(--text-tertiary)">Calendar view coming soon</p>';
}

// ─── TASK ACTIONS ────────────────────────────────────────────────────────
function toggleTask(id) {
  const t = state.tasks.find(x => x.id === id);
  if (t) {
    t.done = !t.done;
    t.status = t.done ? 'done' : 'todo';
    saveState();
    render();
    showToast(t.done ? 'Task completed ✅' : 'Task reopened');
  }
}

function deleteTask(id) {
  if (confirm('Delete this task?')) {
    state.tasks = state.tasks.filter(x => x.id !== id);
    saveState();
    render();
    showToast('Task deleted');
  }
}

function openAddTask() {
  editTaskId = null;
  document.getElementById('task-modal-title').textContent = 'Add New Task';
  document.getElementById('t-title').value = '';
  document.getElementById('t-notes').value = '';
  document.getElementById('t-type').value = 'work';
  document.getElementById('t-priority').value = 'med';
  document.getElementById('t-due').value = todayStr;
  document.getElementById('t-status').value = 'todo';
  document.getElementById('t-category').value = 'SME';
  document.getElementById('task-modal').classList.remove('hidden');
  setTimeout(() => document.getElementById('t-title').focus(), 100);
}

function openAddTaskWithStatus(status) {
  openAddTask();
  document.getElementById('t-status').value = status;
}

function editTask(id) {
  const t = state.tasks.find(x => x.id === id);
  if (!t) return;
  editTaskId = id;
  document.getElementById('task-modal-title').textContent = 'Edit Task';
  document.getElementById('t-title').value = t.title;
  document.getElementById('t-notes').value = t.note || '';
  document.getElementById('t-type').value = t.type;
  document.getElementById('t-priority').value = t.priority;
  document.getElementById('t-due').value = t.due || '';
  document.getElementById('t-status').value = t.status || 'todo';
  document.getElementById('t-category').value = t.category || 'SME';
  document.getElementById('task-modal').classList.remove('hidden');
}

function saveTask() {
  const title = document.getElementById('t-title').value.trim();
  if (!title) {
    showToast('Task title is required');
    return;
  }

  const task = {
    id: editTaskId || uid(),
    type: document.getElementById('t-type').value,
    title,
    note: document.getElementById('t-notes').value.trim(),
    priority: document.getElementById('t-priority').value,
    category: document.getElementById('t-category').value,
    due: document.getElementById('t-due').value,
    status: document.getElementById('t-status').value,
    done: document.getElementById('t-status').value === 'done',
    progress: 0
  };

  if (editTaskId) {
    state.tasks = state.tasks.map(x => (x.id === editTaskId ? task : x));
  } else {
    state.tasks.push(task);
  }

  saveState();
  closeModal('task-modal');
  render();
  showToast(editTaskId ? 'Task updated' : 'Task added ✅');
}

// ─── NOTIFICATIONS ────────────────────────────────────────────────────────
function requestNotifPermission() {
  if (!('Notification' in window)) {
    showToast('Notifications not supported here');
    return;
  }
  if (Notification.permission === 'granted') {
    showToast('Notifications already enabled ✅');
    return;
  }
  Notification.requestPermission().then(p => {
    showToast(p === 'granted' ? 'Notifications enabled ✅' : 'Blocked — enable in browser settings');
  });
}

function scheduleReminders() {
  state.tasks
    .filter(t => !t.done && t.due)
    .forEach(t => {
      const dt = new Date(t.due);
      const ms = dt.getTime() - Date.now();
      if (ms > 0 && ms < 7 * 24 * 60 * 60 * 1000) {
        setTimeout(() => {
          if (Notification.permission === 'granted') {
            new Notification('MyTaskManager', { body: t.title });
          }
        }, ms);
      }
    });
}

// ─── AI CHAT ──────────────────────────────────────────────────────────────
function openAIChat() {
  const el = document.getElementById('ai-messages');
  if (!el.children.length) {
    addAIMsg('assistant', "Hi! I'm your AI assistant. I can help you prioritise tasks, plan your day, or think through challenges.");
  }
  renderAISuggestions();
  document.getElementById('ai-modal').classList.remove('hidden');
  setTimeout(() => document.getElementById('ai-input').focus(), 100);
}

function renderAISuggestions() {
  const overdue = state.tasks.filter(t => !t.done && isPast(t.due)).length;
  const suggestions = [
    'Help me prioritise today',
    overdue > 0 ? `I have ${overdue} overdue tasks — what now?` : 'What should I focus on?',
    'Help me think through decisions'
  ];
  const html = suggestions.map(s => `<button class="ai-quick" onclick="quickAsk('${s.replace(/'/g, "&#39;")}')">${s}</button>`).join('');
  document.getElementById('ai-suggestions').innerHTML = html;
}

function quickAsk(text) {
  document.getElementById('ai-input').value = text;
  sendAI();
}

function addAIMsg(role, text) {
  const el = document.getElementById('ai-messages');
  const div = document.createElement('div');
  div.className = `ai-msg-wrap${role === 'user' ? ' user' : ''}`;
  div.innerHTML = `<div class="ai-msg-label">${role === 'user' ? 'You' : 'AI ✦'}</div><div class="ai-msg-bubble ${role}">${esc(text)}</div>`;
  el.appendChild(div);
  el.scrollTop = el.scrollHeight;
}

async function sendAI() {
  const input = document.getElementById('ai-input');
  const text = input.value.trim();
  if (!text) return;

  input.value = '';
  addAIMsg('user', text);

  const apiKey = (state.anthropicKey || '').trim();
  if (!apiKey) {
    addAIMsg('assistant', '⚠️ No API key found. Open ⚙️ Settings and enter your Anthropic API key.');
    return;
  }

  try {
    const openTasks = state.tasks
      .filter(t => !t.done)
      .slice(0, 10)
      .map(t => `- [${t.type}/${t.priority}] ${t.title}${t.due ? ' (due ' + fmtDate(t.due) + ')' : ''}`)
      .join('\n');

    const sys = `You are a smart personal task manager assistant. Today is ${todayStr}.

Open tasks (top 10):
${openTasks || 'none'}

Be concise, practical, and direct.`;

    const resp = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': apiKey,
        'anthropic-version': '2023-06-01'
      },
      body: JSON.stringify({
        model: 'claude-opus-4-1',
        max_tokens: 500,
        system: sys,
        messages: [...state.aiHistory.slice(-4), { role: 'user', content: text }]
      })
    });

    const data = await resp.json();

    if (!resp.ok || data.error) {
      addAIMsg('assistant', '❌ API error: ' + (data.error?.message || 'Unknown error'));
      return;
    }

    const reply = data.content?.[0]?.text || 'Sorry, no response. Please try again.';
    addAIMsg('assistant', reply);
    state.aiHistory.push({ role: 'user', content: text }, { role: 'assistant', content: reply });
    if (state.aiHistory.length > 20) state.aiHistory = state.aiHistory.slice(-16);
    saveState();
  } catch (e) {
    addAIMsg('assistant', '❌ Network error: ' + e.message);
  }
}

// ─── SETTINGS ─────────────────────────────────────────────────────────────
function openSettings() {
  const k = state.anthropicKey || '';
  document.getElementById('s-apikey').value = k ? k.slice(0, 8) + '...' : '';
  document.getElementById('s-appname').value = state.appName || 'MyTaskManager';
  const status = document.getElementById('s-key-status');
  if (k) {
    status.textContent = '✓ API key saved (' + k.slice(0, 12) + '...)';
    status.style.color = 'var(--success)';
  } else {
    status.textContent = 'No API key set — AI features disabled';
    status.style.color = 'var(--danger)';
  }
  document.getElementById('settings-modal').classList.remove('hidden');
}

function saveSettings() {
  const raw = document.getElementById('s-apikey').value.trim();
  if (raw && !raw.includes('...')) state.anthropicKey = raw;
  const name = document.getElementById('s-appname').value.trim();
  if (name) state.appName = name;
  saveState();
  closeModal('settings-modal');
  showToast('Settings saved ✓');
}

function clearAllData() {
  if (!confirm('Delete ALL tasks, meetings, reminders and settings? This cannot be undone.')) return;
  localStorage.removeItem(SK);
  location.reload();
}

// ─── UTILITIES ────────────────────────────────────────────────────────────
function closeModal(id) {
  document.getElementById(id).classList.add('hidden');
}

let toastT;
function showToast(msg) {
  const el = document.getElementById('toast');
  el.textContent = msg;
  el.classList.remove('hidden');
  clearTimeout(toastT);
  toastT = setTimeout(() => el.classList.add('hidden'), 2800);
}

// ─── INIT ────────────────────────────────────────────────────────────────
window.addEventListener('DOMContentLoaded', async () => {
  await loadState();

  if (!state.tasks.length) {
    seedDefaults();
  }

  if (state.outlookClientId) {
    initMsal(state.outlookClientId);
  }

  updateBannerState();
  switchView('dashboard');

  if (Notification.permission === 'granted') {
    const notifBtn = document.querySelector('[title="Enable reminders"]');
    if (notifBtn) notifBtn.style.color = 'var(--primary)';
  }

  if (state.outlookClientId && msalInstance && state.outlookUser) {
    setTimeout(fetchOutlookEvents, 1500);
  }

  // Sync Outlook on load if connected
  if (state.outlookClientId && msalInstance && state.outlookUser) {
    setTimeout(fetchOutlookEvents, 2000);
  }
});
