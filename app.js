// ─── STATE ───────────────────────────────────────────────────────────────────
const todayDate = new Date();
const todayStr = todayDate.toISOString().slice(0,10);
document.getElementById('today-sub').textContent = todayDate.toLocaleDateString('en-GB',{weekday:'long',day:'numeric',month:'long'});

const SK = 'mytm_v2';
const API_BASE = window.location.origin;
let state = {tasks:[], meetings:[], reminders:[], aiHistory:[], outlookClientId:'', outlookDays:14, outlookMeetings:[], outlookUser:''};
let currentTab = 'today';
let editTaskId = null;
let currentType = 'work';
let fabOpen = false;
let msalInstance = null;
let outlookSyncing = false;

async function loadState(){
  try {
    const res = await fetch(`${API_BASE}/api/tasks`);
    if (!res.ok) throw new Error('Could not load tasks from GitHub');
    const data = await res.json();

    state.tasks = data.tasks || [];
    state.meetings = data.meetings || [];
    state.reminders = data.reminders || [];
    state.aiHistory = data.aiHistory || [];
    state.initialized = true;
  } catch (e) {
    console.error(e);
    showToast('Using local data — GitHub sync failed');
    const s = localStorage.getItem(SK);
    if (s) state = Object.assign(state, JSON.parse(s));
  }
}
async function saveState(){
  try {
    localStorage.setItem(SK, JSON.stringify(state));

    const res = await fetch(`${API_BASE}/api/tasks`, {
      method: 'PUT',
      headers: {'Content-Type': 'application/json'},
      body: JSON.stringify({
        tasks: state.tasks,
        meetings: state.meetings,
        reminders: state.reminders,
        aiHistory: state.aiHistory
      })
    });

    if (!res.ok) throw new Error('GitHub save failed');
    return true;
  } catch (e) {
    console.error(e);
    showToast('Saved locally — GitHub update failed');
    return false;
  }
}
function checkStorageHealth(){
  try{
    const testKey='mytm_storage_test';
    localStorage.setItem(testKey,'1');
    const ok = localStorage.getItem(testKey)==='1';
    localStorage.removeItem(testKey);
    return ok;
  }catch(e){ return false; }
}

function seedDefaults(){
  const d=(o=0)=>{const x=new Date(todayDate);x.setDate(x.getDate()+o);return x.toISOString().slice(0,10)};
  state.tasks=[
    {id:1,type:'work',title:'Finalise Appium V3 build-vs-extend recommendation',note:'CamAPS test automation decision',pri:'high',cat:'SME',due:d(1),done:false,remind:'',progress:45},
    {id:2,type:'work',title:'Review Product A OQ protocol — section 4',note:'',pri:'high',cat:'SME',due:d(0),done:false,remind:'',progress:70},
    {id:3,type:'work',title:'Update FTE capacity model Q3–Q4 2026',note:'',pri:'high',cat:'Manager',due:d(2),done:false,remind:'',progress:60},
    {id:4,type:'work',title:'Product B risk assessment — draft agenda',note:'BCN kickoff',pri:'med',cat:'SME',due:d(3),done:false,remind:'',progress:20},
    {id:5,type:'work',title:'Post cross-site written status update',note:'',pri:'med',cat:'Agile',due:d(0),done:false,remind:'',progress:0},
    {id:6,type:'personal',title:'Review heat pump contractor offers',note:'Compare at least 2 before deciding',pri:'high',cat:'Home',due:d(2),done:false,remind:'',progress:30},
    {id:7,type:'personal',title:'Book dentist appointment',note:'',pri:'low',cat:'Health',due:d(7),done:false,remind:'',progress:0},
    {id:8,type:'work',title:'CamAPS IQ sign-off',note:'',pri:'high',cat:'SME',due:d(-1),done:false,remind:'',progress:90},
  ];
  state.meetings=[
    {id:1,title:'All-team stand-up BCN + CH',date:d(0),time:'08:00',site:'all',link:'https://teams.microsoft.com',note:'15 min max',source:'manual'},
    {id:2,title:'Sprint planning',date:d(0),time:'09:00',site:'all',link:'https://teams.microsoft.com',note:'',source:'manual'},
    {id:3,title:'Cross-site sync',date:d(1),time:'14:00',site:'all',link:'https://teams.microsoft.com',note:'BCN ↔ CH',source:'manual'},
    {id:4,title:'Automation strategy review',date:d(2),time:'10:00',site:'CH',link:'',note:'Appium V3 prep',source:'manual'},
  ];
  saveState();
}

let nextId=300;
function uid(){ return ++nextId+(Date.now()%1000); }

// ─── OUTLOOK AUTH ─────────────────────────────────────────────────────────────
const GRAPH_SCOPES = ['Calendars.Read','User.Read'];

function initMsal(clientId){
  try{
    msalInstance = new msal.PublicClientApplication({
      auth:{
        clientId,
        authority:'https://login.microsoftonline.com/common',
        redirectUri: window.location.href.split('?')[0].split('#')[0]
      },
      cache:{ cacheLocation:'localStorage', storeAuthStateInCookie:true }
    });
    return true;
  }catch(e){ console.error('MSAL init error',e); return false; }
}

async function connectOutlook(){
  const clientId = document.getElementById('ol-client-id').value.trim();
  if(!clientId || clientId.length < 30){ showToast('Please enter a valid Client ID'); return; }
  state.outlookClientId = clientId;
  state.outlookDays = parseInt(document.getElementById('ol-days-ahead').value)||14;
  saveState();
  if(!initMsal(clientId)){ showToast('Could not initialise Microsoft login. Check Client ID.'); return; }
  try{
    setBannerSyncing();
    const result = await msalInstance.loginPopup({scopes: GRAPH_SCOPES});
    state.outlookUser = result.account?.username || result.account?.name || 'Your account';
    saveState();
    await fetchOutlookEvents();
    showOutlookConnected();
    closeModal('outlook-modal');
    showToast('✓ Outlook connected — calendar synced');
  }catch(e){
    console.error('Login error',e);
    showToast('Sign-in cancelled or failed. Try again.');
    updateBannerState();
  }
}

function handleOutlookIconClick(){
  if(state.outlookClientId && msalInstance){
    openOutlookSetup();
  } else {
    openOutlookSetup();
  }
}

async function getGraphToken(){
  if(!msalInstance) return null;
  const accounts = msalInstance.getAllAccounts();
  if(!accounts.length) return null;
  try{
    const r = await msalInstance.acquireTokenSilent({scopes:GRAPH_SCOPES, account:accounts[0]});
    return r.accessToken;
  }catch(e){
    try{
      const r = await msalInstance.acquireTokenPopup({scopes:GRAPH_SCOPES});
      return r.accessToken;
    }catch(e2){ return null; }
  }
}

async function fetchOutlookEvents(){
  const token = await getGraphToken();
  if(!token){ showToast('Could not get calendar access token'); return; }
  outlookSyncing = true;
  updateBannerState();
  try{
    const now = new Date();
    const end = new Date();
    end.setDate(end.getDate()+(state.outlookDays||14));
    const startISO = now.toISOString();
    const endISO   = end.toISOString();
    const url = `https://graph.microsoft.com/v1.0/me/calendarView?startDateTime=${startISO}&endDateTime=${endISO}&$select=subject,start,end,location,onlineMeetingUrl,bodyPreview,organizer&$orderby=start/dateTime&$top=100`;
    const resp = await fetch(url,{headers:{Authorization:'Bearer '+token,'Content-Type':'application/json'}});
    if(!resp.ok){ const err=await resp.json(); throw new Error(err.error?.message||resp.statusText); }
    const data = await resp.json();
    state.outlookMeetings = (data.value||[]).map(ev=>({
      id: 'ol_'+ev.id?.slice(-12)||uid(),
      title: ev.subject||'(No title)',
      date: ev.start?.dateTime?.slice(0,10)||'',
      time: ev.start?.dateTime?.slice(11,16)||'',
      endTime: ev.end?.dateTime?.slice(11,16)||'',
      location: ev.location?.displayName||'',
      link: ev.onlineMeetingUrl||'',
      note: ev.bodyPreview?.slice(0,120)||'',
      organizer: ev.organizer?.emailAddress?.name||'',
      source:'outlook'
    }));
    state.lastOutlookSync = new Date().toISOString();
    saveState();
    outlookSyncing = false;
    updateBannerState();
    render();
  }catch(e){
    console.error('Graph fetch error',e);
    outlookSyncing = false;
    updateBannerState();
    showToast('Sync error: '+e.message);
  }
}

async function syncOutlookNow(){
  const days = parseInt(document.getElementById('ol-days-ahead-2')?.value)||14;
  state.outlookDays = days;
  await fetchOutlookEvents();
  showToast('✓ Calendar synced');
}

function disconnectOutlook(){
  if(!confirm('Disconnect Outlook? Synced meetings will be removed.')) return;
  state.outlookClientId=''; state.outlookUser=''; state.outlookMeetings=[]; state.lastOutlookSync='';
  msalInstance=null; saveState(); updateBannerState(); closeModal('outlook-modal'); render();
  showToast('Outlook disconnected');
}

function openOutlookSetup(){
  const connected = !!(state.outlookClientId && msalInstance && state.outlookUser);
  document.getElementById('ol-setup-view').style.display = connected?'none':'block';
  document.getElementById('ol-connected-view').style.display = connected?'block':'none';
  if(connected){
    document.getElementById('ol-user-display').textContent = state.outlookUser;
    document.getElementById('ol-modal-footer').style.display='none';
  } else {
    document.getElementById('ol-modal-footer').style.display='flex';
    if(state.outlookClientId) document.getElementById('ol-client-id').value=state.outlookClientId;
  }
  document.getElementById('outlook-modal').classList.remove('hidden');
}

// ─── BANNER ───────────────────────────────────────────────────────────────────
function setBannerSyncing(){
  document.getElementById('ob-dot').style.background='var(--outlook)';
  document.getElementById('ob-text').innerHTML='<b>Connecting to Outlook...</b>';
  document.getElementById('ob-action-btn').innerHTML='<span class="ob-spin">↻</span>';
  document.getElementById('ob-action-btn').disabled=true;
  document.getElementById('outlook-banner').className='outlook-banner connected';
}
function showOutlookConnected(){
  updateBannerState();
  document.getElementById('ol-icon-btn').classList.add('active');
}
function updateBannerState(){
  const banner = document.getElementById('outlook-banner');
  const dot = document.getElementById('ob-dot');
  const txt = document.getElementById('ob-text');
  const btn = document.getElementById('ob-action-btn');
  btn.disabled=false;
  if(outlookSyncing){
    banner.className='outlook-banner connected';
    dot.style.background='var(--outlook)';
    txt.innerHTML='<b>Syncing calendar...</b>';
    btn.innerHTML='<span class="ob-spin">↻</span>';
    return;
  }
  if(state.outlookClientId && state.outlookUser){
    banner.className='outlook-banner connected';
    dot.style.background='var(--personal)';
    const last = state.lastOutlookSync ? new Date(state.lastOutlookSync).toLocaleTimeString('en-GB',{hour:'2-digit',minute:'2-digit'}) : '—';
    txt.innerHTML=`<b>Outlook connected</b> · ${state.outlookUser} · Last sync: ${last}`;
    btn.textContent='↻ Sync'; btn.className='ob-btn muted';
    btn.onclick=()=>fetchOutlookEvents();
    document.getElementById('ol-icon-btn').classList.add('active');
  } else {
    banner.className='outlook-banner setup';
    dot.style.background='var(--med)';
    txt.innerHTML='<b>Outlook not connected.</b> Tap to sync your calendar automatically.';
    btn.textContent='Connect'; btn.className='ob-btn primary';
    btn.onclick=openOutlookSetup;
    document.getElementById('ol-icon-btn').classList.remove('active');
  }
}

// ─── RENDER ───────────────────────────────────────────────────────────────────
function render(){
  updateStats();
  const c=document.getElementById('main-content');
  if(currentTab==='today') c.innerHTML=renderToday();
  else if(currentTab==='tasks') c.innerHTML=renderTasks();
  else if(currentTab==='meetings') c.innerHTML=renderMeetings();
  else if(currentTab==='reminders') c.innerHTML=renderReminders();
  scheduleReminders();
}

function allMeetings(){
  const manual = state.meetings;
  const ol = state.outlookMeetings;
  return [...manual,...ol].sort((a,b)=>a.date.localeCompare(b.date)||a.time.localeCompare(b.time));
}

function updateStats(){
  const open=state.tasks.filter(t=>!t.done);
  document.getElementById('s-total').textContent=open.length;
  document.getElementById('s-today').textContent=open.filter(t=>t.due===todayStr).length;
  document.getElementById('s-over').textContent=open.filter(t=>t.due&&t.due<todayStr).length;
  document.getElementById('s-done').textContent=state.tasks.filter(t=>t.done).length;
}

function fmtDate(d){
  if(!d) return '';
  const p=d.split('-');
  const dt=new Date(+p[0],+p[1]-1,+p[2]);
  if(d===todayStr) return 'Today';
  const tom=new Date(todayDate); tom.setDate(tom.getDate()+1);
  if(d===tom.toISOString().slice(0,10)) return 'Tomorrow';
  return dt.toLocaleDateString('en-GB',{weekday:'short',day:'numeric',month:'short'});
}
function isPast(d){ return d&&d<todayStr; }
function esc(s){ return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;'); }

function progColor(pct, done){
  if(done) return 'p-done';
  if(pct===100) return 'p-done';
  if(pct>=60) return 'p-low';
  if(pct>=30) return 'p-med';
  return 'p-high';
}

function taskCard(t){
  const over=!t.done&&isPast(t.due);
  const priLabel=t.pri==='high'?'High':t.pri==='med'?'Medium':'Low';
  const pct = t.done ? 100 : (t.progress||0);
  const pClass = progColor(pct, t.done);
  // 5 clickable step markers: 0,25,50,75,100
  const steps=[0,25,50,75,100];
  const stepsHtml=steps.map(s=>`<div class="prog-step${pct>=s&&pct>0||pct===100?' active':''}" onclick="setProgress(${t.id},${s})" title="${s}%"></div>`).join('');
  return `<div class="task ${t.type}${t.done?' done-task':''}">
    <div class="task-cb${t.done?' done':''}" onclick="toggleTask(${t.id})">
      ${t.done?'<svg width="11" height="11" viewBox="0 0 12 12" fill="none" stroke="#fff" stroke-width="2.2" stroke-linecap="round"><polyline points="2,6 5,9 10,3"/></svg>':''}
    </div>
    <div class="task-body">
      <div class="task-title">${esc(t.title)}</div>
      ${t.note?`<div class="task-note">${esc(t.note)}</div>`:''}
      <div class="task-meta">
        <span class="badge badge-${t.type}">${t.type==='work'?'💼 Work':'🌿 Personal'}</span>
        <span class="badge badge-${t.pri}">${priLabel}</span>
        <span class="badge-cat">${esc(t.cat)}</span>
        ${t.due?`<span class="due-label${over?' overdue':''}">📅 ${fmtDate(t.due)}${over?' · overdue':''}</span>`:''}
      </div>
      <div class="prog-wrap">
        <div class="prog-header">
          <span class="prog-label">Progress</span>
          <span class="prog-pct">${pct}%</span>
        </div>
        <div class="prog-track" onclick="cycleProgress(${t.id},event)" title="Click to update progress">
          <div class="prog-fill ${pClass}" style="width:${pct}%"></div>
        </div>
        <div class="prog-steps">${stepsHtml}</div>
      </div>
    </div>
    <div class="task-actions">
      <button class="task-act-btn ai-btn" onclick="askAIAboutTask(${t.id})" title="Ask AI">✦</button>
      <button class="task-act-btn" onclick="editTask(${t.id})" title="Edit">✏️</button>
      <button class="task-act-btn" onclick="deleteTask(${t.id})" title="Delete">🗑️</button>
    </div>
  </div>`;
}

function mtgCard(m){
  const isOl=m.source==='outlook';
  const siteDot=m.site==='CH'?'dot-ch':m.site==='BCN'?'dot-bcn':m.site==='personal'?'dot-p':'dot-all';
  const siteLabel=m.site==='CH'?'Switzerland':m.site==='BCN'?'Barcelona':m.site==='personal'?'Personal':'Both sites';
  const delBtn=isOl?'':`<button class="mtg-del" onclick="deleteMtg('${m.id}')" title="Delete">✕</button>`;
  const endTime=m.endTime?`–${m.endTime}`:'';
  const loc=m.location?`<span>📍 ${esc(m.location.slice(0,40))}</span>`:'';
  return `<div class="mtg ${isOl?'outlook-mtg':'manual-mtg'}">
    <div class="mtg-time-col">
      <div class="mtg-time">${m.time}${endTime?'<br><span style="font-size:10px;color:var(--text3)">'+endTime+'</span>':''}</div>
      <div class="mtg-date">${fmtDate(m.date)}</div>
    </div>
    <div class="mtg-divider"></div>
    <div class="mtg-body">
      <div class="mtg-title">${esc(m.title)}</div>
      <div class="mtg-meta">
        <span class="mtg-source ${isOl?'ol':'manual'}">${isOl?'📅 Outlook':'Manual'}</span>
        ${isOl?`<span>${esc(m.organizer||'')}</span>`:
          `<span><span class="site-dot ${siteDot}"></span>${siteLabel}</span>`}
        ${loc}
        ${m.note&&!isOl?`<span>${esc(m.note.slice(0,60))}</span>`:''}
      </div>
      ${m.link?`<a class="mtg-link" href="${esc(m.link)}" target="_blank">🔗 Join meeting ↗</a>`:''}
    </div>
    ${delBtn}
  </div>`;
}

function renderToday(){
  const all = allMeetings();
  const todayMtgs = all.filter(m=>m.date===todayStr);
  const highTasks = state.tasks.filter(t=>!t.done&&(t.due===todayStr||isPast(t.due)||t.pri==='high'))
    .sort((a,b)=>(['high','med','low'].indexOf(a.pri)-['high','med','low'].indexOf(b.pri)));
  const workT=highTasks.filter(t=>t.type==='work');
  const persT=highTasks.filter(t=>t.type==='personal');
  let h='';
  if(todayMtgs.length){
    const olCount=todayMtgs.filter(m=>m.source==='outlook').length;
    h+=`<div class="sec-label">Today's meetings <span>${todayMtgs.length}</span>${olCount?`<span class="ol-badge">📅 ${olCount} from Outlook</span>`:''}</div>`;
    h+=todayMtgs.map(mtgCard).join('');
  }
  if(workT.length){ h+=`<div class="sec-label">Work priority <span>${workT.length}</span></div>`+workT.map(taskCard).join(''); }
  if(persT.length){ h+=`<div class="sec-label">Personal priority <span>${persT.length}</span></div>`+persT.map(taskCard).join(''); }
  if(!h) h=emptyState('📋','All clear!','No priority items for today.');
  return h;
}

function renderTasks(){
  const fEl=`<div class="filter-bar">
    <button class="chip on" id="tf-all" onclick="taskFilter('all',this)">All open</button>
    <button class="chip" id="tf-work" onclick="taskFilter('work',this)">💼 Work</button>
    <button class="chip personal" id="tf-personal" onclick="taskFilter('personal',this)">🌿 Personal</button>
    <button class="chip" id="tf-high" onclick="taskFilter('high',this)">🔴 High</button>
    <button class="chip" id="tf-overdue" onclick="taskFilter('overdue',this)">⚠️ Overdue</button>
    <button class="chip" id="tf-done" onclick="taskFilter('done',this)">✅ Done</button>
  </div><div id="task-list-inner"></div>`;
  setTimeout(()=>renderTaskList('all'),0);
  return fEl;
}

function renderTaskList(f){
  let t=[...state.tasks];
  const p={high:0,med:1,low:2};
  if(f==='all') t=t.filter(x=>!x.done);
  else if(f==='work') t=t.filter(x=>!x.done&&x.type==='work');
  else if(f==='personal') t=t.filter(x=>!x.done&&x.type==='personal');
  else if(f==='high') t=t.filter(x=>!x.done&&x.pri==='high');
  else if(f==='overdue') t=t.filter(x=>!x.done&&isPast(x.due));
  else if(f==='done') t=t.filter(x=>x.done);
  t.sort((a,b)=>p[a.pri]-p[b.pri]||(a.due||'z').localeCompare(b.due||'z'));
  const el=document.getElementById('task-list-inner');
  if(el) el.innerHTML=t.length?t.map(taskCard).join(''):emptyState('✅','No tasks here','Add a task with the + button.');
}

function taskFilter(f,btn){
  document.querySelectorAll('.filter-bar .chip').forEach(c=>c.classList.remove('on'));
  btn.classList.add('on');
  renderTaskList(f);
}

function renderMeetings(){
  const filter=`<div class="filter-bar">
    <button class="chip on" id="mf-all" onclick="meetingFilter('all',this)">All</button>
    <button class="chip" id="mf-ol" onclick="meetingFilter('outlook',this)">📅 Outlook</button>
    <button class="chip" id="mf-manual" onclick="meetingFilter('manual',this)">Manual</button>
    <button class="chip" id="mf-today" onclick="meetingFilter('today',this)">Today</button>
    <button class="chip" id="mf-week" onclick="meetingFilter('week',this)">This week</button>
  </div><div id="mtg-list-inner"></div>`;
  setTimeout(()=>renderMeetingList('all'),0);
  return filter;
}

function renderMeetingList(f){
  let m=allMeetings();
  const weekEnd=new Date(todayDate); weekEnd.setDate(weekEnd.getDate()+7);
  const weekEndStr=weekEnd.toISOString().slice(0,10);
  if(f==='outlook') m=m.filter(x=>x.source==='outlook');
  else if(f==='manual') m=m.filter(x=>x.source==='manual'||!x.source);
  else if(f==='today') m=m.filter(x=>x.date===todayStr);
  else if(f==='week') m=m.filter(x=>x.date>=todayStr&&x.date<=weekEndStr);
  const el=document.getElementById('mtg-list-inner');
  if(!el) return;
  if(!m.length){ el.innerHTML=emptyState('📅','No meetings','Add one with + or connect Outlook.'); return; }
  const groups={};
  m.forEach(x=>{ if(!groups[x.date]) groups[x.date]=[]; groups[x.date].push(x); });
  let h='';
  Object.keys(groups).sort().forEach(d=>{
    const olC=groups[d].filter(x=>x.source==='outlook').length;
    h+=`<div class="sec-label">${fmtDate(d)||d} <span>${groups[d].length}</span>${olC?`<span class="ol-badge">📅 ${olC} Outlook</span>`:''}</div>`;
    h+=groups[d].map(mtgCard).join('');
  });
  el.innerHTML=h;
}

function meetingFilter(f,btn){
  document.querySelectorAll('.filter-bar .chip').forEach(c=>c.classList.remove('on'));
  btn.classList.add('on');
  renderMeetingList(f);
}

function renderReminders(){
  const all=[...state.reminders].sort((a,b)=>(a.date+a.time).localeCompare(b.date+b.time));
  if(!all.length) return emptyState('⏰','No reminders','Add reminders to get notified.');
  let h='';
  all.forEach(r=>{
    h+=`<div class="reminder-item">
      <div class="reminder-time" style="${r.date<todayStr?'color:var(--high)':''}">${r.time||'—'}<div style="font-size:9px;color:var(--text3);margin-top:1px">${fmtDate(r.date)||r.date}</div></div>
      <div class="reminder-body"><div class="reminder-title">${esc(r.title)}</div>${r.repeat!=='none'?`<div class="reminder-sub">↻ ${r.repeat}</div>`:''}</div>
      <button class="reminder-del" onclick="deleteReminder(${r.id})">✕</button>
    </div>`;
  });
  return h;
}

function emptyState(icon,title,sub){
  return `<div class="empty-state"><div class="empty-icon">${icon}</div><div class="empty-title">${title}</div><div class="empty-sub">${sub}</div></div>`;
}

// ─── ACTIONS ──────────────────────────────────────────────────────────────────
function toggleTask(id){ const t=state.tasks.find(x=>x.id===id); if(t){t.done=!t.done;saveState();render();showToast(t.done?'Done ✅':'Reopened');} }
function deleteTask(id){ state.tasks=state.tasks.filter(x=>x.id!==id); saveState();render();showToast('Task deleted'); }
function deleteMtg(id){ state.meetings=state.meetings.filter(x=>x.id!=id); saveState();render();showToast('Meeting removed'); }
function deleteReminder(id){ state.reminders=state.reminders.filter(x=>x.id!==id); saveState();render();showToast('Reminder removed'); }

function setProgress(id, pct){
  const t=state.tasks.find(x=>x.id===id); if(!t) return;
  t.progress=pct;
  if(pct===100 && !t.done){ t.done=true; showToast('Task marked complete ✅'); }
  else if(pct<100 && t.done){ t.done=false; }
  saveState(); render();
}
function cycleProgress(id, e){
  const t=state.tasks.find(x=>x.id===id); if(!t) return;
  const track=e.currentTarget;
  const rect=track.getBoundingClientRect();
  const clickX=e.clientX-rect.left;
  const raw=Math.round((clickX/rect.width)*100);
  // snap to nearest 5
  const pct=Math.min(100,Math.max(0,Math.round(raw/5)*5));
  setProgress(id, pct);
}

function switchTab(name,btn){
  currentTab=name;
  document.querySelectorAll('.tabs .tab').forEach(t=>t.classList.remove('on'));
  if(btn) btn.classList.add('on');
  render();
}
function updateBottomNav(name){
  document.querySelectorAll('.nav-item').forEach(n=>n.classList.remove('on'));
  const el=document.getElementById('nav-'+name); if(el) el.classList.add('on');
  document.querySelectorAll('.tabs .tab').forEach((t,i)=>{
    const map=['today','tasks','meetings','reminders'];
    if(map[i]===name) t.classList.add('on'); else t.classList.remove('on');
  });
}
function toggleFab(){
  fabOpen=!fabOpen;
  document.getElementById('fab-menu').classList.toggle('open',fabOpen);
  document.getElementById('fab-btn').textContent=fabOpen?'×':'+';
}

function closeModal(id){ document.getElementById(id).classList.add('hidden'); }

function openAddTask(type){
  editTaskId=null;
  document.getElementById('task-modal-title').textContent='Add task';
  document.getElementById('t-title').value=''; document.getElementById('t-note').value='';
  document.getElementById('t-pri').value='med'; document.getElementById('t-due').value=todayStr;
  document.getElementById('t-remind').value=''; document.getElementById('t-cat').value='SME';
  setType(type||'work');
  document.getElementById('task-modal').classList.remove('hidden');
  setTimeout(()=>document.getElementById('t-title').focus(),100);
}
function editTask(id){
  const t=state.tasks.find(x=>x.id===id); if(!t) return;
  editTaskId=id; document.getElementById('task-modal-title').textContent='Edit task';
  setType(t.type);
  document.getElementById('t-title').value=t.title; document.getElementById('t-note').value=t.note||'';
  document.getElementById('t-pri').value=t.pri; document.getElementById('t-cat').value=t.cat;
  document.getElementById('t-due').value=t.due||''; document.getElementById('t-remind').value=t.remind||'';
  document.getElementById('task-modal').classList.remove('hidden');
}
function setType(type){
  currentType=type;
  document.getElementById('type-work').className='type-opt work'+(type==='work'?' sel':'');
  document.getElementById('type-personal').className='type-opt personal'+(type==='personal'?' sel':'');
}
function saveTask(){
  const title=document.getElementById('t-title').value.trim(); if(!title){document.getElementById('t-title').focus();return;}
  const existing = editTaskId ? state.tasks.find(x=>x.id===editTaskId) : null;
  const task={id:editTaskId||uid(),type:currentType,title,note:document.getElementById('t-note').value.trim(),
    pri:document.getElementById('t-pri').value,cat:document.getElementById('t-cat').value,
    due:document.getElementById('t-due').value,remind:document.getElementById('t-remind').value,
    done: existing ? existing.done||false : false,
    progress: existing ? existing.progress||0 : 0};
  if(editTaskId) state.tasks=state.tasks.map(x=>x.id===editTaskId?task:x); else state.tasks.push(task);
  saveState(); closeModal('task-modal'); render();
  showToast(editTaskId?'Task updated':'Task added ✅');
  if(task.due&&task.remind) scheduleLocalReminder(task);
}

function openAddMeeting(){
  document.getElementById('m-title').value=''; document.getElementById('m-date').value=todayStr;
  document.getElementById('m-time').value='09:00'; document.getElementById('m-site').value='all';
  document.getElementById('m-link').value=''; document.getElementById('m-note').value='';
  document.getElementById('mtg-modal').classList.remove('hidden');
  setTimeout(()=>document.getElementById('m-title').focus(),100);
}
function saveMeeting(){
  const title=document.getElementById('m-title').value.trim(); if(!title) return;
  state.meetings.push({id:uid(),title,date:document.getElementById('m-date').value,
    time:document.getElementById('m-time').value,site:document.getElementById('m-site').value,
    link:document.getElementById('m-link').value.trim(),note:document.getElementById('m-note').value.trim(),source:'manual'});
  saveState(); closeModal('mtg-modal'); render(); showToast('Meeting added 📅');
}

function openAddReminder(){
  document.getElementById('r-title').value=''; document.getElementById('r-date').value=todayStr;
  document.getElementById('r-time').value='08:00'; document.getElementById('r-repeat').value='none';
  document.getElementById('rem-modal').classList.remove('hidden');
  setTimeout(()=>document.getElementById('r-title').focus(),100);
}
function saveReminder(){
  const title=document.getElementById('r-title').value.trim(); if(!title) return;
  const r={id:uid(),title,date:document.getElementById('r-date').value,time:document.getElementById('r-time').value,repeat:document.getElementById('r-repeat').value};
  state.reminders.push(r); saveState(); closeModal('rem-modal'); render(); showToast('Reminder set ⏰');
  scheduleLocalReminder(r);
}

// ─── NOTIFICATIONS ────────────────────────────────────────────────────────────
function requestNotifPermission(){
  if(!('Notification' in window)){showToast('Notifications not supported here');return;}
  if(Notification.permission==='granted'){showToast('Notifications already enabled ✅');return;}
  Notification.requestPermission().then(p=>{
    showToast(p==='granted'?'Notifications enabled ✅':'Blocked — enable in browser settings');
    document.getElementById('notif-btn').classList.toggle('active',p==='granted');
  });
}
function scheduleLocalReminder(item){
  const d=item.due||item.date; const t=item.remind||item.time||'08:00';
  if(!d||!t) return;
  const dt=new Date(d+'T'+t); const ms=dt.getTime()-Date.now();
  if(ms>0&&ms<7*24*60*60*1000){
    setTimeout(()=>{
      if(Notification.permission==='granted') new Notification('MyTaskManager',{body:item.title});
      else showToast('⏰ '+item.title);
    },ms);
  }
}
function scheduleReminders(){
  state.tasks.filter(t=>!t.done&&t.due&&t.remind).forEach(scheduleLocalReminder);
  state.reminders.forEach(scheduleLocalReminder);
}

// ─── TOAST ────────────────────────────────────────────────────────────────────
let toastT;
function showToast(msg){
  const el=document.getElementById('toast'); el.textContent=msg; el.classList.remove('hidden');
  clearTimeout(toastT); toastT=setTimeout(()=>el.classList.add('hidden'),2800);
}

// ─── AI ───────────────────────────────────────────────────────────────────────
function openAIChat(){
  const el=document.getElementById('ai-messages');
  if(!el.children.length) addAIMsg('assistant',"Hi! I'm your AI assistant. I can help you prioritise tasks, plan your day, make decisions, or think through any challenge. What's on your mind?");
  renderAISuggestions();
  document.getElementById('ai-modal').classList.remove('hidden');
  setTimeout(()=>document.getElementById('ai-input').focus(),100);
}
function renderAISuggestions(){
  const over=state.tasks.filter(t=>!t.done&&isPast(t.due)).length;
  const sugs=["Help me prioritise today",over>0?`I have ${over} overdue tasks — what now?`:"What should I focus on this week?","Help me think through the Appium V3 decision","I'm overwhelmed — help me break this down"];
  const el=document.getElementById('ai-suggestions');
  if(el) el.innerHTML=sugs.map(s=>`<button class="ai-quick" onclick="quickAsk('${s.replace(/'/g,"&#39;")}')">${s}</button>`).join('');
}
function quickAsk(text){ document.getElementById('ai-input').value=text; sendAI(); }
function askAIAboutTask(id){
  const t=state.tasks.find(x=>x.id===id); if(!t) return;
  openAIChat();
  setTimeout(()=>{ document.getElementById('ai-input').value=`Help me plan this task: "${t.title}"${t.note?' ('+t.note+')':''}. Priority: ${t.pri}. Due: ${fmtDate(t.due)||'no date'}.`; sendAI(); },300);
}
function addAIMsg(role,text){
  const el=document.getElementById('ai-messages');
  const isUser=role==='user';
  const div=document.createElement('div');
  div.className='ai-msg-wrap'; div.style.alignItems=isUser?'flex-end':'flex-start';
  div.innerHTML=`<div class="ai-msg-label">${isUser?'You':'AI ✦'}</div><div class="ai-msg-bubble ${role}">${esc(text)}</div>`;
  el.appendChild(div); el.scrollTop=el.scrollHeight;
}
async function sendAI(){
  const input=document.getElementById('ai-input');
  const text=input.value.trim(); if(!text) return;
  input.value=''; addAIMsg('user',text);
  const think=document.createElement('div');
  think.style.cssText='display:flex;align-items:center;gap:8px;padding:6px 4px;color:var(--ai);font-size:12px';
  think.innerHTML='<div style="width:8px;height:8px;border-radius:50%;background:var(--ai);animation:pulse 1s ease-in-out infinite"></div> Thinking...';
  document.getElementById('ai-messages').appendChild(think);
  document.getElementById('ai-messages').scrollTop=9999;
  const openTasks=state.tasks.filter(t=>!t.done).slice(0,10).map(t=>`- [${t.type}/${t.pri}] ${t.title}${t.due?' (due '+fmtDate(t.due)+')':''}`).join('\n');
  const over=state.tasks.filter(t=>!t.done&&isPast(t.due));
  const todayMtgs=allMeetings().filter(m=>m.date===todayStr).map(m=>`${m.time} ${m.title}`).join(', ');
  const sys=`You are a smart personal task manager assistant for a Verification Manager at a medical device company, working across two sites: Barcelona and Switzerland. Today is ${todayStr}.

Open tasks (top 10):
${openTasks||'none'}
Overdue: ${over.length} (${over.map(t=>t.title).slice(0,3).join(', ')})
Today's meetings: ${todayMtgs||'none'}
Outlook connected: ${state.outlookUser?'Yes, '+state.outlookUser:'No'}

Be concise, practical, and direct. Use numbered steps when recommending actions.`;
  try{
    const apiKey = (state.anthropicKey||'').trim();
    if(!apiKey){
      think.remove();
      addAIMsg('assistant','⚠️ No API key found. Open ⚙️ Settings and enter your Anthropic API key (starts with sk-ant-...), then tap Save.');
      return;
    }
    const resp=await fetch('https://api.anthropic.com/v1/messages',{method:'POST',
      headers:{
        'Content-Type':'application/json',
        'x-api-key': apiKey,
        'anthropic-version':'2023-06-01',
        'anthropic-dangerous-allow-browser':'true'
      },
      body:JSON.stringify({model:'claude-sonnet-4-6',max_tokens:1000,system:sys,
        messages:[...state.aiHistory.slice(-6),{role:'user',content:text}]})});
    let data;
    try{ data = await resp.json(); }
    catch(parseErr){ think.remove(); addAIMsg('assistant','Could not read API response (status '+resp.status+'). The request may have been blocked before reaching Anthropic.'); return; }
    think.remove();
    if(!resp.ok || data.error){
      const code = resp.status;
      const msg = data.error?.message || resp.statusText || 'Unknown error';
      const type = data.error?.type || '';
      if(code===401||type==='authentication_error'){
        addAIMsg('assistant','❌ Authentication failed (401): '+msg+'\n\nYour API key may be invalid or expired. Open ⚙️ Settings, re-enter it, and Save.');
      } else if(code===403){
        addAIMsg('assistant','❌ Forbidden (403): '+msg+'\n\nThis can happen if the API key doesn\'t have access to model claude-sonnet-4-6, or billing isn\'t set up on your Anthropic account.');
      } else if(code===429){
        addAIMsg('assistant','⏳ Rate limited (429): '+msg+'\n\nWait a moment and try again, or check usage limits at console.anthropic.com.');
      } else if(code===0||!code){
        addAIMsg('assistant','❌ Request blocked (CORS/network). This usually means the page is being opened via file:// instead of http(s)://, or a browser extension/firewall is blocking api.anthropic.com.');
      } else {
        addAIMsg('assistant','❌ API error '+code+': '+msg);
      }
      return;
    }
    const reply=data.content?.[0]?.text||'Sorry, no response. Please try again.';
    addAIMsg('assistant',reply);
    state.aiHistory.push({role:'user',content:text},{role:'assistant',content:reply});
    if(state.aiHistory.length>20) state.aiHistory=state.aiHistory.slice(-16);
    saveState();
  }catch(e){
    think.remove();
    addAIMsg('assistant','❌ Network error: '+e.message+'\n\nMake sure: 1) the page URL starts with http:// or https:// (not file://), 2) you have internet access, 3) no extension is blocking api.anthropic.com.');
  }
}

// ─── INIT ─────────────────────────────────────────────────────────────────────
if(!checkStorageHealth()){
  setTimeout(()=>{
    const t = document.getElementById('toast');
    t.innerHTML = '⚠️ Your browser is blocking saved data (private/incognito mode or cookies disabled for this site). Tasks, settings and API key will NOT be saved between sessions. Use a normal browser window and allow site data. (Tap to dismiss)';
    t.classList.remove('hidden');
    t.style.maxWidth='90%';
    t.style.borderLeftColor='var(--high)';
    t.style.cursor='pointer';
    t.style.pointerEvents='auto';
    t.onclick=()=>t.classList.add('hidden');
  }, 600);
}
loadState().then(() => {
  if(state.outlookClientId){ initMsal(state.outlookClientId); }
  updateBannerState();
  render();
  if(Notification.permission==='granted') document.getElementById('notif-btn').classList.add('active');
  updateSettingsBtn();

  if(state.outlookClientId && msalInstance && state.outlookUser){
    setTimeout(fetchOutlookEvents, 1500);
  }
});
if(Notification.permission==='granted') document.getElementById('notif-btn').classList.add('active');
updateSettingsBtn();
// Auto-sync Outlook on load if connected
if(state.outlookClientId && msalInstance && state.outlookUser){
  setTimeout(fetchOutlookEvents, 1500);
}

function openSettings(){
  const k = state.anthropicKey||'';
  document.getElementById('s-apikey').value = k ? k.slice(0,8)+'...' : '';
  document.getElementById('s-appname').value = state.appName||'MyTaskManager';
  const status = document.getElementById('s-key-status');
  if(k){ status.textContent='✓ API key saved ('+k.slice(0,12)+'...)'; status.style.color='var(--personal)'; }
  else { status.textContent='No API key set — AI features disabled'; status.style.color='var(--high)'; }
  document.getElementById('settings-modal').classList.remove('hidden');
}
function saveSettings(){
  const raw = document.getElementById('s-apikey').value.trim();
  if(raw && !raw.includes('...')) state.anthropicKey = raw;
  const name = document.getElementById('s-appname').value.trim();
  if(name) state.appName = name;
  saveState();
  closeModal('settings-modal');
  showToast('Settings saved ✓');
  updateSettingsBtn();
}
function updateSettingsBtn(){
  const btn = document.getElementById('settings-btn');
  if(btn) btn.style.color = state.anthropicKey ? 'var(--personal)' : 'var(--text2)';
}
function clearAllData(){
  if(!confirm('Delete ALL tasks, meetings, reminders and settings? This cannot be undone.')) return;
  localStorage.removeItem('mytm_v2');
  location.reload();
}


// ─── SIDEBAR TOGGLE (visual/UX addition only, no logic change) ────────────────
function toggleSidebar(){
  document.getElementById('sidebar').classList.toggle('open');
  document.getElementById('sidebar-overlay').classList.toggle('show');
}
