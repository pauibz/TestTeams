/* ================================================================
   CONFIGURACIÓN
================================================================ */
const CONFIG = {
  clientId: 'c6219253-8f6e-48c9-8c4a-766aea58a874',
  tenantId: 'a3cdf8f7-40db-4d0f-8ccb-82310428392a',
};

const SCOPES = [
  'User.Read',
  'Calendars.Read',
  'Tasks.ReadWrite',
  'Mail.Read',
  'Presence.Read',
  'Presence.ReadWrite',
  'People.Read',
  'User.ReadBasic.All',
];

/* ================================================================
   TEAMS DETECTION — COMPATIBLE DESKTOP + WEB
================================================================ */
let _inTeams = false;
let _teamsContext = null;
let _teamsInitialized = false;

async function initTeams() {
  return new Promise(resolve => {
    // Hard timeout: if Teams SDK doesn't respond in 2.5s → browser mode
    const timeout = setTimeout(() => {
      console.log('[Teams] Timeout — running in browser');
      resolve();
    }, 2500);

    try {
      microsoftTeams.app.initialize().then(() => {
        clearTimeout(timeout);
        _inTeams = true;
        _teamsInitialized = true;
        microsoftTeams.app.getContext().then(ctx => {
          _teamsContext = ctx;
          // Apply Teams desktop class for CSS tweaks
          if (ctx?.app?.host?.name === 'Teams') {
            document.body.classList.add('teams-desktop');
          }
        }).catch(() => {});
        console.log('[Teams] Running inside Teams ✅');
        resolve();
      }).catch(() => {
        clearTimeout(timeout);
        console.log('[Teams] Not in Teams — browser mode');
        resolve();
      });
    } catch {
      clearTimeout(timeout);
      resolve();
    }
  });
}

/* ================================================================
   MSAL (browser)
================================================================ */
const msalInstance = new msal.PublicClientApplication({
  auth: {
    clientId: CONFIG.clientId,
    authority: `https://login.microsoftonline.com/${CONFIG.tenantId}`,
    redirectUri: window.location.origin + window.location.pathname,
  },
  cache: { cacheLocation: 'sessionStorage', storeAuthStateInCookie: true },
});

let currentAccount = null;

/* ================================================================
   TOKEN ACQUISITION
   Teams Desktop: getAuthToken → OBO flow would be needed for Graph
   For direct Graph calls we fall back to MSAL popup/silent
================================================================ */
async function getToken() {
  // If in Teams try silent MSAL first (works in web), then Teams SSO
  if (_inTeams) {
    try {
      // Try to get a Graph token via MSAL silently (Teams web supports this)
      const accounts = msalInstance.getAllAccounts();
      if (accounts.length) {
        currentAccount = accounts[0];
        const r = await msalInstance.acquireTokenSilent({ scopes: SCOPES, account: currentAccount });
        return r.accessToken;
      }
    } catch { /* fall through */ }

    // Teams SSO: gives id-token, needs OBO for Graph
    // Here we do a popup to ensure Graph scopes
    try {
      const r = await msalInstance.acquireTokenPopup({ scopes: SCOPES });
      currentAccount = r.account;
      return r.accessToken;
    } catch (e) {
      console.warn('[Teams] acquireTokenPopup failed:', e);
      throw e;
    }
  }

  // Standard MSAL (browser)
  const accounts = msalInstance.getAllAccounts();
  if (!accounts.length) throw new Error('Sin sesión activa');
  currentAccount = accounts[0];
  try {
    const r = await msalInstance.acquireTokenSilent({ scopes: SCOPES, account: currentAccount });
    return r.accessToken;
  } catch {
    const r = await msalInstance.acquireTokenPopup({ scopes: SCOPES, account: currentAccount });
    return r.accessToken;
  }
}

async function callGraph(path, token, options = {}) {
  const res = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
    method: options.method || 'GET',
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json',
    },
    body: options.body ? JSON.stringify(options.body) : undefined,
  });
  if (options.raw) return res;
  if (!res.ok) {
    const err = await res.json().catch(() => ({}));
    throw new Error(`Graph ${res.status}: ${err?.error?.message || path}`);
  }
  if (res.status === 204) return null;
  return res.json();
}

/* ================================================================
   LOGIN / LOGOUT
================================================================ */
async function doLogin() {
  try {
    showLoading('Conectando con Microsoft…');
    if (_inTeams) {
      showApp();
      await loadAllData();
    } else {
      const r = await msalInstance.loginPopup({ scopes: SCOPES });
      currentAccount = r.account;
      showApp();
      await loadAllData();
    }
  } catch (e) {
    hideLoading();
    showError('Error al iniciar sesión: ' + (e.message || 'Inténtalo de nuevo.'));
  }
}

function enterDemo() {
  showApp();
  loadMockData();
}

function doLogout() {
  closeUserMenu();
  if (_inTeams) {
    document.getElementById('appShell').style.display = 'none';
    document.getElementById('loginScreen').style.display = 'flex';
  } else {
    msalInstance.logoutPopup();
  }
}

/* ================================================================
   LOAD ALL DATA
================================================================ */
async function loadAllData() {
  const icon = document.getElementById('refreshIcon');
  icon.classList.add('spinning');
  showLoading('Actualizando desde Microsoft 365…');

  try {
    const token = await getToken();

    const me = await callGraph('/me', token);
    updateUserUI(me);
    loadUserPhoto(me, token);

    const today = new Date().toISOString().split('T')[0];
    const startDT = `${today}T00:00:00`;
    const endDT   = `${today}T23:59:59`;

    const [calRes, tasksRes, mailRes, presRes, peopleRes] = await Promise.allSettled([
      callGraph(
        `/me/calendarView?startDateTime=${startDT}&endDateTime=${endDT}` +
        `&$orderby=start/dateTime&$top=20` +
        `&$select=subject,organizer,start,end,onlineMeeting,location,bodyPreview`,
        token
      ),
      callGraph('/me/planner/tasks?$filter=percentComplete ne 100&$top=20', token),
      callGraph(
        '/me/messages?$top=8&$orderby=receivedDateTime desc' +
        '&$select=subject,from,isRead,bodyPreview,receivedDateTime,webLink,importance',
        token
      ),
      callGraph('/me/presence', token),
      callGraph('/me/people?$top=10&$select=displayName,jobTitle,userPrincipalName,scoredEmailAddresses', token),
    ]);

    const meetings = calRes.status     === 'fulfilled' ? (calRes.value?.value ?? [])    : [];
    const tasks    = tasksRes.status   === 'fulfilled' ? (tasksRes.value?.value ?? [])  : [];
    const mails    = mailRes.status    === 'fulfilled' ? (mailRes.value?.value ?? [])   : [];
    const presence = presRes.status    === 'fulfilled' ? presRes.value                  : null;
    const people   = peopleRes.status  === 'fulfilled' ? (peopleRes.value?.value ?? []) : [];

    renderMeetings(meetings);
    renderTasks(tasks);
    await renderMail(mails, token);
    if (presence) {
      applyPresence(presence.availability);
      syncMyPresenceUI(presence.availability);
    }
    updateKPIs(meetings.length, tasks, mails);
    scheduleMeetingAlerts(meetings);
    loadTeamPresence(people, token);

  } catch (err) {
    console.error('Error cargando datos:', err);
    showToast('Sin conexión — mostrando datos de ejemplo');
    loadMockData();
  }

  document.getElementById('refreshIcon').classList.remove('spinning');
  hideLoading();
  updateLastRefreshTime();
}

/* ================================================================
   AVATAR / PHOTOS
================================================================ */
async function loadUserPhoto(me, token) {
  try {
    const res = await callGraph('/me/photo/$value', token, { raw: true });
    if (!res.ok) return;
    const blob = await res.blob();
    const url  = URL.createObjectURL(blob);
    setAvatarEl('navAvatar',  url, me.displayName);
    setAvatarEl('mppAvatar',  url, me.displayName);
  } catch { /* use initials */ }
}

function setAvatarEl(wrapperId, photoUrl, name) {
  const wrap = document.getElementById(wrapperId);
  if (!wrap) return;
  const ini = wrap.querySelector('span');
  if (ini) ini.style.display = 'none';
  let img = wrap.querySelector('img');
  if (!img) { img = document.createElement('img'); img.alt = name; wrap.appendChild(img); }
  img.src = photoUrl;
}

async function loadContactPhoto(email, token) {
  try {
    const res = await fetch(
      `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(email)}/photo/$value`,
      { headers: { Authorization: `Bearer ${token}` } }
    );
    if (!res.ok) return null;
    return URL.createObjectURL(await res.blob());
  } catch { return null; }
}

/* ================================================================
   TEAM PRESENCE
================================================================ */
let _teamData = [];
let _teamFilter = 'all';
let _teamSearch = '';

async function loadTeamPresence(people, token) {
  if (!people.length) {
    renderTeamGrid([]);
    return;
  }

  // Get user IDs
  const userIds = [];
  const userMap = {};
  for (const p of people) {
    const email = p.scoredEmailAddresses?.[0]?.address;
    if (!email) continue;
    try {
      const u = await callGraph(`/users/${encodeURIComponent(email)}?$select=id,displayName,jobTitle`, token);
      if (u?.id) { userIds.push(u.id); userMap[u.id] = { ...p, ...u }; }
    } catch { /* skip */ }
  }

  if (!userIds.length) { renderTeamGrid([]); return; }

  // Batch presence
  let presences = [];
  try {
    const pr = await callGraph('/communications/getPresencesByUserId', token, {
      method: 'POST',
      body: { ids: userIds.slice(0, 15) },
    });
    presences = pr?.value ?? [];
  } catch {
    // Fallback: individual presence
    for (const id of userIds.slice(0, 8)) {
      try {
        const p = await callGraph(`/users/${id}/presence`, token);
        if (p) presences.push(p);
      } catch { /* skip */ }
    }
  }

  _teamData = presences.map(pr => ({
    id: pr.id,
    name: userMap[pr.id]?.displayName || 'Desconocido',
    role: userMap[pr.id]?.jobTitle || '',
    email: userMap[pr.id]?.mail || '',
    availability: pr.availability || 'PresenceUnknown',
    activity: pr.activity || '',
  }));

  renderTeamGrid(_teamData);

  // Load team photos in background
  loadTeamPhotos(token);
}

async function loadTeamPhotos(token) {
  for (const m of _teamData) {
    if (!m.email) continue;
    const url = await loadContactPhoto(m.email, token);
    if (!url) continue;
    const el = document.getElementById(`tc-avatar-${m.id}`);
    if (el) el.innerHTML = `<img src="${url}" alt="${escHtml(m.name)}" style="width:100%;height:100%;object-fit:cover;border-radius:50%" />`;
  }
}

function renderTeamGrid(data) {
  const el = document.getElementById('teamGrid');
  let filtered = [...data];

  if (_teamFilter !== 'all') {
    filtered = filtered.filter(m => m.availability === _teamFilter ||
      ((_teamFilter === 'Away') && ['Away','BeRightBack'].includes(m.availability)) ||
      ((_teamFilter === 'Busy') && ['Busy','DoNotDisturb'].includes(m.availability)) ||
      ((_teamFilter === 'Offline') && ['Offline','PresenceUnknown'].includes(m.availability))
    );
  }

  if (_teamSearch) {
    const q = _teamSearch.toLowerCase();
    filtered = filtered.filter(m => m.name.toLowerCase().includes(q) || m.role.toLowerCase().includes(q));
  }

  if (!filtered.length) {
    el.innerHTML = '<div class="empty-state" style="grid-column:1/-1">No hay compañeros con este filtro</div>';
    return;
  }

  el.innerHTML = filtered.map(m => {
    const p    = PMAP[m.availability] || PMAP.PresenceUnknown;
    const ini  = m.name.split(' ').map(w => w[0]).join('').slice(0,2).toUpperCase();
    const bcls = p.badge;
    const activityStr = m.activity ? `· ${m.activity}` : '';
    return `
      <div class="team-card">
        <div class="tc-avatar-wrap">
          <div class="tc-avatar" id="tc-avatar-${m.id}" style="background:${p.avatarColor}">${ini}</div>
          <div class="tc-status-dot" style="background:${p.color}"></div>
        </div>
        <div class="tc-name">${escHtml(m.name)}</div>
        <div class="tc-role">${escHtml(m.role)}</div>
        <div class="tc-badge ${bcls}">${p.label}</div>
        <div class="tc-activity">${escHtml(activityStr)}</div>
      </div>`;
  }).join('');
}

function setFilter(filter, el) {
  _teamFilter = filter;
  document.querySelectorAll('.filter-chip').forEach(c => c.classList.remove('active'));
  el.classList.add('active');
  renderTeamGrid(_teamData);
}

function filterTeam(q) {
  _teamSearch = q;
  renderTeamGrid(_teamData);
}

/* ================================================================
   PRESENCE MAP
================================================================ */
const PMAP = {
  Available:       { color: '#6BB700', label: 'Disponible', badge: 'available', avatarColor: '#107C10' },
  Busy:            { color: '#C4314B', label: 'Ocupado/a',  badge: 'busy',      avatarColor: '#A4262C' },
  DoNotDisturb:    { color: '#C4314B', label: 'No molestar',badge: 'busy',      avatarColor: '#A4262C' },
  Away:            { color: '#F4BE22', label: 'Ausente',    badge: 'away',      avatarColor: '#8A5A00' },
  BeRightBack:     { color: '#F4BE22', label: 'Ahora vuelvo',badge:'away',      avatarColor: '#8A5A00' },
  Offline:         { color: '#666',    label: 'Sin conexión',badge: 'offline',  avatarColor: '#444' },
  PresenceUnknown: { color: '#666',    label: 'Desconocido', badge: 'offline',  avatarColor: '#444' },
};

function applyPresence(availability) {
  const p = PMAP[availability] || PMAP.PresenceUnknown;
  // Nav dot
  const navDot = document.getElementById('navPresenceDot');
  if (navDot) navDot.style.background = p.color;
  // Presence page dot
  const mppDot = document.getElementById('mppStatusDot');
  if (mppDot) mppDot.style.background = p.color;
  // Presence page text
  const mppStatus = document.getElementById('mppStatus');
  if (mppStatus) mppStatus.textContent = PRESENCE_LABELS[availability] || '⚫ Desconocido';
}

const PRESENCE_LABELS = {
  Available:    '🟢 Disponible',
  Busy:         '🔴 Ocupado/a',
  DoNotDisturb: '🔴 No molestar',
  Away:         '🟡 Ausente',
  BeRightBack:  '🟡 Ahora vuelvo',
  Offline:      '⚫ Sin conexión',
  PresenceUnknown: '⚫ Desconocido',
};

function syncMyPresenceUI(availability) {
  const p = PMAP[availability] || PMAP.PresenceUnknown;
  const mppDot = document.getElementById('mppStatusDot');
  if (mppDot) mppDot.style.background = p.color;
}

async function setMyPresence(el, availability, color, label) {
  document.querySelectorAll('.presence-opt').forEach(o => o.classList.remove('active'));
  el.classList.add('active');
  const navDot = document.getElementById('navPresenceDot');
  if (navDot) navDot.style.background = color;
  const mppDot = document.getElementById('mppStatusDot');
  if (mppDot) mppDot.style.background = color;
  const mppStatus = document.getElementById('mppStatus');
  if (mppStatus) mppStatus.textContent = label;

  // Try to update via Graph (requires Presence.ReadWrite delegated)
  try {
    const token = await getToken();
    await callGraph('/me/presence/setStatusMessage', token, {
      method: 'POST',
      body: { statusMessage: { message: { content: label, contentType: 'text' } } }
    });
    showToast(`Estado actualizado: ${label}`);
  } catch {
    showToast('Estado actualizado localmente (sin permisos de escritura)');
  }
}

/* ================================================================
   MOCK DATA
================================================================ */
function todayAt(h, m) {
  const d = new Date(); d.setHours(h, m, 0, 0); return d.toISOString();
}

const MOCK_TEAM = [
  { id: 't1', name: 'Laura Sánchez', role: 'Product Designer',         email: 'laura@demo.com',   availability: 'Available' },
  { id: 't2', name: 'Pedro Ruiz',    role: 'Frontend Developer',       email: 'pedro@demo.com',   availability: 'Busy' },
  { id: 't3', name: 'Ana Martínez',  role: 'UX Researcher',            email: 'ana@demo.com',     availability: 'Away' },
  { id: 't4', name: 'Carlos López',  role: 'Backend Developer',        email: 'carlos@demo.com',  availability: 'Available' },
  { id: 't5', name: 'Marta Gil',     role: 'Project Manager',          email: 'marta@demo.com',   availability: 'DoNotDisturb' },
  { id: 't6', name: 'Jorge Pérez',   role: 'Data Engineer',            email: 'jorge@demo.com',   availability: 'Offline' },
  { id: 't7', name: 'Sofía Torres',  role: 'QA Engineer',              email: 'sofia@demo.com',   availability: 'BeRightBack' },
  { id: 't8', name: 'Iván Castro',   role: 'DevOps',                   email: 'ivan@demo.com',    availability: 'Available' },
];

function loadMockData() {
  const meetings = [
    { subject: 'Daily Stand-up · Equipo Digital',     organizer: { emailAddress: { name: 'Pedro Ruiz' } },    start: { dateTime: todayAt(9,0)   }, end: { dateTime: todayAt(9,30)  }, onlineMeeting: { joinUrl: '#' } },
    { subject: 'Revisión de diseño Q2 · App móvil',   organizer: { emailAddress: { name: 'Laura Sánchez' } }, start: { dateTime: todayAt(10,30) }, end: { dateTime: todayAt(11,30) }, onlineMeeting: { joinUrl: '#' } },
    { subject: 'Sync con cliente · Proyecto Atenea',  organizer: { emailAddress: { name: 'Carlos López' } },  start: { dateTime: todayAt(14,0)  }, end: { dateTime: todayAt(14,45) }, onlineMeeting: { joinUrl: '#' } },
    { subject: 'Retrospectiva Sprint 14',              organizer: { emailAddress: { name: 'Ana Martínez' } },  start: { dateTime: todayAt(16,0)  }, end: { dateTime: todayAt(17,0)  }, onlineMeeting: { joinUrl: '#' } },
  ];
  const tasks = [
    { id:'1', title: 'Revisar mockups pantalla de inicio',      percentComplete: 0,   priority: 1, dueDateTime: { dateTime: new Date().toISOString() } },
    { id:'2', title: 'Actualizar documentación de componentes', percentComplete: 0,   priority: 5, dueDateTime: { dateTime: new Date().toISOString() } },
    { id:'3', title: 'Preparar agenda reunión con cliente',     percentComplete: 100, priority: 1, dueDateTime: null },
    { id:'4', title: 'Enviar informe de accesibilidad',         percentComplete: 0,   priority: 5, dueDateTime: { dateTime: new Date().toISOString() } },
    { id:'5', title: 'Pruebas de usabilidad – flujo checkout',  percentComplete: 0,   priority: 9, dueDateTime: null },
    { id:'6', title: 'Responder feedback de Pedro sobre el DS', percentComplete: 0,   priority: 5, dueDateTime: { dateTime: new Date().toISOString() } },
    { id:'7', title: 'Cerrar issues del Sprint 13',             percentComplete: 100, priority: 9, dueDateTime: null },
    { id:'8', title: 'Kick-off Proyecto Hermes',                percentComplete: 0,   priority: 1, dueDateTime: { dateTime: new Date().toISOString() } },
  ];
  const mails = [
    { from: { emailAddress: { name: 'Laura Sánchez' } }, subject: 'Re: Revisión de nuevos componentes del DS',   isRead: false, importance:'high', bodyPreview: 'He revisado los últimos cambios y me parecen muy bien…', receivedDateTime: new Date().toISOString(), webLink: '#' },
    { from: { emailAddress: { name: 'Pedro Ruiz' } },    subject: 'Feedback sprint review – acciones pendientes', isRead: false, importance:'normal', bodyPreview: 'Tras la reunión del viernes quedan estos puntos abiertos…', receivedDateTime: new Date(Date.now()-3e6).toISOString(), webLink: '#' },
    { from: { emailAddress: { name: 'Ana Martínez' } },  subject: 'Retrospectiva hoy 16:00 – agenda adjunta',    isRead: false, importance:'normal', bodyPreview: 'Os comparto la agenda para que podáis prepararlo con tiempo…', receivedDateTime: new Date(Date.now()-5e6).toISOString(), webLink: '#' },
    { from: { emailAddress: { name: 'Carlos López' } },  subject: 'Confirmación reunión cliente Atenea',         isRead: true,  importance:'normal', bodyPreview: 'Confirmo la reunión con Inmobiliaria Atenea para las 14:00.', receivedDateTime: new Date(Date.now()-86400000).toISOString(), webLink: '#' },
    { from: { emailAddress: { name: 'RRHH · Empresa' } },subject: 'Política de teletrabajo actualizada',         isRead: false, importance:'normal', bodyPreview: 'A partir del próximo mes entra en vigor la nueva política…', receivedDateTime: new Date(Date.now()-90000000).toISOString(), webLink: '#' },
  ];

  updateUserUI({ displayName: 'Demo User', jobTitle: 'Modo demostración', mail: 'demo@tuempresa.com' });
  renderMeetings(meetings);
  renderTasks(tasks);
  renderMail(mails, null);
  applyPresence('Available');
  syncMyPresenceUI('Available');
  updateKPIs(meetings.length, tasks, mails);
  scheduleMeetingAlerts(meetings);
  _teamData = MOCK_TEAM.map(m => ({ ...m, activity: '' }));
  renderTeamGrid(_teamData);
  hideLoading();
  updateLastRefreshTime();
  pushNotification('info', 'Modo demo', 'Estás viendo datos de ejemplo');
}

/* ================================================================
   RENDER — MEETINGS TIMELINE
================================================================ */
const pad = n => String(n).padStart(2, '0');

function renderMeetings(events) {
  const el = document.getElementById('meetingsList');
  if (!events?.length) {
    el.innerHTML = '<div class="empty-state">🎉 No tienes reuniones hoy</div>';
    setEl('badgeMeetings', '0');
    return;
  }
  setEl('badgeMeetings', events.length);
  const now = Date.now();
  el.innerHTML = '<div class="timeline" style="padding:0">' + events.map((e, idx) => {
    const rawS = e.start.dateTime;
    const rawE = e.end.dateTime;
    const s    = new Date(rawS.endsWith('Z') ? rawS : rawS + 'Z');
    const en   = new Date(rawE.endsWith('Z') ? rawE : rawE + 'Z');
    const live = now >= s && now < en;
    const past = now >= en;
    const dur  = Math.round((en - s) / 60000);
    const url  = e.onlineMeeting?.joinUrl || '#';
    const org  = e.organizer?.emailAddress?.name || '';
    const dotCls  = live ? 'live' : past ? 'past' : 'future';
    const cardCls = live ? 'live' : past ? 'past' : '';
    const isLast  = idx === events.length - 1;
    return `
      <div class="tl-item">
        <div class="tl-left">
          <div class="tl-hour">${pad(s.getHours())}:${pad(s.getMinutes())}</div>
          <div class="tl-min">${pad(en.getHours())}:${pad(en.getMinutes())}</div>
        </div>
        <div class="tl-marker">
          <div class="tl-dot ${dotCls}"></div>
          <div class="tl-connector${isLast ? ' last' : ''}"></div>
        </div>
        <div class="tl-body">
          <div class="tl-card ${cardCls}">
            <div class="tl-card-left">
              <div class="tl-name">${escHtml(e.subject)}</div>
              <div class="tl-meta">${escHtml(org)}${org ? ' · ' : ''}${dur} min</div>
              <div class="tl-chips">
                ${live ? '<span class="chip live">● En directo</span>' : ''}
                <span class="chip">${dur} min</span>
              </div>
            </div>
            <button class="join-btn${past ? ' past' : ''}"
              ${past ? 'disabled' : `onclick="window.open('${escHtml(url)}','_blank')"`}>
              ${past ? 'Finalizada' : 'Unirse'}
            </button>
          </div>
        </div>
      </div>`;
  }).join('') + '</div>';
}

/* ================================================================
   RENDER — TASKS
================================================================ */
let _tasks = [];

function renderTasks(list) {
  _tasks = list.map(t => ({ ...t, _done: t.percentComplete === 100 }));
  drawTasks();
}

function drawTasks() {
  const el = document.getElementById('tasksList');
  setEl('badgeTasks', _tasks.length);
  if (!_tasks.length) {
    el.innerHTML = '<div class="empty-state">🎉 Sin tareas pendientes</div>';
    return;
  }
  el.innerHTML = _tasks.map((t, i) => {
    const pr  = t.priority <= 2 ? 'high' : t.priority <= 5 ? 'medium' : 'low';
    const due = t.dueDateTime
      ? new Date(t.dueDateTime.dateTime).toLocaleDateString('es', { day: '2-digit', month: 'short' })
      : 'Sin fecha';
    return `
      <div class="task-item" onclick="toggleTask(${i})">
        <div class="task-check${t._done ? ' done' : ''}"></div>
        <div class="task-info">
          <div class="task-name${t._done ? ' done' : ''}">${escHtml(t.title)}</div>
          <div class="task-meta">Vence: ${due}</div>
        </div>
        <div class="priority-dot ${pr}" title="Prioridad ${pr}"></div>
      </div>`;
  }).join('');

  const done = _tasks.filter(t => t._done).length;
  const pend = _tasks.filter(t => !t._done).length;
  setEl('kpiDone', done);
  setEl('kpiPending', pend);
  setEl('statDone', done);
  setEl('statPend', pend);
}

function toggleTask(i) {
  _tasks[i]._done = !_tasks[i]._done;
  drawTasks();
  showToast(_tasks[i]._done ? '✅ Tarea marcada como completada' : '↩ Tarea desmarcada');
}

/* ================================================================
   RENDER — MAIL
================================================================ */
const AVATAR_COLORS = ['#6264A7','#0078D4','#107C10','#C4314B','#8A5A00','#007A7C','#8764B8'];

async function renderMail(msgs, token) {
  const el = document.getElementById('mailList');
  const unread = msgs?.filter(m => !m.isRead).length ?? 0;
  setEl('badgeMail', unread);
  setEl('kpiMail', unread);

  if (!msgs?.length) {
    el.innerHTML = '<div class="empty-state">No hay correos recientes</div>';
    return;
  }

  el.innerHTML = msgs.map((m, i) => {
    const name  = m.from?.emailAddress?.name || 'Desconocido';
    const email = m.from?.emailAddress?.address || '';
    const ini   = name.split(' ').map(w => w[0]).join('').slice(0,2).toUpperCase();
    const color = AVATAR_COLORS[i % AVATAR_COLORS.length];
    const rec   = new Date(m.receivedDateTime);
    const isToday = rec.toDateString() === new Date().toDateString();
    const timeStr = isToday
      ? rec.toLocaleTimeString('es', { hour:'2-digit', minute:'2-digit' })
      : rec.toLocaleDateString('es', { day:'2-digit', month:'short' });
    const isUrgent = m.importance === 'high';
    return `
      <div class="mail-item" onclick="window.open('${m.webLink || '#'}','_blank')">
        <div class="mail-avatar" id="mailAvatar-${i}" style="background:${color}">${ini}</div>
        <div class="mail-content">
          <div class="mail-from">${escHtml(name)} ${isUrgent ? '🔴' : ''}</div>
          <div class="mail-subject${!m.isRead ? ' unread' : ''}">${escHtml(m.subject)}</div>
          <div class="mail-preview">${escHtml(m.bodyPreview)}</div>
        </div>
        <div class="mail-right">
          <span class="mail-time">${timeStr}</span>
          ${!m.isRead ? '<div class="unread-dot"></div>' : ''}
        </div>
      </div>`;
  }).join('');

  if (!token) return;
  msgs.forEach(async (m, i) => {
    const email = m.from?.emailAddress?.address;
    if (!email) return;
    const photoUrl = await loadContactPhoto(email, token);
    if (!photoUrl) return;
    const avatarEl = document.getElementById(`mailAvatar-${i}`);
    if (avatarEl) avatarEl.innerHTML = `<img src="${photoUrl}" alt="${escHtml(m.from?.emailAddress?.name||'')}" />`;
  });
}

/* ================================================================
   KPIs
================================================================ */
function updateKPIs(meetN, taskList, mails) {
  const done   = taskList.filter(t => t.percentComplete === 100).length;
  const pend   = taskList.filter(t => t.percentComplete < 100).length;
  const unread = mails.filter(m => !m.isRead).length;
  const totalMin = meetN * 45;
  const timeStr  = totalMin >= 60
    ? `${Math.floor(totalMin/60)}h ${totalMin%60 ? totalMin%60+'m' : ''}`
    : `${totalMin}m`;

  setEl('kpiMeetings', meetN);
  setEl('kpiDone',     done);
  setEl('kpiPending',  pend);
  setEl('kpiMail',     unread);
  setEl('badgeMeetings', meetN);
  setEl('badgeTasks',  taskList.length);
  setEl('badgeMail',   unread);
  setEl('statMeetTime',meetN ? timeStr : '0m');
  setEl('statDone',    done);
  setEl('statPend',    pend);
  setEl('statUnread',  unread);
  setEl('greetingSub', `Tienes ${meetN} reuniones hoy y ${pend} tareas pendientes.`);
}

/* ================================================================
   SMART ALERTS — upcoming meeting notifications
================================================================ */
const _alertTimers = [];

function scheduleMeetingAlerts(meetings) {
  _alertTimers.forEach(clearTimeout);
  _alertTimers.length = 0;

  const now = Date.now();
  for (const m of meetings) {
    const rawS = m.start.dateTime;
    const s    = new Date(rawS.endsWith('Z') ? rawS : rawS + 'Z');
    const msUntil = s - now;
    if (msUntil > 0 && msUntil < 30 * 60 * 1000) {
      const minLeft = Math.round(msUntil / 60000);
      // Show alert now for meetings in next 30 min
      showMeetingAlert(m.subject, minLeft, m.onlineMeeting?.joinUrl);
      pushNotification('warning', `Reunión en ${minLeft} min`, m.subject);
    } else if (msUntil >= 14.5 * 60 * 1000 && msUntil <= 15.5 * 60 * 1000) {
      const t = setTimeout(() => {
        showMeetingAlert(m.subject, 15, m.onlineMeeting?.joinUrl);
        pushNotification('warning', 'Reunión en 15 min', m.subject);
      }, msUntil - 15 * 60 * 1000);
      _alertTimers.push(t);
    }
  }
}

function showMeetingAlert(subject, minLeft, url) {
  const c = document.getElementById('alertBannerContainer');
  if (!c) return;
  const id = `alert-${Date.now()}`;
  const div = document.createElement('div');
  div.id = id;
  div.className = 'alert-banner';
  div.innerHTML = `
    <div class="alert-banner-left">
      <span style="font-size:20px">📅</span>
      <div class="alert-banner-text">
        <strong>Reunión en ${minLeft} min</strong>
        <span>${escHtml(subject)}</span>
      </div>
    </div>
    <div style="display:flex;gap:8px;align-items:center">
      ${url && url !== '#' ? `<button class="join-btn" onclick="window.open('${escHtml(url)}','_blank')">Unirse</button>` : ''}
      <button class="alert-dismiss" onclick="document.getElementById('${id}').remove()">×</button>
    </div>`;
  c.appendChild(div);
  setTimeout(() => div.remove(), 60000);
}

/* ================================================================
   NOTIFICATIONS
================================================================ */
let _notifications = [];

function pushNotification(type, title, desc) {
  _notifications.unshift({ type, title, desc, time: new Date() });
  updateNotifBadge();
  renderNotifList();
}

function updateNotifBadge() {
  const badge = document.getElementById('badgeNotif');
  if (!badge) return;
  const count = _notifications.length;
  badge.style.display = count ? 'flex' : 'none';
  badge.textContent = count > 9 ? '9+' : count;
}

function renderNotifList() {
  const el = document.getElementById('notifList');
  if (!el) return;
  if (!_notifications.length) {
    el.innerHTML = '<div class="empty-state">Sin notificaciones</div>';
    return;
  }
  el.innerHTML = _notifications.map(n => {
    const t = n.time.toLocaleTimeString('es', { hour:'2-digit', minute:'2-digit' });
    const badgeMap = { warning: 'warning Próximamente', urgent: 'urgent Urgente', info: 'info Info' };
    return `
      <div class="notif-item">
        <div class="notif-badge ${n.type}">${badgeMap[n.type] || n.type}</div>
        <div class="notif-item-header">
          <div class="notif-title">${escHtml(n.title)}</div>
          <div class="notif-time">${t}</div>
        </div>
        <div class="notif-desc">${escHtml(n.desc)}</div>
      </div>`;
  }).join('');
}

function toggleNotifPanel() {
  const panel  = document.getElementById('notifPanel');
  const dimmer = document.getElementById('dimmer');
  const isOpen = panel.classList.contains('open');
  if (isOpen) {
    panel.classList.remove('open');
    dimmer.classList.remove('show');
  } else {
    panel.classList.add('open');
    dimmer.classList.add('show');
    // Clear badge when opened
    _notifications = [];
    updateNotifBadge();
  }
}

function closeNotifPanel() {
  document.getElementById('notifPanel').classList.remove('open');
  document.getElementById('dimmer').classList.remove('show');
}

/* ================================================================
   DAILY SUMMARY (auto-generated)
================================================================ */
function generateSummary() {
  const el = document.getElementById('summaryContent');
  const meetings = document.querySelectorAll('.tl-card').length;
  const pending  = parseInt(getEl('kpiPending')) || 0;
  const done     = parseInt(getEl('kpiDone')) || 0;
  const unread   = parseInt(getEl('kpiMail')) || 0;

  const hour = new Date().getHours();
  const greet = hour < 12 ? 'esta mañana' : hour < 18 ? 'esta tarde' : 'esta noche';

  el.innerHTML = `
    <div style="display:flex;flex-direction:column;gap:8px">
      <div>📅 <strong>Reuniones:</strong> Tienes ${meetings} reuniones ${greet}. Gestiona bien tu tiempo entre ellas.</div>
      <div>🗂 <strong>Tareas:</strong> ${pending} tareas pendientes, ${done} ya completadas. ${pending > 3 ? '¡Buen trabajo en progreso!' : 'Al ritmo actual terminarás antes de las 17:00.'}</div>
      <div>📧 <strong>Correos:</strong> ${unread} mensajes sin leer. ${unread > 3 ? 'Reserva 20 min para responder los más urgentes.' : 'Bandeja de entrada bajo control.'}</div>
      <div style="margin-top:6px;padding:10px;background:var(--brand-l);border-radius:8px;font-size:12px;color:var(--ts)">
        🤖 <em>Consejo del día:</em> Agrupa las respuestas de correo en bloques de 20-30 min para mantener el foco en tareas de alto valor.
      </div>
    </div>`;
  showToast('Resumen generado ✨');
}

/* ================================================================
   POWER BI
================================================================ */
function loadPowerBI() {
  const url = document.getElementById('powerbiUrl')?.value?.trim();
  if (!url) { showToast('Introduce una URL válida de Power BI'); return; }
  document.getElementById('powerbiContent').innerHTML =
    `<iframe id="powerbiFrame" src="${escHtml(url)}" allowfullscreen></iframe>`;
  showToast('Panel de Power BI cargado');
}

function togglePowerBI() {
  const c = document.getElementById('powerbiContent');
  const iframe = document.getElementById('powerbiFrame');
  if (iframe) {
    c.innerHTML = `<div class="powerbi-body"><div class="powerbi-icon">📊</div><div class="powerbi-msg">Pega la URL de tu informe de Power BI embebido para visualizarlo.</div><div class="powerbi-input"><input type="text" id="powerbiUrl" placeholder="https://app.powerbi.com/reportEmbed?reportId=…" /><button onclick="loadPowerBI()">Cargar</button></div></div>`;
  }
}

/* ================================================================
   GLOBAL SEARCH (simple filter)
================================================================ */
function handleSearch(q) {
  if (!q) return;
  // Simple: filter tasks
  const lower = q.toLowerCase();
  const taskItems = document.querySelectorAll('.task-item');
  taskItems.forEach(item => {
    const name = item.querySelector('.task-name')?.textContent?.toLowerCase() || '';
    item.style.display = name.includes(lower) ? '' : 'none';
  });
  const mailItems = document.querySelectorAll('.mail-item');
  mailItems.forEach(item => {
    const text = item.querySelector('.mail-subject')?.textContent?.toLowerCase() || '';
    const from = item.querySelector('.mail-from')?.textContent?.toLowerCase() || '';
    item.style.display = (text.includes(lower) || from.includes(lower)) ? '' : 'none';
  });
}

/* ================================================================
   USER UI
================================================================ */
function updateUserUI(me) {
  const name = me.displayName || 'Usuario';
  const ini  = name.split(' ').map(w => w[0]).join('').slice(0,2).toUpperCase();

  setEl('navAvatarInitials', ini);
  setEl('mppInitials',       ini);
  setEl('mppName',           name);
  setEl('mppTitle',          me.jobTitle || me.mail || '');
  setEl('menuName',          name);
  setEl('menuEmail',         me.mail || me.userPrincipalName || '');
  renderGreeting(name.split(' ')[0]);
}

function renderGreeting(first) {
  const h = new Date().getHours();
  const g = h < 12 ? 'Buenos días' : h < 20 ? 'Buenas tardes' : 'Buenas noches';
  setEl('greetingText', first ? `${g}, ${first} 👋` : `${g} 👋`);
}

function renderDate() {
  const DAYS   = ['domingo','lunes','martes','miércoles','jueves','viernes','sábado'];
  const MONTHS = ['enero','febrero','marzo','abril','mayo','junio','julio','agosto','septiembre','octubre','noviembre','diciembre'];
  const d = new Date();
  setEl('headerDate', `${DAYS[d.getDay()]}, ${d.getDate()} de ${MONTHS[d.getMonth()]} de ${d.getFullYear()}`);
}

function updateLastRefreshTime() {
  const now = new Date().toLocaleTimeString('es', { hour:'2-digit', minute:'2-digit' });
  setEl('lastUpdateTime', `Actualizado ${now}`);
}

/* ================================================================
   THEME
================================================================ */
function toggleTheme() {
  const html = document.documentElement;
  const isDark = html.getAttribute('data-theme') === 'dark';
  html.setAttribute('data-theme', isDark ? 'light' : 'dark');
  const icon = document.getElementById('themeIcon');
  if (icon) {
    icon.innerHTML = isDark
      ? '<path d="M12 7c-2.76 0-5 2.24-5 5s2.24 5 5 5 5-2.24 5-5-2.24-5-5-5zM2 13h2c.55 0 1-.45 1-1s-.45-1-1-1H2c-.55 0-1 .45-1 1s.45 1 1 1zm18 0h2c.55 0 1-.45 1-1s-.45-1-1-1h-2c-.55 0-1 .45-1 1s.45 1 1 1zM11 2v2c0 .55.45 1 1 1s1-.45 1-1V2c0-.55-.45-1-1-1s-1 .45-1 1zm0 18v2c0 .55.45 1 1 1s1-.45 1-1v-2c0-.55-.45-1-1-1s-1 .45-1 1zM5.99 4.58c-.39-.39-1.03-.39-1.41 0-.39.39-.39 1.03 0 1.41l1.06 1.06c.39.39 1.03.39 1.41 0s.39-1.03 0-1.41L5.99 4.58zm12.37 12.37c-.39-.39-1.03-.39-1.41 0-.39.39-.39 1.03 0 1.41l1.06 1.06c.39.39 1.03.39 1.41 0 .39-.39.39-1.03 0-1.41l-1.06-1.06zm1.06-10.96c.39-.39.39-1.03 0-1.41-.39-.39-1.03-.39-1.41 0l-1.06 1.06c-.39.39-.39 1.03 0 1.41s1.03.39 1.41 0l1.06-1.06zM7.05 18.36c.39-.39.39-1.03 0-1.41-.39-.39-1.03-.39-1.41 0l-1.06 1.06c-.39.39-.39 1.03 0 1.41s1.03.39 1.41 0l1.06-1.06z"/>'
      : '<path d="M12 3c-4.97 0-9 4.03-9 9s4.03 9 9 9 9-4.03 9-9c0-.46-.04-.92-.1-1.36-.98 1.37-2.58 2.26-4.4 2.26-2.98 0-5.4-2.42-5.4-5.4 0-1.81.89-3.42 2.26-4.4-.44-.06-.9-.1-1.36-.1z"/>';
  }
  showToast(isDark ? '☀️ Modo claro activado' : '🌙 Modo oscuro activado');
}

/* ================================================================
   PAGE NAVIGATION
================================================================ */
function switchPage(name, btn) {
  document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
  document.querySelectorAll('.nav-btn').forEach(b => b.classList.remove('active'));
  const page = document.getElementById('page' + name.charAt(0).toUpperCase() + name.slice(1));
  if (page) page.classList.add('active');
  if (btn) btn.classList.add('active');
  const titles = { dashboard: 'Dashboard', presence: 'Estado del equipo' };
  setEl('pageTitle', titles[name] || name);
}

/* ================================================================
   AUTO REFRESH (every 5 minutes)
================================================================ */
setInterval(() => {
  if (document.getElementById('appShell')?.style.display !== 'none') {
    loadAllData();
  }
}, 5 * 60 * 1000);

/* ================================================================
   UI HELPERS
================================================================ */
function showApp() {
  document.getElementById('loginScreen').style.display = 'none';
  document.getElementById('appShell').style.display = 'flex';
  renderDate();
  renderGreeting('');
}
function showLoading(msg) {
  setEl('loadingText', msg || 'Cargando…');
  document.getElementById('loadingOverlay').style.display = 'flex';
}
function hideLoading() {
  document.getElementById('loadingOverlay').style.display = 'none';
}
function showError(msg) {
  const el = document.getElementById('loginError');
  el.textContent = msg; el.style.display = 'block';
}
function toggleUserMenu()  { document.getElementById('userMenu').classList.toggle('open'); }
function closeUserMenu()   { document.getElementById('userMenu').classList.remove('open'); }

let _toastTimer;
function showToast(msg) {
  const el = document.getElementById('toast');
  el.textContent = msg; el.classList.add('show');
  clearTimeout(_toastTimer);
  _toastTimer = setTimeout(() => el.classList.remove('show'), 3500);
}

function setEl(id, val)  { const el = document.getElementById(id); if (el) el.textContent = val; }
function getEl(id)       { return document.getElementById(id)?.textContent || '0'; }
function escHtml(str) {
  return String(str || '')
    .replace(/&/g,'&amp;').replace(/</g,'&lt;')
    .replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

// Close menus on outside click
document.addEventListener('click', e => {
  if (!document.getElementById('navAvatar')?.contains(e.target)) closeUserMenu();
});

/* ================================================================
   INIT
================================================================ */
(async () => {
  // Show login immediately — don't wait for anything
  document.getElementById('loginScreen').style.display = 'flex';

  await initTeams();

  if (_inTeams) {
    // In Teams: skip login screen, load data directly
    document.getElementById('loginScreen').style.display = 'none';
    showApp();
    showLoading('Conectando con Microsoft 365…');
    await loadAllData();
    return;
  }

  // Browser: check for existing MSAL session silently
  try {
    await msalInstance.handleRedirectPromise();
    const accounts = msalInstance.getAllAccounts();
    if (accounts.length > 0) {
      currentAccount = accounts[0];
      document.getElementById('loginScreen').style.display = 'none';
      showApp();
      showLoading('Recuperando tu sesión…');
      await loadAllData();
      return;
    }
  } catch (e) {
    console.warn('Error al recuperar sesión:', e);
  }

  // No session — login screen is already visible, nothing more to do
})();