/* ================================================================
   CONFIGURACIÓN — Rellena estos valores
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
];

/* ================================================================
   DETECTAR ENTORNO TEAMS
================================================================ */
let _inTeams = false;
let _teamsInitialized = false;

async function initTeams() {
  try {
    await microsoftTeams.app.initialize();
    _inTeams = true;
    _teamsInitialized = true;
    console.log('[Teams] Running inside Teams client ✅');
  } catch {
    _inTeams = false;
    console.log('[Teams] Running standalone (browser) ✅');
  }
}

/* ================================================================
   MSAL (para web/browser)
================================================================ */
const msalInstance = new msal.PublicClientApplication({
  auth: {
    clientId: CONFIG.clientId,
    authority: `https://login.microsoftonline.com/${CONFIG.tenantId}`,
    redirectUri: window.location.origin + window.location.pathname,
  },
  cache: { cacheLocation: 'sessionStorage' },
});

let currentAccount = null;

/* ================================================================
   GET TOKEN — Teams SSO si estamos en cliente, MSAL si no
================================================================ */
async function getToken() {
  if (_inTeams) {
    // Teams SSO: obtiene token sin popup — funciona en desktop y web
    try {
      const authToken = await microsoftTeams.authentication.getAuthToken({
        resources: [`api://pauibz.github.io/TestTeams/${CONFIG.clientId}`],
        silent: true,
      });
      // OBO (On-Behalf-Of): cambiar el Teams token por un Graph token
      // Si el servidor OBO no está configurado, usamos MSAL como fallback
      return authToken;
    } catch (teamsErr) {
      console.warn('[Teams SSO] Silent failed, trying MSAL popup:', teamsErr);
      // Fallback: Teams web permite popup, Teams desktop podría bloquearlo
      // pero lo intentamos igualmente
    }
  }

  // MSAL estándar (browser / fallback)
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

async function callGraph(path, token, raw = false) {
  const res = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!res.ok) throw new Error(`Graph ${res.status}: ${path}`);
  if (raw) return res;
  return res.json();
}

/* ================================================================
   LOGIN / LOGOUT
================================================================ */
async function doLogin() {
  try {
    showLoading('Conectando con Microsoft…');
    if (_inTeams) {
      // En Teams: auto-autenticar con SSO
      showApp();
      await loadAllData();
    } else {
      await msalInstance.loginPopup({ scopes: SCOPES });
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
    // En Teams no hay "logout" real — volver a pantalla de login local
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

    const [calRes, tasksRes, mailRes, presRes] = await Promise.allSettled([
      callGraph(
        `/me/calendarView?startDateTime=${startDT}&endDateTime=${endDT}` +
        `&$orderby=start/dateTime&$top=20` +
        `&$select=subject,organizer,start,end,onlineMeeting,location`,
        token
      ),
      callGraph('/me/planner/tasks', token),
      callGraph(
        '/me/messages?$top=7&$orderby=receivedDateTime desc' +
        '&$select=subject,from,isRead,bodyPreview,receivedDateTime,webLink',
        token
      ),
      callGraph('/me/presence', token),
    ]);

    const meetings = calRes.status   === 'fulfilled' ? calRes.value.value   : [];
    const tasks    = tasksRes.status === 'fulfilled' ? tasksRes.value.value : [];
    const mails    = mailRes.status  === 'fulfilled' ? mailRes.value.value  : [];
    const presence = presRes.status  === 'fulfilled' ? presRes.value        : null;

    renderMeetings(meetings);
    renderTasks(tasks);
    await renderMail(mails, token);
    if (presence) applyPresence(presence.availability);
    updateKPIs(meetings.length, tasks, mails);

  } catch (err) {
    console.error('Error cargando datos:', err);
    showToast('Error de conexión — mostrando datos de ejemplo');
    loadMockData();
  }

  document.getElementById('refreshIcon').classList.remove('spinning');
  hideLoading();
}

/* ================================================================
   AVATAR — foto real de Graph API
================================================================ */
async function loadUserPhoto(me, token) {
  try {
    const res = await callGraph('/me/photo/$value', token, true);
    if (!res.ok) return;
    const blob = await res.blob();
    const url  = URL.createObjectURL(blob);
    setPhotoEl('avatarBtn',          url, me.displayName);
    setPhotoEl('presenceAvatarWrap', url, me.displayName);
  } catch { /* sin foto, usar iniciales */ }
}

function setPhotoEl(wrapperId, photoUrl, name) {
  const wrap = document.getElementById(wrapperId);
  if (!wrap) return;
  // ocultar iniciales, mostrar img
  const ini = wrap.querySelector('span');
  if (ini) ini.style.display = 'none';
  let img = wrap.querySelector('img');
  if (!img) {
    img = document.createElement('img');
    img.alt = name;
    wrap.appendChild(img);
  }
  img.src = photoUrl;
}

async function loadContactPhoto(email, token) {
  try {
    const res = await fetch(
      `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(email)}/photo/$value`,
      { headers: { Authorization: `Bearer ${token}` } }
    );
    if (!res.ok) return null;
    const blob = await res.blob();
    return URL.createObjectURL(blob);
  } catch { return null; }
}

/* ================================================================
   DATOS MOCK
================================================================ */
function todayAt(h, m) {
  const d = new Date(); d.setHours(h, m, 0, 0); return d.toISOString();
}

function loadMockData() {
  const meetings = [
    { subject: 'Daily Stand-up · Equipo Digital',     organizer: { emailAddress: { name: 'Pedro Ruiz' } },    start: { dateTime: todayAt(9,0)   }, end: { dateTime: todayAt(9,30)  }, onlineMeeting: { joinUrl: '#' } },
    { subject: 'Revisión de diseño Q2 · App móvil',   organizer: { emailAddress: { name: 'Laura Sánchez' } }, start: { dateTime: todayAt(10,30) }, end: { dateTime: todayAt(11,30) }, onlineMeeting: { joinUrl: '#' } },
    { subject: 'Sync con cliente · Proyecto Atenea',  organizer: { emailAddress: { name: 'Carlos López' } },  start: { dateTime: todayAt(12,0)  }, end: { dateTime: todayAt(12,45) }, onlineMeeting: { joinUrl: '#' } },
    { subject: 'Retrospectiva Sprint 14',              organizer: { emailAddress: { name: 'Ana Martínez' } },  start: { dateTime: todayAt(16,0)  }, end: { dateTime: todayAt(17,0)  }, onlineMeeting: { joinUrl: '#' } },
  ];
  const tasks = [
    { title: 'Revisar mockups pantalla de inicio',        percentComplete: 0,   priority: 1, dueDateTime: { dateTime: new Date().toISOString() } },
    { title: 'Actualizar documentación de componentes',   percentComplete: 0,   priority: 5, dueDateTime: { dateTime: new Date().toISOString() } },
    { title: 'Preparar agenda reunión con cliente',       percentComplete: 100, priority: 1, dueDateTime: null },
    { title: 'Enviar informe de accesibilidad',           percentComplete: 0,   priority: 5, dueDateTime: { dateTime: new Date().toISOString() } },
    { title: 'Pruebas de usabilidad – flujo checkout',    percentComplete: 0,   priority: 9, dueDateTime: null },
    { title: 'Responder feedback de Pedro sobre el DS',   percentComplete: 0,   priority: 5, dueDateTime: { dateTime: new Date().toISOString() } },
    { title: 'Cerrar issues del Sprint 13',               percentComplete: 100, priority: 9, dueDateTime: null },
    { title: 'Revisar propuesta de color tokens',         percentComplete: 100, priority: 5, dueDateTime: null },
    { title: 'Kick-off Proyecto Hermes',                  percentComplete: 0,   priority: 1, dueDateTime: { dateTime: new Date().toISOString() } },
  ];
  const mails = [
    { from: { emailAddress: { name: 'Laura Sánchez' } }, subject: 'Re: Revisión de nuevos componentes del DS',   isRead: false, bodyPreview: 'He revisado los últimos cambios y me parecen muy bien...', receivedDateTime: new Date().toISOString(), webLink: '#' },
    { from: { emailAddress: { name: 'Pedro Ruiz' } },    subject: 'Feedback sprint review – acciones pendientes', isRead: false, bodyPreview: 'Tras la reunión del viernes quedan estos puntos abiertos...', receivedDateTime: new Date(Date.now()-3e6).toISOString(), webLink: '#' },
    { from: { emailAddress: { name: 'Ana Martínez' } },  subject: 'Retrospectiva hoy 16:00 – agenda adjunta',    isRead: false, bodyPreview: 'Os comparto la agenda para que podáis prepararlo con tiempo...', receivedDateTime: new Date(Date.now()-5e6).toISOString(), webLink: '#' },
    { from: { emailAddress: { name: 'Carlos López' } },  subject: 'Confirmación reunión cliente Atenea',         isRead: true,  bodyPreview: 'Confirmo la reunión con Inmobiliaria Atenea para las 12:00.', receivedDateTime: new Date(Date.now()-86400000).toISOString(), webLink: '#' },
    { from: { emailAddress: { name: 'RRHH · Empresa' } },subject: 'Política de teletrabajo actualizada',         isRead: false, bodyPreview: 'A partir del próximo mes entra en vigor la nueva política...', receivedDateTime: new Date(Date.now()-90000000).toISOString(), webLink: '#' },
  ];

  updateUserUI({ displayName: 'Demo User', jobTitle: 'Modo demostración', mail: 'demo@tuempresa.com' });
  renderMeetings(meetings);
  renderTasks(tasks);
  renderMail(mails, null);
  applyPresence('Available');
  updateKPIs(meetings.length, tasks, mails);
  hideLoading();
}

/* ================================================================
   RENDER — MEETINGS (TIMELINE)
================================================================ */
const pad = n => String(n).padStart(2, '0');

function renderMeetings(events) {
  const el = document.getElementById('meetingsList');
  if (!events?.length) {
    el.innerHTML = '<div class="empty-state">🎉 No tienes reuniones hoy</div>';
    return;
  }
  const now = Date.now();
  el.innerHTML = '<div class="timeline-wrap">' + events.map((e, idx) => {
    const rawS = e.start.dateTime;
    const rawE = e.end.dateTime;
    const s  = new Date(rawS.endsWith('Z') ? rawS : rawS + 'Z');
    const en = new Date(rawE.endsWith('Z') ? rawE : rawE + 'Z');
    const live = now >= s && now < en;
    const past = now >= en;
    const dur  = Math.round((en - s) / 60000);
    const url  = e.onlineMeeting?.joinUrl || '#';
    const org  = e.organizer?.emailAddress?.name || '';
    const dotClass = live ? 'live' : past ? 'past' : 'future';
    const isLast = idx === events.length - 1;

    return `
      <div class="timeline-item">
        <div class="timeline-left">
          <span class="tl-time">${pad(s.getHours())}:${pad(s.getMinutes())}</span>
          <span class="tl-end">${pad(en.getHours())}:${pad(en.getMinutes())}</span>
        </div>
        <div class="tl-line-wrap">
          <div class="tl-dot ${dotClass}"></div>
          <div class="tl-connector${isLast ? ' last' : ''}"></div>
        </div>
        <div class="tl-body">
          <div class="tl-details">
            <div class="tl-title">${escHtml(e.subject)}</div>
            <div class="tl-sub">${escHtml(org)}${org ? ' · ' : ''}${dur} min</div>
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
  setEl('kpiDone',    done);
  setEl('kpiPending', pend);
  setEl('statDone',   done);
  setEl('statPend',   pend);
}

function toggleTask(i) {
  _tasks[i]._done = !_tasks[i]._done;
  drawTasks();
}

/* ================================================================
   RENDER — MAIL (con fotos reales)
================================================================ */
const AVATAR_COLORS = ['#6264A7','#0078D4','#107C10','#C4314B','#8A5A00','#007A7C','#8764B8'];

async function renderMail(msgs, token) {
  const el = document.getElementById('mailList');
  if (!msgs?.length) {
    el.innerHTML = '<div class="empty-state">No hay correos recientes</div>';
    return;
  }

  // Render inmediato con iniciales
  el.innerHTML = msgs.map((m, i) => {
    const name  = m.from?.emailAddress?.name || 'Desconocido';
    const email = m.from?.emailAddress?.address || '';
    const ini   = name.split(' ').map(w => w[0]).join('').slice(0, 2).toUpperCase();
    const color = AVATAR_COLORS[i % AVATAR_COLORS.length];
    const rec   = new Date(m.receivedDateTime);
    const isToday = rec.toDateString() === new Date().toDateString();
    const timeStr = isToday
      ? rec.toLocaleTimeString('es', { hour: '2-digit', minute: '2-digit' })
      : rec.toLocaleDateString('es', { day: '2-digit', month: 'short' });

    return `
      <div class="mail-item" onclick="window.open('${m.webLink || '#'}','_blank')">
        <div class="mail-avatar" id="mailAvatar-${i}" style="background:${color}" data-email="${escHtml(email)}">${ini}</div>
        <div class="mail-content">
          <div class="mail-from">${escHtml(name)}</div>
          <div class="mail-subject${!m.isRead ? ' unread' : ''}">${escHtml(m.subject)}</div>
          <div class="mail-preview">${escHtml(m.bodyPreview)}</div>
        </div>
        <div class="mail-right">
          <span class="mail-time">${timeStr}</span>
          ${!m.isRead ? '<div class="unread-dot"></div>' : ''}
        </div>
      </div>`;
  }).join('');

  // Cargar fotos en background si tenemos token
  if (!token) return;
  msgs.forEach(async (m, i) => {
    const email = m.from?.emailAddress?.address;
    if (!email) return;
    const photoUrl = await loadContactPhoto(email, token);
    if (!photoUrl) return;
    const avatarEl = document.getElementById(`mailAvatar-${i}`);
    if (!avatarEl) return;
    avatarEl.innerHTML = `<img src="${photoUrl}" alt="${escHtml(m.from?.emailAddress?.name || '')}" />`;
  });
}

/* ================================================================
   PRESENCE
================================================================ */
const PMAP = {
  Available:       { color: '#6BB700', label: '🟢 Disponible' },
  Busy:            { color: '#C4314B', label: '🔴 Ocupado/a' },
  DoNotDisturb:    { color: '#C4314B', label: '🔴 No molestar' },
  Away:            { color: '#F4BE22', label: '🟡 Ausente' },
  BeRightBack:     { color: '#F4BE22', label: '🟡 Ahora vuelvo' },
  Offline:         { color: '#888',    label: '⚫ Sin conexión' },
  PresenceUnknown: { color: '#888',    label: '⚫ Desconocido' },
};

function applyPresence(availability) {
  const p = PMAP[availability] || PMAP.PresenceUnknown;
  document.getElementById('presenceIndicator').style.background = p.color;
  document.getElementById('presenceStatus').textContent = p.label;
}

function setPresence(el, color, label) {
  document.querySelectorAll('.presence-opt').forEach(o => o.classList.remove('active'));
  el.classList.add('active');
  document.getElementById('presenceIndicator').style.background = color;
  document.getElementById('presenceStatus').textContent = label;
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

  setEl('kpiMeetings',   meetN);
  setEl('kpiDone',       done);
  setEl('kpiPending',    pend);
  setEl('kpiMail',       unread);
  setEl('badgeMeetings', meetN);
  setEl('badgeTasks',    taskList.length);
  setEl('badgeMail',     unread);
  setEl('statMeetTime',  meetN ? timeStr : '0m');
  setEl('statDone',      done);
  setEl('statPend',      pend);
  setEl('statUnread',    unread);
  setEl('greetingSub',   `Tienes ${meetN} reuniones hoy y ${pend} tareas pendientes.`);
}

/* ================================================================
   USER UI
================================================================ */
function updateUserUI(me) {
  const name = me.displayName || 'Usuario';
  const ini  = name.split(' ').map(w => w[0]).join('').slice(0, 2).toUpperCase();

  // Iniciales (placeholder hasta cargar foto)
  const avatarBtn = document.getElementById('avatarBtn');
  const avatarSpan = avatarBtn.querySelector('span');
  if (avatarSpan) avatarSpan.textContent = ini;

  const presWrap = document.getElementById('presenceAvatarWrap');
  const presSpan = presWrap.querySelector('span');
  if (presSpan) presSpan.textContent = ini;

  setEl('presenceName',   name);
  setEl('presenceRole',   me.jobTitle || me.mail || '');
  setEl('menuName',       name);
  setEl('menuEmail',      me.mail || me.userPrincipalName || '');
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
  el.textContent = msg;
  el.style.display = 'block';
}
function toggleUserMenu() { document.getElementById('userMenu').classList.toggle('open'); }
function closeUserMenu()  { document.getElementById('userMenu').classList.remove('open'); }

let _toastTimer;
function showToast(msg) {
  const el = document.getElementById('toast');
  el.textContent = msg;
  el.classList.add('show');
  clearTimeout(_toastTimer);
  _toastTimer = setTimeout(() => el.classList.remove('show'), 3500);
}

function setEl(id, val) { const el = document.getElementById(id); if (el) el.textContent = val; }
function getEl(id) { return document.getElementById(id)?.textContent || '0'; }
function escHtml(str) {
  return String(str || '')
    .replace(/&/g,'&amp;').replace(/</g,'&lt;')
    .replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

document.addEventListener('click', e => {
  if (!document.getElementById('avatarBtn')?.contains(e.target)) closeUserMenu();
});

/* ================================================================
   INIT
================================================================ */
(async () => {
  // 1. Intentar inicializar Teams SDK
  await initTeams();

  // 2. Si estamos en Teams, auto-login directo
  if (_inTeams) {
    showApp();
    showLoading('Conectando con Microsoft 365…');
    await loadAllData();
    return;
  }

  // 3. Si estamos en browser, comprobar sesión MSAL existente
  try {
    await msalInstance.handleRedirectPromise();
    const accounts = msalInstance.getAllAccounts();
    if (accounts.length > 0) {
      currentAccount = accounts[0];
      showApp();
      showLoading('Recuperando tu sesión…');
      await loadAllData();
    }
  } catch (e) {
    console.warn('Error al recuperar sesión:', e);
  }
})();
