/* ================================================================
   CONFIGURACIÓN
================================================================ */
const CONFIG = {
  clientId: 'c6219253-8f6e-48c9-8c4a-766aea58a874',
  tenantId: 'a3cdf8f7-40db-4d0f-8ccb-82310428392a',
};

// NOTA: Presence.Read.All y User.ReadBasic.All requieren consentimiento del admin en Azure AD
// Azure Portal → App registrations → API permissions → Grant admin consent
const SCOPES = [
  'User.Read',
  'Calendars.Read',
  'Tasks.ReadWrite',
  'Mail.Read',
  'Presence.Read',
  'Presence.Read.All',
  'People.Read',
  'User.ReadBasic.All',
  'Presence.ReadWrite',
];

/* ================================================================
   TEAMS DETECTION
================================================================ */
let _inTeams = false;
let _teamsInitialized = false;

async function initTeams() {
  return new Promise(resolve => {
    const timeout = setTimeout(() => {
      console.log('[Teams] Timeout — browser mode');
      resolve();
    }, 2500);
    try {
      microsoftTeams.app.initialize().then(() => {
        clearTimeout(timeout);
        _inTeams = true;
        _teamsInitialized = true;
        microsoftTeams.app.getContext().then(ctx => {
          if (ctx?.app?.host?.name === 'Teams') document.body.classList.add('teams-desktop');
        }).catch(() => {});
        console.log('[Teams] Running inside Teams OK');
        resolve();
      }).catch(() => { clearTimeout(timeout); resolve(); });
    } catch { clearTimeout(timeout); resolve(); }
  });
}

/* ================================================================
   MSAL
================================================================ */
const msalInstance = new msal.PublicClientApplication({
  auth: {
    clientId: CONFIG.clientId,
    authority: `https://login.microsoftonline.com/${CONFIG.tenantId}`,
    redirectUri: window.location.origin + window.location.pathname,
    navigateToLoginRequestUrl: true,
  },
  cache: { cacheLocation: 'localStorage', storeAuthStateInCookie: true },
  system: {
    allowNativeBroker: false,
    windowHashTimeout: 9000,
    iframeHashTimeout: 9000,
    loadFrameTimeout: 9000,
  },
});

let currentAccount = null;
let _cachedToken   = null;
let _tokenExpiry   = 0;

// Detectar si popups estan bloqueados (Teams Desktop, IE, algunos browsers)
function canUsePopup() {
  if (_inTeams) return false; // Teams Desktop bloquea popups
  try {
    const test = window.open('', '_blank', 'width=1,height=1');
    if (!test) return false;
    test.close();
    return true;
  } catch { return false; }
}

/* ================================================================
   TOKEN — cached, renovacion automatica
================================================================ */
async function getToken() {
  if (_cachedToken && Date.now() < _tokenExpiry - 120000) {
    return _cachedToken;
  }

  const accounts = msalInstance.getAllAccounts();
  if (accounts.length) {
    currentAccount = accounts[0];
    try {
      const r = await msalInstance.acquireTokenSilent({ scopes: SCOPES, account: currentAccount });
      _cachedToken = r.accessToken;
      _tokenExpiry = r.expiresOn?.getTime() || (Date.now() + 3600000);
      return _cachedToken;
    } catch (silentErr) {
      console.warn('[Token] Silent failed:', silentErr.errorCode);
      // Si el error es interaccion requerida, lanzar flujo interactivo
      if (!silentErr.errorCode?.includes('interaction_required') &&
          !silentErr.errorCode?.includes('consent_required') &&
          !silentErr.errorCode?.includes('login_required')) {
        throw silentErr;
      }
    }
  } else if (accounts.length === 0 && !_inTeams) {
    throw new Error('Sin sesion activa');
  }

  // Flujo interactivo: redirect (funciona en Teams Desktop y browsers con popups bloqueados)
  // Guardar estado para saber que volvemos de un redirect de token
  sessionStorage.setItem('msal_token_redirect_pending', '1');
  await msalInstance.acquireTokenRedirect({ scopes: SCOPES, account: currentAccount || undefined });
  // acquireTokenRedirect navega la pagina — el codigo despues de esta linea no se ejecuta
  throw new Error('Redirigiendo para autenticacion...');
}

/* ================================================================
   GRAPH HELPER con logging
================================================================ */
async function callGraph(path, token, options = {}) {
  const url = path.startsWith('http') ? path : `https://graph.microsoft.com/v1.0${path}`;
  const headers = { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' };
  if (options.eventual) headers['ConsistencyLevel'] = 'eventual';

  const res = await fetch(url, {
    method: options.method || 'GET',
    headers,
    body: options.body ? JSON.stringify(options.body) : undefined,
  });

  if (options.raw) return res;

  if (!res.ok) {
    let errMsg = `HTTP ${res.status}`;
    try {
      const errBody = await res.json();
      errMsg = errBody?.error?.message || errBody?.error?.code || errMsg;
    } catch { /* ignore */ }
    console.warn(`[Graph] ${options.method || 'GET'} ${path} -> ${res.status}: ${errMsg}`);
    const err = new Error(errMsg);
    err.status = res.status;
    err.path   = path;
    throw err;
  }

  if (res.status === 204) return null;
  return res.json();
}

/* ================================================================
   LOGIN / LOGOUT
================================================================ */
async function doLogin() {
  try {
    showLoading('Conectando con Microsoft...');
    if (_inTeams) {
      showApp();
      await loadAllData();
      return;
    }
    // Siempre usar redirect — funciona en Teams Desktop, browsers con popups bloqueados y browser normal
    await msalInstance.loginRedirect({ scopes: SCOPES });
    // loginRedirect navega la pagina, el codigo de abajo no se ejecuta durante el login
  } catch (e) {
    hideLoading();
    showError('Error al iniciar sesion: ' + (e.message || 'Intentalo de nuevo.'));
  }
}

function enterDemo() {
  showApp();
  loadMockData();
}

function doLogout() {
  closeUserMenu();
  _cachedToken = null;
  _tokenExpiry = 0;
  if (_inTeams) {
    document.getElementById('appShell').style.display = 'none';
    document.getElementById('loginScreen').style.display = 'flex';
  } else {
    msalInstance.logoutRedirect({ postLogoutRedirectUri: window.location.origin + window.location.pathname });
  }
}

/* ================================================================
   LOAD ALL DATA — secciones independientes, errores localizados
================================================================ */
async function loadAllData() {
  const icon = document.getElementById('refreshIcon');
  if (icon) icon.classList.add('spinning');
  showLoading('Actualizando desde Microsoft 365...');

  let token;
  try {
    token = await getToken();
  } catch (err) {
    console.error('[Auth] No se pudo obtener token:', err);
    hideLoading();
    if (icon) icon.classList.remove('spinning');
    showToast('Error de autenticacion: ' + err.message);
    showSectionError('meetingsList', 'Error de autenticacion');
    showSectionError('mailList',     'Error de autenticacion');
    showSectionError('tasksList',    'Error de autenticacion');
    showSectionError('teamGrid',     'Error de autenticacion');
    return;
  }

  // Datos del usuario
  let me = null;
  try {
    me = await callGraph('/me', token);
    updateUserUI(me);
    loadUserPhoto(me, token);
  } catch (err) {
    console.warn('[Graph] /me fallo:', err);
  }

  const today   = new Date().toISOString().split('T')[0];
  const startDT = `${today}T00:00:00`;
  const endDT   = `${today}T23:59:59`;

  const [calRes, tasksRes, mailRes, presRes, peopleRes] = await Promise.allSettled([
    callGraph(
      `/me/calendarView?startDateTime=${startDT}&endDateTime=${endDT}` +
      `&$orderby=start/dateTime&$top=20` +
      `&$select=subject,organizer,start,end,onlineMeeting,location`,
      token
    ),
    callGraph('/me/planner/tasks', token),
    callGraph(
      '/me/messages?$top=8&$orderby=receivedDateTime desc' +
      '&$select=subject,from,isRead,bodyPreview,receivedDateTime,webLink,importance',
      token
    ),
    callGraph('/me/presence', token),
    callGraph(
      '/me/people?$top=20&$select=id,displayName,jobTitle,userPrincipalName,scoredEmailAddresses',
      token
    ),
  ]);

  // Reuniones
  if (calRes.status === 'fulfilled') {
    renderMeetings(calRes.value?.value ?? []);
    scheduleMeetingAlerts(calRes.value?.value ?? []);
  } else {
    console.warn('[Graph] Calendario:', calRes.reason?.message);
    showSectionError('meetingsList', 'No se pudieron cargar las reuniones');
  }

  // Tareas
  if (tasksRes.status === 'fulfilled') {
    renderTasks(tasksRes.value?.value ?? []);
  } else {
    console.warn('[Graph] Planner:', tasksRes.reason?.message);
    showSectionError('tasksList', 'No se pudieron cargar las tareas (requiere Planner activo)');
  }

  // Correos
  if (mailRes.status === 'fulfilled') {
    await renderMail(mailRes.value?.value ?? [], token);
  } else {
    console.warn('[Graph] Mail:', mailRes.reason?.message);
    showSectionError('mailList', 'No se pudieron cargar los correos');
  }

  // Mi presencia
  if (presRes.status === 'fulfilled' && presRes.value) {
    applyPresence(presRes.value.availability);
    syncMyPresenceUI(presRes.value.availability);
  } else {
    console.warn('[Graph] Presence:', presRes.reason?.message);
  }

  // KPIs
  const meetings = calRes.status   === 'fulfilled' ? (calRes.value?.value   ?? []) : [];
  const tasks    = tasksRes.status === 'fulfilled' ? (tasksRes.value?.value ?? []) : [];
  const mails    = mailRes.status  === 'fulfilled' ? (mailRes.value?.value  ?? []) : [];
  updateKPIs(meetings.length, tasks, mails);

  // Presencia del equipo en background
  if (peopleRes.status === 'fulfilled') {
    const people = peopleRes.value?.value ?? [];
    console.log(`[Graph] People: ${people.length} contactos`);
    loadTeamPresence(people, token);
  } else {
    console.warn('[Graph] People:', peopleRes.reason?.message);
    showSectionError('teamGrid', 'No se pudieron cargar los contactos (permiso People.Read requerido)');
  }

  if (icon) icon.classList.remove('spinning');
  hideLoading();
  updateLastRefreshTime();
}

function showSectionError(elId, msg) {
  const el = document.getElementById(elId);
  if (el) el.innerHTML = `<div class="empty-state" style="color:#FF6B6B">Error: ${msg}</div>`;
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
    setAvatarEl('navAvatar', url, me.displayName);
    setAvatarEl('mppAvatar', url, me.displayName);
  } catch { /* usar iniciales */ }
}

function setAvatarEl(wrapperId, photoUrl, name) {
  const wrap = document.getElementById(wrapperId);
  if (!wrap) return;
  const ini = wrap.querySelector('span');
  if (ini) ini.style.display = 'none';
  let img = wrap.querySelector('img');
  if (!img) { img = document.createElement('img'); img.alt = name || ''; wrap.appendChild(img); }
  img.src = photoUrl;
}

async function loadContactPhoto(userId, token) {
  try {
    const res = await fetch(
      `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(userId)}/photo/$value`,
      { headers: { Authorization: `Bearer ${token}` } }
    );
    if (!res.ok) return null;
    return URL.createObjectURL(await res.blob());
  } catch { return null; }
}

/* ================================================================
   TEAM PRESENCE — logica robusta paso a paso
================================================================ */
let _teamData   = [];
let _teamFilter = 'all';
let _teamSearch = '';

async function loadTeamPresence(people, token) {
  const teamGridEl = document.getElementById('teamGrid');
  if (!teamGridEl) return;

  if (!people.length) {
    teamGridEl.innerHTML = '<div class="empty-state" style="grid-column:1/-1">No hay contactos en la API People. Verifica permisos People.Read</div>';
    return;
  }

  teamGridEl.innerHTML = '<div class="empty-state" style="grid-column:1/-1">Resolviendo usuarios...</div>';

  // Paso 1: Resolver IDs de usuario via UPN
  const userMap = {};
  const userIds = [];

  for (const p of people) {
    const upn   = p.userPrincipalName;
    const email = p.scoredEmailAddresses?.[0]?.address;
    const identifier = upn || email;
    if (!identifier) continue;
    if (identifier.includes('#EXT#')) continue; // saltar externos

    try {
      const u = await callGraph(
        `/users/${encodeURIComponent(identifier)}?$select=id,displayName,jobTitle,mail,userPrincipalName`,
        token
      );
      if (u?.id && !userMap[u.id]) {
        userIds.push(u.id);
        userMap[u.id] = {
          displayName:       u.displayName || p.displayName || 'Desconocido',
          jobTitle:          u.jobTitle    || p.jobTitle    || '',
          mail:              u.mail        || email         || '',
          userPrincipalName: u.userPrincipalName || upn    || '',
        };
      }
    } catch (err) {
      console.warn(`[Graph] Usuario ${identifier}:`, err.message);
    }
  }

  console.log(`[Presence] ${userIds.length} IDs resueltos de ${people.length} contactos`);

  if (!userIds.length) {
    teamGridEl.innerHTML = '<div class="empty-state" style="grid-column:1/-1">No se resolvieron IDs de usuario. Verifica permisos User.ReadBasic.All</div>';
    return;
  }

  teamGridEl.innerHTML = '<div class="empty-state" style="grid-column:1/-1">Cargando presencia...</div>';

  // Paso 2: Presencia en batch (requiere Presence.Read.All)
  let presences = [];
  const batchIds = userIds.slice(0, 15);

  try {
    const pr = await callGraph('/communications/getPresencesByUserId', token, {
      method: 'POST',
      body: { ids: batchIds },
    });
    presences = pr?.value ?? [];
    console.log(`[Presence] Batch OK: ${presences.length} presencias`);
  } catch (batchErr) {
    console.warn('[Presence] Batch fallo (requiere Presence.Read.All admin consent):', batchErr.message);

    // Fallback: presencia individual
    for (const id of batchIds) {
      try {
        const p = await callGraph(`/users/${id}/presence`, token);
        if (p) presences.push(p);
      } catch (indErr) {
        console.warn(`[Presence] Individual ${id}:`, indErr.message);
        presences.push({ id, availability: 'PresenceUnknown', activity: '' });
      }
    }
    console.log(`[Presence] Fallback individual: ${presences.length} resultados`);
  }

  // Paso 3: Combinar
  _teamData = [];
  const seenIds = new Set();

  for (const pr of presences) {
    const info = userMap[pr.id];
    if (!info || seenIds.has(pr.id)) continue;
    seenIds.add(pr.id);
    _teamData.push({
      id:           pr.id,
      name:         info.displayName,
      role:         info.jobTitle,
      mail:         info.mail,
      upn:          info.userPrincipalName,
      availability: pr.availability || 'PresenceUnknown',
      activity:     pr.activity     || '',
    });
  }

  // Usuarios sin presencia
  for (const id of batchIds) {
    if (!seenIds.has(id) && userMap[id]) {
      _teamData.push({
        id,
        name:         userMap[id].displayName,
        role:         userMap[id].jobTitle,
        mail:         userMap[id].mail,
        upn:          userMap[id].userPrincipalName,
        availability: 'PresenceUnknown',
        activity:     '',
      });
    }
  }

  console.log(`[Presence] Renderizando ${_teamData.length} miembros`);
  renderTeamGrid(_teamData);

  // Paso 4: Fotos en background
  loadTeamPhotos(token);
}

async function loadTeamPhotos(token) {
  for (const m of _teamData) {
    if (!m.id) continue;
    const url = await loadContactPhoto(m.id, token);
    if (!url) continue;
    const el = document.getElementById(`tc-avatar-${m.id}`);
    if (el) {
      el.textContent = '';
      const img = document.createElement('img');
      img.src = url; img.alt = m.name;
      img.style.cssText = 'width:100%;height:100%;object-fit:cover;border-radius:50%';
      el.appendChild(img);
    }
  }
}

function renderTeamGrid(data) {
  const el = document.getElementById('teamGrid');
  if (!el) return;

  let filtered = [...data];

  if (_teamFilter !== 'all') {
    filtered = filtered.filter(m => {
      if (_teamFilter === 'Available') return m.availability === 'Available';
      if (_teamFilter === 'Busy')      return ['Busy','DoNotDisturb'].includes(m.availability);
      if (_teamFilter === 'Away')      return ['Away','BeRightBack'].includes(m.availability);
      if (_teamFilter === 'Offline')   return ['Offline','PresenceUnknown'].includes(m.availability);
      return true;
    });
  }

  if (_teamSearch) {
    const q = _teamSearch.toLowerCase();
    filtered = filtered.filter(m =>
      m.name.toLowerCase().includes(q) || m.role.toLowerCase().includes(q)
    );
  }

  if (!filtered.length) {
    el.innerHTML = '<div class="empty-state" style="grid-column:1/-1">No hay resultados</div>';
    return;
  }

  el.innerHTML = filtered.map(m => {
    const p      = PMAP[m.availability] || PMAP.PresenceUnknown;
    const ini    = m.name.split(' ').map(w => w[0]).join('').slice(0, 2).toUpperCase();
    const actStr = m.activity && m.activity !== m.availability ? m.activity : '';
    return `
      <div class="team-card">
        <div class="tc-avatar-wrap">
          <div class="tc-avatar" id="tc-avatar-${m.id}" style="background:${p.avatarColor}">${ini}</div>
          <div class="tc-status-dot" style="background:${p.color}"></div>
        </div>
        <div class="tc-name">${escHtml(m.name)}</div>
        <div class="tc-role">${escHtml(m.role)}</div>
        <div class="tc-badge ${p.badge}">${p.label}</div>
        ${actStr ? `<div class="tc-activity">${escHtml(actStr)}</div>` : ''}
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
  Available:       { color: '#6BB700', label: 'Disponible',   badge: 'available', avatarColor: '#107C10' },
  Busy:            { color: '#C4314B', label: 'Ocupado/a',    badge: 'busy',      avatarColor: '#A4262C' },
  DoNotDisturb:    { color: '#C4314B', label: 'No molestar',  badge: 'busy',      avatarColor: '#A4262C' },
  Away:            { color: '#F4BE22', label: 'Ausente',      badge: 'away',      avatarColor: '#8A5A00' },
  BeRightBack:     { color: '#F4BE22', label: 'Ahora vuelvo', badge: 'away',      avatarColor: '#8A5A00' },
  Offline:         { color: '#666',    label: 'Sin conexion', badge: 'offline',   avatarColor: '#444'    },
  PresenceUnknown: { color: '#666',    label: 'Desconocido',  badge: 'offline',   avatarColor: '#444'    },
};

const PRESENCE_LABELS = {
  Available:       'Disponible',
  Busy:            'Ocupado/a',
  DoNotDisturb:    'No molestar',
  Away:            'Ausente',
  BeRightBack:     'Ahora vuelvo',
  Offline:         'Sin conexion',
  PresenceUnknown: 'Desconocido',
};

function applyPresence(availability) {
  const p      = PMAP[availability] || PMAP.PresenceUnknown;
  const navDot = document.getElementById('navPresenceDot');
  if (navDot) navDot.style.background = p.color;
}

function syncMyPresenceUI(availability) {
  const p         = PMAP[availability] || PMAP.PresenceUnknown;
  const mppDot    = document.getElementById('mppStatusDot');
  const mppStatus = document.getElementById('mppStatus');
  if (mppDot)    mppDot.style.background = p.color;
  if (mppStatus) mppStatus.textContent   = PRESENCE_LABELS[availability] || 'Desconocido';

  document.querySelectorAll('.presence-opt').forEach(o => o.classList.remove('active'));
  const MAP_IDX = { Available: 0, Busy: 1, BeRightBack: 2, DoNotDisturb: 3, Away: 4 };
  const idx = MAP_IDX[availability];
  if (idx !== undefined) {
    const opts = document.querySelectorAll('.presence-opt');
    if (opts[idx]) opts[idx].classList.add('active');
  }
}

// Mapa availability -> activity valido en Graph API
const PRESENCE_ACTIVITY_MAP = {
  Available:    'Available',
  Busy:         'Busy',
  DoNotDisturb: 'DoNotDisturb',
  BeRightBack:  'BeRightBack',
  Away:         'Away',
  Offline:      'OffWork',
};

async function setMyPresence(el, availability, color, label) {
  document.querySelectorAll('.presence-opt').forEach(o => o.classList.remove('active'));
  el.classList.add('active');

  const navDot    = document.getElementById('navPresenceDot');
  const mppDot    = document.getElementById('mppStatusDot');
  const mppStatus = document.getElementById('mppStatus');
  if (navDot)    navDot.style.background = color;
  if (mppDot)    mppDot.style.background = color;
  if (mppStatus) mppStatus.textContent   = label;

  let token;
  try {
    token = await getToken();
  } catch (err) {
    showToast('No se pudo obtener token: ' + err.message);
    return;
  }

  // Verificar que el token tiene el scope Presence.ReadWrite
  // Graph: POST /me/presence/setPresence
  // Requiere: Presence.ReadWrite (delegated) — sin admin consent
  const activity = PRESENCE_ACTIVITY_MAP[availability] || 'Available';

  try {
    const res = await fetch('https://graph.microsoft.com/v1.0/me/presence/setPresence', {
      method: 'POST',
      headers: {
        Authorization: 'Bearer ' + token,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        sessionId:          CONFIG.clientId,
        availability:       availability,
        activity:           activity,
        expirationDuration: 'PT4H',
      }),
    });

    if (res.status === 204 || res.ok) {
      showToast('Estado actualizado en Teams: ' + label);
      return;
    }

    // Leer error detallado
    let errDetail = `HTTP ${res.status}`;
    try {
      const errJson = await res.json();
      errDetail = errJson?.error?.message || errJson?.error?.code || errDetail;
      console.warn('[Presence Write] Error:', errJson);
    } catch { /* ignore */ }

    // 403 = falta permiso Presence.ReadWrite en Azure
    if (res.status === 403) {
      showToast('Sin permiso de escritura. Contacta al admin para conceder Presence.ReadWrite');
    } else {
      showToast('Error al actualizar: ' + errDetail);
    }

  } catch (err) {
    console.error('[Presence Write] Fetch error:', err);
    showToast('Error de red al actualizar presencia');
  }
}

/* ================================================================
   MOCK DATA
================================================================ */
function todayAt(h, m) {
  const d = new Date(); d.setHours(h, m, 0, 0); return d.toISOString();
}

const MOCK_TEAM = [
  { id: 't1', name: 'Laura Sanchez',  role: 'Product Designer', mail: 'laura@demo.com',  availability: 'Available',    activity: 'Available' },
  { id: 't2', name: 'Pedro Ruiz',     role: 'Frontend Dev',     mail: 'pedro@demo.com',  availability: 'Busy',         activity: 'InACall' },
  { id: 't3', name: 'Ana Martinez',   role: 'UX Researcher',    mail: 'ana@demo.com',    availability: 'Away',         activity: 'Away' },
  { id: 't4', name: 'Carlos Lopez',   role: 'Backend Dev',      mail: 'carlos@demo.com', availability: 'Available',    activity: 'Available' },
  { id: 't5', name: 'Marta Gil',      role: 'Project Manager',  mail: 'marta@demo.com',  availability: 'DoNotDisturb', activity: 'Presenting' },
  { id: 't6', name: 'Jorge Perez',    role: 'Data Engineer',    mail: 'jorge@demo.com',  availability: 'Offline',      activity: 'Offline' },
  { id: 't7', name: 'Sofia Torres',   role: 'QA Engineer',      mail: 'sofia@demo.com',  availability: 'BeRightBack',  activity: 'BeRightBack' },
  { id: 't8', name: 'Ivan Castro',    role: 'DevOps',           mail: 'ivan@demo.com',   availability: 'Available',    activity: 'Available' },
];

function loadMockData() {
  const meetings = [
    { subject: 'Daily Stand-up Equipo Digital',   organizer: { emailAddress: { name: 'Pedro Ruiz' } },   start: { dateTime: todayAt(9,0)   }, end: { dateTime: todayAt(9,30)  }, onlineMeeting: { joinUrl: '#' } },
    { subject: 'Revision de diseno Q2 App movil', organizer: { emailAddress: { name: 'Laura Sanchez' } }, start: { dateTime: todayAt(10,30) }, end: { dateTime: todayAt(11,30) }, onlineMeeting: { joinUrl: '#' } },
    { subject: 'Sync con cliente Proyecto Atenea',organizer: { emailAddress: { name: 'Carlos Lopez' } },  start: { dateTime: todayAt(14,0)  }, end: { dateTime: todayAt(14,45) }, onlineMeeting: { joinUrl: '#' } },
    { subject: 'Retrospectiva Sprint 14',          organizer: { emailAddress: { name: 'Ana Martinez' } }, start: { dateTime: todayAt(16,0)  }, end: { dateTime: todayAt(17,0)  }, onlineMeeting: { joinUrl: '#' } },
  ];
  const tasks = [
    { id:'1', title: 'Revisar mockups pantalla de inicio',      percentComplete: 0,   priority: 1, dueDateTime: { dateTime: new Date().toISOString() } },
    { id:'2', title: 'Actualizar documentacion de componentes', percentComplete: 0,   priority: 5, dueDateTime: { dateTime: new Date().toISOString() } },
    { id:'3', title: 'Preparar agenda reunion con cliente',     percentComplete: 100, priority: 1, dueDateTime: null },
    { id:'4', title: 'Enviar informe de accesibilidad',         percentComplete: 0,   priority: 5, dueDateTime: { dateTime: new Date().toISOString() } },
    { id:'5', title: 'Pruebas de usabilidad flujo checkout',    percentComplete: 0,   priority: 9, dueDateTime: null },
    { id:'6', title: 'Responder feedback de Pedro sobre el DS', percentComplete: 0,   priority: 5, dueDateTime: { dateTime: new Date().toISOString() } },
    { id:'7', title: 'Cerrar issues del Sprint 13',             percentComplete: 100, priority: 9, dueDateTime: null },
    { id:'8', title: 'Kick-off Proyecto Hermes',                percentComplete: 0,   priority: 1, dueDateTime: { dateTime: new Date().toISOString() } },
  ];
  const mails = [
    { from: { emailAddress: { name: 'Laura Sanchez' } }, subject: 'Revision de nuevos componentes del DS',    isRead: false, importance: 'high',   bodyPreview: 'He revisado los ultimos cambios...', receivedDateTime: new Date().toISOString(), webLink: '#' },
    { from: { emailAddress: { name: 'Pedro Ruiz' } },    subject: 'Feedback sprint review acciones pendientes', isRead: false, importance: 'normal', bodyPreview: 'Tras la reunion del viernes quedan estos puntos...', receivedDateTime: new Date(Date.now()-3e6).toISOString(), webLink: '#' },
    { from: { emailAddress: { name: 'Ana Martinez' } },  subject: 'Retrospectiva hoy 16:00 agenda adjunta',   isRead: false, importance: 'normal', bodyPreview: 'Os comparto la agenda para que podais prepararla...', receivedDateTime: new Date(Date.now()-5e6).toISOString(), webLink: '#' },
    { from: { emailAddress: { name: 'Carlos Lopez' } },  subject: 'Confirmacion reunion cliente Atenea',       isRead: true,  importance: 'normal', bodyPreview: 'Confirmo la reunion con Atenea para las 14:00.', receivedDateTime: new Date(Date.now()-86400000).toISOString(), webLink: '#' },
    { from: { emailAddress: { name: 'RRHH Empresa' } },  subject: 'Politica de teletrabajo actualizada',       isRead: false, importance: 'normal', bodyPreview: 'A partir del proximo mes entra en vigor la nueva politica...', receivedDateTime: new Date(Date.now()-90000000).toISOString(), webLink: '#' },
  ];

  updateUserUI({ displayName: 'Demo User', jobTitle: 'Modo demostracion', mail: 'demo@tuempresa.com' });
  renderMeetings(meetings);
  renderTasks(tasks);
  renderMail(mails, null);
  applyPresence('Available');
  syncMyPresenceUI('Available');
  updateKPIs(meetings.length, tasks, mails);
  scheduleMeetingAlerts(meetings);
  _teamData = MOCK_TEAM;
  renderTeamGrid(_teamData);
  hideLoading();
  updateLastRefreshTime();
  pushNotification('info', 'Modo demo', 'Datos de ejemplo. Inicia sesion con Microsoft para ver datos reales.');
}

/* ================================================================
   RENDER MEETINGS
================================================================ */
const pad = n => String(n).padStart(2, '0');

function renderMeetings(events) {
  const el = document.getElementById('meetingsList');
  if (!events?.length) {
    el.innerHTML = '<div class="empty-state">No tienes reuniones hoy</div>';
    setEl('badgeMeetings', '0');
    return;
  }
  setEl('badgeMeetings', events.length);
  const now = Date.now();
  el.innerHTML = events.map((e, idx) => {
    const rawS = e.start.dateTime;
    const rawE = e.end.dateTime;
    const s    = new Date(rawS.endsWith('Z') ? rawS : rawS + 'Z');
    const en   = new Date(rawE.endsWith('Z') ? rawE : rawE + 'Z');
    const live = now >= s.getTime() && now < en.getTime();
    const past = now >= en.getTime();
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
              <div class="tl-meta">${escHtml(org)}${org ? ' - ' : ''}${dur} min</div>
              <div class="tl-chips">
                ${live ? '<span class="chip live">En directo</span>' : ''}
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
  }).join('');
}

/* ================================================================
   RENDER TASKS
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
    el.innerHTML = '<div class="empty-state">Sin tareas pendientes</div>';
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
  setEl('kpiDone', done); setEl('kpiPending', pend);
  setEl('statDone', done); setEl('statPend', pend);
}

function toggleTask(i) {
  _tasks[i]._done = !_tasks[i]._done;
  drawTasks();
  showToast(_tasks[i]._done ? 'Tarea completada' : 'Tarea desmarcada');
}

/* ================================================================
   RENDER MAIL
================================================================ */
const AVATAR_COLORS = ['#6264A7','#0078D4','#107C10','#C4314B','#8A5A00','#007A7C','#8764B8'];

async function renderMail(msgs, token) {
  const el     = document.getElementById('mailList');
  const unread = msgs?.filter(m => !m.isRead).length ?? 0;
  setEl('badgeMail', unread);
  setEl('kpiMail',   unread);

  if (!msgs?.length) {
    el.innerHTML = '<div class="empty-state">No hay correos recientes</div>';
    return;
  }

  el.innerHTML = msgs.map((m, i) => {
    const name    = m.from?.emailAddress?.name    || 'Desconocido';
    const ini     = name.split(' ').map(w => w[0]).join('').slice(0, 2).toUpperCase();
    const color   = AVATAR_COLORS[i % AVATAR_COLORS.length];
    const rec     = new Date(m.receivedDateTime);
    const isToday = rec.toDateString() === new Date().toDateString();
    const timeStr = isToday
      ? rec.toLocaleTimeString('es', { hour: '2-digit', minute: '2-digit' })
      : rec.toLocaleDateString('es',  { day: '2-digit', month: 'short' });
    const urgent = m.importance === 'high';
    return `
      <div class="mail-item" onclick="window.open('${m.webLink || '#'}','_blank')">
        <div class="mail-avatar" id="mailAvatar-${i}" style="background:${color}">${ini}</div>
        <div class="mail-content">
          <div class="mail-from">${escHtml(name)} ${urgent ? '🔴' : ''}</div>
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
    if (avatarEl) {
      avatarEl.textContent = '';
      const img = document.createElement('img');
      img.src = photoUrl;
      img.style.cssText = 'width:100%;height:100%;object-fit:cover;border-radius:50%';
      avatarEl.appendChild(img);
    }
  });
}

/* ================================================================
   KPIs
================================================================ */
function updateKPIs(meetN, taskList, mails) {
  const done   = taskList.filter(t => t.percentComplete === 100).length;
  const pend   = taskList.filter(t => t.percentComplete  <  100).length;
  const unread = mails.filter(m => !m.isRead).length;
  const totalMin = meetN * 45;
  const timeStr  = totalMin >= 60
    ? `${Math.floor(totalMin / 60)}h ${totalMin % 60 ? totalMin % 60 + 'm' : ''}`
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
   MEETING ALERTS
================================================================ */
const _alertTimers = [];

function scheduleMeetingAlerts(meetings) {
  _alertTimers.forEach(clearTimeout);
  _alertTimers.length = 0;
  const now = Date.now();
  for (const m of meetings) {
    const rawS    = m.start.dateTime;
    const s       = new Date(rawS.endsWith('Z') ? rawS : rawS + 'Z');
    const msUntil = s.getTime() - now;
    if (msUntil > 0 && msUntil < 30 * 60 * 1000) {
      const minLeft = Math.round(msUntil / 60000);
      showMeetingAlert(m.subject, minLeft, m.onlineMeeting?.joinUrl);
      pushNotification('warning', `Reunion en ${minLeft} min`, m.subject);
    }
  }
}

function showMeetingAlert(subject, minLeft, url) {
  const c = document.getElementById('alertBannerContainer');
  if (!c) return;
  const id  = `alert-${Date.now()}`;
  const div = document.createElement('div');
  div.id = id; div.className = 'alert-banner';
  div.innerHTML = `
    <div class="alert-banner-left">
      <span style="font-size:20px">📅</span>
      <div class="alert-banner-text">
        <strong>Reunion en ${minLeft} min</strong>
        <span>${escHtml(subject)}</span>
      </div>
    </div>
    <div style="display:flex;gap:8px;align-items:center">
      ${url && url !== '#' ? `<button class="join-btn" onclick="window.open('${escHtml(url)}','_blank')">Unirse</button>` : ''}
      <button class="alert-dismiss" onclick="document.getElementById('${id}').remove()">x</button>
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
  badge.style.display = _notifications.length ? 'flex' : 'none';
  badge.textContent   = _notifications.length > 9 ? '9+' : _notifications.length;
}

function renderNotifList() {
  const el = document.getElementById('notifList');
  if (!el) return;
  if (!_notifications.length) { el.innerHTML = '<div class="empty-state">Sin notificaciones</div>'; return; }
  el.innerHTML = _notifications.map(n => {
    const t      = n.time.toLocaleTimeString('es', { hour: '2-digit', minute: '2-digit' });
    const cls    = { warning: 'warning', urgent: 'urgent', info: 'info' }[n.type] || 'info';
    const lbl    = { warning: 'Proximamente', urgent: 'Urgente', info: 'Info' }[n.type] || n.type;
    return `
      <div class="notif-item">
        <div class="notif-badge ${cls}">${lbl}</div>
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
  if (panel.classList.contains('open')) {
    panel.classList.remove('open'); dimmer.classList.remove('show');
  } else {
    panel.classList.add('open'); dimmer.classList.add('show');
    _notifications = []; updateNotifBadge();
  }
}

function closeNotifPanel() {
  document.getElementById('notifPanel')?.classList.remove('open');
  document.getElementById('dimmer')?.classList.remove('show');
}

/* ================================================================
   DAILY SUMMARY
================================================================ */
function generateSummary() {
  const el     = document.getElementById('summaryContent');
  const pending = parseInt(getEl('kpiPending')) || 0;
  const done    = parseInt(getEl('kpiDone'))    || 0;
  const unread  = parseInt(getEl('kpiMail'))    || 0;
  const meetN   = parseInt(getEl('kpiMeetings'))|| 0;
  const h       = new Date().getHours();
  const greet   = h < 12 ? 'esta manana' : h < 18 ? 'esta tarde' : 'esta noche';
  el.innerHTML = `
    <div style="display:flex;flex-direction:column;gap:8px">
      <div>📅 <strong>Reuniones:</strong> Tienes ${meetN} reuniones ${greet}.</div>
      <div>🗂 <strong>Tareas:</strong> ${pending} pendientes, ${done} completadas.</div>
      <div>📧 <strong>Correos:</strong> ${unread} sin leer. ${unread > 3 ? 'Reserva 20 min para responder los urgentes.' : 'Bandeja bajo control.'}</div>
      <div style="margin-top:6px;padding:10px;background:var(--brand-l);border-radius:8px;font-size:12px;color:var(--ts)">
        Consejo: Agrupa las respuestas de correo en bloques de 20-30 min para mantener el foco.
      </div>
    </div>`;
  showToast('Resumen generado');
}

/* ================================================================
   POWER BI
================================================================ */
function loadPowerBI() {
  const url = document.getElementById('powerbiUrl')?.value?.trim();
  if (!url) { showToast('Introduce una URL valida de Power BI'); return; }
  document.getElementById('powerbiContent').innerHTML =
    `<iframe id="powerbiFrame" src="${escHtml(url)}" allowfullscreen></iframe>`;
}

function togglePowerBI() {
  const c = document.getElementById('powerbiContent');
  if (document.getElementById('powerbiFrame')) {
    c.innerHTML = `<div class="powerbi-body"><div class="powerbi-icon">📊</div><div class="powerbi-msg">Pega la URL de tu informe de Power BI embebido.</div><div class="powerbi-input"><input type="text" id="powerbiUrl" placeholder="https://app.powerbi.com/reportEmbed?reportId=..." /><button onclick="loadPowerBI()">Cargar</button></div></div>`;
  }
}

/* ================================================================
   GLOBAL SEARCH
================================================================ */
function handleSearch(q) {
  const lower = q.toLowerCase();
  document.querySelectorAll('.task-item').forEach(item => {
    const text = item.querySelector('.task-name')?.textContent?.toLowerCase() || '';
    item.style.display = !q || text.includes(lower) ? '' : 'none';
  });
  document.querySelectorAll('.mail-item').forEach(item => {
    const subj = item.querySelector('.mail-subject')?.textContent?.toLowerCase() || '';
    const from = item.querySelector('.mail-from')?.textContent?.toLowerCase()    || '';
    item.style.display = !q || subj.includes(lower) || from.includes(lower) ? '' : 'none';
  });
}

/* ================================================================
   USER UI
================================================================ */
function updateUserUI(me) {
  const name = me.displayName || 'Usuario';
  const ini  = name.split(' ').map(w => w[0]).join('').slice(0, 2).toUpperCase();
  setEl('navAvatarInitials', ini);
  setEl('mppInitials',  ini);
  setEl('mppName',      name);
  setEl('mppTitle',     me.jobTitle || me.mail || '');
  setEl('menuName',     name);
  setEl('menuEmail',    me.mail || me.userPrincipalName || '');
  renderGreeting(name.split(' ')[0]);
}

function renderGreeting(first) {
  const h = new Date().getHours();
  const g = h < 12 ? 'Buenos dias' : h < 20 ? 'Buenas tardes' : 'Buenas noches';
  setEl('greetingText', first ? `${g}, ${first}` : `${g}`);
}

function renderDate() {
  const DAYS   = ['domingo','lunes','martes','miercoles','jueves','viernes','sabado'];
  const MONTHS = ['enero','febrero','marzo','abril','mayo','junio','julio','agosto','septiembre','octubre','noviembre','diciembre'];
  const d = new Date();
  setEl('headerDate', `${DAYS[d.getDay()]}, ${d.getDate()} de ${MONTHS[d.getMonth()]} de ${d.getFullYear()}`);
}

function updateLastRefreshTime() {
  const now = new Date().toLocaleTimeString('es', { hour: '2-digit', minute: '2-digit' });
  setEl('lastUpdateTime', `Actualizado ${now}`);
}

/* ================================================================
   THEME
================================================================ */
function toggleTheme() {
  const html   = document.documentElement;
  const isDark = html.getAttribute('data-theme') === 'dark';
  html.setAttribute('data-theme', isDark ? 'light' : 'dark');
  showToast(isDark ? 'Modo claro activado' : 'Modo oscuro activado');
}

/* ================================================================
   PAGE NAVIGATION
================================================================ */
function switchPage(name, btn) {
  document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
  document.querySelectorAll('.nav-btn').forEach(b => b.classList.remove('active'));
  const page = document.getElementById('page' + name.charAt(0).toUpperCase() + name.slice(1));
  if (page) page.classList.add('active');
  if (btn)  btn.classList.add('active');
  const titles = { dashboard: 'Dashboard', presence: 'Estado del equipo' };
  setEl('pageTitle', titles[name] || name);
}

/* ================================================================
   AUTO REFRESH
================================================================ */
setInterval(() => {
  if (document.getElementById('appShell')?.style.display !== 'none') loadAllData();
}, 5 * 60 * 1000);

/* ================================================================
   UI HELPERS
================================================================ */
function showApp() {
  document.getElementById('loginScreen').style.display = 'none';
  document.getElementById('appShell').style.display    = 'flex';
  renderDate();
  renderGreeting('');
}
function showLoading(msg) {
  setEl('loadingText', msg || 'Cargando...');
  document.getElementById('loadingOverlay').style.display = 'flex';
}
function hideLoading() { document.getElementById('loadingOverlay').style.display = 'none'; }
function showError(msg) {
  const el = document.getElementById('loginError');
  if (el) { el.textContent = msg; el.style.display = 'block'; }
}
function toggleUserMenu() { document.getElementById('userMenu').classList.toggle('open'); }
function closeUserMenu()  { document.getElementById('userMenu').classList.remove('open'); }

let _toastTimer;
function showToast(msg) {
  const el = document.getElementById('toast');
  el.textContent = msg; el.classList.add('show');
  clearTimeout(_toastTimer);
  _toastTimer = setTimeout(() => el.classList.remove('show'), 3500);
}

function setEl(id, val) { const el = document.getElementById(id); if (el) el.textContent = val; }
function getEl(id)      { return document.getElementById(id)?.textContent || '0'; }
function escHtml(str) {
  return String(str || '')
    .replace(/&/g,'&amp;').replace(/</g,'&lt;')
    .replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

document.addEventListener('click', e => {
  if (!document.getElementById('navAvatar')?.contains(e.target)) closeUserMenu();
});

/* ================================================================
   INIT
================================================================ */
(async () => {
  // Mostrar login de inmediato
  document.getElementById('loginScreen').style.display = 'flex';

  await initTeams();

  if (_inTeams) {
    document.getElementById('loginScreen').style.display = 'none';
    showApp();
    showLoading('Conectando con Microsoft 365...');
    await loadAllData();
    return;
  }

  // SIEMPRE llamar handleRedirectPromise primero — captura el token tras volver del redirect de login
  let redirectResult = null;
  try {
    redirectResult = await msalInstance.handleRedirectPromise();
  } catch (e) {
    console.warn('[MSAL] handleRedirectPromise error:', e);
    showError('Error al procesar autenticacion: ' + e.message);
  }

  if (redirectResult?.account) {
    // Acabamos de volver de un loginRedirect o acquireTokenRedirect — tenemos cuenta
    currentAccount = redirectResult.account;
    if (redirectResult.accessToken) {
      _cachedToken = redirectResult.accessToken;
      _tokenExpiry = redirectResult.expiresOn?.getTime() || (Date.now() + 3600000);
    }
    document.getElementById('loginScreen').style.display = 'none';
    showApp();
    showLoading('Cargando tu espacio de trabajo...');
    await loadAllData();
    return;
  }

  // Buscar cuenta existente en cache
  const accounts = msalInstance.getAllAccounts();
  if (accounts.length > 0) {
    currentAccount = accounts[0];
    document.getElementById('loginScreen').style.display = 'none';
    showApp();
    showLoading('Recuperando tu sesion...');
    await loadAllData();
    return;
  }

  // Sin sesion: mostrar login
  // (loginScreen ya esta visible)
})();