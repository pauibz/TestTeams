¡Perfecto! Para conectarlo a la Graph API real necesitas seguir 3 pasos. Te explico todo:

Paso 1 — Registrar la app en Azure AD

Ve a portal.azure.com → Microsoft Entra ID → Registros de aplicaciones → Nueva registro
Rellena:

Nombre: Teams Dashboard
Tipo de cuenta: Cuentas de este directorio organizativo (o multitenant si lo necesitas)
URI de redirección: Web → http://localhost (o la URL donde hosteas el HTML)


Al crearse, copia el Application (client) ID y el Directory (tenant) ID — los necesitarás en el código
Ve a Permisos de API → Agregar permiso → Microsoft Graph → Permisos delegados y añade:

Calendars.Read
Tasks.ReadWrite (Planner)
Mail.Read
Presence.Read
User.Read


Haz clic en Conceder consentimiento de administrador


Paso 2 — Configurar MSAL.js en el HTML
Reemplaza la sección <script> del HTML actual con esto (manteniendo el mismo HTML/CSS):
html<!-- Añadir en el <head> -->
<script src="https://alcdn.msauth.net/browser/2.38.0/js/msal-browser.min.js"></script>
javascript// ── CONFIGURACIÓN MSAL ──
const msalConfig = {
  auth: {
    clientId: "TU_CLIENT_ID_AQUÍ",
    authority: "https://login.microsoftonline.com/TU_TENANT_ID_AQUÍ",
    redirectUri: window.location.origin,
  },
  cache: { cacheLocation: "sessionStorage" }
};

const msalInstance = new msal.PublicClientApplication(msalConfig);

const SCOPES = [
  "User.Read",
  "Calendars.Read",
  "Tasks.ReadWrite",
  "Mail.Read",
  "Presence.Read"
];

// ── LOGIN Y OBTENCIÓN DE TOKEN ──
async function getToken() {
  const accounts = msalInstance.getAllAccounts();
  if (accounts.length === 0) {
    // Redirige al login de Microsoft
    await msalInstance.loginPopup({ scopes: SCOPES });
  }
  const account = msalInstance.getAllAccounts()[0];
  const result = await msalInstance.acquireTokenSilent({ scopes: SCOPES, account });
  return result.accessToken;
}

// ── LLAMADAS A GRAPH API ──
async function callGraph(endpoint, token) {
  const res = await fetch(`https://graph.microsoft.com/v1.0${endpoint}`, {
    headers: { Authorization: `Bearer ${token}` }
  });
  return res.json();
}

Paso 3 — Reemplazar los datos mock por Graph API
Sustituye la función // ── INIT ── por esta:
javascriptasync function init() {
  try {
    const token = await getToken();

    // Datos del usuario
    const me = await callGraph("/me", token);
    document.getElementById('greetingText').textContent =
      `${getGreeting()}, ${me.displayName.split(' ')[0]} 👋`;

    // Reuniones de hoy
    const today = new Date();
    const start = today.toISOString().split('T')[0] + 'T00:00:00';
    const end   = today.toISOString().split('T')[0] + 'T23:59:59';
    const calData = await callGraph(
      `/me/calendarView?startDateTime=${start}&endDateTime=${end}&$orderby=start/dateTime&$top=10`, token
    );
    renderMeetingsFromGraph(calData.value);

    // Tareas de Planner
    const tasksData = await callGraph("/me/planner/tasks?$filter=percentComplete lt 100", token);
    renderTasksFromGraph(tasksData.value);

    // Correos
    const mailData = await callGraph(
      "/me/messages?$top=5&$orderby=receivedDateTime desc&$filter=isRead eq false", token
    );
    renderMailFromGraph(mailData.value);

    // Presencia
    const presenceData = await callGraph(`/me/presence`, token);
    updatePresenceFromGraph(presenceData);

  } catch (err) {
    console.error("Error conectando con Microsoft Graph:", err);
  }
}

// ── ADAPTADORES (Graph → render) ──

function renderMeetingsFromGraph(events) {
  const el = document.getElementById('meetingsList');
  if (!events.length) { el.innerHTML = '<p style="padding:16px;color:#888">No hay reuniones hoy</p>'; return; }
  el.innerHTML = events.map(e => {
    const start = new Date(e.start.dateTime + 'Z');
    const end   = new Date(e.end.dateTime + 'Z');
    const h = start.getHours(), m = start.getMinutes();
    const dur = Math.round((end - start) / 60000);
    const live = isLive(h, m, dur);
    const past = isPast(h, m, dur);
    return `
      <div class="meeting-item">
        <div class="meeting-time-col">
          <span class="meeting-time">${pad(h)}:${pad(m)}</span>
          <span class="meeting-end">${pad(end.getHours())}:${pad(end.getMinutes())}</span>
        </div>
        <div class="meeting-bar ${live ? 'green' : past ? 'gray' : ''}"></div>
        <div class="meeting-details">
          <div class="meeting-title">${e.subject}</div>
          <div class="meeting-sub">${e.organizer?.emailAddress?.name || ''} · ${dur} min</div>
          <div class="meeting-chips">
            ${live ? '<span class="chip live">● En directo</span>' : ''}
          </div>
        </div>
        <button class="join-btn ${past ? 'disabled' : ''}" 
          onclick="${e.onlineMeeting?.joinUrl ? `window.open('${e.onlineMeeting.joinUrl}')` : ''}" 
          ${past ? 'disabled' : ''}>
          ${past ? 'Finalizada' : 'Unirse'}
        </button>
      </div>`;
  }).join('');
}

function renderTasksFromGraph(taskItems) {
  const el = document.getElementById('tasksList');
  if (!taskItems.length) { el.innerHTML = '<p style="padding:16px;color:#888">Sin tareas pendientes 🎉</p>'; return; }
  el.innerHTML = taskItems.map(t => `
    <div class="task-item">
      <div class="task-check ${t.percentComplete === 100 ? 'done' : ''}"></div>
      <div class="task-info">
        <div class="task-name ${t.percentComplete === 100 ? 'done' : ''}">${t.title}</div>
        <div class="task-meta">Vence: ${t.dueDateTime ? new Date(t.dueDateTime.dateTime).toLocaleDateString('es') : 'Sin fecha'}</div>
      </div>
      <div class="priority-dot ${t.priority <= 2 ? 'high' : t.priority <= 5 ? 'medium' : 'low'}"></div>
    </div>`).join('');
}

function renderMailFromGraph(messages) {
  const colors = ["#6264A7","#0078D4","#107C10","#C4314B","#8A5A00"];
  const el = document.getElementById('mailList');
  el.innerHTML = messages.map((m, i) => {
    const from = m.from?.emailAddress?.name || 'Desconocido';
    const initials = from.split(' ').map(w=>w[0]).join('').slice(0,2).toUpperCase();
    const time = new Date(m.receivedDateTime).toLocaleTimeString('es', {hour:'2-digit', minute:'2-digit'});
    return `
      <div class="mail-item">
        <div class="mail-avatar" style="background:${colors[i%colors.length]}">${initials}</div>
        <div class="mail-content">
          <div class="mail-from">${from}</div>
          <div class="mail-subject ${!m.isRead ? 'unread' : ''}">${m.subject}</div>
          <div class="mail-preview">${m.bodyPreview}</div>
        </div>
        <div class="mail-right">
          <span class="mail-time">${time}</span>
          ${!m.isRead ? '<div class="unread-dot"></div>' : ''}
        </div>
      </div>`;
  }).join('');
}

function updatePresenceFromGraph(data) {
  const map = {
    Available:  { color: '#6BB700', label: '🟢 Disponible' },
    Busy:       { color: '#C4314B', label: '🔴 Ocupada' },
    Away:       { color: '#F4BE22', label: '🟡 Ausente' },
    Offline:    { color: '#C0C0C0', label: '⚫ No disponible' },
    DoNotDisturb: { color: '#C4314B', label: '🔴 No molestar' },
  };
  const p = map[data.availability] || map['Offline'];
  document.getElementById('presenceIndicator').style.background = p.color;
  document.getElementById('presenceStatusText').textContent = p.label;
}

// Reemplaza la última línea del init
init();

Notas importantes
TemaDetalleTeams como pestañaUsa @microsoft/teams-js + microsoftTeams.authentication.authenticate() en lugar de loginPopupPresence APIRequiere licencia Teams activa en el tenantPlanner TasksEl endpoint puede devolver tareas de todos los planes; filtra por planId si necesitas uno específicoCORS localPara pruebas locales usa Live Server de VS Code o python -m http.server, no abras el HTML directamente como file://