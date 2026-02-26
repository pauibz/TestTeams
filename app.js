/* CONFIGURACIÓN */

const ALLOWED_HOST = "pauibz.github.io";
const ALLOWED_PATH_PREFIX = "/TestTeams";

if (
  location.hostname !== ALLOWED_HOST ||
  !location.pathname.startsWith(ALLOWED_PATH_PREFIX)
) {
  document.body.innerHTML = "<h2>Acceso no autorizado</h2>";
  throw new Error("Dominio no autorizado");
}

if (window.top !== window.self) {
  window.top.location = window.self.location;
}

const CLIENT_ID = "c6219253-8f6e-48c9-8c4a-766aea58a874";
const TENANT_ID = "a3cdf8f7-40db-4d0f-8ccb-82310428392a";

const msalConfig = {
  auth: {
    clientId: CLIENT_ID,
    authority: `https://login.microsoftonline.com/${TENANT_ID}`,
    redirectUri: window.location.origin + ALLOWED_PATH_PREFIX,
    navigateToLoginRequestUrl: false
  },
  cache: {
    cacheLocation: "sessionStorage"
  }
};

const msalInstance = new msal.PublicClientApplication(msalConfig);

const SCOPES = [
  "User.Read",
  "Calendars.Read",
  "Tasks.ReadWrite",
  "Mail.Read",
  "Presence.Read"
];

/* UTILIDADES */

function showLoading(msg) {
  document.getElementById("loadingText").textContent = msg;
  document.getElementById("loadingOverlay").style.display = "block";
}

function hideLoading() {
  document.getElementById("loadingOverlay").style.display = "none";
}

function showError(msg) {
  const el = document.getElementById("loginError");
  el.textContent = msg;
  el.style.display = "block";
}

/* LOGIN */

async function doLogin() {
  try {
    showLoading("Iniciando sesión...");
    await msalInstance.loginPopup({ scopes: SCOPES });
    document.getElementById("loginScreen").style.display = "none";
    document.getElementById("appShell").style.display = "block";
    await loadAllData();
  } catch (err) {
    hideLoading();
    showError("Error de autenticación");
    console.error(err);
  }
}

function doLogout() {
  msalInstance.logoutPopup();
  sessionStorage.clear();
  location.reload();
}

/* TOKEN */

async function getToken() {
  let account = msalInstance.getAllAccounts()[0];

  if (!account) {
    const login = await msalInstance.loginPopup({ scopes: SCOPES });
    account = login.account;
  }

  const token = await msalInstance.acquireTokenSilent({
    scopes: SCOPES,
    account
  });

  return token.accessToken;
}

/* GRAPH */

async function callGraph(endpoint, token) {
  const res = await fetch(
    `https://graph.microsoft.com/v1.0${endpoint}`,
    {
      headers: { Authorization: `Bearer ${token}` }
    }
  );

  if (!res.ok) throw new Error("Graph error");

  return await res.json();
}

/* CARGA DATOS */

async function loadAllData() {
  showLoading("Cargando datos...");

  try {
    const token = await getToken();
    const me = await callGraph("/me", token);

    document.getElementById("greetingText").textContent =
      "Hola " + me.displayName;

  } catch (err) {
    console.error(err);
    showError("Error cargando datos.");
  }

  hideLoading();
}
