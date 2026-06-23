(function () {
  const config = window.PANEL_CONFIG;
  if (!config) {
    console.error("[gdrive] PANEL_CONFIG missing — load config.js before gdrive.js");
    return;
  }

  const FILE_ID = config.GOOGLE_DRIVE_FILE_ID;
  const API_KEY = config.GOOGLE_API_KEY;
  const CLIENT_ID = config.GOOGLE_CLIENT_ID;
  // Project number = primera parte del CLIENT_ID antes del primer guion.
  // Necesario para que el Picker asocie el grant drive.file con este cliente OAuth.
  const APP_ID = CLIENT_ID.split("-")[0];
  const SCOPE = "https://www.googleapis.com/auth/drive.file";
  const MIME_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

  let tokenClient = null;
  let accessToken = null;
  let tokenExpiresAt = 0;
  let pendingTokenResolvers = [];
  let pickerLoaded = false;
  let refreshTimer = null;

  // Persistencia del token durante la sesión del navegador (sessionStorage): así
  // recargar la página NO vuelve a pedir login mientras la pestaña siga viva.
  // sessionStorage se borra al cerrar la pestaña y el token caduca igualmente.
  const TOKEN_STORAGE_KEY = `gdriveToken:${FILE_ID}`;

  function hasSessionStorage() {
    try {
      return typeof window !== "undefined" && !!window.sessionStorage;
    } catch (_) {
      return false;
    }
  }

  function persistToken() {
    if (!hasSessionStorage() || !accessToken) return;
    try {
      window.sessionStorage.setItem(
        TOKEN_STORAGE_KEY,
        JSON.stringify({ accessToken, tokenExpiresAt })
      );
    } catch (_) { /* ignore */ }
  }

  function clearPersistedToken() {
    if (!hasSessionStorage()) return;
    try {
      window.sessionStorage.removeItem(TOKEN_STORAGE_KEY);
    } catch (_) { /* ignore */ }
  }

  function restoreToken() {
    if (!hasSessionStorage()) return false;
    try {
      const raw = window.sessionStorage.getItem(TOKEN_STORAGE_KEY);
      if (!raw) return false;
      const parsed = JSON.parse(raw);
      if (!parsed?.accessToken || !(Date.now() < Number(parsed.tokenExpiresAt))) {
        clearPersistedToken();
        return false;
      }
      accessToken = parsed.accessToken;
      tokenExpiresAt = Number(parsed.tokenExpiresAt);
      // Reprogramar el refresco según el tiempo restante (con margen de 60 s).
      scheduleTokenRefresh(Math.max((tokenExpiresAt - Date.now()) / 1000 + 60, 60));
      return true;
    } catch (_) {
      return false;
    }
  }

  // Refresca el token EN SILENCIO un par de minutos antes de que caduque, para
  // que el usuario no se tope con un re-login a media sesión al guardar (que es
  // donde se arriesgaba a perder trabajo). prompt:"" no muestra UI si la sesión
  // de Google sigue activa.
  function scheduleTokenRefresh(expiresInSec) {
    if (refreshTimer) {
      clearTimeout(refreshTimer);
      refreshTimer = null;
    }
    const leadSec = 120;
    const delayMs = Math.max(Number(expiresInSec) - leadSec, 30) * 1000;
    refreshTimer = setTimeout(() => {
      if (!tokenClient) return;
      try {
        tokenClient.requestAccessToken({ prompt: "" });
      } catch (e) {
        console.warn("[gdrive] refresco silencioso de token falló:", e);
      }
    }, delayMs);
  }

  window.GoogleDrive = {
    init,
    isSignedIn,
    signIn,
    signOut,
    showGate,
    hideGate,
    loadXlsxBuffer,
    saveXlsxBuffer,
    getAppProperties,
    patchAppProperties,
  };

  function init() {
    if (tokenClient) return true;
    if (!window.google?.accounts?.oauth2) return false;
    tokenClient = window.google.accounts.oauth2.initTokenClient({
      client_id: CLIENT_ID,
      scope: SCOPE,
      callback: handleTokenResponse,
    });
    return true;
  }

  async function handleTokenResponse(resp) {
    if (resp.error) {
      console.error("[gdrive] token error:", resp);
      pendingTokenResolvers.forEach((p) => p.reject(new Error(resp.error)));
      pendingTokenResolvers = [];
      return;
    }
    accessToken = resp.access_token;
    tokenExpiresAt = Date.now() + (Number(resp.expires_in) - 60) * 1000;
    scheduleTokenRefresh(resp.expires_in);
    persistToken();

    // drive.file: comprobar que el archivo del panel está autorizado.
    // Si no lo está (primer login con esta cuenta), mostrar Picker para que
    // el usuario lo seleccione → eso concede el permiso per-file.
    try {
      const accessible = await checkFileAccess(accessToken);
      if (!accessible) {
        hideGate();
        await showPickerForAuth(accessToken);
      }
    } catch (err) {
      console.error("[gdrive] file authorization failed:", err);
      showGate();
      pendingTokenResolvers.forEach((p) => p.reject(err));
      pendingTokenResolvers = [];
      accessToken = null;
      tokenExpiresAt = 0;
      return;
    }

    hideGate();
    document.dispatchEvent(new CustomEvent("gdrive:signedin"));
    pendingTokenResolvers.forEach((p) => p.resolve(accessToken));
    pendingTokenResolvers = [];
  }

  function isSignedIn() {
    return !!accessToken && Date.now() < tokenExpiresAt;
  }

  function signIn(opts = {}) {
    if (!init()) {
      console.error("[gdrive] GIS not loaded yet, cannot sign in");
      return;
    }
    // prompt="" → Google muestra solo lo estrictamente necesario:
    //   - Si nunca se ha consentido: muestra consentimiento.
    //   - Si ya hay consent server-side: lo salta (mejor UX).
    tokenClient.requestAccessToken({ prompt: "" });
  }

  function signOut() {
    if (refreshTimer) {
      clearTimeout(refreshTimer);
      refreshTimer = null;
    }
    if (accessToken && window.google?.accounts?.oauth2?.revoke) {
      window.google.accounts.oauth2.revoke(accessToken, () => {});
    }
    accessToken = null;
    tokenExpiresAt = 0;
    clearPersistedToken();
    showGate();
  }

  function ensureToken() {
    if (isSignedIn()) return Promise.resolve(accessToken);
    return new Promise((resolve, reject) => {
      pendingTokenResolvers.push({ resolve, reject });
      try {
        signIn({ silent: true });
      } catch (e) {
        reject(e);
      }
    });
  }

  async function checkFileAccess(token) {
    const url = `https://www.googleapis.com/drive/v3/files/${FILE_ID}?fields=id`;
    const resp = await fetch(url, {
      headers: { Authorization: `Bearer ${token}` },
      credentials: "omit",
    });
    return resp.ok;
  }

  function loadPicker() {
    return new Promise((resolve, reject) => {
      if (pickerLoaded) return resolve();
      if (!window.gapi) {
        return reject(new Error("gapi no cargado — falta <script src='https://apis.google.com/js/api.js'>"));
      }
      window.gapi.load("picker", {
        callback: () => { pickerLoaded = true; resolve(); },
        onerror: () => reject(new Error("No se pudo cargar Google Picker")),
      });
    });
  }

  async function showPickerForAuth(token) {
    await loadPicker();

    // El primer uso de la API key + Picker desde un edge concreto de Google
    // a veces dispara "La clave de desarrollador no es válida" por
    // propagación de cachés. Se manifiesta como CANCEL inmediato del picker.
    // Reintentamos en silencio hasta 2 veces si detectamos cierre rápido.
    const MAX_ATTEMPTS = 3;
    const TRANSIENT_THRESHOLD_MS = 2500;
    const RETRY_DELAY_MS = 1500;

    const tryShow = (attempt) => new Promise((resolve, reject) => {
      const openedAt = Date.now();

      const view = new google.picker.DocsView()
        .setIncludeFolders(false)
        .setMimeTypes(MIME_XLSX)
        .setQuery("PANEL_CONTROL_DATA")
        .setMode(google.picker.DocsViewMode.LIST);

      const picker = new google.picker.PickerBuilder()
        .addView(view)
        .setOAuthToken(token)
        .setDeveloperKey(API_KEY)
        .setAppId(APP_ID)
        .setTitle("Selecciona el archivo del Panel de Control")
        .setLocale("es")
        .setCallback((data) => {
          if (data.action === google.picker.Action.PICKED) {
            const picked = data.docs && data.docs[0];
            if (picked && picked.id === FILE_ID) {
              resolve();
            } else {
              reject(new Error(
                `El archivo seleccionado no es el correcto. ` +
                `Por favor, selecciona PANEL_CONTROL_DATA.xlsx`
              ));
            }
          } else if (data.action === google.picker.Action.CANCEL) {
            const elapsed = Date.now() - openedAt;
            const looksTransient = elapsed < TRANSIENT_THRESHOLD_MS && attempt < MAX_ATTEMPTS;
            if (looksTransient) {
              console.warn(
                `[gdrive] Picker cerrado en ${elapsed}ms (intento ${attempt}/${MAX_ATTEMPTS}). ` +
                `Probable error transitorio de Google, reintentando…`
              );
              setTimeout(() => {
                tryShow(attempt + 1).then(resolve).catch(reject);
              }, RETRY_DELAY_MS);
            } else {
              reject(new Error("Selección de archivo cancelada"));
            }
          }
        })
        .build();
      picker.setVisible(true);
    });

    return tryShow(1);
  }

  async function loadXlsxBuffer({ useAuth = false } = {}) {
    if (useAuth) {
      let token = await ensureToken();
      const url = `https://www.googleapis.com/drive/v3/files/${FILE_ID}?alt=media`;
      const doFetch = (t) => fetch(url, {
        headers: { Authorization: `Bearer ${t}` },
        credentials: "omit",
      });
      let resp = await doFetch(token);
      // Si el token restaurado de sesión estuviera caducado/revocado (401),
      // forzar re-login interactivo y reintentar una vez.
      if (resp.status === 401) {
        accessToken = null;
        tokenExpiresAt = 0;
        clearPersistedToken();
        token = await ensureToken();
        resp = await doFetch(token);
      }
      if (!resp.ok) throw new Error(`Drive fetch failed: HTTP ${resp.status}`);
      return await resp.arrayBuffer();
    }
    const url = `https://www.googleapis.com/drive/v3/files/${FILE_ID}?alt=media&key=${API_KEY}`;
    const resp = await fetch(url, { credentials: "omit" });
    if (!resp.ok) throw new Error(`Drive fetch failed: HTTP ${resp.status}`);
    return await resp.arrayBuffer();
  }

  async function saveXlsxBuffer(buffer) {
    let token = await ensureToken();
    let resp = await uploadOnce(buffer, token);
    if (resp.status === 401) {
      accessToken = null;
      tokenExpiresAt = 0;
      clearPersistedToken();
      token = await ensureToken();
      resp = await uploadOnce(buffer, token);
    }
    if (!resp.ok) {
      const errText = await resp.text().catch(() => "");
      throw new Error(`Drive upload failed: HTTP ${resp.status} ${errText}`);
    }
    return await resp.json().catch(() => ({}));
  }

  function uploadOnce(buffer, token) {
    const url = `https://www.googleapis.com/upload/drive/v3/files/${FILE_ID}?uploadType=media`;
    return fetch(url, {
      method: "PATCH",
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": MIME_XLSX,
      },
      body: buffer,
    });
  }

  // Fetch the file's appProperties bag (private metadata visible only to this
  // OAuth client). Used for the presence indicator — each browser session
  // writes its heartbeat there and reads everyone else's.
  async function getAppProperties() {
    const token = await ensureToken();
    const url = `https://www.googleapis.com/drive/v3/files/${FILE_ID}?fields=appProperties`;
    const resp = await fetch(url, {
      headers: { Authorization: `Bearer ${token}` },
      credentials: "omit",
    });
    if (!resp.ok) {
      throw new Error(`Drive appProperties GET failed: HTTP ${resp.status}`);
    }
    const data = await resp.json();
    return data.appProperties || {};
  }

  // Patch the file's appProperties. Drive merges the supplied keys with the
  // existing bag; setting a key value to null removes it. We use this to
  // publish our heartbeat and to evict our own entry on tab close.
  async function patchAppProperties(updates) {
    const token = await ensureToken();
    const url = `https://www.googleapis.com/drive/v3/files/${FILE_ID}?fields=appProperties`;
    const resp = await fetch(url, {
      method: "PATCH",
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({ appProperties: updates }),
      credentials: "omit",
    });
    if (!resp.ok) {
      throw new Error(`Drive appProperties PATCH failed: HTTP ${resp.status}`);
    }
    return await resp.json();
  }

  let gateInjected = false;

  function showGate() {
    ensureGateInjected();
    document.getElementById("gdrive-gate").classList.add("visible");
    const app = document.getElementById("app");
    if (app) app.style.display = "none";
  }

  function hideGate() {
    if (!gateInjected) return;
    const gate = document.getElementById("gdrive-gate");
    if (gate) gate.classList.remove("visible");
    const app = document.getElementById("app");
    if (app) app.style.display = "";
  }

  function ensureGateInjected() {
    if (gateInjected) return;
    const style = document.createElement("style");
    style.textContent = `
      #gdrive-gate {
        position: fixed; inset: 0;
        background: rgba(15, 18, 28, 0.92);
        display: none; align-items: center; justify-content: center;
        z-index: 9999; font-family: system-ui, -apple-system, sans-serif;
      }
      #gdrive-gate.visible { display: flex; }
      #gdrive-gate .gdrive-card {
        background: #fff; padding: 32px 36px; border-radius: 14px;
        box-shadow: 0 8px 40px rgba(0,0,0,0.5); max-width: 440px; text-align: center;
      }
      #gdrive-gate h2 {
        margin: 0 0 12px; font-size: 22px; color: #1f2937;
      }
      #gdrive-gate p {
        margin: 0 0 24px; color: #4b5563; line-height: 1.55; font-size: 14.5px;
      }
      #gdrive-gate button {
        background: #1a73e8; color: #fff; border: 0;
        padding: 12px 28px; border-radius: 6px; font-size: 15px; font-weight: 600;
        cursor: pointer; transition: background 0.15s;
      }
      #gdrive-gate button:hover { background: #1557b0; }
      #gdrive-gate button:disabled { opacity: 0.6; cursor: wait; }
    `;
    document.head.appendChild(style);

    const gate = document.createElement("div");
    gate.id = "gdrive-gate";
    gate.innerHTML = `
      <div class="gdrive-card">
        <h2>Acceso al editor</h2>
        <p>Para editar el Panel de Control necesitas iniciar sesión con la cuenta autorizada.</p>
        <button id="gdrive-signin-btn" type="button">Iniciar sesión</button>
      </div>
    `;
    document.body.appendChild(gate);
    document.getElementById("gdrive-signin-btn").addEventListener("click", () => {
      signIn({ silent: false });
    });
    gateInjected = true;
  }

  function tryAutoInit(retriesLeft = 50) {
    if (init()) return;
    if (retriesLeft <= 0) return;
    setTimeout(() => tryAutoInit(retriesLeft - 1), 100);
  }

  // Recuperar el token de la sesión del navegador ANTES de que app.js consulte
  // isSignedIn() en su arranque: si sigue vigente, no se vuelve a pedir login.
  restoreToken();
  tryAutoInit();
})();
