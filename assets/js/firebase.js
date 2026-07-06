// Firebase integration — Fase 0
//
// Este archivo es la puerta de entrada al SDK modular de Firebase para el panel.
// Se carga como <script type="module"> para poder importar el SDK v10 vía CDN
// sin necesidad de un build step. Expone la instancia inicializada en
// window.PanelFirebase para que el resto del código (que no es modular) pueda
// consumirla.
//
// FASE 0: solo inicializa el SDK y verifica la conexión con Firestore. No
// escribe ni lee datos reales del panel. Todo el flujo actual (Excel en
// Drive + merge) sigue funcionando idéntico.
//
// FASE 1 y sucesivas: se irán añadiendo aquí las funciones de sync
// (crearRow, actualizarRow, borrarRow, escucharCambios, etc.)

import { initializeApp } from "https://www.gstatic.com/firebasejs/10.13.2/firebase-app.js";
import {
  getFirestore,
  doc,
  setDoc,
  getDoc,
  serverTimestamp,
} from "https://www.gstatic.com/firebasejs/10.13.2/firebase-firestore.js";
import { getAuth } from "https://www.gstatic.com/firebasejs/10.13.2/firebase-auth.js";

const config = window.PANEL_FIREBASE_CONFIG;
if (!config || !config.projectId) {
  console.error("[firebase] falta window.PANEL_FIREBASE_CONFIG — carga config.js antes que firebase.js");
} else {
  try {
    const app = initializeApp(config);
    const db = getFirestore(app);
    const auth = getAuth(app);

    // Exponer al resto del panel (código no-modular) mediante window.
    window.PanelFirebase = {
      app,
      db,
      auth,
      projectId: config.projectId,
      // Utilidades de Firestore que el resto del panel usará. Se exportan
      // desde aquí para que otros archivos no tengan que reimportar del CDN.
      utils: {
        doc,
        setDoc,
        getDoc,
        serverTimestamp,
      },
    };

    console.info(`[firebase] SDK inicializado contra proyecto: ${config.projectId}`);

    // FASE 0 — Test de conexión: escribimos un documento de "heartbeat" del
    // arranque y luego lo leemos. Sirve para verificar que las reglas de
    // seguridad y la config están correctas antes de empezar Fase 1.
    // Este código se elimina cuando la Fase 1 arranque en serio.
    (async () => {
      try {
        const testRef = doc(db, "_bootstrap", "phase0-heartbeat");
        await setDoc(testRef, {
          projectId: config.projectId,
          lastCheck: serverTimestamp(),
        }, { merge: true });
        const snapshot = await getDoc(testRef);
        if (snapshot.exists()) {
          console.info("[firebase] ✅ conexión Firestore verificada (read + write OK)");
        } else {
          console.warn("[firebase] la escritura no dejó documento visible en la lectura posterior");
        }
      } catch (err) {
        console.error("[firebase] ❌ conexión Firestore fallida:", err);
        console.error("[firebase]    revisar Security Rules — probablemente están cerradas por defecto");
      }
    })();

    document.dispatchEvent(new CustomEvent("firebase:ready", { detail: { projectId: config.projectId } }));
  } catch (err) {
    console.error("[firebase] error al inicializar el SDK:", err);
  }
}
