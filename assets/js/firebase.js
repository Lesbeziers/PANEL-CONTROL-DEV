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
  initializeFirestore,
  persistentLocalCache,
  persistentMultipleTabManager,
  doc,
  setDoc,
  getDoc,
  getDocs,
  collection,
  serverTimestamp,
  writeBatch,
  onSnapshot,
  deleteDoc,
  addDoc,
  query,
  orderBy,
  limit,
} from "https://www.gstatic.com/firebasejs/10.13.2/firebase-firestore.js";
import {
  getAuth,
  GoogleAuthProvider,
  signInWithPopup,
  signOut,
  onAuthStateChanged,
} from "https://www.gstatic.com/firebasejs/10.13.2/firebase-auth.js";

const config = window.PANEL_FIREBASE_CONFIG;
if (!config || !config.projectId) {
  console.error("[firebase] falta window.PANEL_FIREBASE_CONFIG — carga config.js antes que firebase.js");
} else {
  try {
    const app = initializeApp(config);
    // Firestore con caché local persistente en IndexedDB.
    // - Cada carga del panel tras la primera arranca desde caché (0 lecturas
    //   contra el servidor); el listener sincroniza diferencias en segundo
    //   plano. Reduce el gasto de la cuota Spark ~80-90 % en uso normal.
    // - persistentMultipleTabManager permite compartir la caché entre
    //   varias pestañas del navegador sin errores.
    // - Si el navegador no admite IndexedDB (modo incógnito con restricciones,
    //   Safari con storage bloqueado, etc.), el SDK cae a caché en memoria
    //   automáticamente sin lanzar excepción.
    const db = initializeFirestore(app, {
      localCache: persistentLocalCache({
        tabManager: persistentMultipleTabManager(),
      }),
    });
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
        getDocs,
        collection,
        serverTimestamp,
        writeBatch,
      },
      // Herramientas de migración (Fase 1). No forman parte del flujo normal
      // del panel — se invocan a mano desde la consola del navegador.
      migrateBlocksToFirestore,
      countFirestoreRows,
      clearAllFirestoreRows,
      // Capa de escritura en vivo (Fase 1.3b). El resto del panel las llama
      // cada vez que hay una edición, un insert o un delete.
      writeRowToFirestore,
      softDeleteRowInFirestore,
      syncBlockOrderIndicesToFirestore,
      // Capa de lectura en tiempo real (Fase 1.3c). El panel se suscribe una
      // vez tras el load inicial y recibe callbacks por cada doc que cambia.
      listenToPanelRows,
      // Locks visuales — quién está editando qué celda (Fase 1.3d).
      writeCellLock,
      releaseCellLock,
      listenToCellLocks,
      // Historial "Últimos cambios" en Firestore (Fase 1.3e).
      appendHistoryEntryToFirestore,
      listenToHistoryEntries,
      // Firebase Auth — necesario para las Security Rules restrictivas (Fase 1.4).
      signInPanelUser,
      signOutPanelUser,
      onAuthChanged,
      getCurrentAuthUser,
    };

    // Notificar cambios de sesión al resto del panel.
    onAuthStateChanged(auth, (user) => {
      document.dispatchEvent(new CustomEvent("firebase:auth-changed", { detail: { user } }));
      if (user) {
        console.info(`[auth] sesión activa: ${user.email}`);
      } else {
        console.info("[auth] sin sesión");
      }
    });

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

// =============================================================================
// MIGRATION TOOLS — Fase 1
//
// Estas funciones se llaman a mano desde la consola del navegador. No forman
// parte del flujo normal del panel. Se usan una sola vez para volcar los datos
// actuales del Excel a Firestore.
// =============================================================================

/**
 * Vuelca el array `blocks` en memoria (tal como lo tiene el panel tras cargar
 * el Excel) a Firestore, un documento por fila.
 *
 * Uso desde la consola del navegador:
 *   await PanelFirebase.migrateBlocksToFirestore(blocks)
 *
 * Idempotente: como usa el rowKey como ID de documento, ejecutar dos veces
 * simplemente sobrescribe los documentos con los mismos datos.
 */
async function migrateBlocksToFirestore(blocks, options = {}) {
  const { editor = "sistema-migracion", panelId = "main" } = options;
  if (!window.PanelFirebase?.db) {
    console.error("[migrate] Firebase no está inicializado todavía");
    return null;
  }
  if (!Array.isArray(blocks)) {
    console.error("[migrate] esperaba un array — pasa la variable `blocks` del panel");
    return null;
  }

  const db = window.PanelFirebase.db;
  let count = 0;
  let skipped = 0;
  const now = serverTimestamp();

  // Escribir en batches para eficiencia. Firestore acepta hasta 500 writes por
  // batch; usamos 400 para dejar margen.
  const BATCH_SIZE = 400;
  let batch = writeBatch(db);
  let batchCount = 0;

  const flushBatch = async () => {
    if (batchCount === 0) return;
    await batch.commit();
    batch = writeBatch(db);
    batchCount = 0;
  };

  for (const block of blocks) {
    if (!block || block.isSeparator) continue;
    if (!Array.isArray(block.rows)) continue;

    for (let orderIndex = 0; orderIndex < block.rows.length; orderIndex += 1) {
      const row = block.rows[orderIndex];
      if (!row || row._autoPlaceholder) { skipped += 1; continue; }
      if (!row.rowKey) { skipped += 1; continue; }

      const rowDoc = {
        blockId: block.id,
        title: row.title || "",
        id: row.id || "",
        genre: row.genre || "",
        startDateText: row.startDateText || "",
        startDateISO: row.startDateISO || "",
        endDateText: row.endDateText || "",
        endDateISO: row.endDateISO || "",
        listoByMonth: Object.keys(row.listoByMonth || {}).filter((k) => row.listoByMonth[k]),
        actualizado: !!row.actualizado,
        homeMonth: Number.isInteger(row.homeMonth) ? row.homeMonth : null,
        homeYear: Number.isInteger(row.homeYear) ? row.homeYear : null,
        orderIndex,
        createdAt: now,
        updatedAt: now,
        updatedBy: editor,
        deleted: false,
      };

      const rowRef = doc(db, "panels", panelId, "rows", row.rowKey);
      batch.set(rowRef, rowDoc);
      batchCount += 1;
      count += 1;

      if (batchCount >= BATCH_SIZE) {
        await flushBatch();
        console.info(`[migrate] ${count} filas escritas…`);
      }
    }
  }

  await flushBatch();

  // Documento raíz del panel
  try {
    const panelRef = doc(db, "panels", panelId);
    await setDoc(panelRef, {
      createdAt: now,
      schemaVersion: 1,
      lastMigrationAt: now,
    }, { merge: true });
  } catch (err) {
    console.error("[migrate] error escribiendo panels/main:", err);
  }

  console.info(`[migrate] ✅ Completado: ${count} filas migradas, ${skipped} placeholder/vacías omitidas`);
  return { count, skipped };
}

// =============================================================================
// LIVE WRITE LAYER — Fase 1.3b
//
// Estas funciones sustituyen al flujo de guardado xlsx-a-Drive. Se llaman
// desde app.js cada vez que se edita una celda, se inserta o se borra una
// fila. Escriben directo a Firestore con setDoc({merge:true}) — la primera
// escritura crea el documento, las siguientes solo tocan los campos que
// cambiaron.
// =============================================================================

/**
 * Persiste una fila entera en `panels/{panelId}/rows/{rowKey}`.
 * Idempotente. Si el documento no existe, lo crea; si existe, sobrescribe
 * campos con merge (los campos NO enviados se preservan).
 */
async function writeRowToFirestore(blockId, row, options = {}) {
  const { editor = "anon", panelId = "main" } = options;
  if (!window.PanelFirebase?.db) return false;
  if (!row?.rowKey || !blockId) {
    console.warn("[firestore-write] rowKey o blockId faltan, aborto:", row?.rowKey, blockId);
    return false;
  }
  const db = window.PanelFirebase.db;
  // listoByMonth se guarda como ARRAY de meses activos (no como map).
  // Motivo: setDoc({merge:true}) hace merge PROFUNDO en maps — enviar {} para
  // "borrar todo" no borra nada, se preservan las claves anteriores. Los
  // arrays sí se reemplazan wholesale con merge:true.
  const listoByMonthArray = Object.keys(row.listoByMonth || {})
    .filter((k) => row.listoByMonth[k]);
  const payload = {
    blockId,
    title: row.title || "",
    id: row.id || "",
    genre: row.genre || "",
    startDateText: row.startDateText || "",
    startDateISO: row.startDateISO || "",
    endDateText: row.endDateText || "",
    endDateISO: row.endDateISO || "",
    listoByMonth: listoByMonthArray,
    actualizado: !!row.actualizado,
    homeMonth: Number.isInteger(row.homeMonth) ? row.homeMonth : null,
    homeYear: Number.isInteger(row.homeYear) ? row.homeYear : null,
    deleted: false,
    updatedAt: serverTimestamp(),
    updatedBy: editor,
  };
  if (Number.isInteger(options.orderIndex)) payload.orderIndex = options.orderIndex;
  const rowRef = doc(db, "panels", panelId, "rows", row.rowKey);
  await setDoc(rowRef, payload, { merge: true });
  return true;
}

/**
 * Marca una fila como borrada. NO borra el documento físicamente — así
 * queda auditoría y se puede recuperar. Los loaders ignoran deleted:true.
 */
async function softDeleteRowInFirestore(rowKey, options = {}) {
  const { editor = "anon", panelId = "main" } = options;
  if (!window.PanelFirebase?.db || !rowKey) return false;
  const db = window.PanelFirebase.db;
  const rowRef = doc(db, "panels", panelId, "rows", rowKey);
  await setDoc(rowRef, {
    deleted: true,
    deletedAt: serverTimestamp(),
    deletedBy: editor,
    updatedAt: serverTimestamp(),
    updatedBy: editor,
  }, { merge: true });
  return true;
}

/**
 * Reescribe orderIndex para todas las filas visibles de un bloque, en
 * batch. Se llama después de insertar/borrar filas para mantener el orden
 * consistente en Firestore. Ignora placeholders (los que aún no existen
 * como documento).
 */
async function syncBlockOrderIndicesToFirestore(blockId, rows, options = {}) {
  const { editor = "anon", panelId = "main" } = options;
  if (!window.PanelFirebase?.db || !Array.isArray(rows)) return false;
  const db = window.PanelFirebase.db;
  const now = serverTimestamp();

  const batch = writeBatch(db);
  let count = 0;
  rows.forEach((row, index) => {
    if (!row?.rowKey || row._autoPlaceholder) return;
    const rowRef = doc(db, "panels", panelId, "rows", row.rowKey);
    batch.set(rowRef, {
      blockId,
      orderIndex: index,
      updatedAt: now,
      updatedBy: editor,
    }, { merge: true });
    count += 1;
  });
  if (count === 0) return true;
  await batch.commit();
  return count;
}

// =============================================================================
// LIVE READ LAYER — Fase 1.3c
//
// Suscripción única a la colección `panels/{panelId}/rows`. Cada vez que un
// documento cambia (added / modified / removed) — sea por otro editor o por
// nosotros mismos — se invoca onChange con el evento normalizado. La
// deduplicación de ecos locales (nuestro propio write rebotando) se hace
// aquí filtrando `hasPendingWrites`.
// =============================================================================

/**
 * Abre una suscripción viva a la colección de filas del panel.
 *
 * El callback recibe la batch entera de un solo snapshot en una llamada:
 *
 *   onBatch({ isInitial, changes })
 *     isInitial → true solo la primera vez (equivale al load inicial;
 *                 antes se hacía con getDocs aparte, ya no)
 *     changes   → array de { type, rowKey, data, fromLocal }
 *       type    → "added" | "modified" | "removed"
 *       rowKey  → id del documento (rowKey de la fila)
 *       data    → los campos del documento; null si type === "removed"
 *       fromLocal → true si es un eco de nuestro propio write pendiente de
 *                   confirmar por el servidor (útil para ignorar)
 *
 * Devuelve la función de desuscripción; llamarla cierra el listener.
 */
function listenToPanelRows(onBatch, onError, options = {}) {
  const { panelId = "main" } = options;
  if (!window.PanelFirebase?.db) {
    console.error("[firestore-live] Firebase no está inicializado");
    return null;
  }
  const db = window.PanelFirebase.db;
  const q = collection(db, "panels", panelId, "rows");

  let firstSnapshotHandled = false;
  const unsub = onSnapshot(
    q,
    (snap) => {
      const isInitial = !firstSnapshotHandled;
      firstSnapshotHandled = true;
      const changes = snap.docChanges().map((change) => ({
        type: change.type,
        rowKey: change.doc.id,
        data: change.type === "removed" ? null : change.doc.data(),
        fromLocal: change.doc.metadata.hasPendingWrites,
      }));
      try {
        onBatch({ isInitial, changes });
      } catch (err) {
        console.error("[firestore-live] onBatch threw:", err);
      }
    },
    (err) => {
      console.error("[firestore-live] snapshot error:", err);
      if (typeof onError === "function") onError(err);
    }
  );
  console.info("[firestore-live] suscripción activa a panels/" + panelId + "/rows");
  return unsub;
}

// =============================================================================
// CELL LOCKS — Fase 1.3d
//
// Cuando un editor entra en modo edición de una celda, crea un lock en
// `panels/{panelId}/locks/{lockId}` con su alias, sessionId y updatedAt
// (serverTimestamp). Un heartbeat en cliente reescribe updatedAt cada 10 s
// mientras la edición sigue activa. Al salir, el cliente borra el doc.
//
// Si el cliente muere (crash, cierre brusco), el lock queda huérfano.
// Como safety net, los consumidores ignoran locks con updatedAt mayor de
// 30 s — TTL client-side. No se hace enforcement server-side (rules
// permisivas para dev), solo visual: marco rojo informativo.
// =============================================================================

const LOCK_SEPARATOR = "__";

function lockIdOf(rowKey, columnKey) {
  return `${rowKey}${LOCK_SEPARATOR}${columnKey}`;
}

async function writeCellLock(rowKey, columnKey, options = {}) {
  const { editor = "anon", sessionId = "unknown", panelId = "main" } = options;
  if (!window.PanelFirebase?.db || !rowKey || !columnKey) return false;
  const db = window.PanelFirebase.db;
  const lockRef = doc(db, "panels", panelId, "locks", lockIdOf(rowKey, columnKey));
  await setDoc(lockRef, {
    rowKey,
    columnKey,
    editor,
    sessionId,
    updatedAt: serverTimestamp(),
  }, { merge: true });
  return true;
}

async function releaseCellLock(rowKey, columnKey, options = {}) {
  const { panelId = "main" } = options;
  if (!window.PanelFirebase?.db || !rowKey || !columnKey) return false;
  const db = window.PanelFirebase.db;
  const lockRef = doc(db, "panels", panelId, "locks", lockIdOf(rowKey, columnKey));
  try {
    await deleteDoc(lockRef);
    return true;
  } catch (err) {
    // Si ya no existe, no pasa nada.
    return false;
  }
}

function listenToCellLocks(onChange, onError, options = {}) {
  const { panelId = "main" } = options;
  if (!window.PanelFirebase?.db) return null;
  const db = window.PanelFirebase.db;
  const q = collection(db, "panels", panelId, "locks");
  const unsub = onSnapshot(
    q,
    (snap) => {
      snap.docChanges().forEach((change) => {
        try {
          onChange({
            type: change.type,
            lockId: change.doc.id,
            data: change.type === "removed" ? null : change.doc.data(),
          });
        } catch (err) {
          console.error("[firestore-locks] onChange threw:", err);
        }
      });
    },
    (err) => {
      console.error("[firestore-locks] snapshot error:", err);
      if (typeof onError === "function") onError(err);
    }
  );
  console.info("[firestore-locks] suscripción activa a panels/" + panelId + "/locks");
  return unsub;
}

// =============================================================================
// FIREBASE AUTH — Fase 1.4
//
// Añade una segunda capa de sesión (independiente del OAuth de Drive) contra
// Firebase Auth con proveedor Google. Es la que consumen las Security Rules
// para permitir/denegar writes en Firestore.
//
// Diseño: cuenta compartida (panel.editormp@gmail.com). Las Rules validan
// el email; cualquier otro usuario recibe permission-denied al escribir.
//
// El editor llama a signInPanelUser() cuando ha completado la autenticación
// de Drive. Como el usuario ya está firmado en Google, el popup suele ser
// silencioso o mínimo (Google ofrece la cuenta ya elegida). En sesiones
// posteriores, la sesión Firebase queda cacheada en el navegador y ni
// siquiera pide popup.
// =============================================================================

const googleAuthProvider = (() => {
  const p = new GoogleAuthProvider();
  // Muestra el selector de cuenta si hay más de una (útil para no coger la
  // cuenta personal por defecto).
  p.setCustomParameters({ prompt: "select_account" });
  return p;
})();

async function signInPanelUser() {
  if (!window.PanelFirebase?.auth) {
    console.error("[auth] Firebase no está inicializado");
    return null;
  }
  const auth = window.PanelFirebase.auth;
  // Ya hay sesión previa cacheada — no molestar al usuario.
  if (auth.currentUser) return auth.currentUser;
  try {
    const result = await signInWithPopup(auth, googleAuthProvider);
    return result.user;
  } catch (err) {
    // Casos típicos:
    //   auth/popup-closed-by-user   → el usuario cerró el popup
    //   auth/popup-blocked          → bloqueador de popups activo
    //   auth/operation-not-allowed  → Google provider no está habilitado en la Console
    console.error("[auth] sign-in falló:", err?.code || err);
    return null;
  }
}

async function signOutPanelUser() {
  if (!window.PanelFirebase?.auth) return false;
  try {
    await signOut(window.PanelFirebase.auth);
    return true;
  } catch (err) {
    console.error("[auth] sign-out falló:", err);
    return false;
  }
}

function onAuthChanged(callback) {
  if (!window.PanelFirebase?.auth || typeof callback !== "function") return () => {};
  return onAuthStateChanged(window.PanelFirebase.auth, callback);
}

function getCurrentAuthUser() {
  return window.PanelFirebase?.auth?.currentUser || null;
}

// =============================================================================
// HISTORY LOG — Fase 1.3e
//
// Cada mutación (edit, insert, delete, toggle) añade un doc a
// `panels/{panelId}/history`. El panel lateral "Últimos cambios" se suscribe
// a esa colección y se refresca en tiempo real.
//
// Se guarda un `atLocalMs` (Date.now del cliente) como fallback para pintar
// el "hace X min" inmediatamente sin esperar a que el servidor confirme el
// serverTimestamp. Cuando el server confirma, `at` es la fuente oficial.
// =============================================================================

async function appendHistoryEntryToFirestore(entry, options = {}) {
  const { panelId = "main" } = options;
  if (!window.PanelFirebase?.db || !entry) return null;
  const db = window.PanelFirebase.db;
  try {
    const ref = collection(db, "panels", panelId, "history");
    const docRef = await addDoc(ref, {
      ...entry,
      at: serverTimestamp(),
      atLocalMs: Date.now(),
    });
    return docRef.id;
  } catch (err) {
    console.error("[history] append error:", err);
    return null;
  }
}

function listenToHistoryEntries(onChange, onError, options = {}) {
  const { panelId = "main", limitN = 200 } = options;
  if (!window.PanelFirebase?.db) return null;
  const db = window.PanelFirebase.db;
  const q = query(
    collection(db, "panels", panelId, "history"),
    orderBy("atLocalMs", "desc"),
    limit(limitN)
  );
  const unsub = onSnapshot(
    q,
    (snap) => {
      const entries = [];
      snap.forEach((docSnap) => {
        const data = docSnap.data();
        entries.push({
          _id: docSnap.id,
          ...data,
          // ts como ISO string — el resto del panel espera ese formato.
          ts: data.at?.toDate?.().toISOString?.() ||
              (Number.isFinite(data.atLocalMs) ? new Date(data.atLocalMs).toISOString() : ""),
        });
      });
      try { onChange(entries); }
      catch (err) { console.error("[history] onChange threw:", err); }
    },
    (err) => {
      console.error("[history] snapshot error:", err);
      if (typeof onError === "function") onError(err);
    }
  );
  console.info("[history] suscripción activa a panels/" + panelId + "/history");
  return unsub;
}

/**
 * Cuenta cuántos documentos hay en `panels/main/rows`. Útil para verificar
 * después de la migración.
 *
 *   const n = await PanelFirebase.countFirestoreRows()
 */
async function countFirestoreRows(panelId = "main") {
  if (!window.PanelFirebase?.db) return null;
  const db = window.PanelFirebase.db;
  const snap = await getDocs(collection(db, "panels", panelId, "rows"));
  console.info(`[migrate] Firestore contiene ${snap.size} filas en panels/${panelId}/rows`);
  return snap.size;
}

/**
 * Borra TODOS los documentos de `panels/main/rows`. Uso: reiniciar la
 * migración si hay que rehacerla. IRREVERSIBLE. Solo funciona sobre el
 * proyecto DEV; en PROD las reglas lo impedirán.
 *
 *   await PanelFirebase.clearAllFirestoreRows()
 */
async function clearAllFirestoreRows(panelId = "main") {
  if (!window.PanelFirebase?.db) return null;
  const db = window.PanelFirebase.db;
  const snap = await getDocs(collection(db, "panels", panelId, "rows"));
  console.warn(`[migrate] borrando ${snap.size} filas de panels/${panelId}/rows…`);

  const BATCH_SIZE = 400;
  let batch = writeBatch(db);
  let batchCount = 0;

  for (const docSnap of snap.docs) {
    batch.delete(docSnap.ref);
    batchCount += 1;
    if (batchCount >= BATCH_SIZE) {
      await batch.commit();
      batch = writeBatch(db);
      batchCount = 0;
    }
  }
  if (batchCount > 0) await batch.commit();

  console.info(`[migrate] ✅ Borrado ${snap.size} filas`);
  return snap.size;
}
