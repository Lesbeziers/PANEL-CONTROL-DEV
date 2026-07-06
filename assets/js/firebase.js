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
  getDocs,
  collection,
  serverTimestamp,
  writeBatch,
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
        listoByMonth: row.listoByMonth || {},
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
