// Configuración del entorno DEV.
// En la rama `dev` este archivo apunta al Excel de test y al proyecto Firebase de test.
// En la rama `main` este archivo apunta al Excel real y al proyecto Firebase de producción.
// Los dos proyectos Firebase son completamente independientes y nunca se cruzan.

window.PANEL_CONFIG = {
  // Google Drive (legacy, se mantiene mientras convivimos con Drive hasta cutover a Firebase)
  GOOGLE_CLIENT_ID: "679270086294-uforgvhb3j32mp2pst4gu148sgnrmtek.apps.googleusercontent.com",
  GOOGLE_API_KEY: "AIzaSyCfP4msV5hM8D3nCoGGvF56lhWHuKMNSPQ",
  GOOGLE_DRIVE_FILE_ID: "1_ftcdXkF2Pa6NQouFVvNXuhqqMODqUxI",
  GOOGLE_DRIVE_FILE_NAME: "PANEL_CONTROL_DATA_TEST.xlsx",
};

// Firebase — proyecto DEV.
// Autogenerado por Firebase Console. Estas claves son públicas por diseño
// (los proyectos web las exponen); la seguridad real está en las Security Rules
// de Firestore, no en ocultar la apiKey.
window.PANEL_FIREBASE_CONFIG = {
  apiKey: "AIzaSyA_-YF4mIb98gnLiWg7LeTg7aLUi-IXkH4",
  authDomain: "panel-control-movistar-dev.firebaseapp.com",
  projectId: "panel-control-movistar-dev",
  storageBucket: "panel-control-movistar-dev.firebasestorage.app",
  messagingSenderId: "506263058708",
  appId: "1:506263058708:web:cfab26d630018e8b906a01",
};
