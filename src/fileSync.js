// Sincronizzazione tramite File System Access API: l'app collega un file
// (tipicamente dentro una cartella sincronizzata da Google Drive Desktop o
// OneDrive) e lo tiene aggiornato automaticamente ad ogni salvataggio.
// Aprendo l'app su un altro dispositivo con lo stesso file sincronizzato,
// i dati più recenti vengono caricati automaticamente.
//
// Funziona solo su Chrome, Edge e Opera (non Firefox/Safari): isFileSyncSupported()
// va sempre controllato prima di mostrare l'opzione all'utente.

const HANDLE_DB_NAME = "football_club_app_filesync";
const HANDLE_STORE = "handles";
const HANDLE_KEY = "syncFileHandle";

function openHandleDB() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(HANDLE_DB_NAME, 1);
    req.onupgradeneeded = () => {
      req.result.createObjectStore(HANDLE_STORE);
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

async function saveHandle(handle) {
  const db = await openHandleDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(HANDLE_STORE, "readwrite");
    tx.objectStore(HANDLE_STORE).put(handle, HANDLE_KEY);
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  });
}

export async function loadSyncHandle() {
  const db = await openHandleDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(HANDLE_STORE, "readonly");
    const req = tx.objectStore(HANDLE_STORE).get(HANDLE_KEY);
    req.onsuccess = () => resolve(req.result || null);
    req.onerror = () => reject(req.error);
  });
}

export async function clearSyncHandle() {
  const db = await openHandleDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(HANDLE_STORE, "readwrite");
    tx.objectStore(HANDLE_STORE).delete(HANDLE_KEY);
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  });
}

export function isFileSyncSupported() {
  return typeof window !== "undefined" && "showSaveFilePicker" in window;
}

// Crea un nuovo file di sincronizzazione (prima configurazione su un dispositivo)
export async function createSyncFile() {
  const handle = await window.showSaveFilePicker({
    suggestedName: "football-club-sync.json",
    types: [{ description: "File di sincronizzazione", accept: { "application/json": [".json"] } }],
  });
  await saveHandle(handle);
  return handle;
}

// Collega un file di sincronizzazione già esistente (es. su un secondo dispositivo,
// puntando allo stesso file già creato e sincronizzato dal primo)
export async function linkExistingSyncFile() {
  const [handle] = await window.showOpenFilePicker({
    types: [{ description: "File di sincronizzazione", accept: { "application/json": [".json"] } }],
  });
  await saveHandle(handle);
  return handle;
}

export async function verifySyncPermission(handle, readWrite = true) {
  const options = readWrite ? { mode: "readwrite" } : {};
  if ((await handle.queryPermission(options)) === "granted") return true;
  if ((await handle.requestPermission(options)) === "granted") return true;
  return false;
}

export async function readSyncFile(handle) {
  const file = await handle.getFile();
  const text = await file.text();
  return text ? JSON.parse(text) : null;
}

export async function writeSyncFile(handle, dataObj) {
  const writable = await handle.createWritable();
  await writable.write(JSON.stringify(dataObj));
  await writable.close();
}
