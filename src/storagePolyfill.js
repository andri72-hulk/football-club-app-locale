// Ricrea l'API window.storage usata dall'app appoggiandosi a IndexedDB, il
// database integrato nel browser stesso. Nessun servizio esterno, nessun
// login, nessuna connessione internet richiesta dopo il primo caricamento
// della pagina. Stessa identica interfaccia esterna (get/set/delete/list
// con chiave + flag "shared") già usata dalle versioni cloud — qui "shared"
// è solo una distinzione logica tra due spazi separati nello stesso database
// locale, non un vero meccanismo di condivisione tra persone diverse.

const DB_NAME = "football_club_app";
const STORE_NAME = "kv";
const DB_VERSION = 1;

function openDB() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, DB_VERSION);
    req.onupgradeneeded = () => {
      const db = req.result;
      if (!db.objectStoreNames.contains(STORE_NAME)) {
        db.createObjectStore(STORE_NAME);
      }
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

function idbKey(key, shared) {
  return `${shared ? "shared" : "personal"}:${key}`;
}

function idbGet(key, shared) {
  return openDB().then(
    (db) =>
      new Promise((resolve, reject) => {
        const tx = db.transaction(STORE_NAME, "readonly");
        const req = tx.objectStore(STORE_NAME).get(idbKey(key, shared));
        req.onsuccess = () => resolve(req.result);
        req.onerror = () => reject(req.error);
      })
  );
}

function idbSet(key, shared, value) {
  return openDB().then(
    (db) =>
      new Promise((resolve, reject) => {
        const tx = db.transaction(STORE_NAME, "readwrite");
        tx.objectStore(STORE_NAME).put(value, idbKey(key, shared));
        tx.oncomplete = () => resolve();
        tx.onerror = () => reject(tx.error);
      })
  );
}

function idbDelete(key, shared) {
  return openDB().then(
    (db) =>
      new Promise((resolve, reject) => {
        const tx = db.transaction(STORE_NAME, "readwrite");
        tx.objectStore(STORE_NAME).delete(idbKey(key, shared));
        tx.oncomplete = () => resolve();
        tx.onerror = () => reject(tx.error);
      })
  );
}

function idbListKeys(prefix, shared) {
  return openDB().then(
    (db) =>
      new Promise((resolve, reject) => {
        const tx = db.transaction(STORE_NAME, "readonly");
        const req = tx.objectStore(STORE_NAME).getAllKeys();
        req.onsuccess = () => {
          const nsPrefix = `${shared ? "shared" : "personal"}:`;
          const fullPrefix = nsPrefix + prefix;
          const keys = req.result
            .filter((k) => typeof k === "string" && k.startsWith(fullPrefix))
            .map((k) => k.slice(nsPrefix.length));
          resolve(keys);
        };
        req.onerror = () => reject(req.error);
      })
  );
}

const storagePolyfill = {
  async get(key, shared = false) {
    const value = await idbGet(key, shared);
    if (value === undefined) {
      const err = new Error(`Storage key not found: ${key}`);
      err.code = "NOT_FOUND";
      throw err;
    }
    return { key, value, shared: !!shared };
  },

  async set(key, value, shared = false) {
    await idbSet(key, shared, value);
    return { key, value, shared: !!shared };
  },

  async delete(key, shared = false) {
    await idbDelete(key, shared);
    return { key, deleted: true, shared: !!shared };
  },

  async list(prefix = "", shared = false) {
    const keys = await idbListKeys(prefix, shared);
    return { keys, prefix, shared: !!shared };
  },
};

if (typeof window !== "undefined") {
  window.storage = storagePolyfill;
}

export default storagePolyfill;
