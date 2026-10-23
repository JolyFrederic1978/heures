// Cache local (IndexedDB) : l'app fonctionne hors ligne, OneDrive est synchronisé dès que possible.
const DB = (() => {
  const NAME = 'heures', STORE = 'kv', KEY = 'state';

  function open() {
    return new Promise((resolve, reject) => {
      const req = indexedDB.open(NAME, 1);
      req.onupgradeneeded = () => req.result.createObjectStore(STORE);
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    });
  }

  async function tx(mode, fn) {
    const db = await open();
    return new Promise((resolve, reject) => {
      const t = db.transaction(STORE, mode);
      const r = fn(t.objectStore(STORE));
      t.oncomplete = () => resolve(r && r.result);
      t.onerror = () => reject(t.error);
    });
  }

  function defaults() {
    return {
      version: 3,
      // Partagé entre appareils via OneDrive
      clients: [{ id: 'c_vedi', name: 'Vedi', billitName: 'V.T.S. SA', defaultOnSite: true, defaultAlescom: false, updatedAt: 0 }],
      entries: [],
      bills: [],
      timer: null, timerAt: 0,
      settings: {
        alescomBillit: 'ALESCOM SRL',
        rates: {
          client:  { h: 90,   m: 1.5,  dep: 60, hLabel: "Main d'œuvre client heures (mai-003)", mLabel: "Main d'œuvre client minutes (mai-004)", depLabel: 'Déplacement forfaitaire client (dep-002)' },
          partner: { h: 76.8, m: 1.28, dep: 55, hLabel: "Main d'œuvre pour partenaire heure (mai-001)", mLabel: "Main d'œuvre partenaire minutes (mai-002)", depLabel: 'Déplacement forfaitaire Partner (dep-001)' }
        }
      },
      settingsAt: 0,
      // Propre à cet appareil
      onedrive: { clientId: '', tenant: 'organizations', folder: 'Compta/TimeSheet' },
      published: {}, lastClientId: null, lastSync: 0
    };
  }

  return {
    async load() {
      const s = await tx('readonly', st => st.get(KEY));
      return s ? { ...defaults(), ...s } : defaults();
    },
    save(state) { return tx('readwrite', st => st.put(state, KEY)); },
    defaults
  };
})();
