// Environnement : production = gestion.onelevelup.be (OVH) ; toute autre adresse (dev-gestion sur GitHub, localhost) = développement.
// En développement : données dans Compta-DEV, Billit en sandbox uniquement, bandeau DEV.
const ENV = location.hostname === 'gestion.onelevelup.be' ? 'prod' : 'dev';
const DEFAULT_FOLDER = ENV === 'prod' ? 'Compta/TimeSheet' : 'Compta-DEV/TimeSheet';

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
          zone1:   { dep: 20, depLabel: 'Déplacement client zone 1 (dep-003)' },
          partner: { h: 76.8, m: 1.28, dep: 55, hLabel: "Main d'œuvre pour partenaire heure (mai-001)", mLabel: "Main d'œuvre partenaire minutes (mai-002)", depLabel: 'Déplacement forfaitaire Partner (dep-001)' }
        }
      },
      settingsAt: 0,
      // Propre à cet appareil
      // Inscription Entra « Heures OLU » (One Level UP sprl) – identifiants publics, pas des secrets.
      onedrive: { clientId: '0b486fd1-e8f9-4933-94cd-49c5df20d6d6', tenant: 'eb79b929-d806-4ec9-8282-a93399a8ca66', folder: DEFAULT_FOLDER },
      published: {}, lastClientId: null, lastSync: 0
    };
  }

  return {
    async load() {
      const s = await tx('readonly', st => st.get(KEY));
      if (!s) return defaults();
      const d = defaults();
      if (!s.onedrive?.clientId) s.onedrive = { ...d.onedrive, folder: s.onedrive?.folder || d.onedrive.folder };
      return { ...d, ...s };
    },
    save(state) { return tx('readwrite', st => st.put(state, KEY)); },
    defaults
  };
})();
