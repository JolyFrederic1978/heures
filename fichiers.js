// © 2026 One Level UP – LevelDesk. Logiciel propriétaire, tous droits réservés. Copie, modification, redistribution et rétro-ingénierie interdites sans autorisation écrite (voir LICENCE.txt).
// Fichiers de la société (Excel de facturation, fichiers Prestations ; plus tard PDF des achats) : OneDrive du compte, ou serveur LevelDesk
// (architecture.md, section 10 ; phase 3a). Choix dans Réglages → Ma société → Fichiers : « auto » (OneDrive si le compte en a un, sinon le serveur),
// « onedrive » ou « server ». Sans serveur (fichiers OneDrive, production actuelle) : toujours OneDrive, rien ne change.
// Même forme de réponse que OneDrive.put / get / list, pour que l'app n'ait pas à savoir où va le fichier.
const Fichiers = (() => {
  let cfg = () => ({});   // { srv, store, folder }
  // Sur le serveur, les chemins partent de la racine de la société : le dossier OneDrive (« Compta/TimeSheet ») est retiré.
  const rel = p => { const f = String(cfg().folder || '').replace(/\/+$/, ''); return f && p.startsWith(f + '/') ? p.slice(f.length + 1) : p.replace(/^\/+/, ''); };
  async function where() {
    const c = cfg();
    if (!c.srv) return 'onedrive';
    if (c.store === 'server' || c.store === 'onedrive') return c.store;
    // Serveur seulement si Microsoft confirme que le compte n'a pas de OneDrive ; en cas de doute (réseau), OneDrive comme avant.
    return (await OneDrive.hasDrive().catch(() => null)) === false ? 'server' : 'onedrive';
  }
  async function toB64(body) {
    const buf = body instanceof Blob ? await body.arrayBuffer() : body instanceof ArrayBuffer ? body : new TextEncoder().encode(String(body)).buffer;
    const u = new Uint8Array(buf);
    let s = '';
    for (let i = 0; i < u.length; i += 0x8000) s += String.fromCharCode(...u.subarray(i, i + 0x8000));
    return btoa(s);
  }
  const fromB64 = s => Uint8Array.from(atob(s || ''), c => c.charCodeAt(0)).buffer;
  const status = e => +(/\((\d{3})\)/.exec(e.message)?.[1]) || 500;
  const op = body => Serveur.json('fichiers', { method: 'POST', body });

  return {
    configure(fn) { cfg = fn; },
    where,
    label: w => w === 'server' ? 'serveur LevelDesk' : 'OneDrive',
    // { ok, status, where } (comme une réponse fetch) ; etag (OneDrive) ignoré sur le serveur.
    async put(path, body, opts = {}) {
      const w = await where();
      if (w === 'onedrive') { const r = await OneDrive.put(path, body, opts); return { ok: r.ok, status: r.status, where: w }; }
      try {
        await op({ op: 'put', path: rel(path), content: await toB64(body), mime: opts.type || '', overwrite: !!(opts.overwrite || opts.etag) });
        return { ok: true, status: 200, where: w };
      } catch (e) { return { ok: false, status: status(e), where: w, error: e.message }; }
    },
    // { data, etag } ou null si le fichier n'existe pas.
    async get(path, as = 'text') {
      if (await where() === 'onedrive') return OneDrive.get(path, as);
      const r = await op({ op: 'get', path: rel(path) });
      if (!r.exists) return null;
      const buf = fromB64(r.content);
      return { data: as === 'buffer' ? buf : new TextDecoder().decode(buf), etag: null };
    },
    // Fichiers rangés sur le serveur (Réglages → Fichiers) : { files: [{ path, size, at }], used, quota }.
    serverList: (dir = '') => op({ op: 'list', dir }),
    async serverDownload(path) {
      const r = await op({ op: 'get', path });
      if (!r.exists) throw new Error('Fichier introuvable sur le serveur.');
      return { name: r.name, blob: new Blob([fromB64(r.content)], { type: r.mime }) };
    },
    serverDelete: path => op({ op: 'delete', path })
  };
})();
