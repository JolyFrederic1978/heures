// Serveur LevelDesk chez OVH (PHP + MySQL, dossier serveur/ du projet ; adresse API_URL dans db.js).
// Pour l'instant : relais Billit (la clé API reste sur le serveur, chiffrée). L'app n'envoie que son jeton Microsoft.
const Serveur = (() => {
  let info = null;   // dernière réponse de « me » : { email, admin, env, billit: { test, live } }

  // Requête brute (Response) : route + paramètres ; body = texte JSON déjà prêt.
  async function call(route, { method = 'GET', query = {}, body } = {}) {
    const u = new URL(API_URL + 'index.php');
    u.searchParams.set('r', route);
    for (const [k, v] of Object.entries(query)) u.searchParams.set(k, v ?? '');
    const headers = { Authorization: 'Bearer ' + await OneDrive.apiToken() };
    if (body) headers['Content-Type'] = 'application/json';
    try { return await fetch(u, { method, headers, body, cache: 'no-store' }); }
    catch (_) { throw new Error('Serveur OVH injoignable (' + new URL(API_URL).host + ').'); }
  }

  async function json(route, { body, ...opts } = {}) {
    const r = await call(route, { ...opts, body: body && JSON.stringify(body) });
    const t = await r.text();
    let d = null;
    try { d = t ? JSON.parse(t) : null; } catch (_) { d = t; }
    if (!r.ok) throw new Error(`Serveur (${r.status}) : ${d?.error || String(t).slice(0, 200)}`);
    return d;
  }

  return {
    call, json,
    info: () => info,
    async me() { return (info = await json('me')); },
    // Enregistre (après contrôle par Billit) ou efface (apiKey vide) la clé d'un mode. Réponse : { billit, companies }.
    async saveBillitKey(mode, apiKey) {
      const r = await json('billit-key', { method: 'POST', body: { mode, apiKey } });
      if (info?.billit) info.billit[mode] = r.billit;
      return r;
    }
  };
})();
