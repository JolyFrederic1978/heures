// Connexion Microsoft 365 (OAuth2 code + PKCE, sans bibliothèque) et accès OneDrive via Microsoft Graph.
const OneDrive = (() => {
  const TK = 'heures.tokens', PK = 'heures.pkce';
  const SCOPE = 'openid profile offline_access Files.ReadWrite';
  let cfg = () => ({});

  class NeedLogin extends Error { constructor() { super('Connexion OneDrive requise.'); } }

  const redirectUri = () => new URL('./', location.href).href;
  const authBase = () => `https://login.microsoftonline.com/${cfg().tenant || 'organizations'}/oauth2/v2.0`;
  const b64url = buf => btoa(String.fromCharCode(...new Uint8Array(buf))).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
  const tokens = () => { try { return JSON.parse(localStorage.getItem(TK)); } catch (_) { return null; } };

  async function login() {
    if (!cfg().clientId) throw new Error("Renseigne d'abord l'ID d'application (client ID) dans Réglages.");
    const verifier = b64url(crypto.getRandomValues(new Uint8Array(32)));
    const challenge = b64url(await crypto.subtle.digest('SHA-256', new TextEncoder().encode(verifier)));
    const st = b64url(crypto.getRandomValues(new Uint8Array(12)));
    localStorage.setItem(PK, JSON.stringify({ verifier, st }));
    const p = new URLSearchParams({
      client_id: cfg().clientId, response_type: 'code', redirect_uri: redirectUri(), scope: SCOPE,
      code_challenge: challenge, code_challenge_method: 'S256', state: st, response_mode: 'query'
    });
    if (tokens()?.user) p.set('login_hint', tokens().user);
    location.assign(`${authBase()}/authorize?${p}`);
  }

  function logout() { localStorage.removeItem(TK); }

  async function tokenRequest(params) {
    const r = await fetch(`${authBase()}/token`, {
      method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({ client_id: cfg().clientId, scope: SCOPE, ...params })
    });
    const j = await r.json();
    if (!r.ok) { const e = new Error(j.error_description || j.error); e.invalid = r.status === 400; throw e; }
    let user = tokens()?.user;
    try { user = JSON.parse(atob(j.id_token.split('.')[1].replace(/-/g, '+').replace(/_/g, '/'))).preferred_username; } catch (_) { }
    const t = { access: j.access_token, refresh: j.refresh_token, exp: Date.now() + (j.expires_in - 120) * 1000, user };
    localStorage.setItem(TK, JSON.stringify(t));
    return t;
  }

  // À appeler au démarrage : termine la connexion si on revient de la page Microsoft.
  async function handleRedirect() {
    const u = new URL(location.href);
    const code = u.searchParams.get('code'), err = u.searchParams.get('error');
    if (!code && !err) return false;
    history.replaceState(null, '', redirectUri());
    if (err) throw new Error(u.searchParams.get('error_description') || err);
    const pk = JSON.parse(localStorage.getItem(PK) || 'null');
    localStorage.removeItem(PK);
    if (!pk || pk.st !== u.searchParams.get('state')) throw new Error('Réponse de connexion invalide, réessaie.');
    await tokenRequest({ grant_type: 'authorization_code', code, redirect_uri: redirectUri(), code_verifier: pk.verifier });
    return true;
  }

  async function accessToken() {
    const t = tokens();
    if (!t) throw new NeedLogin();
    if (t.exp > Date.now()) return t.access;
    if (!t.refresh) throw new NeedLogin();
    try { return (await tokenRequest({ grant_type: 'refresh_token', refresh_token: t.refresh })).access; }
    catch (e) { if (e.invalid) { logout(); throw new NeedLogin(); } throw e; }   // hors ligne : on garde le jeton
  }

  async function graph(path, opts = {}) {
    const tk = await accessToken();
    return fetch('https://graph.microsoft.com/v1.0' + path, { ...opts, headers: { Authorization: 'Bearer ' + tk, ...(opts.headers || {}) } });
  }
  const itemPath = p => `/me/drive/root:/${p.split('/').filter(Boolean).map(encodeURIComponent).join('/')}:`;

  // Renvoie { data, etag } (data en texte ou ArrayBuffer) ou null si le fichier n'existe pas.
  async function get(path, as = 'text') {
    // Pas de $select : OneDrive Entreprise omet alors @microsoft.graph.downloadUrl.
    const r = await graph(itemPath(path));
    if (r.status === 404) return null;
    if (!r.ok) throw new Error(`OneDrive (${r.status}) : lecture de ${path} impossible.`);
    const meta = await r.json();
    const url = meta['@microsoft.graph.downloadUrl'];
    const c = url ? await fetch(url, { cache: 'no-store' }) : await graph(itemPath(path) + '/content', { cache: 'no-store' });
    if (!c.ok) throw new Error(`OneDrive (${c.status}) : téléchargement de ${path} impossible.`);
    return { data: as === 'buffer' ? await c.arrayBuffer() : await c.text(), etag: meta.eTag };
  }

  // Écrit le fichier (dossiers créés au besoin).
  // etag : échoue (412) si modifié entre-temps ; overwrite : remplace sans condition ; sinon : création seulement.
  async function put(path, body, { etag, type = 'application/octet-stream', overwrite = false } = {}) {
    const headers = { 'Content-Type': type };
    if (etag) headers['If-Match'] = etag; else if (!overwrite) headers['If-None-Match'] = '*';
    const r = await graph(itemPath(path) + '/content', { method: 'PUT', headers, body });
    if (r.status === 423) throw new Error(`${path.split('/').pop()} est ouvert/verrouillé dans OneDrive : réessai à la prochaine synchro.`);
    return r;
  }

  // Fichiers d'un dossier : [{ name, size, modified }] ([] si le dossier n'existe pas).
  async function list(path) {
    let url = itemPath(path) + '/children?$top=200', out = [];
    while (url) {
      const r = url.startsWith('http') ? await fetch(url, { headers: { Authorization: 'Bearer ' + await accessToken() } }) : await graph(url);
      if (r.status === 404) return [];
      if (!r.ok) throw new Error(`OneDrive (${r.status}) : lecture du dossier ${path} impossible.`);
      const j = await r.json();
      out.push(...j.value.filter(x => x.file).map(x => ({ name: x.name, size: x.size, modified: x.lastModifiedDateTime })));
      url = j['@odata.nextLink'];
    }
    return out;
  }

  // Crée le dossier (et ses parents) s'il n'existe pas.
  async function mkdir(path) {
    const parts = path.split('/').filter(Boolean);
    for (let i = 0; i < parts.length; i++) {
      const parent = parts.slice(0, i).join('/');
      const r = await graph((parent ? itemPath(parent) : '/me/drive/root') + '/children', {
        method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ name: parts[i], folder: {}, '@microsoft.graph.conflictBehavior': 'fail' })
      });
      if (!r.ok && r.status !== 409) throw new Error(`OneDrive (${r.status}) : création du dossier ${parts.slice(0, i + 1).join('/')} impossible.`);
    }
  }

  // Déplace un fichier dans un autre dossier (créé au besoin), en le renommant si un fichier du même nom y est déjà.
  // Renvoie le nom final du fichier.
  async function move(path, folder, newName) {
    await mkdir(folder);
    const name = newName || path.split('/').pop();
    const r = await graph(itemPath(path) + '?@microsoft.graph.conflictBehavior=rename', {
      method: 'PATCH', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ parentReference: { path: '/drive/root:/' + folder.split('/').filter(Boolean).join('/') }, name })
    });
    if (!r.ok) throw new Error(`OneDrive (${r.status}) : déplacement de ${name} impossible.`);
    try { return (await r.json()).name || name; } catch (_) { return name; }
  }

  return {
    NeedLogin, login, logout, handleRedirect, get, put, list, mkdir, move, redirectUri,
    configure(fn) { cfg = fn; },
    configured: () => !!cfg().clientId,
    signedIn: () => !!tokens(),
    user: () => tokens()?.user || ''
  };
})();
