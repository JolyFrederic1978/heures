// LevelDesk (anciennement OLU Gestion, puis Heures OLU) – pointage des interventions, facturation et archivage, synchronisés sur OneDrive.
// Les noms techniques gardent « heures » (heures-data.json, base locale, cache) : les changer ferait perdre données et connexion.
// Source de vérité : App/heures-data.json. Les fichiers Excel sont régénérés à partir de ces données :
//   Interventions-A-Facturer.xlsx, Archives/Archive-Interventions.xlsx, A facturer/Prestations - Client date.xlsx
let state;
let tab = 'today';
let installPrompt = null;
let tick;
let syncStatus = { kind: 'off', msg: '' };

// ---------- utilitaires ----------
const $ = (s, root = document) => root.querySelector(s);
const $$ = (s, root = document) => [...root.querySelectorAll(s)];
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const uid = p => p + '_' + Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
const pad = n => String(n).padStart(2, '0');
const isoDate = d => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
const hhmm = d => `${pad(d.getHours())}:${pad(d.getMinutes())}`;
const today = () => isoDate(new Date());
const frDate = d => d ? d.split('-').reverse().join('/') : '';
const longDate = d => new Date(d + 'T00:00').toLocaleDateString('fr-BE', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' });
const fmtDur = min => `${Math.floor(min / 60)}:${pad(min % 60)}`;
const eur = n => new Intl.NumberFormat('fr-BE', { style: 'currency', currency: 'EUR' }).format(n);
const toMin = hm => { const [h, m] = (hm || '0:0').split(':').map(Number); return h * 60 + m; };
const fromMin = m => { m = ((m % 1440) + 1440) % 1440; return `${pad(Math.floor(m / 60))}:${pad(m % 60)}`; };
const diffMin = (a, b) => { let d = toMin(b) - toMin(a); if (d < 0) d += 1440; return d; };
const fold = s => String(s || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '');   // recherche sans accents
const hash = s => { let h = 5381; for (let i = 0; i < s.length; i++) h = (h * 33 ^ s.charCodeAt(i)) >>> 0; return h.toString(36) + s.length; };

function parseDur(s) {
  s = (s || '').trim().toLowerCase().replace(/\s/g, '');
  let m;
  if (s === '') return 0;
  if ((m = /^(\d+)[:h](\d{1,2})?$/.exec(s))) return +m[1] * 60 + +(m[2] || 0);
  if ((m = /^(\d+)[,.](\d+)$/.exec(s))) return Math.round(parseFloat(m[1] + '.' + m[2]) * 60);
  if (/^\d+$/.test(s)) return +s;              // nombre seul = minutes
  return NaN;
}

function toast(msg) {
  const t = $('#toast'); t.textContent = msg; t.classList.remove('hidden');
  clearTimeout(toast.t); toast.t = setTimeout(() => t.classList.add('hidden'), 2800);
}
// --- dictée vocale : reconnaissance du navigateur (Chrome → Google, Edge → Microsoft ; absente de Firefox : pas de bouton) ---
// mic(id) : bouton 🎤 qui ajoute le texte dicté au champ id ; submit = envoie ensuite le formulaire du champ (ajout rapide des tâches).
const Speech = window.SpeechRecognition || window.webkitSpeechRecognition;
const mic = (id, submit) => Speech ? `<button type="button" class="btn chip mic" data-mic="${id}" ${submit ? 'data-mic-submit' : ''} title="Dicter">🎤</button>` : '';
const lblMic = (text, id) => `<div class="lblmic"><label>${text}</label>${mic(id)}</div>`;
let dictation = null;
document.addEventListener('click', ev => {
  const b = ev.target.closest('[data-mic]'); if (!b) return;
  ev.preventDefault(); ev.stopPropagation();
  if (dictation) { dictation.stop(); return; }
  const el = document.getElementById(b.dataset.mic); if (!el || el.disabled) return;
  const r = new Speech(); r.lang = 'fr-BE'; r.interimResults = true; r.continuous = false;
  const base = el.value.replace(/\s+$/, ''), cap = !base || /[.!?]$/.test(base);
  let heard = '';
  r.onresult = e => {
    heard = [...e.results].map(x => x[0].transcript).join('').trim();
    if (cap) heard = heard.charAt(0).toUpperCase() + heard.slice(1);
    el.value = base + (base && heard ? ' ' : '') + heard;
    el.dispatchEvent(new Event('input', { bubbles: true }));
  };
  r.onerror = e => toast(e.error === 'not-allowed' ? 'Micro refusé : autorise-le dans le navigateur.' : e.error === 'no-speech' ? 'Rien entendu.' : 'Dictée impossible (' + e.error + ').');
  r.onend = () => {
    dictation = null; b.classList.remove('on');
    if (!heard) return;
    el.dispatchEvent(new Event('change', { bubbles: true }));
    if (b.hasAttribute('data-mic-submit')) el.form?.requestSubmit();
  };
  dictation = r; b.classList.add('on'); r.start();
}, true);
function download(blob, name) {
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob); a.download = name;
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(a.href), 5000);
}

async function persist() { await DB.save(state); render(); scheduleSync(); }

const live = () => state.entries.filter(e => !e.deleted);
const client = id => state.clients.find(c => c.id === id) || { name: '(client supprimé)', billitName: '' };
const activeClients = () => state.clients.filter(c => !c.deleted && !c.archived).sort((a, b) => a.name.localeCompare(b.name));
const sortEntries = es => es.sort((a, b) => (a.date + a.start).localeCompare(b.date + b.start));
const bill = id => state.bills.find(b => b.id === id);
const liveBills = () => state.bills.filter(b => !b.deleted);
const touch = o => { o.updatedAt = Date.now(); return o; };
function setTimer(t) { state.timer = t; state.timerAt = Date.now(); }

// ---------- sociétés (commanditaires) et utilisateurs ----------
// One Level UP = société « maison » (id 'olu') ; Alescom = premier commanditaire (id 'alescom' : tarifs partenaire, échéance 10 j).
// Une intervention a un commanditaire (cmdId) et un technicien (techId). Anciennes données : case alescom, technicien = administrateur.
// La case alescom reste écrite (cmdId === 'alescom') pour les appareils qui ont encore une ancienne version.
const HOME = 'olu', ADMIN = 'u_admin';
const liveSocs = () => state.societes.filter(s => !s.deleted).sort((a, b) => (b.id === HOME) - (a.id === HOME) || a.name.localeCompare(b.name));
const activeSocs = () => liveSocs().filter(s => s.active !== false);
const soc = id => state.societes.find(s => s.id === id) || { id, name: '(société supprimée)' };
const liveUsers = () => state.users.filter(u => !u.deleted).sort((a, b) => a.name.localeCompare(b.name));
const user = id => state.users.find(u => u.id === id) || { id, name: '(utilisateur supprimé)' };
// Utilisateur connecté : reconnu par son adresse Microsoft. Compte inconnu, désactivé ou supprimé : aucun droit (écran « accès refusé »).
// Seule exception : tant que l'administrateur n'a pas d'adresse (toute 1re connexion), c'est lui (ensureSocietes la renseigne alors).
const GUEST = { id: 'u_guest', name: 'Compte non reconnu', roles: [], clients: {} };
const me = () => {
  const m = fold(OneDrive.user()), u = m && state.users.find(x => x.email && fold(x.email) === m);
  if (u) return u.deleted || u.active === false ? GUEST : u;
  return user(ADMIN).email ? GUEST : user(ADMIN);
};
const isAdmin = () => (me().roles || ['admin']).includes('admin');
const cmdOf = e => e.cmdId || (e.alescom ? 'alescom' : HOME);
const techOf = e => e.techId || ADMIN;
const techSoc = e => state.users.find(u => u.id === techOf(e))?.societeId || HOME;
const clientSoc = c => c?.societeId || (c?.defaultAlescom ? 'alescom' : HOME);
// Facturé par One Level UP : commanditaire One Level UP (au client) ou technicien One Level UP pour un autre commanditaire (au commanditaire).
const billable = e => cmdOf(e) === HOME || techSoc(e) === HOME;
const socRates = s => s.rates || state.settings.rates.partner;
// Clients visibles par l'utilisateur connecté (réglés dans Administration, par société) ; un administrateur voit tout.
function canSee(c) {
  if (isAdmin()) return true;
  const v = me().clients?.[clientSoc(c)];
  return v === 'all' || (Array.isArray(v) && v.includes(c.id));
}

// Sociétés et administrateur créés une fois, à partir des réglages existants (ID fixes : pas de doublon entre appareils).
function ensureSocietes() {
  state.societes ||= []; state.users ||= [];
  const s = state.settings;
  if (!state.societes.some(x => x.id === HOME)) state.societes.push({ id: HOME, name: 'One Level UP', home: true, active: true, updatedAt: 0 });
  if (!state.societes.some(x => x.id === 'alescom')) state.societes.push({ id: 'alescom', name: 'Alescom', active: true, tenant: '',
    billitName: s.alescomBillit || 'ALESCOM SRL', billitIds: { ...s.alescomBillitIds }, rates: { ...s.rates.partner }, due: 10, updatedAt: 0 });
  let a = state.users.find(u => u.id === ADMIN);
  if (!a) state.users.push(a = { id: ADMIN, name: 'Fred', email: '', societeId: HOME, roles: ['admin', 'resp', 'tech'], active: true, clients: {}, updatedAt: 0 });
  if (!a.email && OneDrive.user()) { a.email = OneDrive.user(); touch(a); }
  // Clients sans société : One Level UP (ou Alescom s'ils avaient « alescom par défaut »), enregistré dans la fiche.
  state.clients.forEach(c => { if (!c.societeId) { c.societeId = c.defaultAlescom ? 'alescom' : HOME; touch(c); } });
}
const socOptions = sel => {
  const l = activeSocs(), cur = sel && !l.some(s => s.id === sel) && state.societes.find(s => s.id === sel);
  return (cur ? `<option value="${cur.id}" selected>${esc(cur.name)} (inactive)</option>` : '')
    + l.map(s => `<option value="${s.id}" ${s.id === sel ? 'selected' : ''}>${esc(s.name)}</option>`).join('');
};
const userLabel = u => u.name + (u.societeId && u.societeId !== HOME ? ` (${soc(u.societeId).name})` : '');

// Un seul client par société : « Vedi distance » / « Afm remote » sont fusionnés dans « Vedi » / « Afm »,
// leurs interventions passent en « à distance » (case Sur site décochée). Renvoie true si quelque chose a changé.
function mergeRemoteClients() {
  let changed = false;
  for (const c of state.clients.filter(x => !x.deleted && /\b(remote|distance)\b/i.test(x.name))) {
    const name = c.name.replace(/\b(à|a)\s+distance\b/i, '').replace(/\b(remote|distance)\b/ig, '').replace(/\s+/g, ' ').trim();
    if (!name) continue;
    let base = state.clients.find(x => x !== c && !x.deleted && x.name.toLowerCase() === name.toLowerCase());
    if (!base) { Object.assign(touch(c), { name, defaultOnSite: false }); changed = true; continue; }
    state.entries.forEach(e => { if (e.clientId === c.id) Object.assign(touch(e), { clientId: base.id, onSite: false }); });
    if (state.timer?.clientId === c.id) { state.timer.clientId = base.id; state.timerAt = Date.now(); }
    if (state.lastClientId === c.id) state.lastClientId = base.id;
    c.deleted = true; touch(c); changed = true;
  }
  return changed;
}

// ---------- facturation ----------
// Un groupe de facturation = un client (commanditaire One Level UP), ou un autre commanditaire qui regroupe plusieurs clients
// ('alescom' : clé des factures existantes ; 's:<id>' pour les suivants).
const groupKey = e => { const c = cmdOf(e); return c === HOME ? 'c:' + e.clientId : c === 'alescom' ? 'alescom' : 's:' + c; };
function groupInfo(key) {
  const rates = state.settings.rates;
  if (!key.startsWith('c:')) {
    const s = soc(key === 'alescom' ? 'alescom' : key.slice(2));
    return { key, label: s.name, billitName: s.billitName || s.name, rates: socRates(s), partner: true, soc: s, due: s.due ?? 15 };
  }
  const c = client(key.slice(2));
  // Client proche : le déplacement est facturé avec l'article zone 1 (dep-003).
  const r = c.near ? { ...rates.client, dep: rates.zone1.dep, depLabel: rates.zone1.depLabel, depRef: rates.zone1.depRef } : rates.client;
  return { key, label: c.name, billitName: c.billitName || c.name, rates: r, partner: false, near: !!c.near, due: 15 };
}
function totals(entries, r) {
  let h = 0, m = 0, dep = 0;
  for (const e of entries) { h += Math.floor(e.minutes / 60); m += e.minutes % 60; dep += e.onSite ? 1 : 0; }
  return { h, m, dep, amount: h * r.h + m * r.m + dep * r.dep, minutes: entries.reduce((s, e) => s + e.minutes, 0) };
}
const entryAmount = e => totals([e], groupInfo(groupKey(e)).rates).amount;

// Valeurs ajoutées après la première version : complétées si absentes (données déjà synchronisées).
function normalizeSettings() {
  const r = state.settings.rates;
  r.zone1 ||= { dep: 20, depLabel: 'Déplacement client zone 1 (dep-003)' };
  // Codes des articles Billit utilisés pour chaque ligne de facture.
  Object.assign(r.client, { hRef: 'mai-003', mRef: 'mai-004', depRef: 'dep-002', ...pick(r.client) });
  Object.assign(r.partner, { hRef: 'mai-001', mRef: 'mai-002', depRef: 'dep-001', ...pick(r.partner) });
  r.zone1.depRef ||= 'dep-003';
  state.settings.billit ||= { mode: 'test', test: { apiKey: '', partyId: '' }, live: { apiKey: '', partyId: '' } };
  // Clés API Billit : propres à chaque appareil (state.billitKeys, jamais écrites dans heures-data.json, effacées à la déconnexion).
  // Une clé encore présente dans les réglages synchronisés (versions ≤ v42) est reprise ici puis retirée : la synchro suivante réécrit le fichier sans elle.
  state.billitKeys ||= { test: '', live: '' };
  for (const k of ['test', 'live']) {
    const b = state.settings.billit[k] ||= { apiKey: '', partyId: '' };
    if (b.apiKey) { state.billitKeys[k] ||= b.apiKey; b.apiKey = ''; }
  }
  // Clés sur le serveur OVH (réglage synchronisé) : celles de l'appareil sont effacées.
  if (state.settings.billit.relay) state.billitKeys = { test: '', live: '' };
  ensureSocietes();
}
// Échéance des factures Billit (jours) : 15 j pour les clients, celle du commanditaire sinon (Alescom 10 j) ; à modifier au besoin dans Billit.
const dueDays = info => info.due ?? 15;
const pick = o => Object.fromEntries(['hRef', 'mRef', 'depRef'].filter(k => o[k]).map(k => [k, o[k]]));

// Liste de clients fournie par Fred (oct. 2026) : ajoutée une seule fois, sans doublon (comparaison sans casse).
const CLIENTS_SEED = ['Ace', 'Acofisc', 'AFM', 'All Clean', 'Atelier C', 'BAEV', 'BC Béton', 'BDS', 'BHE', 'BKLZ', 'Burima',
  'CDDH', 'Clinquart', 'Clio', 'Defitec', 'Deroover', 'Dewilde', 'Dufaux', 'Duthoit', 'Enercetec', 'Energen', 'Home', 'Hugo',
  'INT', 'JMD', 'Krack', 'LCTS', 'Marvelle', 'MCS', 'Metalik Pro', 'MF', 'Nextgen', 'P4P', 'PAPV', 'Pitau', 'Quertinmont',
  'SFER', 'Trav', 'Vedi', 'Viart', 'Warnier'];
function seedClients() {
  if (state.settings.clientsSeed >= 1) return false;
  const have = new Set(state.clients.filter(c => !c.deleted).map(c => c.name.toLowerCase()));
  for (const name of CLIENTS_SEED) {
    if (have.has(name.toLowerCase())) continue;
    state.clients.push(touch({ id: uid('c'), name, billitName: '', near: false, societeId: HOME, server: false, servers: 0 }));
    have.add(name.toLowerCase());
  }
  state.settings.clientsSeed = 1; state.settingsAt = Date.now();
  return true;
}
const xrow = e => ({ ...e, client: client(e.clientId).name, cmd: cmdOf(e) === HOME ? '' : soc(cmdOf(e)).name, tech: user(techOf(e)).name });

const prestationsName = b => `Prestations - ${b.label} ${b.date}.xlsx`;
// Facture à un commanditaire : colonne client en plus.
const prestationsFile = b => Xlsx.prestations(sortEntries(live().filter(e => e.billId === b.id)).map(xrow), !b.key.startsWith('c:'));

// Achats refacturés (module Achats, achats.js) : ils peuvent être ajoutés à la facture d'un client.
const achSum = ach => ach.reduce((s, x) => s + x.total, 0);
const achLabel = x => `${x.ref} ${x.label}${x.domain ? ' – ' + x.domain : ''}`;
// Objet de la facture : interventions (période) et/ou achats refacturés.
function billTitle(b) {
  const parts = [];
  if (b.from) parts.push(`Interventions du ${frDate(b.from)} au ${frDate(b.to)}`);
  if (Achats.ofBill(b.id).length) parts.push(parts.length ? 'abonnements' : 'Abonnements et services');
  return parts.join(' et ');
}

function recapBillit(b) {
  const info = groupInfo(b.key), r = info.rates;
  const s = sortEntries(live().filter(e => e.billId === b.id)), t = totals(s, r), ach = Achats.ofBill(b.id);
  const lines = [`Client Billit : ${info.billitName}`, `Objet : ${billTitle(b)}`, ''];
  if (t.h) lines.push(`${r.hLabel} : ${t.h} × ${eur(r.h)} = ${eur(t.h * r.h)}`);
  if (t.m) lines.push(`${r.mLabel} : ${t.m} × ${eur(r.m)} = ${eur(t.m * r.m)}`);
  if (t.dep) lines.push(`${r.depLabel} : ${t.dep} × ${eur(r.dep)} = ${eur(t.dep * r.dep)}`);
  ach.forEach(x => lines.push(`${achLabel(x)} : ${x.qty} × ${eur(x.unit)} = ${eur(x.total)}`));
  const inst = s.filter(e => e.install);
  if (inst.length) lines.push('', 'Installations à ajouter :', ...inst.map(e => `- ${frDate(e.date)} ${e.install}`));
  lines.push('', `Total HTVA (hors installations) : ${eur(t.amount + achSum(ach))}`);
  if (s.length) lines.push(`Pièce jointe : ${prestationsName(b)}`);
  return lines.join('\n');
}
async function copy(txt) {
  try { await navigator.clipboard.writeText(txt); toast('Copié'); } catch (_) { prompt('Copie ce texte :', txt); }
}

// ---------- synchronisation OneDrive ----------
const P = {
  data: () => `${state.onedrive.folder}/App/heures-data.json`,
  open: () => `${state.onedrive.folder}/Interventions-A-Facturer.xlsx`,
  archive: () => `${state.onedrive.folder}/Archives/Archive-Interventions.xlsx`,
  prest: b => `${state.onedrive.folder}/A facturer/${prestationsName(b)}`,
  template: () => `${state.onedrive.folder}/Archives/ArchiveTemplate.xlsx`
};

function setSync(kind, msg = '') {
  syncStatus = { kind, msg };
  const p = $('#syncPill');
  p.textContent = { off: 'Local', login: 'Non connecté', busy: 'Synchro…', ok: '☁ OneDrive', err: '⚠ Erreur', offline: 'Hors ligne' }[kind];
  p.title = msg;
  if ($('#odStatus')) $('#odStatus').textContent = statusText();
}
const statusText = () => syncStatus.kind === 'ok'
  ? `Synchronisé à ${new Date(state.lastSync).toLocaleTimeString('fr-BE', { hour: '2-digit', minute: '2-digit' })} · fichiers Excel à jour`
  : (syncStatus.msg || $('#syncPill').textContent);

function mergeRemote(remote) {
  const before = JSON.stringify([state.entries, state.clients, state.bills, state.timer, state.settings, state.societes, state.users]);
  const mergeArr = (a = [], b = []) => {
    const m = new Map(a.map(x => [x.id, x]));
    for (const x of b) { const y = m.get(x.id); if (!y || (x.updatedAt || 0) > (y.updatedAt || 0)) m.set(x.id, x); }
    return [...m.values()];
  };
  state.entries = mergeArr(state.entries, remote.entries);
  state.clients = mergeArr(state.clients, remote.clients);
  state.bills = mergeArr(state.bills, remote.bills);
  state.societes = mergeArr(state.societes, remote.societes);
  state.users = mergeArr(state.users, remote.users);
  if ((remote.timerAt || 0) > (state.timerAt || 0)) { state.timer = remote.timer; state.timerAt = remote.timerAt; }
  if (remote.settings && (remote.settingsAt || 0) > (state.settingsAt || 0)) { state.settings = remote.settings; state.settingsAt = remote.settingsAt; }
  normalizeSettings();
  return before !== JSON.stringify([state.entries, state.clients, state.bills, state.timer, state.settings, state.societes, state.users]);
}

// Régénère les fichiers Excel seulement si leur contenu a changé.
async function publishExcel() {
  const open = sortEntries(live().filter(e => !e.billId && billable(e))).map(xrow);
  const arch = sortEntries(live().filter(e => e.billId && bill(e.billId))).map(e => {
    const b = bill(e.billId); return { ...xrow(e), billDate: b.date, billRef: b.ref, amount: entryAmount(e) };
  }).sort((a, b) => (a.billDate + a.client + a.date + a.start).localeCompare(b.billDate + b.client + b.date + b.start));
  const jobs = [['open', open, Xlsx.openList, P.open()], ['archive', arch, Xlsx.archive, P.archive()]];
  for (const [k, rows, make, path] of jobs) {
    const sig = hash(JSON.stringify(rows) + path);
    if (state.published[k] === sig) continue;
    const r = await OneDrive.put(path, await make(rows), { type: Xlsx.MIME, overwrite: true });
    if (!r.ok) throw new Error(`OneDrive (${r.status}) : écriture de ${path.split('/').pop()} impossible.`);
    state.published[k] = sig;
  }
}

let syncing = false, syncAgain = false, syncTimer;
function scheduleSync() { clearTimeout(syncTimer); syncTimer = setTimeout(sync, 1500); }

async function sync() {
  if (!OneDrive.configured()) return setSync('off', 'OneDrive non configuré : données sur cet appareil uniquement.');
  if (!OneDrive.signedIn()) return setSync('login', 'Connecte-toi à OneDrive dans Réglages.');
  if (!navigator.onLine) return setSync('offline', 'Hors ligne : synchro au retour du réseau.');
  if (syncing) { syncAgain = true; return; }
  syncing = true; setSync('busy');
  try {
    let changed = false, written = false;
    for (let i = 0; i < 3 && !written; i++) {
      const cur = await OneDrive.get(P.data());
      if (cur && !state.lastSync) { adoptRemote(JSON.parse(cur.data)); changed = true; }
      else if (cur) changed = mergeRemote(JSON.parse(cur.data)) || changed;
      changed = mergeRemoteClients() || changed;
      changed = seedClients() || changed;      // après la fusion : pas de doublon entre appareils
      const payload = JSON.stringify({ version: 3, entries: state.entries, clients: state.clients, bills: state.bills,
        societes: state.societes, users: state.users,
        timer: state.timer, timerAt: state.timerAt, settings: state.settings, settingsAt: state.settingsAt });
      const r = await OneDrive.put(P.data(), payload, { etag: cur?.etag, type: 'application/json' });
      if (r.status === 412) continue;                       // modifié par un autre appareil : on refusionne
      if (!r.ok) throw new Error(`OneDrive (${r.status}) : écriture des données impossible.`);
      written = true;
    }
    // Trois refus de suite (autre appareil en train d'écrire) : pas « synchronisé », nouvel essai à la prochaine synchro.
    if (!written) throw new Error('heures-data.json est modifié en même temps par un autre appareil : réessai à la prochaine synchro.');
    await publishExcel();
    state.lastSync = Date.now();
    await DB.save(state);
    setSync('ok');
    if (changed && !$('#dlg').open) render();
    Achats.sync();                          // gère ses propres erreurs, affichées dans l'onglet Achats
    Taches.sync();                          // idem, onglet Tâches
  } catch (e) {
    if (e instanceof OneDrive.NeedLogin) { setSync('login', 'Session expirée : reconnecte-toi.'); render(); }
    else setSync('err', e.message);
  } finally {
    syncing = false;
    if (syncAgain) { syncAgain = false; scheduleSync(); }
  }
}

// ---------- rendu ----------
const TITLES = { today: "Aujourd'hui", taches: 'Tâches', list: 'Interventions', billing: 'Facturer', archive: 'Archives', achats: 'Achats', articles: 'Articles', admin: 'Administration', settings: 'Réglages' };
// Achats et Articles (achats.js), Tâches (taches.js) : modules séparés, avec leurs propres données et leur propre synchro.
const VIEWS = () => ({ today: viewToday, taches: Taches.view, list: viewList, billing: viewBilling, archive: viewArchive, achats: Achats.view, articles: Achats.articles, admin: viewAdmin, settings: viewSettings });
// Bandeau de la version de développement, avec le dossier de données utilisé (en rouge s'il ne s'agit pas d'un dossier -DEV).
function envBar() {
  if (ENV === 'prod') return;
  let bar = $('#envBar');
  if (!bar) {
    $('.topbar').insertAdjacentHTML('afterend', '<div id="envBar" class="envbar"></div>');
    bar = $('#envBar'); document.body.classList.add('env-dev'); document.title = 'DEV – ' + document.title;
  }
  const folder = state?.onedrive?.folder || DEFAULT_FOLDER, real = !/-DEV\b/i.test(folder);
  bar.classList.toggle('real', real);
  bar.textContent = `DEV · données : ${folder}${real ? ' ⚠ DONNÉES RÉELLES' : ''} · Billit sandbox`;
}

function render() {
  clearInterval(tick);
  envBar();
  // Accès : connexion Microsoft obligatoire, puis code PIN s'il est défini.
  const gate = !OneDrive.signedIn() ? 'login' : locked ? 'pin' : me() === GUEST ? 'denied' : null;
  document.body.classList.toggle('locked', !!gate);
  document.body.classList.toggle('dashwide', !gate && tab === 'today' && wideToday());
  if (gate) { if ($('#dlg').open) $('#dlg').close(); $('#viewTitle').textContent = 'LevelDesk'; return gate === 'login' ? viewLogin() : gate === 'pin' ? viewPin() : viewDenied(); }
  // Onglet Administration : administrateurs seulement.
  $('.tabs [data-tab="admin"]').classList.toggle('hidden', !isAdmin());
  if (tab === 'admin' && !isAdmin()) tab = 'today';
  $('#viewTitle').textContent = TITLES[tab];
  $$('.tabs button[data-tab]').forEach(b => b.classList.toggle('active', b.dataset.tab === tab));
  $('#tabMore').classList.toggle('active', !!$(`.tabs .sec[data-tab="${tab}"]`));   // onglet secondaire ouvert : « Plus » allumé
  VIEWS()[tab]();
}

function entryLi(e, { showDate = false, showBill = true } = {}) {
  const b = e.billId && bill(e.billId);
  return `<li class="tap" data-edit="${e.id}">
    <div class="grow">
      <div><strong>${esc(client(e.clientId).name)}</strong> ${e.onSite ? '<span class="pill">sur site</span>' : '<span class="pill">à distance</span>'}
        ${cmdOf(e) !== HOME ? `<span class="pill warn">${esc(soc(cmdOf(e)).name)}</span>` : ''}
        ${techOf(e) !== me().id ? `<span class="pill">👤 ${esc(user(techOf(e)).name)}</span>` : ''}
        ${e.taskId ? '<span class="pill" title="Faite pour une tâche">📋</span>' : ''}
        ${showBill && b ? `<span class="pill ok">facturé ${frDate(b.date)}${b.ref ? ' · ' + esc(b.ref) : ''}</span>` : ''}</div>
      <div class="small muted">${showDate ? frDate(e.date) + ' · ' : ''}${e.forfait || !e.start ? 'forfait' : `${esc(e.start)}–${esc(e.end)}`}${e.pauseMin ? ` · pause ${fmtDur(e.pauseMin)}` : ''}${e.install ? ' · ' + esc(e.install) : ''}</div>
      ${e.description ? `<div class="small">${esc(e.description)}</div>` : ''}
    </div>
    <div class="dur">${fmtDur(e.minutes)}</div></li>`;
}
function bindEdits(root = document) {
  $$('[data-edit]', root).forEach(li => li.onclick = () => openEntry(state.entries.find(e => e.id === li.dataset.edit)));
}
// Le client choisi reste dans la liste même s'il est archivé ou supprimé : sinon le menu prendrait un autre client à l'enregistrement.
// Seulement les clients visibles par l'utilisateur connecté (Administration).
function clientOptions(sel) {
  const list = activeClients().filter(canSee), cur = sel && !list.some(c => c.id === sel) && state.clients.find(c => c.id === sel);
  return (cur ? `<option value="${cur.id}" selected>${esc(cur.name)}${cur.deleted ? ' (supprimé)' : cur.archived ? ' (archivé)' : ''}</option>` : '')
    + list.map(c => `<option value="${c.id}" ${c.id === sel ? 'selected' : ''}>${esc(c.name)}</option>`).join('');
}
// Sur site, commanditaire et technicien (x : intervention ou chrono ; p : préfixe des id des champs).
// Technicien = l'utilisateur connecté pour une nouvelle intervention, celui d'origine pour une intervention existante : jamais modifiable (demande de Fred).
// Commanditaire = société du client (fiche client), suit le client choisi : pas modifiable à la main non plus (demande de Fred, 30/10).
const flags = (x, p) => `
  <label class="check"><input type="checkbox" id="${p}Site" ${x.onSite ? 'checked' : ''}> Sur site</label>
  <div class="grid2">
    <div><label>Commanditaire</label><select id="${p}Cmd" disabled title="Société du client (fiche client)">${socOptions(cmdOf(x))}</select></div>
    <div><label>Technicien</label><input id="${p}Tech" data-id="${esc(techOf(x))}" value="${esc(userLabel(user(techOf(x))))}" disabled></div>
  </div>`;
const readWho = p => { const cmdId = $(`#${p}Cmd`).value; return { cmdId, techId: $(`#${p}Tech`).dataset.id, alescom: cmdId === 'alescom' }; };
const newWho = c => ({ cmdId: clientSoc(c), techId: me().id, alescom: clientSoc(c) === 'alescom' });

// Heure de début ramenée à aujourd'hui (ou hier si elle est dans le futur).
function startFromTime(hm) {
  const d = new Date(); const [h, m] = hm.split(':').map(Number);
  d.setHours(h, m, 0, 0);
  if (d > new Date()) d.setDate(d.getDate() - 1);
  return d;
}
// Chrono : temps net = maintenant (ou début de la pause en cours) − début − pauses.
const timerNow = t => t.pausedAt ? new Date(t.pausedAt) : new Date();
const timerPausedMs = t => (t.pausedMs || 0);
const timerElapsed = t => Math.max(0, timerNow(t) - new Date(t.start) - timerPausedMs(t));

// ---- Aujourd'hui ----
// Tableau de bord à partir de 1100 px de large (PC) ; en dessous, la page d'origine.
const WIDE = window.matchMedia('(min-width: 1100px)');
const wideToday = () => WIDE.matches;
WIDE.addEventListener('change', () => { if (tab === 'today' && !$('#dlg').open && OneDrive.signedIn() && !locked) render(); });

// Barre latérale du tableau de bord : heures, à facturer, mises à jour serveur du mois, achats.
function dashSide(todays, week, sum) {
  const d = today(), month = d.slice(0, 7), day = +d.slice(8);
  const pm = new Date(d + 'T00:00'); pm.setDate(1); pm.setMonth(pm.getMonth() - 1);
  const prev = isoDate(pm).slice(0, 7), prevName = pm.toLocaleDateString('fr-BE', { month: 'short' });
  const mine = sum(live().filter(e => e.date.startsWith(month)));
  // Comparaison avec la même période du mois précédent (du 1er au même jour).
  const before = sum(live().filter(e => e.date.startsWith(prev) && +e.date.slice(8) <= day));
  const diff = before ? Math.round((mine - before) / before * 100) : null;
  const { groups } = billingGroups(), total = groups.reduce((s, g) => s + g.amount, 0);
  const oldest = groups.filter(g => g.age >= 20).sort((a, b) => b.age - a.age)[0];
  const srv = activeClients().filter(c => c.server);
  const done = srv.filter(c => live().some(e => e.forfait === 'update' && e.clientId === c.id && e.date.startsWith(month)));
  const left = srv.filter(c => !done.includes(c));
  const monthName = new Date(d + 'T00:00').toLocaleDateString('fr-BE', { month: 'long' });
  return `
    <div class="dtile"><div class="l">Heures</div>
      <div class="dgrid"><div><b>${fmtDur(sum(todays))}</b>aujourd'hui</div><div><b>${fmtDur(sum(week))}</b>semaine</div><div><b>${fmtDur(mine)}</b>${esc(monthName)}</div></div>
      ${diff !== null ? `<div class="small ${diff >= 0 ? 'okc' : 'muted'}">${diff >= 0 ? '+' : ''}${diff} % par rapport au début de ${esc(prevName)}</div>` : ''}</div>
    <div class="dtile tap" data-dgo="billing"><div class="l">À facturer ›</div><div class="v">${eur(total)}</div>
      <div class="small muted">${groups.length} client(s)${oldest ? ` · <span class="warnc">${esc(oldest.info.label)} depuis ${oldest.age} j</span>` : ''}</div></div>
    ${srv.length ? `<div class="dtile"><div class="l">Mises à jour serveur · ${esc(monthName)}</div><div class="v">${done.length}/${srv.length}</div>
      <div class="dbar"><i style="width:${done.length / srv.length * 100}%"></i></div>
      <div class="small muted">${left.length ? 'Reste : ' + left.slice(0, 8).map(c => esc(c.name)).join(', ') + (left.length > 8 ? '…' : '') : 'Toutes faites ✓'}</div></div>` : ''}
    <div class="dtile tap" data-dgo="achats" id="achTile"><div class="l">Achats ›</div><div class="small muted">Chargement…</div></div>`;
}

function viewToday() {
  const t = state.timer, d = today();
  const todays = sortEntries(live().filter(e => e.date === d));
  const now = new Date(), monday = new Date(now); monday.setDate(now.getDate() - ((now.getDay() + 6) % 7));
  const week = live().filter(e => e.date >= isoDate(monday) && e.date <= d);
  const sum = es => es.reduce((s, e) => s + e.minutes, 0);
  const lastC = state.clients.find(c => c.id === state.lastClientId && !c.deleted) || activeClients()[0];

  // Mise en page « aérée » (maquette A choisie par Fred le 30/10) : salutation, chrono compact, 3 actions, chiffres du jour, liste.
  const dayLabel = new Date().toLocaleDateString('fr-BE', { weekday: 'long', day: 'numeric', month: 'long' });
  const hello = `${new Date().getHours() >= 18 ? 'Bonsoir' : 'Bonjour'} ${esc(me().name || '')}`.trim();
  const helloSub = todays.length ? `${todays.length} intervention${todays.length > 1 ? 's' : ''} · ${fmtDur(sum(todays))} aujourd'hui` : 'Aucune intervention encore aujourd\'hui';
  // Réglages rarement changés (sur site, commanditaire, heure de début) repliés sous « Options ».
  const timerCard = `
    <div class="card chrono ${t ? 'running' : ''}">
      ${t ? `
        <div class="crow">
          <div class="grow"><div class="cl">Chrono en cours ${t.pausedAt ? '<span class="pill warn">en pause</span>' : ''}</div>
            <div class="cname">${esc(client(t.clientId).name)}</div></div>
          <div class="clock ${t.pausedAt ? 'muted' : ''}" id="clock">0:00:00</div>
        </div>
        ${t.pausedAt ? '<div class="small muted">Pause en cours : <strong id="pauseClock">0:00:00</strong></div>' : ''}
        <div class="row" style="margin-top:10px"><input id="tDesc" class="grow" placeholder="Description (optionnel)" value="${esc(t.description || '')}">${mic('tDesc')}</div>
        <details class="opts"><summary>Début, pause, sur site, commanditaire</summary>
          <div class="grid2">
            <div><label>Début</label><input type="time" id="tStartAt" value="${hhmm(new Date(t.start))}"></div>
            <div><label>Pause déduite (H:MM)</label><input id="tPauseMin" value="${fmtDur(Math.round(timerPausedMs(t) / 60000))}"></div>
          </div>
          ${flags(t, 't')}
        </details>
        <div class="acts"><button class="btn" id="tCancel">Annuler</button>
          <button class="btn" id="tPause">${t.pausedAt ? '▶ Reprendre' : '❚❚ Pause'}</button>
          <button class="btn primary" id="tStop">■ Arrêter</button></div>`
      : lastC ? `
        <div class="crow">
          <div class="grow"><div class="cl">Chrono</div>
            <select id="tClient">${clientOptions(lastC.id)}</select>
            <div class="row"><input id="tDesc0" class="grow" placeholder="Description de l'intervention">${mic('tDesc0')}</div></div>
          <button class="playbtn" id="tStart" title="Démarrer le chrono" aria-label="Démarrer le chrono">▶</button>
        </div>
        <details class="opts"><summary>Sur site, commanditaire, heure de début</summary>
          ${flags({ onSite: false, ...newWho(lastC) }, 't')}
          <div class="row"><span class="small muted">Début</span><input type="time" id="tAt" style="width:auto"><span class="small muted">(vide = maintenant)</span></div>
        </details>`
      : `<p class="muted">Commence par ajouter un client dans Réglages.</p>`}
    </div>`;
  const buttons = `
    <div class="qa">
      <button class="btn" id="addManual"><span class="qi">＋</span>Intervention</button>
      <button class="btn" id="addMaint"><span class="qi">🛠</span>Maintenance</button>
      <button class="btn" id="addQuick"><span class="qi">⚡</span>Rapide 5 min</button>
    </div>`;
  const dayList = todays.length ? `<ul class="list">${todays.map(e => entryLi(e)).join('')}</ul>` : '<div class="empty">Rien pour l\'instant.</div>';
  const toBill = billingGroups().groups.reduce((s, g) => s + g.amount, 0);
  const head = `<div class="hello"><div><h2>${hello}</h2><div class="sub">${esc(helloSub)}</div></div><div class="date">${esc(dayLabel)}</div></div>`;

  // Grand écran (PC) : journée + barre latérale (maquette B du 29/10, au nouveau style) ; téléphone : une colonne.
  $('#view').innerHTML = wideToday() ? `
    ${head}
    <div class="dash">
      <div class="dmain">
        <div id="tkNotif"></div><div id="achNotif"></div>
        ${timerCard}${buttons}
        <div class="card"><h2>Interventions du jour <span class="tot">${fmtDur(sum(todays))}</span></h2>${dayList}</div>
        <div class="card"><h2>Tâches du jour et en retard</h2><div id="tkDay"></div></div>
      </div>
      <aside class="dside">${dashSide(todays, week, sum)}</aside>
    </div>` : `
    ${head}
    <div id="tkNotif"></div>
    <div id="achNotif"></div>
    ${timerCard}${buttons}
    <div class="card stats">
      <div><b>${fmtDur(sum(todays))}</b><span>aujourd'hui</span></div>
      <div><b>${fmtDur(sum(week))}</b><span>semaine</span></div>
      <div class="tap money" id="goBill"><b>${eur(toBill)}</b><span>à facturer ›</span></div>
    </div>
    <div class="card"><h2>Aujourd'hui <span class="tot">${fmtDur(sum(todays))}</span></h2>${dayList}</div>`;

  $('#addManual').onclick = () => openEntry(null);
  $('#addMaint').onclick = () => openForfait(null, 'maint');
  // Intervention rapide : forfait de 5 minutes (sans heures), il reste à choisir le client et à écrire la description.
  // Les mises à jour serveur passent par leurs tâches récurrentes (onglet Tâches), plus par un bouton ici (demande de Fred, 30/10).
  $('#addQuick').onclick = () => openEntry(null, { start: '', end: '', minutes: 5, pauseMin: 0 });
  $$('#goBill, [data-dgo]').forEach(b => b.onclick = () => { tab = b.dataset.dgo || 'billing'; render(); });
  bindEdits();
  Achats.todayBanner($('#achNotif'));         // nouvelles factures fournisseurs (module Achats)
  Taches.todayBanner($('#tkNotif'));          // tâches du jour, en retard, à prendre (module Tâches)
  if ($('#tkDay')) Taches.dayList($('#tkDay'));
  if ($('#achTile')) Achats.dashTile($('#achTile'));

  if (t) {
    const hms = ms => { const s = Math.max(0, Math.floor(ms / 1000)); return `${Math.floor(s / 3600)}:${pad(Math.floor(s / 60) % 60)}:${pad(s % 60)}`; };
    const upd = () => {
      $('#clock').textContent = hms(timerElapsed(t));
      if (t.pausedAt) $('#pauseClock').textContent = hms(Date.now() - new Date(t.pausedAt));
    };
    upd(); tick = setInterval(upd, 1000);
    const save = () => { state.timerAt = Date.now(); DB.save(state); scheduleSync(); };
    $('#tStartAt').onchange = ev => { if (ev.target.value) { t.start = startFromTime(ev.target.value).toISOString(); save(); upd(); } };
    $('#tPauseMin').onchange = ev => {
      const m = parseDur(ev.target.value);
      if (isNaN(m)) return alert('Pause invalide (ex. 0:15).');
      t.pausedMs = m * 60000; ev.target.value = fmtDur(m); save(); upd();
    };
    $('#tDesc').onchange = ev => { t.description = ev.target.value; save(); };
    $('#tSite').onchange = ev => { t.onSite = ev.target.checked; save(); };
    $('#tPause').onclick = async () => {
      t.description = $('#tDesc').value;
      if (t.pausedAt) { t.pausedMs = (t.pausedMs || 0) + (Date.now() - new Date(t.pausedAt)); t.pausedAt = null; }
      else t.pausedAt = new Date().toISOString();
      state.timerAt = Date.now(); await persist();
    };
    $('#tCancel').onclick = async () => { if (confirm('Abandonner ce chrono ?')) { setTimer(null); await persist(); } };
    $('#tStop').onclick = () => {
      t.description = $('#tDesc').value;
      const start = new Date(t.start), end = timerNow(t);
      const pauseMin = Math.round(timerPausedMs(t) / 60000);
      openEntry(null, {
        clientId: t.clientId, date: isoDate(start), start: hhmm(start), end: hhmm(end), pauseMin,
        minutes: Math.max(1, Math.round(timerElapsed(t) / 60000)), description: t.description,
        onSite: t.onSite, cmdId: cmdOf(t), techId: techOf(t), alescom: cmdOf(t) === 'alescom', ...(t.taskId ? { taskId: t.taskId } : {})
      }, true);
    };
  } else if ($('#tStart')) {
    $('#tClient').onchange = ev => { $('#tCmd').value = clientSoc(client(ev.target.value)); };
    $('#tStart').onclick = async () => {
      const at = $('#tAt').value;
      const start = at ? startFromTime(at) : new Date();
      state.lastClientId = $('#tClient').value;
      setTimer({ clientId: state.lastClientId, start: start.toISOString(), description: $('#tDesc0').value.trim(),
        onSite: $('#tSite').checked, ...readWho('t'), pausedMs: 0, pausedAt: null });
      await persist();
    };
  }
}

// ---- dialogue intervention ----
// onSaved(intervention) : appelé après l'enregistrement d'une nouvelle intervention (ex. intervention lue sur une facture Alescom).
// draft : valeurs déjà saisies à réafficher (retour de la fiche « + Nouveau client… ») ; l'intervention elle-même n'est pas modifiée avant « Enregistrer ».
function openEntry(entry, preset = {}, fromTimer = false, onSaved = null, draft = null) {
  if (entry?.forfait) return openForfait(entry);
  // Client proposé : celui du preset (tâche, facture Alescom…), sinon le dernier utilisé ; le commanditaire suit ce client.
  const c0 = (preset.clientId && state.clients.find(c => c.id === preset.clientId && !c.deleted))
    || state.clients.find(c => c.id === state.lastClientId && !c.deleted) || activeClients()[0];
  const e = { ...(entry || {
    clientId: c0?.id, date: today(), start: hhmm(new Date(Date.now() - 3600e3)), end: hhmm(new Date()), minutes: 60, pauseMin: 0,
    description: '', install: '', onSite: false, ...newWho(c0), ...preset
  }), ...(draft || {}) };
  if (!e.clientId) return alert("Ajoute d'abord un client dans Réglages.");
  const b = e.billId && bill(e.billId);
  // Forfait : seule la durée facturée est encodée (pas d'heure de début / fin). Intervention existante sans heures = forfait.
  const flat0 = draft ? draft.flat : entry ? !entry.start && !entry.end : preset.start === '' && preset.end === '' && preset.minutes > 0;
  const dlg = $('#dlg'), f = $('#dlgForm');
  f.innerHTML = `
    <h2>${entry ? 'Modifier' : fromTimer ? 'Fin du chrono' : 'Nouvelle'} intervention</h2>
    ${b ? `<p class="small"><span class="pill ok">facturée le ${frDate(b.date)}${b.ref ? ' · ' + esc(b.ref) : ''}</span>
      ${b.key !== 'import' ? '<button class="link" id="fBill">Voir la facture</button>' : ''}</p>` : ''}
    ${e.taskId ? `<p class="small">📋 Tâche : <button class="link" id="fTask">${esc(Taches.titleOf(e.taskId) || "voir la tâche d'origine")}</button></p>` : ''}
    <label>Client</label>
    <select id="fClient">${clientOptions(e.clientId)}<option value="__new">+ Nouveau client…</option></select>
    <div class="grid2">
      <div><label>Date</label><input type="date" id="fDate" value="${e.date}" required></div>
      <div><label>Durée facturée (H:MM)</label><input id="fDur" value="${fmtDur(e.minutes)}"></div>
    </div>
    ${fromTimer ? '' : `<label class="check"><input type="checkbox" id="fFlat" ${flat0 ? 'checked' : ''}> Forfait (seulement la durée facturée, sans heure de début ni de fin)</label>`}
    <div class="grid3" id="fTimes">
      <div><label>Début</label><input type="time" id="fStart" value="${esc(e.start)}"></div>
      <div><label>Fin</label><input type="time" id="fEnd" value="${esc(e.end)}"></div>
      <div><label>Pause</label><input id="fPause" value="${e.pauseMin ? fmtDur(e.pauseMin) : ''}" placeholder="0:00"></div>
    </div>
    ${lblMic('Description', 'fDesc')}<textarea id="fDesc">${esc(e.description)}</textarea>
    ${flags(e, 'f')}
    <label>Installation / réinstallation (optionnel)</label><input id="fInst" value="${esc(e.install)}" placeholder="ex. Réinstallation PC">
    <div class="actions">
      ${entry && !b ? '<button class="btn danger" id="fDel" style="margin-right:auto">Supprimer</button>' : ''}
      <button class="btn" value="cancel">${fromTimer ? 'Reprendre' : 'Annuler'}</button>
      <button class="btn primary" id="fSave">Enregistrer</button>
    </div>`;

  if ($('#fBill')) $('#fBill').onclick = () => { dlg.close(); showBill(b); };
  if ($('#fTask')) $('#fTask').onclick = () => Taches.open(e.taskId);
  const flat = () => !!$('#fFlat')?.checked;
  const pause = () => flat() ? 0 : parseDur($('#fPause').value) || 0;
  const upDur = () => $('#fDur').value = fmtDur(Math.max(0, diffMin($('#fStart').value, $('#fEnd').value) - pause()));
  $('#fStart').oninput = upDur; $('#fEnd').oninput = upDur; $('#fPause').onchange = upDur;
  $('#fDur').onchange = () => {
    const m = parseDur($('#fDur').value);
    if (!isNaN(m)) { $('#fDur').value = fmtDur(m); if (!flat()) $('#fEnd').value = fromMin(toMin($('#fStart').value) + m + pause()); }
  };
  // Forfait : début, fin et pause grisés (et vidés à l'enregistrement) ; décoché, les heures reviennent.
  const setFlat = () => {
    ['#fStart', '#fEnd', '#fPause'].forEach(s => $(s).disabled = flat());
    $('#fTimes').style.opacity = flat() ? .45 : '';
    if (!flat() && !$('#fStart').value) { $('#fEnd').value = hhmm(new Date()); $('#fStart').value = fromMin(toMin($('#fEnd').value) - (parseDur($('#fDur').value) || 60)); }
  };
  if ($('#fFlat')) { $('#fFlat').onchange = setFlat; setFlat(); }
  $('#fClient').onchange = ev => {
    if (ev.target.value === '__new') {
      // Fiche client complète (avec la liste des clients Billit pas encore reliés), puis retour au formulaire avec ce qui était saisi.
      const d = { clientId: e.clientId, date: $('#fDate').value, minutes: parseDur($('#fDur').value) || e.minutes, start: $('#fStart').value, end: $('#fEnd').value,
        pauseMin: parseDur($('#fPause').value) || 0, description: $('#fDesc').value, onSite: $('#fSite').checked, ...readWho('f'),
        install: $('#fInst').value, flat: flat() };
      const back = x => openEntry(entry, preset, fromTimer, onSaved, x);
      editClient(null, { onSaved: c => back({ ...d, clientId: c.id, ...(entry ? {} : { cmdId: clientSoc(c), alescom: clientSoc(c) === 'alescom' }) }), onCancel: () => back(d) });
    } else $('#fCmd').value = clientSoc(client(ev.target.value));
  };
  if ($('#fDel')) $('#fDel').onclick = async () => {
    if (!confirm('Supprimer cette intervention ?')) return;
    entry.deleted = true; touch(entry);
    dlg.close(); await persist(); toast('Intervention supprimée');
  };
  $('#fSave').onclick = async ev => {
    ev.preventDefault();
    const minutes = parseDur($('#fDur').value), pauseMin = pause();
    const clientId = $('#fClient').value;
    if (!clientId || clientId === '__new') return alert('Choisis un client.');
    if (isNaN(minutes) || minutes <= 0) return alert('Durée invalide (ex. 1:30).');
    if (minutes > 12 * 60 && !confirm(`Durée de ${fmtDur(minutes)} : oubli de dépointage ?\nEnregistrer quand même ?`)) return;
    const data = {
      clientId, date: $('#fDate').value, start: flat() ? '' : $('#fStart').value, end: flat() ? '' : $('#fEnd').value, minutes, pauseMin,
      description: $('#fDesc').value.trim(), onSite: $('#fSite').checked, ...readWho('f'), install: $('#fInst').value.trim()
    };
    if (b && !confirm('Cette intervention est déjà facturée. Modifier quand même l\'archive ?')) return;
    let saved = entry;
    if (entry) Object.assign(touch(entry), data);
    else state.entries.push(saved = touch({ id: uid('e'), billId: null, ...data, ...(e.taskId ? { taskId: e.taskId } : {}) }));
    state.lastClientId = clientId;
    if (fromTimer) setTimer(null);
    if (!entry && e.taskId) await Taches.linkEntry(e.taskId, saved.id);     // intervention faite pour une tâche (module Tâches)
    if (!entry && onSaved) await onSaved(saved);
    dlg.close(); await persist(); toast('Enregistré');
  };
  if (!dlg.open) dlg.showModal();
}

// ---- forfaits serveur : une date et une durée, sans heures de début / fin ----
const FORFAITS = { maint: 'Maintenance serveur', update: 'Mise à jour serveur' };
const serverCount = c => c?.server ? Math.max(1, +c.servers || 1) : 0;
// Mise à jour serveur : 15 min pour 1 serveur, 10 min par serveur au-delà.
const updateMinutes = c => { const n = serverCount(c) || 1; return n === 1 ? 15 : 10 * n; };

// pre (module Tâches) : { clientId, taskId } = client proposé et tâche « Mise à jour serveur » à terminer.
function openForfait(entry, kind, pre = {}) {
  kind = entry?.forfait || kind;
  const label = FORFAITS[kind];
  // Seulement les clients avec serveur(s) dont ce forfait n'a pas encore été encodé ce mois-ci
  // (en modification, le client de l'intervention reste proposé).
  const month = today().slice(0, 7);
  const done = new Set(live().filter(e => e.forfait === kind && e.date.startsWith(month) && e.id !== entry?.id).map(e => e.clientId));
  const clients = activeClients().filter(c => (c.server && !done.has(c.id)) || c.id === entry?.clientId);
  if (!clients.length) {
    const n = activeClients().filter(c => c.server).length;
    return alert(n ? `${label} : déjà encodée ce mois-ci pour les ${n} client(s) avec serveur.`
      : 'Aucun client avec serveur : coche « Serveur(s) » dans la fiche client (Réglages).');
  }
  const c0 = entry ? client(entry.clientId) : clients.find(c => c.id === pre.clientId) || clients.find(c => c.id === state.lastClientId) || clients[0];
  const e = entry || { clientId: c0.id, date: today(), minutes: kind === 'update' ? updateMinutes(c0) : 0,
    description: label, onSite: false, ...newWho(c0) };
  const b = e.billId && bill(e.billId);
  const opts = clients.map(c => `<option value="${c.id}" ${c.id === e.clientId ? 'selected' : ''}>${esc(c.name)}${c.server ? ` · ${serverCount(c)} serveur(s)` : ''}</option>`).join('');
  const hint = c => kind !== 'update' ? 'Forfait : seule la durée est facturée.'
    : c.server ? `${serverCount(c)} serveur(s) → ${fmtDur(updateMinutes(c))}` : 'Client sans serveur renseigné (fiche client) : 15 min par défaut.';

  const dlg = $('#dlg'), f = $('#dlgForm');
  f.innerHTML = `<h2>${esc(label)}</h2>
    ${b ? `<p class="small"><span class="pill ok">facturée le ${frDate(b.date)}${b.ref ? ' · ' + esc(b.ref) : ''}</span></p>` : ''}
    <label>Client</label><select id="kClient">${opts}</select>
    <div class="grid2">
      <div><label>Date</label><input type="date" id="kDate" value="${e.date}"></div>
      <div><label>Durée (H:MM)</label><input id="kDur" value="${e.minutes ? fmtDur(e.minutes) : ''}" placeholder="ex. 1:30"></div>
    </div>
    <p class="small muted" id="kHint">${esc(hint(client(e.clientId)))}</p>
    ${lblMic('Description', 'kDesc')}<textarea id="kDesc">${esc(e.description)}</textarea>
    ${flags(e, 'k')}
    <div class="actions">
      ${entry && !b ? '<button class="btn danger" id="kDel" style="margin-right:auto">Supprimer</button>' : ''}
      <button class="btn" value="cancel">Annuler</button>
      <button class="btn primary" id="kSave">Enregistrer</button>
    </div>`;
  $('#kClient').onchange = ev => {
    const c = client(ev.target.value);
    $('#kHint').textContent = hint(c);
    if (kind === 'update') $('#kDur').value = fmtDur(updateMinutes(c));
    $('#kCmd').value = clientSoc(c);
  };
  if ($('#kDel')) $('#kDel').onclick = async () => {
    if (!confirm('Supprimer cette intervention ?')) return;
    entry.deleted = true; touch(entry); dlg.close(); await persist(); toast('Intervention supprimée');
  };
  $('#kSave').onclick = async ev => {
    ev.preventDefault();
    const minutes = parseDur($('#kDur').value);
    if (isNaN(minutes) || minutes <= 0) return alert('Encode une durée (ex. 0:45 ou 1:30).');
    if (b && !confirm('Cette intervention est déjà facturée. Modifier quand même l\'archive ?')) return;
    const data = { clientId: $('#kClient').value, date: $('#kDate').value, start: '', end: '', minutes, pauseMin: 0,
      description: $('#kDesc').value.trim(), onSite: $('#kSite').checked, ...readWho('k'), install: '', forfait: kind };
    let saved = entry;
    if (entry) Object.assign(touch(entry), data);
    else state.entries.push(saved = touch({ id: uid('e'), billId: null, ...data, ...(pre.taskId && data.clientId === pre.clientId ? { taskId: pre.taskId } : {}) }));
    state.lastClientId = data.clientId;
    // Mise à jour serveur : la tâche du mois de ce client (module Tâches) est reliée et terminée.
    if (!entry && kind === 'update') await Taches.forfaitSaved(saved);
    dlg.close(); await persist(); toast('Enregistré');
  };
  if (!dlg.open) dlg.showModal();
  if (kind === 'maint' && !entry) $('#kDur').focus();
}

// ---- Interventions ----
let listMonth = today().slice(0, 7), listClient = '', listStatus = '';
function viewList() {
  const es = sortEntries(live().filter(e => (!listMonth || e.date.startsWith(listMonth)) && (!listClient || e.clientId === listClient)
    && (!listStatus || (listStatus === 'open') === !e.billId))).reverse();
  const days = [...new Set(es.map(e => e.date))];
  const months = [...new Set([today().slice(0, 7), ...live().map(e => e.date.slice(0, 7))])].sort().reverse();
  const monthLabel = m => new Date(m + '-01T00:00').toLocaleDateString('fr-BE', { month: 'long', year: 'numeric' });
  $('#view').innerHTML = `
    <div class="card"><div class="grid3">
      <div><label>Mois</label><select id="hMonth"><option value="">Tous</option>
        ${months.map(m => `<option value="${m}" ${m === listMonth ? 'selected' : ''}>${monthLabel(m)}</option>`).join('')}</select></div>
      <div><label>Client</label><select id="hClient"><option value="">Tous</option>${clientOptions(listClient)}</select></div>
      <div><label>Statut</label><select id="hStatus"><option value="">Toutes</option>
        <option value="open" ${listStatus === 'open' ? 'selected' : ''}>À facturer</option>
        <option value="billed" ${listStatus === 'billed' ? 'selected' : ''}>Facturées</option></select></div>
    </div>
    <p class="small muted" style="margin:10px 0 0">${es.length} intervention(s) · <strong>${fmtDur(es.reduce((s, e) => s + e.minutes, 0))}</strong></p></div>
    ${days.length ? days.map(d => `<div class="day">${longDate(d)}</div>
      <div class="card" style="padding:4px 14px"><ul class="list">${es.filter(e => e.date === d).map(e => entryLi(e)).join('')}</ul></div>`).join('')
      : '<div class="empty">Aucune intervention.</div>'}`;
  $('#hMonth').onchange = e => { listMonth = e.target.value; render(); };
  $('#hClient').onchange = e => { listClient = e.target.value; render(); };
  $('#hStatus').onchange = e => { listStatus = e.target.value; render(); };
  bindEdits();
}

// ---- Facturer ----
let billingTarget = null, billUntil = today(), excluded = new Set(), excludedAch = new Set();
// À facturer, par client / commanditaire (onglet Facturer et tableau de bord).
function billingGroups() {
  // Seulement ce que One Level UP facture (pas le travail d'un autre technicien pour un autre commanditaire).
  const open = live().filter(e => !e.billId && billable(e));
  // Un client sans intervention en attente apparaît aussi s'il a des achats prêts à refacturer.
  const keys = [...new Set([...open.map(groupKey), ...Achats.readyClients().map(id => 'c:' + id)])];
  const groups = keys.map(k => {
    const info = groupInfo(k), es = sortEntries(open.filter(e => groupKey(e) === k));
    const ach = k.startsWith('c:') ? Achats.readyFor(k.slice(2)) : [], t = totals(es, info.rates);
    // age : jours d'attente de la plus ancienne intervention (affiché « depuis n j » à partir de 20 j).
    const age = es.length ? Math.round((new Date(today() + 'T00:00') - new Date(es[0].date + 'T00:00')) / 864e5) : 0;
    return { info, es, ach, t, achAmount: achSum(ach), amount: t.amount + achSum(ach), age };
  }).sort((a, b) => b.amount - a.amount);
  return { open, keys, groups };
}

function viewBilling() {
  const { open, keys, groups } = billingGroups();
  if (billingTarget && !keys.includes(billingTarget)) billingTarget = null;
  if (billingTarget) return viewBillingTarget(billingTarget, open.filter(e => groupKey(e) === billingTarget));

  const sum = f => groups.reduce((s, g) => s + f(g), 0);
  const short = d => frDate(d).slice(0, 5);
  // Couleur de la pastille : fixe pour un client (tirée de sa clé), pas de sa place dans la liste.
  const AV = ['#1f5f8b', '#2e7d4f', '#8a4fbf', '#b26a00', '#b3261e', '#00796b', '#5d6d7e'];
  const avColor = key => AV[parseInt(hash(key), 36) % AV.length];

  // Mise en page « cartes clients » (choisie par Fred le 26/10) : bandeau du total, puis une carte par client.
  $('#view').innerHTML = `
    <div class="fhead"><div class="sub">À facturer · HTVA estimé</div><div class="big">${eur(sum(g => g.amount))}</div>
      <div class="split"><div><b>${eur(sum(g => g.t.amount))}</b>${open.length} intervention(s)</div>
        <div><b>${eur(sum(g => g.achAmount))}</b>${sum(g => g.ach.length)} achat(s)</div>
        <div><b>${groups.length}</b>client(s)</div></div></div>
    ${groups.length ? `<div class="flegend"><i></i>interventions <u></u>achats</div>
      ${groups.map(g => `
      <div class="ccard tap" data-target="${g.info.key}">
        <div class="av" style="background:${avColor(g.info.key)}">${esc((g.info.label.trim()[0] || '?').toUpperCase())}</div>
        <div class="mid"><div class="nm">${esc(g.info.label)} ${g.info.partner ? '<span class="pill warn">commanditaire</span>' : ''}</div>
          <div class="meta">${[g.es.length && `${g.es.length} interv. · ${fmtDur(g.t.minutes)} · ${short(g.es[0].date)} → ${short(g.es[g.es.length - 1].date)}`,
            g.ach.length && `${g.ach.length} achat${g.ach.length > 1 ? 's' : ''}`].filter(Boolean).join(' · ')}</div>
          <div class="bar"><i style="width:${g.amount ? g.t.amount / g.amount * 100 : 100}%"></i><u style="width:${g.amount ? g.achAmount / g.amount * 100 : 0}%"></u></div></div>
        <div class="amt">${eur(g.amount)}${g.age >= 20 ? `<small>depuis ${g.age} j</small>` : ''}</div><span class="chev">›</span>
      </div>`).join('')}`
      : '<div class="card"><div class="empty">Tout est facturé 🎉</div></div>'}`;
  $$('[data-target]').forEach(li => li.onclick = () => { billingTarget = li.dataset.target; billUntil = today(); excluded = new Set(); excludedAch = new Set(); render(); });
}

// ach : achats refacturés sur la même facture (facultatif).
function totalsTable(t, r, ach = []) {
  const interv = t.h || t.m || t.dep || !ach.length;
  return `<table class="tot">
    ${interv ? `<tr><td>${esc(r.hLabel)}<br><span class="small muted">${t.h} × ${eur(r.h)}</span></td><td>${eur(t.h * r.h)}</td></tr>
    <tr><td>${esc(r.mLabel)}<br><span class="small muted">${t.m} × ${eur(r.m)}</span></td><td>${eur(t.m * r.m)}</td></tr>
    <tr><td>${esc(r.depLabel)}<br><span class="small muted">${t.dep} × ${eur(r.dep)}</span></td><td>${eur(t.dep * r.dep)}</td></tr>` : ''}
    ${ach.map(x => `<tr><td>${esc(achLabel(x))}<br><span class="small muted">${x.qty} × ${eur(x.unit)}${x.from ? ` · du ${frDate(x.from)} au ${frDate(x.to)}` : ''}</span></td><td>${eur(x.total)}</td></tr>`).join('')}
    <tr class="sum"><td>Total HTVA</td><td>${eur(t.amount + achSum(ach))}</td></tr>
    <tr><td class="muted">TVAC (21 %)</td><td class="muted">${eur((t.amount + achSum(ach)) * 1.21)}</td></tr></table>`;
}

function viewBillingTarget(key, all) {
  const info = groupInfo(key), r = info.rates;
  const es = sortEntries(all.filter(e => e.date <= billUntil));
  const sel = es.filter(e => !excluded.has(e.id));
  const t = totals(sel, r);
  const ach = key.startsWith('c:') ? Achats.readyFor(key.slice(2)) : [];
  const selA = ach.filter(x => !excludedAch.has(x.id));
  const what = [sel.length && `${sel.length} intervention(s)`, selA.length && `${selA.length} achat(s)`].filter(Boolean).join(' + ');

  $('#view').innerHTML = `
    <button class="link" id="back">← Retour</button>
    <div class="card" style="margin-top:8px"><h2>${esc(info.label)} <span class="small muted">→ Billit : ${esc(info.billitName)}</span></h2>
      ${es.length ? `<label>Inclure jusqu'au</label><input type="date" id="until" value="${billUntil}">` : ''}
      <div style="margin-top:12px">${totalsTable(t, r, selA)}</div>
      ${sel.some(e => e.install) ? '<p class="small"><span class="pill warn">Contient des installations : article à ajouter dans Billit.</span></p>' : ''}
      <button class="btn primary big" id="doBill" style="margin-top:12px" ${what ? '' : 'disabled'}>Facturer ${what || '…'}</button>
      <p class="small muted">${sel.length ? `Crée « Prestations - ${esc(info.label)} ${today()}.xlsx » dans OneDrive/A facturer et déplace ces interventions vers l'archive.` : ''}
        ${selA.length ? 'Les achats passent dans l\'archive des achats.' : ''}</p>
    </div>
    ${es.length ? `<div class="card"><h2>Interventions (${sel.length}/${es.length})</h2>
      <ul class="list">${es.map(e => `<li><input type="checkbox" data-inc="${e.id}" ${excluded.has(e.id) ? '' : 'checked'}>
        <div class="grow tap" data-edit="${e.id}"><div><strong>${frDate(e.date)}</strong> ${e.forfait || !e.start ? 'forfait' : `${esc(e.start)}–${esc(e.end)}`}
          ${info.partner ? '· ' + esc(client(e.clientId).name) : ''} ${e.onSite ? '<span class="pill">sur site</span>' : ''}
          ${techOf(e) !== me().id ? `<span class="pill">👤 ${esc(user(techOf(e)).name)}</span>` : ''}</div>
          <div class="small muted">${esc(e.description)}${e.install ? ' · ' + esc(e.install) : ''}</div></div>
        <div class="dur">${fmtDur(e.minutes)}</div></li>`).join('')}</ul></div>` : ''}
    ${ach.length ? `<div class="card"><h2>Achats à refacturer (${selA.length}/${ach.length})</h2>
      <ul class="list">${ach.map(x => `<li><input type="checkbox" data-inca="${x.id}" ${excludedAch.has(x.id) ? '' : 'checked'}>
        <div class="grow"><div><strong>${esc(x.ref)}</strong> ${esc(x.label)}</div>
          <div class="small muted">${[x.supplier, x.domain, x.from && `du ${frDate(x.from)} au ${frDate(x.to)}`, `${x.qty} × ${eur(x.unit)}`].filter(Boolean).map(esc).join(' · ')}</div></div>
        <div class="dur">${eur(x.total)}</div></li>`).join('')}</ul></div>` : ''}`;

  $('#back').onclick = () => { billingTarget = null; render(); };
  if ($('#until')) $('#until').onchange = ev => { billUntil = ev.target.value || today(); render(); };
  $$('[data-inc]').forEach(c => c.onchange = () => { c.checked ? excluded.delete(c.dataset.inc) : excluded.add(c.dataset.inc); render(); });
  $$('[data-inca]').forEach(c => c.onchange = () => { c.checked ? excludedAch.delete(c.dataset.inca) : excludedAch.add(c.dataset.inca); render(); });
  bindEdits();
  $('#doBill').onclick = async () => {
    // Pas de question sur le n° de facture : il se saisit dans la fenêtre de la facture (champ « N° de facture Billit »).
    const total = t.amount + achSum(selA);
    const b = touch({ id: uid('b'), key, label: info.label, date: today(), ref: '', entryIds: sel.map(e => e.id),
      from: sel[0]?.date || null, to: sel[sel.length - 1]?.date || null, amount: total });
    state.bills.push(b);
    sel.forEach(e => { e.billId = b.id; touch(e); });
    await Achats.attach(selA.map(x => x.id), b);
    billingTarget = null;
    await persist();
    showBill(b, true);
  };
}

// Dialogue d'une facture : fichier Prestations, récap Billit, référence, annulation.
async function showBill(b, justCreated = false) {
  const info = groupInfo(b.key);
  const es = sortEntries(live().filter(e => e.billId === b.id));
  const t = totals(es, info.rates), ach = Achats.ofBill(b.id);
  const dlg = $('#dlg'), f = $('#dlgForm');
  f.innerHTML = `<h2>${justCreated ? 'Facturé : ' : ''}${esc(b.label)}</h2>
    <p class="small muted">Facturé le ${frDate(b.date)} · ${[es.length && `${es.length} interv. · du ${frDate(b.from)} au ${frDate(b.to)}`, ach.length && `${ach.length} achat(s)`].filter(Boolean).join(' · ')}</p>
    <p class="small" id="bFile">${justCreated && es.length ? 'Envoi du fichier Prestations vers OneDrive…' : ''}</p>
    <div class="card" style="margin:10px 0;padding:10px">
      <div class="row"><strong class="grow">Billit</strong>
        <button class="btn ${b.billit ? '' : 'primary'}" id="bSend">📤 ${b.billit ? 'Renvoyer' : 'Envoyer'} vers Billit (${Billit.mode() === 'live' ? 'réel' : 'test'})</button></div>
      <div class="small" id="bBillitMsg" style="margin-top:6px">${billitStatus(b)}</div>
    </div>
    <label>N° de facture Billit</label><input id="bRef" value="${esc(b.ref)}">
    ${totalsTable(t, info.rates, ach)}
    <pre class="recap">${esc(recapBillit(b))}</pre>
    <ul class="list small" style="margin-top:8px">${es.map(e => entryLi(e, { showDate: true, showBill: false })).join('')}</ul>
    <div class="actions">
      <button class="btn danger" id="bUndo" style="margin-right:auto">Annuler la facturation</button>
      ${es.length ? '<button class="btn" id="bDl">⬇ Prestations</button>' : ''}
      <button class="btn" id="bCopy">Copier récap</button>
      <button class="btn primary" id="bOk">Fermer</button>
    </div>`;
  const saveRef = async () => { const v = $('#bRef').value.trim(); if (v !== (b.ref || '')) { b.ref = v; touch(b); await persist(); } };
  $('#bOk').onclick = async ev => { ev.preventDefault(); await saveRef(); dlg.close(); };
  $('#bCopy').onclick = () => copy(recapBillit(b));
  $('#bSend').onclick = async () => { const btn = $('#bSend'); btn.disabled = true; try { await sendToBillit(b); } finally { if (btn.isConnected) btn.disabled = false; } };
  if ($('#bDl')) $('#bDl').onclick = async () => { try { download(await prestationsFile(b), prestationsName(b)); } catch (e) { alert(e.message); } };
  $('#bUndo').onclick = async () => {
    if (!confirm(`Remettre ${[es.length && 'ces interventions', ach.length && 'ces achats'].filter(Boolean).join(' et ')} dans « à facturer » ? (le fichier Prestations déjà créé n'est pas supprimé)`
      + (b.billit ? `\n\n⚠ La facture créée dans Billit (${billitModeLabel(b.billit.mode)}${b.billit.number ? ', n° ' + b.billit.number : ''}) n'est PAS supprimée : supprime-la toi-même dans Billit.` : ''))) return;
    es.forEach(e => { e.billId = null; touch(e); });
    await Achats.detach(b.id);
    b.deleted = true; touch(b);
    dlg.close(); await persist(); toast('Facturation annulée');
  };
  $$('[data-edit]', f).forEach(li => li.onclick = null);
  if (!dlg.open) dlg.showModal();

  if (justCreated && es.length) {
    const el = $('#bFile');
    try {
      if (!OneDrive.signedIn()) throw new Error('OneDrive non connecté');
      const r = await OneDrive.put(P.prest(b), await prestationsFile(b), { type: Xlsx.MIME, overwrite: true });
      if (!r.ok) throw new Error(`OneDrive (${r.status})`);
      if (el) el.innerHTML = `<span class="pill ok">✓ ${esc(prestationsName(b))} créé dans A facturer</span>`;
    } catch (e) {
      if (el) el.innerHTML = `<span class="pill warn">Fichier non envoyé (${esc(e.message)}) : utilise « ⬇ Prestations ».</span>`;
    }
  }
}

// ---- Envoi vers Billit (API) ----
// Crée la facture dans Billit avec les articles (mai-xxx, dep-xxx) et le fichier Prestations en pièce jointe.
// L'API ne permet pas d'ajouter des lignes à une facture existante : si une facture non envoyée existe déjà
// pour le client, on prévient et on laisse le choix.
const billitModeLabel = m => m === 'live' ? 'réel' : 'test (sandbox)';
function billitStatus(b) {
  const x = b.billit;
  if (!x) return Billit.ready() ? '<span class="muted">Pas encore envoyée.</span>'
    : '<span class="muted">Renseigne la clé API et le PartyID Billit dans Réglages.</span>';
  return `<span class="pill ok">✓ Envoyée (${billitModeLabel(x.mode)}) le ${new Date(x.at).toLocaleString('fr-BE', { dateStyle: 'short', timeStyle: 'short' })}</span>
    ${x.number ? ` · n° ${esc(x.number)}` : ''}${x.status ? ` · statut ${esc(x.status)}` : ''}
    ${/^https:\/\//.test(x.url || '') ? `· <a href="${esc(x.url)}" target="_blank" rel="noopener">Ouvrir dans Billit</a>` : ''}`;
}

const toBase64 = blob => new Promise((resolve, reject) => {
  const r = new FileReader();
  r.onload = () => resolve(String(r.result).split(',')[1]);
  r.onerror = () => reject(r.error);
  r.readAsDataURL(blob);
});

// ID du client Billit (mémorisé par mode : les ID du sandbox et du réel sont différents).
async function resolveBillitCustomer(info, say) {
  const m = Billit.mode();
  const c = info.partner ? null : client(info.key.slice(2));
  const stored = info.partner ? info.soc.billitIds?.[m] : c.billitIds?.[m];
  if (stored) return stored;
  say('Recherche du client dans Billit…');
  const all = await Billit.customers();
  const want = fold(info.billitName);
  const exact = all.filter(p => [p.Name, p.CommercialName, p.DisplayName].some(n => n && fold(n) === want));
  let id = exact.length === 1 ? exact[0].PartyID : null;
  if (!id) {
    // Choix manuel : les correspondances partielles en premier.
    const score = p => fold(Billit.partyName(p)).includes(want) || want.includes(fold(Billit.partyName(p))) ? 0 : 1;
    const list = all.slice().sort((a, b) => score(a) - score(b) || Billit.partyName(a).localeCompare(Billit.partyName(b)));
    if (!list.length) throw new Error('Aucun client trouvé dans Billit.');
    say(`<label style="margin-top:0">Client Billit pour « ${esc(info.label)} »</label>
      <select id="bCust">${list.map(p => `<option value="${p.PartyID}">${esc(Billit.partyName(p))}${p.VATNumber ? ' · ' + esc(p.VATNumber) : ''}</option>`).join('')}</select>
      <div class="row" style="margin-top:8px"><button class="btn" id="bCustNo">Annuler</button><button class="btn primary" id="bCustOk">Utiliser ce client</button></div>`);
    $$('#bCustNo, #bCustOk').forEach(x => x.type = 'button');
    id = await new Promise(res => { $('#bCustOk').onclick = () => res($('#bCust').value); $('#bCustNo').onclick = () => res(null); });
    if (!id) return null;
  }
  if (info.partner) { info.soc.billitIds = { ...info.soc.billitIds, [m]: id }; touch(info.soc); }
  else { c.billitIds = { ...c.billitIds, [m]: id }; touch(c); }
  await DB.save(state); scheduleSync();
  return id;
}

function billitLines(b, info, products) {
  const r = info.rates, t = totals(live().filter(e => e.billId === b.id), r);
  const byRef = ref => products.find(p => fold(p.Reference) === fold(ref));
  const missing = [];
  const line = (qty, ref, label, price) => {
    const p = byRef(ref);
    if (!p) missing.push(ref);
    return { Quantity: qty, Reference: ref, Description: p?.Description || label,
      UnitPriceExcl: p?.AmountExcl ?? p?.UnitPriceExcl ?? price, VATPercentage: p?.VAT ?? p?.VATPercentage ?? 21,
      ...(p?.ProductID ? { ProductID: p.ProductID } : {}) };
  };
  const lines = [];
  if (t.h) lines.push(line(t.h, r.hRef, r.hLabel, r.h));
  if (t.m) lines.push(line(t.m, r.mRef, r.mLabel, r.m));
  if (t.dep) lines.push(line(t.dep, r.depRef, r.depLabel, r.dep));
  // Achats refacturés : prix calculé par le module Achats (achat + marge).
  const ach = Achats.orderLines(b.id, products);
  lines.push(...ach.lines); missing.push(...ach.missing);
  return { lines, missing };
}

async function sendToBillit(b) {
  const say = h => { const el = $('#bBillitMsg'); if (el) el.innerHTML = h; };
  const info = groupInfo(b.key), m = Billit.mode();
  try {
    if (!Billit.ready()) throw new Error('Renseigne la clé API et le PartyID Billit dans Réglages.');
    if (b.billit && !confirm(`Cette facture a déjà été envoyée vers Billit (${billitModeLabel(b.billit.mode)}).\nEn créer une nouvelle quand même ?`)) return;
    const customerId = await resolveBillitCustomer(info, say);
    if (!customerId) return say('Envoi annulé.');

    say('Vérification des factures non envoyées pour ce client…');
    const open = await Billit.openInvoices(customerId);
    if (open.length) {
      const desc = open.map(o => `- ${o.OrderNumber || '(sans n°)'} du ${frDate(String(o.OrderDate || '').slice(0, 10))}${o.TotalExcl != null ? ` : ${eur(o.TotalExcl)} HTVA` : ''}`).join('\n');
      if (!confirm(`${open.length} facture(s) non envoyée(s) existent déjà pour ${info.billitName} :\n${desc}\n\nL'API Billit ne permet pas d'y ajouter des lignes.\nCréer une NOUVELLE facture quand même ?\n(Annuler = compléter toi-même ou avec la tâche Claude)`)) {
        return say('Envoi annulé. Factures existantes : ' + open.map(o => `<a href="${esc(Billit.orderUrl(o.OrderID))}" target="_blank" rel="noopener">${esc(o.OrderNumber || 'facture ' + o.OrderID)}</a>`).join(', '));
      }
    }

    say('Lecture des articles Billit…');
    const { lines, missing } = billitLines(b, info, await Billit.products());
    if (!lines.length) throw new Error('Aucune ligne à facturer.');
    const total = lines.reduce((s, l) => s + l.Quantity * l.UnitPriceExcl, 0);
    const inst = live().filter(e => e.billId === b.id && e.install);
    const days = dueDays(info), due = new Date(); due.setDate(due.getDate() + days);
    const order = {
      // Draft = « Projet » dans Billit (sans ce statut, la facture est numérotée tout de suite).
      OrderType: 'Invoice', OrderDirection: 'Income', OrderStatus: 'Draft', Currency: 'EUR',
      OrderDate: today(), ExpiryDate: isoDate(due), CustomerID: +customerId || customerId,
      OrderTitle: billTitle(b),
      OrderLines: lines,
      // Fichier Prestations seulement s'il y a des interventions (une facture peut ne contenir que des achats).
      ...(b.from ? { Attachments: [{ FileName: prestationsName(b), FileContent: await toBase64(await prestationsFile(b)), MimeType: Xlsx.MIME }] } : {})
    };
    const summary = lines.map(l => `- ${l.Reference} ${l.Description} : ${l.Quantity} × ${eur(l.UnitPriceExcl)}`).join('\n');
    if (!confirm(`${m === 'live' ? '⚠ MODE RÉEL – vraie comptabilité' : 'Mode TEST (sandbox)'}\n\nFacture pour ${info.billitName}\n${order.OrderTitle}\n${summary}\nTotal HTVA : ${eur(total)}`
      + `\nÉchéance : ${frDate(order.ExpiryDate)} (${days} jours)`
      + (missing.length ? `\n\nArticles introuvables dans Billit (prix de l'app utilisé) : ${missing.join(', ')}` : '')
      + (inst.length ? `\n\nInstallations NON ajoutées (à ajouter dans Billit) : ${inst.map(e => e.install).join(', ')}` : '')
      + '\n\nCréer la facture ?')) return say('Envoi annulé.');

    say('Création de la facture…');
    const res = await Billit.createOrder(order);
    const id = typeof res === 'object' && res ? (res.OrderID ?? res.id ?? res.Id) : res;
    let o = null;
    try { o = await Billit.getOrder(id); } catch (_) { /* la facture est créée, le détail est facultatif */ }
    b.billit = { id, mode: m, at: Date.now(), url: Billit.orderUrl(id), number: o?.OrderNumber || '', status: o?.OrderStatus || '' };
    touch(b); await persist();
    const draft = /draft/i.test(b.billit.status);
    say(billitStatus(b) + (draft
      ? '<p class="small"><span class="pill ok">Créée en Projet</span> Vérifie-la dans Billit avant de la créer définitivement.</p>'
      : `<p class="small"><span class="pill warn">Attention : Billit ne l'a pas créée en Projet (statut ${esc(b.billit.status || 'inconnu')}${b.billit.number ? ', n° ' + esc(b.billit.number) : ''}).</span></p>`));
    toast('Facture créée dans Billit');
  } catch (e) {
    say(`<span class="pill warn">${esc(e.message)}</span>`);
  }
}

// ---- Archives : anciennes interventions (déjà facturées) ----
let archQuery = '', archYear = today().slice(0, 4), archClient = '', archCmd = '';
const ARCH_MAX = 300;
function viewArchive() {
  const q = archQuery.toLowerCase();
  const billed = live().filter(e => e.billId && bill(e.billId));
  const years = [...new Set([today().slice(0, 4), ...billed.map(e => e.date.slice(0, 4))])].sort().reverse();
  const text = e => { const b = bill(e.billId); return [client(e.clientId).name, e.description, e.install, b.ref, b.label].join(' ').toLowerCase(); };
  const es = sortEntries(billed.filter(e => (!archYear || e.date.startsWith(archYear)) && (!archClient || e.clientId === archClient)
    && (!archCmd || cmdOf(e) === archCmd) && (!q || text(e).includes(q)))).reverse();
  const cmds = [...new Set(billed.map(cmdOf))].map(soc).sort((a, b) => (b.id === HOME) - (a.id === HOME) || a.name.localeCompare(b.name));
  const shown = es.slice(0, ARCH_MAX);
  const days = [...new Set(shown.map(e => e.date))];

  $('#view').innerHTML = `
    <div class="card">
      <label style="margin-top:0">Rechercher</label><input type="search" id="aQ" value="${esc(archQuery)}" placeholder="description, client, n° facture…">
      <div class="grid3">
        <div><label>Année</label><select id="aYear"><option value="">Toutes</option>${years.map(y => `<option ${y === archYear ? 'selected' : ''}>${y}</option>`).join('')}</select></div>
        <div><label>Client</label><select id="aClient"><option value="">Tous</option>${clientOptions(archClient)}</select></div>
        <div><label>Commanditaire</label><select id="aCmd"><option value="">Tous</option>${cmds.map(s => `<option value="${s.id}" ${s.id === archCmd ? 'selected' : ''}>${esc(s.name)}</option>`).join('')}</select></div>
      </div>
      <p class="small muted" style="margin:10px 0 0">${es.length} intervention(s) archivée(s) · <strong>${fmtDur(es.reduce((s, e) => s + e.minutes, 0))}</strong>
        ${es.length > ARCH_MAX ? ` · ${ARCH_MAX} plus récentes affichées, affine la recherche` : ''}</p>
    </div>
    ${days.length ? days.map(d => `<div class="day">${longDate(d)}</div>
      <div class="card" style="padding:4px 14px"><ul class="list">${shown.filter(e => e.date === d).map(e => entryLi(e)).join('')}</ul></div>`).join('')
      : '<div class="empty">Aucune intervention archivée.</div>'}`;
  const qi = $('#aQ');
  qi.oninput = () => {
    archQuery = qi.value; clearTimeout(viewArchive.t);
    viewArchive.t = setTimeout(() => { render(); const n = $('#aQ'); n.focus(); n.setSelectionRange(n.value.length, n.value.length); }, 250);
  };
  $('#aYear').onchange = e => { archYear = e.target.value; render(); };
  $('#aClient').onchange = e => { archClient = e.target.value; render(); };
  $('#aCmd').onchange = e => { archCmd = e.target.value; render(); };
  // Archives : consultation seulement (pas de modification d'une intervention facturée).
  $$('[data-edit]').forEach(li => li.onclick = () => viewArchivedEntry(state.entries.find(e => e.id === li.dataset.edit)));
}

// Intervention archivée en lecture seule, pour tous les types (chrono, saisie, forfait serveur), avec accès à sa facture.
function viewArchivedEntry(e) {
  const b = e.billId && bill(e.billId);
  const dlg = $('#dlg'), f = $('#dlgForm');
  const row = (k, v) => v ? `<tr><td class="muted">${k}</td><td>${v}</td></tr>` : '';
  const flat = e.forfait || !e.start;
  f.innerHTML = `<h2>Intervention archivée</h2>
    ${b ? `<p class="small"><span class="pill ok">facturée le ${frDate(b.date)}${b.ref ? ' · ' + esc(b.ref) : ''}</span></p>` : ''}
    <table class="tot ro">
      ${row('Client', esc(client(e.clientId).name))}
      ${row('Date', esc(longDate(e.date)))}
      ${row('Horaire', flat ? `forfait${e.forfait ? ' · ' + esc(FORFAITS[e.forfait] || e.forfait) : ''}` : `${esc(e.start)} – ${esc(e.end)}`)}
      ${row('Pause', !flat && e.pauseMin ? fmtDur(e.pauseMin) : '')}
      ${row('Durée facturée', `<strong>${fmtDur(e.minutes)}</strong>`)}
      ${row('Lieu', e.onSite ? 'sur site' : 'à distance')}
      ${row('Commanditaire', esc(soc(cmdOf(e)).name))}
      ${row('Technicien', esc(userLabel(user(techOf(e)))))}
      ${row('Description', esc(e.description).replace(/\n/g, '<br>'))}
      ${row('Installation', esc(e.install))}
      ${row('Montant HTVA', eur(entryAmount(e)))}
    </table>
    <p class="small muted">Les interventions facturées ne se modifient pas. Pour corriger : « Voir la facture » → « Annuler la facturation ».</p>
    <div class="actions">${e.taskId ? '<button class="btn" id="vTask">📋 Tâche d\'origine</button>' : ''}${b && b.key !== 'import' ? '<button class="btn" id="vBill">Voir la facture</button>' : ''}
      <button class="btn primary" id="vOk">Fermer</button></div>`;
  if ($('#vBill')) $('#vBill').onclick = () => showBill(b);
  if ($('#vTask')) $('#vTask').onclick = () => Taches.open(e.taskId);
  $('#vOk').onclick = ev => { ev.preventDefault(); dlg.close(); };
  if (!dlg.open) dlg.showModal();
}

// ---- Administration : sociétés (commanditaires), utilisateurs, rôles et clients visibles ----
// Étape 1 (multi-societes.md) : seul Fred utilise l'app ; les droits seront appliqués aux autres sociétés avec le serveur (étape 2).
const ROLES = { admin: 'Administrateur', resp: 'Responsable', tech: 'Technicien' };
function viewAdmin() {
  const socs = liveSocs(), us = liveUsers();
  const nUsers = s => us.filter(u => u.societeId === s.id).length;
  const nClients = s => state.clients.filter(c => !c.deleted && clientSoc(c) === s.id).length;
  const seeLabel = u => (u.roles || []).includes('admin') ? 'tous les clients'
    : Object.entries(u.clients || {}).map(([k, v]) => v === 'all' ? `${soc(k).name} : tous` : v?.length ? `${soc(k).name} : ${v.length}` : '').filter(Boolean).join(', ') || 'aucun client';
  $('#view').innerHTML = `
    <div class="card"><div class="row"><h2 class="grow" style="margin:0;border:0;padding:0">Sociétés (${socs.length})</h2><button class="btn" id="socAdd">+ Société</button></div>
      <p class="small muted" style="margin:6px 0 8px">One Level UP et les commanditaires pour qui tu interviens. Leurs tarifs, client Billit et échéance servent à facturer tes interventions pour eux.</p>
      <ul class="list">${socs.map(s => `<li class="tap" data-soc="${s.id}"><div class="grow" style="flex:1;min-width:0"><strong>${esc(s.name)}</strong>
          ${s.home ? '<span class="pill ok">ta société</span>' : ''}${s.active === false ? '<span class="pill">inactive</span>' : ''}
          <div class="small muted">${(s.home ? ['tarifs clients : Réglages'] : [`Billit : ${s.billitName || '—'}`, `${eur(socRates(s).h)} / h`, `échéance ${s.due ?? 15} j`])
            .concat([`${nUsers(s)} utilisateur(s)`, `${nClients(s)} client(s)`]).map(esc).join(' · ')}</div></div><span class="chev">›</span></li>`).join('')}</ul>
    </div>
    <div class="card"><div class="row"><h2 class="grow" style="margin:0;border:0;padding:0">Utilisateurs (${us.length})</h2><button class="btn" id="usrAdd">+ Utilisateur</button></div>
      <p class="small muted" style="margin:6px 0 8px">Techniciens et responsables des sociétés, reconnus par leur adresse Microsoft 365.</p>
      <ul class="list">${us.map(u => `<li class="tap" data-usr="${u.id}"><div class="grow" style="flex:1;min-width:0"><strong>${esc(u.name)}</strong>
          ${u.id === me().id ? '<span class="pill ok">toi</span>' : ''}${u.active === false ? '<span class="pill">inactif</span>' : ''}
          <div class="small muted">${[soc(u.societeId).name, u.email || 'adresse à compléter', (u.roles || []).map(r => ROLES[r]).join(', '), seeLabel(u)].map(esc).join(' · ')}</div></div><span class="chev">›</span></li>`).join('')}</ul>
    </div>
    <div class="card"><h2>Droits</h2>
      <ul class="small" style="margin:0;padding-left:18px">
        <li><strong>Administrateur</strong> : tout, dont cet onglet, la facturation, les achats et les réglages.</li>
        <li><strong>Responsable</strong> : les interventions dont sa société est le commanditaire et celles de ses techniciens, durées seulement.</li>
        <li><strong>Technicien</strong> : saisit, modifie et supprime ses interventions ; voit celles dont sa société est le commanditaire, durées seulement.</li>
        <li>Le technicien d'une intervention est toujours l'utilisateur connecté qui la crée ; il ne se change pas, même par un administrateur.</li>
        <li>Chaque utilisateur ne choisit que parmi ses <strong>clients visibles</strong> (par société).</li></ul>
      <p class="small muted">Pour l'instant, toi seul as accès aux données. Les autres sociétés pourront se connecter quand le serveur commun sera en place (étape 2) :
        les droits ci-dessus seront alors appliqués.</p>
    </div>`;
  $('#socAdd').onclick = () => editSoc(null);
  $('#usrAdd').onclick = () => editUser(null);
  $$('[data-soc]').forEach(li => li.onclick = () => editSoc(soc(li.dataset.soc)));
  $$('[data-usr]').forEach(li => li.onclick = () => editUser(user(li.dataset.usr)));
}

function editSoc(s) {
  const dlg = $('#dlg'), f = $('#dlgForm');
  // Nouvelle société : tarifs et articles d'Alescom (partenaire) proposés.
  const r = s ? socRates(s) : { ...socRates(soc('alescom')) }, home = !!s?.home;
  const rate = (k, label, step = '0.01') => `<div><label>${label}</label><input type="number" step="${step}" id="s_${k}" value="${esc(r[k] ?? '')}"></div>`;
  const txt = (k, label) => `<div><label>${label}</label><input id="s_${k}" value="${esc(r[k] ?? '')}" autocapitalize="off"></div>`;
  f.innerHTML = `<h2>${s ? esc(s.name) : 'Nouvelle société'}</h2>
    <label>Nom</label><input id="sName" value="${esc(s?.name ?? '')}" placeholder="ex. Société X">
    ${home ? '<p class="small muted">Ta société : ses tarifs (clients) se règlent dans Réglages.</p>' : `
    <label>Tenant Microsoft 365 <span class="small">(domaine ou ID, pour la connexion de ses utilisateurs plus tard)</span></label>
    <input id="sTenant" value="${esc(s?.tenant ?? '')}" placeholder="ex. societe-x.be" autocapitalize="off">
    <label>Nom du client dans Billit <span class="small">(à qui tu factures ton travail pour eux)</span></label>
    <input id="sBillit" value="${esc(s?.billitName ?? '')}" placeholder="ex. SOCIETE X SRL">
    <h3 style="margin:14px 0 0">Tes tarifs pour ce commanditaire</h3>
    <div class="grid3">${rate('h', '€ / heure')}${rate('m', '€ / minute')}${rate('dep', '€ déplacement')}</div>
    <div class="grid3">${txt('hRef', 'Article heure')}${txt('mRef', 'Article minute')}${txt('depRef', 'Article déplacement')}</div>
    <label>Échéance des factures (jours)</label><input type="number" step="1" min="0" id="sDue" value="${esc(s?.due ?? 15)}" style="max-width:120px">
    <label class="check"><input type="checkbox" id="sActive" ${s?.active === false ? '' : 'checked'}> Active (proposée comme commanditaire)</label>`}
    <div class="actions"><button class="btn" value="cancel">Annuler</button><button class="btn primary" id="sSave">Enregistrer</button></div>`;
  $('#sSave').onclick = async ev => {
    ev.preventDefault();
    const name = $('#sName').value.trim().replace(/\s+/g, ' ');
    if (!name) return alert('Nom requis.');
    if (liveSocs().some(x => x !== s && fold(x.name) === fold(name))) return alert('Cette société existe déjà.');
    const data = { name };
    if (!home) {
      const num = k => { const v = parseFloat($('#s_' + k).value); return isNaN(v) ? 0 : v; };
      const ref = k => $('#s_' + k).value.trim();
      // Les libellés (recap, fichier Prestations) suivent le nom de la société et les articles.
      const rates = { ...r, h: num('h'), m: num('m'), dep: num('dep'), hRef: ref('hRef'), mRef: ref('mRef'), depRef: ref('depRef') };
      if (!s || s.id !== 'alescom') Object.assign(rates, { hLabel: `Main d'œuvre ${name} heure (${rates.hRef})`, mLabel: `Main d'œuvre ${name} minutes (${rates.mRef})`, depLabel: `Déplacement ${name} (${rates.depRef})` });
      Object.assign(data, { tenant: $('#sTenant').value.trim(), billitName: $('#sBillit').value.trim(), rates,
        due: Math.max(0, parseInt($('#sDue').value, 10) || 0), active: $('#sActive').checked });
    }
    if (s) Object.assign(touch(s), data); else state.societes.push(touch({ id: uid('s'), billitIds: {}, ...data }));
    dlg.close(); await persist(); toast('Enregistré');
  };
  if (!dlg.open) dlg.showModal();
}

function editUser(u) {
  const dlg = $('#dlg'), f = $('#dlgForm');
  const roles = u?.roles || ['tech'], see = { ...(u?.clients || {}) };
  if (!u) see[HOME] = see[HOME] || [];
  const socs = activeSocs();
  const clientsOf = s => state.clients.filter(c => !c.deleted && !c.archived && clientSoc(c) === s.id).sort((a, b) => a.name.localeCompare(b.name));
  // Clients visibles : par société, tous / une sélection / aucun.
  const seeBlock = s => {
    const v = see[s.id], mode = v === 'all' ? 'all' : Array.isArray(v) && v.length ? 'some' : 'none', cs = clientsOf(s);
    return `<div class="card" style="margin:8px 0;padding:8px 10px"><div class="row"><strong class="grow">${esc(s.name)}</strong>
        <select data-see="${s.id}" style="width:auto">
          <option value="none" ${mode === 'none' ? 'selected' : ''}>Aucun client</option>
          <option value="all" ${mode === 'all' ? 'selected' : ''}>Tous (${cs.length})</option>
          <option value="some" ${mode === 'some' ? 'selected' : ''}>Sélection</option></select></div>
      <div class="${mode === 'some' ? '' : 'hidden'}" data-seebox="${s.id}" style="max-height:180px;overflow:auto;margin-top:6px">
        ${cs.length ? cs.map(c => `<label class="check"><input type="checkbox" data-seec="${s.id}" value="${c.id}" ${Array.isArray(v) && v.includes(c.id) ? 'checked' : ''}> ${esc(c.name)}</label>`).join('')
          : '<span class="small muted">Aucun client rattaché à cette société (fiche client → Société).</span>'}</div></div>`;
  };
  f.innerHTML = `<h2>${u ? esc(u.name) : 'Nouvel utilisateur'}</h2>
    <label>Nom</label><input id="uName" value="${esc(u?.name ?? '')}" placeholder="ex. Geoffrey">
    <label>Adresse Microsoft 365</label><input id="uEmail" type="email" value="${esc(u?.email ?? '')}" placeholder="prenom@societe.be" autocapitalize="off">
    <label>Société</label><select id="uSoc">${socOptions(u?.societeId || HOME)}</select>
    <label>Rôles</label>
    <div class="row">${Object.entries(ROLES).map(([k, l]) => `<label class="check grow"><input type="checkbox" data-role="${k}" ${roles.includes(k) ? 'checked' : ''}> ${l}</label>`).join('')}</div>
    <label class="check"><input type="checkbox" id="uActive" ${u?.active === false ? '' : 'checked'}> Actif</label>
    <h3 style="margin:14px 0 0">Clients visibles</h3>
    <p class="small muted" id="uSeeAdmin" style="margin:4px 0">Un administrateur voit tous les clients.</p>
    <div id="uSee">${socs.map(seeBlock).join('')}</div>
    <div class="actions"><button class="btn" value="cancel">Annuler</button><button class="btn primary" id="uSave">Enregistrer</button></div>`;
  const isAdm = () => $('[data-role="admin"]').checked;
  const upAdm = () => { $('#uSee').classList.toggle('hidden', isAdm()); $('#uSeeAdmin').classList.toggle('hidden', !isAdm()); };
  $('[data-role="admin"]').onchange = upAdm; upAdm();
  $$('[data-see]', f).forEach(sel => sel.onchange = () => $(`[data-seebox="${sel.dataset.see}"]`).classList.toggle('hidden', sel.value !== 'some'));
  $('#uSave').onclick = async ev => {
    ev.preventDefault();
    const name = $('#uName').value.trim().replace(/\s+/g, ' '), email = $('#uEmail').value.trim().toLowerCase();
    const rs = $$('[data-role]', f).filter(x => x.checked).map(x => x.dataset.role), active = $('#uActive').checked;
    if (!name) return alert('Nom requis.');
    if (!rs.length) return alert('Choisis au moins un rôle.');
    if (email && !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) return alert('Adresse e-mail invalide.');
    if (email && liveUsers().some(x => x !== u && fold(x.email) === fold(email))) return alert('Cette adresse est déjà utilisée par un autre utilisateur.');
    // Toujours au moins un administrateur actif (sinon plus personne n'accède à cet onglet).
    const admins = liveUsers().filter(x => x !== u && x.active !== false && (x.roles || []).includes('admin'));
    if (!admins.length && !(rs.includes('admin') && active)) return alert('Il faut garder au moins un administrateur actif.');
    const clients = {};
    $$('[data-see]', f).forEach(sel => {
      const id = sel.dataset.see;
      if (sel.value === 'all') clients[id] = 'all';
      else if (sel.value === 'some') clients[id] = $$(`[data-seec="${id}"]`, f).filter(x => x.checked).map(x => x.value);
    });
    const data = { name, email, societeId: $('#uSoc').value, roles: rs, active, clients };
    if (u) Object.assign(touch(u), data); else state.users.push(touch({ id: uid('u'), ...data }));
    dlg.close(); await persist(); toast('Enregistré');
  };
  if (!dlg.open) dlg.showModal();
}

// ---- Réglages ----
function viewSettings() {
  const od = state.onedrive, s = state.settings, r = s.rates;
  const rateBlock = (k, title) => `<h2>${title}</h2><div class="grid3">
      <div><label>€ / heure</label><input type="number" step="0.01" data-rate="${k}.h" value="${r[k].h}"></div>
      <div><label>€ / minute</label><input type="number" step="0.01" data-rate="${k}.m" value="${r[k].m}"></div>
      <div><label>€ déplacement</label><input type="number" step="0.01" data-rate="${k}.dep" value="${r[k].dep}"></div></div>`;
  $('#view').innerHTML = `
    ${installPrompt ? '<button class="btn primary big" id="install" style="margin-bottom:12px">Installer l\'application</button>' : ''}
    <div class="card"><div class="row"><h2 class="grow" style="margin:0">Clients (${state.clients.filter(c => !c.deleted).length})</h2><button class="btn" id="addClient">+ Ajouter</button></div>
      <input type="search" id="cSearch" placeholder="Rechercher un client…" style="margin-top:10px">
      <ul class="list" style="margin-top:8px;max-height:420px;overflow:auto">${state.clients.filter(c => !c.deleted).sort((a, b) => a.name.localeCompare(b.name)).map(c => `
        <li data-name="${esc(fold(c.name + ' ' + (c.billitName || '') + ' ' + (c.domains || []).join(' ')))}">
          <div class="grow tap" data-client="${c.id}"><strong>${esc(c.name)}</strong> ${c.archived ? '<span class="pill">archivé</span>' : ''}
            <div class="small muted">Billit : ${esc(c.billitName || '—')}${c.near ? ' · proche (zone 1)' : ''}${c.server ? ` · ${serverCount(c)} serveur(s)${c.m365Enforce ? ' · M365 Enforce' : ''}${c.esetCloud ? ' · ESET Cloud' : ''}` : ''}${clientSoc(c) !== HOME ? ' · ' + esc(soc(clientSoc(c)).name) : ''}${c.domains?.length ? ' · ' + esc(c.domains.join(', ')) : ''}</div></div>
          <button class="btn danger chip" data-delclient="${c.id}" title="Supprimer">🗑</button></li>`).join('')}</ul>
    </div>

    <div class="card">${rateBlock('client', 'Tarifs clients')}
      <div class="grid2" style="margin-top:4px">
        <div><label>€ déplacement client proche (dep-003)</label><input type="number" step="0.01" data-rate="zone1.dep" value="${r.zone1.dep}"></div>
      </div>
      <p class="small muted">Tarifs des commanditaires (Alescom…) : onglet Administration → Sociétés.</p></div>

    <div class="card"><h2>OneDrive</h2>
      <p class="small" id="odStatus">${esc(statusText())}</p>
      ${OneDrive.signedIn() ? `<p class="small muted">Compte : ${esc(OneDrive.user())}</p>` : ''}
      <div class="row">
        ${OneDrive.signedIn()
          ? '<button class="btn" id="odSync">Synchroniser</button><button class="btn danger" id="odOut">Se déconnecter</button>'
          : `<button class="btn primary" id="odIn" ${od.clientId ? '' : 'disabled'}>Se connecter à Microsoft 365</button>`}
      </div>
      <details style="margin-top:10px" ${od.clientId ? '' : 'open'}><summary class="small">Configuration</summary>
        <label>ID d'application (client ID Entra)</label><input id="odClient" value="${esc(od.clientId)}" placeholder="xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx">
        <label>Tenant (ID ou domaine, sinon "organizations")</label><input id="odTenant" value="${esc(od.tenant)}">
        <label>Dossier TimeSheet dans OneDrive</label><input id="odFolder" value="${esc(od.folder)}">
        <p class="small muted">URI de redirection à déclarer (plateforme « Application monopage ») :<br><code>${esc(OneDrive.redirectUri())}</code></p>
        <p class="small muted">Fichiers générés (ne pas les modifier à la main, ils sont réécrits par l'app) :<br>
          <code>${esc(P.open())}</code><br><code>${esc(P.archive())}</code><br><code>${esc(od.folder)}/A facturer/Prestations - …xlsx</code></p>
      </details>
    </div>

    <div class="card"><h2>Billit</h2>
      <div class="row">
        <label class="check grow"><input type="radio" name="bMode" value="test" ${Billit.mode() !== 'live' ? 'checked' : ''}> Test (sandbox)</label>
        <label class="check grow"><input type="radio" name="bMode" value="live" ${Billit.mode() === 'live' ? 'checked' : ''} ${ENV === 'prod' ? '' : 'disabled'}> Réel</label>
      </div>
      ${ENV === 'prod' ? '' : '<p class="small"><span class="pill warn">Version de développement : Billit toujours en Test (sandbox), le mode Réel est bloqué.</span></p>'}
      <label class="check" style="margin-top:8px"><input type="checkbox" id="bRelay" ${s.billit.relay ? 'checked' : ''}> Clés API sur le serveur OVH (recommandé)</label>
      ${['test', 'live'].map(k => `<div class="grid2">
        <div><label>Clé API (${k === 'live' ? 'réel' : 'test'})</label>${s.billit.relay
          ? `<div class="row"><input type="password" autocomplete="off" data-sk="${k}" placeholder="Saisir pour remplacer" style="flex:1">
              <button class="btn" data-skdel="${k}">Effacer</button></div><span class="small" data-skst="${k}">Serveur…</span>`
          : `<input type="password" autocomplete="off" data-bk="${k}.apiKey" value="${esc(state.billitKeys[k])}">`}</div>
        <div><label>PartyID (${k === 'live' ? 'réel' : 'test'})</label><input inputmode="numeric" data-bk="${k}.partyId" value="${esc(s.billit[k].partyId)}"></div>
      </div>`).join('')}
      <p class="small muted">${s.billit.relay
        ? `🔒 Les clés API sont sur le serveur OVH (${esc(new URL(API_URL).host + new URL(API_URL).pathname)}), chiffrées : jamais sur les appareils ni sur OneDrive.
          Une clé saisie ici est contrôlée par Billit puis enregistrée sur le serveur, une seule fois pour tous les appareils.`
        : '🔒 Les clés API restent sur cet appareil (jamais sur OneDrive) : à saisir une fois sur chaque appareil, effacées à la déconnexion.'}</p>
      <div class="row" style="margin-top:10px"><button class="btn" id="bTest">Tester la connexion</button><span class="small" id="bTestRes"></span></div>
      ${ENV === 'prod' ? '' : `<h3 style="margin:16px 0 4px">Copier Billit réel → sandbox</h3>
      <p class="small muted">Clients, articles (catégories par leur nom) et factures de vente et d'achat des 3 derniers mois, avec leurs fournisseurs.
        Le Billit réel est seulement lu (clé « réel » ci-dessus). Ventes copiées en Projet ; clients et fournisseurs sans e-mail ni Peppol.
        Relancer ne crée pas de doublon. Pour que l'onglet Achats importe les factures d'achat copiées, ouvre-le une fois <em>avant</em> la copie.</p>
      <div class="row"><button class="btn" id="cpGo">Préparer la copie</button><button class="btn" id="cpArt">Copier les articles seulement</button></div>
      <pre class="recap hidden" id="cpLog"></pre>`}
      <p class="small muted">Clé API : dans Billit, menu de ton profil → API. Laisse le PartyID vide puis clique « Tester la connexion » : il est retrouvé automatiquement.
        La clé donne accès à tout ton Billit : tu peux la révoquer à tout moment dans Billit.</p>
    </div>

    <div class="card"><h2>Code PIN</h2>
      <p class="small muted">${state.settings.pin ? 'Un code PIN est demandé à l\'ouverture de l\'app et après 5 minutes en arrière-plan, sur tous tes appareils.'
        : 'Aucun code PIN : l\'app s\'ouvre dès que tu es connecté à Microsoft 365.'}</p>
      <div class="row"><button class="btn" id="pinSet">${state.settings.pin ? 'Changer le code PIN' : 'Définir un code PIN'}</button>
        ${state.settings.pin ? '<button class="btn danger" id="pinDel">Supprimer le code PIN</button>' : ''}</div>
    </div>

    <div class="card"><h2>Reprise des données</h2>
      <div class="row">
        <button class="btn" id="importAT">Importer ArchiveTemplate.xlsx</button>
        <button class="btn" id="importHT">Importer un CSV Hours Tracker</button>
        <button class="btn" id="backup">⬇ Sauvegarde JSON</button>
      </div>
      <p class="small muted">ArchiveTemplate.xlsx : les lignes sont reprises « à facturer ». Lu depuis OneDrive (Archives) si connecté, sinon choisis le fichier.</p>
      <input type="file" id="fileAT" accept=".xlsx" class="hidden">
      <input type="file" id="fileHT" accept=".csv,text/csv" class="hidden">
      <p class="small muted" id="persistInfo"></p>
      <p class="small muted" id="appVersion"></p>
    </div>`;

  if (installPrompt) $('#install').onclick = async () => { installPrompt.prompt(); await installPrompt.userChoice; installPrompt = null; render(); };
  $('#addClient').onclick = () => editClient(null);
  $$('[data-client]').forEach(li => li.onclick = () => editClient(state.clients.find(c => c.id === li.dataset.client)));
  $$('[data-delclient]').forEach(b => b.onclick = () => deleteClient(state.clients.find(c => c.id === b.dataset.delclient)));
  $('#cSearch').oninput = ev => {
    const q = fold(ev.target.value.trim());
    $$('li[data-name]').forEach(li => li.classList.toggle('hidden', !!q && !li.dataset.name.includes(q)));
  };
  const saveSettings = async () => { state.settingsAt = Date.now(); await DB.save(state); scheduleSync(); toast('Enregistré'); };
  $$('[data-rate]').forEach(inp => inp.onchange = () => {
    const [k, fld] = inp.dataset.rate.split('.'); const v = parseFloat(inp.value);
    if (!isNaN(v)) { r[k][fld] = v; saveSettings(); }
  });

  const saveOd = async () => {
    Object.assign(od, { clientId: $('#odClient').value.trim(), tenant: $('#odTenant').value.trim() || 'organizations',
      folder: $('#odFolder').value.trim().replace(/^\/+|\/+$/g, '') || DEFAULT_FOLDER });
    state.published = {};
    await DB.save(state); render(); sync();
  };
  ['#odClient', '#odTenant', '#odFolder'].forEach(sel => $(sel).onchange = saveOd);
  if ($('#odIn')) $('#odIn').onclick = () => OneDrive.login().catch(e => alert(e.message));
  if ($('#odSync')) $('#odSync').onclick = () => { state.published = {}; sync(); };
  if ($('#odOut')) $('#odOut').onclick = () => { if (confirm('Se déconnecter ?\nLa copie locale de cet appareil est effacée ; tes données restent dans OneDrive.')) signOut(); };
  $$('input[name="bMode"]').forEach(x => x.onchange = async () => {
    if (x.value === 'live' && !confirm('Passer en mode RÉEL ?\nLes factures seront créées dans ta vraie comptabilité Billit.')) { render(); return; }
    s.billit.mode = x.value; await saveSettings(); render();
  });
  $('#bRelay').onchange = async ev => {
    if (ev.target.checked && !confirm('Mettre les clés API Billit sur le serveur OVH ?\n\nLes clés enregistrées sur les appareils seront effacées (sur tous les appareils, à leur prochaine synchro).'
      + '\nSaisis ensuite chaque clé ici une seule fois : elle est enregistrée sur le serveur.')) { ev.target.checked = false; return; }
    if (!ev.target.checked && !confirm('Revenir aux clés API sur chaque appareil ?\nIl faudra ressaisir la clé sur chaque appareil (celles du serveur restent, efface-les au besoin avant).')) { ev.target.checked = true; return; }
    s.billit.relay = ev.target.checked;
    normalizeSettings(); await saveSettings(); render();
  };
  // Clés du serveur : état (date, auteur) lu sur le serveur ; saisie = contrôle par Billit puis enregistrement.
  const srvKeyState = (k, b) => {
    const el = $(`[data-skst="${k}"]`);
    if (!el) return;
    el.innerHTML = !b?.access ? '<span class="pill warn">Non autorisé sur ce serveur</span>'
      : b.key ? `<span class="pill ok">✓ Sur le serveur${b.access === 'ro' ? ' (lecture seule)' : ''}</span> <span class="muted">${esc(frDate(String(b.at).slice(0, 10)))} · ${esc(b.by || '')}</span>`
      : '<span class="pill warn">Aucune clé sur le serveur</span>';
  };
  if (s.billit.relay) Serveur.me()
    .then(i => { if (!i.admin) throw new Error('réservé aux administrateurs'); ['test', 'live'].forEach(k => srvKeyState(k, i.billit[k])); })
    .catch(e => $$('[data-skst]').forEach(el => el.innerHTML = `<span class="pill warn">${esc(e.message)}</span>`));
  $$('[data-sk]').forEach(inp => inp.onchange = async () => {
    const k = inp.dataset.sk, v = inp.value.trim();
    if (!v) return;
    $(`[data-skst="${k}"]`).textContent = 'Contrôle par Billit…';
    try {
      const r = await Serveur.saveBillitKey(k, v);
      inp.value = '';
      srvKeyState(k, r.billit);
      const cos = r.companies || [];
      if (!s.billit[k].partyId && cos.length) {
        let co = cos[0];
        if (cos.length > 1) {
          const i = prompt('Plusieurs sociétés pour cette clé :\n' + cos.map((c, n) => `${n + 1}. ${c.Name} (PartyID ${c.PartyID})`).join('\n') + '\n\nNuméro de la société :', '1');
          co = cos[(+i || 1) - 1] || cos[0];
        }
        s.billit[k].partyId = String(co.PartyID);
        await saveSettings(); render();
      }
      toast('Clé enregistrée sur le serveur');
    } catch (e) { inp.value = ''; $(`[data-skst="${k}"]`).innerHTML = `<span class="pill warn">${esc(e.message)}</span>`; }
  });
  $$('[data-skdel]').forEach(b => b.onclick = async () => {
    const k = b.dataset.skdel;
    if (!confirm(`Effacer la clé API Billit (${k === 'live' ? 'réel' : 'test'}) du serveur ?\nPlus aucun appareil ne pourra appeler Billit dans ce mode.`)) return;
    try { srvKeyState(k, (await Serveur.saveBillitKey(k, '')).billit); toast('Clé effacée du serveur'); }
    catch (e) { $(`[data-skst="${k}"]`).innerHTML = `<span class="pill warn">${esc(e.message)}</span>`; }
  });
  $$('[data-bk]').forEach(inp => inp.onchange = () => {
    const [k, fld] = inp.dataset.bk.split('.');
    if (fld === 'apiKey') { state.billitKeys[k] = inp.value.trim(); DB.save(state); }   // clé : cet appareil seulement
    else { s.billit[k][fld] = inp.value.trim(); saveSettings(); }
  });
  $('#bTest').onclick = async () => {
    const out = $('#bTestRes'); out.textContent = 'Connexion…';
    try {
      const a = await Billit.account();
      const companies = a?.Companies || a?.companies || [];
      const k = s.billit[Billit.mode()];
      // PartyID vide : on le reprend de la société liée à la clé (choix si plusieurs sociétés).
      if (!k.partyId) {
        let co = companies[0];
        if (companies.length > 1) {
          const i = prompt('Plusieurs sociétés pour cette clé :\n' + companies.map((c, n) => `${n + 1}. ${c.Name || c.CommercialName} (PartyID ${c.PartyID})`).join('\n') + '\n\nNuméro de la société :', '1');
          co = companies[(+i || 1) - 1];
        }
        if (!co?.PartyID) throw new Error('PartyID introuvable dans la réponse de Billit : ' + JSON.stringify(a).slice(0, 200));
        k.partyId = String(co.PartyID);
        await saveSettings();
        render();
        return toast(`PartyID ${co.PartyID} (${co.Name || ''}) enregistré`);
      }
      const co = companies.find(c => String(c.PartyID) === String(k.partyId)) || companies[0];
      out.innerHTML = `<span class="pill ok">✓ Connecté (${esc(billitModeLabel(Billit.mode()))}) : ${esc(co?.Name || a?.Email || 'OK')} · PartyID ${esc(k.partyId)}</span>`;
    } catch (e) { out.innerHTML = `<span class="pill warn">${esc(e.message)}</span>`; }
  };
  // Copie complète (clients, articles, factures des 3 derniers mois) ou articles seulement.
  const copyBillit = async productsOnly => {
    const out = $('#cpLog'), btns = $$('#cpGo, #cpArt'), log = t => { out.textContent += t + '\n'; out.scrollTop = out.scrollHeight; };
    out.textContent = ''; out.classList.remove('hidden'); btns.forEach(b => b.disabled = true);
    try {
      const plan = await CopieBillit.prepare(3, log, { productsOnly });
      log(CopieBillit.summary(plan));
      if (!confirm(CopieBillit.summary(plan) + '\n\nLancer la copie vers le sandbox ?')) return log('Copie annulée.');
      await CopieBillit.run(plan, log);
    } catch (e) { log('Erreur : ' + e.message); }
    finally { btns.forEach(b => b.disabled = false); }
  };
  if ($('#cpGo')) $('#cpGo').onclick = () => copyBillit(false);
  if ($('#cpArt')) $('#cpArt').onclick = () => copyBillit(true);
  $('#pinSet').onclick = () => editPin();
  if ($('#pinDel')) $('#pinDel').onclick = async () => {
    if (!confirm('Supprimer le code PIN (sur tous tes appareils) ?')) return;
    delete state.settings.pin; state.settingsAt = Date.now(); await persist(); toast('Code PIN supprimé');
  };

  $('#backup').onclick = () => download(new Blob([JSON.stringify({ entries: state.entries, clients: state.clients, bills: state.bills, societes: state.societes, users: state.users }, null, 1)],
    { type: 'application/json' }), `heures-sauvegarde-${today()}.json`);
  $('#importHT').onclick = () => $('#fileHT').click();
  $('#fileHT').onchange = async ev => {
    const f = ev.target.files[0]; ev.target.value = ''; if (!f) return;
    try { importRows(Csv.readHoursTracker(await f.text()), 'CSV Hours Tracker', true); } catch (e) { alert(e.message); }
  };
  $('#importAT').onclick = async () => {
    if (!OneDrive.signedIn()) return $('#fileAT').click();
    try {
      const f = await OneDrive.get(P.template(), 'buffer');
      if (!f) { alert(`${P.template()} introuvable : choisis le fichier.`); return $('#fileAT').click(); }
      importRows(await Xlsx.readArchiveTemplate(f.data), 'ArchiveTemplate.xlsx', false);
    } catch (e) { alert(e.message); }
  };
  $('#fileAT').onchange = async ev => {
    const f = ev.target.files[0]; ev.target.value = ''; if (!f) return;
    try { importRows(await Xlsx.readArchiveTemplate(await f.arrayBuffer()), 'ArchiveTemplate.xlsx', false); } catch (e) { alert(e.message); }
  };
  // Version installée (nom du cache du service worker, ex. heures-v23) : pour vérifier qu'une publication est bien arrivée.
  window.caches?.keys().then(ks => {
    const v = ks.filter(k => /^heures-v\d+$/.test(k)).sort((a, b) => +b.slice(8) - +a.slice(8))[0];
    if ($('#appVersion')) $('#appVersion').textContent = `Version : ${v || 'inconnue (pas de cache)'}`;
  }).catch(() => {});
  navigator.storage?.persisted?.().then(p => {
    const el = $('#persistInfo');
    if (el) el.textContent = p ? 'Cache local persistant.' : 'Installe l\'app pour que le navigateur ne vide pas le cache local.';
  });
}

// Suppression "douce" : le client disparaît des listes, ses interventions gardent son nom dans l'historique.
async function deleteClient(c) {
  const n = live().filter(e => e.clientId === c.id).length;
  const open = live().filter(e => e.clientId === c.id && !e.billId).length;
  if (!confirm(`Supprimer le client « ${c.name} » ?` + (n ? `\n\n${n} intervention(s) liée(s) restent dans l'historique${open ? `, dont ${open} pas encore facturée(s)` : ''}.` : ''))) return;
  c.deleted = true; touch(c);
  if (state.lastClientId === c.id) state.lastClientId = null;
  if ($('#dlg').open) $('#dlg').close();
  await persist(); toast('Client supprimé');
}

// opts (création depuis un autre écran, ex. un article d'achat) : preset = valeurs proposées ({ name, domains }),
// onSaved(client) / onCancel() = retour à cet écran au lieu de fermer la fenêtre.
function editClient(c, opts = {}) {
  const dlg = $('#dlg'), f = $('#dlgForm'), pre = opts.preset || {}, mode = Billit.mode();
  // Lien Billit (choix dans la liste) : seulement si ce client n'est pas encore relié à un client Billit dans ce mode.
  const askBillit = !c?.billitIds?.[mode];
  let billitId = null;
  f.innerHTML = `<h2>${c ? 'Modifier' : 'Nouveau'} client</h2>
    ${askBillit ? `<label>Client Billit (${billitModeLabel(mode)})</label>
      <select id="cBillitPick" disabled><option value="">Chargement des clients Billit…</option></select>
      <p class="small muted" id="cBillitInfo" style="margin:4px 0 0">Clients Billit pas encore reliés à un client de l'app.</p>` : ''}
    <label>Nom</label><input id="cName" value="${esc(c?.name ?? pre.name ?? '')}" placeholder="ex. Vedi">
    <label>Nom du client dans Billit</label><input id="cBillit" value="${esc(c?.billitName)}" placeholder="ex. V.T.S. SA">
    <label>Nom(s) de domaine <span class="small">(séparés par une virgule, sert à attribuer les achats OVH…)</span></label>
    <input id="cDomains" value="${esc((c?.domains || pre.domains || []).join(', '))}" placeholder="ex. vedi.be, vts-sa.com" autocapitalize="off">
    <label class="check"><input type="checkbox" id="cNear" ${c?.near ? 'checked' : ''}> Client proche (déplacement zone 1, dep-003)</label>
    <label>Société du client (commanditaire proposé par défaut)</label><select id="cSoc">${socOptions(c ? clientSoc(c) : HOME)}</select>
    <div class="row">
      <label class="check grow"><input type="checkbox" id="cServer" ${c?.server ? 'checked' : ''}> Serveur(s)</label>
      <div id="cServersBox" class="${c?.server ? '' : 'hidden'}" style="width:140px"><label style="margin-top:0">Nombre de serveurs</label>
        <input type="number" id="cServers" min="1" step="1" value="${c?.servers || 1}"></div>
    </div>
    <div id="cSrvOpts" class="row ${c?.server ? '' : 'hidden'}" style="margin-top:-4px">
      <label class="check grow"><input type="checkbox" id="cM365" ${c?.m365Enforce ? 'checked' : ''}> M365 Enforce</label>
      <label class="check grow"><input type="checkbox" id="cEset" ${c?.esetCloud ? 'checked' : ''}> ESET Cloud</label>
    </div>
    ${c ? `<label class="check"><input type="checkbox" id="cArch" ${c.archived ? 'checked' : ''}> Archivé (masqué des listes)</label>` : ''}
    <div class="actions">
      ${c ? '<button class="btn danger" id="cDel" style="margin-right:auto">Supprimer</button>' : ''}
      <button class="btn" ${opts.onCancel ? 'id="cCancel"' : 'value="cancel"'}>${opts.onCancel ? 'Retour' : 'Annuler'}</button><button class="btn primary" id="cSave">Enregistrer</button></div>`;
  // M365 Enforce et ESET Cloud : seulement pour un client avec serveur(s).
  $('#cServer').onchange = ev => ['#cServersBox', '#cSrvOpts'].forEach(s => $(s).classList.toggle('hidden', !ev.target.checked));
  if ($('#cDel')) $('#cDel').onclick = () => deleteClient(c);
  if (opts.onCancel) $('#cCancel').onclick = () => opts.onCancel();
  if (askBillit) (async () => {
    const sel = $('#cBillitPick'), info = $('#cBillitInfo');
    try {
      if (!Billit.ready()) throw new Error('clé API Billit absente (Réglages) : nom Billit à saisir à la main.');
      // Déjà reliés : ID mémorisé dans ce mode, ou même nom Billit qu'un client de l'app.
      const used = new Set(state.clients.filter(x => !x.deleted && x !== c).flatMap(x => [x.billitIds?.[mode] && 'id:' + x.billitIds[mode], x.billitName && 'n:' + fold(x.billitName)]).filter(Boolean));
      const free = (await Billit.customers()).filter(p => !used.has('id:' + p.PartyID) && !used.has('n:' + fold(Billit.partyName(p))))
        .sort((a, b) => Billit.partyName(a).localeCompare(Billit.partyName(b)));
      if (!sel.isConnected) return;
      sel.innerHTML = `<option value="">— aucun / saisie à la main —</option>` + free.map(p => `<option value="${esc(p.PartyID)}">${esc(Billit.partyName(p))}${p.VATNumber ? ' · ' + esc(p.VATNumber) : ''}</option>`).join('');
      sel.disabled = false;
      info.textContent = `${free.length} client(s) Billit pas encore relié(s) à l'app.`;
      sel.onchange = () => {
        const p = free.find(x => String(x.PartyID) === sel.value);
        billitId = p ? p.PartyID : null;
        if (!p) return;
        $('#cBillit').value = Billit.partyName(p);
        if (!$('#cName').value.trim()) $('#cName').value = p.CommercialName || Billit.partyName(p);
      };
    } catch (e) { if (sel.isConnected) { sel.innerHTML = '<option value="">—</option>'; info.textContent = 'Liste Billit indisponible : ' + e.message; } }
  })();
  $('#cSave').onclick = async ev => {
    ev.preventDefault();
    const name = $('#cName').value.trim().replace(/\s+/g, ' ');
    if (!name) return alert('Nom requis.');
    if (/\b(remote|distance)\b/i.test(name)) return alert('Retire « remote » / « distance » du nom : utilise la case « Sur site ».');
    if (state.clients.some(x => x !== c && !x.deleted && x.name.toLowerCase() === name.toLowerCase())) return alert('Ce client existe déjà.');
    const server = $('#cServer').checked, servers = Math.max(1, parseInt($('#cServers').value, 10) || 1);
    const domains = [...new Set($('#cDomains').value.toLowerCase().split(/[\s,;]+/)
      .map(d => d.replace(/^https?:\/\//, '').replace(/^www\./, '').replace(/\/.*$/, '')).filter(d => /^[a-z0-9-]+(\.[a-z0-9-]+)+$/.test(d)))];
    const taken = domains.map(d => [d, state.clients.find(x => x !== c && !x.deleted && (x.domains || []).includes(d))]).find(([, x]) => x);
    if (taken) return alert(`Le domaine ${taken[0]} est déjà attribué à ${taken[1].name}.`);
    const data = { name, billitName: $('#cBillit').value.trim(), domains, near: $('#cNear').checked, societeId: $('#cSoc').value, defaultAlescom: $('#cSoc').value === 'alescom', archived: !!$('#cArch')?.checked,
      server, servers: server ? servers : 0, m365Enforce: server && $('#cM365').checked, esetCloud: server && $('#cEset').checked };
    // Client Billit choisi dans la liste : lien mémorisé (comme au 1er envoi d'une facture), pas de recherche par le nom ensuite.
    if (billitId) data.billitIds = { ...c?.billitIds, [mode]: billitId };
    let saved = c;
    if (c) Object.assign(touch(c), data); else state.clients.push(saved = touch({ id: uid('c'), ...data }));
    if (opts.onSaved) { await persist(); opts.onSaved(saved); return; }
    dlg.close(); await persist();
  };
  if (!dlg.open) dlg.showModal();
}

// Reprise d'anciennes données (ArchiveTemplate.xlsx = à facturer ; CSV Hours Tracker = au choix).
function importRows(rows, source, askBilled) {
  if (!rows.length) return alert('Aucune ligne trouvée dans ce fichier.');
  const find = n => state.clients.find(c => !c.deleted && c.name.toLowerCase() === n.toLowerCase());
  const newClients = [...new Set(rows.map(r => r.client).filter(n => !find(n)))];
  const dupKey = (cid, r) => `${cid}|${r.date}|${r.start}`;
  const existing = new Set(live().map(e => dupKey(e.clientId, e)));
  const dups = rows.filter(r => find(r.client) && existing.has(dupKey(find(r.client).id, r)));
  const long = rows.filter(r => r.minutes > 12 * 60);

  const dlg = $('#dlg'), f = $('#dlgForm');
  f.innerHTML = `<h2>Import ${esc(source)}</h2>
    <p>${rows.length} ligne(s) lue(s), dont <strong>${rows.length - dups.length}</strong> nouvelle(s)${dups.length ? ` (${dups.length} déjà présente(s), ignorée(s))` : ''}.</p>
    ${newClients.length ? `<p class="small">Clients créés : ${newClients.map(esc).join(', ')}</p>` : ''}
    ${long.length ? `<p class="small"><span class="pill warn">Durées suspectes (&gt; 12 h)</span></p><ul class="small">${long.map(r => `<li>${esc(r.job)} ${frDate(r.date)} : ${fmtDur(r.minutes)}</li>`).join('')}</ul>` : ''}
    ${askBilled ? `<label class="check"><input type="checkbox" id="iBilled" checked> Déjà facturées (archive « Import historique »)</label>`
      : '<p class="small muted">Ces interventions seront « à facturer ». Une fois l\'import vérifié, tu peux vider ou archiver l\'ancien ArchiveTemplate.xlsx.</p>'}
    <div class="actions"><button class="btn" value="cancel">Annuler</button><button class="btn primary" id="iGo">Importer</button></div>`;
  $('#iGo').onclick = async ev => {
    ev.preventDefault();
    for (const n of newClients) {
      const first = rows.find(r => r.client === n);
      state.clients.push(touch({ id: uid('c'), name: n, billitName: '', defaultOnSite: first.onSite, societeId: first.alescom ? 'alescom' : HOME, defaultAlescom: first.alescom }));
    }
    let b = null;
    if (askBilled && $('#iBilled').checked) state.bills.push(b = touch({ id: uid('b'), key: 'import', label: 'Import historique', date: today(), ref: '', entryIds: [], from: '', to: '', amount: 0 }));
    let n = 0;
    for (const r of rows) {
      const c = find(r.client);
      if (existing.has(dupKey(c.id, r))) continue;
      const e = touch({ id: uid('e'), clientId: c.id, date: r.date, start: r.start, end: r.end, minutes: r.minutes, pauseMin: 0,
        description: r.description, onSite: r.onSite, cmdId: r.alescom ? 'alescom' : HOME, techId: ADMIN, alescom: r.alescom, install: r.install || '', billId: b?.id || null });
      state.entries.push(e); existing.add(dupKey(c.id, r)); n++;
      if (b) b.entryIds.push(e.id);
    }
    if (b) {
      const es = sortEntries(live().filter(e => e.billId === b.id));
      if (es.length) Object.assign(b, { from: es[0].date, to: es[es.length - 1].date, amount: es.reduce((s, e) => s + entryAmount(e), 0) });
      else state.bills.pop();
    }
    dlg.close(); await persist(); toast(`${n} intervention(s) importée(s)`);
  };
  dlg.showModal();
}

// ---------- accès : connexion Microsoft obligatoire + code PIN facultatif ----------
// Le PIN est stocké haché (PBKDF2) dans les réglages synchronisés : même code sur tous les appareils.
// C'est un verrou d'écran ; la vraie protection des données reste la connexion Microsoft (OneDrive).
let locked = false, hiddenAt = 0;
const PIN_RELOCK_MS = 5 * 60000, PIN_MAX_FAILS = 5;

async function pinHash(pin, salt) {
  const key = await crypto.subtle.importKey('raw', new TextEncoder().encode(pin), 'PBKDF2', false, ['deriveBits']);
  const bits = await crypto.subtle.deriveBits({ name: 'PBKDF2', hash: 'SHA-256', salt: new TextEncoder().encode(salt), iterations: 150000 }, key, 256);
  return btoa(String.fromCharCode(...new Uint8Array(bits)));
}

// Déconnexion : on oublie le jeton et la copie locale (les données restent dans OneDrive).
async function signOut() {
  OneDrive.logout();
  await Achats.reset();
  await Taches.reset();
  const od = state.onedrive;
  state = DB.defaults(); state.onedrive = od;
  await DB.save(state);
  locked = false; tab = 'today';
  setSync('login', 'Connecte-toi à Microsoft 365.');
  render();
}

// Première synchro d'un appareil : OneDrive fait foi (les changements faits avant connexion ne l'écrasent pas).
// Seules les interventions créées localement et absentes de OneDrive sont conservées.
function adoptRemote(remote) {
  const rE = remote.entries || [], rC = remote.clients || [];
  const ids = new Set(rE.map(e => e.id));
  const localOnly = state.entries.filter(e => !e.deleted && !ids.has(e.id));
  const need = new Set(localOnly.map(e => e.clientId));
  state.entries = [...rE, ...localOnly];
  state.clients = [...rC, ...state.clients.filter(c => !c.deleted && need.has(c.id) && !rC.some(r => r.id === c.id))];
  state.bills = remote.bills || [];
  state.societes = remote.societes || []; state.users = remote.users || [];
  state.timer = remote.timer ?? null; state.timerAt = remote.timerAt || 0;
  if (remote.settings) { state.settings = remote.settings; state.settingsAt = remote.settingsAt || 0; }
  normalizeSettings();
}

function viewLogin() {
  $('#view').innerHTML = `
    <div class="card" style="text-align:center;margin-top:24px">
      <img src="icons/icon-192.png" alt="" width="72" height="72" style="border-radius:16px">
      <h2 style="font-size:18px;margin:12px 0 4px">LevelDesk</h2>
      <p class="muted">Outil interne de gestion de One Level UP.</p>
      <button class="btn primary big" id="loginBtn" ${OneDrive.configured() ? '' : 'disabled'}>Se connecter avec Microsoft 365</button>
      ${OneDrive.configured() ? '' : '<p class="small"><span class="pill warn">ID d\'application manquant dans la configuration.</span></p>'}
      <p class="small muted" style="margin-bottom:0">Connexion sur la page officielle de Microsoft : le mot de passe n'est jamais saisi ici.</p>
    </div>`;
  $('#loginBtn').onclick = () => OneDrive.login().catch(e => alert(e.message));
}

// Compte Microsoft inconnu, désactivé ou supprimé dans Administration.
function viewDenied() {
  $('#view').innerHTML = `
    <div class="card" style="text-align:center;margin-top:24px">
      <h2 style="font-size:18px">⛔ Accès refusé</h2>
      <p>Le compte <strong>${esc(OneDrive.user())}</strong> n'est pas autorisé dans LevelDesk.</p>
      <p class="small muted">Demande à l'administrateur de l'ajouter (ou de le réactiver) dans l'onglet Admin.</p>
      <button class="btn big" id="deniedOut">Se déconnecter</button>
    </div>`;
  $('#deniedOut').onclick = () => signOut();
}

function viewPin() {
  const fails = state.pinFails || 0;
  $('#view').innerHTML = `
    <div class="card" style="text-align:center;margin-top:24px">
      <h2 style="font-size:18px">🔒 Code PIN</h2>
      <p class="small muted">${esc(OneDrive.user())}</p>
      <form id="pinForm"><input type="password" id="pinIn" inputmode="numeric" autocomplete="off" maxlength="8"
        style="text-align:center;font-size:24px;letter-spacing:.4em;max-width:220px" autofocus>
      ${fails ? `<p class="small" style="color:var(--danger)">Code incorrect (${fails}/${PIN_MAX_FAILS})</p>` : ''}
      <button class="btn primary big" style="margin-top:12px">Déverrouiller</button></form>
      <p class="small" style="margin-top:14px"><button class="link" id="pinForgot">Code oublié ? Se déconnecter</button></p>
      <p class="small muted">Après une reconnexion Microsoft, tu peux changer le code dans Réglages.</p>
    </div>`;
  $('#pinIn').focus();
  $('#pinForm').onsubmit = async ev => {
    ev.preventDefault();
    const p = state.settings.pin;
    if (p && await pinHash($('#pinIn').value, p.salt) === p.hash) {
      // Déverrouillé : toujours sur « Aujourd'hui » (demande de Fred), écran Facturer remis à sa liste.
      locked = false; tab = 'today'; billingTarget = null;
      state.pinFails = 0; await DB.save(state); render(); sync(); return;
    }
    state.pinFails = fails + 1; await DB.save(state);
    if (state.pinFails >= PIN_MAX_FAILS) { alert('Trop d\'essais : déconnexion.'); return signOut(); }
    render();
  };
  $('#pinForgot').onclick = () => { if (confirm('Se déconnecter ? Tu devras te reconnecter avec ton compte Microsoft.')) signOut(); };
}

function editPin() {
  const dlg = $('#dlg'), f = $('#dlgForm');
  f.innerHTML = `<h2>${state.settings.pin ? 'Changer' : 'Définir'} le code PIN</h2>
    <p class="small muted">4 à 8 chiffres. Le même code sera demandé sur tous tes appareils.</p>
    <label>Nouveau code</label><input type="password" id="p1" inputmode="numeric" maxlength="8" autocomplete="new-password">
    <label>Confirmer</label><input type="password" id="p2" inputmode="numeric" maxlength="8" autocomplete="new-password">
    <div class="actions"><button class="btn" value="cancel">Annuler</button><button class="btn primary" id="pSave">Enregistrer</button></div>`;
  $('#pSave').onclick = async ev => {
    ev.preventDefault();
    const a = $('#p1').value, b = $('#p2').value;
    if (!/^\d{4,8}$/.test(a)) return alert('Le code doit contenir 4 à 8 chiffres.');
    if (a !== b) return alert('Les deux codes ne correspondent pas.');
    const salt = btoa(String.fromCharCode(...crypto.getRandomValues(new Uint8Array(16))));
    state.settings.pin = { salt, hash: await pinHash(a, salt) };
    state.settingsAt = Date.now(); state.pinFails = 0;
    dlg.close(); await persist(); toast('Code PIN enregistré');
  };
  dlg.showModal();
}

// ---------- démarrage ----------
// Dans les dialogues, seul le bouton principal est "submit" : Entrée enregistre, jamais supprimer/annuler.
new MutationObserver(() => $$('#dlgForm button').forEach(b => {
  if (b.classList.contains('primary')) return;
  b.type = 'button';
  if (b.value === 'cancel') b.onclick = () => $('#dlg').close();
})).observe($('#dlgForm'), { childList: true });
$('#dlgForm').addEventListener('submit', ev => { ev.preventDefault(); $('#dlgForm .btn.primary')?.click(); });

$$('.tabs button[data-tab]').forEach(b => b.onclick = () => { tab = b.dataset.tab; render(); });
// « Plus » (téléphone) : les onglets secondaires dans une fenêtre (Admin seulement pour un administrateur).
$('#tabMore').onclick = () => {
  const f = $('#dlgForm');
  f.innerHTML = `<h2>Plus</h2><div class="moresheet">${$$('.tabs .sec').filter(b => !b.classList.contains('hidden')).map(b =>
    `<button class="btn ${b.dataset.tab === tab ? 'active' : ''}" data-go="${b.dataset.tab}">${b.innerHTML}</button>`).join('')}</div>
    <div class="actions"><button class="btn" value="cancel">Fermer</button></div>`;
  $$('[data-go]', f).forEach(x => x.onclick = () => { $('#dlg').close(); tab = x.dataset.go; render(); });
  if (!$('#dlg').open) $('#dlg').showModal();
};
window.addEventListener('beforeinstallprompt', e => { e.preventDefault(); installPrompt = e; if (tab === 'settings') render(); });
window.addEventListener('online', sync);
window.addEventListener('offline', () => setSync('offline', 'Hors ligne : synchro au retour du réseau.'));
document.addEventListener('visibilitychange', () => {
  if (document.visibilityState === 'hidden') { hiddenAt = Date.now(); return; }
  if (state?.settings.pin && OneDrive.signedIn() && hiddenAt && Date.now() - hiddenAt > PIN_RELOCK_MS) { locked = true; render(); }
  if (!locked) sync();
});

(async () => {
  state = await DB.load();
  normalizeSettings();
  OneDrive.configure(() => state.onedrive);
  // Réglages Billit synchronisés (mode, PartyID) + clés API de cet appareil.
  Billit.configure(() => {
    const b = state.settings.billit, k = state.billitKeys || {};
    return { mode: b.mode, relay: !!b.relay, test: { ...b.test, apiKey: k.test || '' }, live: { ...b.live, apiKey: k.live || '' } };
  }, (m, id) => { state.settings.billit[m].partyId = String(id); });
  if (mergeRemoteClients()) await DB.save(state);
  let justLoggedIn = false;
  try { if ((justLoggedIn = await OneDrive.handleRedirect())) toast('Connecté'); }
  catch (e) { alert('Connexion Microsoft : ' + e.message); }
  // PIN demandé à l'ouverture, sauf juste après une connexion Microsoft (permet de changer un code oublié).
  locked = !!state.settings.pin && OneDrive.signedIn() && !justLoggedIn;
  if (justLoggedIn) { state.pinFails = 0; await DB.save(state); }
  render();
  if (!locked) sync();
  navigator.storage?.persist?.();
  if ('serviceWorker' in navigator && (location.protocol === 'https:' || location.hostname === 'localhost')) {
    // Nouvelle version publiée : le service worker la met en cache puis prend la main → on recharge une fois
    // pour l'afficher tout de suite (sinon l'ancienne version resterait jusqu'au chargement suivant).
    const hadController = !!navigator.serviceWorker.controller;
    let reloaded = false;
    navigator.serviceWorker.addEventListener('controllerchange', () => {
      if (!hadController || reloaded || $('#dlg').open) return;
      reloaded = true; location.reload();
    });
    navigator.serviceWorker.register('sw.js').then(r => r.update()).catch(() => {});
  }
})();
