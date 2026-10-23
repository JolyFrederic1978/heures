// Heures OLU – pointage des interventions, facturation et archivage, synchronisés sur OneDrive.
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
// Un groupe de facturation = un client, ou "alescom" (commanditaire qui regroupe plusieurs clients).
const groupKey = e => e.alescom ? 'alescom' : 'c:' + e.clientId;
function groupInfo(key) {
  const rates = state.settings.rates;
  if (key === 'alescom') return { key, label: 'alescom', billitName: state.settings.alescomBillit || 'alescom', rates: rates.partner, partner: true };
  const c = client(key.slice(2));
  return { key, label: c.name, billitName: c.billitName || c.name, rates: rates.client, partner: false };
}
function totals(entries, r) {
  let h = 0, m = 0, dep = 0;
  for (const e of entries) { h += Math.floor(e.minutes / 60); m += e.minutes % 60; dep += e.onSite ? 1 : 0; }
  return { h, m, dep, amount: h * r.h + m * r.m + dep * r.dep, minutes: entries.reduce((s, e) => s + e.minutes, 0) };
}
const entryAmount = e => totals([e], e.alescom ? state.settings.rates.partner : state.settings.rates.client).amount;
const xrow = e => ({ ...e, client: client(e.clientId).name });

const prestationsName = b => `Prestations - ${b.label} ${b.date}.xlsx`;
const prestationsFile = b => Xlsx.prestations(sortEntries(live().filter(e => e.billId === b.id)).map(xrow), b.key === 'alescom');

function recapBillit(b) {
  const info = groupInfo(b.key), r = info.rates;
  const s = sortEntries(live().filter(e => e.billId === b.id)), t = totals(s, r);
  const lines = [`Client Billit : ${info.billitName}`, `Objet : Intervention du ${frDate(b.from)} au ${frDate(b.to)}`, ''];
  if (t.h) lines.push(`${r.hLabel} : ${t.h} × ${eur(r.h)} = ${eur(t.h * r.h)}`);
  if (t.m) lines.push(`${r.mLabel} : ${t.m} × ${eur(r.m)} = ${eur(t.m * r.m)}`);
  if (t.dep) lines.push(`${r.depLabel} : ${t.dep} × ${eur(r.dep)} = ${eur(t.dep * r.dep)}`);
  const inst = s.filter(e => e.install);
  if (inst.length) lines.push('', 'Installations à ajouter :', ...inst.map(e => `- ${frDate(e.date)} ${e.install}`));
  lines.push('', `Total HTVA (hors installations) : ${eur(t.amount)}`, `Pièce jointe : ${prestationsName(b)}`);
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
  const before = JSON.stringify([state.entries, state.clients, state.bills, state.timer, state.settings]);
  const mergeArr = (a = [], b = []) => {
    const m = new Map(a.map(x => [x.id, x]));
    for (const x of b) { const y = m.get(x.id); if (!y || (x.updatedAt || 0) > (y.updatedAt || 0)) m.set(x.id, x); }
    return [...m.values()];
  };
  state.entries = mergeArr(state.entries, remote.entries);
  state.clients = mergeArr(state.clients, remote.clients);
  state.bills = mergeArr(state.bills, remote.bills);
  if ((remote.timerAt || 0) > (state.timerAt || 0)) { state.timer = remote.timer; state.timerAt = remote.timerAt; }
  if (remote.settings && (remote.settingsAt || 0) > (state.settingsAt || 0)) { state.settings = remote.settings; state.settingsAt = remote.settingsAt; }
  return before !== JSON.stringify([state.entries, state.clients, state.bills, state.timer, state.settings]);
}

// Régénère les fichiers Excel seulement si leur contenu a changé.
async function publishExcel() {
  const open = sortEntries(live().filter(e => !e.billId)).map(xrow);
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
    let changed = false;
    for (let i = 0; i < 3; i++) {
      const cur = await OneDrive.get(P.data());
      if (cur) changed = mergeRemote(JSON.parse(cur.data)) || changed;
      changed = mergeRemoteClients() || changed;
      const payload = JSON.stringify({ version: 3, entries: state.entries, clients: state.clients, bills: state.bills,
        timer: state.timer, timerAt: state.timerAt, settings: state.settings, settingsAt: state.settingsAt });
      const r = await OneDrive.put(P.data(), payload, { etag: cur?.etag, type: 'application/json' });
      if (r.status === 412) continue;                       // modifié par un autre appareil : on refusionne
      if (!r.ok) throw new Error(`OneDrive (${r.status}) : écriture des données impossible.`);
      break;
    }
    await publishExcel();
    state.lastSync = Date.now();
    await DB.save(state);
    setSync('ok');
    if (changed && !$('#dlg').open) render();
  } catch (e) {
    if (e instanceof OneDrive.NeedLogin) setSync('login', 'Session expirée : reconnecte-toi dans Réglages.');
    else setSync('err', e.message);
  } finally {
    syncing = false;
    if (syncAgain) { syncAgain = false; scheduleSync(); }
  }
}

// ---------- rendu ----------
const TITLES = { today: "Aujourd'hui", list: 'Interventions', billing: 'Facturer', archive: 'Archives', settings: 'Réglages' };
const VIEWS = () => ({ today: viewToday, list: viewList, billing: viewBilling, archive: viewArchive, settings: viewSettings });
function render() {
  clearInterval(tick);
  $('#viewTitle').textContent = TITLES[tab];
  $$('.tabs button').forEach(b => b.classList.toggle('active', b.dataset.tab === tab));
  VIEWS()[tab]();
}

function entryLi(e, { showDate = false, showBill = true } = {}) {
  const b = e.billId && bill(e.billId);
  return `<li class="tap" data-edit="${e.id}">
    <div class="grow">
      <div><strong>${esc(client(e.clientId).name)}</strong> ${e.onSite ? '<span class="pill">sur site</span>' : '<span class="pill">à distance</span>'}
        ${e.alescom ? '<span class="pill warn">alescom</span>' : ''}
        ${showBill && b ? `<span class="pill ok">facturé ${frDate(b.date)}${b.ref ? ' · ' + esc(b.ref) : ''}</span>` : ''}</div>
      <div class="small muted">${showDate ? frDate(e.date) + ' · ' : ''}${e.forfait ? 'forfait' : `${e.start}–${e.end}`}${e.pauseMin ? ` · pause ${fmtDur(e.pauseMin)}` : ''}${e.install ? ' · ' + esc(e.install) : ''}</div>
      ${e.description ? `<div class="small">${esc(e.description)}</div>` : ''}
    </div>
    <div class="dur">${fmtDur(e.minutes)}</div></li>`;
}
function bindEdits(root = document) {
  $$('[data-edit]', root).forEach(li => li.onclick = () => openEntry(state.entries.find(e => e.id === li.dataset.edit)));
}
function clientOptions(sel) {
  return activeClients().map(c => `<option value="${c.id}" ${c.id === sel ? 'selected' : ''}>${esc(c.name)}</option>`).join('');
}
const flags = (site, ales, p) => `
  <div class="row">
    <label class="check grow"><input type="checkbox" id="${p}Site" ${site ? 'checked' : ''}> Sur site</label>
    <label class="check grow"><input type="checkbox" id="${p}Ales" ${ales ? 'checked' : ''}> Commanditaire alescom</label>
  </div>`;

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
function viewToday() {
  const t = state.timer, d = today();
  const todays = sortEntries(live().filter(e => e.date === d));
  const now = new Date(), monday = new Date(now); monday.setDate(now.getDate() - ((now.getDay() + 6) % 7));
  const week = live().filter(e => e.date >= isoDate(monday) && e.date <= d);
  const sum = es => es.reduce((s, e) => s + e.minutes, 0);
  const open = live().filter(e => !e.billId);
  const lastC = state.clients.find(c => c.id === state.lastClientId && !c.deleted) || activeClients()[0];

  $('#view').innerHTML = `
    <div class="card timer">
      ${t ? `
        <div><strong>${esc(client(t.clientId).name)}</strong> ${t.pausedAt ? '<span class="pill warn">en pause</span>' : ''}</div>
        <div class="clock ${t.pausedAt ? 'muted' : ''}" id="clock">0:00:00</div>
        ${t.pausedAt ? '<div class="small">Pause en cours : <strong id="pauseClock">0:00:00</strong></div>' : ''}
        <div class="grid2" style="text-align:left">
          <div><label>Début</label><input type="time" id="tStartAt" value="${hhmm(new Date(t.start))}"></div>
          <div><label>Pause déduite (H:MM)</label><input id="tPauseMin" value="${fmtDur(Math.round(timerPausedMs(t) / 60000))}"></div>
        </div>
        <input id="tDesc" placeholder="Description (optionnel)" value="${esc(t.description || '')}" style="margin-top:10px">
        ${flags(t.onSite, t.alescom, 't')}
        <div class="row"><button class="btn" id="tCancel">Annuler</button>
          <button class="btn grow" id="tPause">${t.pausedAt ? '▶ Fin de pause' : '❚❚ Pause'}</button>
          <button class="btn primary grow" id="tStop">■ Arrêter</button></div>`
      : lastC ? `
        <select id="tClient">${clientOptions(lastC.id)}</select>
        <input id="tDesc0" placeholder="Description de l'intervention" style="margin-top:10px">
        ${flags(false, lastC.defaultAlescom, 't')}
        <div class="row" style="justify-content:center;margin-bottom:10px">
          <span class="small muted">Début</span>
          <input type="time" id="tAt" style="width:auto">
          <span class="small muted">(vide = maintenant)</span>
        </div>
        <button class="btn primary big" id="tStart">▶ Démarrer le chrono</button>`
      : `<p class="muted">Commence par ajouter un client dans Réglages.</p>`}
    </div>
    <button class="btn big" id="addManual">+ Ajouter une intervention</button>
    <div class="grid2" style="margin-top:10px">
      <button class="btn" id="addMaint">🛠 Maintenance serveur</button>
      <button class="btn" id="addUpd">⟳ Mise à jour serveur</button>
    </div>
    <div class="kpis" style="margin-top:12px">
      <div class="kpi"><div class="v">${fmtDur(sum(todays))}</div><div class="l">aujourd'hui</div></div>
      <div class="kpi"><div class="v">${fmtDur(sum(week))}</div><div class="l">cette semaine</div></div>
      <div class="kpi tap" id="goBill"><div class="v">${eur(open.reduce((s, e) => s + entryAmount(e), 0))}</div><div class="l">à facturer (${open.length} interv.) →</div></div>
    </div>
    <div class="card"><h2>Interventions du jour</h2>
      ${todays.length ? `<ul class="list">${todays.map(e => entryLi(e)).join('')}</ul>` : '<div class="empty">Rien pour l\'instant.</div>'}
    </div>`;

  $('#addManual').onclick = () => openEntry(null);
  $('#addMaint').onclick = () => openForfait(null, 'maint');
  $('#addUpd').onclick = () => openForfait(null, 'update');
  $('#goBill').onclick = () => { tab = 'billing'; render(); };
  bindEdits();

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
    $('#tAles').onchange = ev => { t.alescom = ev.target.checked; save(); };
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
        onSite: t.onSite, alescom: t.alescom
      }, true);
    };
  } else if ($('#tStart')) {
    $('#tClient').onchange = ev => { $('#tAles').checked = !!client(ev.target.value).defaultAlescom; };
    $('#tStart').onclick = async () => {
      const at = $('#tAt').value;
      const start = at ? startFromTime(at) : new Date();
      state.lastClientId = $('#tClient').value;
      setTimer({ clientId: state.lastClientId, start: start.toISOString(), description: $('#tDesc0').value.trim(),
        onSite: $('#tSite').checked, alescom: $('#tAles').checked, pausedMs: 0, pausedAt: null });
      await persist();
    };
  }
}

// ---- dialogue intervention ----
function openEntry(entry, preset = {}, fromTimer = false) {
  if (entry?.forfait) return openForfait(entry);
  const c0 = state.clients.find(c => c.id === state.lastClientId && !c.deleted) || activeClients()[0];
  const e = entry || {
    clientId: c0?.id, date: today(), start: hhmm(new Date(Date.now() - 3600e3)), end: hhmm(new Date()), minutes: 60, pauseMin: 0,
    description: '', install: '', onSite: false, alescom: !!c0?.defaultAlescom, ...preset
  };
  if (!e.clientId) return alert("Ajoute d'abord un client dans Réglages.");
  const b = e.billId && bill(e.billId);
  const dlg = $('#dlg'), f = $('#dlgForm');
  f.innerHTML = `
    <h2>${entry ? 'Modifier' : fromTimer ? 'Fin du chrono' : 'Nouvelle'} intervention</h2>
    ${b ? `<p class="small"><span class="pill ok">facturée le ${frDate(b.date)}${b.ref ? ' · ' + esc(b.ref) : ''}</span>
      ${b.key !== 'import' ? '<button class="link" id="fBill">Voir la facture</button>' : ''}</p>` : ''}
    <label>Client</label>
    <select id="fClient">${clientOptions(e.clientId)}<option value="__new">+ Nouveau client…</option></select>
    <div class="grid2">
      <div><label>Date</label><input type="date" id="fDate" value="${e.date}" required></div>
      <div><label>Durée facturée (H:MM)</label><input id="fDur" value="${fmtDur(e.minutes)}"></div>
    </div>
    <div class="grid3">
      <div><label>Début</label><input type="time" id="fStart" value="${e.start}"></div>
      <div><label>Fin</label><input type="time" id="fEnd" value="${e.end}"></div>
      <div><label>Pause</label><input id="fPause" value="${e.pauseMin ? fmtDur(e.pauseMin) : ''}" placeholder="0:00"></div>
    </div>
    <label>Description</label><textarea id="fDesc">${esc(e.description)}</textarea>
    ${flags(e.onSite, e.alescom, 'f')}
    <label>Installation / réinstallation (optionnel)</label><input id="fInst" value="${esc(e.install)}" placeholder="ex. Réinstallation PC">
    <div class="actions">
      ${entry && !b ? '<button class="btn danger" id="fDel" style="margin-right:auto">Supprimer</button>' : ''}
      <button class="btn" value="cancel">${fromTimer ? 'Reprendre' : 'Annuler'}</button>
      <button class="btn primary" id="fSave">Enregistrer</button>
    </div>`;

  if ($('#fBill')) $('#fBill').onclick = () => { dlg.close(); showBill(b); };
  const pause = () => parseDur($('#fPause').value) || 0;
  const upDur = () => $('#fDur').value = fmtDur(Math.max(0, diffMin($('#fStart').value, $('#fEnd').value) - pause()));
  $('#fStart').oninput = upDur; $('#fEnd').oninput = upDur; $('#fPause').onchange = upDur;
  $('#fDur').onchange = () => {
    const m = parseDur($('#fDur').value);
    if (!isNaN(m)) { $('#fDur').value = fmtDur(m); $('#fEnd').value = fromMin(toMin($('#fStart').value) + m + pause()); }
  };
  $('#fClient').onchange = ev => {
    if (ev.target.value === '__new') {
      const name = (prompt('Nom du nouveau client :') || '').trim();
      let c = name && state.clients.find(x => !x.deleted && x.name.toLowerCase() === name.toLowerCase());
      if (name && !c) state.clients.push(c = touch({ id: uid('c'), name, billitName: '', defaultOnSite: false, defaultAlescom: false }));
      ev.target.innerHTML = clientOptions(c?.id || e.clientId) + '<option value="__new">+ Nouveau client…</option>';
    } else if (!entry) $('#fAles').checked = !!client(ev.target.value).defaultAlescom;
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
      clientId, date: $('#fDate').value, start: $('#fStart').value, end: $('#fEnd').value, minutes, pauseMin,
      description: $('#fDesc').value.trim(), onSite: $('#fSite').checked, alescom: $('#fAles').checked, install: $('#fInst').value.trim()
    };
    if (b && !confirm('Cette intervention est déjà facturée. Modifier quand même l\'archive ?')) return;
    if (entry) Object.assign(touch(entry), data);
    else state.entries.push(touch({ id: uid('e'), billId: null, ...data }));
    state.lastClientId = clientId;
    if (fromTimer) setTimer(null);
    dlg.close(); await persist(); toast('Enregistré');
  };
  dlg.showModal();
}

// ---- forfaits serveur : une date et une durée, sans heures de début / fin ----
const FORFAITS = { maint: 'Maintenance serveur', update: 'Mise à jour serveur' };
const serverCount = c => c?.server ? Math.max(1, +c.servers || 1) : 0;
// Mise à jour serveur : 15 min pour 1 serveur, 10 min par serveur au-delà.
const updateMinutes = c => { const n = serverCount(c) || 1; return n === 1 ? 15 : 10 * n; };

function openForfait(entry, kind) {
  kind = entry?.forfait || kind;
  const label = FORFAITS[kind];
  const clients = activeClients().sort((a, b) => (serverCount(b) > 0) - (serverCount(a) > 0) || a.name.localeCompare(b.name));
  const c0 = entry ? client(entry.clientId)
    : clients.find(c => c.id === state.lastClientId && (kind === 'maint' || c.server)) || clients.find(c => c.server) || clients[0];
  if (!c0) return alert("Ajoute d'abord un client dans Réglages.");
  const e = entry || { clientId: c0.id, date: today(), minutes: kind === 'update' ? updateMinutes(c0) : 0,
    description: label, onSite: false, alescom: !!c0.defaultAlescom };
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
    <label>Description</label><textarea id="kDesc">${esc(e.description)}</textarea>
    ${flags(e.onSite, e.alescom, 'k')}
    <div class="actions">
      ${entry && !b ? '<button class="btn danger" id="kDel" style="margin-right:auto">Supprimer</button>' : ''}
      <button class="btn" value="cancel">Annuler</button>
      <button class="btn primary" id="kSave">Enregistrer</button>
    </div>`;
  $('#kClient').onchange = ev => {
    const c = client(ev.target.value);
    $('#kHint').textContent = hint(c);
    if (kind === 'update') $('#kDur').value = fmtDur(updateMinutes(c));
    if (!entry) $('#kAles').checked = !!c.defaultAlescom;
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
      description: $('#kDesc').value.trim(), onSite: $('#kSite').checked, alescom: $('#kAles').checked, install: '', forfait: kind };
    if (entry) Object.assign(touch(entry), data);
    else state.entries.push(touch({ id: uid('e'), billId: null, ...data }));
    state.lastClientId = data.clientId;
    dlg.close(); await persist(); toast('Enregistré');
  };
  dlg.showModal();
  if (kind === 'maint' && !entry) $('#kDur').focus();
}

// ---- Interventions ----
let listMonth = today().slice(0, 7), listClient = '', listStatus = '';
function viewList() {
  const es = sortEntries(live().filter(e => e.date.startsWith(listMonth) && (!listClient || e.clientId === listClient)
    && (!listStatus || (listStatus === 'open') === !e.billId))).reverse();
  const days = [...new Set(es.map(e => e.date))];
  $('#view').innerHTML = `
    <div class="card"><div class="grid3">
      <div><label>Mois</label><input type="month" id="hMonth" value="${listMonth}"></div>
      <div><label>Client</label><select id="hClient"><option value="">Tous</option>${clientOptions(listClient)}</select></div>
      <div><label>Statut</label><select id="hStatus"><option value="">Toutes</option>
        <option value="open" ${listStatus === 'open' ? 'selected' : ''}>À facturer</option>
        <option value="billed" ${listStatus === 'billed' ? 'selected' : ''}>Facturées</option></select></div>
    </div>
    <p class="small muted" style="margin:10px 0 0">${es.length} intervention(s) · <strong>${fmtDur(es.reduce((s, e) => s + e.minutes, 0))}</strong></p></div>
    ${days.length ? days.map(d => `<div class="day">${longDate(d)}</div>
      <div class="card" style="padding:4px 14px"><ul class="list">${es.filter(e => e.date === d).map(e => entryLi(e)).join('')}</ul></div>`).join('')
      : '<div class="empty">Aucune intervention.</div>'}`;
  $('#hMonth').onchange = e => { listMonth = e.target.value || today().slice(0, 7); render(); };
  $('#hClient').onchange = e => { listClient = e.target.value; render(); };
  $('#hStatus').onchange = e => { listStatus = e.target.value; render(); };
  bindEdits();
}

// ---- Facturer ----
let billingTarget = null, billUntil = today(), excluded = new Set();
function viewBilling() {
  const open = live().filter(e => !e.billId);
  const keys = [...new Set(open.map(groupKey))];
  if (billingTarget && !keys.includes(billingTarget)) billingTarget = null;
  if (billingTarget) return viewBillingTarget(billingTarget, open.filter(e => groupKey(e) === billingTarget));

  const groups = keys.map(k => {
    const info = groupInfo(k), es = sortEntries(open.filter(e => groupKey(e) === k));
    return { info, es, t: totals(es, info.rates) };
  }).sort((a, b) => b.t.amount - a.t.amount);

  $('#view').innerHTML = `
    <div class="kpis"><div class="kpi"><div class="v">${eur(groups.reduce((s, g) => s + g.t.amount, 0))}</div><div class="l">à facturer (HTVA estimé)</div></div>
      <div class="kpi"><div class="v">${open.length}</div><div class="l">interventions en attente</div></div></div>
    <div class="card"><h2>En attente de facturation</h2>
      ${groups.length ? `<ul class="list">${groups.map(g => `
        <li class="tap" data-target="${g.info.key}"><div class="grow"><strong>${esc(g.info.label)}</strong>
          ${g.info.partner ? '<span class="pill warn">partenaire</span>' : ''}
          <div class="small muted">${g.es.length} interv. · ${fmtDur(g.t.minutes)} · du ${frDate(g.es[0].date)} au ${frDate(g.es[g.es.length - 1].date)}</div></div>
          <div class="dur">${eur(g.t.amount)}</div></li>`).join('')}</ul>`
      : '<div class="empty">Tout est facturé 🎉</div>'}
    </div>`;
  $$('[data-target]').forEach(li => li.onclick = () => { billingTarget = li.dataset.target; billUntil = today(); excluded = new Set(); render(); });
}

function totalsTable(t, r) {
  return `<table class="tot">
    <tr><td>${esc(r.hLabel)}<br><span class="small muted">${t.h} × ${eur(r.h)}</span></td><td>${eur(t.h * r.h)}</td></tr>
    <tr><td>${esc(r.mLabel)}<br><span class="small muted">${t.m} × ${eur(r.m)}</span></td><td>${eur(t.m * r.m)}</td></tr>
    <tr><td>${esc(r.depLabel)}<br><span class="small muted">${t.dep} × ${eur(r.dep)}</span></td><td>${eur(t.dep * r.dep)}</td></tr>
    <tr class="sum"><td>Total HTVA</td><td>${eur(t.amount)}</td></tr>
    <tr><td class="muted">TVAC (21 %)</td><td class="muted">${eur(t.amount * 1.21)}</td></tr></table>`;
}

function viewBillingTarget(key, all) {
  const info = groupInfo(key), r = info.rates;
  const es = sortEntries(all.filter(e => e.date <= billUntil));
  const sel = es.filter(e => !excluded.has(e.id));
  const t = totals(sel, r);

  $('#view').innerHTML = `
    <button class="link" id="back">← Retour</button>
    <div class="card" style="margin-top:8px"><h2>${esc(info.label)} <span class="small muted">→ Billit : ${esc(info.billitName)}</span></h2>
      <label>Inclure jusqu'au</label><input type="date" id="until" value="${billUntil}">
      <div style="margin-top:12px">${totalsTable(t, r)}</div>
      ${sel.some(e => e.install) ? '<p class="small"><span class="pill warn">Contient des installations : article à ajouter dans Billit.</span></p>' : ''}
      <button class="btn primary big" id="doBill" style="margin-top:12px" ${sel.length ? '' : 'disabled'}>Facturer ${sel.length} intervention(s)</button>
      <p class="small muted">Crée « Prestations - ${esc(info.label)} ${today()}.xlsx » dans OneDrive/A facturer et déplace ces interventions vers l'archive.</p>
    </div>
    <div class="card"><h2>Interventions (${sel.length}/${es.length})</h2>
      <ul class="list">${es.map(e => `<li><input type="checkbox" data-inc="${e.id}" ${excluded.has(e.id) ? '' : 'checked'}>
        <div class="grow tap" data-edit="${e.id}"><div><strong>${frDate(e.date)}</strong> ${e.forfait ? 'forfait' : `${e.start}–${e.end}`}
          ${info.partner ? '· ' + esc(client(e.clientId).name) : ''} ${e.onSite ? '<span class="pill">sur site</span>' : ''}</div>
          <div class="small muted">${esc(e.description)}${e.install ? ' · ' + esc(e.install) : ''}</div></div>
        <div class="dur">${fmtDur(e.minutes)}</div></li>`).join('')}</ul></div>`;

  $('#back').onclick = () => { billingTarget = null; render(); };
  $('#until').onchange = ev => { billUntil = ev.target.value || today(); render(); };
  $$('[data-inc]').forEach(c => c.onchange = () => { c.checked ? excluded.delete(c.dataset.inc) : excluded.add(c.dataset.inc); render(); });
  bindEdits();
  $('#doBill').onclick = async () => {
    const ref = prompt(`Facturer ${sel.length} intervention(s) de ${info.label} pour ${eur(t.amount)} HTVA.\n\nN° de facture Billit (optionnel, modifiable plus tard dans Archives) :`, '');
    if (ref === null) return;
    const b = touch({ id: uid('b'), key, label: info.label, date: today(), ref: ref.trim(), entryIds: sel.map(e => e.id),
      from: sel[0].date, to: sel[sel.length - 1].date, amount: t.amount });
    state.bills.push(b);
    sel.forEach(e => { e.billId = b.id; touch(e); });
    billingTarget = null;
    await persist();
    showBill(b, true);
  };
}

// Dialogue d'une facture : fichier Prestations, récap Billit, référence, annulation.
async function showBill(b, justCreated = false) {
  const info = groupInfo(b.key);
  const es = sortEntries(live().filter(e => e.billId === b.id));
  const t = totals(es, info.rates);
  const dlg = $('#dlg'), f = $('#dlgForm');
  f.innerHTML = `<h2>${justCreated ? 'Facturé : ' : ''}${esc(b.label)}</h2>
    <p class="small muted">Facturé le ${frDate(b.date)} · ${es.length} interv. · du ${frDate(b.from)} au ${frDate(b.to)}</p>
    <p class="small" id="bFile">${justCreated ? 'Envoi du fichier Prestations vers OneDrive…' : ''}</p>
    <label>N° de facture Billit</label><input id="bRef" value="${esc(b.ref)}">
    ${totalsTable(t, info.rates)}
    <pre class="recap">${esc(recapBillit(b))}</pre>
    <ul class="list small" style="margin-top:8px">${es.map(e => entryLi(e, { showDate: true, showBill: false })).join('')}</ul>
    <div class="actions">
      <button class="btn danger" id="bUndo" style="margin-right:auto">Annuler la facturation</button>
      <button class="btn" id="bDl">⬇ Prestations</button>
      <button class="btn" id="bCopy">Copier récap</button>
      <button class="btn primary" id="bOk">Fermer</button>
    </div>`;
  const saveRef = async () => { const v = $('#bRef').value.trim(); if (v !== (b.ref || '')) { b.ref = v; touch(b); await persist(); } };
  $('#bOk').onclick = async ev => { ev.preventDefault(); await saveRef(); dlg.close(); };
  $('#bCopy').onclick = () => copy(recapBillit(b));
  $('#bDl').onclick = async () => { try { download(await prestationsFile(b), prestationsName(b)); } catch (e) { alert(e.message); } };
  $('#bUndo').onclick = async () => {
    if (!confirm('Remettre ces interventions dans « à facturer » ? (le fichier Prestations déjà créé n\'est pas supprimé)')) return;
    es.forEach(e => { e.billId = null; touch(e); });
    b.deleted = true; touch(b);
    dlg.close(); await persist(); toast('Facturation annulée');
  };
  $$('[data-edit]', f).forEach(li => li.onclick = null);
  dlg.showModal();

  if (justCreated) {
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

// ---- Archives : anciennes interventions (déjà facturées) ----
let archQuery = '', archYear = today().slice(0, 4), archClient = '';
const ARCH_MAX = 300;
function viewArchive() {
  const q = archQuery.toLowerCase();
  const billed = live().filter(e => e.billId && bill(e.billId));
  const years = [...new Set([today().slice(0, 4), ...billed.map(e => e.date.slice(0, 4))])].sort().reverse();
  const text = e => { const b = bill(e.billId); return [client(e.clientId).name, e.description, e.install, b.ref, b.label].join(' ').toLowerCase(); };
  const es = sortEntries(billed.filter(e => (!archYear || e.date.startsWith(archYear)) && (!archClient || e.clientId === archClient)
    && (!q || text(e).includes(q)))).reverse();
  const shown = es.slice(0, ARCH_MAX);
  const days = [...new Set(shown.map(e => e.date))];

  $('#view').innerHTML = `
    <div class="card">
      <label style="margin-top:0">Rechercher</label><input type="search" id="aQ" value="${esc(archQuery)}" placeholder="description, client, n° facture…">
      <div class="grid2">
        <div><label>Année</label><select id="aYear"><option value="">Toutes</option>${years.map(y => `<option ${y === archYear ? 'selected' : ''}>${y}</option>`).join('')}</select></div>
        <div><label>Client</label><select id="aClient"><option value="">Tous</option>${clientOptions(archClient)}</select></div>
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
  bindEdits();
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
    <div class="card"><div class="row"><h2 class="grow" style="margin:0">Clients</h2><button class="btn" id="addClient">+ Ajouter</button></div>
      <ul class="list" style="margin-top:8px">${state.clients.filter(c => !c.deleted).sort((a, b) => a.name.localeCompare(b.name)).map(c => `
        <li class="tap" data-client="${c.id}"><div class="grow"><strong>${esc(c.name)}</strong> ${c.archived ? '<span class="pill">archivé</span>' : ''}
          <div class="small muted">Billit : ${esc(c.billitName || '—')}${c.server ? ` · ${serverCount(c)} serveur(s)` : ''}${c.defaultAlescom ? ' · alescom' : ''}</div></div></li>`).join('')}</ul>
      <label>Nom Billit du commanditaire alescom</label><input id="alesBillit" value="${esc(s.alescomBillit)}">
    </div>

    <div class="card">${rateBlock('client', 'Tarifs clients')}<div style="height:12px"></div>${rateBlock('partner', 'Tarifs alescom (partenaire)')}</div>

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
    </div>`;

  if (installPrompt) $('#install').onclick = async () => { installPrompt.prompt(); await installPrompt.userChoice; installPrompt = null; render(); };
  $('#addClient').onclick = () => editClient(null);
  $$('[data-client]').forEach(li => li.onclick = () => editClient(state.clients.find(c => c.id === li.dataset.client)));
  const saveSettings = async () => { state.settingsAt = Date.now(); await DB.save(state); scheduleSync(); toast('Enregistré'); };
  $('#alesBillit').onchange = ev => { s.alescomBillit = ev.target.value.trim(); saveSettings(); };
  $$('[data-rate]').forEach(inp => inp.onchange = () => {
    const [k, fld] = inp.dataset.rate.split('.'); const v = parseFloat(inp.value);
    if (!isNaN(v)) { r[k][fld] = v; saveSettings(); }
  });

  const saveOd = async () => {
    Object.assign(od, { clientId: $('#odClient').value.trim(), tenant: $('#odTenant').value.trim() || 'organizations',
      folder: $('#odFolder').value.trim().replace(/^\/+|\/+$/g, '') || 'Compta/TimeSheet' });
    state.published = {};
    await DB.save(state); render(); sync();
  };
  ['#odClient', '#odTenant', '#odFolder'].forEach(sel => $(sel).onchange = saveOd);
  if ($('#odIn')) $('#odIn').onclick = () => OneDrive.login().catch(e => alert(e.message));
  if ($('#odSync')) $('#odSync').onclick = () => { state.published = {}; sync(); };
  if ($('#odOut')) $('#odOut').onclick = () => { if (confirm('Se déconnecter de OneDrive ?')) { OneDrive.logout(); sync(); render(); } };

  $('#backup').onclick = () => download(new Blob([JSON.stringify({ entries: state.entries, clients: state.clients, bills: state.bills }, null, 1)],
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
  navigator.storage?.persisted?.().then(p => {
    const el = $('#persistInfo');
    if (el) el.textContent = p ? 'Cache local persistant.' : 'Installe l\'app pour que le navigateur ne vide pas le cache local.';
  });
}

function editClient(c) {
  const dlg = $('#dlg'), f = $('#dlgForm');
  const used = c && live().some(e => e.clientId === c.id);
  f.innerHTML = `<h2>${c ? 'Modifier' : 'Nouveau'} client</h2>
    <label>Nom</label><input id="cName" value="${esc(c?.name)}" placeholder="ex. Vedi">
    <label>Nom du client dans Billit</label><input id="cBillit" value="${esc(c?.billitName)}" placeholder="ex. V.T.S. SA">
    <label class="check"><input type="checkbox" id="cAles" ${c?.defaultAlescom ? 'checked' : ''}> Commanditaire alescom par défaut</label>
    <div class="row">
      <label class="check grow"><input type="checkbox" id="cServer" ${c?.server ? 'checked' : ''}> Serveur(s)</label>
      <div id="cServersBox" class="${c?.server ? '' : 'hidden'}" style="width:140px"><label style="margin-top:0">Nombre de serveurs</label>
        <input type="number" id="cServers" min="1" step="1" value="${c?.servers || 1}"></div>
    </div>
    ${c ? `<label class="check"><input type="checkbox" id="cArch" ${c.archived ? 'checked' : ''}> Archivé (masqué des listes)</label>` : ''}
    <div class="actions">
      ${c && !used ? '<button class="btn danger" id="cDel" style="margin-right:auto">Supprimer</button>' : ''}
      <button class="btn" value="cancel">Annuler</button><button class="btn primary" id="cSave">Enregistrer</button></div>`;
  $('#cServer').onchange = ev => $('#cServersBox').classList.toggle('hidden', !ev.target.checked);
  if ($('#cDel')) $('#cDel').onclick = async () => { c.deleted = true; touch(c); dlg.close(); await persist(); };
  $('#cSave').onclick = async ev => {
    ev.preventDefault();
    const name = $('#cName').value.trim().replace(/\s+/g, ' ');
    if (!name) return alert('Nom requis.');
    if (/\b(remote|distance)\b/i.test(name)) return alert('Retire « remote » / « distance » du nom : utilise la case « Sur site ».');
    if (state.clients.some(x => x !== c && !x.deleted && x.name.toLowerCase() === name.toLowerCase())) return alert('Ce client existe déjà.');
    const server = $('#cServer').checked, servers = Math.max(1, parseInt($('#cServers').value, 10) || 1);
    const data = { name, billitName: $('#cBillit').value.trim(), defaultAlescom: $('#cAles').checked, archived: !!$('#cArch')?.checked,
      server, servers: server ? servers : 0 };
    if (c) Object.assign(touch(c), data); else state.clients.push(touch({ id: uid('c'), ...data }));
    dlg.close(); await persist();
  };
  dlg.showModal();
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
      state.clients.push(touch({ id: uid('c'), name: n, billitName: '', defaultOnSite: first.onSite, defaultAlescom: first.alescom }));
    }
    let b = null;
    if (askBilled && $('#iBilled').checked) state.bills.push(b = touch({ id: uid('b'), key: 'import', label: 'Import historique', date: today(), ref: '', entryIds: [], from: '', to: '', amount: 0 }));
    let n = 0;
    for (const r of rows) {
      const c = find(r.client);
      if (existing.has(dupKey(c.id, r))) continue;
      const e = touch({ id: uid('e'), clientId: c.id, date: r.date, start: r.start, end: r.end, minutes: r.minutes, pauseMin: 0,
        description: r.description, onSite: r.onSite, alescom: r.alescom, install: r.install || '', billId: b?.id || null });
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

// ---------- démarrage ----------
// Dans les dialogues, seul le bouton principal est "submit" : Entrée enregistre, jamais supprimer/annuler.
new MutationObserver(() => $$('#dlgForm button').forEach(b => {
  if (b.classList.contains('primary')) return;
  b.type = 'button';
  if (b.value === 'cancel') b.onclick = () => $('#dlg').close();
})).observe($('#dlgForm'), { childList: true });
$('#dlgForm').addEventListener('submit', ev => { ev.preventDefault(); $('#dlgForm .btn.primary')?.click(); });

$$('.tabs button').forEach(b => b.onclick = () => { tab = b.dataset.tab; render(); });
window.addEventListener('beforeinstallprompt', e => { e.preventDefault(); installPrompt = e; if (tab === 'settings') render(); });
window.addEventListener('online', sync);
window.addEventListener('offline', () => setSync('offline', 'Hors ligne : synchro au retour du réseau.'));
document.addEventListener('visibilitychange', () => { if (document.visibilityState === 'visible') sync(); });

(async () => {
  state = await DB.load();
  if (mergeRemoteClients()) await DB.save(state);
  OneDrive.configure(() => state.onedrive);
  try { if (await OneDrive.handleRedirect()) { tab = 'settings'; toast('Connecté à OneDrive'); } }
  catch (e) { alert('Connexion OneDrive : ' + e.message); }
  render();
  sync();
  navigator.storage?.persist?.();
  if ('serviceWorker' in navigator && (location.protocol === 'https:' || location.hostname === 'localhost')) {
    navigator.serviceWorker.register('sw.js');
  }
})();
