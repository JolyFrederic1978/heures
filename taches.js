// Module Tâches (voir taches.md) : tâches « à la Todoist », liées à un client et à un technicien ; une tâche mène à une ou plusieurs interventions.
// - Données : TimeSheet/App/taches-data.json, avec sa propre synchro (comme Achats). heures-data.json : seulement entries[].taskId.
// - Technicien vide = « à prendre » : un technicien peut se l'attribuer si le client fait partie de ses clients visibles (Admin → Utilisateurs).
// - Tout technicien peut assigner à tout autre ; le créateur et le technicien assigné peuvent modifier, réassigner, terminer, supprimer.
// Utilise les outils de app.js ($, esc, uid, touch, state, client, me, canSee, openEntry…) : chargé avant lui, appelé après le démarrage.
const Taches = (() => {
  let T = null, busy = false, msg = '';
  let filter = 'today', fClient = '', fTech = '';

  // --- cache local (base IndexedDB à part) ---
  function idb(mode, fn) {
    return new Promise((resolve, reject) => {
      const req = indexedDB.open('heures-taches', 1);
      req.onupgradeneeded = () => req.result.createObjectStore('kv');
      req.onerror = () => reject(req.error);
      req.onsuccess = () => {
        const t = req.result.transaction('kv', mode);
        const r = fn(t.objectStore('kv'));
        t.oncomplete = () => resolve(r && r.result);
        t.onerror = () => reject(t.error);
      };
    });
  }
  const empty = () => ({ tasks: [], lastSync: 0 });
  async function load() { if (!T) T = { ...empty(), ...(await idb('readonly', s => s.get('data')) || {}) }; return T; }
  const save = () => idb('readwrite', s => s.put(T, 'data'));
  async function changed() { await save(); render(); scheduleSync(); }
  const dataPath = () => `${state.onedrive.folder}/App/taches-data.json`;

  // --- règles ---
  const alive = () => T.tasks.filter(t => !t.deleted);
  const task = id => T.tasks.find(t => t.id === id);
  const socOfUser = id => state.users.find(u => u.id === id)?.societeId;
  const isResp = () => (me().roles || []).includes('resp');
  // Visible : assignée à moi, créée par moi, à prendre sur un de mes clients ; responsable : tout ce qui touche sa société ; administrateur : tout.
  function visible(t) {
    const u = me();
    if (isAdmin() || t.assigneeId === u.id || t.createdBy === u.id) return true;
    if (!t.assigneeId && canTakeClient(t)) return true;
    return isResp() && [socOfUser(t.assigneeId), socOfUser(t.createdBy), t.clientId && clientSoc(client(t.clientId))].includes(u.societeId);
  }
  const canTakeClient = t => !t.clientId || canSee(client(t.clientId));
  const canEdit = t => isAdmin() || t.assigneeId === me().id || t.createdBy === me().id;
  const canTake = t => !t.assigneeId && t.status !== 'done' && canTakeClient(t);
  const isOpen = t => t.status !== 'done';
  const overdue = t => isOpen(t) && t.due && t.due < today();
  const dueToday = t => isOpen(t) && t.due === today();
  const sortKey = t => `${t.due || '9999-99-99'} ${t.time || '99:99'} ${t.priority || 4}`;
  const PRIO = { 1: 'P1 urgent', 2: 'P2 haute', 3: 'P3 moyenne', 4: 'P4 normale' };

  // Assignation : le nouveau technicien est prévenu dans son bandeau (newFor, effacé quand il ouvre la tâche).
  function assign(t, id) {
    if ((t.assigneeId || '') === (id || '')) return;
    t.assigneeId = id || null;
    if (id && id !== me().id) t.newFor = id; else delete t.newFor;
  }

  // --- synchronisation OneDrive (appelée par la synchro du pointage, app.js) ---
  function merge(remote) {
    const before = JSON.stringify(T.tasks);
    const x = new Map(T.tasks.map(o => [o.id, o]));
    for (const o of remote.tasks || []) { const y = x.get(o.id); if (!y || (o.updatedAt || 0) > (y.updatedAt || 0)) x.set(o.id, o); }
    T.tasks = [...x.values()];
    return before !== JSON.stringify(T.tasks);
  }
  async function sync() {
    await load();
    if (busy || !OneDrive.signedIn() || !navigator.onLine) return;
    busy = true;
    let redraw = false;
    try {
      let cur = await OneDrive.get(dataPath());
      if (cur) redraw = merge(JSON.parse(cur.data));
      for (let i = 0; i < 3; i++) {
        const body = JSON.stringify({ version: 1, tasks: T.tasks });
        if (cur && hash(cur.data) === hash(body)) break;            // rien de neuf à écrire
        const r = await OneDrive.put(dataPath(), body, { etag: cur?.etag, type: 'application/json' });
        if (r.ok) break;
        if (r.status !== 412) throw new Error(`OneDrive (${r.status}) : écriture de taches-data.json impossible.`);
        cur = await OneDrive.get(dataPath());                       // modifié par un autre appareil : on refusionne
        if (cur) redraw = merge(JSON.parse(cur.data)) || redraw;
      }
      T.lastSync = Date.now(); msg = '';
      await save();
    } catch (e) {
      msg = e instanceof OneDrive.NeedLogin ? '' : e.message;
      redraw = true;
    } finally { busy = false; }
    if (redraw && tab === 'taches' && !$('#dlg').open) render();
    if (tab === 'today' && $('#tkNotif')) todayBanner($('#tkNotif'));
  }

  // --- ajout rapide : « Réinstaller PC accueil Vedi demain 14h p1 @Geoffrey » ---
  // Reconnu : p1…p4, @technicien, #client (sinon le nom d'un client dans le texte), aujourd'hui / demain / après-demain / jour de la semaine /
  // JJ/MM[/AAAA] / « dans N jours », heure 14h, 14h30, 14:30. Le reste devient le titre ; tout est vérifié dans la fiche avant d'enregistrer.
  const DAYS = ['dimanche', 'lundi', 'mardi', 'mercredi', 'jeudi', 'vendredi', 'samedi'];
  const addDays = n => { const d = new Date(); d.setDate(d.getDate() + n); return isoDate(d); };
  function parseQuick(s) {
    let txt = ` ${s} `;
    const out = {};
    // fn renvoie false : le morceau reste dans le titre.
    const take = (re, fn) => { txt = txt.replace(re, (...m) => fn(m) === false ? m[0] : ' '); };
    take(/\s[pP]([1-4])(?=\s)/, m => { out.priority = +m[1]; });
    take(/\s@(\S+)/, m => {
      const u = liveUsers().find(x => x.active !== false && fold(x.name).startsWith(fold(m[1])));
      if (!u) return false; out.assigneeId = u.id;
    });
    take(/\s#(\S+)/, m => {
      const c = activeClients().filter(canSee).find(x => fold(x.name).replace(/\s+/g, '').startsWith(fold(m[1])));
      if (!c) return false; out.clientId = c.id;
    });
    take(/\s(\d{1,2})(?:h(\d{2})?|:(\d{2}))(?=\s)/i, m => {
      const h = +m[1], mi = +(m[2] || m[3] || 0);
      if (h > 23 || mi > 59) return false; out.time = `${pad(h)}:${pad(mi)}`;
    });
    take(/\saprès-demain(?=\s)/i, () => { out.due = addDays(2); });
    take(/\sdemain(?=\s)/i, () => { out.due = addDays(1); });
    take(/\s(aujourd'hui|auj)(?=\s)/i, () => { out.due = today(); });
    take(/\sdans (\d{1,3}) jours?(?=\s)/i, m => { out.due = addDays(+m[1]); });
    take(new RegExp(`\\s(${DAYS.join('|')})(?=\\s)`, 'i'), m => {
      const n = (DAYS.indexOf(m[1].toLowerCase()) - new Date().getDay() + 7) % 7 || 7; out.due = addDays(n);
    });
    take(/\s(\d{1,2})\/(\d{1,2})(?:\/(\d{2,4}))?(?=\s)/, m => {
      const d = +m[1], mo = +m[2];
      if (!d || d > 31 || !mo || mo > 12) return false;
      let y = m[3] ? +(m[3].length === 2 ? '20' + m[3] : m[3]) : new Date().getFullYear();
      let iso = `${y}-${pad(mo)}-${pad(d)}`;
      if (!m[3] && iso < today()) iso = `${y + 1}-${pad(mo)}-${pad(d)}`;     // date passée sans année : l'an prochain
      out.due = iso;
    });
    if (out.time && !out.due) out.due = today();
    out.title = txt.replace(/\s+/g, ' ').trim();
    // Pas de #client : nom d'un client dans le texte (le plus long d'abord), laissé dans le titre.
    if (!out.clientId) {
      const f = ` ${fold(out.title)} `;
      const c = activeClients().filter(canSee).sort((a, b) => b.name.length - a.name.length)
        .find(x => x.name.length > 2 && f.includes(` ${fold(x.name)} `));
      if (c) out.clientId = c.id;
    }
    return out;
  }

  // --- affichage ---
  const userName = id => id ? user(id).name : 'à prendre';
  const dueLabel = t => {
    if (!t.due) return '';
    const d = t.due === today() ? "aujourd'hui" : t.due === addDays(1) ? 'demain' : new Date(t.due + 'T00:00').toLocaleDateString('fr-BE', { weekday: 'short', day: '2-digit', month: '2-digit' });
    return d + (t.time ? ' ' + t.time : '');
  };
  function taskLi(t) {
    const cl = t.checklist || [], done = cl.filter(x => x.done).length;
    const meta = [
      t.clientId && esc(client(t.clientId).name),
      t.assigneeId ? `👤 ${esc(userName(t.assigneeId))}` : '<span class="pill warn">à prendre</span>',
      t.due && `<span class="${overdue(t) ? 'late' : dueToday(t) ? 'now' : ''}">📅 ${esc(dueLabel(t))}</span>`,
      cl.length && `☑ ${done}/${cl.length}`,
      t.comments?.length && `💬 ${t.comments.length}`,
      t.entryIds?.length && `🔗 ${t.entryIds.length} interv.`
    ].filter(Boolean).join(' · ');
    return `<li class="tk p${t.priority || 4} ${isOpen(t) ? '' : 'done'}" data-tk="${t.id}">
      <button class="tkdone" data-tkdone="${t.id}" title="${isOpen(t) ? 'Terminer' : 'Rouvrir'}" ${canEdit(t) ? '' : 'disabled'}>${isOpen(t) ? '' : '✓'}</button>
      <div class="grow tap" style="flex:1;min-width:0"><div><strong>${esc(t.title)}</strong>${t.newFor === me().id ? ' <span class="pill ok">nouvelle</span>' : ''}</div>
        <div class="small muted">${meta}</div></div></li>`;
  }

  function view() {
    if (!T) { $('#view').innerHTML = '<div class="empty">Chargement…</div>'; load().then(() => tab === 'taches' && render()); return; }
    const u = me(), all = alive().filter(visible);
    const FILTERS = {
      today: ["Aujourd'hui et en retard", t => isOpen(t) && (overdue(t) || dueToday(t))],
      upcoming: ['À venir', t => isOpen(t) && t.due > today()],
      free: ['À prendre', t => canTake(t)],
      mine: ['Assignées par moi', t => isOpen(t) && t.createdBy === u.id && t.assigneeId !== u.id],
      open: ['Toutes', t => isOpen(t)],
      done: ['Terminées', t => !isOpen(t)]
    };
    const list = all.filter(FILTERS[filter][1]).filter(t => (!fClient || t.clientId === fClient) && (!fTech || (fTech === '-' ? !t.assigneeId : t.assigneeId === fTech)));
    filter === 'done' ? list.sort((a, b) => (b.doneAt || 0) - (a.doneAt || 0)) : list.sort((a, b) => sortKey(a).localeCompare(sortKey(b)));
    const shown = filter === 'done' ? list.slice(0, 100) : list;
    const count = k => all.filter(FILTERS[k][1]).length;
    $('#view').innerHTML = `
      <div class="card">
        <form id="tkQuick" class="row" style="margin:0"><input id="tkQ" class="grow" style="flex:1" placeholder="Nouvelle tâche : « Réinstaller PC Vedi demain 14h p1 @Geoffrey »" autocomplete="off">
          <button class="btn primary" type="submit">+ Ajouter</button></form>
        <p class="small muted" style="margin:6px 0 0">Reconnu : date (demain, lundi, 12/11…), heure (14h), p1 à p4, @technicien, #client. Tu vérifies avant d'enregistrer.</p>
      </div>
      <div class="row" style="flex-wrap:wrap;margin-bottom:8px">${Object.entries(FILTERS).map(([k, [l]]) =>
        `<button class="btn chip ${k === filter ? 'active' : ''}" data-tkf="${k}">${l}${k === 'done' ? '' : ` (${count(k)})`}</button>`).join('')}</div>
      <div class="grid2" style="margin-bottom:10px">
        <select id="tkFC"><option value="">Tous les clients</option>${clientOptions(fClient)}</select>
        <select id="tkFT"><option value="">Tous les techniciens</option><option value="-" ${fTech === '-' ? 'selected' : ''}>À prendre</option>
          ${liveUsers().map(x => `<option value="${x.id}" ${x.id === fTech ? 'selected' : ''}>${esc(userLabel(x))}</option>`).join('')}</select>
      </div>
      ${msg ? `<p class="small"><span class="pill warn">${esc(msg)}</span></p>` : ''}
      <div class="card" style="padding:4px 14px">${shown.length ? `<ul class="list tklist">${shown.map(taskLi).join('')}</ul>`
        : '<div class="empty">Aucune tâche ici.</div>'}</div>
      ${list.length > shown.length ? `<p class="small muted">${shown.length} plus récentes sur ${list.length}.</p>` : ''}`;
    $('#tkQuick').onsubmit = ev => {
      ev.preventDefault();
      const q = $('#tkQ').value.trim();
      if (q) editTask(null, parseQuick(q));
    };
    $$('[data-tkf]').forEach(b => b.onclick = () => { filter = b.dataset.tkf; render(); });
    $('#tkFC').onchange = ev => { fClient = ev.target.value; render(); };
    $('#tkFT').onchange = ev => { fTech = ev.target.value; render(); };
    $$('[data-tk] .tap').forEach(el => el.onclick = () => editTask(task(el.closest('[data-tk]').dataset.tk)));
    $$('[data-tkdone]').forEach(b => b.onclick = ev => { ev.stopPropagation(); const t = task(b.dataset.tkdone); isOpen(t) ? finish(t) : reopen(t); });
  }

  // --- fiche d'une tâche (création : preset = champs de l'ajout rapide) ---
  function editTask(t, preset = {}) {
    if (t && t.newFor === me().id) { delete t.newFor; touch(t); save(); scheduleSync(); }
    const d = t ? { ...t } : { title: '', notes: '', clientId: '', assigneeId: me().id, due: '', time: '', priority: 4, ...preset };
    const cl = (t?.checklist || []).map(x => ({ ...x }));
    const edit = !t || canEdit(t), entries = (t?.entryIds || []).map(id => state.entries.find(e => e.id === id && !e.deleted)).filter(Boolean);
    const dlg = $('#dlg'), f = $('#dlgForm');
    const ro = edit ? '' : 'disabled';
    f.innerHTML = `<h2>${t ? (isOpen(t) ? 'Tâche' : 'Tâche terminée') : 'Nouvelle tâche'}</h2>
      ${t ? `<p class="small muted" style="margin-top:0">Créée par ${esc(userName(t.createdBy))} le ${frDate(isoDate(new Date(t.createdAt || Date.now())))}
        ${t.doneAt ? ` · terminée le ${frDate(isoDate(new Date(t.doneAt)))} par ${esc(userName(t.doneBy))}` : ''}</p>` : ''}
      <label>Titre</label><input id="kTitle" value="${esc(d.title)}" ${ro}>
      <label>Description</label><textarea id="kNotes" ${ro}>${esc(d.notes || '')}</textarea>
      <div class="grid2">
        <div><label>Client</label><select id="kCl" ${ro}><option value="">— aucun —</option>${clientOptions(d.clientId)}</select></div>
        <div><label>Technicien</label><select id="kAs" ${ro}><option value="">— à prendre —</option>
          ${liveUsers().filter(x => x.active !== false || x.id === d.assigneeId).map(x => `<option value="${x.id}" ${x.id === d.assigneeId ? 'selected' : ''}>${esc(userLabel(x))}</option>`).join('')}</select></div>
      </div>
      <div class="grid3">
        <div><label>Échéance</label><input type="date" id="kDue" value="${esc(d.due || '')}" ${ro}></div>
        <div><label>Heure</label><input type="time" id="kTime" value="${esc(d.time || '')}" ${ro}></div>
        <div><label>Priorité</label><select id="kPr" ${ro}>${[1, 2, 3, 4].map(p => `<option value="${p}" ${+d.priority === p ? 'selected' : ''}>${PRIO[p]}</option>`).join('')}</select></div>
      </div>
      <label>Points à cocher</label><ul class="list compact" id="kList"></ul>
      ${edit ? `<div class="row"><input id="kItem" class="grow" style="flex:1" placeholder="Ajouter un point…"><button class="btn" id="kAdd" type="button">+</button></div>` : ''}
      ${t ? `<label>Commentaires</label><ul class="list compact">${(t.comments || []).map(c => `<li><div class="grow" style="flex:1;min-width:0">${esc(c.text)}
          <div class="small muted">${esc(userName(c.by))} · ${new Date(c.at).toLocaleString('fr-BE', { dateStyle: 'short', timeStyle: 'short' })}</div></div></li>`).join('') || '<li class="small muted">Aucun commentaire.</li>'}</ul>
        <div class="row"><input id="kCom" class="grow" style="flex:1" placeholder="Ajouter un commentaire…"><button class="btn" id="kComAdd" type="button">Ajouter</button></div>` : ''}
      ${entries.length ? `<label>Interventions liées</label><ul class="list compact">${entries.map(e => `<li class="tap" data-tke="${e.id}"><span class="grow">${frDate(e.date)} · ${esc(client(e.clientId).name)} · ${esc(user(techOf(e)).name)}</span>
          <span class="dur">${fmtDur(e.minutes)}</span></li>`).join('')}</ul>` : ''}
      <div class="actions" style="flex-wrap:wrap">
        ${t && edit ? '<button class="btn danger" id="kDel" style="margin-right:auto">Supprimer</button>' : ''}
        ${t && canTake(t) ? '<button class="btn" id="kTake">🙋 Me l\'attribuer</button>' : ''}
        ${t && isOpen(t) && edit ? '<button class="btn" id="kTimer">▶ Chrono</button><button class="btn" id="kEntry">+ Intervention</button><button class="btn" id="kDone">✓ Terminer</button>' : ''}
        ${t && !isOpen(t) && edit ? '<button class="btn" id="kReopen">↺ Rouvrir</button>' : ''}
        <button class="btn" value="cancel">${edit ? 'Annuler' : 'Fermer'}</button>
        ${edit ? '<button class="btn primary" id="kSave">Enregistrer</button>' : ''}
      </div>`;
    const drawList = () => {
      $('#kList').innerHTML = cl.map((x, i) => `<li><label class="check grow" style="margin:0"><input type="checkbox" data-kci="${i}" ${x.done ? 'checked' : ''} ${ro}> ${esc(x.text)}</label>
        ${edit ? `<button class="btn chip" type="button" data-kcx="${i}" title="Retirer">✕</button>` : ''}</li>`).join('') || '<li class="small muted">Aucun.</li>';
      $$('[data-kci]').forEach(c => c.onchange = () => { cl[+c.dataset.kci].done = c.checked; });
      $$('[data-kcx]').forEach(b => b.onclick = () => { cl.splice(+b.dataset.kcx, 1); drawList(); });
    };
    drawList();
    const addItem = () => { const v = $('#kItem').value.trim(); if (v) { cl.push({ id: uid('ck'), text: v, done: false }); $('#kItem').value = ''; drawList(); } $('#kItem').focus(); };
    if ($('#kAdd')) {
      $('#kAdd').onclick = addItem;
      $('#kItem').onkeydown = ev => { if (ev.key === 'Enter') { ev.preventDefault(); addItem(); } };
    }
    // Valeurs du formulaire → tâche (sans enregistrer).
    const read = () => ({ title: $('#kTitle').value.trim().replace(/\s+/g, ' '), notes: $('#kNotes').value.trim(), clientId: $('#kCl').value || null,
      assigneeId: $('#kAs').value || null, due: $('#kDue').value, time: $('#kTime').value, priority: +$('#kPr').value, checklist: cl });
    const store = async () => {
      const v = read();
      if (!v.title) { alert('Titre requis.'); return null; }
      if (v.time && !v.due) v.due = today();
      let x = t;
      if (!x) T.tasks.push(x = { id: uid('tk'), createdBy: me().id, createdAt: Date.now(), status: 'open', entryIds: [], comments: [] });
      const as = v.assigneeId; delete v.assigneeId;
      Object.assign(x, v); assign(x, as); touch(x);
      await save(); scheduleSync();
      return x;
    };
    if ($('#kSave')) $('#kSave').onclick = async ev => { ev.preventDefault(); if (await store()) { dlg.close(); render(); toast('Tâche enregistrée'); } };
    if ($('#kComAdd')) {
      const addCom = async () => {
        const v = $('#kCom').value.trim(); if (!v) return;
        t.comments = [...(t.comments || []), { id: uid('cm'), by: me().id, at: Date.now(), text: v }]; touch(t);
        await save(); scheduleSync(); editTask(t);
      };
      $('#kComAdd').onclick = addCom;
      $('#kCom').onkeydown = ev => { if (ev.key === 'Enter') { ev.preventDefault(); addCom(); } };
    }
    if ($('#kTake')) $('#kTake').onclick = async () => { assign(t, me().id); touch(t); await save(); scheduleSync(); toast('Tâche attribuée'); editTask(t); };
    if ($('#kDel')) $('#kDel').onclick = async () => {
      if (!confirm(`Supprimer la tâche « ${t.title} » ?${t.entryIds?.length ? '\n(les interventions liées sont conservées)' : ''}`)) return;
      t.deleted = true; touch(t); dlg.close(); await changed(); toast('Tâche supprimée');
    };
    // Chrono / intervention / terminer : les modifications en cours sont d'abord enregistrées.
    if ($('#kTimer')) $('#kTimer').onclick = async () => { const x = await store(); if (x) startTimer(x); };
    if ($('#kEntry')) $('#kEntry').onclick = async () => { const x = await store(); if (x) addEntry(x); };
    if ($('#kDone')) $('#kDone').onclick = async () => { const x = await store(); if (x) finish(x); };
    if ($('#kReopen')) $('#kReopen').onclick = () => reopen(t);
    $$('[data-tke]', f).forEach(li => li.onclick = () => openEntry(state.entries.find(e => e.id === li.dataset.tke)));
    if (!dlg.open) dlg.showModal();
    if (!t) $('#kTitle').focus();
  }

  // --- liens avec le pointage (app.js) ---
  function startTimer(t) {
    if (state.timer && !confirm('Un chrono est déjà en cours. Le remplacer par celui de cette tâche ?')) return;
    const c = t.clientId && state.clients.find(x => x.id === t.clientId && !x.deleted) || state.clients.find(x => x.id === state.lastClientId && !x.deleted) || activeClients()[0];
    if (!c) return alert("Ajoute d'abord un client dans Réglages.");
    setTimer({ clientId: c.id, start: new Date().toISOString(), description: t.title, onSite: false, ...newWho(c), taskId: t.id, pausedMs: 0, pausedAt: null });
    state.lastClientId = c.id;
    $('#dlg').close(); tab = 'today'; persist(); toast('Chrono démarré');
  }
  // Intervention préremplie (client, description) ; le lien avec la tâche est fait à l'enregistrement (linkEntry, appelé par openEntry).
  const addEntry = t => openEntry(null, { ...(t.clientId ? { clientId: t.clientId } : {}), description: t.title, taskId: t.id });
  async function finish(t) {
    Object.assign(t, { status: 'done', doneAt: Date.now(), doneBy: me().id }); touch(t);
    await save(); scheduleSync();
    const f = $('#dlgForm');
    f.innerHTML = `<h2>✓ Tâche terminée</h2><p><strong>${esc(t.title)}</strong>${t.clientId ? ` · ${esc(client(t.clientId).name)}` : ''}</p>
      <p class="small muted">${t.entryIds?.length ? `${t.entryIds.length} intervention(s) déjà liée(s). ` : ''}Ajouter l'intervention correspondante ?</p>
      <div class="actions"><button class="btn" value="cancel">Plus tard</button><button class="btn primary" id="kGoEntry">+ Ajouter l'intervention</button></div>`;
    $('#kGoEntry').onclick = ev => { ev.preventDefault(); addEntry(t); };
    if (!$('#dlg').open) $('#dlg').showModal();
    if (tab === 'taches') view();
  }
  async function reopen(t) {
    Object.assign(t, { status: 'open' }); delete t.doneAt; delete t.doneBy; touch(t);
    if ($('#dlg').open) $('#dlg').close();
    await changed(); toast('Tâche rouverte');
  }
  async function linkEntry(taskId, entryId) {
    await load();
    const t = task(taskId);
    if (!t || (t.entryIds || []).includes(entryId)) return;
    t.entryIds = [...(t.entryIds || []), entryId]; touch(t);
    await save(); scheduleSync();
  }
  const titleOf = id => T && task(id)?.title;
  async function open(id) { await load(); const t = task(id); if (t && !t.deleted) editTask(t); else alert('Tâche introuvable (supprimée ?).'); }

  // --- bandeau de la page Aujourd'hui (app.js : <div id="tkNotif">) ---
  function todayBanner(el) {
    if (!T) { load().then(() => el.isConnected && todayBanner(el)); return; }
    const u = me(), ts = alive().filter(visible);
    const mineOpen = ts.filter(t => t.assigneeId === u.id && isOpen(t));
    const n = { now: mineOpen.filter(dueToday).length, late: mineOpen.filter(overdue).length, free: ts.filter(canTake).length, fresh: mineOpen.filter(t => t.newFor === u.id).length };
    if (!n.now && !n.late && !n.free && !n.fresh) { el.innerHTML = ''; return; }
    el.innerHTML = `<div class="notifbar tk tap" title="Ouvrir l'onglet Tâches"><span class="ico">📋</span>
      <div class="txt"><strong>${[n.now && `${n.now} tâche(s) aujourd'hui`, n.late && `${n.late} en retard`].filter(Boolean).join(' · ') || 'Tâches'}</strong>
        <div class="small state">${[n.fresh && `${n.fresh} nouvelle(s) pour toi`, n.free && `${n.free} à prendre`].filter(Boolean).join(' · ') || 'rien de nouveau'}</div></div>
      <span class="go">›</span></div>`;
    $('.notifbar', el).onclick = () => { filter = n.now || n.late ? 'today' : n.fresh ? 'open' : 'free'; tab = 'taches'; render(); };
  }

  async function reset() { T = null; msg = ''; await idb('readwrite', s => s.delete('data')); }

  return { view, sync, todayBanner, linkEntry, titleOf, open, reset, parseQuick };
})();
