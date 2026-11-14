// © 2026 One Level UP – LevelDesk. Logiciel propriétaire, tous droits réservés. Copie, modification, redistribution et rétro-ingénierie interdites sans autorisation écrite (voir LICENCE.txt).
// Module Tâches (voir taches.md) : tâches « à la Todoist », liées à un client et à un technicien ; une tâche mène à une ou plusieurs interventions.
// - Données : TimeSheet/App/taches-data.json, avec sa propre synchro (comme Achats). heures-data.json : seulement entries[].taskId.
//   Avec le serveur OVH (srvOn(), architecture.md) : table tasks, synchronisée par app.js (srvSync) ; plus de fichier OneDrive.
// - Technicien vide = « à prendre » : un technicien peut se l'attribuer si le client fait partie de ses clients visibles (Admin → Utilisateurs).
// - Tout technicien peut assigner à tout autre ; le créateur et le technicien assigné peuvent modifier, réassigner, terminer, supprimer.
// Utilise les outils de app.js ($, esc, uid, touch, state, client, me, canSee, openEntry…) : chargé avant lui, appelé après le démarrage.
const Taches = (() => {
  let T = null, busy = false, msg = '';
  // Préférences d'affichage de cet appareil (filtre choisi, groupes repliés) : localStorage, sans importance si perdu.
  const UI_KEY = 'heures.tkui';
  const ui = (() => { try { return JSON.parse(localStorage.getItem(UI_KEY)) || {}; } catch { return {}; } })();
  const saveUi = () => { try { localStorage.setItem(UI_KEY, JSON.stringify(ui)); } catch { /* stockage indisponible */ } };
  let filter = ui.filter || 'open', fClient = '', fTech = '', q = '';   // filtre par défaut : « Toutes » (demande de Fred, 30/10)

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
  // Privée (sans client) : seulement sa personne (créateur et technicien assigné), même pour un administrateur, et jamais « à prendre ».
  const isPrivate = t => !t.clientId;
  // Visible : assignée à moi, créée par moi, à prendre sur un de mes clients ; responsable : tout ce qui touche sa société ; administrateur : tout.
  function visible(t) {
    const u = me();
    if (t.assigneeId === u.id || t.createdBy === u.id) return true;
    if (isPrivate(t)) return false;
    if (isAdmin()) return true;
    if (!t.assigneeId && canTakeClient(t)) return true;
    return isResp() && [socOfUser(t.assigneeId), socOfUser(t.createdBy), t.clientId && clientSoc(client(t.clientId))].includes(u.societeId);
  }
  const canTakeClient = t => !isPrivate(t) && canSee(client(t.clientId));
  const canEdit = t => (isAdmin() && !isPrivate(t)) || t.assigneeId === me().id || t.createdBy === me().id;
  const canTake = t => !t.assigneeId && t.status !== 'done' && canTakeClient(t);
  const isOpen = t => t.status !== 'done';
  const overdue = t => isOpen(t) && t.due && t.due < today();
  const dueToday = t => isOpen(t) && t.due === today();
  const sortKey = t => `${t.due || '9999-99-99'} ${t.time || '99:99'} ${t.priority || 4}`;
  // Planifiée : créée automatiquement chaque mois (mises à jour serveur) ; rangée à part dans la liste.
  const isPlanned = t => t.kind === 'upd' || !!t.recur;
  // Importée depuis Outlook (plus tard : assistant de boîte mail, t.source = 'outlook') ; rangée à part aussi.
  const isOutlook = t => t.source === 'outlook';
  const lateDays = t => Math.round((new Date(today() + 'T00:00') - new Date(t.due + 'T00:00')) / 864e5);
  const PRIO = { 1: 'P1 urgent', 2: 'P2 haute', 3: 'P3 moyenne', 4: 'P4 normale' };

  // Assignation : le nouveau technicien est prévenu dans son bandeau (newFor, effacé quand il ouvre la tâche).
  function assign(t, id) {
    if ((t.assigneeId || '') === (id || '')) return;
    t.assigneeId = id || null;
    if (id && id !== me().id) t.newFor = id; else delete t.newFor;
  }

  // --- récurrence : mises à jour serveur, le 2ᵉ mercredi de chaque mois (taches.md, section 5) ---
  // Une tâche par client avec serveur(s), assignée à Fred (administrateur), créée pour le mois en cours.
  // ID fixe upd_<client>_<AAAA-MM> : pas de doublon entre appareils ; une tâche supprimée n'est pas recréée.
  // Pas créée si la mise à jour du mois est déjà encodée (forfait « Mise à jour serveur »). Celles du mois suivant apparaissent le 1er.
  const updId = (clientId, month) => `upd_${clientId}_${month}`;
  function secondWednesday(month) {
    const d = new Date(month + '-01T00:00');
    d.setDate(1 + (3 - d.getDay() + 7) % 7 + 7);
    return isoDate(d);
  }
  function ensureRecurring() {
    if (!isAdmin() || !mod('forfaits')) return false;        // seul un administrateur les crée (pour lui-même)
    // Serveur : une société gérée par une autre (Alescom, gérée par One Level UP) ne les crée pas, c'est la société qui la gère qui le fait ;
    // et seulement pour ses propres clients et ceux des sociétés qu'elle gère. Vu le 13/11 : l'administrateur d'Alescom recréait toutes les 2 s
    // celles de 4 clients One Level UP (BDS, JMD, MC&S, Vedi), déjà créées par Fred et invisibles pour lui → refus « interdit » en boucle.
    if (srvOn() && soc(HOME).managedBy) return false;
    const month = today().slice(0, 7);
    let ch = false;
    for (const c of activeClients().filter(x => x.server && (!srvOn() || clientSoc(x) === HOME || managedByMe(clientSoc(x))))) {
      const id = updId(c.id, month);
      if (task(id) || srvRefused('tasks', id)) continue;      // refusée par le serveur : pas recréée (sinon boucle de synchro)
      if (live().some(e => e.forfait === 'update' && e.clientId === c.id && e.date.startsWith(month))) continue;
      T.tasks.push(touch({ id, kind: 'upd', recur: { every: 'month', nth: 2, weekday: 3 }, title: `Mise à jour serveur – ${c.name}`, notes: '',
        clientId: c.id, assigneeId: me().id, createdBy: me().id, createdAt: Date.now(), due: secondWednesday(month), time: '', priority: 3,
        status: 'open', checklist: [], comments: [], entryIds: [] }));
      ch = true;
    }
    return ch;
  }
  // Forfait « Mise à jour serveur » enregistré (app.js) : relié à sa tâche du mois, qui est terminée.
  async function forfaitSaved(e) {
    await load();
    const t = task(e.taskId) || task(updId(e.clientId, e.date.slice(0, 7)));
    if (!t || t.deleted) return;
    if (!(t.entryIds || []).includes(e.id)) t.entryIds = [...(t.entryIds || []), e.id];
    if (isOpen(t)) Object.assign(t, { status: 'done', doneAt: Date.now(), doneBy: me().id });
    if (!e.taskId) { e.taskId = t.id; touch(e); }
    touch(t); await save(); scheduleSync();
  }

  // --- synchronisation OneDrive (appelée par la synchro du pointage, app.js) ---
  function merge(remote) {
    const before = JSON.stringify(T.tasks);
    const x = new Map(T.tasks.map(o => [o.id, o]));
    for (const o of remote.tasks || []) { const y = x.get(o.id); if (!y || (o.updatedAt || 0) > (y.updatedAt || 0)) x.set(o.id, o); }
    T.tasks = [...x.values()];
    return before !== JSON.stringify(T.tasks);
  }
  // --- tâches de l'assistant de boîte mail (tâche planifiée de claude.ai, voir Claude/mailbox_assistant/PLAYBOOK.md, section 3) ---
  // L'assistant ne peut pas appeler le serveur LevelDesk (pas de jeton Microsoft) : il écrit ses tâches dans le OneDrive de Fred,
  // <dossier>/App/taches-outlook.json = { version: 1, items: [{ id, title, notes, client, from, subject, webLink, priority, due, time, createdAt,
  //   comments: [{ id, at, text, priority }] }] }. Ce fichier appartient à l'assistant : LevelDesk le lit seulement.
  // LevelDesk crée une tâche ol_<id> par élément (source 'outlook', assignée à la personne connectée) et ajoute les commentaires nouveaux.
  // <dossier>/App/taches-outlook-etat.json (écrit par LevelDesk, lu par l'assistant) = { version: 1, tasks: { ol_<id>: { status, title } } } :
  // dit à l'assistant ce qui est ouvert / terminé / supprimé, et aux autres appareils ce qui est déjà importé (une tâche supprimée n'est pas recréée).
  // Seulement pour un administrateur dont le compte a un OneDrive ; au plus toutes les 5 minutes.
  const olPath = () => `${state.onedrive.folder}/App/taches-outlook.json`;
  const olStatePath = () => `${state.onedrive.folder}/App/taches-outlook-etat.json`;
  const OL_EVERY = 5 * 60e3;
  let olAt = 0, olBusy = false;
  const olStr = (v, n) => String(v ?? '').replace(/\s+\n/g, '\n').trim().slice(0, n);
  const olId = v => { const s = String(v ?? '').replace(/[^\w-]/g, '').slice(0, 60); return s ? 'ol_' + s : ''; };
  const olTime = v => { const n = Date.parse(v); return n > 0 ? n : Date.now(); };
  const mailLink = v => { try { const u = new URL(v); return u.protocol === 'https:' && /(^|\.)(office365\.com|office\.com|live\.com|cloud\.microsoft)$/.test(u.hostname) ? u.href : ''; } catch { return ''; } };
  // Client : nom exact (ou nom Billit), sinon domaine de l'expéditeur dans la fiche d'un client. Sans client : tâche privée.
  function olClient(it) {
    const cs = activeClients().filter(canSee), n = fold(olStr(it.client, 100)), dom = fold(olStr(it.from, 200)).match(/@([\w.-]+\.[a-z]{2,})/)?.[1];
    const c = (n && cs.find(x => fold(x.name) === n || fold(x.billitName) === n))
      || (dom && cs.find(x => (x.domains || []).some(d => dom === fold(d) || dom.endsWith('.' + fold(d)))));
    return c ? c.id : null;
  }
  async function importOutlook() {
    if (olBusy || !isAdmin() || !navigator.onLine || !OneDrive.signedIn() || Date.now() - olAt < OL_EVERY) return false;
    olBusy = true; olAt = Date.now();
    let ch = false, created = 0;
    try {
      if (!(await OneDrive.hasDrive())) return false;
      const f = await OneDrive.get(olPath());
      if (!f) return false;
      let items;
      try { items = JSON.parse(f.data.replace(/^﻿/, '')).items; } catch { throw new Error('taches-outlook.json illisible (JSON invalide) : tâches Outlook pas importées.'); }
      if (!Array.isArray(items)) return false;
      const st = await OneDrive.get(olStatePath());
      let known = {};
      try { known = JSON.parse(st.data.replace(/^﻿/, '')).tasks || {}; } catch { /* pas encore de fichier d'état */ }
      const pr = v => [1, 2, 3, 4].includes(+v) ? +v : 0;
      for (const it of items.slice(-500)) {
        const id = olId(it?.id), title = olStr(it?.title, 200);
        if (!id || !title) continue;
        let t = task(id);
        if (!t && !known[id]) {
          const due = /^\d{4}-\d{2}-\d{2}$/.test(it.due) ? it.due : '', link = mailLink(it.webLink);
          T.tasks.push(t = touch({ id, source: 'outlook', title, notes: olStr(it.notes, 4000), clientId: olClient(it), assigneeId: me().id, newFor: me().id,
            createdBy: me().id, createdAt: olTime(it.createdAt), due, time: due && /^\d{2}:\d{2}$/.test(it.time) ? it.time : '', priority: pr(it.priority) || 3,
            status: 'open', checklist: [], comments: [], entryIds: [],
            mail: { from: olStr(it.from, 200), subject: olStr(it.subject, 300), link } }));
          ch = true; created++;
        }
        if (!t || t.deleted) continue;
        // Suite du fil : commentaires ajoutés par l'assistant (une fois chacun) ; la priorité ne peut que monter.
        for (const c of Array.isArray(it.comments) ? it.comments.slice(-100) : []) {
          const cid = olId(c?.id).replace(/^ol_/, 'olc_'), text = olStr(c?.text, 2000);
          if (!cid || !text || (t.comments || []).some(x => x.id === cid)) continue;
          t.comments = [...(t.comments || []), { id: cid, by: 'outlook', at: olTime(c.at), text }];
          if (pr(c.priority) && pr(c.priority) < (t.priority || 4)) t.priority = pr(c.priority);
          if (isOpen(t) && t.assigneeId) t.newFor = t.assigneeId;
          touch(t); ch = true;
        }
      }
      // État pour l'assistant et les autres appareils (écrit seulement s'il change).
      const tasks = { ...known };
      for (const t of T.tasks.filter(isOutlook)) tasks[t.id] = { status: t.deleted ? 'deleted' : isOpen(t) ? 'open' : 'done', title: t.title };
      const body = JSON.stringify({ version: 1, tasks }, null, 1);
      if (!st || hash(st.data) !== hash(body)) {
        const r = await OneDrive.put(olStatePath(), body, { etag: st?.etag, type: 'application/json' });
        if (!r.ok && ![409, 412].includes(r.status)) throw new Error(`OneDrive (${r.status}) : écriture de taches-outlook-etat.json impossible.`);
      }
      if (created) toast(`📧 ${created} tâche${created > 1 ? 's' : ''} importée${created > 1 ? 's' : ''} depuis Outlook`);
    } catch (e) {
      if (!(e instanceof OneDrive.NeedLogin)) olMsg = e.message;
      return ch;
    } finally { olBusy = false; }
    olMsg = '';
    return ch;
  }
  let olMsg = '';

  // Avec le serveur : les tâches sont envoyées et reçues par srvSync (app.js) ; ici, seulement les tâches récurrentes et l'affichage.
  async function serverSync() {
    await load();
    const hadOl = olMsg;
    const ch = ensureRecurring() | await importOutlook(), had = !!msg || hadOl !== olMsg;
    if (ch) { await save(); scheduleSync(); }
    msg = '';
    // La page Tâches n'est redessinée que s'il y a du neuf ici : les tâches reçues du serveur le sont déjà par srvSync (app.js).
    // Avant : redessinée après CHAQUE synchro (toutes les minutes), ce qui effaçait la saisie et la position dans la page.
    afterSync(ch || had);
  }
  function afterSync(redraw) {
    if (redraw && tab === 'taches' && !$('#dlg').open) render();
    if (tab === 'today' && $('#tkNotif')) todayBanner($('#tkNotif'));
    if (tab === 'today' && $('#tkDay')) dayList($('#tkDay'));
  }
  async function sync() {
    if (srvOn()) return serverSync();
    await load();
    if (busy || !OneDrive.signedIn() || !navigator.onLine) return;
    busy = true;
    let redraw = false;
    try {
      let cur = await OneDrive.get(dataPath());
      if (cur) redraw = merge(JSON.parse(cur.data));
      redraw = ensureRecurring() || redraw;            // après la fusion : les tâches du mois déjà créées ailleurs sont connues
      const hadOl = olMsg;
      redraw = await importOutlook() || hadOl !== olMsg || redraw;   // tâches de l'assistant de boîte mail (écrites plus bas avec les autres)
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
    afterSync(redraw);
  }

  // --- ajout rapide : « Réinstaller PC accueil Vedi demain 14h p1 @Geoffrey » (au clavier ou dicté) ---
  // Reconnu : p1…p4, « priorité 1 », « urgent » ; @technicien ; #client (sinon le nom d'un client dans le texte) ;
  // aujourd'hui / ce soir / demain / après-demain / lundi [prochain] / la semaine prochaine / dans N jours ou semaines / 12/11[/2026] / le 12 novembre ;
  // heure 14h, 14h30, 14 h 30, 14 heures, 14:30, midi. Les petits mots devant (« à », « le », « pour »…) partent avec. Le reste devient le titre ;
  // tout est vérifié dans la fiche avant d'enregistrer.
  const DAYS = ['dimanche', 'lundi', 'mardi', 'mercredi', 'jeudi', 'vendredi', 'samedi'];
  const MONTHS = ['janvier', 'fevrier', 'mars', 'avril', 'mai', 'juin', 'juillet', 'aout', 'septembre', 'octobre', 'novembre', 'decembre'];
  const MONTH_RE = 'janvier|f[ée]vrier|mars|avril|mai|juin|juillet|ao[uû]t|septembre|octobre|novembre|d[ée]cembre';
  const addDays = n => { const d = new Date(); d.setDate(d.getDate() + n); return isoDate(d); };
  const inDays = wd => (wd - new Date().getDay() + 7) % 7 || 7;          // prochain jour de semaine wd (jamais aujourd'hui)
  // Date sans année déjà passée : l'an prochain.
  const dayMonth = (d, mo, y) => {
    if (!d || d > 31 || !mo || mo > 12) return null;
    const yy = y ? +(String(y).length === 2 ? '20' + y : y) : new Date().getFullYear();
    const iso = `${yy}-${pad(mo)}-${pad(d)}`;
    return !y && iso < today() ? `${yy + 1}-${pad(mo)}-${pad(d)}` : iso;
  };
  function parseQuick(s) {
    // Apostrophes du téléphone et de la dictée (’) ; ponctuation de la dictée (« demain. »).
    let txt = ` ${String(s).replace(/[’‘]/g, "'").replace(/[,;!?]|\.(?=\s|$)/g, ' ')} `;
    const out = {};
    // fn renvoie false : le morceau reste dans le titre.
    const take = (re, fn) => { txt = txt.replace(re, (...m) => fn(m) === false ? m[0] : ' '); };
    const pre = '(?:(?:pour|le|ce|d\'ici|avant)\\s)*';                   // petits mots devant une date, retirés avec elle
    take(/\s(?:[pP]\s?([1-4])|priorit[ée]\s?([1-4]))(?=\s)/, m => { out.priority = +(m[1] || m[2]); });
    take(/\surgente?(?=\s)/i, () => { out.priority = 1; });
    take(/\s@(\S+)/, m => {
      const u = liveUsers().find(x => x.active !== false && fold(x.name).startsWith(fold(m[1])));
      if (!u) { out.unknownUser = m[1]; return false; }
      out.assigneeId = u.id;
    });
    take(/\s#(\S+)/, m => {
      const c = activeClients().filter(canSee).find(x => fold(x.name).replace(/\s+/g, '').startsWith(fold(m[1])));
      if (!c) { out.unknownClient = m[1]; return false; }
      out.clientId = c.id;
    });
    take(/\s(?:(?:à|a|vers|pour)\s)?(\d{1,2})\s?(?:h|heures?)(?:\s?(\d{2}))?(?=\s)/i, m => {
      const h = +m[1], mi = +(m[2] || 0);
      if (h > 23 || mi > 59) return false; out.time = `${pad(h)}:${pad(mi)}`;
    });
    take(/\s(?:(?:à|a|vers)\s)?(\d{1,2}):(\d{2})(?=\s)/, m => {
      if (+m[1] > 23 || +m[2] > 59) return false; out.time = `${pad(+m[1])}:${m[2]}`;
    });
    take(/\s(?:(?:à|a|vers)\s)?midi(?=\s)/i, () => { out.time ||= '12:00'; });
    take(new RegExp(`\\s${pre}après[- ]demain(?=\\s)`, 'i'), () => { out.due = addDays(2); });
    take(new RegExp(`\\s${pre}demain(?:\\s(?:matin|soir))?(?=\\s)`, 'i'), () => { out.due = addDays(1); });
    take(new RegExp(`\\s${pre}(?:aujourd'hui|auj|ce\\s(?:soir|matin)|cet\\saprès-midi)(?=\\s)`, 'i'), () => { out.due = today(); });
    take(new RegExp(`\\s${pre}(?:dans|d'ici)\\s(\\d{1,3})\\s(jours?|semaines?)(?=\\s)`, 'i'), m => { out.due = addDays(+m[1] * (/^s/i.test(m[2]) ? 7 : 1)); });
    take(new RegExp(`\\s${pre}(?:la\\s)?semaine\\sprochaine(?=\\s)`, 'i'), () => { out.due = addDays(inDays(1)); });
    take(new RegExp(`\\s${pre}(${DAYS.join('|')})(?:\\s(?:prochain|matin|soir|après-midi))?(?=\\s)`, 'i'), m => { out.due = addDays(inDays(DAYS.indexOf(m[1].toLowerCase()))); });
    take(new RegExp(`\\s${pre}(\\d{1,2})(?:er)?\\s(${MONTH_RE})(?:\\s(\\d{4}))?(?=\\s)`, 'i'), m => {
      const iso = dayMonth(+m[1], MONTHS.indexOf(fold(m[2])) + 1, m[3]);
      if (!iso) return false; out.due = iso;
    });
    take(new RegExp(`\\s${pre}(\\d{1,2})\\/(\\d{1,2})(?:\\/(\\d{2,4}))?(?=\\s)`), m => {
      const iso = dayMonth(+m[1], +m[2], m[3]);
      if (!iso) return false; out.due = iso;
    });
    if (out.time && !out.due) out.due = today();
    out.title = txt.replace(/\s+/g, ' ').trim();
    out.title = out.title.charAt(0).toUpperCase() + out.title.slice(1);
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
  // Report rapide (liste et groupe « En retard ») : liste déroulante du système, simple au doigt comme à la souris.
  const POSTPONE = [['0', "Aujourd'hui"], ['1', 'Demain'], ['mon', 'Lundi prochain'], ['7', 'Dans 1 semaine'], ['-', 'Sans date']];
  const postponeSel = (attr, title, first = '📅') => `<select class="tkpost" ${attr} title="${title}" aria-label="${title}"><option value="">${first}</option>
    ${POSTPONE.map(([k, l]) => `<option value="${k}">${l}</option>`).join('')}</select>`;
  const postDate = k => k === '-' ? '' : addDays(k === 'mon' ? inDays(1) : +k);
  function taskLi(t) {
    const cl = t.checklist || [], done = cl.filter(x => x.done).length, mine = t.assigneeId === me().id;
    const late = overdue(t) ? lateDays(t) : 0;
    const meta = [
      t.clientId ? esc(client(t.clientId).name) : '🔒 privé',
      t.assigneeId ? (mine ? '' : `👤 ${esc(userName(t.assigneeId))}`) : '<span class="pill warn">à prendre</span>',
      t.due && (late ? `<span class="late">${esc(dueLabel(t))}</span> <span class="pill bad">${late} j de retard</span>`
        : `<span class="${dueToday(t) ? 'now' : ''}">${esc(dueLabel(t))}</span>`),
      isPlanned(t) && '🔁', isOutlook(t) && '📧 Outlook',
      cl.length && `☑ ${done}/${cl.length}`,
      t.comments?.length && `💬 ${t.comments.length}`,
      t.entryIds?.length && `🔗 ${t.entryIds.length}`
    ].filter(Boolean).join(' · ');
    return `<li class="tk p${t.priority || 4} ${isOpen(t) ? '' : 'done'} ${t.id === flashId ? 'flash' : ''}" data-tk="${t.id}">
      <button class="tkdone" data-tkdone="${t.id}" title="${isOpen(t) ? 'Terminer' : 'Rouvrir'}" aria-label="${isOpen(t) ? 'Terminer' : 'Rouvrir'}" ${canEdit(t) ? '' : 'disabled'}>${isOpen(t) ? '' : '✓'}</button>
      <div class="grow tap tkbody"><div class="tkt">${esc(t.title)}${t.newFor === me().id ? ' <span class="pill ok">nouvelle</span>' : ''}${t.priority && t.priority < 3 ? ` <span class="pill pr${t.priority}">P${t.priority}</span>` : ''}</div>
        <div class="small muted">${meta}</div></div>
      ${isOpen(t) && canEdit(t) ? postponeSel(`data-tkpost="${t.id}"`, 'Reporter') : ''}</li>`;
  }
  // Clics d'une liste de tâches (page Tâches et tableau de bord) : ouvrir, terminer / rouvrir, reporter.
  function bindRows(el, after) {
    $$('[data-tk] .tap', el).forEach(x => x.onclick = () => editTask(task(x.closest('[data-tk]').dataset.tk)));
    $$('[data-tkdone]', el).forEach(b => b.onclick = async ev => { ev.stopPropagation(); const t = task(b.dataset.tkdone); await (isOpen(t) ? finish(t) : reopen(t)); after?.(); });
    $$('[data-tkpost]', el).forEach(s => s.onchange = async () => {
      if (!s.value) return;
      await postpone([task(s.dataset.tkpost)], s.value); after?.();
    });
  }
  async function postpone(ts, k) {
    const due = postDate(k);
    for (const t of ts) { t.due = due; if (!due) t.time = ''; touch(t); }
    await save(); scheduleSync();
    toast(`${ts.length > 1 ? ts.length + ' tâches reportées' : 'Tâche reportée'} · ${due ? frDate(due) : 'sans date'}`);
    if (tab === 'taches') render();
  }

  const QUICK_HELP = "Exemples : « demain 14h », « lundi », « le 12 novembre », « p1 » ou « urgent », « @Geoffrey », « #Vedi ». Champ vide + Ajouter : fiche vierge.";
  // Aperçu de l'ajout rapide : ce que la fiche contiendra.
  function quickPreview(o) {
    const due = o.due ? dueLabel(o) : '';
    return [`<strong>${esc(o.title || '(sans titre)')}</strong>`,
      o.clientId ? `🏢 ${esc(client(o.clientId).name)}` : '🔒 privé',
      `👤 ${esc(userName(o.assigneeId || me().id))}`,
      due && `📅 ${esc(due)}`,
      o.priority && o.priority < 4 && `<span class="pill">${PRIO[o.priority]}</span>`,
      o.unknownUser && `<span class="pill warn">technicien « ${esc(o.unknownUser)} » inconnu</span>`,
      o.unknownClient && `<span class="pill warn">client « ${esc(o.unknownClient)} » inconnu</span>`
    ].filter(Boolean).join(' · ');
  }
  // Filtres (puces) ; « Toutes » par défaut, le dernier choix est gardé sur l'appareil.
  const filters = u => ({
    me: ['Mes tâches', t => isOpen(t) && t.assigneeId === u.id],
    open: ['Toutes', t => isOpen(t)],
    free: ['À prendre', t => canTake(t)],
    byme: ['Assignées par moi', t => isOpen(t) && t.createdBy === u.id && t.assigneeId !== u.id],
    done: ['Terminées', t => !isOpen(t)]
  });
  const setFilter = k => { filter = k; ui.filter = k; saveUi(); };
  // Après création : filtre qui montre la tâche (sinon elle semblait perdue).
  const FILTER_OF = x => !isOpen(x) ? 'done' : (filters(me())[filter]?.[1](x) ? filter : 'open');
  const FILTER_NAME = { me: 'Mes tâches', open: 'Toutes', free: 'À prendre', byme: 'Assignées par moi', done: 'Terminées' };
  let flashId = null;
  // Groupes par échéance (maquette A, 13/11) ; les tâches planifiées et celles d'Outlook sont rangées à part, en dessous.
  const GROUPS = [
    ['late', 'En retard'], ['today', "Aujourd'hui"], ['tomorrow', 'Demain'], ['week', 'Cette semaine'], ['later', 'Plus tard'], ['nodate', 'Sans date'],
    ['outlook', '📧 Depuis Outlook'], ['planned', '🔁 Planifiées · mises à jour serveur']
  ];
  const groupOf = t => isPlanned(t) ? 'planned' : isOutlook(t) ? 'outlook' : !t.due ? 'nodate' : overdue(t) ? 'late' : dueToday(t) ? 'today'
    : t.due === addDays(1) ? 'tomorrow' : t.due <= addDays(7) ? 'week' : 'later';
  const matches = t => !q || fold([t.title, t.notes, t.clientId && client(t.clientId).name, t.assigneeId && userName(t.assigneeId)].join(' ')).includes(fold(q));

  function view() {
    if (!T) { $('#view').innerHTML = '<div class="empty">Chargement…</div>'; load().then(() => tab === 'taches' && render()); return; }
    const u = me(), all = alive().filter(visible), F = filters(u);
    if (!F[filter]) setFilter('open');
    const count = k => all.filter(F[k][1]).length;
    const fOn = (fClient ? 1 : 0) + (fTech ? 1 : 0);
    $('#view').innerHTML = `
      <div class="card tkadd">
        <form id="tkQuick" class="row" style="margin:0"><input id="tkQ" class="grow" style="flex:1" placeholder="Nouvelle tâche : « Réinstaller PC Vedi demain 14h p1 »" autocomplete="off">
          ${mic('tkQ', true)}<button class="btn primary" type="submit" aria-label="Ajouter">＋</button></form>
        <p class="small muted" id="tkQPrev" style="margin:6px 0 0">${QUICK_HELP}</p>
      </div>
      <div class="tkchips">${Object.entries(F).map(([k, [l]]) =>
        `<button class="btn chip ${k === filter ? 'active' : ''}" data-tkf="${k}">${l}${k === 'done' ? '' : ` <span class="n">${count(k)}</span>`}</button>`).join('')}</div>
      <div class="tktools"><input id="tkS" type="search" placeholder="🔍 Rechercher : titre, client, technicien…" value="${esc(q)}" autocomplete="off">
        <button class="btn chip ${fOn ? 'active' : ''}" type="button" id="tkFx">Filtres${fOn ? ` <span class="n">${fOn}</span>` : ''}</button></div>
      <div class="grid2" id="tkFbox" style="margin-bottom:10px" ${fOn || ui.fbox ? '' : 'hidden'}>
        <select id="tkFC"><option value="">Tous les clients</option>${clientOptions(fClient)}</select>
        <select id="tkFT"><option value="">Tous les techniciens</option><option value="-" ${fTech === '-' ? 'selected' : ''}>À prendre</option>
          ${liveUsers().map(x => `<option value="${x.id}" ${x.id === fTech ? 'selected' : ''}>${esc(userLabel(x))}</option>`).join('')}</select>
      </div>
      ${msg || olMsg ? `<p class="small"><span class="pill warn">${esc(msg || olMsg)}</span></p>` : ''}
      <div id="tkBody"></div>`;
    // La liste seule est redessinée pendant la recherche (le champ garde le focus).
    const drawBody = () => {
      const list = all.filter(F[filter][1]).filter(t => (!fClient || t.clientId === fClient) && (!fTech || (fTech === '-' ? !t.assigneeId : t.assigneeId === fTech))).filter(matches);
      const body = $('#tkBody');
      if (filter === 'done') {
        list.sort((a, b) => (b.doneAt || 0) - (a.doneAt || 0));
        const shown = list.slice(0, 100);
        body.innerHTML = `<div class="card tkcard">${shown.length ? `<ul class="list tklist">${shown.map(taskLi).join('')}</ul>` : '<div class="empty">Aucune tâche terminée.</div>'}</div>
          ${list.length > shown.length ? `<p class="small muted">${shown.length} plus récentes sur ${list.length}.</p>` : ''}`;
      } else {
        list.sort((a, b) => sortKey(a).localeCompare(sortKey(b)));
        const by = {}; for (const t of list) (by[groupOf(t)] ||= []).push(t);
        // « Depuis Outlook » : affiché vide pour un administrateur, en attendant l'import (assistant de boîte mail).
        const showEmpty = k => k === 'outlook' && isAdmin() && !q && !fOn && filter !== 'free';
        const secs = GROUPS.filter(([k]) => by[k]?.length || showEmpty(k)).map(([k, l]) => {
          const ts = by[k] || [], closed = !!ui.closed?.[k], lateN = ts.filter(overdue).length;
          const editable = ts.filter(canEdit);
          return `<section class="tkgrp g-${k} ${closed ? 'closed' : ''}">
            <div class="tkgh"><button type="button" class="tkgt" data-tkg="${k}" aria-expanded="${!closed}"><span class="arr">${closed ? '▸' : '▾'}</span> ${l} <span class="n">${ts.length}</span>
              ${k === 'planned' && lateN ? ` <span class="pill bad">${lateN} en retard</span>` : ''}</button>
              ${k === 'late' && editable.length > 1 ? postponeSel('data-tkpostall="late"', 'Tout reporter', 'Tout reporter…') : ''}</div>
            ${closed ? '' : ts.length ? `<ul class="list tklist">${ts.map(taskLi).join('')}</ul>`
              : '<p class="small muted tkempty">Les tâches créées par l\'assistant de boîte mail arrivent ici.</p>'}</section>`;
        }).join('');
        body.innerHTML = secs ? `<div class="card tkcard">${secs}</div>`
          : `<div class="empty">${q || fOn ? 'Aucune tâche ne correspond.' : filter === 'me' ? 'Rien pour toi 👍' : 'Aucune tâche ici.'}</div>`;
        $$('[data-tkg]', body).forEach(b => b.onclick = () => { ui.closed = { ...(ui.closed || {}), [b.dataset.tkg]: !ui.closed?.[b.dataset.tkg] }; saveUi(); drawBody(); });
        const all2 = $('[data-tkpostall]', body);
        if (all2) all2.onchange = () => all2.value && postpone((by.late || []).filter(canEdit), all2.value);
      }
      bindRows(body);
    };
    drawBody();
    // Champ vide : fiche vierge ; sinon fiche préremplie. Sous le champ, ce qui sera reconnu (mis à jour en tapant ou pendant la dictée).
    $('#tkQuick').onsubmit = ev => {
      ev.preventDefault();
      const q = $('#tkQ').value.trim();
      editTask(null, q ? parseQuick(q) : {});
    };
    $('#tkQ').oninput = () => { const q = $('#tkQ').value.trim(); $('#tkQPrev').innerHTML = q ? quickPreview(parseQuick(q)) : QUICK_HELP; };
    $$('[data-tkf]').forEach(b => b.onclick = () => { setFilter(b.dataset.tkf); render(); });
    $('#tkS').oninput = ev => { q = ev.target.value.trim(); drawBody(); };
    $('#tkFx').onclick = () => { const b = $('#tkFbox'); b.hidden = !b.hidden; ui.fbox = !b.hidden; saveUi(); };
    $('#tkFC').onchange = ev => { fClient = ev.target.value; render(); };
    $('#tkFT').onchange = ev => { fTech = ev.target.value; render(); };
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
      ${t?.kind === 'upd' ? `<p class="small"><span class="pill">🔁 chaque mois, 2ᵉ mercredi</span> « Terminer » ouvre le forfait « Mise à jour serveur » du client ;
        une fois enregistré, la tâche est terminée.</p>` : ''}
      ${t?.mail ? `<p class="small" style="margin-top:0"><span class="pill">📧 Outlook</span> ${esc([t.mail.from, t.mail.subject].filter(Boolean).join(' · '))}
        ${mailLink(t.mail.link) ? ` · <a href="${esc(mailLink(t.mail.link))}" target="_blank" rel="noopener noreferrer">Ouvrir l'e-mail</a>` : ''}</p>` : ''}
      ${t ? (() => {
        // Deux actions en vue (la plus utile d'abord), les autres sous « ⋯ » (maquette A, 14/11).
        const acts = [
          isOpen(t) && edit && '<button class="btn tkok" id="kDone" type="button">✓ Terminer</button>',
          !isOpen(t) && edit && '<button class="btn" id="kReopen" type="button">↺ Rouvrir</button>',
          canTake(t) && '<button class="btn" id="kTake" type="button">🙋 Me l\'attribuer</button>',
          isOpen(t) && edit && t.kind !== 'upd' && !isPrivate(t) && '<button class="btn" id="kTimer" type="button">▶ Chrono</button>',
          isOpen(t) && edit && t.kind !== 'upd' && !isPrivate(t) && '<button class="btn" id="kEntry" type="button">+ Intervention</button>',
          isOpen(t) && edit && `<button class="btn" id="kOut" type="button" title="Rendez-vous de 1 h à l'heure de la tâche, sinon journée entière">📅 ${t.outlook?.[me().id] ? 'Mettre à jour Outlook' : 'Agenda Outlook'}</button>`
        ].filter(Boolean);
        // Supprimer : toujours sous « ⋯ » (jamais à portée d'un clic par erreur).
        const main = acts.slice(0, 2), more = [...acts.slice(2), edit && '<button class="btn danger" id="kDel" type="button">🗑 Supprimer</button>'].filter(Boolean);
        if (!main.length && !more.length) return '';
        return `<div class="tkacts">${main.join('')}
          ${more.length ? `<button class="btn tkmore" id="kMore" type="button" aria-expanded="false" title="Autres actions" aria-label="Autres actions">⋯</button>` : ''}</div>
          ${more.length ? `<div class="tkacts tkmorebox" id="kMoreBox" hidden>${more.join('')}</div>` : ''}`;
      })() : ''}
      <label>Titre</label><input id="kTitle" value="${esc(d.title)}" ${ro}>
      ${edit ? lblMic('Description', 'kNotes') : '<label>Description</label>'}<textarea id="kNotes" ${ro}>${esc(d.notes || '')}</textarea>
      <div class="grid2">
        <div><label>Client</label><select id="kCl" ${ro}><option value="">🔒 Privé (sans client)</option>${clientOptions(d.clientId)}</select></div>
        <div><label>Technicien</label><select id="kAs" ${ro}><option value="">— à prendre —</option>
          ${liveUsers().filter(x => x.active !== false || x.id === d.assigneeId).map(x => `<option value="${x.id}" ${x.id === d.assigneeId ? 'selected' : ''}>${esc(userLabel(x))}</option>`).join('')}</select></div>
      </div>
      <div class="grid2">
        <div><label>Échéance</label><input type="date" id="kDue" value="${esc(d.due || '')}" ${ro}></div>
        <div><label>Heure</label><input type="time" id="kTime" value="${esc(d.time || '')}" ${ro}></div>
      </div>
      ${edit ? `<div class="row" style="flex-wrap:wrap;margin-top:6px">${[['0', "Aujourd'hui"], ['1', 'Demain'], ['mon', 'Lundi'], ['7', 'Dans 1 sem.'], ['', 'Sans date']]
        .map(([k, l]) => `<button class="btn chip" type="button" data-kd="${k}">${l}</button>`).join('')}</div>` : ''}
      <label>Priorité</label><input type="hidden" id="kPr" value="${+d.priority || 4}">
      <div class="row prio" style="flex-wrap:wrap">${[1, 2, 3, 4].map(p => `<button class="btn chip pr${p} ${(+d.priority || 4) === p ? 'active' : ''}" type="button" data-kp="${p}" ${ro}>${PRIO[p]}</button>`).join('')}</div>
      <label>Points à cocher</label><ul class="list compact" id="kList"></ul>
      ${edit ? `<div class="row"><input id="kItem" class="grow" style="flex:1" placeholder="Ajouter un point…"><button class="btn" id="kAdd" type="button">+</button></div>` : ''}
      ${t ? `<label>Commentaires</label><ul class="list compact">${(t.comments || []).map(c => `<li><div class="grow" style="flex:1;min-width:0">${esc(c.text)}
          <div class="small muted">${c.by === 'outlook' ? '📧 Assistant de boîte mail' : esc(userName(c.by))} ·${new Date(c.at).toLocaleString('fr-BE', { dateStyle: 'short', timeStyle: 'short' })}</div></div></li>`).join('') || '<li class="small muted">Aucun commentaire.</li>'}</ul>
        <div class="row"><input id="kCom" class="grow" style="flex:1" placeholder="Ajouter un commentaire…"><button class="btn" id="kComAdd" type="button">Ajouter</button></div>` : ''}
      ${entries.length ? `<label>Interventions liées</label><ul class="list compact">${entries.map(e => `<li class="tap" data-tke="${e.id}"><span class="grow">${frDate(e.date)} · ${esc(client(e.clientId).name)} · ${esc(user(techOf(e)).name)}</span>
          <span class="dur">${fmtDur(e.minutes)}</span></li>`).join('')}</ul>` : ''}
      <div class="actions" style="flex-wrap:wrap">
        ${!t && edit ? `<button class="btn" id="kOut" type="button" title="Rendez-vous de 1 h à l'heure de la tâche, sinon journée entière" style="margin-right:auto">📅 Agenda Outlook</button>` : ''}
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
    // Priorité : boutons (une liste déroulante était écrasée sur téléphone).
    $$('[data-kp]', f).forEach(b => b.onclick = () => {
      $('#kPr').value = b.dataset.kp;
      $$('[data-kp]', f).forEach(x => x.classList.toggle('active', x === b));
    });
    // Raccourcis d'échéance.
    $$('[data-kd]', f).forEach(b => b.onclick = () => {
      const k = b.dataset.kd;
      $('#kDue').value = k === '' ? '' : addDays(k === 'mon' ? inDays(1) : +k);
      if (k === '') $('#kTime').value = '';
    });
    // Valeurs du formulaire → tâche (sans enregistrer). Un point à cocher tapé mais pas encore ajouté (« + ») est gardé.
    const read = () => ({ title: $('#kTitle').value.trim().replace(/\s+/g, ' '), notes: $('#kNotes').value.trim(), clientId: $('#kCl').value || null,
      assigneeId: $('#kAs').value || null, due: $('#kDue').value, time: $('#kTime').value, priority: +$('#kPr').value, checklist: cl });
    const store = async () => {
      const pending = $('#kItem')?.value.trim();
      if (pending) { cl.push({ id: uid('ck'), text: pending, done: false }); $('#kItem').value = ''; }
      const v = read();
      if (!v.title) { alert('Titre requis.'); return null; }
      if (v.time && !v.due) v.due = today();
      // Privée : jamais « à prendre », elle reste à son créateur si personne n'est choisi.
      if (!v.clientId && !v.assigneeId) v.assigneeId = t?.createdBy || me().id;
      let x = t;
      if (!x) T.tasks.push(x = { id: uid('tk'), createdBy: me().id, createdAt: Date.now(), status: 'open', entryIds: [], comments: [] });
      const as = v.assigneeId; delete v.assigneeId;
      Object.assign(x, v); assign(x, as); touch(x);
      await save(); scheduleSync();
      return x;
    };
    // Nouvelle tâche : on affiche le filtre où elle se trouve et elle clignote brièvement (sinon, sans date, elle semblait perdue).
    if ($('#kSave')) $('#kSave').onclick = async ev => {
      ev.preventDefault();
      const x = await store();
      if (!x) return;
      dlg.close();
      if (!t && tab === 'taches') {
        setFilter(FILTER_OF(x)); fClient = ''; fTech = ''; q = ''; flashId = x.id;
        setTimeout(() => { flashId = null; }, 2500);
        toast(`Tâche ajoutée · ${FILTER_NAME[filter]}`);
      } else toast('Tâche enregistrée');
      render();
      if (!t) $('[data-tk].flash')?.scrollIntoView({ block: 'center', behavior: 'smooth' });
    };
    if (!t && edit) setTimeout(() => { const el = $('#kTitle'); el.focus(); el.setSelectionRange(el.value.length, el.value.length); }, 50);
    if ($('#kComAdd')) {
      const addCom = async () => {
        const v = $('#kCom').value.trim(); if (!v) return;
        t.comments = [...(t.comments || []), { id: uid('cm'), by: me().id, at: Date.now(), text: v }]; touch(t);
        await save(); scheduleSync(); editTask(t);
      };
      $('#kComAdd').onclick = addCom;
      $('#kCom').onkeydown = ev => { if (ev.key === 'Enter') { ev.preventDefault(); addCom(); } };
    }
    if ($('#kMore')) $('#kMore').onclick = () => {
      const box = $('#kMoreBox'), open = box.hidden;
      box.hidden = !open; $('#kMore').setAttribute('aria-expanded', open); $('#kMore').classList.toggle('active', open);
    };
    if ($('#kTake')) $('#kTake').onclick = async () => { assign(t, me().id); touch(t); await save(); scheduleSync(); toast('Tâche attribuée'); editTask(t); };
    if ($('#kDel')) $('#kDel').onclick = async () => {
      if (!confirm(`Supprimer la tâche « ${t.title} » ?${t.entryIds?.length ? '\n(les interventions liées sont conservées)' : ''}`)) return;
      t.deleted = true; touch(t); dlg.close(); await changed(); toast('Tâche supprimée');
    };
    // Chrono / intervention / terminer : les modifications en cours sont d'abord enregistrées.
    if ($('#kOut')) $('#kOut').onclick = async () => {
      if (!$('#kDue').value) return alert("Ajoute d'abord une échéance (date) à la tâche.");
      const x = await store(); if (x) toOutlook(x);
    };
    if ($('#kTimer')) $('#kTimer').onclick = async () => { const x = await store(); if (x) startTimer(x); };
    if ($('#kEntry')) $('#kEntry').onclick = async () => { const x = await store(); if (x) addEntry(x); };
    if ($('#kDone')) $('#kDone').onclick = async () => { const x = await store(); if (x) finish(x); };
    if ($('#kReopen')) $('#kReopen').onclick = () => reopen(t);
    $$('[data-tke]', f).forEach(li => li.onclick = () => openEntry(state.entries.find(e => e.id === li.dataset.tke)));
    if (!dlg.open) dlg.showModal();
    if (!t) $('#kTitle').focus();
  }

  // --- agenda Outlook (taches.md, section 7) ---
  // Dans l'agenda de la personne connectée : rendez-vous de 1 h à l'heure de la tâche, sinon journée entière.
  // L'événement est gardé par personne (t.outlook[userId]) : cliquer à nouveau le met à jour (date, titre…) au lieu d'en créer un autre.
  // Une tâche dans l'agenda de son technicien est « planifiée » : pas d'e-mail de rappel la veille (serveur, cron.php).
  const TZ = 'Romance Standard Time';   // Bruxelles, nom attendu par Outlook
  const localDT = d => `${isoDate(d)}T${pad(d.getHours())}:${pad(d.getMinutes())}:00`;
  async function toOutlook(t) {
    if (!t.due) return alert("Ajoute d'abord une échéance (date) à la tâche.");
    const c = t.clientId && state.clients.find(x => x.id === t.clientId);
    const start = new Date(`${t.due}T${t.time || '00:00'}`);
    const end = t.time ? new Date(start.getTime() + 3600e3) : new Date(new Date(`${t.due}T12:00`).getTime() + 864e5);
    const ev = {
      subject: t.title + (c ? ` – ${c.name}` : ''),
      body: { contentType: 'text', content: [t.notes, ...(t.checklist || []).map(x => (x.done ? '☑ ' : '☐ ') + x.text),
        'Tâche LevelDesk : ' + location.origin + location.pathname].filter(Boolean).join('\n') },
      start: { dateTime: t.time ? localDT(start) : `${t.due}T00:00:00`, timeZone: TZ },
      end: { dateTime: t.time ? localDT(end) : `${isoDate(end)}T00:00:00`, timeZone: TZ },
      isAllDay: !t.time, showAs: t.time ? 'busy' : 'free'
    };
    const known = t.outlook?.[me().id];
    try {
      let r = known ? await OneDrive.calendar('PATCH', '/me/events/' + encodeURIComponent(known), ev) : null;
      if (!r || r.status === 404) r = await OneDrive.calendar('POST', '/me/events', ev);   // supprimé dans Outlook : recréé
      const j = await r.json().catch(() => ({}));
      if (!r.ok) throw new Error(/MailboxNotEnabled/i.test(j.error?.code || '') ? "Ton compte Microsoft n'a pas de boîte Outlook : pas d'agenda disponible."
        : `Outlook (${r.status}) : ${j.error?.message || r.statusText}`);
      t.outlook = { ...(t.outlook || {}), [me().id]: j.id }; touch(t);
      await save(); scheduleSync();
      toast(known ? 'Rendez-vous mis à jour dans Outlook' : 'Tâche ajoutée à ton agenda Outlook');
      editTask(t);
    } catch (e) {
      if (e.consent) {
        if (confirm("Pour ajouter tes tâches à ton agenda, LevelDesk doit y avoir accès (page Microsoft, une seule fois).\nLa tâche est déjà enregistrée : reclique sur « Agenda Outlook » au retour.\n\nContinuer ?")) OneDrive.calendarConsent();
        return;
      }
      alert(e instanceof OneDrive.NeedLogin ? 'Session Microsoft expirée : reconnecte-toi.' : e.message);
    }
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
  // Mise à jour serveur : forfait « Mise à jour serveur » du client (durée selon le nombre de serveurs) ; enregistré, il termine la tâche (forfaitSaved).
  const addEntry = t => t.kind === 'upd' ? openForfait(null, 'update', { clientId: t.clientId, taskId: t.id })
    : openEntry(null, { ...(t.clientId ? { clientId: t.clientId } : {}), description: t.title, taskId: t.id });
  async function finish(t) {
    if (t.kind === 'upd' && !t.entryIds?.length) return addEntry(t);
    Object.assign(t, { status: 'done', doneAt: Date.now(), doneBy: me().id }); touch(t);
    await save(); scheduleSync();
    // Tâche privée : pas d'intervention à proposer.
    if (isPrivate(t)) { if ($('#dlg').open) $('#dlg').close(); toast('Tâche terminée'); if (tab === 'taches') view(); return; }
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
    $('.notifbar', el).onclick = () => { setFilter(n.now || n.late || n.fresh ? 'me' : 'free'); q = ''; tab = 'taches'; render(); };
  }

  // --- tableau de bord PC (app.js : <div id="tkDay">) : mes tâches du jour et en retard ---
  function dayList(el) {
    if (!T) { el.innerHTML = '<div class="empty">Chargement…</div>'; load().then(() => el.isConnected && dayList(el)); return; }
    const u = me(), ts = alive().filter(visible);
    const list = ts.filter(t => t.assigneeId === u.id && (overdue(t) || dueToday(t))).sort((a, b) => sortKey(a).localeCompare(sortKey(b)));
    const free = ts.filter(canTake).length;
    el.innerHTML = (list.length ? `<ul class="list tklist">${list.map(taskLi).join('')}</ul>` : '<div class="empty">Rien d\'urgent 👍</div>')
      + (free ? `<p class="small" style="margin:6px 0 0"><a href="#" id="tkDayFree">${free} tâche(s) à prendre ›</a></p>` : '');
    bindRows(el, () => el.isConnected && dayList(el));
    if ($('#tkDayFree', el)) $('#tkDayFree', el).onclick = ev => { ev.preventDefault(); setFilter('free'); q = ''; tab = 'taches'; render(); };
  }

  async function reset() { T = null; msg = ''; olMsg = ''; olAt = 0; await idb('readwrite', s => s.delete('data')); }
  // Pour la synchro du serveur (app.js) : liste des tâches, puis enregistrement local après réception.
  async function raw() { return (await load()).tasks; }
  async function setRaw(tasks) { await load(); T.tasks = tasks; await save(); }

  return { view, sync, todayBanner, dayList, linkEntry, forfaitSaved, titleOf, open, reset, parseQuick, raw, setRaw, afterSync };
})();
