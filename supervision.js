// © 2026 One Level UP – LevelDesk. Logiciel propriétaire, tous droits réservés. Copie, modification, redistribution et rétro-ingénierie interdites sans autorisation écrite (voir LICENCE.txt).
// Module Supervision (supervision.md) : serveurs des clients suivis par un agent en lecture seule, installé sur chaque serveur.
// - L'agent envoie chaque jour son export au serveur LevelDesk (?r=agent, jeton propre au serveur) ; cet onglet lit le résultat (?r=supervision).
// - EN LIGNE SEULEMENT : rien n'est gardé dans la copie hors ligne de l'app ; les données sont relues à l'ouverture de l'onglet.
// - Tout texte venant d'un export passe par esc() : un export est une donnée reçue d'une machine tierce, jamais du HTML.
// - Administrateur : fiches serveur, script d'installation (nouveau jeton à chaque fois), révocation, retrait, reprise de clients.json.
// Utilise les outils de app.js ($, $$, esc, uid, state, client, activeClients, isAdmin, mod, srvOn, toast, download, frDate) : chargé avant lui, appelé après le démarrage.
const Supervision = (() => {
  let S = null;                 // dernière réponse de « etat » : { agentVersion, servers }
  let at = 0, busy = false, err = '';
  let flt = '';                 // filtre de l'onglet (clic sur un compteur) : '', ok, warn, crit, mute
  const MUTE_H = 36;            // sans contact depuis plus de 36 h : signalé (la vraie alerte « serveur muet » viendra du serveur, étape 2)
  const ROLES = ['HyperV', 'AD', 'DNS', 'DHCP', 'Files', 'Print', 'RDS'];
  const BACKUPS = ['', 'BackupAssist', 'Windows Server Backup', 'Veeam'];

  const call = body => Serveur.json('supervision', { method: 'POST', body });
  const on = () => srvOn() && mod('supervision');
  const hoursSince = iso => iso ? (Date.now() - Date.parse(iso)) / 3600000 : Infinity;
  const muted = s => !!s.lastSeen && hoursSince(s.lastSeen) > MUTE_H;
  function ago(iso) {
    if (!iso) return 'jamais vu';
    const h = hoursSince(iso);
    if (h < 1) return 'vu il y a ' + Math.max(1, Math.round(h * 60)) + ' min';
    if (h < 48) return 'vu il y a ' + Math.round(h) + ' h';
    return 'vu il y a ' + Math.round(h / 24) + ' j';
  }
  const lvCls = l => ({ Critique: 'crit', Attention: 'warn', OK: 'ok' }[l] || 'none');
  // NAS (14/11) : pas d'agent ; état = pire statut du contrôle du mois saisi à la main (mêmes rubriques que le rapport de maintenance).
  const isNas = s => s?.data?.kind === 'nas';
  const NAS_CHECKS = [['jobs', 'Sauvegardes / réplications (dernier job + état)'], ['restore', 'Test de restauration depuis le NAS'], ['espace', 'Espace disque (libre / total)'],
    ['disques', 'État des disques (SMART, RAID / SHR)'], ['maj_dsm', 'Mises à jour du système (DSM : dernière / disponible)'], ['paquets', 'Paquets (Package Center) à jour']];
  const NAS_ST = ['OK', 'Attention', 'Critique', 'N/A'];
  const checkOf = s => s?.data?.check || {};
  const checkAge = s => checkOf(s).date ? (Date.now() - Date.parse(checkOf(s).date)) / 86400000 : Infinity;
  const NAS_OLD_D = 40;          // contrôle de plus de 40 jours : signalé « à refaire »
  // Catégorie d'un serveur pour les compteurs et les filtres : ok, warn, crit, mute (sans contact / jamais vu / NAS pas contrôlé).
  const catOf = s => isNas(s) ? (checkOf(s).date || Object.keys(checkOf(s).items || {}).length ? (lvCls(s.level) === 'none' ? 'mute' : lvCls(s.level)) : 'mute')
    : !s.lastSeen || muted(s) ? 'mute' : lvCls(s.level) === 'none' ? 'mute' : lvCls(s.level);
  const lvPill = s => isNas(s) ? (!s.level ? '<span class="pill">pas encore contrôlé</span>'
      : `<span class="pill supv-${lvCls(s.level)}">${esc(s.level === 'Attention' ? 'À surveiller' : s.level)}</span>${checkAge(s) > NAS_OLD_D ? ' <span class="pill supv-mute">contrôle à refaire</span>' : ''}`)
    : !s.lastSeen ? '<span class="pill">jamais vu</span>'
    : muted(s) ? '<span class="pill supv-mute">sans contact</span>'
    : `<span class="pill supv-${lvCls(s.level)}">${esc(s.level === 'Attention' ? 'À surveiller' : s.level || '—')}</span>`;
  const cname = id => state.clients.find(c => c.id === id)?.name || '(client)';
  const num = (v, unit = '') => v === null || v === undefined || v === '' ? '—' : esc(String(v).replace('.', ',')) + unit;

  function dialog(html, large) {
    const d = $('#dlg');
    d.classList.toggle('large', !!large);
    $('#dlgForm').innerHTML = html;
    if (!d.open) { d.showModal(); d.addEventListener('close', () => d.classList.remove('large'), { once: true }); }
    return $('#dlgForm');
  }

  async function load(force) {
    if (busy || (!force && S && Date.now() - at < 60000)) return;
    busy = true; err = '';
    try { S = await call({ op: 'etat' }); at = Date.now(); }
    catch (e) { err = e.message || String(e); }
    busy = false;
    if (typeof tab !== 'undefined' && tab === 'serveurs' && !editing()) draw();
  }

  // ---------- onglet ----------
  // page : null = tous les serveurs ; { cli } = page d'un client (serveurs, maintenances, rapports) ; { cli, month } = rapport de maintenance.
  let page = null;
  const editing = () => !!(page?.month && $('#rpForm'));      // rapport en cours de saisie : jamais redessiné par une synchro
  function view() {
    if (editing()) return;
    draw();
    load(false);
  }
  const goAll = () => { page = null; flt = ''; draw(); window.scrollTo(0, 0); };
  const goClient = cid => { page = { cli: cid }; R = null; draw(); window.scrollTo(0, 0); };

  const MOIS = ['janvier', 'février', 'mars', 'avril', 'mai', 'juin', 'juillet', 'août', 'septembre', 'octobre', 'novembre', 'décembre'];
  const monthName = m => `${MOIS[+m.slice(5, 7) - 1] || '?'} ${m.slice(0, 4)}`;
  const deMonth = m => (/^[aeiouhéèêœ]/i.test(monthName(m)) ? 'd\'' : 'de ') + monthName(m);
  // Interventions « Maintenance serveur » d'un client (forfait de app.js), les plus récentes d'abord.
  const maintOf = (cid, month) => live().filter(e => e.forfait === 'maint' && e.clientId === cid && (!month || e.date.startsWith(month))).sort((a, b) => b.date.localeCompare(a.date));
  function addMaint(cid) {
    const done = maintOf(cid, today().slice(0, 7))[0];
    done ? (done.billId || done.locked ? viewArchivedEntry(done) : openForfait(done)) : openForfait(null, 'maint', { clientId: cid });
  }

  // Ligne d'un serveur ou d'un NAS dans une liste (onglet, page du client).
  function srvLine(s, withClient) {
    const old = s.agentVersion && s.agentVersion !== S.agentVersion;
    const c = s.counts || {}, d = s.data || {}, ck = checkOf(s);
    const sub = isNas(s) ? ['NAS', [d.brand, d.model].filter(Boolean).join(' '), d.role, ck.date ? 'contrôlé le ' + frDate(ck.date) : 'aucun contrôle saisi'].filter(Boolean).map(esc).join(' · ')
      : `${s.lastSeen ? esc(ago(s.lastSeen)) : s.hasToken ? 'agent prévu, aucun envoi reçu' : 'pas d\'agent installé'}${s.agentVersion ? ' · agent ' + esc(s.agentVersion) + (old ? ' <span class="supv-old">(à mettre à jour : ' + esc(S.agentVersion) + ')</span>' : '') : ''}${s.lastSeen && !s.hasToken ? ' · jeton révoqué' : ''}`;
    return `<li class="tap" data-srv="${esc(s.id)}"><div class="grow"><strong>${isNas(s) ? '🗄 ' : ''}${esc(s.name)}</strong>${withClient ? ` <span class="muted">· ${esc(cname(s.clientId))}</span>` : ''} ${lvPill(s)}
      <div class="small muted">${sub}</div></div>
      <div class="dur small">${c.Critique ? `<span class="pill supv-crit">${c.Critique} critique${c.Critique > 1 ? 's' : ''}</span> ` : ''}${c.Attention ? `<span class="pill supv-warn">${c.Attention} à surveiller</span>` : ''}</div></li>`;
  }

  function draw() {
    const el = $('#view');
    if (!on()) { el.innerHTML = '<div class="card"><p class="empty">Module Supervision non activé pour ta société.</p></div>'; return; }
    if (S && page?.month) return drawReport();
    if (S && page?.cli) return drawClient();
    if (!S) {
      el.innerHTML = `<div class="card"><p class="empty">${err ? 'Serveurs : consultation en ligne seulement.<br><span class="small">' + esc(err) + '</span>' : 'Chargement…'}</p>
        ${err ? '<div class="actions"><button class="btn primary" id="svRetry">Réessayer</button></div>' : ''}</div>`;
      if ($('#svRetry')) $('#svRetry').onclick = () => { el.querySelector('.empty').textContent = 'Chargement…'; load(true); };
      return;
    }
    const all = S.servers;
    const n = { ok: 0, warn: 0, crit: 0, mute: 0 };
    all.forEach(s => n[catOf(s)]++);
    const FLT = { ok: 'OK', warn: 'À surveiller', crit: 'Critiques', mute: 'Sans contact ou pas encore contrôlés' };
    if (flt && !n[flt]) flt = '';
    const line = srvLine;
    // Filtre (clic sur un compteur) : une seule liste des serveurs concernés, client indiqué, les critiques d'abord par nombre d'alertes.
    const shown = flt ? all.filter(s => catOf(s) === flt).sort((a, b) => ((b.counts?.Critique || 0) - (a.counts?.Critique || 0)) || ((b.counts?.Attention || 0) - (a.counts?.Attention || 0))
      || cname(a.clientId).localeCompare(cname(b.clientId)) || a.name.localeCompare(b.name)) : all;
    const byClient = new Map();
    shown.forEach(s => { if (!byClient.has(s.clientId)) byClient.set(s.clientId, []); byClient.get(s.clientId).push(s); });
    const groups = [...byClient.entries()].sort((a, b) => cname(a[0]).localeCompare(cname(b[0])));
    const kpi = (k, cls, label) => `<button type="button" class="kpi supv-kpi ${flt === k ? 'on' : ''}" data-flt="${k}" ${n[k] ? '' : 'disabled'} title="${n[k] ? 'Voir la liste' : ''}">
      <div class="v ${cls}">${n[k]}</div><div class="l">${label}</div></button>`;
    // Pied de la carte d'un client : maintenance du mois (forfait « Maintenance serveur ») et accès à sa page (maintenances, rapports).
    const month = today().slice(0, 7);
    const foot = cid => { const m = mod('forfaits') ? maintOf(cid, month)[0] : null;
      return `<div class="supv-foot">${mod('forfaits') ? `<span class="small ${m ? 'supv-t-ok' : 'muted'}">🛠 Maintenance ${esc(deMonth(month))} : ${m ? `faite le ${esc(frDate(m.date))} (${fmtDur(m.minutes)})` : 'pas encore encodée'}</span>` : '<span></span>'}
        <button type="button" class="btn link" data-cli="${esc(cid)}">${isAdmin() ? 'Maintenance et rapports' : 'Détail du client'} ›</button></div>`; };
    el.innerHTML = `${err ? `<div class="notifbar"><span class="ico">⚠</span><span class="txt">Données non actualisées : ${esc(err)}</span></div>` : ''}
      <div class="kpis">
        <button type="button" class="kpi supv-kpi ${flt ? '' : 'on'}" data-flt="" title="Tous les serveurs, par client"><div class="v">${all.length}</div><div class="l">Tous les appareils</div></button>
        ${kpi('ok', 'supv-t-ok', 'OK')}${kpi('warn', 'supv-t-warn', 'À surveiller')}${kpi('crit', 'supv-t-crit', 'Critiques')}
        ${kpi('mute', 'supv-t-mute', `Sans contact (> ${MUTE_H} h ou jamais vu) · NAS pas contrôlé`)}
      </div>
      ${flt ? `<div class="row supv-back"><button type="button" class="btn" id="svAll">‹ Tous les serveurs</button></div>
        <div class="card"><h2>${esc(FLT[flt])}<span class="tot">${shown.length} appareil${shown.length > 1 ? 's' : ''}</span></h2>
          <ul class="list">${shown.map(s => line(s, true)).join('')}</ul></div>`
        : groups.length ? groups.map(([cid, list]) => `<div class="card"><h2 class="tap" data-cli="${esc(cid)}" title="Ouvrir la page du client">${esc(cname(cid))}<span class="tot">${list.length} appareil${list.length > 1 ? 's' : ''}</span></h2>
        <ul class="list">${list.map(s => line(s, false)).join('')}</ul>${foot(cid)}</div>`).join('')
        : `<div class="card"><p class="empty">Aucun serveur supervisé.${isAdmin() ? '<br><span class="small">Ajoute un serveur, puis installe son agent, ou ajoute un NAS.</span>' : ''}</p></div>`}
      <div class="row" style="margin:4px 0 16px">
        <button class="btn" id="svRefresh">↻ Actualiser</button>
        ${isAdmin() ? '<button class="btn primary" id="svAdd">+ Serveur</button><button class="btn" id="svAddNas">+ NAS</button><button class="btn" id="svImport" style="margin-left:auto">Importer clients.json</button>' : ''}
      </div>`;
    $$('[data-srv]', el).forEach(li => li.onclick = () => openServer(li.dataset.srv));
    $$('[data-flt]', el).forEach(b => b.onclick = () => { flt = flt === b.dataset.flt ? '' : b.dataset.flt; draw(); });
    $$('[data-cli]', el).forEach(b => b.onclick = () => goClient(b.dataset.cli));
    if ($('#svAll')) $('#svAll').onclick = goAll;
    $('#svRefresh').onclick = () => { toast('Actualisation…'); load(true); };
    if ($('#svAdd')) $('#svAdd').onclick = () => editServer(null);
    if ($('#svAddNas')) $('#svAddNas').onclick = () => editNas(null);
    if ($('#svImport')) $('#svImport').onclick = importClients;
  }

  // ---------- fiche d'un serveur ----------
  async function openServer(id) {
    const f = dialog('<h2>Serveur</h2><p class="empty">Chargement…</p><div class="actions"><button class="btn" value="cancel">Fermer</button></div>', true);
    let d;
    try { d = await call({ op: 'serveur', id }); }
    catch (e) { f.querySelector('.empty').innerHTML = 'Lecture impossible.<br><span class="small">' + esc(e.message) + '</span>'; return; }
    if (isNas(d.server)) return nasView(d.server);
    const s = d.server, r = d.report, m = r?.summary || {}, sy = m.system || {}, hw = m.hardware;
    const row = (l, v) => v === '' || v === null || v === undefined ? '' : `<tr><td>${esc(l)}</td><td>${v}</td></tr>`;
    const t = v => v ? esc(v) : '';
    const order = { Critique: 0, Attention: 1, Info: 2 };
    const alerts = [...(m.alerts || [])].sort((a, b) => order[a.level] - order[b.level]);
    const vols = (m.volumes || []).map(v => {
      const used = v.sizeGB ? Math.max(0, Math.min(100, 100 - (v.freePct ?? 0))) : 0;
      const cls = v.freePct < 10 ? 'crit' : v.freePct < 20 ? 'warn' : 'ok';
      return `<div class="supv-vol"><div class="small"><strong>${esc(v.drive)}</strong> ${esc(v.label)} <span class="muted">· ${num(v.freeGB, ' Go')} libres sur ${num(v.sizeGB, ' Go')} (${num(v.freePct, ' %')})</span></div>
        <div class="bar"><i class="supv-b-${cls}" style="width:${used}%"></i></div></div>`;
    }).join('');
    f.innerHTML = `<h2>${esc(s.name)} <span class="small muted">· ${esc(cname(s.clientId))}</span> ${lvPill(s)}</h2>
      <p class="small muted" style="margin:0 0 10px">${s.lastSeen ? esc(ago(s.lastSeen)) : s.hasToken ? 'Agent prévu, aucun envoi reçu' : 'Pas d\'agent installé'}${r ? ' · collecte du ' + esc(r.collectedAt || frDate(r.day)) + (m.meta?.durationSec ? ' (' + esc(Math.round(m.meta.durationSec / 60)) + ' min)' : '') : ''}
        ${m.meta && m.meta.runAsAdmin === false ? ' · <span class="supv-old">collecte sans droits d\'administrateur</span>' : ''}</p>
      ${!r ? `<p class="empty">Aucun export reçu.${s.hasToken ? ' L\'agent est-il installé sur le serveur ?' : ' Installe l\'agent sur ce serveur.'}</p>` : `
      ${alerts.length ? `<div class="supv-sec">Points d'attention</div><ul class="list supv-alerts">${alerts.map(a => `<li><span class="pill supv-${lvCls(a.level === 'Info' ? '' : a.level)}">${esc(a.level)}</span>
        <div class="grow"><strong>${esc(a.category)}</strong> · ${esc(a.message)}${a.action ? `<div class="small muted">→ ${esc(a.action)}</div>` : ''}</div></li>`).join('')}</ul>`
        : '<p class="small supv-t-ok">Aucun point d\'attention dans le dernier export.</p>'}
      ${(m.checkErrors || []).length ? `<p class="small supv-old">Contrôles en erreur pendant la collecte : ${esc(m.checkErrors.join(', '))}</p>` : ''}
      <div class="supv-sec">Inventaire</div>
      <table class="tot ro">
        ${row('Ordinateur', t(sy.computerName) + (sy.domain ? ' <span class="muted">· ' + esc(sy.domain) + '</span>' : ''))}
        ${row('Système', t(sy.osCaption) + (sy.osBuild ? ' <span class="muted">(build ' + esc(sy.osBuild) + ')</span>' : ''))}
        ${row('Fin de support', sy.osEndOfSupport ? esc(frDate(sy.osEndOfSupport)) : '')}
        ${row('Matériel', [sy.manufacturer, sy.model].filter(Boolean).map(esc).join(' ') + (sy.isVirtual ? ' <span class="muted">(virtuel)</span>' : ''))}
        ${row('N° de série', t(sy.serialNumber))}
        ${row('BIOS / ROM', t(sy.biosVersion))}
        ${row('Processeur', t(sy.cpu) + (sy.cpuCores ? ' <span class="muted">· ' + esc(sy.cpuCores) + ' cœurs</span>' : ''))}
        ${row('Mémoire', sy.ramGB ? num(sy.ramGB, ' Go') + (m.perf?.memUsedPct != null ? ' <span class="muted">· utilisée à ' + num(m.perf.memUsedPct, ' %') + '</span>' : '') : '')}
        ${row('Adresses IP', (sy.ips || []).map(esc).join(', '))}
        ${row('Dernier démarrage', sy.lastBoot ? esc(sy.lastBoot) + (sy.uptimeDays != null ? ' <span class="muted">(' + num(sy.uptimeDays, ' j') + ')</span>' : '') : '')}
        ${hw ? row('SPP HPE', t(hw.sppDetected)) + row('iLO', t(hw.iloVersion || hw.iloFirmware)) + row('RAID', t(hw.raid)) : ''}
      </table>
      ${vols ? `<div class="supv-sec">Volumes</div>${vols}` : ''}
      ${upsBlock(m.ups)}
      <div class="supv-sec">État</div>
      <table class="tot ro">
        ${row('Sauvegarde', (m.backup?.expected ? esc(m.backup.expected) + ' · ' : '') + (m.backup?.lastSuccess ? 'dernière faite le ' + esc(m.backup.lastSuccess) + (m.backup.withWarnings ? ' <span class="supv-old">(avec avertissements)</span>' : '') : (m.backup?.expected ? 'aucune réussie détectée' : '<span class="muted">pas de sauvegarde à contrôler sur ce serveur</span>')))}
        ${row('Mises à jour', (m.updates ? esc(m.updates.pending) + ' en attente' + (m.updates.rebootPending ? ' · <strong>redémarrage en attente</strong>' : '') + (m.updates.lastSecurityDate ? ' <span class="muted">· dernière cumulative ' + esc(m.updates.lastSecurityDate) + '</span>' : '') : ''))}
        ${row('Antivirus', [m.av?.eset, m.av?.defender].filter(Boolean).map(esc).join(' · ') || '<span class="muted">non détecté</span>')}
        ${row('Événements (31 j)', m.events ? `${num(m.events.sysCritical)} critiques · ${num(m.events.sysErrors)} erreurs système · ${num(m.events.appErrors)} erreurs applications · ${num(m.events.unexpectedShutdowns)} arrêt(s) inattendu(s)` : '')}
      </table>`}
      ${s.data?.description || (s.data?.roles || []).length ? `<p class="small muted" style="margin-top:10px">Fiche : ${esc([s.data.type, (s.data.roles || []).join(', '), s.data.description, s.data.backup].filter(Boolean).join(' · '))}</p>` : ''}
      <div class="actions supv-actions">
        ${isAdmin() ? `<button type="button" class="btn danger" id="svDel" style="margin-right:auto">Retirer</button>
          <button type="button" class="btn" id="svEdit">Modifier</button>
          ${s.hasToken ? '<button type="button" class="btn" id="svRevoke">Révoquer le jeton</button>' : ''}
          <button type="button" class="btn" id="svInstall">Installer l'agent</button>` : ''}
        ${r ? '<button type="button" class="btn" id="svExport">Export complet (JSON)</button>' : ''}
        <button class="btn" value="cancel">Fermer</button>
      </div>`;
    if ($('#svExport')) $('#svExport').onclick = async () => {
      try {
        const x = await call({ op: 'export', id });
        download(new Blob([JSON.stringify(x.export, null, 2)], { type: 'application/json' }), `${cname(s.clientId)}_${s.name}_${x.day}.json`.replace(/[^\w.\-]+/g, '_'));
      } catch (e) { toast(e.message); }
    };
    if ($('#svEdit')) $('#svEdit').onclick = () => editServer(s);
    if ($('#svInstall')) $('#svInstall').onclick = () => installAgent(s);
    if ($('#svRevoke')) $('#svRevoke').onclick = () => confirmBox(`Révoquer le jeton de ${s.name} ?`,
      'L\'agent installé sur ce serveur ne pourra plus rien envoyer tant qu\'il n\'est pas réinstallé.', 'Révoquer', async () => {
        await call({ op: 'serveur.jeton', id, revoke: true }); toast('Jeton révoqué'); await load(true); openServer(id);
      }, () => openServer(id));
    if ($('#svDel')) $('#svDel').onclick = () => confirmBox(`Retirer ${s.name} de la supervision ?`,
      'Son jeton est supprimé. L\'agent reste sur le serveur tant qu\'il n\'est pas désinstallé (Desinstaller-LevelDesk.ps1), mais ses envois seront refusés.', 'Retirer', async () => {
        await call({ op: 'serveur.delete', id }); $('#dlg').close(); toast('Serveur retiré'); load(true);
      }, () => openServer(id));
  }

  // Onduleur (agent 1.3.1, PowerChute Serial Shutdown) : état de la batterie et besoin de la remplacer.
  function upsBlock(u) {
    if (!u) return '';
    const row = (l, v) => v === '' || v === null || v === undefined ? '' : `<tr><td>${esc(l)}</td><td>${v}</td></tr>`;
    const warn = t => `<span class="supv-old">${t}</span>`;
    const st = u.selfTestResult ? (/^(ok|pass(ed)?)$/i.test(u.selfTestResult) ? '<span class="supv-t-ok">Réussi</span>' : `<strong class="supv-t-crit">${esc(u.selfTestResult)} : batterie à remplacer</strong>`) : '';
    const age = u.batteryAgeMonths != null ? `${num(Math.round(u.batteryAgeMonths / 12 * 10) / 10, ' an(s)')} <span class="muted">(depuis ${esc(u.batteryDate)}, ${u.batteryDateSource === 'saisie' ? 'date saisie' : 'date de fabrication'})</span>` : '';
    const ver = u.version ? esc(u.version) + (u.upToDate === false ? ' ' + warn(`(version ${esc(u.latestVersion)} disponible)`) : u.upToDate ? ' <span class="muted">(à jour)</span>' : '') : '';
    return `<div class="supv-sec">Onduleur</div>
      <table class="tot ro">
        ${row('Dernier auto-test', st + (u.lastSelfTest ? ` <span class="muted">· ${esc(u.lastSelfTest)}${u.selfTests ? ' · ' + esc(u.selfTests) + ' sur 31 j' : ''}</span>` : ''))}
        ${row('Capacité de la batterie', u.batteryCapacityPct != null ? (u.batteryCapacityPct < 90 ? warn(num(u.batteryCapacityPct, ' %')) : num(u.batteryCapacityPct, ' %')) + (u.minCapacityPct != null && u.minCapacityPct < u.batteryCapacityPct ? ` <span class="muted">(min. ${num(u.minCapacityPct, ' %')} sur 31 j)</span>` : '') : '')}
        ${row('Âge de la batterie', u.batteryAgeMonths > 48 ? warn(age) : age)}
        ${row('Tension batterie', u.batteryVoltage != null ? num(u.batteryVoltage, ' V') : '')}
        ${row('Température', u.internalTempC != null ? (u.internalTempC > 35 ? warn(num(u.internalTempC, ' °C')) : num(u.internalTempC, ' °C')) : '')}
        ${row('Charge', u.loadPct != null ? num(u.loadPct, ' %') + (u.powerW != null ? ` <span class="muted">· ${num(u.powerW, ' W')}</span>` : '') : '')}
        ${row('Réseau électrique', u.lineVoltage != null ? num(u.lineVoltage, ' V') + (u.onBatteryEvents ? ' · ' + warn(`${esc(u.onBatteryEvents)} passage(s) sur batterie (dernier ${esc(u.lastOnBattery)})`) : ' <span class="muted">· aucun passage sur batterie sur 31 j</span>') : '')}
        ${row('Dernière mesure', u.lastSample ? esc(u.lastSample) + (u.sampleAgeHours > 2 ? ' ' + warn('(liaison perdue ?)') : '') : '')}
        ${row('N° de série', u.serial ? esc(u.serial) + (u.manufactured ? ` <span class="muted">(fabriqué ${esc(u.manufactured)})</span>` : '') : '')}
        ${row('PowerChute', [esc(u.software || ''), ver].filter(Boolean).join(' ') + (u.serviceState && u.serviceState !== 'Running' ? ' ' + warn('service ' + esc(u.serviceState)) : ''))}
        ${u.note ? row('Remarque', esc(u.note)) : ''}
      </table>`;
  }

  // Confirmation dans la fenêtre (pas de confirm() du navigateur : bloqué dans certains navigateurs intégrés).
  function confirmBox(title, text, okLabel, onOk, onBack) {
    const f = dialog(`<h2>${esc(title)}</h2><p>${esc(text)}</p><p class="small supv-old hidden" id="svCErr"></p>
      <div class="actions"><button type="button" class="btn" id="svCBack">Retour</button><button type="button" class="btn primary" id="svCOk">${esc(okLabel)}</button></div>`);
    $('#svCBack', f).onclick = onBack;
    $('#svCOk', f).onclick = async ev => {
      ev.target.disabled = true;
      try { await onOk(); } catch (e) { ev.target.disabled = false; const p = $('#svCErr', f); if (p) { p.textContent = e.message; p.classList.remove('hidden'); } }
    };
  }

  // ---------- création / modification ----------
  function editServer(s) {
    const d = s?.data || {};
    const clients = activeClients().filter(c => c.server || c.id === s?.clientId);
    const list = clients.length ? clients : activeClients();
    const f = dialog(`<h2>${s ? 'Modifier le serveur' : 'Nouveau serveur'}</h2>
      <label>Client</label><select id="svClient">${list.map(c => `<option value="${esc(c.id)}" ${c.id === s?.clientId ? 'selected' : ''}>${esc(c.name)}</option>`).join('')}</select>
      <label>Nom du serveur <span class="small">(nom de l'ordinateur : lettres, chiffres, point, tiret)</span></label>
      <input id="svName" value="${esc(s?.name || '')}" placeholder="ex. HV01" autocapitalize="characters" autocomplete="off" maxlength="40">
      <label>Type</label><select id="svType"><option value="physique">Physique</option><option value="virtuel" ${d.type === 'virtuel' ? 'selected' : ''}>Virtuel</option></select>
      <label>Rôles attendus</label><div class="row">${ROLES.map(r => `<label class="check"><input type="checkbox" data-role="${r}" ${(d.roles || []).includes(r) ? 'checked' : ''}> ${r}</label>`).join('')}</div>
      <label>Description</label><input id="svDesc" value="${esc(d.description || '')}" maxlength="300">
      <label>Logiciels métier <span class="small">(séparés par une virgule)</span></label><input id="svLob" value="${esc((d.lob || []).join(', '))}">
      <label>Sauvegarde exécutée sur ce serveur</label>
      <select id="svBackup">${[...new Set([...BACKUPS, d.backup || ''])].map(b => `<option value="${esc(b)}" ${b === (d.backup || '') ? 'selected' : ''}>${b ? esc(b) : 'Aucune (rien à contrôler ici)'}</option>`).join('')}</select>
      <label>Date d'installation</label><input type="date" id="svDate" value="${esc(d.dateInstallation || '')}">
      <label>Batterie de l'onduleur remplacée en <span class="small">(mois ; vide = date de fabrication de l'onduleur)</span></label>
      <input type="month" id="svUpsBat" value="${esc(d.upsBatteryDate || '')}">
      <p class="small muted" style="margin:4px 0 0">Prise en compte par l'agent à sa prochaine installation.</p>
      <p class="small supv-old hidden" id="svErr"></p>
      <div class="actions"><button type="button" class="btn" id="svBack">${s ? 'Retour' : 'Annuler'}</button><button type="button" class="btn primary" id="svSave">Enregistrer</button></div>`);
    $('#svBack', f).onclick = () => s ? openServer(s.id) : $('#dlg').close();
    $('#svSave', f).onclick = async ev => {
      const name = $('#svName', f).value.trim();
      const fail = t => { const p = $('#svErr', f); p.textContent = t; p.classList.remove('hidden'); };
      if (!/^[A-Za-z0-9_.-]{1,40}$/.test(name)) return fail('Nom invalide : lettres, chiffres, point, tiret, 40 caractères au plus, sans espace.');
      ev.target.disabled = true;
      const id = s?.id || uid('sv').replace(/[^A-Za-z0-9_-]/g, '');
      try {
        await call({ op: 'serveur.save', id, clientId: $('#svClient', f).value, name, data: {
          type: $('#svType', f).value, roles: $$('[data-role]', f).filter(x => x.checked).map(x => x.dataset.role), description: $('#svDesc', f).value.trim(),
          lob: $('#svLob', f).value.split(',').map(x => x.trim()).filter(Boolean), backup: $('#svBackup', f).value, dateInstallation: $('#svDate', f).value,
          upsBatteryDate: $('#svUpsBat', f).value } });
        toast('Enregistré'); await load(true); openServer(id);
      } catch (e) { ev.target.disabled = false; fail(e.message); }
    };
  }

  // ---------- NAS (pas d'agent : fiche + contrôle du mois saisi à la main, pour le rapport de maintenance) ----------
  const stCls = st => ({ OK: 'ok', Attention: 'warn', Critique: 'crit' }[st] || 'none');
  function nasView(s) {
    const d = s.data || {}, ck = checkOf(s), it = ck.items || {};
    const row = (l, v) => v ? `<tr><td>${esc(l)}</td><td>${esc(v)}</td></tr>` : '';
    const f = dialog(`<h2>🗄 ${esc(s.name)} <span class="small muted">· ${esc(cname(s.clientId))}</span> ${lvPill(s)}</h2>
      <div class="supv-sec">Fiche</div>
      <table class="tot ro">
        ${row('Type', 'NAS')}${row('Marque / modèle', [d.brand, d.model].filter(Boolean).join(' '))}${row('Version du système (DSM)', d.dsm)}
        ${row('Rôle', d.role)}${row('Adresse IP', d.ip)}${row('N° de série', d.serial)}${row('Disques', d.disks)}
        ${row('Date d\'installation', d.dateInstallation ? frDate(d.dateInstallation) : '')}${row('Description', d.description)}
      </table>
      ${!d.dsm ? '<p class="small supv-old">Version DSM à compléter (demandée par le rapport de maintenance).</p>' : ''}
      <div class="supv-sec">Contrôle du mois ${ck.date ? `<span class="small muted">· le ${esc(frDate(ck.date))}${ck.by ? ' par ' + esc(user(ck.by).name) : ''}</span>` : ''}</div>
      ${ck.date || Object.keys(it).length ? `<ul class="list supv-alerts">${NAS_CHECKS.map(([k, label]) => { const x = it[k] || {};
          return `<li><span class="pill supv-${stCls(x.st)}">${esc(x.st || 'À compléter')}</span><div class="grow"><strong>${esc(label)}</strong>${x.com ? `<div class="small muted">${esc(x.com)}</div>` : ''}</div></li>`; }).join('')}</ul>
        ${checkAge(s) > NAS_OLD_D ? `<p class="small supv-old">Contrôle de plus de ${NAS_OLD_D} jours : à refaire pour le prochain rapport.</p>` : ''}`
        : '<p class="empty">Aucun contrôle saisi. Le NAS n\'a pas d\'agent : les points du rapport de maintenance se saisissent ici chaque mois.</p>'}
      <div class="actions supv-actions">
        ${isAdmin() ? `<button type="button" class="btn danger" id="svDel" style="margin-right:auto">Retirer</button>
          <button type="button" class="btn" id="svEdit">Modifier la fiche</button>
          <button type="button" class="btn primary" id="svCheck">✓ Saisir le contrôle du mois</button>` : ''}
        <button type="button" class="btn" id="svClose">Fermer</button>
      </div>`, true);
    $('#svClose', f).onclick = () => $('#dlg').close();
    if ($('#svEdit', f)) $('#svEdit', f).onclick = () => editNas(s);
    if ($('#svCheck', f)) $('#svCheck', f).onclick = () => checkNas(s);
    if ($('#svDel', f)) $('#svDel', f).onclick = () => confirmBox(`Retirer ${s.name} de la supervision ?`, 'La fiche du NAS et son dernier contrôle sont retirés.', 'Retirer', async () => {
      await call({ op: 'serveur.delete', id: s.id }); $('#dlg').close(); toast('NAS retiré'); load(true);
    }, () => nasView(s));
  }

  const saveNas = (s, clientId, name, data) => call({ op: 'serveur.save', id: s?.id || uid('nas').replace(/[^A-Za-z0-9_-]/g, ''), clientId, name, data: { ...data, kind: 'nas' } });

  function editNas(s) {
    const d = s?.data || {};
    const list = activeClients();
    const f = dialog(`<h2>${s ? 'Modifier le NAS' : 'Nouveau NAS'}</h2>
      <label>Client</label><select id="nsClient">${list.map(c => `<option value="${esc(c.id)}" ${c.id === s?.clientId ? 'selected' : ''}>${esc(c.name)}</option>`).join('')}</select>
      <label>Nom <span class="small">(lettres, chiffres, point, tiret)</span></label>
      <input id="nsName" value="${esc(s?.name || '')}" placeholder="ex. PAPV-SYNOLOGY" autocomplete="off" maxlength="40">
      <div class="row"><div class="grow"><label>Marque</label><input id="nsBrand" value="${esc(d.brand ?? 'Synology')}" maxlength="40"></div>
        <div class="grow"><label>Modèle</label><input id="nsModel" value="${esc(d.model || '')}" placeholder="ex. DS224+" maxlength="60"></div></div>
      <div class="row"><div class="grow"><label>Version du système (DSM)</label><input id="nsDsm" value="${esc(d.dsm || '')}" placeholder="ex. 7.2.2-72806 Update 3" maxlength="40"></div>
        <div class="grow"><label>Adresse IP</label><input id="nsIp" value="${esc(d.ip || '')}" maxlength="60"></div></div>
      <label>Rôle</label><input id="nsRole" value="${esc(d.role ?? 'Sauvegarde')}" placeholder="ex. Sauvegarde (Active Backup for Business)" maxlength="120">
      <div class="row"><div class="grow"><label>N° de série</label><input id="nsSerial" value="${esc(d.serial || '')}" maxlength="60"></div>
        <div class="grow"><label>Date d'installation</label><input type="date" id="nsDate" value="${esc(d.dateInstallation || '')}"></div></div>
      <label>Disques <span class="small">(ex. 2 × 8 To WD Red, SHR-1)</span></label><input id="nsDisks" value="${esc(d.disks || '')}" maxlength="200">
      <label>Description</label><input id="nsDesc" value="${esc(d.description || '')}" maxlength="300">
      <p class="small supv-old hidden" id="nsErr"></p>
      <div class="actions"><button type="button" class="btn" id="nsBack">${s ? 'Retour' : 'Annuler'}</button><button type="button" class="btn primary" id="nsSave">Enregistrer</button></div>`);
    $('#nsBack', f).onclick = () => s ? nasView(s) : $('#dlg').close();
    $('#nsSave', f).onclick = async ev => {
      const name = $('#nsName', f).value.trim();
      const fail = t => { const p = $('#nsErr', f); p.textContent = t; p.classList.remove('hidden'); };
      if (!/^[A-Za-z0-9_.-]{1,40}$/.test(name)) return fail('Nom invalide : lettres, chiffres, point, tiret, 40 caractères au plus, sans espace.');
      ev.target.disabled = true;
      const v = id => $(id, f).value.trim();
      try {
        const r = await saveNas(s, $('#nsClient', f).value, name, { brand: v('#nsBrand'), model: v('#nsModel'), dsm: v('#nsDsm'), ip: v('#nsIp'), role: v('#nsRole'),
          serial: v('#nsSerial'), dateInstallation: v('#nsDate'), disks: v('#nsDisks'), description: v('#nsDesc'), check: d.check || {} });
        toast('Enregistré'); await load(true); nasView(r.server);
      } catch (e) { ev.target.disabled = false; fail(e.message); }
    };
  }

  // Contrôle du mois : un statut (OK / Attention / Critique / N/A) et un commentaire par point du rapport de maintenance.
  // Prérempli avec le contrôle précédent (les commentaires se mettent à jour, la date passe à aujourd'hui).
  function checkNas(s) {
    const d = s.data || {}, it = checkOf(s).items || {};
    const today = new Date(Date.now() - new Date().getTimezoneOffset() * 60000).toISOString().slice(0, 10);
    const f = dialog(`<h2>Contrôle du mois : ${esc(s.name)}</h2>
      <p class="small muted" style="margin-top:0">${esc(cname(s.clientId))} · ${esc([d.brand, d.model].filter(Boolean).join(' '))}. Ces points alimentent la partie NAS du rapport de maintenance.</p>
      <label>Date du contrôle</label><input type="date" id="ncDate" value="${today}">
      <label>Version du système (DSM) constatée</label><input id="ncDsm" value="${esc(d.dsm || '')}" maxlength="40">
      ${NAS_CHECKS.map(([k, label]) => { const x = it[k] || {};
        return `<div class="supv-nck"><div class="small"><strong>${esc(label)}</strong></div>
          <div class="row">${NAS_ST.map(st => `<label class="check"><input type="radio" name="nc_${k}" value="${st}" ${x.st === st ? 'checked' : ''}> ${st}</label>`).join('')}</div>
          <input data-ncc="${k}" value="${esc(x.com || '')}" placeholder="commentaire (ex. dernier job le 13/11 : réussi)" maxlength="500"></div>`; }).join('')}
      <p class="small supv-old hidden" id="ncErr"></p>
      <div class="actions"><button type="button" class="btn" id="ncBack">Retour</button><button type="button" class="btn primary" id="ncSave">Enregistrer le contrôle</button></div>`, true);
    $('#ncBack', f).onclick = () => nasView(s);
    $('#ncSave', f).onclick = async ev => {
      const fail = t => { const p = $('#ncErr', f); p.textContent = t; p.classList.remove('hidden'); };
      const items = {};
      for (const [k] of NAS_CHECKS) {
        const st = $(`input[name="nc_${k}"]:checked`, f)?.value || '', com = $(`[data-ncc="${k}"]`, f).value.trim();
        if (st || com) items[k] = { st, com };
      }
      const missing = NAS_CHECKS.filter(([k]) => !items[k]?.st).map(([, l]) => l);
      if (missing.length === NAS_CHECKS.length) return fail('Choisis au moins un statut.');
      if (!$('#ncDate', f).value) return fail('Date du contrôle manquante.');
      ev.target.disabled = true;
      try {
        const r = await saveNas(s, s.clientId, s.name, { ...d, dsm: $('#ncDsm', f).value.trim(), check: { date: $('#ncDate', f).value, items } });
        toast(missing.length ? `Contrôle enregistré (${missing.length} point(s) sans statut)` : 'Contrôle enregistré'); await load(true); nasView(r.server);
      } catch (e) { ev.target.disabled = false; fail(e.message); }
    };
  }

  // ---------- script d'installation de l'agent ----------
  // Même nom que celui donné par le serveur (lib/supervision.php, op serveur.installeur).
  const safeName = x => String(x).replace(/[^A-Za-z0-9_.-]+/g, '_').replace(/^_+|_+$/g, '') || 'x';
  const installerName = s => `Installer-LevelDesk-${safeName(cname(s.clientId))}-${safeName(s.name)}.ps1`;
  function installAgent(s) {
    const f = dialog(`<h2>Installer l'agent sur ${esc(s.name)}</h2>
      <p>${s.hasToken ? '<strong>Un nouveau jeton sera créé</strong> : l\'agent déjà installé sur ce serveur cessera d\'envoyer tant qu\'il n\'est pas réinstallé avec ce nouveau fichier.'
        : 'Un jeton propre à ce serveur sera créé et placé dans le fichier d\'installation.'}</p>
      <ol class="small">
        <li>Copier le fichier téléchargé sur le serveur (il contient le jeton : ne pas l'envoyer par e-mail).</li>
        <li>Sur le serveur, ouvrir PowerShell <strong>en tant qu'administrateur</strong> (le clic droit « Exécuter avec PowerShell » ne suffit pas), aller dans le dossier du fichier et lancer :<br>
          <code>powershell.exe -ExecutionPolicy Bypass -File .\\${esc(installerName(s))}</code></li>
        <li>Attendre « Premier envoi : OK », puis supprimer le fichier.</li>
      </ol>
      <p class="small muted">L'agent lit seulement : il n'ouvre aucun port et ne reçoit aucune instruction. Il envoie son export chaque jour vers 6 h.</p>
      <p class="small supv-old hidden" id="svErr"></p>
      <div class="actions"><button type="button" class="btn" id="svBack">Retour</button><button type="button" class="btn primary" id="svGo">Télécharger le fichier d'installation</button></div>`);
    $('#svBack', f).onclick = () => openServer(s.id);
    $('#svGo', f).onclick = async ev => {
      ev.target.disabled = true;
      try {
        const r = await call({ op: 'serveur.installeur', id: s.id });
        // PowerShell 5.1 ne lit correctement les accents que dans un fichier UTF-8 AVEC marque BOM.
        download(new Blob([new Uint8Array([0xEF, 0xBB, 0xBF]), r.content], { type: 'application/octet-stream' }), r.filename);
        toast('Fichier téléchargé'); await load(true); openServer(s.id);
      } catch (e) { ev.target.disabled = false; const p = $('#svErr', f); p.textContent = e.message; p.classList.remove('hidden'); }
    };
  }

  // ---------- fiche client : serveurs supervisés (lecture ; la gestion se fait dans l'onglet Serveurs) ----------
  async function clientBox(el, c) {
    if (!el || !c || !on() || !isAdmin()) return;
    el.innerHTML = '<p class="small muted">Serveurs supervisés : chargement…</p>';
    await load(false);
    if (!el.isConnected) return;
    if (!S) { el.innerHTML = '<p class="small muted">Serveurs supervisés : consultation en ligne seulement.</p>'; return; }
    const list = S.servers.filter(s => s.clientId === c.id);
    const diff = c.server && list.length && list.length !== Math.max(1, +c.servers || 1);
    el.innerHTML = `<div class="supv-box"><div class="small"><strong>Serveurs supervisés</strong> : ${list.length ? list.map(s => `${esc(s.name)} ${lvPill(s)}`).join(' ') : 'aucun'}</div>
      ${diff ? `<div class="small supv-old">⚠ ${list.length} serveur(s) supervisé(s) pour « Nombre de serveurs » = ${esc(Math.max(1, +c.servers || 1))}.</div>` : ''}
      <button type="button" class="btn link" id="svGoTab">Gérer dans l'onglet Serveurs ›</button></div>`;
    $('#svGoTab', el).onclick = () => { $('#dlg').close(); tab = 'serveurs'; render(); };
  }

  // ---------- reprise de maintenance\Template\outils\clients.json ----------
  function importClients() {
    const inp = document.createElement('input');
    inp.type = 'file'; inp.accept = '.json,application/json';
    inp.onchange = async () => {
      const file = inp.files[0]; if (!file) return;
      let cfg = null;
      try { cfg = JSON.parse((await file.text()).replace(/^﻿/, '')); } catch (_) { /* pas du JSON */ }
      const src = cfg && cfg.clients && typeof cfg.clients === 'object' && !Array.isArray(cfg.clients) ? Object.values(cfg.clients) : null;
      const rows = (src || []).filter(c => c && typeof c === 'object' && Array.isArray(c.serveurs) && c.serveurs.length);
      if (!rows.length) { toast('Fichier non reconnu : rien n\'a été créé.'); return; }
      importDialog(rows);
    };
    inp.click();
  }

  function importDialog(rows) {
    const norm = s => String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().replace(/[^a-z0-9]+/g, '');
    const cl = activeClients();
    const guess = r => cl.find(c => [r.code, r.nom].filter(Boolean).some(x => norm(x) === norm(c.name)))?.id || '';
    const have = (cid, name) => (S?.servers || []).some(s => s.clientId === cid && s.name.toLowerCase() === String(name).toLowerCase());
    const f = dialog(`<h2>Importer clients.json</h2>
      <p class="small muted">Chaque client du fichier est relié à un client LevelDesk. Un serveur déjà présent (même client, même nom) est ignoré. Les NAS qui ont un nom sont repris aussi (sans contrôle).</p>
      ${rows.map((r, i) => `<div class="supv-imp"><div class="small"><strong>${esc(r.code || '?')}</strong> ${esc(r.nom || '')}
          <span class="muted">· ${r.serveurs.map(s => esc(s && s.nom)).join(', ')}</span></div>
        <select data-imp="${i}"><option value="">Ne pas importer</option>${cl.map(c => `<option value="${esc(c.id)}" ${c.id === guess(r) ? 'selected' : ''}>${esc(c.name)}</option>`).join('')}</select></div>`).join('')}
      <p class="small hidden" id="svImpRes"></p>
      <div class="actions"><button class="btn" value="cancel">Fermer</button><button type="button" class="btn primary" id="svImpGo">Importer</button></div>`, true);
    $('#svImpGo', f).onclick = async ev => {
      ev.target.disabled = true;
      let made = 0, there = 0, bad = 0; const skipped = [], errors = [];
      for (const sel of $$('[data-imp]', f)) {
        const r = rows[+sel.dataset.imp], cid = sel.value;
        if (!cid) { skipped.push(r.code || r.nom || '?'); continue; }
        for (const sv of r.serveurs) {
          const name = String(sv?.nom || '').trim();
          if (!/^[A-Za-z0-9_.-]{1,40}$/.test(name)) { bad++; continue; }
          if (have(cid, name)) { there++; continue; }
          try {
            await call({ op: 'serveur.save', id: uid('sv').replace(/[^A-Za-z0-9_-]/g, ''), clientId: cid, name, data: {
              type: sv.type, roles: sv.roles || [], description: sv.description || '', lob: sv.lob || [], backup: sv.backup || '', dateInstallation: sv.dateInstallation || '' } });
            made++;
          } catch (e) { if (/\(409\)/.test(e.message)) there++; else errors.push(`${name} : ${e.message}`); }
        }
        // NAS du fichier (nom, modèle, rôle) : repris s'ils ont un nom (14/11). Nom avec espaces → tirets.
        for (const nv of (Array.isArray(r.nas) ? r.nas : r.nas ? [r.nas] : [])) {
          const name = String(nv?.nom || '').trim().replace(/\s+/g, '-');
          if (!name) continue;
          if (!/^[A-Za-z0-9_.-]{1,40}$/.test(name)) { bad++; continue; }
          if (have(cid, name)) { there++; continue; }
          try { await saveNas(null, cid, name, { brand: /^DS|^RS/i.test(nv.modele || '') || /synology/i.test(name) ? 'Synology' : '', model: nv.modele || '', role: nv.role || '' }); made++; }
          catch (e) { if (/\(409\)/.test(e.message)) there++; else errors.push(`${name} : ${e.message}`); }
        }
      }
      await load(true);
      const p = $('#svImpRes', f);
      p.textContent = `${made} serveur(s) créé(s), ${there} déjà présent(s)` + (bad ? `, ${bad} nom(s) invalide(s)` : '') + (skipped.length ? ` · non reliés : ${skipped.join(', ')}` : '') + (errors.length ? ` · erreurs : ${errors.join(' ; ')}` : '');
      p.classList.remove('hidden');
      ev.target.disabled = false;
    };
  }

  // ---------- page d'un client : ses serveurs, ses maintenances, ses rapports ----------
  const reps = {};              // rapports connus par client : { <clientId>: [{ month, status, updatedAt, updatedBy }] }
  async function loadReps(cid) {
    try { reps[cid] = (await call({ op: 'rapports', clientId: cid })).reports || []; } catch (e) { reps[cid] = reps[cid] || []; toast(e.message); }
    if (tab === 'serveurs' && page?.cli === cid && !page.month) drawClient();
  }

  function drawClient() {
    const el = $('#view'), cid = page.cli, c = client(cid), month = today().slice(0, 7);
    const list = S.servers.filter(s => s.clientId === cid);
    const forf = mod('forfaits');
    const done = maintOf(cid, month)[0];
    const hist = live().filter(e => e.clientId === cid && e.forfait).sort((a, b) => b.date.localeCompare(a.date)).slice(0, 12);
    const prevM = isoDate(new Date(+month.slice(0, 4), +month.slice(5, 7) - 2, 1)).slice(0, 7);
    const rl = reps[cid];
    const has = m => (rl || []).some(r => r.month === m);
    el.innerHTML = `<div class="row supv-back"><button type="button" class="btn" id="svAll">‹ Tous les serveurs</button><h2 class="grow">${esc(c.name)}</h2></div>
      <div class="card"><h2>Serveurs et NAS<span class="tot">${list.length} appareil${list.length > 1 ? 's' : ''}</span></h2>
        ${list.length ? `<ul class="list">${list.map(s => srvLine(s, false)).join('')}</ul>` : '<p class="empty">Aucun serveur supervisé pour ce client.</p>'}</div>
      ${forf ? `<div class="card"><h2>Maintenance<span class="tot">${esc(monthName(month))}</span></h2>
        <p class="${done ? 'supv-t-ok' : 'supv-old'}" style="margin:0 0 8px">${done ? `✓ Maintenance ${esc(deMonth(month))} faite le ${esc(frDate(done.date))} (${fmtDur(done.minutes)}${done.techId ? ', ' + esc(user(done.techId).name) : ''}).` : `Maintenance ${esc(deMonth(month))} pas encore encodée.`}</p>
        <div class="row"><button type="button" class="btn ${done ? '' : 'primary'}" id="svMaint">${done ? 'Ouvrir l\'intervention' : '🛠 Encoder la maintenance du mois'}</button></div>
        ${hist.length ? `<div class="supv-sec">Dernières maintenances et mises à jour</div><ul class="list">${hist.map(e => entryLi(e, { showDate: true })).join('')}</ul>`
          : '<p class="small muted" style="margin-top:10px">Aucune intervention « Maintenance serveur » ou « Mise à jour serveur » pour ce client.</p>'}</div>` : ''}
      ${isAdmin() ? `<div class="card"><h2>Rapports de maintenance</h2>
        <p class="small muted" style="margin-top:0">Rapport mensuel pour le client, préparé à partir des exports des agents, des contrôles des NAS et des interventions du mois.</p>
        <div class="row" style="flex-wrap:wrap;gap:8px">
          ${has(month) ? '' : `<button type="button" class="btn primary" data-rp="${month}">📄 Rapport ${esc(deMonth(month))}</button>`}
          ${has(prevM) ? '' : `<button type="button" class="btn" data-rp="${prevM}">Rapport ${esc(deMonth(prevM))}</button>`}
          <input type="month" id="svRpMonth" max="${month}" style="width:auto;margin:0" title="Autre mois"><button type="button" class="btn" id="svRpGo">Ouvrir</button>
        </div>
        ${!rl ? '<p class="small muted">Chargement…</p>' : rl.length ? `<ul class="list" style="margin-top:8px">${rl.map(r => `<li class="tap" data-rp="${esc(r.month)}"><div class="grow"><strong>${esc(monthName(r.month))}</strong>
            ${r.status === 'valid' ? '<span class="pill supv-ok">validé</span>' : '<span class="pill supv-warn">brouillon</span>'}
            <div class="small muted">modifié le ${esc(frDate(String(r.updatedAt).slice(0, 10)))}${r.updatedBy ? ' par ' + esc(user(r.updatedBy).name) : ''}</div></div><div class="dur small">›</div></li>`).join('')}</ul>`
          : '<p class="small muted">Aucun rapport enregistré pour ce client.</p>'}</div>` : ''}`;
    $('#svAll').onclick = goAll;
    $$('[data-srv]', el).forEach(li => li.onclick = () => openServer(li.dataset.srv));
    bindEdits(el);
    if ($('#svMaint')) $('#svMaint').onclick = () => addMaint(cid);
    $$('[data-rp]', el).forEach(b => b.onclick = () => openReport(cid, b.dataset.rp));
    if ($('#svRpGo')) $('#svRpGo').onclick = () => { const m = $('#svRpMonth').value; /^\d{4}-\d{2}$/.test(m) && m <= month ? openReport(cid, m) : toast('Choisis un mois (pas dans le futur).'); };
    if (isAdmin() && !rl) loadReps(cid);
  }

  // ---------- rapport de maintenance mensuel ----------
  // Reprise de l'ancien générateur Word (maintenance\Template\outils\build_report.py) : mêmes rubriques et mêmes règles, calculées ici depuis les exports
  // des agents (dernier export jusqu'à la date de l'intervention + un export plus ancien pour les tendances), les contrôles des NAS et les interventions.
  // Chaque ligne calculée peut être corrigée (statut, commentaire) ; un rapport validé garde une copie figée (snap) : il ne change plus.
  let R = null;                 // rapport ouvert : { cli, month, status, rec, data, A, days, preview }
  const expCache = new Map();
  const SPP = { verif: '2026-10-19', gen10: '2026.07.00.00', gen10plus: '2026.07.00.00', gen11: '2026.07.00.00', gen12: '2026.07.00.00' };   // dernier Service Pack for ProLiant par génération (hpe_spp.json)
  const ORDER = { Critique: 0, Attention: 1, 'À compléter': 2, Info: 3, OK: 4, 'N/A': 5 };
  const RP_ST = ['OK', 'Attention', 'Critique', 'À compléter', 'N/A', 'hide'];
  const L = x => x === null || x === undefined || x === '' ? [] : Array.isArray(x) ? x : [x];
  const obj = x => x && typeof x === 'object' && !Array.isArray(x) ? x : null;
  const pd = s => {
    if (s instanceof Date) return s;
    const m = String(s || '').match(/^(\d{4})-(\d{2})-(\d{2})(?:[ T](\d{2}):(\d{2}))?/);
    if (m) return new Date(+m[1], m[2] - 1, +m[3], +(m[4] || 0), +(m[5] || 0));
    const f = String(s || '').match(/^(\d{2})\/(\d{2})\/(\d{2,4})/);
    return f ? new Date(f[3].length === 2 ? 2000 + +f[3] : +f[3], f[2] - 1, +f[1]) : null;
  };
  const fd = (s, year = true) => { const d = pd(s); return d ? `${pad(d.getDate())}/${pad(d.getMonth() + 1)}${year ? '/' + d.getFullYear() : ''}` : String(s || ''); };
  const days = (a, b) => Math.floor((a - b) / 86400000);
  const gb = v => v === null || v === undefined || isNaN(+v) ? '?' : +v >= 1000 ? (v / 1024).toFixed(2).replace('.', ',') + ' To' : Math.round(v) + ' Go';
  const pct = v => v === null || v === undefined || isNaN(+v) ? '?' : String(Math.round(+v));
  const worst = a => { const s = a.filter(x => x in ORDER); return s.length ? s.reduce((m, x) => ORDER[x] < ORDER[m] ? x : m) : 'OK'; };
  const osShort = c => { const m = /Windows Server (\d{4}( R2)?)/.exec(c || ''); return m ? 'Windows Server ' + m[1] : String(c || '?').replace('Microsoft ', ''); };

  // srvs : fiches des serveurs (physiques d'abord) ; nasList : fiches des NAS ; cur / prev : { nom: export } ; date : jour de l'intervention (Date).
  function analyse(srvs, nasList, cur, prev, date) {
    const rows = {}, alerts = [], missing = [], actions = [];
    const have = srvs.filter(s => obj(cur[s.name])).map(s => [s.name, cur[s.name]]);
    const endDay = new Date(date.getFullYear(), date.getMonth(), date.getDate(), 23, 59);
    srvs.forEach(s => { if (!obj(cur[s.name])) missing.push(`Pas d'export récent du serveur ${s.name} (agent installé ? dernier envoi ?)`); });
    for (const [n, e] of have) {
      L(e.alerts).forEach(a => obj(a) && alerts.push({ level: String(a.level || 'Info'), category: String(a.category || ''), message: String(a.message || ''), action: String(a.action || ''), server: n }));
      L(e.checkErrors).forEach(ce => alerts.push({ level: 'Info', category: 'Script', server: n, message: `Contrôle « ${obj(ce)?.check ?? ce} » en erreur pendant la collecte`, action: 'Vérifier manuellement' }));
    }
    const alOf = (cats, server) => alerts.filter(a => cats.includes(a.category) && (!server || a.server === server));
    const stAl = (al, def = 'OK') => al.some(a => a.level === 'Critique') ? 'Critique' : al.some(a => a.level === 'Attention') ? 'Attention' : def;
    const add = (sec, id, label, st, com, hint) => (rows[sec] = rows[sec] || []).push({ id, label, st, com: com || '', hint: hint || '' });
    const todo = (sec, id, label, com, miss) => { add(sec, id, label, 'À compléter', com); if (miss) missing.push(miss); };
    // consigne pour celui qui remplit : affichée en grisé dans le champ, jamais imprimée
    const ask = (sec, id, label, hint, miss) => { add(sec, id, label, 'À compléter', '', hint); if (miss) missing.push(miss); };
    const msgs = al => al.map(a => `${a.server} : ${a.message}`).join(' ; ');
    const isPhys = s => String(s.data?.type || '').startsWith('phys');
    const roles = s => s.data?.roles || [];
    const phys = srvs.filter(isPhys);
    const physName = (phys[0] || srvs[0] || {}).name;
    const pe = obj(cur[physName]);

    // inventaire
    let vmi = 0;
    const inv = srvs.map(s => {
      const e = obj(cur[s.name]), p = isPhys(s);
      const prefix = p ? 'SRV-PHYSIQUE' : 'VM-' + pad(++vmi);
      const ip = e ? L(e.system?.ipAddresses).filter(i => !/^(vethernet \(default|loopback)/i.test(String(obj(i)?.interface || ''))).map(i => obj(i) ? i.ip : i).filter(Boolean).join(', ') : '?';
      const role = p && roles(s).includes('HyperV') ? 'Hôte Hyper-V' : s.data?.description || roles(s).join(', ');
      const st = !e ? 'Non collecté' : alerts.some(a => a.server === s.name && a.level === 'Critique') ? 'Critique' : 'OK';
      return { nom: `${prefix} - ${e?.system?.computerName || s.name}`, role, os: e ? osShort(e.system?.osCaption) : '?', ip: ip || '?', st };
    });
    let hwLine = '';
    if (pe && !pe.system?.isVirtual) {
      const sy = pe.system || {};
      hwLine = `Serveur physique ${[sy.manufacturer, sy.model].filter(Boolean).join(' ')} – n° de série ${sy.serialNumber || '?'} – date d'installation ${fd(phys[0]?.data?.dateInstallation || sy.installDate)}`;
    }

    // sauvegardes
    // serveurs qui font une sauvegarde : déclarée dans la fiche du serveur, ou vue par l'agent
    const bsrv = srvs.filter(s => s.data?.backup || cur[s.name]?.backup?.expected || cur[s.name]?.backup?.lastSuccessDetected);
    const btype = [...new Set(bsrv.map(s => s.data?.backup || cur[s.name]?.backup?.expected).filter(Boolean))].sort().join(', ');
    let last = null, warnBk = false;
    for (const s of bsrv) { const b = cur[s.name]?.backup, d = pd(b?.lastSuccessDetected); if (d && (!last || d > last)) { last = d; warnBk = !!b.lastSuccessWithWarnings; } }
    const LB = 'État des sauvegardes (dernière exécution réussie)';
    if (last) {
      const errs = alOf(['Sauvegardes']);
      const st = days(endDay, last) > 2 ? 'Critique' : warnBk || errs.length ? 'Attention' : 'OK';
      add('backup', 'backup_etat', LB, st, `Dernière sauvegarde réussie le ${fd(last, false)} (${btype})` + (warnBk ? ' – terminée avec des avertissements' : '') + (errs.length ? ' – ' + errs[0].message : ''));
    } else todo('backup', 'backup_etat', LB, `Non détectée automatiquement (${btype || 'aucune sauvegarde déclarée dans les fiches des serveurs'}) – vérifier la console`, 'Date de la dernière sauvegarde réussie (non détectée par l\'agent)');
    ask('backup', 'backup_restore', 'Test mensuel de restauration de fichiers', 'backup restauré, élément restauré, résultat', 'Test mensuel de restauration (backup restauré, élément, résultat)');
    const usb = bsrv.flatMap(s => L(cur[s.name]?.storage?.usbDisks)).filter(obj);
    const tgt = bsrv.map(s => cur[s.name]?.backup?.wsb?.target).filter(Boolean);
    usb.length || tgt.length ? todo('backup', 'backup_espace', 'Espace de stockage de sauvegarde disponible', usb.length ? 'Disque(s) USB : ' + usb.map(u => `${u.name} ${gb(u.sizeGB)}`).join(', ') : 'Cible : ' + tgt.join(' ; '), 'Espace libre / total du support de sauvegarde')
      : ask('backup', 'backup_espace', 'Espace de stockage de sauvegarde disponible', 'espace libre / total de la cible', 'Espace libre / total du support de sauvegarde');
    const sc = have.map(([n, e]) => [n, L(e.backup?.shadowCopies).filter(x => obj(x) && /:\\$/.test(String(x.volume || '')))]).filter(([, v]) => v.length);
    if (sc.length) add('backup', 'backup_vss', 'Clichés instantanés (versions précédentes des fichiers)', 'OK',
      sc.map(([n, v]) => `${n} : ` + v.map(x => `${String(x.volume).replace(/\\$/, '')} ${x.count} versions (dernière ${fd(x.newest, false)})`).join(', ')).join(' ; '));

    // mises à jour
    const secC = [], sysC = [], secS = [], sysS = [];
    for (const [n, e] of have) {
      const u = obj(e.updates) || {};
      const ls = u.lastSecurityUpdate?.date || u.lastInstallSuccess?.date || u.lastHotfix?.date;
      const pend = L(u.pending).filter(obj);
      const secp = pend.filter(p => /Security|Sécurité|Critical|Critique/.test(p.categories || '') || /cumulati/i.test(p.title || ''));
      secS.push(secp.length || (pd(ls) && days(endDay, pd(ls)) > 45) ? 'Attention' : 'OK');
      secC.push(`${n} : ${ls ? fd(ls, false) : '?'}` + (secp.length ? ` (${secp.length} en attente)` : ''));
      const fails = L(u.failedPeriod), other = pend.length - secp.length;
      sysS.push(u.rebootPending || fails.length || other > 0 ? 'Attention' : 'OK');
      const bits = [other > 0 ? `${other} mise(s) à jour disponible(s)` : '', u.rebootPending ? 'redémarrage requis' : '', fails.length ? `${fails.length} échec(s)` : ''].filter(Boolean);
      sysC.push(`${n} : ` + (bits.join(', ') || 'à jour'));
    }
    if (have.length) {
      add('maj', 'maj_secu', 'Mises à jour de sécurité Windows', worst(secS), 'Dernière mise à jour cumulative – ' + secC.join(' ; '));
      add('maj', 'maj_sys', 'Mises à jour système / correctifs', worst(sysS), sysC.join(' ; '));
    } else { todo('maj', 'maj_secu', 'Mises à jour de sécurité Windows', ''); todo('maj', 'maj_sys', 'Mises à jour système / correctifs', ''); }
    const av = [];
    for (const [n, e] of have) { const es = obj(e.security?.eset?.info) || {}, ver = es.ProductVersion || es.ProductVersionString; if (ver) av.push(`${n} : ${es.ProductName || 'ESET'} ${ver}`); }
    for (const x of srvs.flatMap(s => s.data?.lob || [])) {
      const sw = have.flatMap(([, e]) => L(e.software)).find(w => obj(w) && String(w.name || '').toLowerCase().includes(String(x).toLowerCase()));
      if (sw) av.push(`${sw.name} ${sw.version || ''}`.trim());
    }
    add('maj', 'maj_logiciels', 'Mises à jour logicielles (antivirus, applications métier…)', 'OK', av.join(' ; ') || 'À jour');
    const hw = obj(pe?.hardware) || {}, model = String(pe?.system?.model || '');
    const gen = [['gen12', /Gen\s?12/i], ['gen11', /Gen\s?11/i], ['gen10plus', /Gen\s?10\s?Plus/i], ['gen10', /Gen\s?10/i], ['gen9', /Gen\s?9/i]].find(([, re]) => re.test(model))?.[0];
    const latest = gen ? SPP[gen] : null;
    let inst = hw.sppDetected ? (/20\d{2}\.\d{2}\.\d{2}\.\d{2}|20\d{2}\.\d{2}/.exec(hw.sppDetected) || [String(hw.sppDetected)])[0] : '';
    const LS = 'Mises à jour des pilotes (drivers) — serveur physique';
    if (phys.length && !pe) todo('maj', 'maj_spp', LS, '');
    else if (phys.length && /HPE?|Hewlett/i.test(String(pe.system?.manufacturer || '') + model)) {
      if (inst && latest && inst.slice(0, 7) < latest.slice(0, 7)) {
        add('maj', 'maj_spp', LS, 'Attention', `Support Pack installé ${inst} – nouvelle version ${latest} disponible, installation proposée`);
        actions.push(`Installer le Service Pack for ProLiant ${latest} sur ${physName} (installé : ${inst}) – prévoir un redémarrage`);
      } else if (inst) add('maj', 'maj_spp', LS, 'OK', `Support Pack ${inst}` + (latest ? ' (dernière version)' : ''));
      else todo('maj', 'maj_spp', LS, `Version du Support Pack non détectée – dernière disponible : ${latest || '?'}`, `Version du Support Pack HPE installée sur ${physName} (iLO > Firmware & OS Software)`);
    }
    if (phys.length && pe) add('maj', 'maj_firmware', 'Mise à jour du firmware / BIOS — serveur physique', 'OK',
      `ROM ${hw.romVersion || pe.system?.biosVersion || '?'} du ${fd(hw.romDate || pe.system?.biosDate)}` + (hw.iloVersion ? ` ; iLO ${hw.iloVersion}` : ''));

    // NAS : fiche et contrôle du mois saisis dans l'onglet Serveurs
    const nas = nasList.map((s, i) => {
      const d = s.data || {}, ck = obj(d.check) || {}, it = obj(ck.items) || {};
      const old = ck.date && days(endDay, pd(ck.date)) > NAS_OLD_D;
      if (!ck.date) missing.push(`${s.name} : contrôle du mois du NAS à saisir (fiche du NAS)`);
      else if (old) missing.push(`${s.name} : contrôle du NAS du ${fd(ck.date)} trop ancien, à refaire`);
      if (!d.dsm) missing.push(`${s.name} : version DSM`);
      return { info: { nom: s.name, modele: [d.brand, d.model].filter(Boolean).join(' '), dsm: d.dsm || '', role: d.role || '', synology: /synology/i.test(d.brand || '') || /^(DS|RS)/i.test(d.model || '') },
        rows: NAS_CHECKS.map(([k]) => { const x = !old && obj(it[k]) || {};
          return { id: `nas${i + 1}_${k}`, label: NAS_LABELS[k], st: NAS_ST.includes(x.st) ? x.st : 'À compléter', com: x.com || '' }; }) };
    });

    // santé
    const evC = [], evS = [];
    for (const [n, e] of have) {
      const s_ = obj(e.events?.system) || {}, a_ = obj(e.events?.application) || {};
      const crit = (+s_.critical || 0) + (+a_.critical || 0), er = (+s_.errors || 0) + (+a_.errors || 0);
      evC.push(`${n} : ${crit} crit. / ${er} err.`); evS.push(crit > 0 || er > 150 ? 'Attention' : 'OK');
    }
    have.length ? add('sante', 'sante_journaux', 'Journaux d\'événements (erreurs / avertissements critiques)', worst(evS), evC.join(' ; ')) : todo('sante', 'sante_journaux', 'Journaux d\'événements (erreurs / avertissements critiques)', '');
    const hwal = alOf(['Matériel', 'Santé']).filter(a => !a.message.includes('arrêt'));
    add('sante', 'sante_alertes', 'Alertes système', stAl(hwal), msgs(hwal.slice(0, 3)) || 'Aucune alerte matérielle');
    const sv = alOf(['Services']);
    add('sante', 'sante_services', 'État des services critiques', stAl(sv), msgs(sv) || 'Tous les services critiques sont démarrés');
    const unexp = have.map(([n, e]) => [n, L(e.events?.unexpectedShutdowns).length]).filter(([, x]) => x);
    add('sante', 'sante_reboot', 'Redémarrages inattendus', unexp.length ? 'Attention' : 'OK', unexp.map(([n, x]) => `${n} : ${x}`).join(', ') || 'Aucun');
    const ue = have.map(([n, e]) => [n, obj(e.ups)]).find(([, u]) => u);
    if (ue) {
      const u = ue[1], bad = u.selfTestResult && !/^(ok|pass(ed)?)$/i.test(u.selfTestResult), ual = alOf(['Onduleur']);
      const st = bad ? 'Critique' : worst([stAl(ual), (u.batteryCapacityPct ?? 100) < 90 || u.batteryAgeMonths > 48 || u.internalTempC > 35 ? 'Attention' : 'OK']);
      add('sante', 'sante_ups', 'Onduleur (UPS)', st, [u.selfTestResult ? `auto-test ${bad ? 'en échec : batterie à remplacer' : 'réussi'}${u.lastSelfTest ? ' le ' + fd(u.lastSelfTest, false) : ''}` : '',
        u.batteryCapacityPct != null ? `capacité de la batterie ${pct(u.batteryCapacityPct)} %` : '', u.batteryDate ? `batterie de ${String(u.batteryDate).split('-').reverse().join('/')}` : '',
        u.internalTempC != null ? `${pct(u.internalTempC)} °C` : '', u.onBatteryEvents ? `${u.onBatteryEvents} passage(s) sur batterie` : 'aucune coupure de courant'].filter(Boolean).join(' – '));
    } else add('sante', 'sante_ups', 'Onduleur (UPS)', 'hide', '');
    const near = have.map(([n, e]) => [n, e.system?.osEndOfSupport]).filter(([, d]) => pd(d) && days(pd(d), endDay) < 365);
    if (near.length) add('sante', 'sante_eol', 'Support du système d\'exploitation', 'Attention', near.map(([n, d]) => `${n} : fin de support le ${fd(d)}`).join(' ; '));

    // ressources : valeur du mois, valeur précédente, tendance
    const res = [];
    const trend = (a, b) => a == null || b == null ? '—' : Math.abs(a - b) <= Math.max(0.03 * Math.max(Math.abs(b), 1), 0.5) ? 'stable' : a > b ? '↑' : '↓';
    for (const s of srvs) {
      const e = obj(cur[s.name]), p = obj(prev[s.name]);
      if (!e) continue;
      const pv = Object.fromEntries(L(p?.storage?.volumes).filter(obj).map(v => [v.drive, v]));
      for (const v of L(e.storage?.volumes).filter(v => obj(v) && (+v.sizeGB || 0) >= 1)) {
        const o = pv[v.drive];
        res.push([`Espace disque disponible — ${s.name} ${v.drive}`, `${gb(v.freeGB)} libres sur ${gb(v.sizeGB)} (${pct(v.freePct)} %)`, o ? `${gb(o.freeGB)} libres` : '—', trend(v.freeGB, o ? o.freeGB : null)]);
      }
      const pf = obj(e.performance), pp = obj(p?.performance) || {};
      if (pf) {
        res.push([`Mémoire utilisée — ${s.name}`, `${pct(pf.memUsedPct)} % – ${String(pf.memUsedGB ?? '?').replace('.', ',')}/${String(pf.memTotalGB ?? '?').replace('.', ',')} Go`, pp.memUsedPct != null ? pct(pp.memUsedPct) + ' %' : '—', trend(pf.memUsedPct, pp.memUsedPct)]);
        res.push([`Charge processeur (moy.) — ${s.name}`, `${pct(pf.cpuAvgPct)} %`, pp.cpuAvgPct != null ? pct(pp.cpuAvgPct) + ' %' : '—', trend(pf.cpuAvgPct, pp.cpuAvgPct)]);
      }
    }

    // sécurité
    const avS = [], avC = [];
    for (const [n, e] of have) {
      const es = obj(e.security?.eset), de = obj(e.security?.defender);
      if (es && es.service === 'Running') { avS.push('OK'); avC.push(`${n} : ESET actif`); }
      else if (de && de.realTime && de.mode === 'Normal') { avS.push((+de.signatureAgeDays || 0) <= 3 ? 'OK' : 'Attention'); avC.push(`${n} : Defender actif`); }
      else { avS.push('Critique'); avC.push(`${n} : aucun antivirus actif détecté`); }
    }
    have.length ? add('secu', 'secu_av', 'Contrôle antivirus / antimalware', worst(avS), avC.join(' ; ')) : todo('secu', 'secu_av', 'Contrôle antivirus / antimalware', '');
    const dets = have.map(([n, e]) => [n, L(e.security?.defenderDetections).length]).filter(([, x]) => x).map(([n, x]) => `${n} : ${x} (Defender)`);
    dets.length ? add('secu', 'secu_quarantaine', 'Détections en quarantaine', 'Attention', dets.join(' ; '))
      : ask('secu', 'secu_quarantaine', 'Détections en quarantaine', 'à vérifier dans la console ESET', 'Détections / quarantaine dans la console ESET');
    ask('secu', 'secu_adware', 'Adware / logiciels indésirables', 'à vérifier dans la console ESET', 'Adware / applications indésirables (console ESET)');
    const acC = [], acS = [];
    for (const [n, e] of have) {
      const fl = +e.security?.failedLogons?.count || 0, lk = +e.security?.lockouts?.count || 0;
      if (fl || lk) acC.push(`${n} : ${fl} échec(s) de connexion, ${lk} verrouillage(s)`);
      acS.push(fl > 500 ? 'Attention' : 'OK');
      const ad = obj(e.activeDirectory);
      if (ad) {
        if (L(ad.usersInactive90d).length) acC.push(`${L(ad.usersInactive90d).length} compte(s) inactif(s) depuis plus de 90 j`);
        if (L(ad.lockedOut).length) acC.push('verrouillé(s) : ' + L(ad.lockedOut).map(x => obj(x) ? x.name || x.samAccountName || '' : x).filter(Boolean).join(', '));
      }
    }
    have.length ? add('secu', 'secu_comptes', 'Comptes utilisateurs (verrouillages, tentatives suspectes)', worst(acS), acC.join(' ; ') || 'Aucune activité suspecte') : todo('secu', 'secu_comptes', 'Comptes utilisateurs (verrouillages, tentatives suspectes)', '');
    const certs = alOf(['Sécurité']).filter(a => /ertificat/.test(a.message));
    if (certs.length) add('secu', 'secu_certs', 'Certificats (expiration)', 'Attention', msgs(certs));
    const conf = alOf(['Sécurité']).filter(a => /SMB1|Pare-feu/.test(a.message));
    if (conf.length) add('secu', 'secu_config', 'Configuration (pare-feu, protocoles)', 'Attention', msgs(conf));

    // machines virtuelles
    const hv = have.map(([, e]) => obj(e.hyperV)).filter(Boolean);
    if (hv.length) {
      const vms = hv.flatMap(h => L(h.vms)).filter(obj), down = vms.filter(v => v.state !== 'Running').map(v => v.name), al = alOf(['Hyper-V']);
      add('vm', 'vm_etat', 'État général des VM', down.length || al.some(a => /intégration|introuvable/.test(a.message)) ? 'Attention' : 'OK', `${vms.length} VM, ` + (down.length ? `arrêtée(s) : ${down.join(', ')}` : 'toutes démarrées'));
      const cps = hv.flatMap(h => L(h.checkpoints)).filter(obj);
      add('vm', 'vm_snap', 'Snapshots (nombre, ancienneté, cohérence)', cps.some(c => (+c.ageDays || 0) > 7) ? 'Attention' : 'OK', cps.map(c => `${c.vm} : « ${c.name} » du ${fd(c.created)}`).join(' ; ') || 'Aucun point de contrôle');
      const h = hv[0], vcpu = vms.reduce((n, v) => n + (+v.cpuCount || 0), 0);
      add('vm', 'vm_ress', 'Utilisation des ressources allouées', h.hostRamGB && (+h.vmMemoryAssignedGB || 0) / h.hostRamGB > 0.9 ? 'Attention' : 'OK',
        `Mémoire assignée ${String(h.vmMemoryAssignedGB ?? '?').replace('.', ',')} Go / ${String(h.hostRamGB ?? '?').replace('.', ',')} Go – ${vcpu} processeurs virtuels / ${h.hostLogicalProcessors ?? '?'} processeurs logiques`);
    } else if (srvs.some(s => roles(s).includes('HyperV'))) {
      for (const [id, lab] of [['vm_etat', 'État général des VM'], ['vm_snap', 'Snapshots (nombre, ancienneté, cohérence)'], ['vm_ress', 'Utilisation des ressources allouées']]) todo('vm', id, lab, '');
    }

    // services essentiels
    const DEF = {
      dns: d => `Résolution interne et externe OK – redirecteurs : ${L(d.forwarders).join(', ') || 'aucun (indications de racine)'}`,
      dhcp: d => L(d.scopes).filter(obj).map(s => `Étendue ${s.scopeId} : ${s.percentInUse} % utilisée (${s.free} libres)`).join(' ; '),
      activeDirectory: d => `Domaine ${d.domain || '?'} – ${L(d.domainControllers).length} contrôleur(s) de domaine – ${d.usersEnabled ?? '?'} comptes actifs`,
      rds: d => `Licences : ${d.licensingMode || '?'} – ${d.sessionCount || 0} session(s) au moment du contrôle`,
      printServer: d => `${L(d.printers).length} imprimante(s), spouleur démarré` };
    const SVC = [['svc_dns', 'DNS', 'dns', 'DNS', 'DNS'], ['svc_dhcp', 'DHCP', 'dhcp', 'DHCP', 'DHCP'], ['svc_ad', 'Active Directory', 'activeDirectory', 'Active Directory', 'AD'],
      ['svc_rds', 'Services Bureau à distance (RDS)', 'rds', 'RDS', 'RDS'], ['svc_print', 'Serveur d\'impression', 'printServer', 'Impression', 'Print']];
    for (const [id, lab, key, cat, role] of SVC) {
      const es = have.map(([, e]) => obj(e[key])).filter(Boolean);
      if (es.length) { const al = alOf([cat]); add('svc', id, lab, stAl(al), msgs(al) || DEF[key](es[0]) || 'OK'); }
      else { const s = srvs.find(x => roles(x).includes(role) && obj(cur[x.name]));
        if (s) todo('svc', id, lab, `Contrôle automatique non collecté sur ${s.name}`, `${lab} : contrôle non collecté sur ${s.name}`);
        else if (!have.length && ['svc_dns', 'svc_dhcp', 'svc_ad'].includes(id)) todo('svc', id, lab, ''); }
    }
    const fs = have.map(([n, e]) => [n, L(e.fileServer?.shares).length]).filter(([, x]) => x);
    if (fs.length) add('svc', 'svc_files', 'Serveur de fichiers', stAl(alOf(['Fichiers'])), fs.map(([n, x]) => `${n} : ${x} partage(s)`).join(' ; '));

    // actions pour le technicien (pas imprimées)
    const seen = new Set();
    for (const a of [...alerts].sort((x, y) => (ORDER[x.level] ?? 9) - (ORDER[y.level] ?? 9))) {
      if (!['Critique', 'Attention'].includes(a.level) || !a.action) continue;
      const k = a.server + '|' + a.action + '|' + a.message.slice(0, 40);
      if (!seen.has(k)) { seen.add(k); actions.push(`[${a.level}] ${a.server} – ${a.message} → ${a.action}`); }
    }
    return { rows, inv, hwLine, nas, res, alerts, missing, actions };
  }
  const NAS_LABELS = { jobs: 'État des sauvegardes NAS (jobs de sauvegarde / réplication)', restore: 'Test de restauration depuis le NAS', espace: 'Espace disque disponible (volumes / pools)',
    disques: 'État des disques (SMART, RAID / SHR)', maj_dsm: 'Mises à jour du système (DSM)', paquets: 'Mises à jour des paquets installés (Package Center)' };

  const autoResume = A => {
    const crit = A.alerts.filter(a => a.level === 'Critique'), att = A.alerts.filter(a => a.level === 'Attention');
    if (!crit.length && !att.length) return 'Les serveurs sont en bon état, les sauvegardes et les mises à jour sont en ordre. Aucun point bloquant n\'a été relevé ce mois-ci.';
    return [crit.length ? 'Point(s) critique(s) : ' + crit.slice(0, 3).map(a => a.message).join(' ; ') + '.' : '', att.length ? 'Point(s) à surveiller : ' + att.slice(0, 4).map(a => a.message).join(' ; ') + '.' : ''].filter(Boolean).join(' ');
  };
  const autoRecos = A => [...new Set(A.alerts.filter(a => ['Critique', 'Attention'].includes(a.level) && a.action).map(a => `${a.action} (${a.server})`))].slice(0, 8);
  // Interventions du client pendant le mois, proposées pour la rubrique « Interventions correctives » (sans durée ni prix : document pour le client).
  const monthEntries = (cid, month) => live().filter(e => e.clientId === cid && e.date.startsWith(month)).sort((a, b) => a.date.localeCompare(b.date))
    .filter(e => e.description && !Object.values(FORFAITS).includes(e.description.trim())).map(e => `${fd(e.date, false)} : ${e.description.trim().replace(/\s*\n\s*/g, ' – ')}`);

  function openReport(cid, month) { page = { cli: cid, month }; R = null; draw(); window.scrollTo(0, 0); loadReport(); }

  async function exportOf(id, upTo) {
    const k = id + '|' + upTo;
    if (!expCache.has(k)) {
      try { const x = await call({ op: 'export', id, upTo }); expCache.set(k, { day: x.day, e: x.export }); }
      catch (e) { if (!/\(404\)/.test(e.message)) throw e; expCache.set(k, null); }
    }
    return expCache.get(k);
  }

  async function loadReport() {
    const { cli, month } = page, c = client(cli);
    try {
      const r = await call({ op: 'rapport', clientId: cli, month });
      if (page?.cli !== cli || page.month !== month) return;
      const d = obj(r.report?.data) || {}, p = obj(r.previous?.data) || {}, pi = obj(p.info) || {};
      const mt = maintOf(cli, month)[0];
      const lastDay = isoDate(new Date(+month.slice(0, 4), +month.slice(5, 7), 0));
      const info = { brand: pi.brand || soc(clientSoc(c)).name.toUpperCase(), nom: pi.nom || c.billitName || c.name, adresse: pi.adresse || c.address || '', contact: pi.contact || '',
        date: mt?.date || (month === today().slice(0, 7) ? today() : lastDay), intervenant: mt?.techId ? user(mt.techId).name : me().name, ...(obj(d.info) || {}) };
      R = { cli, month, status: r.report?.status || 'new', rec: r.report, A: null, days: {}, preview: false, fresh: !r.report,
        data: { etat: '', resume: '', inter: '', recos: '', ov: {}, ...d, info, next: { date: '', nature: obj(p.next)?.nature || 'Maintenance préventive mensuelle standard', ...(obj(d.next) || {}) } } };
      if (!obj(R.data.ov)) R.data.ov = {};
      if (R.status !== 'valid' || !obj(R.data.snap)) await compute();
      if (R.fresh) { R.data.inter = monthEntries(cli, month).join('\n'); R.data.recos = autoRecos(R.A).join('\n'); R.data.resume = autoResume(R.A); }
    } catch (e) { R = { cli, month, error: e.message }; }
    if (tab === 'serveurs' && page?.cli === cli && page.month === month) drawReport();
  }

  // Exports utilisés : le dernier reçu jusqu'au jour de l'intervention (20 jours au plus avant), et un autre d'au moins 15 jours plus ancien pour les tendances.
  async function compute() {
    if (!S) await load(true);
    const all = (S?.servers || []).filter(s => s.clientId === R.cli);
    const srvs = all.filter(s => !isNas(s)).sort((a, b) => (a.data?.type === 'virtuel') - (b.data?.type === 'virtuel') || a.name.localeCompare(b.name));
    const date = pd(R.data.info.date) || new Date(), cur = {}, prev = {};
    R.days = {};
    await Promise.all(srvs.map(async s => {
      let x = await exportOf(s.id, isoDate(date));
      // Rien jusqu'à ce jour (agent installé après l'intervention) : le dernier export du mois du rapport.
      if (!x || days(date, pd(x.day)) > 20) { x = await exportOf(s.id, isoDate(new Date(+R.month.slice(0, 4), +R.month.slice(5, 7), 0))); if (x && !x.day.startsWith(R.month)) x = null; }
      if (!x) return;
      cur[s.name] = x.e; R.days[s.name] = x.day;
      const o = await exportOf(s.id, isoDate(new Date(pd(x.day).getTime() - 15 * 86400000)));
      if (o) prev[s.name] = o.e;
    }));
    R.A = analyse(srvs, all.filter(isNas), cur, prev, date);
    R.srvCount = srvs.length;
  }

  // Lignes dans l'ordre du rapport, avec les corrections (ov) : [{ sec, title, kind, info?, rows: [{ id, label, st, com, auto: { st, com } }] }]
  const SECS = [['backup', 'Sauvegardes'], ['maj', 'Mises à jour'], ['nas'], ['sante', 'Santé générale des serveurs'], ['res'], ['secu', 'Sécurité'], ['vm', 'Machines virtuelles'], ['svc', 'Services essentiels']];
  function blocks() {
    const A = R.A, ov = R.data.ov || {};
    const eff = r => { const o = obj(ov[r.id]); return { id: r.id, label: r.label, st: o && RP_ST.includes(o.st) ? o.st : r.st, com: o ? String(o.com ?? '') : r.com, auto: { st: r.st, com: r.com }, hint: r.hint || '', mod: !!o }; };
    const out = [];
    for (const [sec, title] of SECS) {
      if (sec === 'nas') A.nas.forEach(n => out.push({ kind: 'nas', title: `NAS ${n.info.synology ? 'Synology ' : ''}- ${n.info.nom}`, info: n.info, rows: n.rows.map(eff) }));
      else if (sec === 'res') out.push({ kind: 'res', title: 'Ressources et performances', res: A.res });
      else if ((A.rows[sec] || []).length) out.push({ kind: 'checks', title, rows: A.rows[sec].map(eff) });
    }
    return out;
  }
  const etatAuto = bl => ({ Critique: 'Critique', Attention: 'À surveiller' }[worst(bl.flatMap(b => b.rows || []).map(r => r.st).filter(s => ['Critique', 'Attention', 'OK'].includes(s)))] || 'Bon');
  const lines = t => String(t || '').split('\n').map(x => x.replace(/^\s*[-•*]\s*/, '').trim()).filter(Boolean);

  // Copie complète du rapport tel qu'il est imprimé (gardée dans le rapport validé).
  function model() {
    const d = R.data, i = d.info, bl = blocks(), c = client(R.cli), m = R.month;
    const lastDay = new Date(+m.slice(0, 4), +m.slice(5, 7), 0).getDate();
    const show = rows => rows.filter(r => r.st !== 'hide').map(r => ({ label: r.label, st: r.st, com: r.com }));
    return { brand: i.brand || 'ONE LEVEL UP', code: c.name, month: m,
      info: [['Client', i.nom || c.name], ['Site / Adresse', i.adresse || ''], ['Référence du rapport', `MAINT-${c.name.replace(/[^A-Za-z0-9]+/g, '').toUpperCase() || 'CLIENT'}-${m}`],
        ['Période couverte', `01/${m.slice(5, 7)}/${m.slice(0, 4)} – ${pad(lastDay)}/${m.slice(5, 7)}/${m.slice(0, 4)}`], ['Date de l\'intervention', fd(i.date)], ['Intervenant', i.intervenant || ''], ['Contact client', i.contact || '']],
      etat: d.etat || etatAuto(bl), resume: d.resume || autoResume(R.A), inv: R.A.inv, hwLine: R.A.hwLine,
      blocks: bl.map(b => b.kind === 'res' ? { kind: 'res', title: b.title, res: b.res } : { kind: b.kind, title: b.title, info: b.info, rows: show(b.rows) }).filter(b => b.kind === 'res' || b.rows.length),
      inter: lines(d.inter), recos: lines(d.recos), next: { date: d.next?.date ? fd(d.next.date) : '', nature: d.next?.nature || '' },
      sign: { name: i.intervenant || me().name, email: me().email || '' }, days: R.days, at: new Date().toISOString() };
  }

  const stCss = st => ({ OK: 'ok', Attention: 'warn', Critique: 'crit' }[st] || 'todo');
  // Le rapport sur « papier » (aperçu à l'écran et impression) : couleurs et rubriques de l'ancien rapport Word.
  function paper(M) {
    const e = esc, kv = rows => `<table>${rows.map(([k, v]) => `<tr><td class="rp-k">${e(k)}</td><td>${e(v) || '<span class="rp-st-todo">[ à compléter ]</span>'}</td></tr>`).join('')}</table>`;
    const checks = rows => `<table><tr><th style="width:44%">Élément contrôlé</th><th style="width:16%">Statut</th><th>Commentaire</th></tr>
      ${rows.map(r => `<tr><td>${e(r.label)}</td><td class="rp-st rp-st-${stCss(r.st)}">${e(r.st)}</td><td class="rp-c">${e(r.com)}</td></tr>`).join('')}</table>`;
    const all = M.blocks.flatMap(b => b.rows || []), cnt = st => all.filter(r => r.st === st).length;
    const sym = { Bon: ['Bon ✔', 'ok'], 'À surveiller': ['À surveiller ⚠', 'warn'], Critique: ['Critique ✖', 'crit'] }[M.etat] || [M.etat, 'todo'];
    let n = 3;
    const h = t => `<h3>${++n}. ${e(t)}</h3>`;
    return `<div class="rp-paper">
      <div class="rp-brand">${e(M.brand)}</div><div class="rp-sub">Services informatiques — Maintenance &amp; infogérance</div>
      <div class="rp-title">RAPPORT DE MAINTENANCE MENSUELLE</div><div class="rp-mois">${e(M.code)} — ${e(monthName(M.month).replace(/^./, x => x.toUpperCase()))}</div>
      <h3>1. Informations générales</h3>${kv(M.info)}
      <h3>2. Résumé</h3>
      <p><strong>État général de l'infrastructure : </strong><span class="rp-st rp-st-${sym[1]}">${e(sym[0])}</span></p>
      <div class="rp-tiles"><div class="rp-tile rp-t-ok"><b>${cnt('OK')}</b> contrôle(s) en ordre</div><div class="rp-tile rp-t-warn"><b>${cnt('Attention')}</b> à surveiller</div><div class="rp-tile rp-t-crit"><b>${cnt('Critique')}</b> critique(s)</div></div>
      ${String(M.resume || '').split('\n').filter(x => x.trim()).map(x => `<p class="rp-c">${e(x)}</p>`).join('')}
      <h3>3. Inventaire des serveurs et machines virtuelles</h3>
      ${M.hwLine ? `<p class="rp-note">${e(M.hwLine)}</p>` : ''}
      ${M.inv.length ? `<table><tr><th style="width:27%">Nom du serveur / VM</th><th>Rôle</th><th style="width:21%">Système d'exploitation</th><th style="width:17%">Adresse IP</th><th style="width:12%">Statut</th></tr>
        ${M.inv.map(i => `<tr><td>${e(i.nom)}</td><td class="rp-c">${e(i.role)}</td><td class="rp-c">${e(i.os)}</td><td class="rp-c">${e(i.ip)}</td><td class="rp-st rp-st-${stCss(i.st)}">${e(i.st)}</td></tr>`).join('')}</table>` : '<p class="rp-note">Aucun serveur supervisé.</p>'}
      ${M.blocks.map(b => b.kind === 'res' ? `${h(b.title)}<p class="rp-note">Le suivi mois par mois permet d'anticiper les saturations avant qu'elles ne deviennent critiques.</p>
          <table><tr><th style="width:34%">Ressource</th><th>Valeur ce mois</th><th style="width:22%">Valeur précédente</th><th style="width:13%">Tendance</th></tr>
          ${(b.res.length ? b.res : [['[ export manquant ]', '', '', '']]).map(r => `<tr><td>${e(r[0])}</td><td class="rp-c">${e(r[1])}</td><td class="rp-c">${e(r[2])}</td><td class="rp-c">${e(r[3])}</td></tr>`).join('')}</table>`
        : b.kind === 'nas' ? `${h(b.title)}${kv([['Modèle', b.info.modele], ['Version DSM', b.info.dsm], ['Rôle', b.info.role || ' ']])}${checks(b.rows)}`
        : `${h(b.title)}${checks(b.rows)}`).join('')}
      ${h('Interventions correctives durant la maintenance')}
      ${M.inter.length ? `<ul>${M.inter.map(x => `<li>${e(x)}</li>`).join('')}</ul>` : '<p class="rp-c">Aucune intervention corrective n\'a été nécessaire durant cette maintenance.</p>'}
      ${h('Recommandations et points de vigilance')}<p class="rp-note">Actions suggérées pour le mois suivant (préventif, capacité, sécurité…).</p>
      <ul>${(M.recos.length ? M.recos : ['Aucune recommandation particulière ce mois-ci.']).map(x => `<li>${e(x)}</li>`).join('')}</ul>
      ${h('Prochaine intervention')}${kv([['Date prévue', M.next.date], ['Nature', M.next.nature || 'Maintenance préventive mensuelle standard']])}
      <p style="margin-top:16px">Pour toute question concernant ce rapport ou pour planifier une intervention complémentaire, je reste à votre disposition.</p>
      <p>Cordialement,</p><p class="rp-sign">${e(M.sign.name)}${M.sign.email ? '<br>' + e(M.sign.email) : ''}</p></div>`;
  }

  // Impression (→ « Enregistrer au format PDF ») : seul le rapport est imprimé (styles.css, body.printing).
  function printReport(M) {
    let a = $('#printArea');
    if (!a) { a = document.createElement('div'); a.id = 'printArea'; document.body.appendChild(a); }
    const safe = s => String(s || '').replace(/[^\p{L}\p{N} .,'&()\-—]/gu, ' ').trim();      // texte placé dans une règle CSS
    const brand = safe(M.brand).toLowerCase().replace(/(^|\s)\p{L}/gu, x => x.toUpperCase());
    a.innerHTML = `<style>@page { size: A4; margin: 16mm 14mm 18mm;
      @top-right { content: "Rapport de maintenance mensuelle — ${safe(M.code)}"; font: italic 8pt Calibri, Arial, sans-serif; color: #595959; }
      @bottom-center { content: "${brand} — Document confidentiel — Page " counter(page) " / " counter(pages); font: italic 8pt Calibri, Arial, sans-serif; color: #595959; } }</style>` + paper(M);
    const title = document.title;
    document.title = `Rapport de maintenance ${safe(M.code)} ${M.month}`;       // nom proposé pour le PDF
    document.body.classList.add('printing');
    const end = () => { document.body.classList.remove('printing'); a.innerHTML = ''; document.title = title; window.removeEventListener('afterprint', end); };
    window.addEventListener('afterprint', end);
    window.print();
  }

  const saveReport = async status => {
    const d = { ...R.data };
    if (status !== 'valid') delete d.snap;
    const r = await call({ op: 'rapport.save', clientId: R.cli, month: R.month, status, data: d });
    R.rec = r.report; R.status = r.report.status; R.fresh = false; R.dirty = false; R.data = { ...R.data, ...d, snap: d.snap };
    delete reps[R.cli];
  };

  function drawReport() {
    const el = $('#view');
    const cid = page.cli, c = client(cid), head = sub => `<div class="row supv-back"><button type="button" class="btn" id="rpBack">‹ ${esc(c.name)}</button>
      <h2 class="grow">Rapport ${esc(deMonth(page.month))} ${sub || ''}</h2></div>`;
    const back = async () => {
      if ($('#rpForm') && R?.dirty && R.status !== 'valid') { readForm(); try { await saveReport('draft'); toast('Brouillon enregistré'); } catch (e) { return toast(e.message); } }
      goClient(cid);
    };
    if (!R || R.error) {
      el.innerHTML = `${head()}<div class="card"><p class="empty">${R?.error ? 'Rapport indisponible.<br><span class="small">' + esc(R.error) + '</span>' : 'Préparation du rapport : lecture des exports des serveurs…'}</p></div>`;
      $('#rpBack').onclick = back; return;
    }
    // Rapport validé (copie figée) ou aperçu avant validation.
    if (R.status === 'valid' || R.preview) {
      const M = R.status === 'valid' && obj(R.data.snap) ? R.data.snap : model();
      el.innerHTML = `${head(R.status === 'valid' ? '<span class="pill supv-ok">validé</span>' : '<span class="pill supv-warn">aperçu</span>')}
        <div class="row rp-bar">${R.status === 'valid' ? '<button type="button" class="btn" id="rpReopen">Rouvrir pour modifier</button>' : '<button type="button" class="btn" id="rpEdit">‹ Modifier</button>'}
          <button type="button" class="btn primary" id="rpPrint">🖨 Imprimer / PDF</button></div>
        <p class="small muted">Dans la fenêtre d'impression, choisir « Enregistrer au format PDF » pour obtenir le fichier à envoyer au client.</p>${paper(M)}`;
      $('#rpBack').onclick = back;
      $('#rpPrint').onclick = () => printReport(M);
      if ($('#rpEdit')) $('#rpEdit').onclick = () => { R.preview = false; drawReport(); };
      if ($('#rpReopen')) $('#rpReopen').onclick = async ev => {
        ev.target.disabled = true;
        try { await saveReport('draft'); await compute(); drawReport(); toast('Rapport rouvert : les lignes sont recalculées avec les exports'); } catch (e) { ev.target.disabled = false; toast(e.message); }
      };
      return;
    }
    const d = R.data, i = d.info, A = R.A, bl = blocks(), auto = etatAuto(bl);
    const opt = (v, cur, t) => `<option value="${esc(v)}" ${v === cur ? 'selected' : ''}>${esc(t || v)}</option>`;
    const rowHtml = r => `<div class="rp-row ${r.mod ? 'mod' : ''}" data-rid="${esc(r.id)}" data-ast="${esc(r.auto.st)}" data-acom="${esc(r.auto.com)}">
      <div class="rp-lab">${esc(r.label)}</div>
      <select data-rst class="rp-s-${stCss(r.st)}" aria-label="Statut">${RP_ST.map(s => opt(s, r.st, s === 'hide' ? 'Ne pas afficher' : s)).join('')}</select>
      <textarea data-rcom rows="2" placeholder="${esc(r.hint || 'commentaire')}" maxlength="1500" aria-label="Commentaire">${esc(r.com)}</textarea>
      <button type="button" class="btn link rp-reset ${r.mod ? '' : 'hidden'}" title="Revenir à la valeur calculée">↺</button></div>`;
    const used = Object.entries(R.days).map(([n, day]) => `${n} : ${fd(day, false)}`).join(' · ');
    el.innerHTML = `${head(R.rec ? '<span class="pill supv-warn">brouillon</span>' : '<span class="pill">nouveau</span>')}
      <form id="rpForm" autocomplete="off">
      ${A.missing.length || A.actions.length ? `<div class="card"><h2>À faire avant d'envoyer</h2>
        ${A.missing.length ? `<div class="supv-sec" style="margin-top:0">Éléments à compléter</div><ul class="small rp-ul">${A.missing.map(x => `<li>${esc(x)}</li>`).join('')}</ul>` : ''}
        ${A.actions.length ? `<details ${A.missing.length ? '' : 'open'}><summary class="small"><strong>${A.actions.length} action(s) conseillée(s) par les agents</strong> (pour toi, pas imprimées)</summary><ul class="small rp-ul">${A.actions.map(x => `<li>${esc(x)}</li>`).join('')}</ul></details>` : ''}</div>` : ''}
      <div class="card"><h2>Informations générales</h2>
        <div class="grid2"><div><label>Client</label><input id="rpNom" value="${esc(i.nom)}" maxlength="200"></div><div><label>Contact chez le client</label><input id="rpContact" value="${esc(i.contact)}" maxlength="200"></div></div>
        <label>Site / adresse</label><input id="rpAdr" value="${esc(i.adresse)}" maxlength="300">
        <div class="grid2"><div><label>Date de l'intervention</label><input type="date" id="rpDate" value="${esc(i.date)}"></div><div><label>Intervenant</label><input id="rpTech" value="${esc(i.intervenant)}" maxlength="100"></div></div>
        <label>En-tête du rapport</label><input id="rpBrand" value="${esc(i.brand)}" maxlength="80">
        <p class="small muted" style="margin:8px 0 0">${R.srvCount ? (used ? 'Exports utilisés – ' + esc(used) : 'Aucun export dans les 20 jours avant cette date ni plus tard dans le mois.') : 'Ce client n\'a pas de serveur supervisé.'}
          <button type="button" class="btn link" id="rpCalc">↻ Recalculer avec cette date</button></p></div>
      <div class="card"><h2>Résumé</h2>
        <label>État général de l'infrastructure</label><select id="rpEtat">${opt('', d.etat, 'Automatique : ' + auto)}${['Bon', 'À surveiller', 'Critique'].map(x => opt(x, d.etat)).join('')}</select>
        <label>Résumé pour le client <button type="button" class="btn link" id="rpResAuto">Proposer</button></label><textarea id="rpResume" rows="4" maxlength="3500">${esc(d.resume)}</textarea></div>
      ${bl.map(b => b.kind === 'res' ? `<div class="card"><h2>${esc(b.title)}<span class="tot">calculé</span></h2>${b.res.length ? `<table class="tot ro rp-res">${b.res.map(r => `<tr><td>${esc(r[0])}</td><td>${esc(r[1])} <span class="muted">· avant : ${esc(r[2])} · ${esc(r[3])}</span></td></tr>`).join('')}</table>` : '<p class="small muted">Aucun export.</p>'}</div>`
        : `<div class="card"><h2>${esc(b.title)}</h2>${b.kind === 'nas' ? `<p class="small muted" style="margin-top:0">${esc([b.info.modele, b.info.dsm ? 'DSM ' + b.info.dsm : '', b.info.role].filter(Boolean).join(' · '))} — contrôle saisi dans la fiche du NAS.</p>` : ''}${b.rows.map(rowHtml).join('')}</div>`).join('')}
      <div class="card"><h2>Interventions correctives</h2>
        <label>Une ligne par intervention <button type="button" class="btn link" id="rpInterAuto">Reprendre les interventions ${esc(deMonth(R.month))}</button></label>
        <textarea id="rpInter" rows="4" maxlength="3500" placeholder="vide = « Aucune intervention corrective n'a été nécessaire »">${esc(d.inter)}</textarea></div>
      <div class="card"><h2>Recommandations</h2>
        <label>Une ligne par recommandation <button type="button" class="btn link" id="rpRecoAuto">Proposer</button></label><textarea id="rpRecos" rows="4" maxlength="3500">${esc(d.recos)}</textarea></div>
      <div class="card"><h2>Prochaine intervention</h2>
        <div class="grid2"><div><label>Date prévue</label><input type="date" id="rpNextDate" value="${esc(d.next.date)}"></div><div><label>Nature</label><input id="rpNextNat" value="${esc(d.next.nature)}" maxlength="200"></div></div></div>
      <p class="small supv-old hidden" id="rpErr"></p>
      <div class="row rp-bar">${R.rec ? '<button type="button" class="btn danger" id="rpDel" style="margin-right:auto">Supprimer</button>' : ''}
        <button type="button" class="btn" id="rpSave">Enregistrer le brouillon</button><button type="button" class="btn" id="rpPrev">👁 Aperçu</button><button type="button" class="btn primary" id="rpValid">✓ Valider le rapport</button></div>
      </form>`;
    const f = $('#rpForm');
    f.onsubmit = ev => ev.preventDefault();
    f.addEventListener('input', () => { R.dirty = true; });
    f.addEventListener('change', readForm);
    $$('.rp-row', f).forEach(row => {
      const sel = $('[data-rst]', row), com = $('[data-rcom]', row), rs = $('.rp-reset', row);
      const upd = () => { sel.className = 'rp-s-' + stCss(sel.value); const m = sel.value !== row.dataset.ast || com.value.trim() !== row.dataset.acom.trim(); row.classList.toggle('mod', m); rs.classList.toggle('hidden', !m); };
      sel.onchange = upd; com.oninput = upd;
      rs.onclick = () => { sel.value = row.dataset.ast; com.value = row.dataset.acom; upd(); R.dirty = true; readForm(); };
    });
    const fail = t => { const p = $('#rpErr'); p.textContent = t; p.classList.remove('hidden'); };
    const run = fn => async ev => { const b = ev.currentTarget; b.disabled = true; try { await fn(); } catch (e) { fail(e.message); } if (b.isConnected) b.disabled = false; };
    $('#rpBack').onclick = back;
    $('#rpCalc').onclick = run(async () => { readForm(); if (!R.data.info.date) throw new Error('Date de l\'intervention manquante.'); await compute(); drawReport(); toast('Recalculé'); });
    $('#rpResAuto').onclick = () => { $('#rpResume').value = autoResume(A); R.dirty = true; readForm(); };
    $('#rpRecoAuto').onclick = () => { $('#rpRecos').value = autoRecos(A).join('\n'); R.dirty = true; readForm(); };
    $('#rpInterAuto').onclick = () => { const l = monthEntries(cid, R.month); $('#rpInter').value = l.join('\n'); R.dirty = true; readForm(); toast(l.length ? `${l.length} intervention(s) reprise(s)` : 'Aucune intervention avec description ce mois-là'); };
    $('#rpSave').onclick = run(async () => { readForm(); await saveReport('draft'); drawReport(); toast('Brouillon enregistré'); });
    $('#rpPrev').onclick = () => { readForm(); R.preview = true; drawReport(); window.scrollTo(0, 0); };
    $('#rpValid').onclick = () => {
      readForm();
      const left = blocks().flatMap(b => b.rows || []).filter(r => r.st === 'À compléter').map(r => r.label);
      const go = async () => { R.data.snap = model(); await saveReport('valid'); $('#dlg').open && $('#dlg').close(); drawReport(); window.scrollTo(0, 0); toast('Rapport validé'); };
      confirmBox('Valider le rapport ?', left.length ? `${left.length} ligne(s) encore « À compléter » : ${left.join(' ; ')}. Elles seront imprimées telles quelles.` : 'Le rapport est figé tel qu\'il est : il ne changera plus avec les nouveaux exports. Il peut être rouvert.',
        left.length ? 'Valider quand même' : 'Valider', go, () => $('#dlg').close());
    };
    if ($('#rpDel')) $('#rpDel').onclick = () => confirmBox('Supprimer ce brouillon ?', 'Les saisies de ce rapport sont perdues.', 'Supprimer', async () => {
      await call({ op: 'rapport.delete', clientId: cid, month: R.month }); delete reps[cid]; $('#dlg').close(); R = null; goClient(cid); toast('Brouillon supprimé');
    }, () => $('#dlg').close());
  }

  // Formulaire → R.data (appelé à chaque changement : un nouvel affichage de la page ne perd rien).
  function readForm() {
    const f = $('#rpForm'); if (!f || !R?.data) return;
    const v = id => ($(id, f)?.value ?? '').trim(), d = R.data;
    d.info = { brand: v('#rpBrand'), nom: v('#rpNom'), adresse: v('#rpAdr'), contact: v('#rpContact'), date: v('#rpDate'), intervenant: v('#rpTech') };
    d.etat = v('#rpEtat'); d.resume = v('#rpResume'); d.inter = v('#rpInter'); d.recos = v('#rpRecos');
    d.next = { date: v('#rpNextDate'), nature: v('#rpNextNat') };
    d.ov = {};
    $$('.rp-row', f).forEach(row => {
      const st = $('[data-rst]', row).value, com = $('[data-rcom]', row).value.trim();
      if (st !== row.dataset.ast || com !== row.dataset.acom.trim()) d.ov[row.dataset.rid] = { st, com };
    });
  }

  function reset() { S = null; at = 0; err = ''; page = null; R = null; expCache.clear(); Object.keys(reps).forEach(k => delete reps[k]); }

  return { view, clientBox, reset };
})();
