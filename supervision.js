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
    if (typeof tab !== 'undefined' && tab === 'serveurs') draw();
  }

  // ---------- onglet ----------
  function view() {
    draw();
    load(false);
  }

  function draw() {
    const el = $('#view');
    if (!on()) { el.innerHTML = '<div class="card"><p class="empty">Module Supervision non activé pour ta société.</p></div>'; return; }
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
    const line = (s, withClient) => {
      const old = s.agentVersion && s.agentVersion !== S.agentVersion;
      const c = s.counts || {}, d = s.data || {}, ck = checkOf(s);
      const sub = isNas(s) ? ['NAS', [d.brand, d.model].filter(Boolean).join(' '), d.role, ck.date ? 'contrôlé le ' + frDate(ck.date) : 'aucun contrôle saisi'].filter(Boolean).map(esc).join(' · ')
        : `${s.lastSeen ? esc(ago(s.lastSeen)) : s.hasToken ? 'agent prévu, aucun envoi reçu' : 'pas d\'agent installé'}${s.agentVersion ? ' · agent ' + esc(s.agentVersion) + (old ? ' <span class="supv-old">(à mettre à jour : ' + esc(S.agentVersion) + ')</span>' : '') : ''}${s.lastSeen && !s.hasToken ? ' · jeton révoqué' : ''}`;
      return `<li class="tap" data-srv="${esc(s.id)}"><div class="grow"><strong>${isNas(s) ? '🗄 ' : ''}${esc(s.name)}</strong>${withClient ? ` <span class="muted">· ${esc(cname(s.clientId))}</span>` : ''} ${lvPill(s)}
        <div class="small muted">${sub}</div></div>
        <div class="dur small">${c.Critique ? `<span class="pill supv-crit">${c.Critique} critique${c.Critique > 1 ? 's' : ''}</span> ` : ''}${c.Attention ? `<span class="pill supv-warn">${c.Attention} à surveiller</span>` : ''}</div></li>`;
    };
    // Filtre (clic sur un compteur) : une seule liste des serveurs concernés, client indiqué, les critiques d'abord par nombre d'alertes.
    const shown = flt ? all.filter(s => catOf(s) === flt).sort((a, b) => ((b.counts?.Critique || 0) - (a.counts?.Critique || 0)) || ((b.counts?.Attention || 0) - (a.counts?.Attention || 0))
      || cname(a.clientId).localeCompare(cname(b.clientId)) || a.name.localeCompare(b.name)) : all;
    const byClient = new Map();
    shown.forEach(s => { if (!byClient.has(s.clientId)) byClient.set(s.clientId, []); byClient.get(s.clientId).push(s); });
    const groups = [...byClient.entries()].sort((a, b) => cname(a[0]).localeCompare(cname(b[0])));
    const kpi = (k, cls, label) => `<button type="button" class="kpi supv-kpi ${flt === k ? 'on' : ''}" data-flt="${k}" ${n[k] ? '' : 'disabled'} title="${n[k] ? 'Voir la liste' : ''}">
      <div class="v ${cls}">${n[k]}</div><div class="l">${label}</div></button>`;
    el.innerHTML = `${err ? `<div class="notifbar"><span class="ico">⚠</span><span class="txt">Données non actualisées : ${esc(err)}</span></div>` : ''}
      <div class="kpis">
        ${kpi('ok', 'supv-t-ok', 'OK')}${kpi('warn', 'supv-t-warn', 'À surveiller')}${kpi('crit', 'supv-t-crit', 'Critiques')}
        ${kpi('mute', 'supv-t-mute', `Sans contact (> ${MUTE_H} h ou jamais vu) · NAS pas contrôlé`)}
      </div>
      ${flt ? `<div class="card"><h2>${esc(FLT[flt])}<span class="tot">${shown.length} appareil${shown.length > 1 ? 's' : ''}</span></h2>
          <ul class="list">${shown.map(s => line(s, true)).join('')}</ul>
          <div class="actions" style="margin-top:8px"><button type="button" class="btn" id="svAll">Tout afficher</button></div></div>`
        : groups.length ? groups.map(([cid, list]) => `<div class="card"><h2>${esc(cname(cid))}<span class="tot">${list.length} appareil${list.length > 1 ? 's' : ''}</span></h2>
        <ul class="list">${list.map(s => line(s, false)).join('')}</ul></div>`).join('')
        : `<div class="card"><p class="empty">Aucun serveur supervisé.${isAdmin() ? '<br><span class="small">Ajoute un serveur, puis installe son agent, ou ajoute un NAS.</span>' : ''}</p></div>`}
      <div class="row" style="margin:4px 0 16px">
        <button class="btn" id="svRefresh">↻ Actualiser</button>
        ${isAdmin() ? '<button class="btn primary" id="svAdd">+ Serveur</button><button class="btn" id="svAddNas">+ NAS</button><button class="btn" id="svImport" style="margin-left:auto">Importer clients.json</button>' : ''}
      </div>`;
    $$('[data-srv]', el).forEach(li => li.onclick = () => openServer(li.dataset.srv));
    $$('[data-flt]', el).forEach(b => b.onclick = () => { flt = flt === b.dataset.flt ? '' : b.dataset.flt; draw(); });
    if ($('#svAll')) $('#svAll').onclick = () => { flt = ''; draw(); };
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

  function reset() { S = null; at = 0; err = ''; }

  return { view, clientBox, reset };
})();
