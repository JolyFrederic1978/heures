// Module Achats : refacturation des articles des factures fournisseurs (PDF), séparé du pointage.
// - Les PDF sont déposés dans OneDrive Compta/Achats/Entrant ou arrivent dans Billit (Peppol). L'app les lit à l'ouverture,
//   puis au plus toutes les heures, ou sur demande, et les range dans Compta/Achats/Traités/AAAA.
// - Données : TimeSheet/App/achats-data.json, avec sa propre synchro (celle du pointage n'y touche pas).
// - Les articles à 0 € ne sont pas importés.
// - Client : trouvé par le nom de domaine (fiche client). Article Billit : table de correspondance, apprise au fil des imports.
// - Prix de vente = prix d'achat × (1 + marge) ; marge de l'article si elle est définie, sinon celle du fournisseur.
//   Si le prix calculé diffère du prix de l'article dans Billit, l'app propose d'ajuster Billit.
// - Fichiers Excel régénérés dans Compta/Achats : Achats-A-Facturer.xlsx, Archives/Archive-Achats.xlsx, Correspondance-Articles.xlsx.
// Utilise les outils de app.js ($, esc, uid, touch, state, client, persist…) : chargé avant lui, appelé après le démarrage.
const Achats = (() => {
  const CHECK_EVERY = 60 * 60000;     // Entrant et Billit : à l'ouverture de l'app, puis au plus toutes les heures (et sur « Vérifier maintenant »)
  // Marges par défaut données par Fred (25/10/2026). Identifiants fixes : pas de doublon entre appareils.
  const SUPPLIERS_SEED = [{ id: 'sup_ovh', name: 'OVH', margin: 100 }, { id: 'sup_mgk', name: 'MGK', margin: 35 },
    { id: 'sup_lisam', name: 'lisam.cloud', margin: 100 }, { id: 'sup_alescom', name: 'Alescom', margin: '' },
    { id: 'sup_dsd', name: 'DSD', margin: '' }];
  let A = null;                       // données du module
  let lastCheck = 0, busy = false, msg = '';
  let invFind = '', invSup = '';      // recherche et filtre fournisseur des « Factures importées »
  let products = null, productsErr = '', productsMode = '', productsAt = 0;   // articles Billit (mémoire seulement)

  // --- cache local (base IndexedDB à part : n'interfère pas avec celle du pointage) ---
  function idb(mode, fn) {
    return new Promise((resolve, reject) => {
      const req = indexedDB.open('heures-achats', 1);
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
  // hints : contact (ou autre texte de la facture) → client, pour les fournisseurs sans domaine (ex. MGK).
  const empty = () => ({ invoices: [], lines: [], mapping: [], suppliers: [], hints: [], files: [], bdocs: [], published: {}, lastSync: 0 });
  async function load() {
    if (!A) A = { ...empty(), ...(await idb('readonly', s => s.get('data')) || {}) };
    A.bdocs ||= [];
    for (const s of SUPPLIERS_SEED) if (!A.suppliers.some(x => x.id === s.id)) A.suppliers.push({ ...s, updatedAt: 0 });
    return A;
  }
  const save = () => idb('readwrite', s => s.put(A, 'data'));
  async function changed() { await save(); render(); scheduleSync(); }

  const base = () => state.onedrive.folder.split('/').slice(0, -1).concat('Achats').join('/');
  const path = {
    data: () => `${state.onedrive.folder}/App/achats-data.json`,
    inbox: () => `${base()}/Entrant`,
    done: year => `${base()}/Traités/${year}`,
    open: () => `${base()}/Achats-A-Facturer.xlsx`,
    billed: () => `${base()}/Archives/Archive-Achats.xlsx`,
    map: () => `${base()}/Correspondance-Articles.xlsx`
  };

  // --- règles ---
  const alive = arr => arr.filter(x => !x.deleted);
  const inv = id => A.invoices.find(i => i.id === id) || {};
  const mapOf = l => l.mapId ? A.mapping.find(m => m.id === l.mapId && !m.deleted) : null;
  const findMap = (supplier, key) => alive(A.mapping).find(m => m.supplier === supplier && m.key === key);
  const supplierOf = name => alive(A.suppliers).find(s => fold(s.name) === fold(name));
  const clientByDomain = d => !d ? null
    : state.clients.find(c => !c.deleted && (c.domains || []).some(x => d === x || d.endsWith('.' + x))) || null;
  const hintKey = s => fold(s).replace(/\s+/g, ' ').trim();
  const hintOf = (supplier, hint) => hint ? alive(A.hints).find(h => h.supplier === supplier && h.text === hintKey(hint)) : null;
  const clientByHint = (supplier, hint) => { const h = hintOf(supplier, hint); const c = h && state.clients.find(x => x.id === h.clientId && !x.deleted); return c || null; };
  const clientFor = l => clientByDomain(l.domain) || clientByHint(l.supplier, l.hint);
  // Client écrit dans les détails d'une facture (ex. « vedi » chez Alescom) : nom de la fiche ou nom Billit, sans accents ni casse.
  const clientByName = name => {
    const k = fold(name).replace(/[^a-z0-9]+/g, ' ').trim();
    if (!k) return null;
    const cs = state.clients.filter(c => !c.deleted), n = c => [c.name, c.billitName].map(s => fold(s).replace(/[^a-z0-9]+/g, ' ').trim());
    return cs.find(c => n(c).includes(k)) || cs.find(c => n(c).some(x => x && (x.startsWith(k + ' ') || k.startsWith(x + ' ')))) || null;
  };
  const whereOf = l => l.domain || l.hint || '';          // domaine (OVH) ou contact (MGK)
  const isSet = v => v !== '' && v != null;
  const round2 = n => Math.round(n * 100) / 100;
  // Marge appliquée (%) : celle de l'article, sinon celle du fournisseur.
  function marginOf(l, m = mapOf(l)) {
    if (isSet(m?.margin)) return { pct: +m.margin, from: 'article' };
    const s = supplierOf(l.supplier);
    return s && isSet(s.margin) ? { pct: +s.margin, from: 'fournisseur' } : null;
  }
  const saleUnit = (l, m) => { const g = marginOf(l, m); return g ? round2(l.unit * (1 + g.pct / 100)) : null; };
  // Prix figé (billedUnit) une fois la ligne mise sur une facture de l'app : un changement de marge ne la modifie plus.
  const lineUnit = l => l.billedUnit ?? saleUnit(l);
  const saleTotal = l => { const u = lineUnit(l); return u == null ? null : round2(u * l.qty); };
  // N° de facture : celui de la facture de l'app (onglet Facturer) si la ligne y est, sinon celui saisi ici.
  const billRefOf = l => l.billId ? (bill(l.billId)?.ref || '') : (l.billRef || '');
  // todo = client, article Billit ou marge manquant ; ready = prêt à refacturer.
  function lineState(l) {
    if (l.status === 'billed') return 'billed';
    if (l.status === 'ignored' || mapOf(l)?.ignore) return 'ignored';
    return l.clientId && mapOf(l)?.billitRef && saleUnit(l) != null ? 'ready' : 'todo';
  }
  const missing = l => [!l.clientId && 'client', !mapOf(l)?.billitRef && 'article Billit', saleUnit(l) == null && 'marge'].filter(Boolean);
  const clientName = l => l.clientId ? client(l.clientId).name : '';
  const pct = n => `${String(n).replace('.', ',')} %`;
  // Couleur par fournisseur (styles.css : sup-ovh, sup-mgk…) et étiquette colorée.
  const supCls = s => ({ ovh: 'sup-ovh', mgk: 'sup-mgk', 'lisam.cloud': 'sup-lisam', alescom: 'sup-alescom', dsd: 'sup-dsd' })[String(s || '').toLowerCase()] || 'sup-autre';
  const supPill = s => `<span class="pill sup-p ${supCls(s)}">${esc(s)}</span>`;

  // --- articles Billit ---
  const product = ref => ref && products?.find(p => fold(p.Reference) === fold(ref));
  const productPrice = p => p ? +(p.AmountExcl ?? p.UnitPriceExcl ?? 0) : null;
  // Relue si plus vieille que maxAge (ms) : les prix peuvent être changés dans Billit pendant que l'app est ouverte.
  // Les dialogues (article, client, fiche d'un article Billit) demandent une liste fraîche (maxAge 0).
  async function loadProducts(maxAge = 60000) {
    if (!Billit.ready()) { products = null; productsErr = 'Clé API Billit absente (Réglages) : prix Billit non vérifiés.'; return; }
    if (products && productsMode === Billit.mode() && Date.now() - productsAt < maxAge) return;
    try { products = await Billit.products(); productsErr = ''; productsMode = Billit.mode(); productsAt = Date.now(); }
    catch (e) { products = null; productsErr = e.message; }
  }
  // Écart entre le prix calculé et le prix de l'article dans Billit (null si pas d'écart ou pas vérifiable).
  function priceGap(l, m = mapOf(l)) {
    const p = product(m?.billitRef), u = saleUnit(l, m);
    return p && u != null && Math.abs(productPrice(p) - u) > 0.005 ? { p, billit: productPrice(p), calc: u } : null;
  }
  // Champs connus d'un article Billit, à renvoyer en entier lors d'une mise à jour.
  const productBody = p => Object.fromEntries(['ProductID', 'Reference', 'Description', 'AmountExcl', 'VAT', 'Unit', 'GroupID',
    'BuyExcl', 'StockQuantity', 'MinimumBilledQuantity', 'InternalInformation'].filter(k => p[k] != null).map(k => [k, p[k]]));
  // buy : prix d'achat de la dernière facture fournisseur, enregistré aussi dans Billit (BuyExcl).
  async function setBillitPrice(p, price, buy) {
    if (!confirm(`Changer le prix de l'article ${p.Reference} dans Billit (${Billit.mode() === 'live' ? 'réel' : 'test'}) :\n${eur(productPrice(p))} → ${eur(price)} HTVA ?`)) return false;
    try { await Billit.saveProduct({ ...productBody(p), AmountExcl: price, ...(buy != null ? { BuyExcl: buy } : {}) }); await loadProducts(0); toast('Prix mis à jour dans Billit'); return true; }
    catch (e) { alert(e.message); return false; }
  }

  // Lignes sans client dont le domaine est maintenant connu (domaine ajouté dans une fiche client).
  function assignClients() {
    let ch = false;
    for (const l of alive(A.lines)) {
      if (l.clientId || l.status === 'billed' || (!l.domain && !l.hint)) continue;
      const c = clientFor(l);
      if (c) { l.clientId = c.id; touch(l); ch = true; }
    }
    return ch;
  }

  // --- synchronisation OneDrive ---
  const KINDS = ['invoices', 'lines', 'mapping', 'suppliers', 'hints', 'files', 'bdocs'];
  function merge(remote) {
    const before = JSON.stringify(KINDS.map(k => A[k]));
    const m = (a = [], b = []) => {
      const x = new Map(a.map(o => [o.id, o]));
      for (const o of b) { const y = x.get(o.id); if (!y || (o.updatedAt || 0) > (y.updatedAt || 0)) x.set(o.id, o); }
      return [...x.values()];
    };
    for (const k of KINDS) A[k] = m(A[k], remote[k]);
    // Même facture importée sur deux appareils à la fois : on garde la première.
    const seen = new Set();
    for (const i of alive(A.invoices).sort((a, b) => (a.importedAt || 0) - (b.importedAt || 0))) {
      const k = i.supplier + '|' + i.number;
      if (!i.number || !seen.has(k)) { seen.add(k); continue; }
      i.deleted = true; touch(i);
      A.lines.forEach(l => { if (l.invoiceId === i.id && !l.deleted) { l.deleted = true; touch(l); } });
    }
    // Même correspondance créée sur deux appareils : on garde la première.
    const maps = new Map();
    for (const mp of alive(A.mapping).sort((a, b) => (a.createdAt || 0) - (b.createdAt || 0))) {
      const k = mp.supplier + '|' + mp.key, first = maps.get(k);
      if (!first) { maps.set(k, mp); continue; }
      mp.deleted = true; touch(mp);
      A.lines.forEach(l => { if (l.mapId === mp.id) { l.mapId = first.id; touch(l); } });
    }
    return before !== JSON.stringify(KINDS.map(k => A[k]));
  }
  const payload = () => JSON.stringify({ version: 1, ...Object.fromEntries(KINDS.map(k => [k, A[k]])) });

  // Appelé par la synchro du pointage (app.js) quand elle a réussi. force = vérifier le dossier Entrant tout de suite.
  async function sync(force = false) {
    await load();
    if (busy || !OneDrive.signedIn() || !navigator.onLine) return;
    busy = true;
    let redraw = false;
    try {
      let ch = false, cur = await OneDrive.get(path.data());
      if (cur) ch = merge(JSON.parse(cur.data));
      ch = assignClients() || ch;
      let billitErr = '';
      if (force || Date.now() - lastCheck > CHECK_EVERY) {
        lastCheck = Date.now(); ch = (await checkInbox()) || ch;
        // Une erreur Billit n'empêche pas d'enregistrer le reste dans OneDrive.
        try { ch = (await checkBillit()) || ch; } catch (e) { billitErr = 'Billit : ' + e.message; }
      }
      for (let i = 0; i < 3; i++) {
        const body = payload();
        if (cur && hash(cur.data) === hash(body)) break;           // rien de neuf à écrire
        const r = await OneDrive.put(path.data(), body, { etag: cur?.etag, type: 'application/json' });
        if (r.ok) break;
        if (r.status !== 412) throw new Error(`OneDrive (${r.status}) : écriture de achats-data.json impossible.`);
        cur = await OneDrive.get(path.data());                     // modifié par un autre appareil : on refusionne
        if (cur) ch = merge(JSON.parse(cur.data)) || ch;
      }
      await publishExcel();
      A.lastSync = Date.now(); msg = billitErr;
      await save();
      redraw = ch || force;
    } catch (e) {
      msg = e instanceof OneDrive.NeedLogin ? '' : e.message;
      redraw = true;
    } finally { busy = false; }
    if (redraw && tab === 'achats' && !$('#dlg').open) render();
    // Page Aujourd'hui : seul le bandeau est redessiné (pas toute la page, pour ne pas effacer une saisie en cours).
    if (tab === 'today' && $('#achNotif')) todayBanner($('#achNotif'));
  }

  // --- bandeau « nouvelles factures » de la page Aujourd'hui (app.js : <div id="achNotif">) ---
  // Factures importées pas encore vues ; toucher = onglet Achats ; ✕ = marquées vues (sur tous les appareils, via achats-data.json).
  function todayBanner(el) {
    if (!A) { load().then(() => el.isConnected && todayBanner(el)); return; }
    const fresh = alive(A.invoices).filter(i => i.seen === false).sort((a, b) => (b.importedAt || 0) - (a.importedAt || 0));
    if (!fresh.length) { el.innerHTML = ''; return; }
    const ids = new Set(fresh.map(i => i.id));
    const todo = alive(A.lines).filter(l => ids.has(l.invoiceId) && lineState(l) === 'todo').length;
    const ready = alive(A.lines).filter(l => ids.has(l.invoiceId) && lineState(l) === 'ready').length;
    const names = fresh.slice(0, 3).map(i => `${i.supplier} ${i.number || ''}`.trim()).join(', ') + (fresh.length > 3 ? ', …' : '');
    el.innerHTML = `<div class="notifbar tap" id="achNotifGo" title="Ouvrir l'onglet Achats">
      <span class="ico">📥</span>
      <div class="txt"><strong>${fresh.length === 1 ? '1 nouvelle facture' : `${fresh.length} nouvelles factures`}</strong> <span class="small muted">· ${esc(names)}</span>
        <div class="small state">${[todo && `${todo} article(s) à compléter`, ready && `${ready} prêt(s) à refacturer`].filter(Boolean).join(' · ') || 'rien à refacturer'}</div></div>
      <span class="go">›</span>
      <button class="x" id="achNotifX" title="Marquer comme vues">✕</button></div>`;
    $('#achNotifGo', el).onclick = () => { tab = 'achats'; render(); };
    $('#achNotifX', el).onclick = async ev => {
      ev.stopPropagation();
      fresh.forEach(i => { i.seen = true; touch(i); });
      el.innerHTML = '';
      await save(); scheduleSync();
    };
  }

  // --- import des PDF ---
  // Lecteurs disponibles (ex. « ovh,mgk ») : un PDF « non reconnu » est relu quand un lecteur a été ajouté depuis.
  const readersSig = () => Fournisseurs.readers().map(r => r.id).join(',');
  async function checkInbox() {
    const files = (await OneDrive.list(path.inbox())).filter(f => /\.pdf$/i.test(f.name));
    let ch = false;
    // Problèmes signalés pour des fichiers qui ne sont plus dans Entrant : on les oublie.
    for (const rec of alive(A.files)) if (!files.some(f => f.name === rec.name)) { rec.deleted = true; touch(rec); ch = true; }
    for (const f of files) {
      const rec = alive(A.files).find(x => x.name === f.name && x.size === f.size);
      // Déjà signalé (illisible, doublon…) : pas relu, sauf « Réessayer » ou nouveau lecteur depuis l'échec.
      if (rec && rec.status !== 'retry' && !(rec.status === 'error' && rec.readers !== readersSig())) continue;
      const got = await OneDrive.get(`${path.inbox()}/${f.name}`, 'buffer');
      if (!got) continue;
      await importPdf(got.data, f.name, f.size, true);
      ch = true;
    }
    return ch;
  }

  // --- factures reçues dans Billit (Peppol) ---
  // Seulement les nouvelles factures, jamais les anciennes. À la première vérification (par mode Billit), toutes les factures déjà
  // présentes dans Billit sont notées « old » sans être importées ; ensuite, toute facture qui apparaît est nouvelle.
  // (Les dates de Billit n'ont pas de fuseau horaire : une comparaison de dates ne serait pas sûre.)
  // Date de la mise en service, synchronisée (Réglages de l'app) : sert aussi à ne demander à Billit que les factures modifiées depuis.
  const billitFrom = () => state.settings.achatsBillitFrom?.[Billit.mode()] || '';
  // Fournisseur Billit (nom de la société) → nom du lecteur.
  const BILLIT_SUP = [[/\bovh\b/i, 'OVH'], [/\bmgk\b/i, 'MGK'], [/lisam/i, 'lisam.cloud'], [/alescom/i, 'Alescom'], [/\bdsd\b/i, 'DSD']];
  const supOf = o => BILLIT_SUP.find(([re]) => re.test(o.CounterParty?.DisplayName || ''))?.[1];
  // Date « AAAA-MM-JJ » d'un champ Billit (tout autre contenu est ignoré : il est affiché dans l'app).
  const dayOf = s => /^\d{4}-\d{2}-\d{2}/.exec(String(s || ''))?.[0] || '';
  async function checkBillit() {
    if (!Billit.ready()) return false;
    const mode = Billit.mode(), from = billitFrom(), bid = o => `bd_${mode}_${o.OrderID}`;   // même id sur tous les appareils
    // Fournisseurs déjà mis en service dans ce mode (les 3 premiers l'ont été en même temps que la date de départ).
    // Un fournisseur ajouté plus tard (Alescom) : ses factures déjà dans Billit sont notées « old » à sa première vérification.
    const done = state.settings.achatsBillitSups?.[mode] || (from ? ['OVH', 'MGK', 'lisam.cloud'] : []);
    const todo = BILLIT_SUP.map(x => x[1]).filter(s => !done.includes(s));
    if (todo.length) {
      for (const o of await Billit.costInvoices('2000-01-01')) {
        if (todo.includes(supOf(o)) && !A.bdocs.some(x => x.id === bid(o)))
          A.bdocs.push(touch({ id: bid(o), orderId: o.OrderID, mode, supplier: supOf(o), number: o.OrderNumber || '', date: dayOf(o.OrderDate), status: 'old' }));
      }
      state.settings.achatsBillitFrom = { ...state.settings.achatsBillitFrom, [mode]: from || today() };
      state.settings.achatsBillitSups = { ...state.settings.achatsBillitSups, [mode]: [...done, ...todo] };
      state.settingsAt = Date.now();
      await DB.save(state); scheduleSync();
      return true;
    }
    let ch = false;
    for (const o of await Billit.costInvoices(from)) {
      const sup = supOf(o);
      if (!sup) continue;
      const id = bid(o);
      let rec = A.bdocs.find(x => x.id === id);
      // Déjà traitée ; une erreur est réessayée seulement sur « Réessayer » ou quand un lecteur a été ajouté.
      if (rec && rec.status !== 'retry' && !(rec.status === 'error' && rec.readers !== readersSig())) continue;
      if (!rec) A.bdocs.push(rec = { id, orderId: o.OrderID, mode });
      Object.assign(rec, { supplier: sup, number: o.OrderNumber || '', date: dayOf(o.OrderDate), total: o.TotalExcl ?? null });
      let res;
      try { res = await importBillit(o, sup); } catch (e) { res = { status: 'error', msg: e.message }; }
      Object.assign(touch(rec), { text: '', ...res, readers: readersSig() });
      ch = true;
    }
    return ch;
  }
  // Une facture Billit : PDF de la facture, sinon PDF d'origine joint à l'e-facture, sinon lignes de l'e-facture.
  async function importBillit(o, sup) {
    const ok = r => r && !r.error && r.invoice?.supplier === sup;
    const tryPdf = async (buffer, name) => { try { return await Fournisseurs.read(buffer.slice(0), name); } catch (e) { return { error: e.message, rows: [] }; } };
    let pdf = null, r = null;
    if (o.OrderPDF?.FileID) { const f = await Billit.file(o.OrderPDF.FileID); pdf = f.buffer; r = await tryPdf(pdf, f.FileName); }
    const firstErr = r;
    if (!ok(r)) {
      const xmlAtt = ((await Billit.getOrder(o.OrderID))?.Attachments || []).find(a => /\.xml$/i.test(a.FileName || '') || /xml/i.test(a.MimeType || ''));
      const xml = xmlAtt ? new TextDecoder().decode((await Billit.file(xmlAtt.FileID)).buffer) : '';
      for (const e of xml ? Fournisseurs.ublPdfs(xml) : []) {
        const x = await tryPdf(e.buffer, e.name);
        if (ok(x)) { r = x; pdf = e.buffer; break; }
      }
      if (!ok(r) && xml && pdf) { const x = Fournisseurs.fromUbl(xml, sup); if (ok(x)) r = x; }
    }
    if (!ok(r)) return { status: 'error', msg: (r?.error || firstErr?.error || 'Facture non reconnue.') + ' (facture reçue dans Billit)',
      text: (firstErr?.rows || []).map(x => x.cells.join(' | ')).join('\n').slice(0, 30000) };
    const i = r.invoice;
    if (!i.lines.length) return { status: 'error', msg: `Facture ${sup} lue, mais aucun article trouvé.` };
    const dup = alive(A.invoices).find(x => x.supplier === sup && ((i.number && x.number === i.number) || x.billitOrderId === o.OrderID));
    if (dup) return { status: 'known', msg: `déjà importée (${dup.number || 'sans n°'})` };
    const before = A.invoices.length;
    const name = `Facture_${String(i.number || o.OrderNumber || o.OrderID).replace(/[#\\/:*?"<>|]/g, '-')}.pdf`;
    const m = await importPdf(pdf, name, pdf.byteLength, false, { ...r, billitTotal: o.TotalExcl ?? null });
    if (A.invoices.length === before) return { status: 'error', msg: m };
    Object.assign(touch(A.invoices[A.invoices.length - 1]), { billitOrderId: o.OrderID, via: r.reader?.id === 'ubl' ? 'billit-ubl' : 'billit' });
    return { status: 'imported', msg: m };
  }

  // Lit un PDF et crée la facture et ses lignes. fromInbox : le PDF est dans Entrant (on le range dans Traités),
  // sinon il vient de l'appareil (on en dépose une copie dans Traités). Renvoie un message pour l'utilisateur.
  // pre : facture déjà lue (import depuis Billit) ; le PDF n'est alors pas relu.
  async function importPdf(buffer, name, size, fromInbox, pre) {
    let r = pre;
    if (!r) try { r = await Fournisseurs.read(buffer.slice(0), name); } catch (e) { r = { error: 'Lecture du PDF impossible : ' + e.message, rows: [] }; }
    const i = r.invoice;
    const dup = !r.error && i.number && alive(A.invoices).find(x => x.supplier === i.supplier && x.number === i.number);
    const problem = r.error || (dup && `Facture ${i.number} déjà importée le ${frDate(isoDate(new Date(dup.importedAt)))}.`);
    let rec = alive(A.files).find(x => x.name === name && x.size === size);
    if (problem) {
      if (!fromInbox) return `${name} : ${problem}`;
      rec ||= (A.files.push(rec = { id: uid('af'), name, size }), rec);
      Object.assign(touch(rec), { status: dup ? 'duplicate' : 'error', msg: problem, readers: readersSig(), text: (r.rows || []).map(x => x.cells.join(' | ')).join('\n').slice(0, 30000) });
      return `${name} : ${problem}`;
    }
    const year = (i.date || today()).slice(0, 4);
    // Nom dans Traités : « Fournisseur - nom d'origine » (ex. OVH - Facture_FR78898017.pdf). OneDrive peut encore le renommer s'il existe déjà.
    const sup = i.supplier.replace(/[\\/:*?"<>|]/g, '');
    let stored = fold(name).startsWith(fold(`${sup} - `)) ? name : `${sup} - ${name}`;
    if (fromInbox) stored = await OneDrive.move(`${path.inbox()}/${name}`, path.done(year), stored) || stored;
    else {
      const up = await OneDrive.put(`${path.done(year)}/${stored}`, buffer, { type: 'application/pdf' });
      if (!up.ok && up.status !== 412) throw new Error(`OneDrive (${up.status}) : copie de ${name} dans Traités impossible.`);
    }
    if (rec) { rec.deleted = true; touch(rec); }
    // Articles à 0 € : pas pris en compte. Lignes « skip » (Alescom) : frais de transport jamais refacturés,
    // interventions et déplacement repris au pointage (invoice.interventions) au lieu d'être refacturés comme achats.
    const kept = i.lines.filter(l => (l.total || l.unit) && !l.skip);
    const skips = i.lines.filter(l => l.skip), zero = i.lines.filter(l => !l.skip && !(l.total || l.unit)).length;
    const notes = [['transport', 'frais de transport non refacturé(s)'], ['intervention', "ligne(s) d'intervention → pointage"], ['deplacement', 'déplacement → pointage (sur site)']]
      .map(([k, t]) => { const x = skips.filter(l => l.skip === k); return x.length ? `${x.length} ${t} (${eur(x.reduce((s, l) => s + l.total, 0))})` : ''; }).filter(Boolean);
    const invoice = touch({ id: uid('ai'), supplier: i.supplier, number: i.number, date: i.date, totalExcl: i.totalExcl,
      // seen: false = signalée par le bandeau de la page Aujourd'hui jusqu'au ✕ (factures importées avant le 26/10 : pas de champ = déjà vues).
      file: `${path.done(year)}/${stored}`, importedAt: Date.now(), seen: false, skipped: zero, notes, warn: totalWarn(i, r.billitTotal),
      interventions: (i.interventions || []).map(x => ({ id: uid('aw'), ...x, clientId: clientByName(x.client)?.id || null, entryId: null, dismissed: false })) });
    A.invoices.push(invoice);
    for (const l of kept) {
      const m = findMap(i.supplier, l.key), c = clientFor({ ...l, supplier: i.supplier });
      A.lines.push(touch({ id: uid('al'), invoiceId: invoice.id, supplier: i.supplier, key: l.key, label: l.label, detail: l.detail,
        domain: l.domain, hint: l.hint || '', from: l.from || null, to: l.to || null, qty: l.qty, unit: l.unit, total: l.total,
        clientId: c?.id || null, mapId: m?.id || null, status: 'open' }));
    }
    return `${i.supplier} ${i.number || name} : ${kept.length} article(s) importé(s)${invoice.skipped ? `, ${invoice.skipped} à 0 € ignoré(s)` : ''}`
      + `${notes.length ? ', ' + notes.join(', ') : ''}${invoice.interventions.length ? `, ${invoice.interventions.length} intervention(s) à ajouter au pointage` : ''}.`
      + (invoice.warn ? ` ⚠ ${invoice.warn}` : '');
  }
  // Contrôle à l'import : total des articles lus (0 € compris) = total HTVA de la facture, et = total connu par Billit pour une facture reçue dans Billit.
  // Renvoie '' si tout correspond, sinon le problème (la facture est alors signalée « Total à vérifier »).
  function totalWarn(i, billitTotal) {
    const sum = round2(i.lines.reduce((s, l) => s + (l.total || 0), 0));
    const off = t => t != null && !isNaN(t) && Math.abs(sum - t) > 0.02;
    if (i.totalExcl == null && billitTotal == null) return `Total HTVA de la facture introuvable : vérifie que les articles (${eur(sum)}) sont complets.`;
    if (off(i.totalExcl)) return `Total des articles ${eur(sum)} ≠ total HTVA de la facture ${eur(i.totalExcl)} (écart ${eur(round2(sum - i.totalExcl))}).`;
    if (off(billitTotal)) return `Total des articles ${eur(sum)} ≠ total HTVA dans Billit ${eur(billitTotal)} (écart ${eur(round2(sum - billitTotal))}).`;
    return '';
  }
  const toCheck = i => i.warn && !i.warnOk;

  // --- fichiers Excel (lecture seule : tout se modifie dans l'app) ---
  // Colonnes reprises du modèle de Fred (Classeur2.xlsx, 25/10/2026), puis des colonnes d'information.
  const D = d => { if (!d) return null; const [y, m, j] = d.split('-').map(Number); return new Date(Date.UTC(y, m - 1, j)); };
  const HEAD = ['numéro', 'nom du produit', 'fournisseur', 'nom du client', 'quantité', "prix d'achat HTVA", 'référence article dans billit',
    'marge fournisseur', 'marge article', 'prix de vente HTVA', 'domaine / contact', 'du', 'au', 'date facture fournisseur'];
  const W = [14, 45, 13, 18, 10, 16, 24, 17, 14, 17, 20, 11, 11, 14];
  const row = l => {
    const i = inv(l.invoiceId), m = mapOf(l), s = supplierOf(l.supplier);
    return [i.number || null, l.label, l.supplier, clientName(l) || null, l.qty, l.total, m?.billitRef || null,
      s && isSet(s.margin) ? s.margin / 100 : null, isSet(m?.margin) ? m.margin / 100 : null, saleTotal(l),
      whereOf(l) || null, D(l.from), D(l.to), D(i.date)];
  };
  const byClient = (a, b) => (clientName(a) + inv(a.invoiceId).date).localeCompare(clientName(b) + inv(b.invoiceId).date);
  const shift = (cols, n) => cols.map(c => c + n);

  async function publishExcel() {
    const ls = alive(A.lines);
    const open = ls.filter(l => ['todo', 'ready'].includes(lineState(l))).sort(byClient);
    const billed = ls.filter(l => l.status === 'billed').sort((a, b) => (a.billedAt + clientName(a)).localeCompare(b.billedAt + clientName(b)));
    const maps = alive(A.mapping).sort((a, b) => (a.supplier + a.label).localeCompare(b.supplier + b.label));
    const fmt = { dateCols: [11, 12, 13], pctCols: [7, 8], moneyCols: [5, 9] };
    const jobs = [
      ['open', path.open(), () => Xlsx.build([...HEAD, 'statut'],
        open.map(l => [...row(l), lineState(l) === 'ready' ? 'prêt' : 'à compléter : ' + missing(l).join(', ')]),
        { ...fmt, widths: [...W, 28], filter: true }), open.map(row)],
      ['billed', path.billed(), () => Xlsx.build(['facturé le', 'n° facture Billit', ...HEAD],
        billed.map(l => [D(l.billedAt), billRefOf(l) || null, ...row(l)]),
        { dateCols: [0, ...shift(fmt.dateCols, 2)], pctCols: shift(fmt.pctCols, 2), moneyCols: shift(fmt.moneyCols, 2), widths: [12, 16, ...W], filter: true }),
        billed.map(l => [l.billedAt, billRefOf(l), ...row(l)])],
      ['map', path.map(), () => Xlsx.build(['fournisseur', 'nom du produit', 'clé de correspondance', 'référence article dans billit', 'marge fournisseur', 'marge article', 'ne pas refacturer'],
        maps.map(m => { const s = supplierOf(m.supplier); return [m.supplier, m.label, m.key, m.billitRef || null,
          s && isSet(s.margin) ? s.margin / 100 : null, isSet(m.margin) ? m.margin / 100 : null, m.ignore ? 'oui' : null]; }),
        { pctCols: [4, 5], widths: [13, 45, 35, 24, 17, 14, 16], filter: true }),
        [maps.map(m => [m.supplier, m.label, m.key, m.billitRef, m.margin, m.ignore]), alive(A.suppliers).map(s => [s.name, s.margin])]]
    ];
    for (const [k, p, make, data] of jobs) {
      const sig = hash(JSON.stringify(data) + p);
      if (A.published[k] === sig) continue;
      const r = await OneDrive.put(p, await make(), { type: Xlsx.MIME, overwrite: true });
      if (!r.ok) throw new Error(`OneDrive (${r.status}) : écriture de ${p.split('/').pop()} impossible.`);
      A.published[k] = sig;
    }
  }

  // --- écran ---
  const openDlg = () => { if (!$('#dlg').open) $('#dlg').showModal(); };
  const lineLi = l => {
    const st = lineState(l), m = mapOf(l), sale = saleTotal(l), gap = priceGap(l);
    return `<li class="tap" data-line="${l.id}"><div class="grow" style="flex:1;min-width:0">
      <div><strong>${esc(clientName(l) || '— client ?')}</strong>
        ${st === 'todo' ? `<span class="pill warn">${missing(l).join(' + ')} à choisir</span>` : ''}
        ${st === 'ignored' ? '<span class="pill">non refacturé</span>' : ''}
        ${st === 'billed' ? `<span class="pill ok">facturé ${frDate(l.billedAt)}${billRefOf(l) ? ' · ' + esc(billRefOf(l)) : ''}</span>` : ''}
        ${m?.billitRef ? `<span class="pill">${esc(m.billitRef)}</span>` : ''}
        ${gap && st !== 'billed' ? `<span class="pill warn">prix Billit ${eur(gap.billit)}</span>` : ''}</div>
      <div class="small">${esc(l.label)}</div>
      <div class="small muted">${[l.supplier, whereOf(l), l.from && `du ${frDate(l.from)} au ${frDate(l.to)}`, `achat ${l.qty} × ${eur(l.unit)}`].filter(Boolean).map(esc).join(' · ')}</div>
    </div><div class="dur">${sale != null && st !== 'ignored' ? eur(sale) : ''}<div class="small muted" style="font-weight:400">${eur(l.total)}</div></div></li>`;
  };
  // Article facturé via l'onglet Facturer : ouvre la facture (envoi vers Billit, annulation) ; sinon, l'article.
  const bindLines = (root = document) => $$('[data-line]', root).forEach(li => li.onclick = () => {
    const l = A.lines.find(x => x.id === li.dataset.line), b = l.billId && bill(l.billId);
    if (b && !b.deleted) showBill(b); else editLine(l);
  });

  function view() {
    if (!A) { $('#view').innerHTML = '<div class="empty">Chargement…</div>'; load().then(() => tab === 'achats' && render()); return; }
    if (Billit.ready() && !productsErr && (!products || Date.now() - productsAt > 60000)) loadProducts().then(() => tab === 'achats' && !$('#dlg').open && render());
    const ls = alive(A.lines);
    const todo = ls.filter(l => lineState(l) === 'todo').sort(byClient);
    const ready = ls.filter(l => lineState(l) === 'ready');
    const groups = [...new Set(ready.map(l => l.clientId))].map(id => ({ id, ls: ready.filter(l => l.clientId === id) }))
      .sort((a, b) => client(a.id).name.localeCompare(client(b.id).name));
    const invoices = alive(A.invoices).sort((a, b) => (b.date || '').localeCompare(a.date || ''));
    const sups = [...new Set(invoices.map(i => i.supplier))].sort();
    // Texte cherché : n°, fournisseur, date, montant, clients, domaines / contacts et noms des articles.
    const findText = i => fold([i.supplier, i.number, frDate(i.date), i.totalExcl != null ? eur(i.totalExcl) : '', i.totalExcl,
      ...alive(A.lines).filter(l => l.invoiceId === i.id).flatMap(l => [clientName(l), whereOf(l), l.label])].join(' '));
    const problems = alive(A.files), checks = invoices.filter(toCheck);
    const works = invoices.flatMap(i => (i.interventions || []).filter(x => !x.entryId && !x.dismissed).map(x => ({ i, x })));
    const bdone = alive(A.bdocs).filter(b => b.mode === Billit.mode()), bErr = bdone.filter(b => b.status === 'error');
    const gaps = ready.filter(l => priceGap(l)).length;
    $('#view').innerHTML = `
      <div class="kpis"><div class="kpi"><div class="v">${eur(ready.reduce((s, l) => s + saleTotal(l), 0))}</div><div class="l">${ready.length} article(s) prêt(s) à refacturer (HTVA)</div></div>
        <div class="kpi"><div class="v">${todo.length}</div><div class="l">article(s) à compléter</div></div></div>
      ${works.length ? `<div class="card"><h2>Interventions à ajouter au pointage (${works.length})</h2>
        <p class="small muted" style="margin-top:0">Lues sur les factures Alescom (lignes « Intervention ») : elles ne sont pas refacturées comme achats.
          « Ajouter » ouvre l'intervention préremplie ; vérifie-la puis enregistre.</p>
        <ul class="list">${works.map(({ i, x }) => workLi(i, x)).join('')}</ul></div>` : ''}
      ${todo.length ? `<div class="card"><h2>À compléter (${todo.length})</h2><ul class="list">${todo.map(lineLi).join('')}</ul></div>` : ''}
      <div class="card"><h2>Prêt à refacturer</h2>
        <p class="small muted" style="margin-top:0">Ces articles sont proposés dans l'onglet <strong>Facturer</strong>, sur la facture du client (avec ses interventions).</p>
        ${gaps ? `<p class="small"><span class="pill warn">${gaps} article(s) dont le prix Billit diffère du prix calculé : ouvre le client pour ajuster.</span></p>` : ''}
        ${productsErr ? `<p class="small muted">${esc(productsErr)}</p>` : ''}
        ${groups.length ? `<ul class="list">${groups.map(g => `<li class="tap" data-aclient="${g.id}"><div class="grow" style="flex:1;min-width:0"><strong>${esc(client(g.id).name)}</strong>
            <div class="small muted">${g.ls.length} article(s) · achat ${eur(g.ls.reduce((s, l) => s + l.total, 0))}</div>
            <ul class="arts">${g.ls.sort(byClient).map(l => `<li>${supPill(l.supplier)}<span style="flex:1;min-width:0">${[mapOf(l).billitRef, l.label, whereOf(l)].filter(Boolean).map(esc).join(' · ')}${l.qty !== 1 ? ` (× ${esc(l.qty)})` : ''}</span>
              <span class="dur">${eur(saleTotal(l))}</span></li>`).join('')}</ul></div>
            <div class="dur">${eur(g.ls.reduce((s, l) => s + saleTotal(l), 0))}</div></li>`).join('')}</ul>`
          : '<div class="empty">Rien à refacturer pour l\'instant.</div>'}</div>
      <div class="card"><h2>Factures importées (${invoices.length})</h2>
        ${invoices.length ? `<input type="search" id="iFind" placeholder="Rechercher : n°, fournisseur, client, domaine, montant" value="${esc(invFind)}">
          <div class="row" style="margin:8px 0">${['', ...sups].map(s => `<button class="btn chip ${s === invSup ? 'active' : ''}" data-isup="${esc(s)}">${s ? esc(s) : 'Tous'}</button>`).join('')}</div>
          <ul class="list" id="iList">${invoices.map(i => `<li class="tap sup ${supCls(i.supplier)}" data-inv="${i.id}" data-sup="${esc(i.supplier)}" data-find="${esc(findText(i))}"><div class="grow" style="flex:1;min-width:0">${supPill(i.supplier)} <strong>${esc(i.number)}</strong>
            ${toCheck(i) ? '<span class="pill warn">⚠ total à vérifier</span>' : ''}
            <div class="small muted">${[frDate(i.date), `${alive(A.lines).filter(l => l.invoiceId === i.id).length} article(s)`, i.skipped && `${i.skipped} à 0 € ignoré(s)`,
              [...new Set(alive(A.lines).filter(l => l.invoiceId === i.id).map(l => clientName(l) || whereOf(l)).filter(Boolean))].join(', ')].filter(Boolean).map(esc).join(' · ')}</div></div>
            <div class="dur">${i.totalExcl != null ? eur(i.totalExcl) : ''}</div>
            ${i.file ? `<button class="btn chip" data-ipdf="${i.id}" title="Voir la facture">📄 PDF</button>` : ''}</li>`).join('')}</ul>
          <p class="small muted" id="iMore"></p>` : '<div class="empty">Aucune facture importée.</div>'}</div>
      <div class="card"><h2>Factures fournisseurs</h2>
        <p class="small muted" style="margin-top:0">À l'ouverture de l'app, puis toutes les heures et sur « Vérifier maintenant », l'app lit :</p>
        <ul class="small muted" style="margin:0 0 8px;padding-left:18px">
          <li>les PDF déposés dans OneDrive <code>${esc(path.inbox())}</code> ;</li>
          <li>${Billit.ready() ? `les nouvelles factures ${BILLIT_SUP.map(x => x[1]).join(', ').replace(/, ([^,]+)$/, ' et $1')} reçues dans Billit${billitFrom() ? ` depuis le ${frDate(billitFrom())}` : ''}
            (${Billit.mode() === 'live' ? 'réel' : 'test'}, lecture seule ; les anciennes ne sont pas reprises)${bdone.length
              ? ` : ${bdone.filter(b => b.status === 'imported').length} importée(s), ${bdone.filter(b => b.status === 'known').length} déjà présente(s)` : ''}.`
            : 'les factures reçues dans Billit : clé API Billit absente (Réglages).'}</li></ul>
        <p class="small muted">Les factures sont rangées dans <code>${esc(base())}/Traités</code>. Les articles à 0 € sont ignorés.${A.lastSync ? ` Dernière vérification : ${new Date(lastCheck || A.lastSync).toLocaleTimeString('fr-BE', { hour: '2-digit', minute: '2-digit' })}.` : ''}</p>
        ${msg ? `<p class="small"><span class="pill warn">${esc(msg)}</span></p>` : ''}
        ${checks.length ? `<p class="small" style="margin:0 0 4px"><strong style="color:var(--warn)">⚠ ${checks.length} facture(s) dont le total des articles ne correspond pas au total de la facture :</strong>
            ouvre-la, compare avec le PDF, puis corrige ou confirme.</p>
          <ul class="list" style="margin-bottom:10px">${checks.map(i => `<li class="tap" data-inv="${i.id}"><div class="grow" style="flex:1;min-width:0">${supPill(i.supplier)} <strong>${esc(i.number)}</strong>
            <div class="small" style="color:var(--warn)">${esc(i.warn)}</div></div></li>`).join('')}</ul>` : ''}
        <div class="row"><button class="btn primary" id="aCheck" ${busy ? 'disabled' : ''}>🔄 Vérifier maintenant</button>
          <button class="btn" id="aPick">📄 Importer un PDF…</button><input type="file" id="aFile" accept="application/pdf,.pdf" multiple hidden></div>
        ${bErr.length ? `<ul class="list" style="margin-top:10px">${bErr.map(b => `<li><div class="grow" style="flex:1;min-width:0">${supPill(b.supplier)} <strong>${esc(b.number || 'sans n°')}</strong> <span class="small muted">${frDate(b.date)} · Billit</span>
            <div class="small"><span class="pill warn">non lu</span> ${esc(b.msg)}</div></div>
            ${b.text ? `<button class="btn chip" data-btext="${b.id}">Texte</button>` : ''}<button class="btn chip" data-bretry="${b.id}">Réessayer</button></li>`).join('')}</ul>` : ''}
        ${problems.length ? `<ul class="list" style="margin-top:10px">${problems.map(p => `<li><div class="grow" style="flex:1;min-width:0"><strong>${esc(p.name)}</strong>
            <div class="small"><span class="pill warn">${p.status === 'duplicate' ? 'doublon' : 'non lu'}</span> ${esc(p.msg)}</div></div>
            <button class="btn chip" data-ptext="${p.id}">Texte</button>${p.status === 'error' ? `<button class="btn chip" data-pretry="${p.id}">Réessayer</button>` : ''}</li>`).join('')}</ul>
          <p class="small muted">Retire du dossier Entrant les fichiers qui ne sont pas à traiter : ils disparaîtront de cette liste.</p>` : ''}
      </div>`;

    $('#aCheck').onclick = () => { msg = ''; const p = sync(true); render(); p.then(() => toast('Vérification terminée')); };
    $('#aPick').onclick = () => $('#aFile').click();
    $('#aFile').onchange = async ev => {
      const files = [...ev.target.files]; ev.target.value = '';
      if (!OneDrive.signedIn()) return alert('Connexion OneDrive requise.');
      const out = [];
      for (const f of files) {
        try { out.push(await importPdf(await f.arrayBuffer(), f.name, f.size, false)); } catch (e) { out.push(`${f.name} : ${e.message}`); }
      }
      await changed(); alert(out.join('\n'));
    };
    $$('[data-ptext]').forEach(b => b.onclick = () => showText(A.files.find(p => p.id === b.dataset.ptext)));
    // Factures Billit non lues : texte extrait, réessayer.
    $$('[data-btext]').forEach(b => b.onclick = () => { const d = A.bdocs.find(x => x.id === b.dataset.btext); showText({ name: `${d.supplier} ${d.number}`, text: d.text }); });
    $$('[data-bretry]').forEach(b => b.onclick = async () => { const d = A.bdocs.find(x => x.id === b.dataset.bretry); d.status = 'retry'; touch(d); await save(); sync(true); });
    $$('[data-pretry]').forEach(b => b.onclick = async () => { const p = A.files.find(x => x.id === b.dataset.pretry); p.status = 'retry'; touch(p); await save(); sync(true); });
    $$('[data-aclient]').forEach(li => li.onclick = () => billClient(li.dataset.aclient));
    $$('[data-inv]').forEach(li => li.onclick = () => showInvoice(inv(li.dataset.inv)));
    $$('[data-ipdf]').forEach(b => b.onclick = ev => { ev.stopPropagation(); showPdf(inv(b.dataset.ipdf)); });
    // Recherche / filtre fournisseur, sans redessiner l'écran. Sans critère : les 15 plus récentes.
    const filterInv = () => {
      const words = fold(invFind).split(/\s+/).filter(Boolean), all = $$('#iList > li');
      const hits = all.filter(li => (!invSup || li.dataset.sup === invSup) && words.every(w => li.dataset.find.includes(w)));
      const shown = words.length || invSup ? hits : hits.slice(0, 15);
      all.forEach(li => li.classList.toggle('hidden', !shown.includes(li)));
      $('#iMore').textContent = !hits.length ? 'Aucune facture trouvée.' : shown.length < hits.length ? `${shown.length} plus récentes sur ${hits.length} : utilise la recherche pour les autres.` : '';
    };
    if ($('#iFind')) {
      $('#iFind').oninput = ev => { invFind = ev.target.value; filterInv(); };
      $$('[data-isup]').forEach(b => b.onclick = () => {
        invSup = b.dataset.isup;
        $$('[data-isup]').forEach(x => x.classList.toggle('active', x === b));
        filterInv();
      });
      filterInv();
    }
    bindLines();
    bindWorks();
  }

  // --- interventions lues sur une facture (Alescom) ---
  const workLi = (i, x) => {
    const c = x.clientId && state.clients.find(k => k.id === x.clientId && !k.deleted);
    const st = x.entryId ? '<span class="pill ok">ajoutée au pointage</span>' : x.dismissed ? '<span class="pill">ignorée</span>' : '';
    return `<li><div class="grow" style="flex:1;min-width:0">${supPill(i.supplier)} <strong>${frDate(x.date)}</strong>
        ${x.start ? `${esc(x.start)}–${esc(x.end)}` : ''} (${fmtDur(x.minutes || 0)}) · ${c ? esc(c.name) : `<span class="pill warn">client « ${esc(x.client || '?')} » à choisir</span>`}
        ${x.onSite ? '<span class="pill">sur site</span>' : '<span class="pill">à distance</span>'} ${st}
        <div class="small muted">${[x.notes, x.tech && `par ${x.tech}`, `facture ${i.number}`].filter(Boolean).map(esc).join(' · ')}</div></div>
      ${!x.entryId && !x.dismissed ? `<button class="btn chip" data-wadd="${i.id}|${x.id}">➕ Ajouter</button><button class="btn chip" data-wskip="${i.id}|${x.id}">Ignorer</button>` : ''}</li>`;
  };
  function bindWorks(root = document) {
    const get = v => { const [a, b] = v.split('|'), i = inv(a); return { i, x: (i.interventions || []).find(w => w.id === b) }; };
    $$('[data-wadd]', root).forEach(btn => btn.onclick = ev => {
      ev.preventDefault();
      const { i, x } = get(btn.dataset.wadd);
      const preset = { date: x.date, start: x.start || '', end: x.end || '', minutes: x.minutes || 0, pauseMin: 0, onSite: !!x.onSite, alescom: false, install: '',
        description: [x.notes, `(${[i.supplier, x.tech].filter(Boolean).join(' – ')}, facture ${i.number})`].filter(Boolean).join(' ') };
      if (x.clientId && state.clients.some(c => c.id === x.clientId && !c.deleted)) preset.clientId = x.clientId;
      openEntry(null, preset, false, async e => { x.entryId = e.id; touch(i); await save(); });
    });
    $$('[data-wskip]', root).forEach(btn => btn.onclick = async ev => {
      ev.preventDefault();
      const { i, x } = get(btn.dataset.wskip);
      x.dismissed = true; touch(i);
      if ($('#dlg').open) $('#dlg').close();
      await changed(); toast('Intervention ignorée');
    });
  }

  function showText(p) {
    const f = $('#dlgForm');
    f.innerHTML = `<h2>${esc(p.name)}</h2><p class="small muted">Texte extrait du PDF (colonnes séparées par « | »), pour ajuster le lecteur.</p>
      <pre class="recap" style="max-height:60vh;overflow:auto">${esc(p.text || '(aucun texte)')}</pre>
      <div class="actions"><button class="btn" id="tCopy">Copier</button><button class="btn" value="cancel">Fermer</button></div>`;
    $('#tCopy').onclick = () => copy(p.text || '');
    openDlg();
  }

  const productOptions = () => (products || []).map(p => `<option value="${esc(p.Reference)}">${esc(p.Description || '')} · ${eur(productPrice(p))}</option>`).join('');

  // Une ligne : client + correspondance avec l'article Billit (valable pour tous les articles du même type).
  // draft : valeurs saisies à conserver (retour de la création d'article).
  async function editLine(l, draft = null) {
    await loadProducts(0);
    const m = mapOf(l) || findMap(l.supplier, l.key);
    const i = inv(l.invoiceId), s = supplierOf(l.supplier);
    const d = draft || { clientId: l.clientId, ref: m?.billitRef || '', margin: m?.margin ?? '', ignore: !!m?.ignore, ignoreLine: l.status === 'ignored' };
    const domainOwner = clientByDomain(l.domain), hintKnown = hintOf(l.supplier, l.hint);
    const billed = l.status === 'billed';
    const dlg = $('#dlg'), f = $('#dlgForm');
    f.innerHTML = `<h2>Article ${esc(l.supplier)}</h2>
      <p class="small muted">Facture ${esc(i.number)} du ${frDate(i.date)}</p>
      <p><strong>${esc(l.label)}</strong>${l.detail ? `<br><span class="small">${esc(l.detail)}</span>` : ''}<br>
        <span class="small muted">${[whereOf(l), l.from && `du ${frDate(l.from)} au ${frDate(l.to)}`, `achat ${l.qty} × ${eur(l.unit)} = ${eur(l.total)}`].filter(Boolean).map(esc).join(' · ')}</span></p>
      <label>Client</label><select id="lClient" ${billed ? 'disabled' : ''}><option value="">— choisir —</option>${clientOptions(d.clientId)}
        <option value="__new">+ Nouveau client…</option></select>
      ${l.domain && !domainOwner ? `<label class="check"><input type="checkbox" id="lDom" checked> Ajouter ${esc(l.domain)} à la fiche de ce client</label>` : ''}
      ${!l.domain && l.hint && !hintKnown ? `<label class="check"><input type="checkbox" id="lHint" checked> Retenir « ${esc(l.hint)} » pour ce client (factures ${esc(l.supplier)})</label>` : ''}
      <div class="card" style="margin:12px 0 0;padding:10px">
        <strong class="small">Pour tous les articles « ${esc(l.label)} » de ${esc(l.supplier)}</strong>
        <label>Article Billit (référence)</label>
        <div class="row"><input id="mRef" class="grow" style="flex:1" value="${esc(d.ref)}" list="billitRefs" placeholder="ex. int-004" autocapitalize="off">
          <button class="btn" id="mNew" ${Billit.ready() ? '' : 'disabled'}>+ Créer</button></div>
        <datalist id="billitRefs">${productOptions()}</datalist>
        <label>Marge de l'article en % (vide = marge fournisseur : ${s && isSet(s.margin) ? pct(s.margin) : 'non définie'})</label>
        <input type="number" step="0.1" id="mMargin" value="${esc(d.margin)}">
        <p class="small" id="mCalc" style="margin:8px 0 0"></p>
        <label class="check"><input type="checkbox" id="mIgnore" ${d.ignore ? 'checked' : ''}> Ne jamais refacturer ce type d'article</label>
      </div>
      <label class="check"><input type="checkbox" id="lIgnore" ${d.ignoreLine ? 'checked' : ''} ${billed ? 'disabled' : ''}> Ne pas refacturer cette ligne-ci</label>
      ${productsErr ? `<p class="small muted">${esc(productsErr)}</p>` : ''}
      <div class="actions"><button class="btn" value="cancel">Annuler</button><button class="btn primary" id="lSave">Enregistrer</button></div>`;
    const read = () => ({ clientId: ($('#lClient').value !== '__new' && $('#lClient').value) || null, ref: $('#mRef').value.trim(), margin: $('#mMargin').value.trim(),
      ignore: $('#mIgnore').checked, ignoreLine: $('#lIgnore').checked });
    // « + Nouveau client… » : fiche client (avec choix dans les clients Billit pas encore reliés), domaine de la ligne proposé,
    // puis retour sur cet article avec le nouveau client choisi (ou sans changement avec « Retour »).
    $('#lClient').onchange = ev => {
      if (ev.target.value !== '__new') return;
      const v = { ...read(), clientId: d.clientId };
      editClient(null, { preset: { domains: l.domain ? [l.domain] : [] },
        onSaved: c => editLine(l, { ...v, clientId: c.id }), onCancel: () => editLine(l, v) });
    };
    // Prix calculé et comparaison avec Billit, mis à jour pendant la saisie.
    const calc = () => {
      const v = read(), tmp = { billitRef: v.ref, margin: v.margin };
      const g = marginOf(l, tmp), u = saleUnit(l, tmp), p = product(v.ref);
      const gap = priceGap(l, tmp);
      $('#mCalc').innerHTML = (u == null ? '<span class="pill warn">Aucune marge : définis la marge de l\'article ou du fournisseur.</span>'
        : `Prix de vente : <strong>${eur(u)}</strong> HTVA (marge ${g.from} ${pct(g.pct)})`)
        + (!v.ref ? '' : !products ? '' : !p ? `<br><span class="pill warn">${esc(v.ref)} n'existe pas dans Billit : utilise « + Créer ».</span>`
          : `<br>Billit : ${esc(p.Description || '')} · ${eur(productPrice(p))}` + (gap ? ` <span class="pill warn">≠ prix calculé</span> <button class="btn chip" id="mFix">Mettre ${eur(gap.calc)} dans Billit</button>` : ' <span class="pill ok">✓ identique</span>'));
      if ($('#mFix')) $('#mFix').onclick = async () => { if (await setBillitPrice(gap.p, gap.calc, l.unit)) calc(); };
    };
    ['#mRef', '#mMargin'].forEach(s => $(s).oninput = calc);
    calc();
    $('#mNew').onclick = () => {
      const v = read();
      createArticle({ Reference: v.ref, Description: l.label, AmountExcl: saleUnit(l, { margin: v.margin }) ?? '', BuyExcl: l.unit },
        ref => editLine(l, { ...v, ref: ref || v.ref }));
    };
    $('#lSave').onclick = async ev => {
      ev.preventDefault();
      const v = read();
      if (v.ref && products && !product(v.ref) && !confirm(`L'article ${v.ref} n'existe pas dans Billit (${Billit.mode() === 'live' ? 'réel' : 'test'}).\n\nEnregistrer quand même ? (Annuler pour le créer avec « + Créer »)`)) return;
      let mp = m;
      if (!mp && (v.ref || v.margin !== '' || v.ignore)) A.mapping.push(mp = { id: uid('am'), supplier: l.supplier, key: l.key, label: l.label, createdAt: Date.now() });
      if (mp) {
        Object.assign(touch(mp), { billitRef: v.ref, margin: v.margin === '' ? '' : +v.margin, ignore: v.ignore });
        alive(A.lines).forEach(x => { if (!x.mapId && x.supplier === l.supplier && x.key === l.key) { x.mapId = mp.id; touch(x); } });
      }
      let clientsChanged = false;
      if (!billed) {
        const c = v.clientId && state.clients.find(x => x.id === v.clientId);
        if (c && $('#lDom')?.checked) { c.domains = [...(c.domains || []), l.domain]; touch(c); clientsChanged = true; }
        if (c && $('#lHint')?.checked) A.hints.push(touch({ id: uid('ah'), supplier: l.supplier, text: hintKey(l.hint), label: l.hint, clientId: c.id }));
        Object.assign(touch(l), { clientId: v.clientId, mapId: mp?.id || null, status: v.ignoreLine ? 'ignored' : 'open' });
        assignClients();                     // autres lignes du même domaine / contact
      }
      dlg.close();
      await save();
      if (clientsChanged) await persist(); else await changed();
    };
    openDlg();
  }

  // Catégories Billit (GroupID / GroupName des articles existants : l'API ne les liste pas à part).
  function categories() {
    const m = new Map();
    for (const p of products || []) if (p.GroupID) {
      const g = m.get(p.GroupID) || m.set(p.GroupID, { id: p.GroupID, name: p.GroupName || `Catégorie ${p.GroupID}`, refs: [] }).get(p.GroupID);
      g.refs.push(p.Reference);
    }
    return [...m.values()].sort((a, b) => a.name.localeCompare(b.name));
  }
  // Référence proposée (règle de Fred) : 3 premières lettres de la catégorie, « - », puis le numéro suivant disponible
  // pour ce préfixe (ex. « logiciel » : LOG-001, log-002, log-003 → log-004 ; « matériel » → mat-…). Numéro sur 3 chiffres.
  // Exceptions voulues par Fred : ces catégories gardent leurs anciens préfixes.
  const REF_PREFIX = { 'service internet': 'int', reseau: 'net' };
  const refPrefix = g => REF_PREFIX[fold(g.name).trim()] || fold(g.name).replace(/[^a-z]/g, '').slice(0, 3);
  function nextRef(g) {
    const p = refPrefix(g);
    if (!p) return '';
    const used = (products || []).map(x => new RegExp(`^${p}-(\\d+)$`, 'i').exec(x.Reference || '')?.[1]).filter(Boolean).map(Number);
    return `${p}-${String(Math.max(0, ...used) + 1).padStart(3, '0')}`;
  }

  // Création d'un article dans Billit. back(ref) : retour à l'écran précédent (ref = article créé, ou rien si annulé).
  function createArticle(pre = {}, back = backToArticles) {
    const f = $('#dlgForm');
    f.innerHTML = `<h2>Nouvel article Billit</h2>
      <p class="small muted">Créé dans Billit en mode ${Billit.mode() === 'live' ? 'réel' : 'test (sandbox)'}.</p>
      <label>Catégorie Billit (obligatoire)</label><select id="pGroup"><option value="">— choisir —</option>
        ${categories().map(g => `<option value="${g.id}" ${String(g.id) === String(pre.GroupID ?? '') ? 'selected' : ''}>${esc(g.name)} (${g.refs.length})</option>`).join('')}</select>
      <label>Référence</label><input id="pRef" value="${esc(pre.Reference)}" placeholder="ex. int-005" autocapitalize="off">
      <p class="small muted" id="pRefHint" style="margin:4px 0 0"></p>
      <label>Description</label><input id="pDesc" value="${esc(pre.Description)}">
      <div class="grid2"><div><label>Prix d'achat HTVA</label><input type="number" step="0.01" id="pBuy" value="${esc(pre.BuyExcl)}"></div>
        <div><label>Prix de vente HTVA</label><input type="number" step="0.01" id="pPrice" value="${esc(pre.AmountExcl)}"></div></div>
      <div class="grid2"><div><label>TVA %</label><input type="number" step="1" id="pVat" value="${esc(pre.VAT ?? 21)}"></div>
        <div><label>Marge</label><p class="small muted" id="pMargin" style="margin:10px 0 0"></p></div></div>
      <div class="actions"><button class="btn" id="pBack">Retour</button><button class="btn primary" id="pSave">Créer dans Billit</button></div>`;
    const showMargin = () => {
      const buy = +$('#pBuy').value, sale = +$('#pPrice').value;
      $('#pMargin').textContent = $('#pBuy').value !== '' && $('#pPrice').value !== '' && buy > 0 ? `${pct(Math.round((sale - buy) / buy * 1000) / 10)} (${eur(sale - buy)})` : '—';
    };
    $('#pBuy').oninput = $('#pPrice').oninput = showMargin;
    showMargin();
    // Catégorie choisie : référence proposée (3 lettres de la catégorie + numéro suivant), si le champ est vide ou déjà proposé.
    let suggested = '';
    const onGroup = () => {
      const g = categories().find(x => String(x.id) === $('#pGroup').value), next = g && nextRef(g);
      $('#pRefHint').textContent = !categories().length ? 'Aucune catégorie trouvée dans Billit (liste des articles non chargée ?).'
        : g ? `Articles de cette catégorie : ${g.refs.slice(0, 8).join(', ')}${g.refs.length > 8 ? '…' : ''}` : 'Choisis la catégorie : la référence est proposée.';
      if (next && (!$('#pRef').value.trim() || $('#pRef').value === suggested)) $('#pRef').value = suggested = next;
    };
    $('#pGroup').onchange = onGroup;
    onGroup();
    $('#pBack').onclick = () => back();
    $('#pSave').onclick = async ev => {
      ev.preventDefault();
      const item = { Reference: $('#pRef').value.trim(), Description: $('#pDesc').value.trim(), AmountExcl: +$('#pPrice').value, VAT: +$('#pVat').value,
        ...($('#pBuy').value !== '' ? { BuyExcl: +$('#pBuy').value } : {}), GroupID: +$('#pGroup').value };
      if (!$('#pGroup').value) return alert('Choisis la catégorie Billit de l\'article.');
      if (!item.Reference || !item.Description) return alert('Référence et description requises.');
      if (!(item.AmountExcl >= 0) || $('#pPrice').value === '') return alert('Prix requis.');
      await loadProducts(0);
      if (product(item.Reference)) return alert(`L'article ${item.Reference} existe déjà dans Billit.`);
      const btn = $('#pSave'); btn.disabled = true;
      try {
        await Billit.saveProduct(item);
        await loadProducts(0);
        if (!product(item.Reference)) throw new Error("Billit n'a pas renvoyé d'erreur, mais l'article n'apparaît pas dans la liste. Vérifie dans Billit.");
        toast(`Article ${item.Reference} créé dans Billit`);
        back(item.Reference);
      } catch (e) { alert(e.message); if (btn.isConnected) btn.disabled = false; }
    };
    openDlg();
  }

  // Refacturation d'un client : pour l'instant, la facture Billit est faite à la main ; ici on la note.
  function billClient(clientId) {
    const ls = alive(A.lines).filter(l => l.clientId === clientId && lineState(l) === 'ready').sort(byClient);
    const sel = new Set(ls.map(l => l.id));
    const dlg = $('#dlg'), f = $('#dlgForm');
    const draw = () => {
      const chosen = ls.filter(l => sel.has(l.id));
      f.innerHTML = `<h2>Achats à refacturer : ${esc(client(clientId).name)}</h2>
        <p class="small muted">Achat ${eur(chosen.reduce((s, l) => s + l.total, 0))} · vente <strong>${eur(chosen.reduce((s, l) => s + saleTotal(l), 0))}</strong> HTVA</p>
        <ul class="list small">${ls.map(l => { const g = marginOf(l), gap = priceGap(l); return `<li><input type="checkbox" data-sel="${l.id}" ${sel.has(l.id) ? 'checked' : ''}>
          <div class="grow" style="flex:1;min-width:0"><strong>${esc(mapOf(l).billitRef)}</strong> ${esc(l.label)}
            <div class="muted">${[whereOf(l), l.from && `du ${frDate(l.from)} au ${frDate(l.to)}`, `${l.qty} × ${eur(saleUnit(l))} (achat ${eur(l.unit)} + ${pct(g.pct)} ${g.from})`].filter(Boolean).map(esc).join(' · ')}</div>
            ${gap ? `<div><span class="pill warn">Prix Billit ${eur(gap.billit)}</span> <button class="btn chip" data-fix="${l.id}">Mettre ${eur(gap.calc)} dans Billit</button></div>` : ''}</div>
          <div class="dur">${eur(saleTotal(l))}</div></li>`; }).join('')}</ul>
        <div class="actions"><button class="btn" value="cancel">Fermer</button>
          <button class="btn primary" id="bMark" ${chosen.length ? '' : 'disabled'}>Marquer ${chosen.length} article(s) comme facturé(s)</button></div>`;
      $$('[data-sel]', f).forEach(c => c.onchange = () => { c.checked ? sel.add(c.dataset.sel) : sel.delete(c.dataset.sel); draw(); });
      $$('[data-fix]', f).forEach(b => b.onclick = async () => {
        const l = A.lines.find(x => x.id === b.dataset.fix), gap = priceGap(l);
        if (gap && await setBillitPrice(gap.p, gap.calc, l.unit)) draw();
      });
      $('#bMark').onclick = async ev => {
        ev.preventDefault();
        // Dernier contrôle des prix Billit (ils ont pu changer depuis l'ouverture) : proposition d'ajustement, article par article.
        await loadProducts(0);
        const asked = new Set();
        for (const l of chosen) {
          const gap = priceGap(l);
          if (!gap || asked.has(gap.p.ProductID)) continue;
          asked.add(gap.p.ProductID);
          await setBillitPrice(gap.p, gap.calc, l.unit);
        }
        const ref = prompt('N° de la facture Billit (optionnel) :', '');
        if (ref === null) return;
        chosen.forEach(l => Object.assign(touch(l), { status: 'billed', billedAt: today(), billRef: ref.trim() }));
        dlg.close(); await changed(); toast('Articles marqués comme facturés : déplacés vers l\'archive');
      };
    };
    draw();
    openDlg();
    loadProducts(0).then(() => { if (dlg.open && $('#bMark')) draw(); });
  }

  // PDF d'origine (rangé dans Traités/AAAA) : pages affichées en images + téléchargement. Gardé en mémoire pendant la session.
  const pdfCache = new Map();
  async function showPdf(i) {
    const dlg = $('#dlg'), f = $('#dlgForm'), name = i.file.split('/').pop();
    let zoom = 100;
    f.innerHTML = `<div class="row"><h2 class="grow">${esc(i.supplier)} ${esc(i.number)}</h2>
        <button class="btn chip" id="pOut" title="Réduire">−</button><span class="small" id="pZoom" style="min-width:44px;text-align:center">100 %</span>
        <button class="btn chip" id="pIn" title="Agrandir">+</button><button class="btn chip" id="pFit">Ajuster</button>
        <button class="btn chip" id="pWide">⛶ Grand écran</button></div>
      <p class="small muted" id="pMsg">Chargement de ${esc(name)}…</p>
      <div id="pPages" style="max-height:65vh;overflow:auto"></div>
      <div class="actions"><button class="btn" id="pDl" disabled>⬇ Télécharger</button><button class="btn" id="pBack">Retour</button>
        <button class="btn" value="cancel">Fermer</button></div>`;
    const setZoom = z => {
      zoom = Math.min(300, Math.max(50, z));
      $('#pZoom').textContent = `${zoom} %`;
      $$('#pPages canvas').forEach(c => c.style.width = `${zoom}%`);
    };
    const setWide = on => { dlg.classList.toggle('wide', on); if ($('#pWide')) $('#pWide').textContent = on ? '⛶ Réduire la fenêtre' : '⛶ Grand écran'; };
    $('#pOut').onclick = ev => { ev.preventDefault(); setZoom(zoom - 25); };
    $('#pIn').onclick = ev => { ev.preventDefault(); setZoom(zoom + 25); };
    $('#pFit').onclick = ev => { ev.preventDefault(); setZoom(100); };
    $('#pWide').onclick = ev => { ev.preventDefault(); setWide(!dlg.classList.contains('wide')); };
    dlg.addEventListener('close', () => setWide(false), { once: true });
    $('#pBack').onclick = () => { setWide(false); showInvoice(i); };
    openDlg();
    try {
      let buf = pdfCache.get(i.id);
      if (!buf) {
        const got = await OneDrive.get(i.file, 'buffer');
        if (!got) throw new Error(`${name} introuvable dans ${i.file.split('/').slice(0, -1).join('/')} (déplacé ou renommé dans OneDrive ?).`);
        pdfCache.set(i.id, buf = got.data);
      }
      if (!$('#pPages')) return;                     // dialogue fermé ou changé entre-temps
      $('#pDl').disabled = false;
      $('#pDl').onclick = () => download(new Blob([buf], { type: 'application/pdf' }), name);
      const r = await Fournisseurs.render(buf.slice(0), $('#pPages'));
      setZoom(zoom);                                 // zoom choisi pendant le chargement
      if ($('#pMsg')) $('#pMsg').textContent = `${name} · ${r.pages} page(s)${r.shown < r.pages ? ` (${r.shown} affichées : télécharge pour le reste)` : ''}`;
    } catch (e) { if ($('#pMsg')) $('#pMsg').innerHTML = `<span class="pill warn">${esc(e.message)}</span>`; }
  }

  function showInvoice(i) {
    const ls = alive(A.lines).filter(l => l.invoiceId === i.id);
    // Marqués « facturé » à la main (« Déjà refacturée », « Marquer comme facturé ») : annulable ; une facture de l'onglet Facturer s'annule là-bas.
    const manual = ls.filter(l => l.status === 'billed' && !l.billId);
    const dlg = $('#dlg'), f = $('#dlgForm');
    f.innerHTML = `<h2>${esc(i.supplier)} ${esc(i.number)}</h2>
      <p class="small muted">Du ${frDate(i.date)} · total HTVA ${i.totalExcl != null ? eur(i.totalExcl) : '?'} · importée le ${frDate(isoDate(new Date(i.importedAt)))}
        ${i.skipped ? `· ${i.skipped} article(s) à 0 € ignoré(s)` : ''}<br><code>${esc(i.file)}</code></p>
      ${i.file ? '<button class="btn" id="iPdf">📄 Voir la facture</button>' : ''}
      ${toCheck(i) ? `<p class="small"><span class="pill warn">⚠ ${esc(i.warn)}</span></p>
          <p class="small muted" style="margin-top:0">Compare avec le PDF. Si un article manque, supprime l'import et signale-le pour ajuster le lecteur ;
            si le total est bon (remise, frais…), confirme-le.</p>
          <button class="btn chip" id="iWarnOk">✓ Total vérifié, c'est correct</button>`
        : i.warn ? `<p class="small muted">Total vérifié${i.warnOk ? ` le ${frDate(i.warnOk)}` : ''} (${esc(i.warn)})</p>`
        : `<p class="small"><span class="pill ok">✓ total des articles = total de la facture</span></p>`}
      ${i.notes?.length ? `<p class="small muted">Non importé comme achat : ${esc(i.notes.join(' ; '))}.</p>` : ''}
      <ul class="list">${ls.map(lineLi).join('')}</ul>
      ${i.interventions?.length ? `<p class="small" style="margin:12px 0 0"><strong>Interventions (pointage)</strong></p>
        <ul class="list" id="iWorks">${i.interventions.map(x => workLi(i, x)).join('')}</ul>` : ''}
      <div class="actions"><button class="btn danger" id="iDel" style="margin-right:auto">Supprimer l'import</button>
        ${ls.some(l => ['todo', 'ready'].includes(lineState(l))) ? '<button class="btn" id="iDone">Déjà refacturée</button>' : ''}
        ${manual.length ? `<button class="btn" id="iUndo">↩ Remettre à facturer (${manual.length})</button>` : ''}
        <button class="btn" value="cancel">Fermer</button></div>`;
    bindLines(f);
    if ($('#iPdf')) $('#iPdf').onclick = () => showPdf(i);
    bindWorks(f);
    if ($('#iUndo')) $('#iUndo').onclick = async ev => {
      ev.preventDefault();
      if (!confirm(`Remettre ${manual.length} article(s) de ${i.supplier} ${i.number} à facturer ?\n\nIls avaient été marqués comme facturés (${[...new Set(manual.map(l => l.billRef || 'sans n°'))].join(', ')}) : ils reviennent dans « Prêt à refacturer » ou « À compléter ».`)) return;
      manual.forEach(l => { Object.assign(touch(l), { status: 'open' }); delete l.billedAt; delete l.billRef; delete l.billedUnit; });
      dlg.close(); await changed(); toast('Articles remis à facturer');
    };
    if ($('#iWarnOk')) $('#iWarnOk').onclick = async ev => { ev.preventDefault(); touch(i).warnOk = today(); dlg.close(); await changed(); toast('Total marqué comme vérifié'); };
    // Facture déjà refacturée hors de l'app (ex. avant la mise en service du module) : tout passe en « facturé ».
    if ($('#iDone')) $('#iDone').onclick = async () => {
      const todo = ls.filter(l => ['todo', 'ready'].includes(lineState(l)));
      if (!confirm(`Marquer les ${todo.length} article(s) de ${i.supplier} ${i.number} comme déjà refacturés au client ?\n\nIls ne seront plus proposés à la facturation (ils vont dans l'archive).`)) return;
      todo.forEach(l => Object.assign(touch(l), { status: 'billed', billedAt: today(), billRef: 'déjà refacturé' }));
      dlg.close(); await changed(); toast('Facture marquée comme déjà refacturée');
    };
    $('#iDel').onclick = async () => {
      if (ls.some(l => l.status === 'billed')) return alert('Des articles de cette facture sont déjà marqués comme facturés : import non supprimé.');
      if (!confirm(`Supprimer l'import de ${i.supplier} ${i.number} et ses ${ls.length} article(s) ?\n\nLe PDF reste dans Traités : remets-le dans Entrant pour le réimporter.`)) return;
      touch(i).deleted = true;
      ls.forEach(l => { touch(l).deleted = true; });
      dlg.close(); await changed();
    };
    openDlg();
  }

  // Onglet « Articles » : articles Billit (création sans passer par un import), correspondances, marges des fournisseurs, contacts → clients.
  // Les fiches s'ouvrent en fenêtre ; « Retour » / enregistrement reviennent sur la page.
  let artFind = '', artCat = '';
  const backToArticles = () => { if ($('#dlg').open) $('#dlg').close(); render(); };
  function articles() {
    if (!A) { $('#view').innerHTML = '<div class="empty">Chargement…</div>'; load().then(() => tab === 'articles' && render()); return; }
    if (Billit.ready() && !productsErr && (!products || Date.now() - productsAt > 60000)) loadProducts().then(() => tab === 'articles' && !$('#dlg').open && render());
    if (!Billit.ready() && !productsErr) productsErr = 'Clé API Billit absente (Réglages).';
    const cats = categories(), catOf = p => p.GroupID ? cats.find(g => g.id === p.GroupID)?.name || p.GroupName || '' : '';
    const prods = (products || []).slice().sort((a, b) => String(a.Reference).localeCompare(String(b.Reference)));
    const maps = alive(A.mapping).sort((a, b) => (a.supplier + a.label).localeCompare(b.supplier + b.label));
    const sups = alive(A.suppliers).sort((a, b) => a.name.localeCompare(b.name));
    const hints = alive(A.hints).sort((a, b) => a.label.localeCompare(b.label));
    // Ordre voulu par Fred (26/10) : correspondances d'abord, puis articles Billit (liste compacte, une ligne par article, zone qui défile).
    $('#view').innerHTML = `
      <div class="card"><h2>Correspondance des articles fournisseurs (${maps.length})</h2>
        <p class="small muted" style="margin-top:0">Article d'une facture fournisseur → article Billit. Touche une ligne pour la modifier (nom, article, marge).</p>
        ${maps.length ? `<ul class="list">${maps.map(m => { const p = product(m.billitRef); return `<li class="tap" data-map="${m.id}"><div class="grow" style="flex:1;min-width:0">${supPill(m.supplier)} <strong>${esc(m.label)}</strong>
            <div class="small muted">${[m.ignore && 'non refacturé', isSet(m.margin) && `marge article ${pct(m.margin)}`, p && `Billit ${eur(productPrice(p))}`].filter(Boolean).join(' · ')}</div></div>
            <div class="dur">${esc(m.billitRef || '—')}${m.billitRef && products && !p ? '<div><span class="pill warn">inconnu</span></div>' : ''}</div></li>`; }).join('')}</ul>`
          : '<div class="empty">Vide : elle se remplit quand tu complètes les articles importés (onglet Achats).</div>'}
      </div>
      <div class="card"><div class="row"><h2 class="grow" style="margin:0;border:0;padding:0">Articles Billit (${products ? products.length : '?'})</h2>
          <button class="btn primary" id="pNew" ${Billit.ready() && products ? '' : 'disabled'}>+ Nouvel article</button></div>
        <p class="small muted" style="margin:6px 0 10px">Billit ${Billit.mode() === 'live' ? 'réel' : 'test (sandbox)'}. Touche un article pour le modifier (nom, prix, TVA, catégorie).</p>
        ${productsErr ? `<p class="small"><span class="pill warn">${esc(productsErr)}</span></p>` : ''}
        ${products ? `<div class="grid2"><input type="search" id="artFind" placeholder="Rechercher : référence, nom…" value="${esc(artFind)}">
          <select id="artCat"><option value="">Toutes les catégories</option>
            ${cats.map(g => `<option value="${g.id}" ${String(g.id) === artCat ? 'selected' : ''}>${esc(g.name)} (${g.refs.length})</option>`).join('')}
            <option value="none" ${artCat === 'none' ? 'selected' : ''}>Sans catégorie (${prods.filter(p => !p.GroupID).length})</option></select></div>
          <div class="artbox"><ul class="list compact" id="artList">${prods.map(p => `<li class="tap" data-prod="${esc(p.ProductID)}" data-cat="${esc(p.GroupID || 'none')}"
              data-find="${esc(fold([p.Reference, p.Description, catOf(p)].join(' ')))}"
              title="${esc([catOf(p) || 'sans catégorie', p.VAT != null ? `TVA ${p.VAT} %` : ''].filter(Boolean).join(' · '))}">
            <strong class="ref">${esc(p.Reference || '(sans réf.)')}</strong>
            <span class="name">${esc(p.Description || '')}${catOf(p) ? '' : ' <span class="pill warn">sans catégorie</span>'}</span>
            <span class="dur">${eur(productPrice(p))}</span></li>`).join('')}</ul></div>
          <p class="small muted" id="artMore" style="margin:6px 0 0"></p>` : !productsErr ? '<div class="empty">Chargement des articles Billit…</div>' : ''}
      </div>
      <div class="card"><h2>Marges des fournisseurs</h2>
        <p class="small muted" style="margin-top:0">Utilisée quand l'article n'a pas sa propre marge. Prix de vente = achat × (1 + marge).</p>
        <ul class="list">${sups.map(s => `<li style="align-items:center"><span class="grow" style="flex:1">${supPill(s.name)}</span>
          <input type="number" step="0.1" data-smargin="${s.id}" value="${esc(s.margin)}" style="width:90px"> %</li>`).join('')}</ul>
        <div class="row"><button class="btn chip" id="sAdd">+ Fournisseur</button></div>
      </div>
      <div class="card"><h2>Contacts → clients (${hints.length})</h2>
        <p class="small muted" style="margin-top:0">Pour les fournisseurs sans domaine (MGK) : appris en complétant un article (« Retenir … pour ce client »).</p>
        ${hints.length ? `<ul class="list">${hints.map(h => `<li style="align-items:center"><div class="grow" style="flex:1;min-width:0">
            <strong>${esc(h.label)}</strong> → ${esc(client(h.clientId).name)}<div class="small muted">${esc(h.supplier)}</div></div>
            <button class="btn danger chip" data-delhint="${h.id}" title="Oublier">🗑</button></li>`).join('')}</ul>` : '<div class="empty">Aucun contact appris.</div>'}
      </div>`;

    // Recherche / catégorie sans redessiner la page.
    const filterArt = () => {
      const words = fold(artFind).split(/\s+/).filter(Boolean), all = $$('#artList > li');
      const hits = all.filter(li => (!artCat || li.dataset.cat === artCat) && words.every(w => li.dataset.find.includes(w)));
      all.forEach(li => li.classList.toggle('hidden', !hits.includes(li)));
      if ($('#artMore')) $('#artMore').textContent = hits.length === all.length ? `${all.length} article(s).` : `${hits.length} article(s) sur ${all.length}.`;
    };
    if ($('#artFind')) {
      $('#artFind').oninput = ev => { artFind = ev.target.value; filterArt(); };
      $('#artCat').onchange = ev => { artCat = ev.target.value; filterArt(); };
      filterArt();
    }
    $('#pNew').onclick = () => createArticle(artCat && artCat !== 'none' ? { GroupID: artCat } : {}, backToArticles);
    $$('[data-prod]').forEach(li => li.onclick = () => editProduct(products.find(p => String(p.ProductID) === li.dataset.prod)));
    $$('[data-map]').forEach(li => li.onclick = () => editMap(A.mapping.find(m => m.id === li.dataset.map)));
    $$('[data-smargin]').forEach(inp => inp.onchange = async () => {
      const s = A.suppliers.find(x => x.id === inp.dataset.smargin);
      Object.assign(touch(s), { margin: inp.value === '' ? '' : +inp.value });
      await save(); scheduleSync(); toast(`Marge ${s.name} : ${inp.value === '' ? 'aucune' : pct(inp.value)}`);
    });
    $('#sAdd').onclick = async () => {
      const name = (prompt('Nom du fournisseur (tel qu\'il apparaît dans l\'app) :', '') || '').trim();
      if (!name) return;
      if (supplierOf(name)) return alert('Ce fournisseur existe déjà.');
      const m = prompt(`Marge par défaut de ${name} (%) :`, '100');
      A.suppliers.push(touch({ id: uid('sup'), name, margin: m === null || m.trim() === '' ? '' : +m.replace(',', '.') }));
      await changed();
    };
    $$('[data-delhint]').forEach(b => b.onclick = async () => {
      const h = A.hints.find(x => x.id === b.dataset.delhint);
      if (!confirm(`Oublier « ${h.label} » → ${client(h.clientId).name} ? (les articles déjà attribués ne changent pas)`)) return;
      touch(h).deleted = true; await changed();
    });
  }

  function editMap(m) {
    const f = $('#dlgForm'), s = supplierOf(m.supplier);
    f.innerHTML = `<h2>${esc(m.label)}</h2><p class="small muted">${esc(m.supplier)} · clé « ${esc(m.key)} »</p>
      <label>Nom du produit (affiché dans l'app et les fichiers Excel)</label><input id="mLabel" value="${esc(m.label)}">
      <label>Article Billit (référence)</label><input id="mRef" value="${esc(m.billitRef)}" list="billitRefs" autocapitalize="off">
      <datalist id="billitRefs">${productOptions()}</datalist>
      <label>Marge de l'article en % (vide = marge fournisseur : ${s && isSet(s.margin) ? pct(s.margin) : 'non définie'})</label>
      <input type="number" step="0.1" id="mMargin" value="${esc(m.margin ?? '')}">
      <label class="check"><input type="checkbox" id="mIgnore" ${m.ignore ? 'checked' : ''}> Ne jamais refacturer ce type d'article</label>
      <div class="actions"><button class="btn danger" id="mDel" style="margin-right:auto">Supprimer</button>
        <button class="btn" id="mBack">Retour</button><button class="btn primary" id="mSave">Enregistrer</button></div>`;
    $('#mBack').onclick = backToArticles;
    $('#mSave').onclick = async ev => {
      ev.preventDefault();
      const margin = $('#mMargin').value.trim(), label = $('#mLabel').value.trim().replace(/\s+/g, ' ');
      if (!label) return alert('Nom requis.');
      // Le nom est repris par les lignes de ce type (la clé de correspondance, elle, ne change pas).
      if (label !== m.label) alive(A.lines).forEach(l => { if (l.mapId === m.id && l.label === m.label) { l.label = label; touch(l); } });
      Object.assign(touch(m), { label, billitRef: $('#mRef').value.trim(), margin: margin === '' ? '' : +margin, ignore: $('#mIgnore').checked });
      await save(); scheduleSync(); backToArticles();
    };
    $('#mDel').onclick = async () => {
      if (!confirm('Supprimer cette correspondance ? Les articles non facturés de ce type redeviendront « à compléter ».')) return;
      touch(m).deleted = true;
      alive(A.lines).forEach(l => { if (l.mapId === m.id && l.status !== 'billed') { l.mapId = null; touch(l); } });
      await save(); scheduleSync(); backToArticles();
    };
  }

  // Modification d'un article Billit existant (nom, prix, TVA) : renvoyé en entier avec son ProductID.
  function editProduct(p) {
    if (!p) return;
    const f = $('#dlgForm');
    f.innerHTML = `<h2>Article Billit ${esc(p.Reference || '')}</h2>
      <p class="small muted">Modifié dans Billit en mode ${Billit.mode() === 'live' ? 'réel' : 'test (sandbox)'}.</p>
      <label>Catégorie Billit</label><select id="pGroup"><option value="">${p.GroupID ? '— choisir —' : '— sans catégorie —'}</option>
        ${categories().map(g => `<option value="${g.id}" ${g.id === p.GroupID ? 'selected' : ''}>${esc(g.name)}</option>`).join('')}</select>
      <label>Référence</label><input id="pRef" value="${esc(p.Reference)}" autocapitalize="off">
      <label>Nom (description)</label><input id="pDesc" value="${esc(p.Description)}">
      <div class="grid2"><div><label>Prix de vente HTVA</label><input type="number" step="0.01" id="pPrice" value="${esc(productPrice(p))}"></div>
        <div><label>TVA %</label><input type="number" step="1" id="pVat" value="${esc(p.VAT ?? 21)}"></div></div>
      <div class="actions"><button class="btn" id="pBack">Retour</button><button class="btn primary" id="pSave">Enregistrer dans Billit</button></div>`;
    $('#pBack').onclick = backToArticles;
    $('#pSave').onclick = async ev => {
      ev.preventDefault();
      const item = { ...productBody(p), Reference: $('#pRef').value.trim(), Description: $('#pDesc').value.trim(), AmountExcl: +$('#pPrice').value, VAT: +$('#pVat').value,
        ...($('#pGroup').value ? { GroupID: +$('#pGroup').value } : {}) };
      if (p.GroupID && !$('#pGroup').value) return alert('Choisis une catégorie (un article ne peut pas perdre sa catégorie).');
      if (!item.Reference || !item.Description || $('#pPrice').value === '') return alert('Référence, nom et prix requis.');
      if (item.Reference !== p.Reference && products.some(x => x !== p && fold(x.Reference) === fold(item.Reference))) return alert(`La référence ${item.Reference} existe déjà.`);
      if (item.Reference !== p.Reference && alive(A.mapping).some(m => fold(m.billitRef) === fold(p.Reference))
        && !confirm(`Des correspondances utilisent ${p.Reference}. Elles seront mises à jour vers ${item.Reference}. Continuer ?`)) return;
      const btn = $('#pSave'); btn.disabled = true;
      try {
        await Billit.saveProduct(item);
        if (item.Reference !== p.Reference) {
          alive(A.mapping).forEach(m => { if (fold(m.billitRef) === fold(p.Reference)) { m.billitRef = item.Reference; touch(m); } });
          await save(); scheduleSync();
        }
        await loadProducts(0); toast('Article mis à jour dans Billit'); backToArticles();
      } catch (e) { alert(e.message); if (btn.isConnected) btn.disabled = false; }
    };
    openDlg();
  }

  // --- lien avec l'onglet Facturer (app.js) : les achats prêts d'un client vont sur sa facture, avec ses interventions ---
  const forBilling = l => ({ id: l.id, label: l.label, supplier: l.supplier, domain: whereOf(l), from: l.from, to: l.to,
    qty: l.qty, unit: lineUnit(l), total: saleTotal(l), ref: mapOf(l)?.billitRef || '' });
  // Données pas encore chargées (tout début) : liste vide, puis nouvel affichage une fois chargées.
  const ensure = () => { if (A) return true; load().then(() => render()); return false; };
  const readyClients = () => !ensure() ? [] : [...new Set(alive(A.lines).filter(l => lineState(l) === 'ready').map(l => l.clientId))];
  const readyFor = clientId => !ensure() ? []
    : alive(A.lines).filter(l => l.clientId === clientId && lineState(l) === 'ready').sort(byClient).map(forBilling);
  const ofBill = billId => !ensure() ? [] : alive(A.lines).filter(l => l.billId === billId).map(forBilling);
  async function attach(ids, b) {
    await load();
    for (const id of ids) {
      const l = A.lines.find(x => x.id === id);
      if (l) Object.assign(touch(l), { status: 'billed', billId: b.id, billedAt: b.date, billRef: '', billedUnit: saleUnit(l) });
    }
    await save();
  }
  async function detach(billId) {
    await load();
    alive(A.lines).forEach(l => { if (l.billId === billId) Object.assign(touch(l), { status: 'open', billId: null, billedAt: null, billedUnit: null }); });
    await save();
  }
  // Lignes de la facture Billit : prix figé à la facturation, article Billit trouvé par sa référence.
  function orderLines(billId, prods) {
    const missing = [];
    const lines = ofBill(billId).map(x => {
      const p = prods.find(p => fold(p.Reference) === fold(x.ref));
      if (!p) missing.push(x.ref);
      const what = [p?.Description || x.label, x.domain, x.from && `du ${frDate(x.from)} au ${frDate(x.to)}`].filter(Boolean).join(' – ');
      return { Quantity: x.qty, Reference: x.ref, Description: what, UnitPriceExcl: x.unit, VATPercentage: p?.VAT ?? 21,
        ...(p?.ProductID ? { ProductID: p.ProductID } : {}) };
    });
    return { lines, missing };
  }

  // App ouverte mais pas touchée : la vérification horaire se fait quand même (seulement si l'app est visible et déverrouillée).
  setInterval(() => {
    if (document.visibilityState === 'visible' && !locked && OneDrive.signedIn() && Date.now() - lastCheck > CHECK_EVERY) sync();
  }, 5 * 60000);

  // Déconnexion : on oublie la copie locale (les données restent dans OneDrive).
  async function reset() { A = null; products = null; productsErr = ''; lastCheck = 0; msg = ''; await idb('readwrite', s => s.delete('data')); }

  return { view, articles, sync, reset, todayBanner, readyClients, readyFor, ofBill, attach, detach, orderLines };
})();
