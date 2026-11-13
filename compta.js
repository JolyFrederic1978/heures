// © 2026 One Level UP – LevelDesk. Logiciel propriétaire, tous droits réservés. Copie, modification, redistribution et rétro-ingénierie interdites sans autorisation écrite (voir LICENCE.txt).
// Connecteur de facturation de la société (architecture.md, section 11) : des opérations LevelDesk, quel que soit le logiciel (Billit, Odoo…).
// Odoo : seulement avec le serveur (serveur/api/lib/connecteurs/odoo.php).
// Formats LevelDesk :
//   client  { id, name, tradeName, vat, email, address }
//   article { id, ref, name, price, vat, buy, unit, catId, catName }
//   facture { customerId, date, dueDate, title, lines: [{ articleId, ref, description, qty, price, vat }], files: [{ name, mime, content (base64) }],
//             note (détail des interventions, texte : mis dans la note du devis quand la société crée des devis) }
//   achat   { id, number, supplier, date, total, pdfId }
// Avec le serveur : POST ?r=compta (connecteur et clés de la société sur le serveur, serveur/api/lib/compta.php).
// Sans serveur (fichiers OneDrive, production actuelle) : Billit appelé depuis l'appareil (billit.js), réponses mises ici au même format.
const Compta = (() => {
  // cfg() : { srv: données sur le serveur ?, admin, mode (réglage synchronisé) }
  let cfg = () => ({}), conf = null, confAt = 0;
  const NAMES = { billit: 'Billit', odoo: 'Odoo' };
  const srv = () => !!cfg().srv;
  // En développement : toujours le mode test (sandbox), même si les réglages disent « Réel ».
  const mode = () => ENV === 'prod' ? cfg().mode || 'test' : 'test';
  const type = () => srv() ? conf?.modes?.[mode()]?.type || (conf ? '' : 'billit') : 'billit';

  async function op(name, args = {}, m = mode(), soc = '') { return (await Serveur.json('compta', { method: 'POST', body: { op: name, mode: m, args, ...(soc ? { soc } : {}) } })).data; }

  // Connecteur d'une AUTRE société que je gère (ex. Alescom gérée par One Level UP) : lecture seule (config, test, clients) — ses clients
  // sont liés à SON logiciel (Odoo d'Alescom), pas au mien. Avec le serveur seulement.
  const others = new Map();   // société → { conf, at }
  const other = {
    async config(soc, force) {
      if (!srv() || !cfg().admin) return null;
      const x = others.get(soc);
      if (!force && x && Date.now() - x.at < 300000) return x.conf;
      const c = await Serveur.json('compta', { method: 'POST', body: { op: 'config', soc } });
      others.set(soc, { conf: c, at: Date.now() });
      return c;
    },
    type: soc => others.get(soc)?.conf?.modes?.[mode()]?.type || '',
    ready: soc => { const x = others.get(soc)?.conf?.modes?.[mode()]; return !!(x?.type && x.key); },
    label: soc => NAMES[other.type(soc)] || 'Facturation',
    clients: soc => op('clients', {}, mode(), soc)
  };

  // ---- Billit direct (sans serveur) → format LevelDesk ----
  const num = v => v == null || v === '' ? null : +v;
  function address(p) {
    const all = p?.Addresses || [];
    const a = all.find(x => /deliver|livr/i.test(x.AddressType || '') && x.Street) || (p?.Street ? p : all.find(x => x.Street));
    if (!a) return '';
    const street = [a.Street, a.StreetNumber].filter(Boolean).join(' ') + (a.Box ? ' bte ' + a.Box : '');
    return [street, [a.Zipcode, a.City].filter(Boolean).join(' '), a.CountryCode && !/^BE$/i.test(a.CountryCode) ? a.CountryCode : '']
      .map(s => String(s).trim()).filter(Boolean).join(', ');
  }
  const toClient = p => ({ id: String(p.PartyID ?? ''), name: Billit.partyName(p), tradeName: p.CommercialName || '', vat: p.VATNumber || '', email: p.Email || '', address: address(p) });
  const toArticle = p => ({ id: String(p.ProductID ?? ''), ref: p.Reference || '', name: p.Description || '', price: num(p.AmountExcl ?? p.UnitPriceExcl),
    vat: num(p.VAT ?? p.VATPercentage), buy: num(p.BuyExcl), unit: p.Unit || '', catId: p.GroupID ? String(p.GroupID) : '', catName: p.GroupName || '' });
  const day = s => /^\d{4}-\d{2}-\d{2}/.exec(String(s || ''))?.[0] || '';
  const direct = {
    async test() {
      const a = await Billit.account();
      const cos = (a?.Companies || a?.companies || []).map(c => ({ id: String(c.PartyID), name: c.Name || c.CommercialName || '' }));
      return { companies: cos, partyId: cfg().partyId || '', company: cos.find(c => c.id === String(cfg().partyId))?.name || '' };
    },
    clients: async () => (await Billit.customers()).map(toClient),
    articles: async () => (await Billit.products()).map(toArticle),
    async saveArticle(a) {
      const keep = ['ProductID', 'Reference', 'Description', 'AmountExcl', 'VAT', 'Unit', 'GroupID', 'BuyExcl', 'StockQuantity', 'MinimumBilledQuantity', 'InternalInformation'];
      const old = a.id ? (await Billit.products()).find(p => String(p.ProductID) === String(a.id)) : null;
      if (a.id && !old) throw new Error('Article introuvable dans Billit.');
      const item = old ? Object.fromEntries(keep.filter(k => old[k] != null).map(k => [k, old[k]])) : {};
      for (const [k, b] of [['ref', 'Reference'], ['name', 'Description'], ['price', 'AmountExcl'], ['vat', 'VAT'], ['buy', 'BuyExcl']])
        if (a[k] != null && a[k] !== '') item[b] = ['ref', 'name'].includes(k) ? String(a[k]) : +a[k];
      if (a.catId) item.GroupID = +a.catId;
      const r = await Billit.saveProduct(item);
      return { id: String(typeof r === 'object' && r ? r.ProductID ?? r.id ?? '' : r) };
    },
    async draftInvoice(f) {
      const order = { OrderType: 'Invoice', OrderDirection: 'Income', OrderStatus: 'Draft', Currency: 'EUR', OrderDate: f.date, ExpiryDate: f.dueDate,
        CustomerID: +f.customerId || f.customerId, OrderTitle: f.title,
        OrderLines: f.lines.map(l => ({ Quantity: l.qty, Reference: l.ref, Description: l.description, UnitPriceExcl: l.price, VATPercentage: l.vat ?? 21,
          ...(l.articleId ? { ProductID: +l.articleId || l.articleId } : {}) })),
        ...(f.files?.length ? { Attachments: f.files.map(x => ({ FileName: x.name, FileContent: x.content, MimeType: x.mime })) } : {}) };
      const res = await Billit.createOrder(order);
      const id = String(typeof res === 'object' && res ? (res.OrderID ?? res.id ?? res.Id) : res);
      let o = null;
      try { o = await Billit.getOrder(id); } catch (_) { /* la facture est créée, le détail est facultatif */ }
      return { id, number: o?.OrderNumber || '', status: o?.OrderStatus || '', draft: /draft/i.test(o?.OrderStatus || ''), url: Billit.orderUrl(id) };
    },
    openInvoices: async id => (await Billit.openInvoices(id)).map(o => ({ id: String(o.OrderID), number: o.OrderNumber || '', date: day(o.OrderDate),
      total: num(o.TotalExcl), url: Billit.orderUrl(o.OrderID) })),
    purchases: async since => (await Billit.costInvoices(since)).map(o => ({ id: String(o.OrderID), number: o.OrderNumber || '',
      supplier: o.CounterParty?.DisplayName || o.CounterParty?.Name || '', date: day(o.OrderDate), total: num(o.TotalExcl), pdfId: o.OrderPDF?.FileID || '' })),
    purchase: async id => ({ id: String(id), files: ((await Billit.getOrder(id))?.Attachments || []).map(a => ({ id: a.FileID, name: a.FileName || '', mime: a.MimeType || '' })) }),
    async file(id) { const f = await Billit.file(id); return { name: f.FileName, mime: f.MimeType, buffer: f.buffer }; }
  };
  const b64buf = s => Uint8Array.from(atob(s || ''), c => c.charCodeAt(0)).buffer;

  return {
    configure(fn) { cfg = fn; },
    mode, type,
    label: () => NAMES[type()] || 'Facturation',
    // Prêt à être appelé : avec le serveur, un administrateur dont la société a un connecteur (inconnu tant que la config n'est pas lue : on essaie) ;
    // sans serveur, clé Billit et PartyID sur cet appareil.
    ready: () => srv() ? !!cfg().admin && (!conf || !!(conf.modes?.[mode()]?.type && conf.modes[mode()].key)) : Billit.ready(),
    webName: () => NAMES[type()] || '',
    // Document créé par l'envoi : facture en brouillon, ou devis en brouillon (Odoo, réglage du connecteur « doc » = quote).
    quote: () => srv() && conf?.modes?.[mode()]?.doc === 'quote',
    // Odoo : « Suivre l'inventaire » proposé à la création d'un article (case cochée par défaut selon le connecteur).
    stockOption: () => srv() && type() === 'odoo',
    stockDefault: () => srv() && !!conf?.modes?.[mode()]?.stock,
    // ---- configuration (serveur) ----
    config: () => conf,
    async loadConfig(force) {
      if (!srv() || !cfg().admin) return null;
      if (!force && conf && Date.now() - confAt < 300000) return conf;
      conf = await Serveur.json('compta', { method: 'POST', body: { op: 'config' } }); confAt = Date.now();
      return conf;
    },
    async saveConfig(body) { const r = await Serveur.json('compta', { method: 'POST', body: { op: 'config.save', ...body } }); conf = r; confAt = Date.now(); return r; },
    reset() { conf = null; confAt = 0; others.clear(); },
    other,
    // ---- opérations ----
    test: () => srv() ? op('test') : direct.test(),
    clients: () => srv() ? op('clients') : direct.clients(),
    articles: () => srv() ? op('articles') : direct.articles(),
    // Catégories { id, name } : listées par Odoo ; Billit seulement celles de ses articles.
    categories: () => srv() ? op('categories') : direct.articles().then(ps => [...new Map(ps.filter(p => p.catId).map(p => [p.catId, { id: p.catId, name: p.catName }])).values()]),
    saveArticle: a => srv() ? op('article.save', a) : direct.saveArticle(a),
    draftInvoice: f => srv() ? op('facture.brouillon', f) : direct.draftInvoice(f),
    openInvoices: customerId => srv() ? op('factures.ouvertes', { customerId: String(customerId) }) : direct.openInvoices(customerId),
    purchases: since => srv() ? op('achats', { since }) : direct.purchases(since),
    purchase: id => srv() ? op('achat', { id: String(id) }) : direct.purchase(id),
    // Odoo : valider une facture fournisseur et mettre ses articles en stock (lines : [{ articleId, qty, unit, label }]).
    validatePurchase: (id, lines) => op('achat.valider', { id: String(id), lines }),
    async file(id) {
      if (!srv()) return direct.file(id);
      const f = await op('fichier', { id: String(id) });
      return { name: f.name, mime: f.mime, buffer: b64buf(f.content) };
    },
    address
  };
})();
