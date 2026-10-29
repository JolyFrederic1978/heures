// Copie Billit réel → sandbox, dans la version de développement uniquement (Réglages → Billit).
// Copie : catégories (par correspondance de nom, l'API ne sait pas les créer), clients, articles, et les factures de vente et d'achat
// des derniers mois (avec les fournisseurs de ces factures). Le Billit réel n'est que LU (Billit.live : requêtes GET uniquement).
// Sécurité : factures de vente en Projet ; clients et fournisseurs copiés sans e-mail ni identifiant Peppol, pour qu'aucun envoi
// fait depuis le sandbox n'atteigne un vrai client. Relancer la copie ne crée pas de doublon (clé « copie-reel:<OrderID> » dans Reference).
const CopieBillit = (() => {
  const norm = s => String(s || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^a-z0-9]+/g, ' ').trim();
  const vat = s => String(s || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
  const day = s => /^\d{4}-\d{2}-\d{2}/.exec(String(s || ''))?.[0] || '';
  const tag = id => `copie-reel:${id}`;
  const FILE_OK = /\.(pdf|xlsx?|csv)$/i;      // extensions acceptées par Billit pour les pièces jointes
  const PARTY_FIELDS = ['Name', 'CommercialName', 'Street', 'StreetNumber', 'Box', 'Zipcode', 'City', 'CountryCode', 'VATNumber', 'Language',
    'VATLiable', 'IBAN', 'BIC', 'Phone', 'Mobile', 'Website', 'Contact', 'ContactFirstName', 'ContactLastName', 'DefaultExpiryOffset'];
  const ADDRESS_FIELDS = ['AddressType', 'Tav', 'Name', 'Street', 'StreetNumber', 'Box', 'Zipcode', 'City', 'CountryCode'];
  const LINE_FIELDS = ['Quantity', 'UnitPriceExcl', 'Description', 'DescriptionExtended', 'VATPercentage', 'Reference', 'Unit', 'ReductionPercentage'];
  const pick = (o, fields) => Object.fromEntries(fields.filter(k => o?.[k] != null && o[k] !== '').map(k => [k, o[k]]));
  const counterOf = o => o.OrderDirection === 'Cost'
    ? o.SupplierID ?? o.Supplier?.PartyID ?? o.CounterParty?.PartyID
    : o.CustomerID ?? o.Customer?.PartyID ?? o.CounterParty?.PartyID;

  function guard() {
    if (ENV !== 'dev' || Billit.mode() !== 'test') throw new Error('Copie possible seulement dans la version de développement, vers le sandbox.');
    if (!Billit.ready()) throw new Error('Clé API et PartyID du sandbox (test) manquants dans Réglages.');
    if (!Billit.live.ready()) throw new Error('Clé API du Billit réel manquante (champ « Clé API (réel) »).');
  }

  // 1. Lecture des deux comptes et préparation (rien n'est écrit, sauf le PartyID réel s'il manquait).
  async function prepare(months, log) {
    guard();
    const lk = state.settings.billit.live;
    if (!lk.partyId) {
      const co = ((await Billit.live.account())?.Companies || [])[0];
      if (!co?.PartyID) throw new Error('PartyID du Billit réel introuvable.');
      Billit.live.setPartyId(co.PartyID);
      state.settingsAt = Date.now(); await DB.save(state);
    }
    const since = new Date(); since.setMonth(since.getMonth() - months);
    const from = isoDate(since);
    const invoices = Billit.odata(`OrderType eq 'Invoice' and OrderDate ge DateTime'${from}'`);
    log('Lecture du Billit réel…');
    const [lProducts, lParties, lOrders] = await Promise.all([
      Billit.live.list('/v1/products'), Billit.live.list('/v1/parties'), Billit.live.list('/v1/orders' + invoices)]);
    log('Lecture du sandbox…');
    const [tProducts, tParties, tOrders] = await Promise.all([
      Billit.list('/v1/products'), Billit.list('/v1/parties'), Billit.list('/v1/orders' + Billit.odata("OrderType eq 'Invoice'"))]);

    // Catégories : par nom (celles du sandbox sont lues sur ses articles).
    const tGroups = new Map(tProducts.filter(p => p.GroupID && p.GroupName).map(p => [norm(p.GroupName), p.GroupID]));
    const lGroups = new Map(lProducts.filter(p => p.GroupID && p.GroupName).map(p => [p.GroupID, p.GroupName]));
    const groupMap = new Map([...lGroups].filter(([, n]) => tGroups.has(norm(n))).map(([id, n]) => [id, tGroups.get(norm(n))]));
    const missingGroups = [...new Set([...lGroups.values()].filter(n => !tGroups.has(norm(n))))].sort();

    // Clients : tous ; fournisseurs : ceux des factures d'achat copiées. Déjà dans le sandbox = même n° de TVA ou même nom, même type.
    const orders = lOrders.filter(o => !tOrders.some(t => t.Reference === tag(o.OrderID)))
      .sort((a, b) => day(a.OrderDate).localeCompare(day(b.OrderDate)));
    const supplierIds = new Set(orders.filter(o => o.OrderDirection === 'Cost').map(o => String(counterOf(o))));
    const parties = lParties.filter(p => p.PartyType === 'Customer' || supplierIds.has(String(p.PartyID)));
    const twin = p => tParties.find(x => x.PartyType === p.PartyType && ((vat(p.VATNumber) && vat(x.VATNumber) === vat(p.VATNumber))
      || norm(Billit.partyName(x)) === norm(Billit.partyName(p))));
    const partyMap = new Map(), newParties = [];
    for (const p of parties) { const t = twin(p); t ? partyMap.set(String(p.PartyID), t.PartyID) : newParties.push(p); }

    // Articles : même référence ; un article du sandbox sans catégorie la reçoit si elle existe maintenant.
    const tByRef = new Map(tProducts.filter(p => p.Reference).map(p => [norm(p.Reference), p]));
    const productMap = new Map(), newProducts = [], regroup = [];
    for (const p of lProducts) {
      const t = tByRef.get(norm(p.Reference));
      if (!t) newProducts.push(p);
      else { productMap.set(p.ProductID, t.ProductID); if (!t.GroupID && groupMap.has(p.GroupID)) regroup.push([p, t]); }
    }
    return { months, from, groupMap, missingGroups, partyMap, newParties, productMap, newProducts, regroup, orders,
      already: lOrders.length - orders.length, sales: orders.filter(o => o.OrderDirection !== 'Cost').length, costs: orders.filter(o => o.OrderDirection === 'Cost').length };
  }

  const summary = p => [
    `Billit réel → sandbox, factures depuis le ${p.from.split('-').reverse().join('/')} :`,
    `- clients / fournisseurs à créer : ${p.newParties.length} (${p.partyMap.size} déjà dans le sandbox)`,
    `- articles à créer : ${p.newProducts.length} (${p.productMap.size} déjà dans le sandbox${p.regroup.length ? `, ${p.regroup.length} à ranger dans leur catégorie` : ''})`,
    `- factures à copier : ${p.sales} vente(s) en Projet + ${p.costs} achat(s)${p.already ? ` (${p.already} déjà copiée(s))` : ''}`,
    p.missingGroups.length ? `- catégories absentes du sandbox (l'API ne peut pas les créer) : ${p.missingGroups.join(', ')}.\n  Les articles concernés sont créés sans catégorie ; crée ces catégories dans le sandbox puis relance la copie pour les y ranger.` : '- catégories : toutes présentes dans le sandbox'
  ].join('\n');

  const fileOf = async f => { const x = await Billit.live.get('/v1/files/' + encodeURIComponent(f.FileID)); return { FileName: x.FileName || f.FileName, FileContent: x.FileContent }; };

  // 2. Écriture dans le sandbox. Chaque élément en erreur est noté et la copie continue.
  async function run(p, log) {
    guard();
    let errors = 0, parties = 0, products = 0;
    const fail = (what, e) => { errors++; log(`  ✗ ${what} : ${e.message}`); };
    const productBody = (lp, extra = {}) => ({ ...pick(lp, ['Reference', 'Description', 'AmountExcl', 'VAT', 'Unit', 'BuyExcl']),
      ...(p.groupMap.has(lp.GroupID) ? { GroupID: p.groupMap.get(lp.GroupID) } : {}), ...extra });

    log(`Clients / fournisseurs (${p.newParties.length})…`);
    for (const lp of p.newParties) {
      try {
        const id = await Billit.saveParty({ ...pick(lp, PARTY_FIELDS), PartyType: lp.PartyType,
          Addresses: (lp.Addresses || []).map(a => pick(a, ADDRESS_FIELDS)) });
        p.partyMap.set(String(lp.PartyID), id); parties++;
      } catch (e) { fail(Billit.partyName(lp), e); }
    }
    log(`Articles (${p.newProducts.length})…`);
    for (const lp of p.newProducts) {
      try { p.productMap.set(lp.ProductID, await Billit.saveProduct(productBody(lp))); products++; } catch (e) { fail(lp.Reference, e); }
    }
    for (const [lp, t] of p.regroup) {
      try { await Billit.saveProduct(productBody(lp, { ProductID: t.ProductID })); } catch (e) { fail(lp.Reference + ' (catégorie)', e); }
    }
    log(`Factures (${p.orders.length})…`);
    let done = 0;
    for (const lo of p.orders) {
      const label = `${lo.OrderDirection === 'Cost' ? 'achat' : 'vente'} ${lo.OrderNumber || lo.OrderID} du ${day(lo.OrderDate)}`;
      try {
        const d = await Billit.live.get('/v1/orders/' + lo.OrderID);
        const cost = d.OrderDirection === 'Cost', cp = p.partyMap.get(String(counterOf(d)));
        if (!cp) throw new Error(`${cost ? 'fournisseur' : 'client'} absent du sandbox`);
        const order = { OrderType: 'Invoice', OrderDirection: d.OrderDirection, OrderDate: day(d.OrderDate), ExpiryDate: day(d.ExpiryDate) || day(d.OrderDate),
          OrderTitle: d.OrderTitle || '', Currency: d.Currency || 'EUR', Reference: tag(d.OrderID), [cost ? 'SupplierID' : 'CustomerID']: cp,
          OrderLines: (d.OrderLines || []).map(l => ({ ...pick(l, LINE_FIELDS), ...(p.productMap.has(l.ProductID) ? { ProductID: p.productMap.get(l.ProductID) } : {}) })) };
        if (cost) {
          order.OrderNumber = d.OrderNumber || '';      // n° du fournisseur : sert au module Achats
          if (d.OrderPDF?.FileID) order.OrderPDF = await fileOf(d.OrderPDF);
        } else order.OrderStatus = 'Draft';             // Projet : jamais numérotée ni envoyée
        const atts = (d.Attachments || []).filter(a => a.FileID && FILE_OK.test(a.FileName || '') && a.FileID !== d.OrderPDF?.FileID);
        if (atts.length) order.Attachments = await Promise.all(atts.map(fileOf));
        await Billit.createOrder(order);
        done++;
      } catch (e) { fail(label, e); }
    }
    log(`Terminé : ${parties} client(s)/fournisseur(s), ${products} article(s), ${done} facture(s) copiée(s)${errors ? `, ${errors} erreur(s) (voir ci-dessus)` : ', aucune erreur'}.`);
    return { done, errors };
  }

  return { prepare, summary, run };
})();
