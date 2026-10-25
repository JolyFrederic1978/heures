// Accès à l'API Billit (https://docs.billit.be) : clients, articles, factures.
// Authentification par en-têtes ApiKey + PartyID. Mode « test » = sandbox, « live » = comptabilité réelle.
const Billit = (() => {
  const API = { test: 'https://api.sandbox.billit.be', live: 'https://api.billit.be' };
  const WEB = { test: 'https://my.sandbox.billit.be', live: 'https://my.billit.eu' };
  let cfg = () => ({});

  const keys = () => { const c = cfg(); return c[c.mode] || {}; };
  const mode = () => cfg().mode || 'test';

  async function call(path, opts = {}) {
    const k = keys();
    if (!k.apiKey || (!k.partyId && !opts.noParty)) throw new Error(`Clé API / PartyID Billit (${mode() === 'live' ? 'réel' : 'test'}) manquants dans Réglages.`);
    const headers = { ApiKey: k.apiKey, Accept: 'application/json' };
    if (k.partyId) headers.PartyID = k.partyId;
    if (opts.body) headers['Content-Type'] = 'application/json';
    const r = await fetch(API[mode()] + path, { ...opts, headers });
    const text = await r.text();
    let data = null;
    try { data = text ? JSON.parse(text) : null; } catch (_) { data = text; }
    if (!r.ok) {
      const msg = typeof data === 'string' ? data
        : data?.Message || data?.message || data?.errors?.map(e => e.Description || e.Code).join(' ; ') || JSON.stringify(data);
      throw new Error(`Billit (${r.status}) : ${String(msg || r.statusText).slice(0, 400)}`);
    }
    return data;
  }
  // Les listes peuvent arriver sous forme de tableau ou d'objet { Items: [...] } selon l'endpoint.
  const items = d => Array.isArray(d) ? d : d?.Items || d?.items || d?.value || [];
  const odata = f => '?$filter=' + encodeURIComponent(f);

  const partyName = p => p.Name || p.CommercialName || p.DisplayName || '';

  return {
    configure(fn) { cfg = fn; },
    mode,
    ready: () => !!(keys().apiKey && keys().partyId),
    orderUrl: id => `${WEB[mode()]}/Order/Income/Invoice/${id}/EditOrder`,
    partyName,
    // Fonctionne avec la clé seule : sert aussi à retrouver le PartyID de la société.
    account: () => call('/v1/account/accountInformation', { noParty: true }),
    products: async () => items(await call('/v1/products')),
    async customers() {
      try { return items(await call('/v1/parties' + odata("PartyType eq 'Customer'"))); }
      catch (_) { return items(await call('/v1/parties')); }
    },
    // Factures de vente en projet (Draft) ou pas encore envoyées (ToSend) pour ce client.
    async openInvoices(customerId) {
      const all = items(await call('/v1/orders' + odata("OrderType eq 'Invoice' and OrderDirection eq 'Income' and (OrderStatus eq 'Draft' or OrderStatus eq 'ToSend')")));
      return all.filter(o => String(o.CounterParty?.PartyID ?? o.CustomerID ?? o.Customer?.PartyID) === String(customerId));
    },
    createOrder: order => call('/v1/orders', { method: 'POST', body: JSON.stringify(order) }),
    // Crée un article, ou le met à jour si ProductID est fourni (champs : Reference, Description, AmountExcl, VAT, BuyExcl…).
    // Corps = un seul objet (pas { Items: [...] } comme dans le guide Billit : l'article est alors créé vide). Réponse = ProductID.
    saveProduct: item => call('/v1/products', { method: 'POST', body: JSON.stringify(item) }),
    getOrder: id => call('/v1/orders/' + id),
    // Factures d'achat (OrderDirection Cost) modifiées depuis la date « AAAA-MM-JJ » ; toutes les pages (120 par page).
    async costInvoices(since) {
      let path = '/v1/orders' + odata(`OrderType eq 'Invoice' and OrderDirection eq 'Cost' and LastModified ge DateTime'${since}'`), all = [];
      for (let i = 0; path && i < 50; i++) {
        const d = await call(path);
        all.push(...items(d));
        const next = d?.NextPageLink;
        path = !next ? null : next.startsWith('http') ? next.replace(/^https?:\/\/[^/]+/, '') : next.startsWith('/') ? next : '/v1/' + next;
      }
      return all;
    },
    // Fichier (PDF de la facture, e-facture XML…) : { FileName, MimeType, buffer }.
    async file(id) {
      const f = await call('/v1/files/' + encodeURIComponent(id));
      return { FileName: f.FileName, MimeType: f.MimeType, buffer: Uint8Array.from(atob(f.FileContent || ''), c => c.charCodeAt(0)).buffer };
    }
  };
})();
