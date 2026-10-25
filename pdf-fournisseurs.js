// Lecture des factures fournisseurs en PDF : texte extrait avec pdf.js (rangé en lignes par position),
// puis un lecteur par fournisseur. Un lecteur renvoie :
//   { supplier, number, date, totalExcl, lines: [{ key, label, detail, domain, qty, unit, total, from, to }] }
// label = nom du produit (sans le domaine du client) ; key = ce nom sans dates ni nombres :
// sert à la table de correspondance avec les articles Billit.
const Fournisseurs = (() => {
  // pdf.js 4.10.38 (build « legacy », compatible avec les navigateurs plus anciens), servi depuis lib/ :
  // pas de dépendance à un site extérieur. Version ≥ 4.2.67 : corrige la faille CVE-2024-4367 (code caché dans un PDF).
  const SRC = new URL('lib/pdf.min.mjs', location.href).href;
  const WORKER = new URL('lib/pdf.worker.min.mjs', location.href).href;
  const pad = n => String(n).padStart(2, '0');
  let pdfjsLib, loading;

  function load() {
    return loading ||= import(SRC).then(m => { pdfjsLib = m; pdfjsLib.GlobalWorkerOptions.workerSrc = WORKER; })
      .catch(e => { loading = null; throw new Error('Impossible de charger pdf.js : ' + e.message); });
  }
  // isEvalSupported: false : pdf.js ne compile jamais de code à partir du contenu d'un PDF (protection en plus de la mise à jour).
  const openPdf = buffer => pdfjsLib.getDocument({ data: new Uint8Array(buffer), isEvalSupported: false }).promise;

  // Lignes du PDF de haut en bas : { page, cells: [texte des colonnes], text }.
  async function readRows(buffer) {
    await load();
    const pdf = await openPdf(buffer);
    const rows = [];
    for (let p = 1; p <= pdf.numPages; p++) {
      const { items } = await (await pdf.getPage(p)).getTextContent();
      const lines = [];
      for (const it of items) {
        if (!it.str.trim()) continue;
        const x = it.transform[4], y = it.transform[5];
        let l = lines.find(l => Math.abs(l.y - y) < 2.5);
        if (!l) lines.push(l = { y, parts: [] });
        l.parts.push({ x, end: x + it.width, s: it.str });
      }
      lines.sort((a, b) => b.y - a.y);
      for (const l of lines) {
        l.parts.sort((a, b) => a.x - b.x);
        // Morceaux proches = même cellule ; un écart plus grand = colonne suivante.
        const cells = [];
        let prev = null;
        for (const q of l.parts) {
          const gap = prev ? q.x - prev.end : Infinity;
          if (gap < 6) cells[cells.length - 1] += (gap > 1 ? ' ' : '') + q.s; else cells.push(q.s);
          prev = q;
        }
        const c = cells.map(s => s.replace(/\s+/g, ' ').trim()).filter(Boolean);
        if (c.length) rows.push({ page: p, cells: c, text: c.join(' ') });
      }
    }
    return rows;
  }

  // --- outils communs aux lecteurs ---
  const AMOUNT = /^-?(?:\d{1,3}(?:[ .  ]\d{3})+|\d+)[.,]\d{2}\s*€?$/;
  const isAmount = s => AMOUNT.test(String(s).trim());
  function num(s) {
    s = String(s).replace(/[€\s  ]/g, '');
    s = /,\d{2}$/.test(s) ? s.replace(/\./g, '').replace(',', '.') : s.replace(/,/g, '');
    return parseFloat(s);
  }
  const DATE = /(\d{1,2})[/.-](\d{1,2})[/.-](\d{4})/;
  const MONTHS = ['janvier', 'fevrier', 'mars', 'avril', 'mai', 'juin', 'juillet', 'aout', 'septembre', 'octobre', 'novembre', 'decembre'];
  const plain = s => String(s || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '');
  // « 01/04/2026 » ou « 01 Avril 2026 » → 2026-04-01
  function isoDate(s) {
    let m = DATE.exec(s || '');
    if (m) return `${m[3]}-${pad(m[2])}-${pad(m[1])}`;
    m = /(\d{1,2})\s+([a-z]+)\s+(\d{4})/.exec(plain(s));
    const mo = m ? MONTHS.indexOf(m[2]) : -1;
    return mo >= 0 ? `${m[3]}-${pad(mo + 1)}-${pad(m[1])}` : null;
  }
  const period = s => {
    const m = /(\d{1,2}[/.-]\d{1,2}[/.-]\d{4})\s*(?:au|à|-|–|to)\s*(\d{1,2}[/.-]\d{1,2}[/.-]\d{4})/i.exec(s || '');
    return m ? { from: isoDate(m[1]), to: isoDate(m[2]) } : {};
  };
  // Domaines des clients ; les noms techniques du fournisseur (serveurs, OVH lui-même) ne comptent pas.
  const DOMAIN = /\b(?:[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?\.)+[a-z]{2,24}\b/gi;
  const NOT_CLIENT = /(^|\.)(ovh|ovhcloud|kimsufi|soyoustart|ovh-hosting)\.[a-z]+$|^(ns|vps|ip|dedibox|sd)-?\d|\.ip-[\d-]+\.[a-z]+$/i;
  const domains = s => (String(s || '').toLowerCase().match(DOMAIN) || [])
    .map(d => d.replace(/^www\./, '')).filter(d => !NOT_CLIENT.test(d) && !/^\d+(\.\d+)+$/.test(d) && !/\.(pdf|xlsx?|docx?)$/.test(d));
  const keyOf = s => plain(s).replace(/[^\s@]+@/g, ' ').replace(DOMAIN, ' ').replace(new RegExp(DATE.source, 'g'), ' ')
    .replace(/[^a-z]+/g, ' ').trim();
  const find = (rows, re) => { for (const r of rows) { const m = re.exec(r.text); if (m) return m; } return null; };

  // --- lecteurs ---
  // OVH (factures « Renouvellement » en français, vérifié sur 3 factures de 2026) :
  //   Rubrique Hébergement
  //   Abonnement | Référence | Quantité | Prix HT          ← en-tête du tableau (répété après un saut de page)
  //   Perso rental for 12 months                          ← nom du produit (1 ou 2 lignes)
  //   (01/01/2026-31/12/2026) | jmd.be | 1 | 39,48 € | 39,48 €
  //   Sans engagement
  //   SOUS TOTAL | 39,48 €
  // La période peut aussi venir sur la ligne qui suit les montants. Le nom peut commencer par « jmd.be - .be ».
  const ovh = {
    id: 'ovh', name: 'OVH',
    detect: rows => rows.some(r => /^OVH - .*ROUBAIX/i.test(r.text)) || rows.some(r => /ovh/i.test(r.text) && /Rubrique/i.test(r.text)),
    parse(rows, fileName) {
      const number = find(rows, /Référence de la facture\s*:?\s*([A-Z]{2}\d{5,})/i)?.[1]
        || find(rows, /Facture\s+([A-Z]{2}\d{5,})/)?.[1] || /([A-Z]{2}\d{5,})/.exec(fileName || '')?.[1] || '';
      const date = isoDate(find(rows, /Date d'émission\s*:?\s*(.+)$/i)?.[1]);
      const totalM = find(rows, /Total de la facture HT\s*(-?[\d .  ]+[.,]\d{2})/i);
      const FOOT = /Page \d+\/\d+|^OVH - |^SAS au capital|^SIREN|^FR\d+$/i;
      const PERIOD_ONLY = /^\(?\s*\d{1,2}\/\d{1,2}\/\d{4}\s*-\s*\d{1,2}\/\d{1,2}\/\d{4}\s*\)?$/;
      const lines = [];
      let inTable = false, section = '', pending = [];
      for (const r of rows) {
        const t = r.text;
        const sec = /^Rubrique\s+(.+)$/i.exec(t);
        if (sec) { section = sec[1]; inTable = false; pending = []; continue; }
        if (/^Abonnement\b.*Prix HT/i.test(t)) { inTable = true; pending = []; continue; }
        if (!inTable) continue;
        if (/^SOUS TOTAL/i.test(t) || FOOT.test(t)) { inTable = false; pending = []; continue; }
        if (/^(Prix|unitaire)$/i.test(t) || /engagement/i.test(t)) continue;
        const cells = [...r.cells], amounts = [];
        while (cells.length && isAmount(cells[cells.length - 1])) amounts.unshift(num(cells.pop()));
        if (amounts.length >= 2) {
          const qty = cells.length && /^\d+(?:[.,]\d+)?$/.test(cells[cells.length - 1]) ? parseFloat(cells.pop().replace(',', '.')) : 1;
          const ref = cells.length && !PERIOD_ONLY.test(cells[cells.length - 1]) ? cells.pop() : '';
          const extra = cells.filter(c => !PERIOD_ONLY.test(c));
          const full = [...pending, ...extra].join(' ').replace(/\s+/g, ' ').trim() || ref;
          // « jmd.be - .be demande de renouvellement - 12 mois » → « demande de renouvellement - 12 mois »
          const label = full.replace(/^[a-z0-9.-]+\.[a-z]{2,}\s*-\s*/i, '').replace(/^\.[a-z]{2,}\s+/i, '').trim() || full;
          lines.push({ label, detail: `${section} · ${full}`, key: keyOf(label),
            domain: domains(ref)[0] || domains(full)[0] || '', ref,
            qty, unit: amounts[amounts.length - 2], total: amounts[amounts.length - 1], ...period(t) });
          pending = [];
        } else if (PERIOD_ONLY.test(t)) {
          const last = lines[lines.length - 1];
          if (last && !last.from) Object.assign(last, period(t));
        } else pending.push(t);
      }
      return { supplier: 'OVH', number, date, totalExcl: totalM ? num(totalM[1]) : null, lines };
    }
  };

  // MGK Technologies (licences ESET, vérifié sur 3 factures 2025-2026) :
  //   No. | Référence Description | Informations complémentaires | Remise | Qté | Prix unitaire | Montant   ← en-tête
  //   1 | 106 | ESET NOD32 Antivirus | Hugo Bauwens | 1 | 33.05 | 33.05      ← « Informations » = contact chez le client
  //   Type d'achat: Renew / Période de validité: 1 année / Valide du 21.09.2025 jusqu'au / 20.09.2026
  //   Nombre de postes: 5 / Public ID: … / SKU Code: ENOD-NB5-R-1Y
  //   Date du document: | 30.09.2025 | Total | 33.05 €
  // Pas de domaine : le client est retrouvé par le contact (hint). Correspondance par code SKU (postes, durée et prix différents).
  const mgk = {
    id: 'mgk', name: 'MGK',
    detect: rows => rows.some(r => /MGK Technologies/i.test(r.text)),
    parse(rows, fileName) {
      const number = find(rows, /No\. de document\s*:?\s*([A-Z]{2,}\d{4}-\d+)/i)?.[1] || /(FAC\d{4}-\d+)/i.exec(fileName || '')?.[1] || '';
      const date = isoDate(find(rows, /Date du document\s*:?\s*(\d{1,2}[.\/-]\d{1,2}[.\/-]\d{4})/i)?.[1]);
      const totalM = find(rows, /\bTotal\s+(-?[\d .  ,]+[.,]\d{2})\s*€?\s*$/i);
      const items = [];
      let inTable = false;
      for (const r of rows) {
        const t = r.text;
        if (/^No\.\s.*Montant/i.test(t)) { inTable = true; continue; }
        if (!inTable) continue;
        if (/^Date du document/i.test(t)) { inTable = false; continue; }
        if (/^produit\b/i.test(t)) continue;
        const cells = [...r.cells], amounts = [];
        while (cells.length && isAmount(cells[cells.length - 1])) amounts.unshift(num(cells.pop()));
        if (/^\d+$/.test(cells[0] || '') && amounts.length >= 2) {
          const qty = cells.length > 1 && /^\d+(?:[.,]\d+)?$/.test(cells[cells.length - 1]) ? parseFloat(cells.pop().replace(',', '.')) : 1;
          if (cells.length > 3 && /^-?\d+(?:[.,]\d+)?\s*%?$/.test(cells[cells.length - 1])) cells.pop();   // remise
          cells.shift();                                                                                   // n° de ligne
          let ref = cells.shift() || '';
          if (!/^\d+$/.test(ref)) { const m = /^(\d+)\s+(.*)$/.exec(ref); if (m) { cells.unshift(m[2]); ref = m[1]; } }
          items.push({ ref, desc: cells.shift() || '', hint: cells.join(' '), qty,
            unit: amounts[amounts.length - 2], total: amounts[amounts.length - 1], info: [] });
        } else if (items.length) items[items.length - 1].info.push(t);
      }
      const lines = items.map(it => {
        const info = it.info.join(' ');
        const get = re => (re.exec(info) || [])[1]?.trim() || '';
        return mgkLine(it, date, { sku: get(/SKU Code\s*:\s*(\S+)/i), seats: get(/Nombre de postes\s*:\s*(\d+)/i),
          type: get(/Type d'achat\s*:\s*(\w+)/i), validity: get(/Période de validité\s*:\s*(\d+\s*[^\s,;]+)/i),
          from: isoDate(get(/Valide du\s+(\d{1,2}\.\d{1,2}\.\d{4})/i)), to: isoDate(get(/jusqu'au\s+(\d{1,2}\.\d{1,2}\.\d{4})/i)) });
      });
      return { supplier: 'MGK', number, date, totalExcl: totalM ? num(totalM[1]) : null, lines };
    }
  };
  // Ligne MGK (commune au PDF et à l'e-facture UBL). Période : « du A jusqu'au B », ou « jusqu'au B »,
  // ou seulement la durée (calculée depuis la date du document).
  function mgkLine(it, date, x) {
    let { from, to } = x;
    const years = +(/(\d+)\s*ann/i.exec(x.validity) || [])[1] || 0, months = +(/(\d+)\s*mois/i.exec(x.validity) || [])[1] || 0;
    const shift = (d, sign) => { const t = new Date(d + 'T00:00:00Z'); t.setUTCFullYear(t.getUTCFullYear() + sign * years); t.setUTCMonth(t.getUTCMonth() + sign * months); t.setUTCDate(t.getUTCDate() - sign); return t.toISOString().slice(0, 10); };
    if (!from && to && (years || months)) from = shift(to, -1);
    if (!from && !to && date && (years || months)) { from = date; to = shift(date, 1); }
    const label = [it.desc, x.seats && `${x.seats} poste${+x.seats > 1 ? 's' : ''}`, x.validity].filter(Boolean).join(' – ');
    return { label, detail: [x.type && `Type d'achat : ${x.type}`, x.sku && `SKU ${x.sku}`, it.ref && `réf. MGK ${it.ref}`].filter(Boolean).join(' · '),
      key: x.sku ? 'sku:' + x.sku.toLowerCase() : 'ref:' + it.ref + ':' + keyOf(it.desc),
      domain: '', hint: it.hint, ref: it.ref, qty: it.qty, unit: it.unit, total: it.total, from: from || null, to: to || null };
  }

  // lisam.cloud (vérifié sur 11 factures 2025-2026, reçues par Peppol dans Billit) :
  //   FACTURE 2026#34049 / Date : 2026-07-27
  //   Nom du produit | Quantité | Prix HT | Total | Tva                  ← en-tête
  //   remote.jmd.be | 12 Mois | 4,50 | 54,00 | 21                        ← domaine (ou « C1513 - JMD.BE ») et durée
  //   Certificat SSL de domaine avec gestion de votre CSR (2026-07-01 -> 2027-06-30)
  //   NOS COORDONNEES BANCAIRE | TOTAUX DE CETTE FACTURE … Montant HT | 54,00
  // Une ligne = 1 article au prix total (la durée fait partie du nom).
  const lisam = {
    id: 'lisam', name: 'lisam.cloud',
    detect: rows => rows.some(r => /LISAM\.CLOUD/i.test(r.text)),
    parse(rows) {
      const number = find(rows, /FACTURE\s+(\d{4}#\d+)/i)?.[1] || '';
      const date = find(rows, /^Date\s*:\s*(\d{4}-\d{2}-\d{2})/i)?.[1] || null;
      const totalM = find(rows, /Montant HT\s+(-?[\d .  ]+[.,]\d{2})/i);
      const lines = [];
      let inTable = false;
      for (const r of rows) {
        const t = r.text;
        if (/^Nom du produit/i.test(t)) { inTable = true; continue; }
        if (!inTable) continue;
        if (/^NOS COORDONNEES|TOTAUX DE CETTE FACTURE/i.test(t)) { inTable = false; continue; }
        const c = r.cells, n = c.length;
        if (n >= 4 && /^\d+(?:[.,]\d+)?$/.test(c[n - 1]) && isAmount(c[n - 2]) && isAmount(c[n - 3])) {
          const dur = n >= 5 ? c[n - 4] : '';
          lines.push({ what: c.slice(0, n >= 5 ? n - 4 : n - 3).join(' '), dur, total: num(c[n - 2]), desc: [] });
        } else if (lines.length) lines[lines.length - 1].desc.push(t);
      }
      return { supplier: 'lisam.cloud', number, date, totalExcl: totalM ? num(totalM[1]) : null, lines: lines.map(l => {
        const desc = l.desc.join(' '), per = /(\d{4}-\d{2}-\d{2})[^>]*->\s*(\d{4}-\d{2}-\d{2})/.exec(desc);
        const label = [desc.replace(/\([^)]*->[^)]*\)/, '').replace(/\s+/g, ' ').trim() || l.what, plain(l.dur).replace(/\s+/g, ' ')].filter(Boolean).join(' – ');
        return { label, detail: [l.what, l.dur].filter(Boolean).join(' · '), key: keyOf(label), domain: domains(l.what)[0] || domains(desc)[0] || '',
          ref: l.what, qty: 1, unit: l.total, total: l.total, from: per?.[1] || null, to: per?.[2] || null };
      }) };
    }
  };

  // Alescom (factures « INV/AAAA/NNNNN », vérifié sur 4 factures 2026 reçues par Peppol dans Billit) :
  //   Facture INV/2026/00209 / Date de facturation | Échéance | … puis 07/07/2026 | 22/07/2026 | …
  //   DESCRIPTION | QUANTITÉ | PRIX UNITAIRE | [REM.%] | TAXES | MONTANT                      ← en-tête
  //   [GEODISHMDO] Intervention à distance via | 1,00 | 96,00 | 20,00 | 21% | 76,80 €         ← [code] nom, montant net de remise
  //   TeamViewer et/ou téléphone / heure - Geoffrey                                            ← suite du nom
  //   16/3 / 15h15-16h33 / vedi / répertoire réseau mathieu …                                  ← détails de l'intervention : jour, heures, client, travail
  //   [TTT] Frais de transport | 1,00 | 5,00 | 21% | 5,00 €
  //   Montant hors taxes | 99,84 €
  // Les lignes « Intervention » ne sont pas refacturées comme achats : elles deviennent des interventions à ajouter au pointage
  // (invoice.interventions) ; un « Déplacement » les marque sur site ; les « Frais de transport » ne sont jamais refacturés.
  const alescom = {
    id: 'alescom', name: 'Alescom',
    detect: rows => rows.some(r => /^ALESCOM\b/i.test(r.text)) && rows.some(r => /^Facture\s+INV\//i.test(r.text)),
    parse(rows) {
      const number = find(rows, /^Facture\s+(INV\/\d{4}\/\d+)/i)?.[1] || '';
      const hi = rows.findIndex(r => /^Date de facturation/i.test(r.text));
      const date = hi >= 0 ? isoDate(rows[hi + 1]?.cells[0]) : null;
      const totalM = find(rows, /Montant hors taxes\s+(-?[\d .  ]+[.,]\d{2})/i);
      const END = /^(Conditions de paiement|Paiement dès|Communication de paiement|Montant hors taxes)/i;
      const DAY = /^(\d{1,2})\/(\d{1,2})(?:\/(\d{2,4}))?\b\s*(.*)$/;
      const items = [];
      let inTable = false;
      for (const r of rows) {
        const t = r.text, c = r.cells, n = c.length;
        if (/^DESCRIPTION\b.*MONTANT/i.test(t)) { inTable = true; continue; }
        if (!inTable) continue;
        if (END.test(t)) { inTable = false; continue; }
        if (n >= 4 && isAmount(c[n - 1]) && /^\d+(?:[.,]\d+)?\s*%$/.test(c[n - 2]) && /^-?\d+(?:[.,]\d+)?$/.test(c[1])) {
          items.push({ desc: [c[0]], qty: num(c[1]), total: num(c[n - 1]), details: [] });
        } else if (items.length) {
          const it = items[items.length - 1];
          // Les détails commencent au premier « jour/mois » ; avant, c'est la suite du nom de l'article.
          if (it.details.length || DAY.test(t)) it.details.push(t); else it.desc.push(t);
        }
      }
      const year = (date || '').slice(0, 4), month = +(date || '').slice(5, 7);
      const lines = items.map(it => {
        const full = it.desc.join(' ').replace(/\s+/g, ' ').trim();
        const code = /^\[([^\]]+)\]\s*/.exec(full)?.[1] || '';
        const label = full.replace(/^\[[^\]]+\]\s*/, '');
        const kind = /frais de transport/i.test(label) ? 'transport' : /intervention/i.test(label) ? 'intervention' : /d[ée]placement/i.test(label) ? 'deplacement' : '';
        // Article ordinaire : ce qui suit le nom (ex. « 18/06/2026 au 30/06/2026 : 2 TB ») est un détail, avec la période éventuelle.
        const more = kind === 'intervention' ? '' : it.details.join(' ');
        return { label, detail: [code && `réf. Alescom ${code}`, more].filter(Boolean).join(' · '), key: code ? 'code:' + code.toLowerCase() : keyOf(label),
          domain: '', ref: code, qty: it.qty, unit: it.qty ? Math.round(it.total / it.qty * 100) / 100 : it.total, total: it.total,
          skip: kind, details: it.details, ...period(more) };
      });
      // Interventions : un groupe par « jour/mois » dans les détails (heures « 15h15-16h33 », puis le client, puis le travail fait).
      const work = lines.filter(l => l.skip === 'intervention');
      const onSite = lines.some(l => l.skip === 'deplacement') || work.some(l => /sur site|sur place/i.test(l.label));
      const tech = work.map(l => / - ([A-Za-zÀ-ÿ' -]+)$/.exec(l.label)?.[1]?.trim()).find(Boolean) || '';
      const minutesQty = work.reduce((s, l) => s + (/\/\s*heure/i.test(l.label) ? l.qty * 60 : /\/\s*minute/i.test(l.label) ? l.qty : 0), 0);
      const hm = s => { const m = /^(\d{1,2})h(\d{0,2})$/i.exec(s); return m ? `${pad(m[1])}:${pad(m[2] || 0)}` : ''; };
      const groups = [];
      for (const t of work.flatMap(l => l.details)) {
        const d = DAY.exec(t);
        let rest = t;
        if (d) {
          let y = d[3] ? (d[3].length === 2 ? '20' + d[3] : d[3]) : year;
          if (!d[3] && year && +d[2] > month) y = String(+year - 1);       // facture de janvier pour une intervention de décembre
          groups.push({ date: y ? `${y}-${pad(d[2])}-${pad(d[1])}` : null, start: '', end: '', client: '', notes: [] });
          rest = d[4].trim();
          if (!rest) continue;
        }
        const g = groups[groups.length - 1];
        if (!g) continue;
        const tm = /^(\d{1,2}h\d{0,2})\s*[-–à]\s*(\d{1,2}h\d{0,2})\s*(.*)$/i.exec(rest);
        if (tm && !g.start) { g.start = hm(tm[1]); g.end = hm(tm[2]); rest = tm[3].trim(); if (!rest) continue; }
        if (!g.client) g.client = rest; else g.notes.push(rest);
      }
      const diff = (a, b) => { const [h1, m1] = a.split(':').map(Number), [h2, m2] = b.split(':').map(Number); return (h2 * 60 + m2) - (h1 * 60 + m1); };
      const interventions = (groups.length ? groups : work.length ? [{ date, start: '', end: '', client: '', notes: [] }] : []).map(g => ({
        date: g.date || date, start: g.start, end: g.end,
        minutes: g.start && g.end && diff(g.start, g.end) > 0 ? diff(g.start, g.end) : groups.length <= 1 ? minutesQty : 0,
        client: g.client, notes: g.notes.join(' ; '), onSite, tech }));
      return { supplier: 'Alescom', number, date, totalExcl: totalM ? num(totalM[1]) : null, lines, interventions };
    }
  };

  const READERS = [ovh, mgk, lisam, alescom];

  // --- e-facture UBL (Peppol, reçue dans Billit) ---
  const ublDoc = xml => new DOMParser().parseFromString(xml, 'application/xml');
  const kids = (el, name) => el ? [...el.children].filter(x => x.localName === name) : [];
  const kid = (el, ...names) => names.reduce((e, n) => kids(e, n)[0], el);
  const txt = el => el?.textContent?.trim() || '';
  // PDF joints à l'e-facture (souvent le PDF d'origine du fournisseur) : [{ name, buffer }].
  function ublPdfs(xml) {
    const d = ublDoc(xml);
    return [...d.getElementsByTagNameNS('*', 'EmbeddedDocumentBinaryObject')]
      .filter(e => /pdf/i.test(e.getAttribute('mimeCode') || '') || /\.pdf$/i.test(e.getAttribute('filename') || ''))
      .map(e => ({ name: e.getAttribute('filename') || 'facture.pdf', buffer: Uint8Array.from(atob(e.textContent.replace(/\s/g, '')), c => c.charCodeAt(0)).buffer }));
  }
  // Facture lue directement dans l'e-facture, quand aucun PDF lisible n'est disponible (PDF fabriqué par Billit).
  // supplier : nom du lecteur (« MGK », « OVH », « lisam.cloud ») ; MGK : la note de la ligne contient le contact, la durée, les postes et le SKU.
  function fromUbl(xml, supplier) {
    const d = ublDoc(xml).documentElement;
    if (!d || d.localName === 'parsererror') return { error: 'E-facture illisible.' };
    // Dates de l'e-facture : seulement « AAAA-MM-JJ » (le contenu vient du fournisseur et est affiché dans l'app).
    const day = s => /^\d{4}-\d{2}-\d{2}$/.test(s) ? s : null;
    const date = day(txt(kid(d, 'IssueDate')));
    const lines = [...kids(d, 'InvoiceLine'), ...kids(d, 'CreditNoteLine')].map(l => {
      const note = txt(kid(l, 'Note')), name = txt(kid(l, 'Item', 'Name')) || note.split('\n')[0];
      const qty = parseFloat(txt(kid(l, 'InvoicedQuantity')) || txt(kid(l, 'CreditedQuantity')) || '1') || 1;
      const total = parseFloat(txt(kid(l, 'LineExtensionAmount'))) || 0, unit = parseFloat(txt(kid(l, 'Price', 'PriceAmount'))) || round(total / qty);
      const p = kid(l, 'InvoicePeriod');
      const from = day(txt(kid(p, 'StartDate'))), to = day(txt(kid(p, 'EndDate')));
      if (supplier === 'MGK') {
        const get = re => (re.exec(note) || [])[1]?.trim() || '';
        return mgkLine({ desc: name, hint: note.split(/\r?\n/)[0].trim(), ref: txt(kid(l, 'Item', 'SellersItemIdentification', 'ID')), qty, unit, total }, date,
          { sku: get(/SKU Code\s*:\s*(\S+)/i), seats: get(/(?:Number of stations|Nombre de postes)\s*:\s*(\d+)/i),
            type: get(/(?:Purchase type|Type d'achat)\s*:\s*(\w+)/i), validity: get(/(?:Validity period|Période de validité)\s*:\s*(\d+\s*[^\s,;]+)/i), from, to });
      }
      const label = name || note;
      return { label, detail: note, key: keyOf(label), domain: domains(name + ' ' + note)[0] || '', ref: '', qty, unit, total, from, to };
    });
    const totalExcl = parseFloat(txt(kid(d, 'LegalMonetaryTotal', 'TaxExclusiveAmount')));
    if (!lines.length) return { error: 'E-facture sans ligne.' };
    return { reader: { id: 'ubl', name: supplier }, invoice: { supplier, number: txt(kid(d, 'ID')), date, totalExcl: isNaN(totalExcl) ? null : totalExcl, lines } };
  }
  const round = n => Math.round(n * 100) / 100;

  // Lit un PDF : { reader, invoice } ou { reader: null } si aucun lecteur ne reconnaît le fournisseur.
  async function read(buffer, fileName) { return fromRows(await readRows(buffer), fileName); }
  // Même chose à partir des lignes déjà extraites (sert aussi aux tests).
  function fromRows(rows, fileName) {
    if (!rows.length) return { rows, reader: null, error: 'PDF sans texte (facture scannée ?) : lecture impossible.' };
    const reader = READERS.find(r => r.detect(rows));
    if (!reader) return { rows, reader: null, error: 'Fournisseur non reconnu : aucun lecteur pour ce format.' };
    const invoice = reader.parse(rows, fileName);
    if (!invoice.lines.length) return { rows, reader, invoice, error: `Facture ${reader.name} lue, mais aucun article trouvé : le lecteur doit être ajusté.` };
    return { rows, reader, invoice };
  }

  // Affiche les pages du PDF en images (canvas) dans el : marche aussi sur téléphone, sans lecteur PDF.
  // Dessinées en plus haute résolution que la largeur affichée, pour rester nettes quand on zoome.
  async function render(buffer, el, maxPages = 10) {
    await load();
    const pdf = await openPdf(buffer);
    const n = Math.min(pdf.numPages, maxPages);
    for (let p = 1; p <= n; p++) {
      const page = await pdf.getPage(p);
      const w = Math.min(Math.max(el.clientWidth, 320) * (window.devicePixelRatio || 1) * 2, 2000);
      const vp = page.getViewport({ scale: w / page.getViewport({ scale: 1 }).width });
      const c = document.createElement('canvas');
      c.width = vp.width; c.height = vp.height;
      c.style.cssText = 'width:100%;display:block;margin-bottom:8px;border:1px solid var(--line);border-radius:6px;background:#fff';
      el.appendChild(c);
      await page.render({ canvasContext: c.getContext('2d'), viewport: vp }).promise;
    }
    return { pages: pdf.numPages, shown: n };
  }

  return { read, readRows, fromRows, render, domains, keyOf, ublPdfs, fromUbl, readers: () => READERS.map(r => ({ id: r.id, name: r.name })) };
})();
