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
        const sku = get(/SKU Code\s*:\s*(\S+)/i), seats = get(/Nombre de postes\s*:\s*(\d+)/i);
        const type = get(/Type d'achat\s*:\s*(\w+)/i), validity = get(/Période de validité\s*:\s*(\d+\s*[^\s,;]+)/i);
        // Période : « Valide du A jusqu'au B », ou « Valide jusqu'au B », ou seulement la durée (calculée depuis la date du document).
        let from = isoDate(get(/Valide du\s+(\d{1,2}\.\d{1,2}\.\d{4})/i)), to = isoDate(get(/jusqu'au\s+(\d{1,2}\.\d{1,2}\.\d{4})/i));
        const years = +(/(\d+)\s*ann/i.exec(validity) || [])[1] || 0, months = +(/(\d+)\s*mois/i.exec(validity) || [])[1] || 0;
        const shift = (d, sign) => { const x = new Date(d + 'T00:00:00Z'); x.setUTCFullYear(x.getUTCFullYear() + sign * years); x.setUTCMonth(x.getUTCMonth() + sign * months); x.setUTCDate(x.getUTCDate() - sign); return x.toISOString().slice(0, 10); };
        if (!from && to && (years || months)) from = shift(to, -1);
        if (!from && !to && date && (years || months)) { from = date; to = shift(date, 1); }
        const label = [it.desc, seats && `${seats} poste${+seats > 1 ? 's' : ''}`, validity].filter(Boolean).join(' – ');
        return { label, detail: [type && `Type d'achat : ${type}`, sku && `SKU ${sku}`, `réf. MGK ${it.ref}`].filter(Boolean).join(' · '),
          key: sku ? 'sku:' + sku.toLowerCase() : 'ref:' + it.ref + ':' + keyOf(it.desc),
          domain: '', hint: it.hint, ref: it.ref, qty: it.qty, unit: it.unit, total: it.total, from, to };
      });
      return { supplier: 'MGK', number, date, totalExcl: totalM ? num(totalM[1]) : null, lines };
    }
  };

  const READERS = [ovh, mgk];

  // Lit un PDF : { reader, invoice } ou { reader: null } si aucun lecteur ne reconnaît le fournisseur.
  async function read(buffer, fileName) {
    const rows = await readRows(buffer);
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

  return { read, readRows, render, domains, keyOf, readers: () => READERS.map(r => ({ id: r.id, name: r.name })) };
})();
