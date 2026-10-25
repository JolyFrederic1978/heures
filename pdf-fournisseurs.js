// Lecture des factures fournisseurs en PDF : texte extrait avec pdf.js (rangé en lignes par position),
// puis un lecteur par fournisseur. Un lecteur renvoie :
//   { supplier, number, date, totalExcl, lines: [{ key, label, detail, domain, qty, unit, total, from, to }] }
// label = nom du produit (sans le domaine du client) ; key = ce nom sans dates ni nombres :
// sert à la table de correspondance avec les articles Billit.
const Fournisseurs = (() => {
  const SRC = 'https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.min.js';
  const WORKER = 'https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.worker.min.js';
  const pad = n => String(n).padStart(2, '0');
  let loading;

  function load() {
    if (window.pdfjsLib) return Promise.resolve();
    return loading ||= new Promise((resolve, reject) => {
      const s = document.createElement('script');
      s.src = SRC;
      s.onload = () => { pdfjsLib.GlobalWorkerOptions.workerSrc = WORKER; resolve(); };
      s.onerror = () => { loading = null; reject(new Error('Impossible de charger pdf.js (connexion requise la première fois).')); };
      document.head.appendChild(s);
    });
  }

  // Lignes du PDF de haut en bas : { page, cells: [texte des colonnes], text }.
  async function readRows(buffer) {
    await load();
    const pdf = await pdfjsLib.getDocument({ data: new Uint8Array(buffer) }).promise;
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

  const READERS = [ovh];

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

  return { read, readRows, domains, keyOf, readers: () => READERS.map(r => ({ id: r.id, name: r.name })) };
})();
