// Fichiers Excel : liste à facturer, archive, fichier Prestations (format repris par la facture Billit)
// et lecture de l'ancien ArchiveTemplate.xlsx pour la migration.
const Xlsx = (() => {
  const SRC = 'https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js';
  const MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  const pad = n => String(n).padStart(2, '0');
  let loading;

  function load() {
    if (window.ExcelJS) return Promise.resolve();
    return loading ||= new Promise((resolve, reject) => {
      const s = document.createElement('script');
      s.src = SRC; s.onload = resolve;
      s.onerror = () => { loading = null; reject(new Error('Impossible de charger ExcelJS (connexion requise la première fois).')); };
      document.head.appendChild(s);
    });
  }

  const date = d => { const [y, m, j] = d.split('-').map(Number); return new Date(Date.UTC(y, m - 1, j)); };
  // Forfaits (maintenance / mise à jour serveur) : pas d'heures, cellule vide.
  const time = hm => { if (!hm) return null; const [h, m] = hm.split(':').map(Number); return new Date(Date.UTC(1899, 11, 30, h, m)); };

  async function build(head, rows, { dateCols = [], timeCols = [], pctCols = [], moneyCols = [], widths = [], total = null, filter = false } = {}) {
    await load();
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Feuil1');
    ws.addRow(head);
    rows.forEach(r => ws.addRow(r));
    if (total) ws.addRow(total);
    ws.eachRow((row, n) => row.eachCell({ includeEmpty: true }, c => {
      c.font = { name: 'Arial', size: 11, bold: n === 1 || (!!total && n === ws.rowCount) };
    }));
    dateCols.forEach(i => ws.getColumn(i + 1).numFmt = 'dd/mm/yyyy');
    timeCols.forEach(i => ws.getColumn(i + 1).numFmt = 'hh:mm');
    pctCols.forEach(i => ws.getColumn(i + 1).numFmt = '0%');
    moneyCols.forEach(i => ws.getColumn(i + 1).numFmt = '#,##0.00 €');
    widths.forEach((w, i) => ws.getColumn(i + 1).width = w);
    ws.views = [{ state: 'frozen', ySplit: 1 }];
    if (filter) ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: head.length } };
    return new Blob([await wb.xlsx.writeBuffer()], { type: MIME });
  }

  // r : { client, date, start, end, description, minutes, onSite, alescom, install }
  const BASE = ['client', 'date', 'heure debut', 'heure de fin', 'description', 'heures', 'minutes', 'deplacement', "Donneur d'ordre", 'Installation/réinstallation'];
  const base = r => [r.client, date(r.date), time(r.start), time(r.end), r.description || null,
    Math.floor(r.minutes / 60), r.minutes % 60, r.onSite ? 1 : null, r.alescom ? 'alescom' : null, r.install || null];
  const BASE_W = [16, 12, 11, 11, 50, 8, 8, 11, 15, 24];

  const openList = rows => build(BASE, rows.map(base), { dateCols: [1], timeCols: [2, 3], widths: BASE_W, filter: true });

  // r en plus : { billDate, billRef, amount }
  const archive = rows => build(['facturé le', 'n° facture', ...BASE, 'montant HTVA'],
    rows.map(r => [date(r.billDate), r.billRef || null, ...base(r), Math.round(r.amount * 100) / 100]),
    { dateCols: [0, 3], timeCols: [4, 5], widths: [12, 12, ...BASE_W, 13], filter: true });

  // Même format que les fichiers "Prestations - Client AAAA-MM-JJ.xlsx" existants (ligne Total = sommes brutes).
  function prestations(rows, withClient) {
    const head = [...(withClient ? ['client'] : []), 'date', 'heure debut', 'heure de fin', 'description',
      'heures', 'minutes', 'deplacement', 'Installaton/réinstallation'];
    let sh = 0, sm = 0, sd = 0;
    const data = rows.map(r => {
      const h = Math.floor(r.minutes / 60), m = r.minutes % 60;
      sh += h; sm += m; sd += r.onSite ? 1 : 0;
      return [...(withClient ? [r.client] : []), date(r.date), time(r.start), time(r.end), r.description || null,
        h, m, r.onSite ? 1 : null, r.install || null];
    });
    const off = withClient ? 1 : 0;
    const total = ['Total', ...Array(head.length - 1).fill(null)];
    total[off + 4] = sh; total[off + 5] = sm; total[off + 6] = sd;
    return build(head, data, { dateCols: [off], timeCols: [off + 1, off + 2], widths: [...(withClient ? [16] : []), 12, 11, 11, 50, 8, 8, 11, 24], total });
  }

  // --- lecture ArchiveTemplate.xlsx ---
  const norm = s => String(s || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^a-z]/g, '');
  const val = c => {
    let v = c?.value;
    if (v && typeof v === 'object' && !(v instanceof Date)) v = v.result ?? v.text ?? (v.richText ? v.richText.map(t => t.text).join('') : null);
    return v;
  };
  const EPOCH = Date.UTC(1899, 11, 30);
  function asDate(v) {
    if (typeof v === 'number') v = new Date(EPOCH + v * 864e5);
    if (v instanceof Date) { const d = new Date(v.getTime() + 30000); return `${d.getUTCFullYear()}-${pad(d.getUTCMonth() + 1)}-${pad(d.getUTCDate())}`; }
    const m = /(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})/.exec(v || '');
    return m ? `${m[3].length === 2 ? '20' + m[3] : m[3]}-${pad(m[2])}-${pad(m[1])}` : null;
  }
  function asTime(v) {
    if (typeof v === 'number') { const t = Math.round((v % 1) * 1440); return `${pad(Math.floor(t / 60) % 24)}:${pad(t % 60)}`; }
    if (v instanceof Date) { const t = ((Math.round(v.getTime() / 60000) % 1440) + 1440) % 1440; return `${pad(Math.floor(t / 60))}:${pad(t % 60)}`; }
    const m = /(\d{1,2})[:h](\d{2})/.exec(v || '');
    return m ? `${pad(m[1])}:${m[2]}` : null;
  }

  async function readArchiveTemplate(buffer) {
    await load();
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buffer);
    const ws = wb.worksheets[0];
    const cols = {};
    ws.getRow(1).eachCell((c, i) => { const k = norm(val(c)); cols[k.startsWith('installa') ? 'install' : k] = i; });
    if (!cols.client || !cols.date) throw new Error("Colonnes 'client' / 'date' introuvables en ligne 1.");
    const out = [];
    ws.eachRow((row, n) => {
      if (n === 1) return;
      const g = k => cols[k] ? val(row.getCell(cols[k])) : null;
      const job = String(g('client') || '').trim(), d = asDate(g('date'));
      if (!job || !d || /^total$/i.test(job)) return;
      const start = asTime(g('heuredebut')) || '00:00';
      const minutes = (+g('heures') || 0) * 60 + (+g('minutes') || 0);
      out.push({
        job, client: job.replace(/\b(à|a)\s+distance\b/i, '').replace(/\b(remote|distance)\b/ig, '').replace(/\s+/g, ' ').trim() || job,
        date: d, start, end: asTime(g('heuredefin')) || start, minutes,
        description: String(g('description') || '').trim(),
        onSite: +g('deplacement') > 0, alescom: /alescom/i.test(String(g('donneurdordre') || '')),
        install: String(g('install') || '').trim()
      });
    });
    return out;
  }

  return { openList, archive, prestations, readArchiveTemplate, build, MIME };
})();
