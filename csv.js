// Lecture d'un CSVExport.csv d'Hours Tracker (reprise de l'historique).
// Colonnes : Job,Clocked In,Clocked Out,Duration,Comment,Tags,... — dates "JJ-MM-AA HH:MM", durée "H:MM".
const Csv = (() => {
  const pad = n => String(n).padStart(2, '0');

  function parse(text) {
    const rows = []; let row = [], cur = '', q = false;
    text = (text || '').replace(/^﻿/, '');
    for (let i = 0; i < text.length; i++) {
      const ch = text[i];
      if (q) {
        if (ch === '"' && text[i + 1] === '"') { cur += '"'; i++; }
        else if (ch === '"') q = false;
        else cur += ch;
      } else if (ch === '"') q = true;
      else if (ch === ',') { row.push(cur); cur = ''; }
      else if (ch === '\n' || ch === '\r') {
        if (ch === '\r' && text[i + 1] === '\n') i++;
        row.push(cur); cur = ''; if (row.some(c => c !== '')) rows.push(row); row = [];
      } else cur += ch;
    }
    row.push(cur); if (row.some(c => c !== '')) rows.push(row);
    return rows;
  }

  // "28-08-26 15:25" -> { date:'2026-08-28', time:'15:25' }
  function htDate(s) {
    const m = /^(\d{2})-(\d{2})-(\d{2,4})\s+(\d{1,2}):(\d{2})/.exec((s || '').trim());
    if (!m) return null;
    const y = m[3].length === 2 ? '20' + m[3] : m[3];
    return { date: `${y}-${m[2]}-${m[1]}`, time: `${pad(m[4])}:${m[5]}` };
  }

  // Règles du circuit précédent : "remote"/"distance" dans le Job = pas de déplacement,
  // tag "sur site" = déplacement, tag "alescom" = commanditaire alescom.
  function readHoursTracker(text) {
    const rows = parse(text);
    if (!rows.length) return [];
    const head = rows[0].map(h => h.trim().toLowerCase());
    const col = n => head.indexOf(n);
    const iJob = col('job'), iIn = col('clocked in'), iOut = col('clocked out'),
      iDur = col('duration'), iCom = col('comment'), iTags = col('tags');
    if (iJob < 0 || iIn < 0) throw new Error("Colonnes 'Job' / 'Clocked In' introuvables : ce n'est pas un CSVExport d'Hours Tracker.");
    const out = [];
    for (const r of rows.slice(1)) {
      const job = (r[iJob] || '').trim(); const a = htDate(r[iIn]); if (!job || !a) continue;
      const b = htDate(r[iOut]) || a;
      const tags = (r[iTags] || '').toLowerCase();
      const remote = /remote|distance/i.test(job);
      const client = job.replace(/\b(à|a)\s+distance\b/i, '').replace(/\b(remote|distance)\b/ig, '').replace(/\s+/g, ' ').trim();
      const [dh, dm] = (r[iDur] || '0:00').split(':').map(Number);
      out.push({
        job, client: client || job, date: a.date, start: a.time, end: b.time,
        minutes: (dh || 0) * 60 + (dm || 0), description: (r[iCom] || '').trim(),
        onSite: !remote || tags.includes('sur site'), alescom: tags.includes('alescom'), install: ''
      });
    }
    return out;
  }

  return { readHoursTracker };
})();
