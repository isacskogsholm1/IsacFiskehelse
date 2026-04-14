const http = require('http');
const https = require('https');
const fs = require('fs');
const path = require('path');
const os = require('os');
const crypto = require('crypto');

// ── Password hashing (built-in, no npm needed) ────────────────────────────
function hashPassword(password) {
  return crypto.createHash('sha256').update(password + 'vela_salt_2025').digest('hex');
}
function verifyPassword(password, hash) {
  return hashPassword(password) === hash;
}

// Load .env for local development
try {
  const envFile = path.join(__dirname, '.env');
  fs.readFileSync(envFile, 'utf8').split('\n').forEach(line => {
    const [k, ...v] = line.split('=');
    if (k && v.length && !process.env[k.trim()]) process.env[k.trim()] = v.join('=').trim();
  });
} catch {}

function getLocalIP() {
  const ifaces = os.networkInterfaces();
  for (const name of Object.keys(ifaces)) {
    for (const iface of ifaces[name]) {
      if (iface.family === 'IPv4' && !iface.internal) return iface.address;
    }
  }
  return 'localhost';
}
const ExcelJS = require('exceljs');
const { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, BorderStyle, TableRow, TableCell, Table, WidthType, ShadingType } = require('docx');
const dir  = __dirname;
const port = process.env.PORT || 7823;
const mime = { '.html':'text/html', '.js':'application/javascript', '.css':'text/css', '.json':'application/json', '.png':'image/png', '.ico':'image/x-icon', '.svg':'image/svg+xml' };

// ── BarentsWatch credentials ───────────────────────────────────────────────
const BW_CLIENT_ID     = process.env.BW_CLIENT_ID     || 'isacskogsholm1@live.no:AquAI';
const BW_CLIENT_SECRET = process.env.BW_CLIENT_SECRET || 'm,sbog17ksBrevika';

// ── Groq API (gratis — Whisper + Llama 3.3 70B) ───────────────────────────
// Gratis nøkkel på: console.groq.com
const GROQ_API_KEY   = process.env.GROQ_API_KEY   || '';
const OPENAI_API_KEY = process.env.OPENAI_API_KEY  || ''; // fallback for Whisper


// ── Users / sessions ───────────────────────────────────────────────────────
const USERS_FILE = path.join(dir, 'users.json');
function loadUsers() {
  // Try file first, fall back to VELA_USERS env var (for cloud hosting)
  try { return JSON.parse(fs.readFileSync(USERS_FILE, 'utf8')); } catch {}
  try { return JSON.parse(process.env.VELA_USERS || '[]'); } catch { return []; }
}
function saveUsers(users) {
  try { fs.writeFileSync(USERS_FILE, JSON.stringify(users, null, 2)); } catch {}
}
const _sessions = {}; // token → { userId, expires }
function makeToken() { return Math.random().toString(36).slice(2) + Date.now().toString(36); }
function checkSession(req) {
  const auth = req.headers['authorization'] || '';
  const token = auth.replace('Bearer ', '').trim();
  const s = _sessions[token];
  if (!s || Date.now() > s.expires) return null;
  return s.userId;
}

// ── Token cache (server-side) ──────────────────────────────────────────────
let _cachedToken = null;
let _tokenExpiry = 0;

async function getServerBWToken() {
  if (_cachedToken && Date.now() < _tokenExpiry - 30000) return _cachedToken;
  return new Promise((resolve) => {
    const body = `grant_type=client_credentials&client_id=${encodeURIComponent(BW_CLIENT_ID)}&client_secret=${encodeURIComponent(BW_CLIENT_SECRET)}&scope=api`;
    const options = {
      hostname: 'id.barentswatch.no',
      path: '/connect/token',
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
        'Content-Length': Buffer.byteLength(body),
      }
    };
    const req = https.request(options, (res) => {
      let data = '';
      res.on('data', c => data += c);
      res.on('end', () => {
        try {
          const json = JSON.parse(data);
          if (json.access_token) {
            _cachedToken = json.access_token;
            _tokenExpiry = Date.now() + (json.expires_in || 3600) * 1000;
            resolve(json.access_token);
          } else {
            resolve(null);
          }
        } catch(e) { resolve(null); }
      });
    });
    req.on('error', () => resolve(null));
    req.write(body);
    req.end();
  });
}

// ── Excel export helper ────────────────────────────────────────────────────────
function parseBody(req) {
  return new Promise((resolve, reject) => {
    let body = '';
    req.on('data', c => body += c);
    req.on('end', () => { try { resolve(JSON.parse(body)); } catch(e) { reject(e); } });
  });
}

// ── Real .docx generation ─────────────────────────────────────────────────────
async function buildDocx(payload) {
  const { title = 'Dokument', content = '', aiText = '' } = payload;
  const date = new Date().toLocaleDateString('no-NO', { day:'2-digit', month:'long', year:'numeric' });

  const BLUE    = '054370';
  const MID     = '0B72B5';
  const LIGHT   = 'B3D9F2';
  const GRAY    = '666666';
  const BG_BLUE = 'EDF6FB';

  const children = [];

  // ── Header table: title + branding ──
  children.push(new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: { top:{style:BorderStyle.NONE}, bottom:{style:BorderStyle.NONE}, left:{style:BorderStyle.NONE}, right:{style:BorderStyle.NONE}, insideH:{style:BorderStyle.NONE}, insideV:{style:BorderStyle.NONE} },
    rows: [new TableRow({ children: [
      new TableCell({
        width: { size: 70, type: WidthType.PERCENTAGE },
        borders: { top:{style:BorderStyle.NONE}, bottom:{style:BorderStyle.NONE}, left:{style:BorderStyle.NONE}, right:{style:BorderStyle.NONE} },
        children: [
          new Paragraph({ children: [new TextRun({ text: title, bold:true, size:44, color:BLUE, font:'Calibri' })], spacing:{after:80} }),
          new Paragraph({ children: [
            new TextRun({ text:'Dato: ', bold:true, size:18, color:GRAY, font:'Calibri' }),
            new TextRun({ text:date+' · Generert av Vela', size:18, color:GRAY, font:'Calibri' }),
          ], spacing:{after:200}, border:{ bottom:{ color:LIGHT, size:6, style:BorderStyle.SINGLE, space:4 } } }),
        ]
      }),
      new TableCell({
        width: { size: 30, type: WidthType.PERCENTAGE },
        borders: { top:{style:BorderStyle.NONE}, bottom:{style:BorderStyle.NONE}, left:{style:BorderStyle.NONE}, right:{style:BorderStyle.NONE} },
        shading: { fill: BLUE, type: ShadingType.CLEAR, color: BLUE },
        children: [
          new Paragraph({ alignment: AlignmentType.CENTER, children:[new TextRun({ text:'Vela', bold:true, size:32, color:'FFFFFF', font:'Calibri' })], spacing:{before:80} }),
          new Paragraph({ alignment: AlignmentType.CENTER, children:[new TextRun({ text:'Fiskehelsebiolog', size:18, color:'AADDFF', font:'Calibri' })] }),
        ]
      }),
    ]})]
  }));
  children.push(new Paragraph({ spacing:{ before:0, after:200 } }));

  // ── Parse content into paragraphs ──
  const lines = content.split('\n');
  let sectionKV = [];

  const flushKV = () => { sectionKV = []; };

  lines.forEach(raw => {
    const line = raw.trim();

    if (!line) { flushKV(); children.push(new Paragraph({ spacing:{before:0,after:80} })); return; }

    // Divider
    if (/^[━─=\-]{3,}$/.test(line)) {
      flushKV();
      children.push(new Paragraph({ border:{ bottom:{ color:LIGHT, size:6, style:BorderStyle.SINGLE, space:2 } }, spacing:{before:120,after:120} }));
      return;
    }

    // ALL CAPS section header
    if (/^(\d+[.)]\s+)?[A-ZÆØÅ][A-ZÆØÅ\s\(\)\/\-0-9]{3,}$/.test(line)) {
      flushKV();
      children.push(new Paragraph({
        children: [new TextRun({ text:line, bold:true, size:22, color:MID, font:'Calibri', allCaps:true })],
        spacing:{ before:280, after:80 },
        border:{ bottom:{ color:'C7E8FA', size:4, style:BorderStyle.SINGLE, space:2 } }
      }));
      return;
    }

    // Key: Value
    const m = line.match(/^([^:]{2,40}):\s*(.+)$/);
    if (m) {
      sectionKV.push({ key:m[1].trim(), val:m[2].trim() });
      children.push(new Paragraph({
        children: [
          new TextRun({ text:m[1]+': ', bold:true, size:20, color:BLUE, font:'Calibri' }),
          new TextRun({ text:m[2], size:20, color:'1a2744', font:'Calibri' }),
        ],
        spacing:{ before:40, after:40 }
      }));
      return;
    }

    // Bullet
    if (/^[•·\-*]\s/.test(raw)) {
      flushKV();
      children.push(new Paragraph({
        children: [new TextRun({ text:line.replace(/^[•·\-*]\s+/,''), size:20, font:'Calibri' })],
        bullet:{ level:0 },
        spacing:{ before:40, after:40 }
      }));
      return;
    }

    flushKV();
    children.push(new Paragraph({ children:[new TextRun({ text:line, size:20, font:'Calibri' })], spacing:{before:40,after:40} }));
  });

  // ── AI analyse section ──
  if (aiText && aiText.trim()) {
    children.push(new Paragraph({ spacing:{before:200,after:80} }));
    children.push(new Paragraph({
      children:[new TextRun({ text:'✦ AI-ANALYSE', bold:true, size:22, color:MID, font:'Calibri', allCaps:true })],
      spacing:{before:120,after:80}
    }));
    children.push(new Paragraph({
      children:[new TextRun({ text:aiText.trim(), size:20, font:'Calibri', color:'1a2744' })],
      shading:{ fill:BG_BLUE, type:ShadingType.CLEAR, color:BG_BLUE },
      spacing:{before:80,after:80},
      indent:{ left:200, right:200 }
    }));
  }

  // ── Footer ──
  children.push(new Paragraph({ spacing:{before:400,after:40}, border:{ top:{ color:LIGHT, size:4, style:BorderStyle.SINGLE, space:4 } } }));
  children.push(new Paragraph({
    alignment: AlignmentType.CENTER,
    children:[new TextRun({ text:`Vela · Havbruksdokumentasjon · ${date}`, size:16, color:GRAY, font:'Calibri' })]
  }));

  const doc = new Document({
    creator:'Vela', title, description:'Generert av Vela Havbruksdokumentasjon',
    styles:{ default:{ document:{ run:{ font:'Calibri', size:20 } } } },
    sections:[{ properties:{}, children }]
  });

  return Packer.toBuffer(doc);
}

async function buildVektExcel(payload) {
  const { lok, merd, dato, merknad, fisker, stats, statsLengde, chartPng } = payload;
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Vela'; wb.created = new Date();

  const _fmt = (v) => (v != null && !isNaN(v)) ? +parseFloat(v).toFixed(1) : '—';

  // ── Sheet 1: Snittvekt (single dashboard sheet) ───────────────────────────
  const ws = wb.addWorksheet('Snittvekt', { views: [{ showGridLines: true }] });

  // Row 1: Title bar spanning A1:H1
  ws.mergeCells('A1:H1');
  const titleCell = ws.getCell('A1');
  titleCell.value = `Snittvekt — ${lok} — Merd ${merd} — ${dato}`;
  titleCell.font = { size: 14, bold: true, color: { argb: 'FFFFFFFF' } };
  titleCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0B72B5' } };
  titleCell.alignment = { horizontal: 'left', vertical: 'middle' };
  ws.getRow(1).height = 28;

  // Row 2: Column headers
  const headers = ['Nr', 'Vekt (g)', 'Lengde (cm)', 'K-faktor', 'Avvik (g)', 'Avvik (%)', 'Status'];
  headers.forEach((h, i) => {
    const cell = ws.getCell(2, i + 1);
    cell.value = h;
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11 };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF054370' } };
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
    cell.border = { bottom: { style: 'thin', color: { argb: 'FF0B72B5' } } };
  });
  ws.getRow(2).height = 22;

  // Data rows
  fisker.forEach((f, i) => {
    const vektVal = typeof f === 'number' ? f : f.vekt;
    const lengdeVal = typeof f === 'number' ? null : (f.lengde ?? null);

    const kFaktor = (lengdeVal != null && lengdeVal > 0)
      ? +((vektVal / Math.pow(lengdeVal, 3)) * 100).toFixed(3)
      : '—';
    const avvik = +(vektVal - stats.mean).toFixed(1);
    const pct   = +((vektVal - stats.mean) / stats.mean * 100).toFixed(1);
    const status = vektVal > stats.mean * 1.25 ? '▲ Over' : vektVal < stats.mean * 0.75 ? '▼ Under' : '● OK';
    const fillColor = i % 2 === 0 ? 'FFFAFCFF' : 'FFE6F4FB';
    const statusColor = vektVal > stats.mean * 1.25 ? 'FFE74C3C' : vektVal < stats.mean * 0.75 ? 'FFF39C12' : 'FF27AE60';

    const row = ws.getRow(i + 3);
    [i + 1, vektVal, lengdeVal ?? '—', kFaktor, avvik, pct, status].forEach((val, ci) => {
      const cell = row.getCell(ci + 1);
      cell.value = val;
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: fillColor } };
      cell.alignment = { horizontal: ci === 0 ? 'center' : ci === 6 ? 'center' : 'right', vertical: 'middle' };
      if (ci === 6) cell.font = { bold: true, color: { argb: statusColor } };
      if ((ci === 4 || ci === 5) && avvik < 0) cell.font = { color: { argb: 'FFE74C3C' } };
      cell.border = { bottom: { style: 'hair', color: { argb: 'FFDDEEEE' } } };
    });
    row.height = 18;
  });

  // SNITT row
  const snittRow = ws.getRow(fisker.length + 3);
  const hasLengder = fisker.some(f => typeof f !== 'number' && f.lengde != null);
  const snittLengde = hasLengder && statsLengde ? _fmt(statsLengde.mean) : '—';
  ['SNITT', _fmt(stats.mean), snittLengde, '—', '—', '—', ''].forEach((v, i) => {
    const cell = snittRow.getCell(i + 1);
    cell.value = i === 1 || i === 2 ? (typeof v === 'string' && v !== '—' ? +v : v) : v;
    cell.font = { bold: true, color: { argb: 'FF054370' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD0E8F8' } };
    cell.alignment = { horizontal: i === 0 ? 'left' : 'right', vertical: 'middle' };
  });
  snittRow.height = 20;

  // Stats summary block below SNITT
  const statsStartRow = fisker.length + 5;
  const _statBlock = (label, st, rowOffset) => {
    if (!st) return;
    const headerCell = ws.getCell(statsStartRow + rowOffset, 1);
    headerCell.value = label;
    headerCell.font = { bold: true, size: 11, color: { argb: 'FF054370' } };
    ws.mergeCells(statsStartRow + rowOffset, 1, statsStartRow + rowOffset, 4);

    const statRows = [
      ['n (antall)', st.n ?? fisker.length],
      ['Snitt', _fmt(st.mean)],
      ['Median', _fmt(st.median)],
      ['SD', _fmt(st.sd)],
      ['CV%', st.cv != null ? +parseFloat(st.cv).toFixed(1) : '—'],
      ['Min', st.min ?? '—'],
      ['Maks', st.max ?? '—'],
    ];
    statRows.forEach(([lbl, val], si) => {
      const r = statsStartRow + rowOffset + 1 + si;
      const lCell = ws.getCell(r, 1);
      lCell.value = lbl;
      lCell.font = { bold: true, color: { argb: 'FF054370' } };
      lCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: si % 2 === 0 ? 'FFFAFCFF' : 'FFE6F4FB' } };
      lCell.alignment = { vertical: 'middle' };
      const vCell = ws.getCell(r, 2);
      vCell.value = val;
      vCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: si % 2 === 0 ? 'FFFAFCFF' : 'FFE6F4FB' } };
      vCell.alignment = { horizontal: 'right', vertical: 'middle' };
    });
  };

  _statBlock('Statistikk — Vekt (g)', stats, 0);
  if (statsLengde) {
    _statBlock('Statistikk — Lengde (cm)', statsLengde, 10);
  }
  if (merknad) {
    const mRow = statsStartRow + (statsLengde ? 20 : 10);
    const mLabel = ws.getCell(mRow, 1);
    mLabel.value = 'Merknad';
    mLabel.font = { bold: true, color: { argb: 'FF054370' } };
    const mVal = ws.getCell(mRow, 2);
    mVal.value = merknad;
    ws.mergeCells(mRow, 2, mRow, 4);
  }

  // Column widths
  ws.columns = [
    { width: 10 }, { width: 12 }, { width: 12 }, { width: 12 }, { width: 14 }, { width: 12 }, { width: 16 }
  ];

  // Embed chart PNG
  if (chartPng) {
    const imgId = wb.addImage({ base64: chartPng.replace(/^data:image\/png;base64,/, ''), extension: 'png' });
    ws.addImage(imgId, { tl: { col: 8, row: 1 }, br: { col: 18, row: 26 } });
  }

  return wb;
}

async function buildIndividExcel(payload) {
  const { lok, merd, dato, fisker, avgV, avgL, avgK, scoreParams, chartVektPng, chartWelferdPng } = payload;
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Vela'; wb.created = new Date();

  // Total columns: 4 (biometrics) + 13 (welfare params) + 1 (sum) = 18
  const totalCols = 4 + scoreParams.length + 1;
  // Helper to get last column letter (handles beyond Z)
  const colLetter = (n) => n <= 26 ? String.fromCharCode(64 + n) : String.fromCharCode(64 + Math.floor((n-1)/26)) + String.fromCharCode(65 + ((n-1) % 26));

  // ── Sheet 1: Individkontroll ──────────────────────────────────────────────
  const ws = wb.addWorksheet('Individkontroll', { views: [{ showGridLines: true }] });

  // Row 1: Title spanning all columns
  ws.mergeCells(1, 1, 1, totalCols);
  const tc = ws.getCell('A1');
  tc.value = `Individkontroll — ${lok} — Merd ${merd} — ${dato}`;
  tc.font = { size: 14, bold: true, color: { argb: 'FFFFFFFF' } };
  tc.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0B72B5' } };
  tc.alignment = { horizontal: 'left', vertical: 'middle' };
  ws.getRow(1).height = 28;

  // Row 2: Headers — Nr | Vekt (g) | Lengde (cm) | K-faktor | [welfare params] | Sum
  const heads = ['Nr', 'Vekt (g)', 'Lengde (cm)', 'K-faktor', ...scoreParams, 'Sum'];
  heads.forEach((h, i) => {
    const cell = ws.getCell(2, i + 1);
    cell.value = h;
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 10 };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF054370' } };
    cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    cell.border = { bottom: { style: 'thin', color: { argb: 'FF0B72B5' } } };
  });
  ws.getRow(2).height = 30;

  // Score color helpers
  const scoreItemFg = (sc) => sc === 0 ? 'FF1E8449' : sc === 1 ? 'FFE67E22' : sc === 2 ? 'FFD35400' : 'FFC0392B';
  const scoreItemBg = (sc) => sc === 0 ? 'FFD5F5E3' : sc === 1 ? 'FFFDEBD0' : sc === 2 ? 'FFFAD7A0' : 'FFFDEDEC';
  const sumBg = (s) => s === 0 ? 'FFD5F5E3' : s <= 3 ? 'FFFDEBD0' : s <= 6 ? 'FFFAD7A0' : 'FFFDEDEC';
  const sumFg = (s) => s === 0 ? 'FF1E8449' : s <= 3 ? 'FFE67E22' : s <= 6 ? 'FFD35400' : 'FFC0392B';

  // Data rows
  fisker.forEach((f, i) => {
    const vektVal = typeof f === 'number' ? f : f.vekt;
    const lengdeVal = typeof f === 'number' ? null : (f.lengde ?? null);
    const kVal = typeof f === 'number' ? null : (f.k ?? null);
    const scores = (typeof f === 'number' ? [] : f.scores) || [];
    const scoreSum = typeof f === 'number' ? 0 : (f.score ?? scores.reduce((a, b) => a + b, 0));

    const kColor = kVal == null ? 'FF999999' : kVal >= 1.0 ? 'FF27AE60' : kVal >= 0.8 ? 'FFF39C12' : 'FFE74C3C';
    const rowFill = i % 2 === 0 ? 'FFFAFCFF' : 'FFE6F4FB';
    const row = ws.getRow(i + 3);

    // Cols 1-4: biometrics
    [i + 1, vektVal, lengdeVal ?? '—', kVal != null ? +kVal.toFixed(3) : '—'].forEach((val, ci) => {
      const cell = row.getCell(ci + 1);
      cell.value = val;
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: rowFill } };
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
      if (ci === 3) cell.font = { bold: true, color: { argb: kColor } };
      cell.border = { bottom: { style: 'hair', color: { argb: 'FFDDEEEE' } } };
    });

    // Cols 5 to 4+scoreParams.length: individual welfare scores
    scores.forEach((sc, si) => {
      const cell = row.getCell(5 + si);
      cell.value = sc;
      cell.font = { bold: true, color: { argb: scoreItemFg(sc) } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: scoreItemBg(sc) } };
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
      cell.border = { bottom: { style: 'hair', color: { argb: 'FFDDEEEE' } } };
    });
    // Fill any missing score cells with blank
    for (let si = scores.length; si < scoreParams.length; si++) {
      const cell = row.getCell(5 + si);
      cell.value = '';
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: rowFill } };
      cell.border = { bottom: { style: 'hair', color: { argb: 'FFDDEEEE' } } };
    }

    // Col last: sum
    const sumCell = row.getCell(4 + scoreParams.length + 1);
    sumCell.value = scoreSum;
    sumCell.font = { bold: true, color: { argb: sumFg(scoreSum) } };
    sumCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: sumBg(scoreSum) } };
    sumCell.alignment = { horizontal: 'center', vertical: 'middle' };
    sumCell.border = { bottom: { style: 'hair', color: { argb: 'FFDDEEEE' } } };

    row.height = 18;
  });

  // SNITT row
  const snittRowIdx = fisker.length + 3;
  const snittRow = ws.getRow(snittRowIdx);
  const avgScores = scoreParams.map((_, pi) =>
    +(fisker.map(f => ((f.scores || [])[pi] ?? 0)).reduce((a, b) => a + b, 0) / (fisker.length || 1)).toFixed(2)
  );
  const avgSum = +(avgScores.reduce((a, b) => a + b, 0)).toFixed(2);

  ['SNITT', avgV != null ? +avgV : '—', avgL != null ? +avgL : '—', avgK != null ? +avgK : '—', ...avgScores, avgSum].forEach((v, i) => {
    const cell = snittRow.getCell(i + 1);
    cell.value = v;
    cell.font = { bold: true, color: { argb: 'FF054370' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD0E8F8' } };
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
  });
  snittRow.height = 20;

  // Column widths
  ws.columns = [
    { width: 8 }, { width: 10 }, { width: 11 }, { width: 10 },
    ...scoreParams.map(() => ({ width: 11 })),
    { width: 10 },
  ];

  // Embed charts to the right
  if (chartVektPng) {
    const id = wb.addImage({ base64: chartVektPng.replace(/^data:image\/png;base64,/, ''), extension: 'png' });
    ws.addImage(id, { tl: { col: 19, row: 1 }, br: { col: 30, row: 18 } });
  }
  if (chartWelferdPng) {
    const id2 = wb.addImage({ base64: chartWelferdPng.replace(/^data:image\/png;base64,/, ''), extension: 'png' });
    ws.addImage(id2, { tl: { col: 19, row: 19 }, br: { col: 30, row: 36 } });
  }

  // ── Sheet 2: Velferd-oversikt ─────────────────────────────────────────────
  const ws2 = wb.addWorksheet('Velferd-oversikt');

  // Title
  ws2.mergeCells('A1:M1');
  const vc = ws2.getCell('A1');
  vc.value = `Velferd-oversikt — ${lok} — Merd ${merd} — ${dato}`;
  vc.font = { bold: true, size: 13, color: { argb: 'FFFFFFFF' } };
  vc.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0B72B5' } };
  vc.alignment = { horizontal: 'left', vertical: 'middle' };
  ws2.getRow(1).height = 26;

  // Category summary header
  const catHeaderRow = 3;
  ['Kategori', 'Antall fisk', '% av total'].forEach((h, ci) => {
    const cell = ws2.getCell(catHeaderRow, ci + 1);
    cell.value = h;
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF054370' } };
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
  });
  ws2.getRow(catHeaderRow).height = 20;

  const scoreSums = fisker.map(f => (f.scores || []).reduce((a, b) => a + b, 0));
  const categories = [
    ['OK  (sum = 0)',      scoreSums.filter(s => s === 0).length,       'FF27AE60'],
    ['Mild (sum 1–3)',     scoreSums.filter(s => s >= 1 && s <= 3).length, 'FFF39C12'],
    ['Moderat (sum 4–6)',  scoreSums.filter(s => s >= 4 && s <= 6).length, 'FFE67E22'],
    ['Alvorlig (sum > 6)', scoreSums.filter(s => s > 6).length,          'FFE74C3C'],
  ];
  categories.forEach(([lbl, cnt, color], ri) => {
    const pct = fisker.length ? `${Math.round(cnt / fisker.length * 100)}%` : '0%';
    [lbl, cnt, pct].forEach((val, ci) => {
      const cell = ws2.getCell(catHeaderRow + 1 + ri, ci + 1);
      cell.value = val;
      if (ci === 0) cell.font = { bold: true };
      if (ci === 1) cell.font = { bold: true, color: { argb: color } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: ri % 2 === 0 ? 'FFFAFCFF' : 'FFE6F4FB' } };
      cell.alignment = { horizontal: ci === 0 ? 'left' : 'center', vertical: 'middle' };
    });
    ws2.getRow(catHeaderRow + 1 + ri).height = 18;
  });

  // Average score per parameter table
  const avgTableStart = catHeaderRow + categories.length + 3;
  ['Parameter', 'Gj.snitt score', 'Andel med funn'].forEach((h, ci) => {
    const cell = ws2.getCell(avgTableStart, ci + 1);
    cell.value = h;
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0B72B5' } };
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
  });
  ws2.getRow(avgTableStart).height = 20;

  scoreParams.forEach((param, pi) => {
    const paramScores = fisker.map(f => (f.scores || [])[pi] ?? 0);
    const avg = +(paramScores.reduce((a, b) => a + b, 0) / (paramScores.length || 1)).toFixed(2);
    const withFindings = paramScores.filter(s => s > 0).length;
    const row = ws2.getRow(avgTableStart + 1 + pi);
    [param, avg, fisker.length ? `${Math.round(withFindings / fisker.length * 100)}%` : '0%'].forEach((v, ci) => {
      const cell = row.getCell(ci + 1);
      cell.value = v;
      const scoreColor = avg === 0 ? 'FF27AE60' : avg < 0.5 ? 'FF92400E' : avg < 1 ? 'FFD97706' : 'FFE74C3C';
      if (ci === 1) cell.font = { bold: true, color: { argb: scoreColor } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: pi % 2 === 0 ? 'FFFAFCFF' : 'FFE6F4FB' } };
      cell.alignment = { horizontal: ci === 0 ? 'left' : 'center', vertical: 'middle' };
    });
    row.height = 18;
  });

  ws2.columns = [{ width: 26 }, { width: 16 }, { width: 16 }];

  // Embed welfare chart in Velferd-oversikt sheet
  if (chartWelferdPng) {
    const welferdImgId = wb.addImage({ base64: chartWelferdPng.replace(/^data:image\/png;base64,/, ''), extension: 'png' });
    ws2.addImage(welferdImgId, { tl: { col: 0, row: 25 }, br: { col: 12, row: 48 } });
  }

  return wb;
}

http.createServer(async (req, res) => {

  // ── CORS preflight ─────────────────────────────────────────────────────────
  if (req.method === 'OPTIONS') {
    res.writeHead(204, { 'Access-Control-Allow-Origin': '*', 'Access-Control-Allow-Headers': 'Authorization,Content-Type,Accept', 'Access-Control-Allow-Methods': 'GET,POST,OPTIONS' });
    res.end(); return;
  }

  // ── Word (.docx) export ───────────────────────────────────────────────────
  if (req.method === 'POST' && req.url === '/api/docx') {
    try {
      const payload = await parseBody(req);
      const buf = await buildDocx(payload);
      const safe = (payload.title||'dokument').replace(/[^a-zA-Z0-9æøåÆØÅ\s\-_]/g,'').trim();
      res.writeHead(200, {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        'Content-Disposition': `attachment; filename="${safe}.docx"`,
        'Access-Control-Allow-Origin': '*',
      });
      res.end(buf);
    } catch(e) {
      res.writeHead(500, { 'Content-Type':'application/json', 'Access-Control-Allow-Origin':'*' });
      res.end(JSON.stringify({ error: e.message }));
    }
    return;
  }
  if (req.method === 'OPTIONS' && req.url === '/api/docx') {
    res.writeHead(204, { 'Access-Control-Allow-Origin':'*', 'Access-Control-Allow-Headers':'Content-Type', 'Access-Control-Allow-Methods':'POST' });
    res.end(); return;
  }

  // ── Server info (local IP + tunnel URL for mobile QR) ────────────────────
  if (req.url === '/api/info') {
    const ip = getLocalIP();
    let tunnelUrl = null;
    try { tunnelUrl = fs.readFileSync('/tmp/vela_tunnel_url', 'utf8').trim() || null; } catch(e) {}
    res.writeHead(200, { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': '*', 'Cache-Control': 'no-store' });
    res.end(JSON.stringify({ ip, port, tunnelUrl }));
    return;
  }

  // ── Excel export endpoints ─────────────────────────────────────────────────
  if (req.method === 'POST' && req.url === '/excel-vekt') {
    try {
      const payload = await parseBody(req);
      const wb = await buildVektExcel(payload);
      const buf = await wb.xlsx.writeBuffer();
      res.writeHead(200, {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition': `attachment; filename="${['snittvekt',(payload.lok||''),(payload.merd||''),(payload.dato||'')].map(s=>s.replace(/[^a-zA-Z0-9\-]/g,'_')).join('_')}.xlsx"`,
        'Access-Control-Allow-Origin': '*',
      });
      res.end(buf);
    } catch(e) {
      res.writeHead(500, { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': '*' });
      res.end(JSON.stringify({ error: e.message }));
    }
    return;
  }

  if (req.method === 'POST' && req.url === '/excel-individ') {
    try {
      const payload = await parseBody(req);
      const wb = await buildIndividExcel(payload);
      const buf = await wb.xlsx.writeBuffer();
      res.writeHead(200, {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition': `attachment; filename="${['individkontroll',(payload.lok||''),(payload.merd||''),(payload.dato||'')].map(s=>s.replace(/[^a-zA-Z0-9\-]/g,'_')).join('_')}.xlsx"`,
        'Access-Control-Allow-Origin': '*',
      });
      res.end(buf);
    } catch(e) {
      res.writeHead(500, { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': '*' });
      res.end(JSON.stringify({ error: e.message }));
    }
    return;
  }

  if (req.method === 'OPTIONS' && (req.url === '/excel-vekt' || req.url === '/excel-individ')) {
    res.writeHead(204, { 'Access-Control-Allow-Origin':'*', 'Access-Control-Allow-Headers':'Content-Type', 'Access-Control-Allow-Methods':'POST' });
    res.end(); return;
  }

  // ── BW token proxy (/bw-token) — returns server-side cached token ──────────
  if (req.method === 'POST' && req.url === '/bw-token') {
    // Drain request body (not needed since we use server-side creds)
    req.resume();
    await new Promise(r => req.on('end', r));
    const tok = await getServerBWToken();
    if (tok) {
      res.writeHead(200, { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': '*' });
      res.end(JSON.stringify({ access_token: tok, token_type: 'Bearer', expires_in: Math.floor((_tokenExpiry - Date.now()) / 1000) }));
    } else {
      res.writeHead(502, { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': '*' });
      res.end(JSON.stringify({ error: 'Kunne ikke hente BW-token' }));
    }
    return;
  }

  // ── Fiskeridirektoratet proxy (/fiskdir-api?path=...) ─────────────────────
  if (req.url.startsWith('/fiskdir-api')) {
    if (req.method === 'OPTIONS') {
      res.writeHead(204, { 'Access-Control-Allow-Origin':'*', 'Access-Control-Allow-Headers':'Accept,Content-Type', 'Access-Control-Allow-Methods':'GET' });
      res.end(); return;
    }
    const urlObj2 = new URL(req.url, 'http://localhost');
    const fdPath = urlObj2.searchParams.get('path') || '';
    // pub-aqua paths → api.fiskeridir.no, others → register.fiskeridir.no
    const fdHost = fdPath.startsWith('/pub-aqua') ? 'api.fiskeridir.no' : 'register.fiskeridir.no';
    const fdOpts = { hostname: fdHost, path: fdPath, method: 'GET', headers: { 'Accept': 'application/json' } };
    const fdProxy = https.request(fdOpts, (pRes) => {
      let data = '';
      pRes.on('data', c => data += c);
      pRes.on('end', () => {
        res.writeHead(pRes.statusCode, { 'Content-Type': pRes.headers['content-type'] || 'application/json', 'Access-Control-Allow-Origin': '*' });
        res.end(data);
      });
    });
    fdProxy.on('error', e => {
      res.writeHead(502, { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': '*' });
      res.end(JSON.stringify({ error: e.message }));
    });
    fdProxy.end();
    return;
  }

  // ── BW API proxy (/bw-api?path=...) — injects auth automatically ──────────
  if (req.url.startsWith('/bw-api')) {
    if (req.method === 'OPTIONS') {
      res.writeHead(204, { 'Access-Control-Allow-Origin':'*', 'Access-Control-Allow-Headers':'Authorization,Accept,Content-Type', 'Access-Control-Allow-Methods':'GET,POST' });
      res.end(); return;
    }
    const urlObj = new URL(req.url, 'http://localhost');
    const bwPath = urlObj.searchParams.get('path') || '';
    const method = req.method === 'POST' ? 'POST' : 'GET';

    let reqBody = '';
    req.on('data', c => reqBody += c);
    req.on('end', async () => {
      // Always use server-side token
      const token = await getServerBWToken();
      const hdrs = {
        'Accept': 'application/json',
        'Authorization': token ? `Bearer ${token}` : '',
      };
      if (method === 'POST' && reqBody) {
        hdrs['Content-Type'] = req.headers['content-type'] || 'application/json';
        hdrs['Content-Length'] = Buffer.byteLength(reqBody);
      }
      const options = { hostname: 'www.barentswatch.no', path: bwPath, method, headers: hdrs };
      const proxy = https.request(options, (pRes) => {
        let data = '';
        pRes.on('data', c => data += c);
        pRes.on('end', () => {
          res.writeHead(pRes.statusCode, { 'Content-Type': pRes.headers['content-type'] || 'application/json', 'Access-Control-Allow-Origin': '*' });
          res.end(data);
        });
      });
      proxy.on('error', e => {
        res.writeHead(502, { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': '*' });
        res.end(JSON.stringify({ error: e.message }));
      });
      if (method === 'POST' && reqBody) proxy.write(reqBody);
      proxy.end();
    });
    return;
  }

  // ── Auth: login (/login POST) ──────────────────────────────────────────────
  if (req.method === 'POST' && req.url === '/login') {
    let body = ''; req.on('data', c => body += c);
    req.on('end', () => {
      try {
        const { name, pin } = JSON.parse(body);
        if (!name || !pin) { res.writeHead(400,{'Access-Control-Allow-Origin':'*'}); res.end(JSON.stringify({error:'Mangler felt'})); return; }
        const users = loadUsers();
        const user = users.find(u => u.name.toLowerCase() === name.toLowerCase());
        // Support both hashed password and legacy plain PIN
        const passwordOk = user && (
          (user.password && verifyPassword(String(pin), user.password)) ||
          (!user.password && user.pin && user.pin === String(pin))
        );
        if (!passwordOk) {
          res.writeHead(401, { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': '*' });
          res.end(JSON.stringify({ error: 'Feil navn eller passord' })); return;
        }
        // Upgrade plain PIN to hashed password on first login
        if (!user.password) {
          user.password = hashPassword(String(pin));
          delete user.pin;
          saveUsers(users);
        }
        const token = makeToken();
        _sessions[token] = { userId: user.id, name: user.name, role: user.role || 'biolog', expires: Date.now() + 86400000 };
        res.writeHead(200, { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': '*' });
        res.end(JSON.stringify({ token, name: user.name, role: user.role || 'biolog' }));
      } catch { res.writeHead(400, {'Access-Control-Allow-Origin':'*'}); res.end('{}'); }
    }); return;
  }

  // ── Auth: register (/register POST) ───────────────────────────────────────
  if (req.method === 'POST' && req.url === '/register') {
    let body = ''; req.on('data', c => body += c);
    req.on('end', () => {
      try {
        const { name, pin, invite, role } = JSON.parse(body);
        if (invite !== 'IsacVela') {
          res.writeHead(403, { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': '*' });
          res.end(JSON.stringify({ error: 'Feil invitasjonskode' })); return;
        }
        if (!pin || String(pin).length < 6) {
          res.writeHead(400, { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': '*' });
          res.end(JSON.stringify({ error: 'Passord må ha minst 6 tegn' })); return;
        }
        const validRole = (role === 'driftsteknikker') ? 'driftsteknikker' : 'biolog';
        const users = loadUsers();
        if (users.find(u => u.name.toLowerCase() === name.toLowerCase())) {
          res.writeHead(409, { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': '*' });
          res.end(JSON.stringify({ error: 'Brukernavnet er tatt' })); return;
        }
        const newUser = { id: makeToken().slice(0,8), name: name.trim(), password: hashPassword(String(pin)), role: validRole, created: new Date().toISOString() };
        users.push(newUser);
        saveUsers(users);
        const token = makeToken();
        _sessions[token] = { userId: newUser.id, name: newUser.name, role: newUser.role, expires: Date.now() + 86400000 };
        res.writeHead(200, { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': '*' });
        res.end(JSON.stringify({ token, name: newUser.name, role: newUser.role }));
      } catch { res.writeHead(400, {'Access-Control-Allow-Origin':'*'}); res.end('{}'); }
    }); return;
  }

  // ── Auth: check session (/check-session GET) ───────────────────────────────
  if (req.method === 'GET' && req.url === '/check-session') {
    req.resume();
    const auth = req.headers['authorization'] || '';
    const token = auth.replace('Bearer ', '').trim();
    const s = _sessions[token];
    if (s && Date.now() < s.expires) {
      res.writeHead(200, { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': '*' });
      res.end(JSON.stringify({ ok: true, name: s.name, role: s.role || 'biolog' }));
    } else {
      res.writeHead(401, { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': '*' });
      res.end(JSON.stringify({ ok: false }));
    }
    return;
  }

  // ── Whisper transcription (/transcribe POST) ───────────────────────────────
  if (req.method === 'POST' && req.url === '/transcribe') {
    if (!OPENAI_API_KEY) {
      res.writeHead(503, { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': '*' });
      res.end(JSON.stringify({ error: 'no_key', message: 'OpenAI API-nøkkel ikke konfigurert' })); return;
    }
    const chunks = [];
    req.on('data', c => chunks.push(c));
    req.on('end', () => {
      const body = Buffer.concat(chunks);
      const ct = req.headers['content-type'] || '';
      const boundary = ct.split('boundary=')[1];
      if (!boundary) { res.writeHead(400, {'Access-Control-Allow-Origin':'*'}); res.end('{}'); return; }

      // Forward to Groq Whisper (gratis), fallback to OpenAI
      const whisperKey  = GROQ_API_KEY || OPENAI_API_KEY;
      const whisperHost = GROQ_API_KEY ? 'api.groq.com' : 'api.openai.com';
      const whisperPath = GROQ_API_KEY ? '/openai/v1/audio/transcriptions' : '/v1/audio/transcriptions';
      const options = {
        hostname: whisperHost,
        path: whisperPath,
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${whisperKey}`,
          'Content-Type': ct,
          'Content-Length': body.length,
        }
      };
      const oaiReq = https.request(options, oaiRes => {
        let data = '';
        oaiRes.on('data', c => data += c);
        oaiRes.on('end', () => {
          res.writeHead(oaiRes.statusCode, { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': '*' });
          res.end(data);
        });
      });
      oaiReq.on('error', e => {
        res.writeHead(502, { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': '*' });
        res.end(JSON.stringify({ error: e.message }));
      });
      oaiReq.write(body);
      oaiReq.end();
    }); return;
  }

  // ── AI proxy (/claude POST) — accepts Anthropic format, runs on Groq (gratis) ─────
  // Maps claude-opus/sonnet → llama-3.3-70b-versatile, claude-haiku → llama-3.1-8b-instant
  if (req.method === 'POST' && req.url === '/claude') {
    const chunks = []; req.on('data', c => chunks.push(c));
    req.on('end', () => {
      try {
        const incoming = JSON.parse(Buffer.concat(chunks).toString());

        // Model mapping: best model for main analysis, mini for helpers
        const modelHint = (incoming.model || '').toLowerCase();
        const oaiModel = modelHint.includes('haiku') ? 'llama-3.1-8b-instant' : 'llama-3.3-70b-versatile';

        // Convert Anthropic format → OpenAI format
        const messages = [];
        if (incoming.system) messages.push({ role: 'system', content: incoming.system });
        (incoming.messages || []).forEach(m => {
          const content = Array.isArray(m.content)
            ? m.content.map(c => c.text || c).join('') : m.content;
          messages.push({ role: m.role, content });
        });

        const oaiBody = JSON.stringify({
          model: oaiModel,
          max_tokens: incoming.max_tokens || 1024,
          messages,
          temperature: incoming.temperature ?? 0.3,
        });

        const aiKey  = GROQ_API_KEY || OPENAI_API_KEY;
        const aiHost = GROQ_API_KEY ? 'api.groq.com' : 'api.openai.com';
        const aiPath = GROQ_API_KEY ? '/openai/v1/chat/completions' : '/v1/chat/completions';
        const options = {
          hostname: aiHost,
          path: aiPath,
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
            'Authorization': `Bearer ${aiKey}`,
            'Content-Length': Buffer.byteLength(oaiBody),
          }
        };

        const oaiReq = https.request(options, oaiRes => {
          let data = '';
          oaiRes.on('data', c => data += c);
          oaiRes.on('end', () => {
            try {
              const oaiResp = JSON.parse(data);
              // Convert OpenAI response → Anthropic format so client code works unchanged
              if (oaiResp.choices?.[0]?.message) {
                const text = oaiResp.choices[0].message.content || '';
                const anthropicResp = {
                  id: oaiResp.id,
                  type: 'message',
                  role: 'assistant',
                  content: [{ type: 'text', text }],
                  model: oaiModel,
                  stop_reason: 'end_turn',
                  usage: {
                    input_tokens: oaiResp.usage?.prompt_tokens || 0,
                    output_tokens: oaiResp.usage?.completion_tokens || 0,
                  }
                };
                res.writeHead(200, { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': '*' });
                res.end(JSON.stringify(anthropicResp));
              } else {
                res.writeHead(oaiRes.statusCode, { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': '*' });
                res.end(data);
              }
            } catch(e) {
              res.writeHead(502, { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': '*' });
              res.end(JSON.stringify({ error: { message: e.message } }));
            }
          });
        });
        oaiReq.on('error', e => {
          res.writeHead(502, { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': '*' });
          res.end(JSON.stringify({ error: { message: e.message } }));
        });
        oaiReq.write(oaiBody);
        oaiReq.end();
      } catch(e) {
        res.writeHead(400, { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': '*' });
        res.end(JSON.stringify({ error: { message: e.message } }));
      }
    });
    return;
  }

  // ── Static file serving ────────────────────────────────────────────────────
  let p = path.join(dir, req.url === '/' ? '/landing.html' : req.url.split('?')[0]);
  fs.readFile(p, (err, data) => {
    if (err) { res.writeHead(404); res.end('Not found'); return; }
    res.writeHead(200, { 'Content-Type': mime[path.extname(p)] || 'text/plain', 'Cache-Control': 'no-store' });
    res.end(data);
  });

}).listen(port, () => console.log('Vela server running on port ' + port));
