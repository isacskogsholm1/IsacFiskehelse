const http = require('http');
const https = require('https');
const fs = require('fs');
const path = require('path');
const os = require('os');

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

// ── OpenAI Whisper ─────────────────────────────────────────────────────────
const OPENAI_API_KEY = process.env.OPENAI_API_KEY || '';

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
  const { lok, merd, dato, merknad, fisker, stats, chartPng } = payload;
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Vela'; wb.created = new Date();

  // ── Sheet 1: Dashboard ─────────────────────────────────────────────────────
  const ws = wb.addWorksheet('Snittvekt', { views: [{ showGridLines: true }] });

  // Header row
  ws.mergeCells('A1:F1');
  const titleCell = ws.getCell('A1');
  titleCell.value = `Snittvekt — ${lok} — Merd ${merd} — ${dato}`;
  titleCell.font = { size: 14, bold: true, color: { argb: 'FFFFFFFF' } };
  titleCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0B72B5' } };
  titleCell.alignment = { horizontal: 'left', vertical: 'middle' };
  ws.getRow(1).height = 28;

  // Column headers
  const headers = ['Fisk nr.', 'Vekt (g)', 'Avvik (g)', 'Avvik (%)', 'Status'];
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
  fisker.forEach((v, i) => {
    const row = ws.getRow(i + 3);
    const avvik = +(v - stats.mean).toFixed(1);
    const pct   = +((v - stats.mean) / stats.mean * 100).toFixed(1);
    const status = v > stats.mean * 1.25 ? '▲ Over' : v < stats.mean * 0.75 ? '▼ Under' : '● OK';
    const fillColor = i % 2 === 0 ? 'FFFAFCFF' : 'FFE6F4FB';
    const statusColor = v > stats.mean * 1.25 ? 'FFE74C3C' : v < stats.mean * 0.75 ? 'FFF39C12' : 'FF27AE60';

    [i + 1, v, avvik, pct, status].forEach((val, ci) => {
      const cell = row.getCell(ci + 1);
      cell.value = val;
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: fillColor } };
      cell.alignment = { horizontal: ci === 0 ? 'center' : ci === 4 ? 'center' : 'right', vertical: 'middle' };
      if (ci === 4) cell.font = { bold: true, color: { argb: statusColor } };
      if ((ci === 2 || ci === 3) && avvik < 0) cell.font = { color: { argb: 'FFE74C3C' } };
      cell.border = { bottom: { style: 'hair', color: { argb: 'FFDDEEEE' } } };
    });
    row.height = 18;
  });

  // Snitt row
  const sRow = ws.getRow(fisker.length + 3);
  ['SNITT', stats.mean.toFixed(1), '—', '—', ''].forEach((v, i) => {
    const cell = sRow.getCell(i + 1);
    cell.value = i === 1 ? +v : v;
    cell.font = { bold: true, color: { argb: 'FF054370' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD0E8F8' } };
    cell.alignment = { horizontal: i === 0 ? 'left' : 'right', vertical: 'middle' };
  });

  // Col widths
  ws.columns = [
    { width: 10 }, { width: 12 }, { width: 14 }, { width: 12 }, { width: 14 }
  ];

  // ── Sheet 2: Statistikk ───────────────────────────────────────────────────
  const ws2 = wb.addWorksheet('Statistikk');
  const _fmt = (v) => (v != null && !isNaN(v)) ? +parseFloat(v).toFixed(1) : '—';
  [
    ['Parameter', 'Verdi'],
    ['Lokalitet', lok], ['Merd', merd], ['Dato', dato],
    ['Antall fisk (n)', stats.n || fisker.length],
    ['Gjennomsnitt (g)', _fmt(stats.mean)],
    ['Median (g)',        _fmt(stats.median)],
    ['Standardavvik (g)', _fmt(stats.sd)],
    ['CV%', stats.cv != null ? +stats.cv : '—'],
    ['Min (g)', stats.min ?? '—'], ['Maks (g)', stats.max ?? '—'],
    ['Merknad', merknad || ''],
  ].forEach((row, ri) => {
    row.forEach((val, ci) => {
      const cell = ws2.getCell(ri + 1, ci + 1);
      cell.value = val;
      if (ri === 0) {
        cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF054370' } };
      } else {
        if (ci === 0) cell.font = { bold: true, color: { argb: 'FF054370' } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: ri % 2 === 0 ? 'FFFAFCFF' : 'FFE6F4FB' } };
      }
      cell.alignment = { vertical: 'middle' };
    });
  });
  ws2.columns = [{ width: 26 }, { width: 22 }];

  // ── Embed chart PNG if provided ───────────────────────────────────────────
  if (chartPng) {
    const imgId = wb.addImage({ base64: chartPng.replace(/^data:image\/png;base64,/, ''), extension: 'png' });
    ws.addImage(imgId, { tl: { col: 6, row: 1 }, br: { col: 14, row: Math.min(fisker.length + 4, 22) } });
  }

  return wb;
}

async function buildIndividExcel(payload) {
  const { lok, merd, dato, fisker, avgV, avgL, avgK, scoreParams, chartVektPng, chartKPng, chartWelferdPng } = payload;
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Vela'; wb.created = new Date();

  // ── Sheet 1: Individdata ──────────────────────────────────────────────────
  const ws = wb.addWorksheet('Individkontroll');
  ws.mergeCells('A1:G1');
  const tc = ws.getCell('A1');
  tc.value = `Individkontroll — ${lok} — Merd ${merd} — ${dato}`;
  tc.font = { size: 14, bold: true, color: { argb: 'FFFFFFFF' } };
  tc.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0B72B5' } };
  tc.alignment = { horizontal: 'left', vertical: 'middle' };
  ws.getRow(1).height = 28;

  const heads = ['Fisk nr.', 'Vekt (g)', 'Lengde (cm)', 'K-faktor', 'Kjønn', 'Velferdssum'];
  heads.forEach((h, i) => {
    const cell = ws.getCell(2, i + 1);
    cell.value = h;
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF054370' } };
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
  });
  ws.getRow(2).height = 22;

  fisker.forEach((f, i) => {
    const kColor = f.k == null ? 'FF999999' : f.k >= 1.0 ? 'FF27AE60' : f.k >= 0.8 ? 'FFF39C12' : 'FFE74C3C';
    const sColor = f.score === 0 ? 'FF27AE60' : f.score <= 3 ? 'FFF39C12' : 'FFE74C3C';
    const fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: i % 2 === 0 ? 'FFFAFCFF' : 'FFE6F4FB' } };
    const row = ws.getRow(i + 3);
    [i+1, f.vekt, f.lengde, f.k, f.kjonn, f.score].forEach((val, ci) => {
      const cell = row.getCell(ci + 1);
      cell.value = val;
      cell.fill = fill;
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
      if (ci === 3) cell.font = { bold: true, color: { argb: kColor } };
      if (ci === 5) cell.font = { bold: true, color: { argb: sColor } };
    });
    row.height = 18;
  });

  const sRow = ws.getRow(fisker.length + 3);
  ['SNITT', +avgV, +avgL, +avgK, '', ''].forEach((v, i) => {
    const cell = sRow.getCell(i + 1);
    cell.value = v;
    cell.font = { bold: true, color: { argb: 'FF054370' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD0E8F8' } };
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
  });
  ws.columns = [{ width:10 },{ width:12 },{ width:14 },{ width:12 },{ width:10 },{ width:14 }];

  // Embed charts
  if (chartVektPng) {
    const id = wb.addImage({ base64: chartVektPng.replace(/^data:image\/png;base64,/, ''), extension: 'png' });
    ws.addImage(id, { tl: { col: 7, row: 1 }, br: { col: 16, row: Math.min(fisker.length/2 + 4, 18) } });
  }
  if (chartKPng) {
    const id2 = wb.addImage({ base64: chartKPng.replace(/^data:image\/png;base64,/, ''), extension: 'png' });
    const offset = Math.min(fisker.length/2 + 5, 19);
    ws.addImage(id2, { tl: { col: 7, row: offset }, br: { col: 16, row: offset + 14 } });
  }

  // ── Sheet 2: Velferd ──────────────────────────────────────────────────────
  const ws2 = wb.addWorksheet('Velferd');
  ws2.mergeCells('A1:' + String.fromCharCode(65 + scoreParams.length + 1) + '1');
  const vc = ws2.getCell('A1');
  vc.value = 'Velferdsskår per fisk — 0=ingen funn, 1=mild, 2=moderat, 3=alvorlig';
  vc.font = { bold: true, color: { argb: 'FFFFFFFF' } };
  vc.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF054370' } };

  ['Fisk', ...scoreParams, 'Sum'].forEach((h, i) => {
    const cell = ws2.getCell(2, i + 1);
    cell.value = h;
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0B72B5' } };
    cell.alignment = { horizontal: 'center' };
  });

  fisker.forEach((f, ri) => {
    const scores = f.scores || [];
    [ri + 1, ...scores, f.score].forEach((v, ci) => {
      const cell = ws2.getCell(ri + 3, ci + 1);
      cell.value = v;
      const sc = ci === 0 ? null : v;
      if (sc !== null) {
        const bg = sc === 0 ? 'FF27AE6020' : sc === 1 ? 'FFF39C1220' : sc === 2 ? 'FFE67E2220' : 'FFE74C3C20';
        const fg = sc === 0 ? 'FF1E8449' : sc === 1 ? 'FFE67E22' : sc === 2 ? 'FFD35400' : 'FFC0392B';
        cell.font = { bold: true, color: { argb: fg } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } };
      }
      cell.alignment = { horizontal: 'center' };
    });
  });
  ws2.columns = [{ width: 10 }, ...scoreParams.map(() => ({ width: 14 })), { width: 10 }];

  // ── Velferd summary stats below fish data ──────────────────────────────────
  const summaryStartRow = fisker.length + 4;
  const summaryLabels = [
    ['', '', ''],
    ['SAMMENDRAG', '', ''],
    ['Kategori', 'Antall fisk', '% av total'],
  ];
  const scoreSums = fisker.map(f => (f.scores || []).reduce((a,b) => a+b, 0));
  const categories = [
    ['OK  (sum = 0)',       scoreSums.filter(s => s === 0).length],
    ['Mild (sum 1–3)',      scoreSums.filter(s => s >= 1 && s <= 3).length],
    ['Moderat (sum 4–6)',   scoreSums.filter(s => s >= 4 && s <= 6).length],
    ['Alvorlig (sum > 6)', scoreSums.filter(s => s > 6).length],
  ];
  const catColors = ['FF27AE60','FFF39C12','FFE67E22','FFE74C3C'];
  summaryLabels.concat(categories.map(([l,n]) => [l, n, fisker.length ? +(n/fisker.length*100).toFixed(0)+'%' : '0%']))
    .forEach((row, ri) => {
      row.forEach((val, ci) => {
        const cell = ws2.getCell(summaryStartRow + ri, ci + 1);
        cell.value = val;
        if (ri === 1) {
          cell.font = { bold: true, size: 12, color: { argb: 'FF054370' } };
        } else if (ri === 2) {
          cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF054370' } };
          cell.alignment = { horizontal: 'center' };
        } else if (ri >= 3) {
          const catIdx = ri - 3;
          const isCount = ci === 1;
          if (ci === 0) cell.font = { bold: true };
          if (isCount && categories[catIdx]) {
            cell.font = { bold: true, color: { argb: catColors[catIdx] } };
          }
          cell.alignment = { horizontal: ci === 0 ? 'left' : 'center' };
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: ri % 2 === 0 ? 'FFFAFCFF' : 'FFE6F4FB' } };
        }
      });
    });

  // ── Average score per parameter (for chart context) ────────────────────────
  const avgStartRow = summaryStartRow + categories.length + 5;
  ['Parameter', 'Gj.snitt score', 'Andel med funn'].forEach((h, ci) => {
    const cell = ws2.getCell(avgStartRow, ci + 1);
    cell.value = h;
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0B72B5' } };
    cell.alignment = { horizontal: 'center' };
  });
  scoreParams.forEach((param, pi) => {
    const paramScores = fisker.map(f => (f.scores || [])[pi] ?? 0);
    const avg = +(paramScores.reduce((a,b)=>a+b,0) / (paramScores.length||1)).toFixed(2);
    const withFindings = paramScores.filter(s=>s>0).length;
    const row = ws2.getRow(avgStartRow + 1 + pi);
    [param, avg, fisker.length ? `${Math.round(withFindings/fisker.length*100)}%` : '0%'].forEach((v, ci) => {
      const cell = row.getCell(ci + 1);
      cell.value = v;
      const scoreColor = avg === 0 ? 'FF27AE60' : avg < 0.5 ? 'FF92400E' : avg < 1 ? 'FFD97706' : 'FFE74C3C';
      if (ci === 1) cell.font = { bold: true, color: { argb: scoreColor } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: pi % 2 === 0 ? 'FFFAFCFF' : 'FFE6F4FB' } };
      cell.alignment = { horizontal: ci === 0 ? 'left' : 'center' };
    });
    row.height = 18;
  });

  // ── Embed welfare chart in Velferd sheet ───────────────────────────────────
  if (chartWelferdPng) {
    const chartCol = scoreParams.length + 3;
    const welferdImgId = wb.addImage({ base64: chartWelferdPng.replace(/^data:image\/png;base64,/, ''), extension: 'png' });
    ws2.addImage(welferdImgId, { tl: { col: chartCol, row: 1 }, br: { col: chartCol + 10, row: Math.min(fisker.length + 6, 24) } });
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
        'Content-Disposition': `attachment; filename="snittvekt_${(payload.lok||'').replace(/[^a-zA-Z0-9]/g,'_')}_merd${payload.merd}_${(payload.dato||'').replace(/\./g,'-')}.xlsx"`,
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
        'Content-Disposition': `attachment; filename="individkontroll_${(payload.lok||'').replace(/[^a-zA-Z0-9]/g,'_')}_merd${payload.merd}_${(payload.dato||'').replace(/\./g,'-')}.xlsx"`,
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
        const users = loadUsers();
        const user = users.find(u => u.name.toLowerCase() === name.toLowerCase() && u.pin === String(pin));
        if (!user) {
          res.writeHead(401, { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': '*' });
          res.end(JSON.stringify({ error: 'Feil navn eller PIN' })); return;
        }
        const token = makeToken();
        _sessions[token] = { userId: user.id, name: user.name, expires: Date.now() + 86400000 };
        res.writeHead(200, { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': '*' });
        res.end(JSON.stringify({ token, name: user.name, role: user.role || 'user' }));
      } catch { res.writeHead(400, {'Access-Control-Allow-Origin':'*'}); res.end('{}'); }
    }); return;
  }

  // ── Auth: register (/register POST) ───────────────────────────────────────
  if (req.method === 'POST' && req.url === '/register') {
    let body = ''; req.on('data', c => body += c);
    req.on('end', () => {
      try {
        const { name, pin, invite } = JSON.parse(body);
        if (invite !== 'VELA2025') {
          res.writeHead(403, { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': '*' });
          res.end(JSON.stringify({ error: 'Feil invitasjonskode' })); return;
        }
        const users = loadUsers();
        if (users.find(u => u.name.toLowerCase() === name.toLowerCase())) {
          res.writeHead(409, { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': '*' });
          res.end(JSON.stringify({ error: 'Brukernavnet er tatt' })); return;
        }
        const newUser = { id: makeToken().slice(0,8), name: name.trim(), pin: String(pin), role: 'user', created: new Date().toISOString() };
        users.push(newUser);
        saveUsers(users);
        const token = makeToken();
        _sessions[token] = { userId: newUser.id, name: newUser.name, expires: Date.now() + 86400000 };
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
      res.end(JSON.stringify({ ok: true, name: s.name }));
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

      // Forward multipart directly to OpenAI Whisper
      const options = {
        hostname: 'api.openai.com',
        path: '/v1/audio/transcriptions',
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${OPENAI_API_KEY}`,
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

  // ── Static file serving ────────────────────────────────────────────────────
  let p = path.join(dir, req.url === '/' ? '/landing.html' : req.url.split('?')[0]);
  fs.readFile(p, (err, data) => {
    if (err) { res.writeHead(404); res.end('Not found'); return; }
    res.writeHead(200, { 'Content-Type': mime[path.extname(p)] || 'text/plain', 'Cache-Control': 'no-store' });
    res.end(data);
  });

}).listen(port, () => console.log('Vela server running on port ' + port));
