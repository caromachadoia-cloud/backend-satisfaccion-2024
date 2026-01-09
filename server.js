const express = require('express');
const multer = require('multer');
const ExcelJS = require('exceljs');
const cors = require('cors');

const app = express();
const PORT = process.env.PORT || 3000;

app.use(cors({ origin: '*', methods: ['GET', 'POST'], allowedHeaders: ['Content-Type'] }));
app.use(express.json());

const storage = multer.memoryStorage();
const upload = multer({ storage });

const TOP_WORDS = 50;
const TOP_COMMENTS = 6;

const STOPWORDS = new Set([
  'de','la','que','el','en','y','a','los','del','se','las','por','un','para','con','no','una','su','al','lo','como','mas','pero','sus','le','ya','o',
  'este','ha','me','si','sin','sobre','muy','cuando','tambien','hasta','hay','donde','quien','desde','todo','nos','durante','uno','ni','contra','ese','eso','mi',
  'e','son','fue','gracias','hola','buen','dia','tarde','noche','lugar','servicio','atencion','excelente','buena','mala','regular','bien','mal','hace','falta',
  'mucha','mucho','esta','estos','estaba','fueron','tener','tiene','todo','ok','oka','dale','genial','super','súper','re'
]);

function stripAccents(s) {
  return String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '');
}

function onlyUsefulText(s) {
  let t = String(s || '').trim();
  if (!t) return '';
  t = t.replace(/https?:\/\/\S+/gi, ' ');
  t = t.replace(/[@#]\S+/g, ' ');
  t = t.replace(/[\p{Extended_Pictographic}\p{Emoji_Presentation}]/gu, ' ');
  t = t.replace(/[^\p{L}\p{N}\s]/gu, ' ');
  t = t.replace(/\s+/g, ' ').trim();
  return t;
}

function isSpammyRepeated(s) {
  const t = stripAccents(String(s || '').toLowerCase()).replace(/\s+/g, '');
  if (!t) return true;
  if (/^(.)\1{3,}$/.test(t)) return true;
  if (/(.)\1{6,}/.test(t)) return true;
  return false;
}

function tokenize(text) {
  const clean = onlyUsefulText(text);
  if (!clean || clean.length < 5) return [];
  const base = stripAccents(clean.toLowerCase());
  const words = base.match(/[a-zñ]+/g) || [];
  const out = [];
  for (const w of words) {
    if (w.length < 3) continue;
    if (STOPWORDS.has(w)) continue;
    if (isSpammyRepeated(w)) continue;
    out.push(w);
  }
  return out;
}

function commentQualityScore(text) {
  const t = String(text || '').trim();
  if (!t) return -999;
  if (t.length < 12) return -999;
  if (isSpammyRepeated(t)) return -999;

  const tokens = tokenize(t);
  const uniq = new Set(tokens);
  const uniqCount = uniq.size;

  if (uniqCount < 3) return -999;

  let score = 0;
  score += Math.min(60, t.length);
  score += uniqCount * 6;
  score -= Math.max(0, tokens.length - uniqCount) * 2;

  return score;
}

function freqTop(wordsArr, topN) {
  const map = new Map();
  for (const w of wordsArr) map.set(w, (map.get(w) || 0) + 1);
  return Array.from(map.entries()).sort((a, b) => b[1] - a[1]).slice(0, topN);
}

function pickTopComments(arr, topN) {
  const seen = new Set();
  const scored = [];
  for (const c of arr) {
    const raw = String(c?.texto || '').trim();
    const clean = onlyUsefulText(raw);
    if (!clean) continue;

    const key = stripAccents(clean.toLowerCase()).replace(/\s+/g, ' ').trim();
    if (seen.has(key)) continue;
    seen.add(key);

    const score = commentQualityScore(clean);
    if (score < 0) continue;

    scored.push({ texto: raw, date: c.date, score });
  }

  scored.sort((a, b) => b.score - a.score);

  return scored.slice(0, topN).map(c => ({
    texto: c.texto,
    meta: `${c.date.getDate()}/${c.date.getMonth() + 1} ${String(c.date.getHours()).padStart(2,'0')}:00hs`
  }));
}

app.post('/procesar-anual', upload.single('archivoExcel'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ success: false, message: 'Falta archivo' });

    let datosManuales = {};
    try { datosManuales = JSON.parse(req.body.datosManuales || '{}'); } catch { datosManuales = {}; }

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(req.file.buffer);
    const worksheet = workbook.worksheets[0];

    const colMap = {};
    worksheet.getRow(1).eachCell((cell, colNumber) => {
      const val = String(cell.value || '').toLowerCase().trim();
      if (val.includes('fecha')) colMap.fecha = colNumber;
      if (val.includes('hora')) colMap.hora = colNumber;
      if (val.includes('sector')) colMap.sector = colNumber;
      if (val.includes('ubicacion')) colMap.ubicacion = colNumber;
      if (val === 'calificacion' || val === 'calificación') colMap.val = colNumber;
      if (val.includes('comentario')) colMap.comentario = colNumber;
    });

    if (!colMap.fecha || !colMap.val || !colMap.sector) {
      return res.status(400).json({ success: false, message: 'No se detectaron columnas clave (fecha/calificación/sector).' });
    }

    const sectorsData = {};

    worksheet.eachRow((row, rowNum) => {
      if (rowNum === 1) return;

      const ratingRaw = row.getCell(colMap.val).value;
      const rating = parseInt(String(ratingRaw || '').trim(), 10);
      if (![1,2,3,4].includes(rating)) return;

      const dateVal = row.getCell(colMap.fecha).value;
      let date = (dateVal instanceof Date) ? dateVal : new Date(dateVal);
      if (isNaN(date.getTime())) return;

      const sectorName = String(row.getCell(colMap.sector).value || 'General').trim() || 'General';
      const ubicName = String(row.getCell(colMap.ubicacion).value || 'General').trim() || 'General';
      const rawComment = String(row.getCell(colMap.comentario).value || '').trim();

      if (!sectorsData[sectorName]) {
        sectorsData[sectorName] = {
          meses: Array.from({ length: 12 }, () => ({ mp:0, p:0, n:0, mn:0, total:0 })),
          ubicaciones: {},
          palabrasPos: [],
          palabrasNeg: [],
          horasNeg: Array(24).fill(0),
          comsPos: [],
          comsNeg: []
        };
      }

      const s = sectorsData[sectorName];
      if (!s.ubicaciones[ubicName]) {
        s.ubicaciones[ubicName] = { mp:0, p:0, n:0, mn:0, total:0, horasNeg: Array(24).fill(0) };
      }

      const mesIdx = date.getMonth();
      const statsMes = s.meses[mesIdx];
      const statsUbic = s.ubicaciones[ubicName];

      statsMes.total++;
      statsUbic.total++;

      let hour = 12;
      if (colMap.hora) {
        const hVal = row.getCell(colMap.hora).value;
        if (hVal instanceof Date) hour = hVal.getHours();
        else {
          const p = String(hVal || '').split(':')[0];
          const ph = parseInt(p, 10);
          if (!isNaN(ph)) hour = Math.min(23, Math.max(0, ph));
        }
      }

      if (rating === 4) {
        statsMes.mp++; statsUbic.mp++;
        if (rawComment.length > 8) {
          s.palabrasPos.push(...tokenize(rawComment));
          s.comsPos.push({ texto: rawComment, date });
        }
      } else if (rating === 3) {
        statsMes.p++; statsUbic.p++;
        if (rawComment.length > 8) {
          s.palabrasPos.push(...tokenize(rawComment));
          s.comsPos.push({ texto: rawComment, date });
        }
      } else if (rating === 2) {
        statsMes.n++; statsUbic.n++;
        s.horasNeg[hour]++; statsUbic.horasNeg[hour]++;
        if (rawComment.length > 8) {
          s.palabrasNeg.push(...tokenize(rawComment));
          s.comsNeg.push({ texto: rawComment, date });
        }
      } else if (rating === 1) {
        statsMes.mn++; statsUbic.mn++;
        s.horasNeg[hour]++; statsUbic.horasNeg[hour]++;
        if (rawComment.length > 8) {
          s.palabrasNeg.push(...tokenize(rawComment));
          s.comsNeg.push({ texto: rawComment, date });
        }
      }
    });

    const resultado = Object.entries(sectorsData).map(([nombre, data]) => {
      ['enero','febrero'].forEach((mes, i) => {
        if (datosManuales[mes]) {
          const total = datosManuales[mes].total || 0;
          const mp = datosManuales[mes].muy_positivas || 0;
          data.meses[i].total += total;
          data.meses[i].mp += mp;
        }
      });

      const mesesFinal = data.meses.map((m, i) => {
        const pos = m.mp + m.p;
        const neg = m.n + m.mn;
        const sat = m.total > 0 ? (((pos - neg) / m.total) * 100) : 0;
        return {
          nombre: ['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'][i],
          sat: Number(sat.toFixed(1)),
          total: m.total
        };
      });

      const ranking = Object.entries(data.ubicaciones).map(([key, u]) => {
        let maxH = 0, hC = 0;
        u.horasNeg.forEach((c, h) => { if (c > maxH) { maxH = c; hC = h; } });

        const pos = u.mp + u.p;
        const neg = u.n + u.mn;
        const sat = u.total > 0 ? (((pos - neg) / u.total) * 100) : 0;

        return { nombre: key, total: u.total, sat: N
