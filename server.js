const express = require('express');
const multer = require('multer');
const ExcelJS = require('exceljs');
const cors = require('cors');

const app = express();
const PORT = process.env.PORT || 10000;

app.use(cors());
app.use(express.json());

const storage = multer.memoryStorage();
const upload = multer({ storage: storage });

// DICCIONARIO PARA ANÁLISIS INTELIGENTE POR SECTOR
const CONTEXTO_SECTORES = {
    "atencion": ["atencion", "trato", "amable", "ayudo", "atenta", "chica", "chico", "resolvio", "espera","tarjeta","fun"],
    "gastro": ["comida", "mozo", "frio", "caliente", "rico", "bebida", "mesa", "pedido", "tardo"],
    "cajas": ["pago", "cobro", "fila", "rapido", "dinero", "efectivo", "tarjeta", "atencion","espera","cajero"],
    "traslados": ["chofer", "auto", "camioneta", "viaje", "llego", "tarde", "limpio", "conduccion", "carrito"],
    "baños": ["lugar", "instalaciones", "baño", "limpieza", "seguridad","olor","jabon","papel"]
};

function esComentarioInteligente(texto, sector) {
    if (!texto || texto.length < 20) return false;
    const limpio = texto.toLowerCase();
    if (!/[a-z]{4,}/.test(limpio)) return false; // Filtra basura

    const sectorKey = Object.keys(CONTEXTO_SECTORES).find(k => sector.toLowerCase().includes(k)) || "general";
    const palabrasClave = CONTEXTO_SECTORES[sectorKey];
    
    // Si es un sector específico, validamos que hable de algo relacionado
    if (sectorKey !== "baños") {
        const palabrasProhibidas = ["maquina", "suerte", "paga"].filter(p => !palabrasClave.includes(p));
        if (palabrasProhibidas.some(p => limpio.includes(p))) return false;
    }
    return true;
}

const STOPWORDS = ['g','l','de', 'la', 'que', 'el', 'en', 'y', 'a', 'los', 'del', 'se', 'las', 'por', 'un', 'para', 'con', 'no', 'una', 'su', 'al', 'lo', 'como', 'más', 'pero', 'sus', 'le', 'ya', 'o', 'este', 'ha', 'me', 'si', 'sin', 'sobre', 'muy', 'cuando', 'también', 'hasta', 'hay', 'donde', 'quien', 'desde', 'todo', 'nos', 'uno', 'ni', 'contra', 'ese', 'eso', 'mi', 'qué', 'e', 'son', 'fue', 'gracias', 'hola', 'lugar', 'atencion', 'servicio'];

function getWords(text) {
    return text.toLowerCase().replace(/[.,/#!$%^&*;:{}=\-_`~()]/g, " ").match(/[a-záéíóúñü]+/g)
        ?.filter(word => !STOPWORDS.includes(word) && word.length > 3) || [];
}

app.post('/procesar-anual', upload.single('archivoExcel'), async (req, res) => {
    try {
        if (!req.file) return res.status(400).json({ success: false, message: 'Falta archivo' });

        let manual = JSON.parse(req.body.datosManuales || '{}');
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(req.file.buffer);
        const worksheet = workbook.worksheets[0];

        const colMap = {};
        worksheet.getRow(1).eachCell((cell, colNumber) => {
            const val = cell.value?.toString().toLowerCase().trim() || '';
            if (val.includes('fecha')) colMap.fecha = colNumber;
            if (val.includes('hora')) colMap.hora = colNumber;
            if (val.includes('sector')) colMap.sector = colNumber;
            if (val.includes('comentario')) colMap.comentario = colNumber;
            if (val.includes('calificacion') && !val.includes('desc')) colMap.rating = colNumber;
        });

        const sectores = {};

        worksheet.eachRow((row, rowNum) => {
            if (rowNum === 1) return;
            const rating = parseInt(row.getCell(colMap.rating)?.value);
            const dateVal = row.getCell(colMap.fecha)?.value;
            if (!dateVal || isNaN(rating)) return;

            let hVal = row.getCell(colMap.hora)?.value;
            let horaReal = (hVal instanceof Date) ? hVal.getHours() : (typeof hVal === 'number' ? Math.floor(hVal * 24) : 12);
            let date = (dateVal instanceof Date) ? dateVal : new Date(dateVal);
            const sector = (row.getCell(colMap.sector)?.value || 'General').toString().trim();
            const comment = (row.getCell(colMap.comentario)?.value || '').toString().trim();

            if (!sectores[sector]) {
                sectores[sector] = {
                    meses: Array.from({length: 12}, () => ({ mp:0, p:0, n:0, mn:0, total:0 })),
                    statsHoras: Array.from({length: 24}, () => ({ total: 0, neg: 0 })),
                    comsPos: [], comsNeg: [], palabrasPos: [], palabrasNeg: []
                };
            }

            const s = sectores[sector];
            const statsMes = s.meses[date.getMonth()];
            const statsH = s.statsHoras[horaReal];

            statsMes.total++;
            statsH.total++;
            if (rating === 4) { statsMes.mp++; s.palabrasPos.push(...getWords(comment)); }
            if (rating === 3) statsMes.p++;
            if (rating === 2) { statsMes.n++; statsH.neg++; s.palabrasNeg.push(...getWords(comment)); }
            if (rating === 1) { statsMes.mn++; statsH.neg++; s.palabrasNeg.push(...getWords(comment)); }

            if (esComentarioInteligente(comment, sector)) {
                const info = { texto: comment, meta: `${date.getDate()}/${date.getMonth()+1} ${horaReal}:00hs` };
                if (rating >= 3) s.comsPos.push(info);
                else s.comsNeg.push(info);
            }
        });

        const final = Object.entries(sectores).map(([nombre, data]) => {
            ['enero', 'febrero'].forEach((m, i) => { if (manual[m]) data.meses[i] = manual[m]; });

            let sumaSat = 0, mesesData = 0;
            const mesesFinal = data.meses.map((m) => {
                const f = m.total / 100;
                const val = m.total > 0 ? ((m.mp / f) - ((m.mn + m.n) / f)) : 0;
                if (m.total > 0) { sumaSat += val; mesesData++; }
                return { sat: parseFloat(val.toFixed(1)), total: m.total };
            });

            let hCritica = "--"; let maxNeg = -1; let porcCritico = 0;
            data.statsHoras.forEach((h, i) => {
                if (h.total >= 3) {
                    const p = (h.neg / h.total) * 100;
                    if (p > maxNeg) { maxNeg = p; hCritica = i.toString().padStart(2, '0') + ':00'; porcCritico = p.toFixed(0); }
                }
            });

            const getFreq = (arr) => Object.entries(arr.reduce((a,w)=>(a[w]=(a[w]||0)+1,a),{})).sort((a,b)=>b[1]-a[1]).slice(0, 25);

            return {
                nombre, meses: mesesFinal,
                comentarios: { pos: data.comsPos.sort((a,b)=>b.texto.length-a.texto.length).slice(0,3), neg: data.comsNeg.sort((a,b)=>b.texto.length-a.texto.length).slice(0,3) },
                nubePos: getFreq(data.palabrasPos), nubeNeg: getFreq(data.palabrasNeg),
                satAnual: mesesData > 0 ? (sumaSat / mesesData).toFixed(1) : "0.0",
                infoHora: { hora: hCritica, porcentaje: porcCritico }
            };
        });

        res.json({ success: true, data: { sectores: final } });
    } catch (e) { res.status(500).json({ success: false, message: e.message }); }
});

app.listen(PORT, () => console.log(`Server ON`));
