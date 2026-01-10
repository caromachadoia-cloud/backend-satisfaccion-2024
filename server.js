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

// DICCIONARIO PARA FILTRADO INTELIGENTE POR SECTOR
const CONTEXTO_SECTORES = {
    "atencion": ["atencion", "trato", "amable", "ayudo", "atenta", "chica", "chico", "resolvio", "espera", "tarjeta", "fun", "stand"],
    "gastro": ["comida", "mozo", "frio", "caliente", "rico", "bebida", "mesa", "pedido", "tardo", "restaurant", "confiteria", "hamburguesa", "menu"],
    "cajas": ["pago", "cobro", "fila", "rapido", "dinero", "efectivo", "tarjeta", "atencion", "espera", "cajero", "ticket", "cobrar"],
    "traslados": ["chofer", "auto", "camioneta", "viaje", "llego", "tarde", "limpio", "conduccion", "carrito", "transporte", "valet"],
    "baños": ["baño", "limpieza", "olor", "jabon", "papel", "sucio", "higienico", "sanitario"]
};

const PALABRAS_PROHIBIDAS_GENERAL = ["maquina", "paga", "premio", "suerte", "ruleta", "slot", "ganar"];

function esComentarioRelevante(texto, sector) {
    if (!texto || texto.length < 20) return false;
    const limpio = texto.toLowerCase();
    const sectorKey = Object.keys(CONTEXTO_SECTORES).find(k => sector.toLowerCase().includes(k)) || "general";
    
    // Si el sector no es máquinas, filtramos cualquier comentario que hable de máquinas para evitar ruido
    if (sectorKey !== "general" && PALABRAS_PROHIBIDAS_GENERAL.some(p => limpio.includes(p))) return false;

    // El comentario debe tener relación con el sector
    const palabrasContexto = CONTEXTO_SECTORES[sectorKey] || [];
    return palabrasContexto.some(p => limpio.includes(p)) || sectorKey === "general";
}

function extractWords(text) {
    const STOPWORDS = ['de', 'la', 'que', 'el', 'en', 'y', 'a', 'los', 'del', 'se', 'las', 'por', 'un', 'para', 'con', 'no', 'una', 'su', 'al', 'lo', 'como', 'más', 'pero', 'sus', 'le', 'ya', 'o', 'este', 'ha', 'me', 'si', 'sin', 'sobre', 'muy', 'cuando', 'también', 'hasta', 'hay', 'donde', 'quien', 'desde', 'todo', 'nos', 'uno', 'ni', 'contra', 'ese', 'eso', 'mi', 'qué'];
    return text.toLowerCase().replace(/[.,/#!$%^&*;:{}=\-_`~()]/g, " ").match(/[a-záéíóúñü]+/g)
        ?.filter(word => !STOPWORDS.includes(word) && word.length > 3) || [];
}

app.post('/procesar-anual', upload.single('archivoExcel'), async (req, res) => {
    try {
        if (!req.file) return res.status(400).json({ success: false, message: 'Archivo no recibido' });
        let manual = JSON.parse(req.body.datosManuales || '{}');
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(req.file.buffer);
        const worksheet = workbook.worksheets[0];

        const colMap = {};
        worksheet.getRow(1).eachCell((cell, colNumber) => {
            const val = cell.value?.toString().toLowerCase().trim() || '';
            if (val === 'fecha') colMap.fecha = colNumber;
            if (val === 'hora') colMap.hora = colNumber;
            if (val === 'sector') colMap.sector = colNumber;
            if (val === 'ubicacion') colMap.ubicacion = colNumber;
            if (val.includes('comentario')) colMap.comentario = colNumber;
            if (val === 'calificacion') colMap.rating = colNumber;
        });

        const sectores = {};

        worksheet.eachRow((row, rowNum) => {
            if (rowNum === 1) return;
            const rating = parseInt(row.getCell(colMap.rating)?.value);
            const dateVal = row.getCell(colMap.fecha)?.value;
            if (!dateVal || isNaN(rating)) return;

            // HORA CRÍTICA REAL: Mapeo de formatos Excel
            let hVal = row.getCell(colMap.hora)?.value;
            let horaReal = (hVal instanceof Date) ? hVal.getHours() : (typeof hVal === 'number' ? Math.floor(hVal * 24) : 12);
            let date = (dateVal instanceof Date) ? dateVal : new Date(dateVal);
            
            const sectorName = (row.getCell(colMap.sector)?.value || 'General').toString().trim();
            const ubicName = (row.getCell(colMap.ubicacion)?.value || 'General').toString().trim();
            const comment = (row.getCell(colMap.comentario)?.value || '').toString().trim();

            if (!sectores[sectorName]) {
                sectores[sectorName] = {
                    meses: Array.from({length: 12}, () => ({ mp:0, p:0, n:0, mn:0, total:0 })),
                    statsHoras: Array.from({length: 24}, () => ({ total: 0, neg: 0 })),
                    ubicaciones: {}, comsPos: [], comsNeg: [], palabrasPos: [], palabrasNeg: []
                };
            }

            const s = sectores[sectorName];
            const mIdx = date.getMonth();
            s.meses[mIdx].total++;
            s.statsHoras[horaReal].total++;

            if (!s.ubicaciones[ubicName]) s.ubicaciones[ubicName] = { mp:0, p:0, n:0, mn:0, total:0 };
            s.ubicaciones[ubicName].total++;

            if (rating === 4) { s.meses[mIdx].mp++; s.ubicaciones[ubicName].mp++; }
            if (rating === 3) { s.meses[mIdx].p++; s.ubicaciones[ubicName].p++; }
            if (rating === 2) { s.meses[mIdx].n++; s.ubicaciones[ubicName].n++; s.statsHoras[horaReal].neg++; }
            if (rating === 1) { s.meses[mIdx].mn++; s.ubicaciones[ubicName].mn++; s.statsHoras[horaReal].neg++; }

            if (esComentarioRelevante(comment, sectorName)) {
                const info = { texto: comment, meta: `${date.getDate()}/${date.getMonth()+1} ${horaReal}:00hs` };
                if (rating >= 3) { s.comsPos.push(info); s.palabrasPos.push(...extractWords(comment)); }
                else { s.comsNeg.push(info); s.palabrasNeg.push(...extractWords(comment)); }
            }
        });

        const final = Object.entries(sectores).map(([nombre, data]) => {
            ['enero', 'febrero'].forEach((m, i) => { if (manual[m]) data.meses[i] = manual[m]; });

            let sumaSat = 0, mesesConDato = 0;
            const mesesFinal = data.meses.map((m) => {
                const factor = m.total / 100;
                // FORMULA SOLICITADA EXACTA: (E / A%) - ((B-C) / A%)
                const val = m.total > 0 ? ( (m.mp / factor) - ((m.mn - m.n) / factor) ) : 0;
                if (m.total > 0) { sumaSat += val; mesesConDato++; }
                return { sat: parseFloat(val.toFixed(1)), total: m.total };
            });

            // HORA CRÍTICA BASADA EN TASA DE NEGATIVIDAD
            let hCritica = "12:00"; let maxRate = -1; let porcNeg = 0;
            data.statsHoras.forEach((h, i) => {
                if (h.total >= 5) {
                    const rate = (h.neg / h.total) * 100;
                    if (rate > maxRate) { maxRate = rate; hCritica = i.toString().padStart(2, '0') + ':00'; porcNeg = rate.toFixed(1); }
                }
            });

            const metricsUbic = Object.entries(data.ubicaciones).map(([uNom, uD]) => {
                const f = uD.total / 100;
                const uSat = uD.total > 0 ? ((uD.mp/f) - ((uD.mn - uD.n)/f)).toFixed(1) : 0;
                return { nombre: uNom, totalAnual: uD.total, satProm: uSat, promDiario: (uD.total / 365).toFixed(2) };
            }).sort((a,b) => b.totalAnual - a.totalAnual);

            const freq = (arr) => Object.entries(arr.reduce((a,w)=>(a[w]=(a[w]||0)+1,a),{})).sort((a,b)=>b[1]-a[1]).slice(0, 30);

            return {
                nombre, meses: mesesFinal, ubicaciones: metricsUbic,
                comentarios: { 
                    pos: data.comsPos.sort((a,b)=>b.texto.length-a.texto.length).slice(0,3), 
                    neg: data.comsNeg.sort((a,b)=>b.texto.length-a.texto.length).slice(0,3) 
                },
                nubePos: freq(data.palabrasPos), nubeNeg: freq(data.palabrasNeg),
                satAnual: mesesConDato > 0 ? (sumaSat / mesesConDato).toFixed(1) : "0.0",
                infoHora: { hora: hCritica, porcentaje: porcNeg }
            };
        });

        res.json({ success: true, data: { sectores: final } });
    } catch (e) { res.status(500).json({ success: false, message: e.message }); }
});

app.listen(PORT, () => console.log(`Server ON 2025`));
