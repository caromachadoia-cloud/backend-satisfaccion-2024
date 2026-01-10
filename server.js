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

// DICCIONARIO ESTRICTO POR SECTOR
const CONTEXTO_SECTORES = {
    "atencion": ["atencion", "trato", "amable", "ayudo", "atenta", "chica", "chico", "resolvio", "espera", "tarjeta", "fun", "stand"],
    "gastro": ["comida", "mozo", "frio", "caliente", "rico", "bebida", "mesa", "pedido", "tardo", "restaurant", "confiteria", "hamburguesa"],
    "cajas": ["pago", "cobro", "fila", "rapido", "dinero", "efectivo", "tarjeta", "atencion", "espera", "cajero", "ticket"],
    "traslados": ["chofer", "auto", "camioneta", "viaje", "llego", "tarde", "limpio", "conduccion", "carrito", "estacionamiento", "valet"],
    "baños": ["baño", "limpieza", "olor", "jabon", "papel", "sucio", "higienico"]
};

const PALABRAS_PROHIBIDAS_GENERAL = ["maquina", "paga", "premio", "suerte", "ruleta", "slot"];

function esComentarioRelevante(texto, sector) {
    if (!texto || texto.length < 20) return false;
    const limpio = texto.toLowerCase();
    const sectorKey = Object.keys(CONTEXTO_SECTORES).find(k => sector.toLowerCase().includes(k));
    
    // 1. Si no detectamos el sector, filtramos solo basura general
    if (!sectorKey) return !PALABRAS_PROHIBIDAS_GENERAL.some(p => limpio.includes(p));

    // 2. Filtro estricto: Debe contener al menos una palabra del contexto del sector
    const tieneContexto = CONTEXTO_SECTORES[sectorKey].some(p => limpio.includes(p));
    
    // 3. Si habla de máquinas y el sector no es máquinas, se descarta
    const hablaDeMaquinas = PALABRAS_PROHIBIDAS_GENERAL.some(p => limpio.includes(p));
    
    return tieneContexto && !hablaDeMaquinas;
}

function getWords(text) {
    const STOPWORDS = ['de', 'la', 'que', 'el', 'en', 'y', 'a', 'los', 'del', 'se', 'las', 'por', 'un', 'para', 'con', 'no', 'una', 'su', 'al', 'lo', 'como', 'más', 'pero', 'sus', 'le', 'ya', 'o', 'este', 'ha', 'me', 'si', 'sin', 'sobre', 'muy', 'cuando', 'también', 'hasta', 'hay', 'donde', 'quien', 'desde', 'todo', 'nos', 'uno', 'ni', 'contra', 'ese', 'eso', 'mi', 'qué'];
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

            // EXTRACCIÓN DE HORA MEJORADA
            let hVal = row.getCell(colMap.hora)?.value;
            let horaReal = 12;
            if (hVal instanceof Date) {
                horaReal = hVal.getHours();
            } else if (typeof hVal === 'number') {
                horaReal = Math.floor(hVal * 24);
            } else if (typeof hVal === 'string' && hVal.includes(':')) {
                horaReal = parseInt(hVal.split(':')[0]);
            }

            let date = (dateVal instanceof Date) ? dateVal : new Date(dateVal);
            const sector = (row.getCell(colMap.sector)?.value || 'General').toString().trim();
            const ubicacion = (row.getCell(colMap.ubicacion)?.value || 'Sin Ubicación').toString().trim();
            const comment = (row.getCell(colMap.comentario)?.value || '').toString().trim();

            if (!sectores[sector]) {
                sectores[sector] = {
                    meses: Array.from({length: 12}, () => ({ mp:0, p:0, n:0, mn:0, total:0 })),
                    statsHoras: Array.from({length: 24}, () => ({ total: 0, neg: 0 })),
                    ubicaciones: {}, comsPos: [], comsNeg: [], palabrasPos: [], palabrasNeg: []
                };
            }

            const s = sectores[sector];
            const mIdx = date.getMonth();
            s.meses[mIdx].total++;
            s.statsHoras[horaReal].total++;

            if (!s.ubicaciones[ubicacion]) s.ubicaciones[ubicacion] = { mp:0, p:0, n:0, mn:0, total:0 };
            s.ubicaciones[ubicacion].total++;

            if (rating === 4) { s.meses[mIdx].mp++; s.ubicaciones[ubicacion].mp++; }
            if (rating === 3) { s.meses[mIdx].p++; s.ubicaciones[ubicacion].p++; }
            if (rating === 2) { s.meses[mIdx].n++; s.ubicaciones[ubicacion].n++; s.statsHoras[horaReal].neg++; }
            if (rating === 1) { s.meses[mIdx].mn++; s.ubicaciones[ubicacion].mn++; s.statsHoras[horaReal].neg++; }

            if (esComentarioRelevante(comment, sector)) {
                const info = { texto: comment, meta: `${date.getDate()}/${date.getMonth()+1} ${horaReal}:00hs` };
                if (rating >= 3) { s.comsPos.push(info); s.palabrasPos.push(...getWords(comment)); }
                else { s.comsNeg.push(info); s.palabrasNeg.push(...getWords(comment)); }
            }
        });

        const final = Object.entries(sectores).map(([nombre, data]) => {
            ['enero', 'febrero'].forEach((m, i) => { if (manual[m]) data.meses[i] = manual[m]; });

            const mesesFinal = data.meses.map((m) => {
                const f = m.total / 100;
                const val = m.total > 0 ? ((m.mp / f) - ((m.mn - m.n) / f)) : 0;
                return { sat: parseFloat(val.toFixed(1)), total: m.total };
            });

            // HORA CRÍTICA REAL (Basada en tasa de negatividad)
            let hCritica = "00:00"; let maxRate = -1; let porc = 0;
            data.statsHoras.forEach((h, i) => {
                if (h.total >= 5) {
                    const rate = (h.neg / h.total) * 100;
                    if (rate > maxRate) { maxRate = rate; hCritica = i.toString().padStart(2, '0') + ':00'; porc = rate.toFixed(1); }
                }
            });

            // MÉTRICAS POR UBICACIÓN
            const rankingUbicaciones = Object.entries(data.ubicaciones).map(([uNom, uData]) => {
                const uf = uData.total / 100;
                const uSat = uData.total > 0 ? ((uData.mp / uf) - ((uData.mn - uData.n) / uf)).toFixed(1) : 0;
                return {
                    nombre: uNom,
                    totalAnual: uData.total,
                    satPromedio: uSat,
                    promedioDiario: (uData.total / 365).toFixed(2)
                };
            }).sort((a,b) => b.totalAnual - a.totalAnual);

            const getFreq = (arr) => Object.entries(arr.reduce((a,w)=>(a[w]=(a[w]||0)+1,a),{})).sort((a,b)=>b[1]-a[1]).slice(0, 30);

            return {
                nombre, meses: mesesFinal,
                ubicaciones: rankingUbicaciones,
                comentarios: { 
                    pos: data.comsPos.sort((a,b)=>b.texto.length-a.texto.length).slice(0,3), 
                    neg: data.comsNeg.sort((a,b)=>b.texto.length-a.texto.length).slice(0,3) 
                },
                nubePos: getFreq(data.palabrasPos), nubeNeg: getFreq(data.palabrasNeg),
                satAnual: (mesesFinal.reduce((a, b) => a + b.sat, 0) / 12).toFixed(1),
                infoHora: { hora: hCritica, porcentaje: porc }
            };
        });

        res.json({ success: true, data: { sectores: final } });
    } catch (e) { res.status(500).json({ success: false, message: e.message }); }
});

app.listen(PORT, () => console.log(`Server ON`));
