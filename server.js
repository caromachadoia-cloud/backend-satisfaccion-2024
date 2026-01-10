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
    "atencion": ["atencion", "trato", "amable", "ayudo", "atenta", "chica", "chico", "resolvio", "espera"],
    "gastro": ["comida", "mozo", "frio", "caliente", "rico", "bebida", "mesa", "pedido", "tardo"],
    "cajas": ["pago", "cobro", "fila", "rapido", "dinero", "efectivo", "tarjeta", "atencion"],
    "traslados": ["chofer", "auto", "camioneta", "viaje", "llego", "tarde", "limpio", "conduccion"],
    "general": ["lugar", "instalaciones", "baño", "limpieza", "seguridad"]
};

function esComentarioInteligente(texto, sector) {
    if (!texto || texto.length < 20) return false;
    const limpio = texto.toLowerCase();
    
    // 1. Filtrar basura (letras repetidas, sin sentido o solo emojis)
    if (!/[a-z]{4,}/.test(limpio)) return false; 

    // 2. Filtrar off-topic (ej: si es traslados y habla de "maquinas")
    const sectorKey = Object.keys(CONTEXTO_SECTORES).find(k => sector.toLowerCase().includes(k)) || "general";
    const palabrasProhibidas = ["maquina", "paga", "premio", "suerte"].filter(p => !CONTEXTO_SECTORES[sectorKey].includes(p));
    
    if (palabrasProhibidas.some(p => limpio.includes(p))) return false;

    return true;
}

const STOPWORDS = ['de', 'la', 'que', 'el', 'en', 'y', 'a', 'los', 'del', 'se', 'las', 'por', 'un', 'para', 'con', 'no', 'una', 'su', 'al', 'lo', 'como', 'más', 'pero', 'sus', 'le', 'ya', 'o', 'este', 'ha', 'me', 'si', 'sin', 'sobre', 'muy', 'cuando', 'también', 'hasta', 'hay', 'donde', 'quien', 'desde', 'todo', 'nos', 'uno', 'ni', 'contra', 'ese', 'eso', 'mi', 'qué', 'e', 'son', 'fue', 'gracias', 'hola'];

function getWords(text) {
    return text.toLowerCase().replace(/[.,/#!$%^&*;:{}=\-_`~()]/g, " ").match(/[a-záéíóúñü]+/g)
        ?.filter(word => !STOPWORDS.includes(word) && word.length > 3) || [];
}

app.post('/procesar-anual', upload.single('archivoExcel'), async (req, res) => {
    try {
        if (!req.file) return res.status(400).json({ success: false, message: 'No hay archivo' });

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

            // FIX HORA: Intentar obtener hora real
            let hVal = row.getCell(colMap.hora)?.value;
            let horaReal = 12;
            if (hVal instanceof Date) horaReal = hVal.getHours();
            else if (typeof hVal === 'number') horaReal = Math.floor(hVal * 24);
            else if (typeof hVal === 'string' && hVal.includes(':')) horaReal = parseInt(hVal.split(':')[0]);

            let date = (dateVal instanceof Date) ? dateVal : new Date(dateVal);
            const sector = (row.getCell(colMap.sector)?.value || 'General').toString().trim();
            const comment = (row.getCell(colMap.comentario)?.value || '').toString().trim();

            if (!sectores[sector]) {
                sectores[sector] = {
                    meses: Array.from({length: 12}, () => ({ mp:0, p:0, n:0, mn:0, total:0 })),
                    comsPos: [], comsNeg: [], palabrasPos: [], palabrasNeg: []
                };
            }

            const s = sectores[sector];
            const stats = s.meses[date.getMonth()];
            stats.total++;
            if (rating === 4) stats.mp++;
            if (rating === 3) stats.p++;
            if (rating === 2) stats.n++;
            if (rating === 1) stats.mn++;

            const info = { texto: comment, meta: `${date.getDate()}/${date.getMonth()+1} ${horaReal}:00hs` };
            
            if (esComentarioInteligente(comment, sector)) {
                if (rating >= 3) { s.comsPos.push(info); s.palabrasPos.push(...getWords(comment)); }
                else { s.comsNeg.push(info); s.palabrasNeg.push(...getWords(comment)); }
            }
        });

        const final = Object.entries(sectores).map(([nombre, data]) => {
            ['enero', 'febrero'].forEach((m, i) => {
                if (manual[m]) data.meses[i] = manual[m];
            });

            const mesesFinal = data.meses.map((m, i) => {
                const f = m.total / 100;
                const val = m.total > 0 ? ( (m.mp / f) - ((m.mn + m.n) / f) ).toFixed(1) : 0;
                return {
                    nombre: ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic'][i],
                    sat: parseFloat(val),
                    total: m.total
                };
            });

            const getFreq = (arr) => Object.entries(arr.reduce((a,w)=>(a[w]=(a[w]||0)+1,a),{})).sort((a,b)=>b[1]-a[1]).slice(0, 30);

            return {
                nombre, meses: mesesFinal,
                comentarios: { 
                    pos: data.comsPos.sort((a,b)=>b.texto.length-a.texto.length).slice(0,3), 
                    neg: data.comsNeg.sort((a,b)=>b.texto.length-a.texto.length).slice(0,3) 
                },
                nubePos: getFreq(data.palabrasPos),
                nubeNeg: getFreq(data.palabrasNeg),
                satAnual: (mesesFinal.reduce((a, b) => a + b.sat, 0) / 12).toFixed(1)
            };
        });

        res.json({ success: true, data: { sectores: final } });
    } catch (e) { res.status(500).json({ success: false, message: e.message }); }
});

app.listen(PORT, () => console.log(`Server ON`));
