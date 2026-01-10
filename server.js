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

// --- UTILIDADES ---

const normalizarTexto = (t) => {
    return t?.toString().toLowerCase().trim().normalize("NFD").replace(/[\u0300-\u036f]/g, "") || "";
};

function parsearFecha(valor) {
    if (!valor) return null;
    if (valor instanceof Date) return valor;
    if (typeof valor === 'number') return new Date(Math.round((valor - 25569) * 86400 * 1000));
    if (typeof valor === 'string') {
        let soloFecha = valor.split(' ')[0].trim();
        const separador = soloFecha.includes('/') ? '/' : (soloFecha.includes('-') ? '-' : null);
        if (separador) {
            const partes = soloFecha.split(separador);
            if (partes.length === 3) {
                const dia = parseInt(partes[0]);
                const mes = parseInt(partes[1]) - 1;
                const anio = parseInt(partes[2]);
                const fechaGen = new Date(anio, mes, dia);
                if (!isNaN(fechaGen.getTime())) return fechaGen;
            }
        }
        const intento = new Date(valor);
        return isNaN(intento.getTime()) ? null : intento;
    }
    return null;
}

const CONTEXTO_SECTORES = {
    "atencion": ["atencion", "trato", "amable", "ayudo", "atenta", "chica", "chico", "resolvio", "espera", "tarjeta", "fun", "stand", "personal", "empleado", "gente", "recepcion", "guardarropas"],
    "gastro": ["comida", "mozo", "frio", "caliente", "rico", "bebida", "mesa", "pedido", "tardo", "restaurant", "confiteria", "hamburguesa", "menu", "sabor", "cafe", "barra"],
    "cajas": ["pago", "cobro", "fila", "rapido", "dinero", "efectivo", "tarjeta", "atencion", "espera", "cajero", "ticket", "cobrar", "caja"],
    "traslados": ["chofer", "auto", "camioneta", "viaje", "llego", "tarde", "limpio", "conduccion", "carrito", "transporte", "valet", "bus"],
    "baños": ["baño", "limpieza", "olor", "jabon", "papel", "sucio", "higienico", "sanitario", "agua", "inodoro"]
};

const PROHIBIDAS_POR_RUIDO = ["maquina", "paga", "premio", "suerte", "ruleta", "slot", "ganar", "perder", "apuesta", "pozo"];

function esComentarioValido(texto, sector) {
    if (!texto || texto.toString().length < 4) return false; 
    const limpio = normalizarTexto(texto);
    const sectorKey = Object.keys(CONTEXTO_SECTORES).find(k => normalizarTexto(sector).includes(k)) || "general";
    if (sectorKey !== "general" && PROHIBIDAS_POR_RUIDO.some(p => limpio.includes(p))) return false;
    const palabrasContexto = CONTEXTO_SECTORES[sectorKey] || [];
    return palabrasContexto.some(p => limpio.includes(p)) || sectorKey === "general";
}

function extractWords(text) {
    const STOPWORDS = ['de', 'la', 'que', 'el', 'en', 'y', 'a', 'los', 'del', 'se', 'las', 'por', 'un', 'para', 'con', 'no', 'una', 'su', 'al', 'lo', 'como', 'más', 'pero', 'sus', 'le', 'ya', 'o', 'este', 'ha', 'me', 'si', 'sin', 'sobre', 'muy', 'cuando', 'también', 'hasta', 'hay', 'donde', 'quien', 'desde', 'todo', 'nos', 'uno', 'ni', 'contra', 'ese', 'eso', 'mi', 'qué', 'porque', 'estaba', 'fui', 'era', 'son', 'fue'];
    return normalizarTexto(text).replace(/[.,/#!$%^&*;:{}=\-_`~()]/g, " ").match(/[a-z]+/g)
        ?.filter(word => !STOPWORDS.includes(word) && word.length > 3) || [];
}

function convertirRating(valor) {
    if (typeof valor === 'number') return valor;
    if (!valor) return NaN;
    let num = parseInt(valor);
    if (!isNaN(num)) return num;
    const t = normalizarTexto(valor);
    if (t.includes('muy positiva') || t.includes('excelente')) return 4;
    if (t.includes('positiva') || t.includes('buena')) return 3;
    if (t.includes('muy negativa') || t.includes('muy mala') || t.includes('pesima')) return 1;
    if (t.includes('negativa') || t.includes('mala') || t.includes('regular')) return 2;
    return NaN;
}

app.post('/procesar-anual', upload.single('archivoExcel'), async (req, res) => {
    try {
        if (!req.file) return res.status(400).json({ success: false, message: 'Archivo no recibido' });
        
        let manual = {};
        try { manual = JSON.parse(req.body.datosManuales || '{}'); } catch(e) {}

        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(req.file.buffer);
        const worksheet = workbook.worksheets[0];

        console.log(`📂 Hoja: ${worksheet.name} - Filas totales: ${worksheet.rowCount}`);

        const colMap = { fecha: null, hora: null, sector: null, ubicacion: null, comentario: null, rating: null };
        let headerRowIndex = 1;
        let ratingPriority = 0;

        // 1. Detección Inteligente de Columnas
        for(let r = 1; r <= 5; r++) {
            const row = worksheet.getRow(r);
            row.eachCell((cell, colNumber) => {
                const val = normalizarTexto(cell.value);
                
                if (val.includes('fecha')) colMap.fecha = colNumber;
                if (val.includes('hora')) colMap.hora = colNumber;
                if (val.includes('sector')) colMap.sector = colNumber;
                if (val.includes('ubicacion')) colMap.ubicacion = colNumber;
                if (val.includes('comentario')) colMap.comentario = colNumber;
                
                let currentPriority = 0;
                if (val === 'calificacion' || val === 'rating') currentPriority = 3;
                else if (val.includes('calificacion') && !val.includes('desc')) currentPriority = 2;
                else if (val === 'nota' || (val.includes('puntos') && !val.includes('criticos'))) currentPriority = 1;

                if (currentPriority > ratingPriority) {
                    colMap.rating = colNumber;
                    ratingPriority = currentPriority;
                }
            });
            if (colMap.fecha && colMap.rating && ratingPriority >= 2) {
                headerRowIndex = r;
                break;
            }
        }

        console.log("📍 Mapeo final:", colMap);

        if (!colMap.fecha || !colMap.rating) {
            return res.json({ success: false, message: 'Faltan columnas Fecha o Calificación.' });
        }

        const sectores = {};
        let filasProcesadas = 0;

        // 2. Procesamiento de filas
        worksheet.eachRow((row, rowNum) => {
            if (rowNum <= headerRowIndex) return;

            const ratingVal = row.getCell(colMap.rating)?.value;
            let rating = convertirRating(ratingVal);
            if (typeof ratingVal === 'object' && ratingVal?.result) rating = convertirRating(ratingVal.result);

            if (isNaN(rating)) return;

            const dateVal = row.getCell(colMap.fecha)?.value;
            let date = parsearFecha(dateVal);
            if (!date) return;
            
            const mIdx = date.getMonth();

            let horaReal = 12;
            const hVal = colMap.hora ? row.getCell(colMap.hora)?.value : null;
            if (hVal instanceof Date) horaReal = hVal.getHours();
            else if (typeof hVal === 'number') horaReal = Math.floor(hVal * 24);
            else if (typeof hVal === 'string') {
                const partesHora = hVal.trim().split(':');
                if (partesHora.length >= 1) horaReal = parseInt(partesHora[0]);
            }
            if (isNaN(horaReal)) horaReal = 12;
            if (horaReal < 0) horaReal = 0; if (horaReal > 23) horaReal = 23;

            const sectorName = colMap.sector ? (row.getCell(colMap.sector)?.value || 'General').toString().trim() : 'General';
            const ubicName = colMap.ubicacion ? (row.getCell(colMap.ubicacion)?.value || 'General').toString().trim() : 'General';
            const comment = colMap.comentario ? (row.getCell(colMap.comentario)?.value || '').toString().trim() : '';

            if (!sectores[sectorName]) {
                sectores[sectorName] = {
                    meses: Array.from({length: 12}, () => ({ mp:0, p:0, n:0, mn:0, total:0 })),
                    statsHoras: Array.from({length: 24}, () => ({ total: 0, neg: 0 })),
                    ubicaciones: {}, 
                    comsPos: [], comsNeg: [], 
                    palabrasPos: [], palabrasNeg: []
                };
            }

            const s = sectores[sectorName];
            filasProcesadas++;

            s.meses[mIdx].total++;
            s.statsHoras[horaReal].total++;

            if (!s.ubicaciones[ubicName]) s.ubicaciones[ubicName] = { mp:0, p:0, n:0, mn:0, total:0 };
            s.ubicaciones[ubicName].total++;

            if (rating >= 4) { s.meses[mIdx].mp++; s.ubicaciones[ubicName].mp++; }
            else if (rating === 3) { s.meses[mIdx].p++; s.ubicaciones[ubicName].p++; }
            else if (rating === 2) { s.meses[mIdx].n++; s.ubicaciones[ubicName].n++; s.statsHoras[horaReal].neg++; }
            else if (rating <= 1) { s.meses[mIdx].mn++; s.ubicaciones[ubicName].mn++; s.statsHoras[horaReal].neg++; }

            if (esComentarioValido(comment, sectorName)) {
                const info = { texto: comment, meta: `${date.getDate()}/${mIdx+1} ${horaReal}:00hs` };
                if (rating >= 3) { s.comsPos.push(info); s.palabrasPos.push(...extractWords(comment)); }
                else { s.comsNeg.push(info); s.palabrasNeg.push(...extractWords(comment)); }
            }
        });

        console.log(`✅ Procesado. Filas OK: ${filasProcesadas}.`);

        if (filasProcesadas === 0) {
            return res.json({ success: false, message: `0 filas válidas.` });
        }

        // 3. Cálculos Finales
        const final = Object.entries(sectores).map(([nombre, data]) => {
            
            // --- CORRECCIÓN NaN EN DATOS MANUALES ---
            ['enero', 'febrero'].forEach((m, i) => { 
                if (manual[m] && manual[m].total > 0) {
                    const total = manual[m].total || 0;
                    const mp = manual[m].mp || 0;
                    const mn = manual[m].mn || 0;
                    const n = manual[m].n || 0;
                    
                    // Calculamos los Positivos (p) por diferencia para evitar NaN
                    let p = total - mp - mn - n;
                    if (p < 0) p = 0;

                    // Reemplazamos el mes con los datos completos
                    data.meses[i] = { total, mp, mn, n, p };
                }
            });

            // Satisfacción Mensual
            let sumaSat = 0, mesesConDato = 0;
            const mesesFinal = data.meses.map((m) => {
                let val = 0;
                if(m.total > 0) {
                    // Aseguramos que existan los valores (|| 0) para evitar NaN
                    const positivos = (m.mp || 0) + (m.p || 0);
                    val = (positivos / m.total) * 100;
                }
                
                if (m.total > 0) { sumaSat += val; mesesConDato++; }
                return { sat: parseFloat(val.toFixed(1)), total: m.total };
            });

            // Hora Crítica
            let hCritica = "12:00"; let maxNegVol = -1; let volNegTotal = 0; let totalEnEsaHora = 0;
            data.statsHoras.forEach((h, i) => {
                if (h.neg > maxNegVol) {
                    maxNegVol = h.neg; 
                    volNegTotal = h.neg; 
                    totalEnEsaHora = h.total;
                    hCritica = i.toString().padStart(2, '0') + ':00';
                }
            });
            let porcNegCritica = totalEnEsaHora > 0 ? ((volNegTotal / totalEnEsaHora) * 100).toFixed(1) : "0.0";

            // Métricas Ubicación
            const metricsUbic = Object.entries(data.ubicaciones).map(([uNom, uD]) => {
                const positivos = (uD.mp || 0) + (uD.p || 0);
                const uSat = uD.total > 0 ? ((positivos / uD.total) * 100).toFixed(1) : "0.0";
                return { nombre: uNom, totalAnual: uD.total, satProm: uSat, promDiario: (uD.total / 365).toFixed(2) };
            }).sort((a,b) => b.totalAnual - a.totalAnual);

            // Nubes
            const freq = (arr) => Object.entries(arr.reduce((a,w)=>(a[w]=(a[w]||0)+1,a),{})).sort((a,b)=>b[1]-a[1]).slice(0, 40);

            return {
                nombre, meses: mesesFinal, ubicaciones: metricsUbic,
                comentarios: { pos: data.comsPos.slice(0,5), neg: data.comsNeg.slice(0,5) },
                nubePos: freq(data.palabrasPos), nubeNeg: freq(data.palabrasNeg),
                satAnual: mesesConDato > 0 ? (sumaSat / mesesConDato).toFixed(1) : "0.0",
                infoHora: { hora: hCritica, porcentaje: porcNegCritica, volumenNeg: maxNegVol }
            };
        });

        res.json({ success: true, data: { sectores: final } });
    } catch (e) { 
        console.error(e);
        res.status(500).json({ success: false, message: e.message }); 
    }
});

app.listen(PORT, () => console.log(`Server ON port ${PORT}`));
