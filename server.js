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

// --- FUNCIONES DE AYUDA ---

// Limpia textos para comparar (quita tildes, mayúsculas y espacios extra)
const normalizarTexto = (t) => {
    return t?.toString().toLowerCase().trim().normalize("NFD").replace(/[\u0300-\u036f]/g, "") || "";
};

const CONTEXTO_SECTORES = {
    "atencion": ["atencion", "trato", "amable", "ayudo", "atenta", "chica", "chico", "resolvio", "espera", "tarjeta", "fun", "stand", "personal", "empleado", "gente"],
    "gastro": ["comida", "mozo", "frio", "caliente", "rico", "bebida", "mesa", "pedido", "tardo", "restaurant", "confiteria", "hamburguesa", "menu", "sabor", "cafe"],
    "cajas": ["pago", "cobro", "fila", "rapido", "dinero", "efectivo", "tarjeta", "atencion", "espera", "cajero", "ticket", "cobrar", "caja"],
    "traslados": ["chofer", "auto", "camioneta", "viaje", "llego", "tarde", "limpio", "conduccion", "carrito", "transporte", "valet", "bus"],
    "baños": ["baño", "limpieza", "olor", "jabon", "papel", "sucio", "higienico", "sanitario", "agua", "inodoro"]
};

const PROHIBIDAS_POR_RUIDO = ["maquina", "paga", "premio", "suerte", "ruleta", "slot", "ganar", "perder", "apuesta"];

function esComentarioValido(texto, sector) {
    if (!texto || texto.length < 10) return false;
    const limpio = normalizarTexto(texto);
    const sectorKey = Object.keys(CONTEXTO_SECTORES).find(k => normalizarTexto(sector).includes(k)) || "general";
    
    // Si NO es sector general y aparece palabra prohibida (ej. queja de slot en gastronomía), se descarta
    if (sectorKey !== "general" && PROHIBIDAS_POR_RUIDO.some(p => limpio.includes(p))) return false;

    const palabrasContexto = CONTEXTO_SECTORES[sectorKey] || [];
    // Debe contener palabras del contexto o ser general
    return palabrasContexto.some(p => limpio.includes(p)) || sectorKey === "general";
}

function extractWords(text) {
    const STOPWORDS = ['de', 'la', 'que', 'el', 'en', 'y', 'a', 'los', 'del', 'se', 'las', 'por', 'un', 'para', 'con', 'no', 'una', 'su', 'al', 'lo', 'como', 'más', 'pero', 'sus', 'le', 'ya', 'o', 'este', 'ha', 'me', 'si', 'sin', 'sobre', 'muy', 'cuando', 'también', 'hasta', 'hay', 'donde', 'quien', 'desde', 'todo', 'nos', 'uno', 'ni', 'contra', 'ese', 'eso', 'mi', 'qué', 'porque', 'estaba', 'fui', 'era', 'son'];
    return normalizarTexto(text).replace(/[.,/#!$%^&*;:{}=\-_`~()]/g, " ").match(/[a-z]+/g)
        ?.filter(word => !STOPWORDS.includes(word) && word.length > 3) || [];
}

// --- ENDPOINT PRINCIPAL ---

app.post('/procesar-anual', upload.single('archivoExcel'), async (req, res) => {
    try {
        if (!req.file) return res.status(400).json({ success: false, message: 'Archivo no recibido' });
        
        let manual = {};
        try { manual = JSON.parse(req.body.datosManuales || '{}'); } catch(e) {}

        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(req.file.buffer);
        const worksheet = workbook.worksheets[0]; // Primera hoja

        console.log(`📂 Procesando archivo. Hoja: ${worksheet.name}`);

        const colMap = { fecha: null, hora: null, sector: null, ubicacion: null, comentario: null, rating: null };
        let headerRowIndex = 1;

        // 1. BUSCAR ENCABEZADOS (Hasta fila 5)
        for(let r = 1; r <= 5; r++) {
            const row = worksheet.getRow(r);
            row.eachCell((cell, colNumber) => {
                const val = normalizarTexto(cell.value);
                if (val.includes('fecha') || val === 'dia') colMap.fecha = colNumber;
                if (val.includes('hora')) colMap.hora = colNumber;
                if (val.includes('sector') || val.includes('area')) colMap.sector = colNumber;
                if (val.includes('ubicacion') || val.includes('lugar')) colMap.ubicacion = colNumber;
                if (val.includes('comentario') || val.includes('observacion')) colMap.comentario = colNumber;
                if (val.includes('calificacion') || val.includes('rating') || val.includes('valoracion') || val.includes('puntos') || val === 'nota') colMap.rating = colNumber;
            });
            if (colMap.fecha && colMap.rating) {
                headerRowIndex = r;
                break;
            }
        }

        console.log("📍 Mapeo de columnas:", colMap);

        if (!colMap.fecha || !colMap.rating) {
            return res.json({ success: false, message: 'No se encontraron las columnas Fecha y Calificación en las primeras filas.' });
        }

        const sectores = {};
        let filasProcesadas = 0;

        // 2. RECORRER FILAS
        worksheet.eachRow((row, rowNum) => {
            if (rowNum <= headerRowIndex) return;

            // Obtener Rating
            const ratingVal = row.getCell(colMap.rating)?.value;
            let rating = parseInt(ratingVal);
            if (typeof ratingVal === 'object' && ratingVal?.result) rating = parseInt(ratingVal.result);

            if (isNaN(rating) || rating === null) return;

            // Obtener Fecha
            const dateVal = row.getCell(colMap.fecha)?.value;
            if (!dateVal) return;
            
            let date = (dateVal instanceof Date) ? dateVal : new Date(dateVal);
            if (isNaN(date.getTime())) return;
            
            const mIdx = date.getMonth(); // 0 (Ene) a 11 (Dic)

            // Obtener Hora (Normalización compleja)
            let horaReal = 12; 
            const hVal = colMap.hora ? row.getCell(colMap.hora)?.value : null;
            
            if (hVal instanceof Date) {
                horaReal = hVal.getHours();
            } else if (typeof hVal === 'number') {
                // Excel decimal (0.5 = 12:00)
                horaReal = Math.floor(hVal * 24);
            } else if (typeof hVal === 'string' && hVal.includes(':')) {
                horaReal = parseInt(hVal.split(':')[0]);
            }
            // Limites 0-23
            if (horaReal < 0) horaReal = 0;
            if (horaReal > 23) horaReal = 23;

            // Textos
            const sectorName = colMap.sector ? (row.getCell(colMap.sector)?.value || 'General').toString().trim() : 'General';
            const ubicName = colMap.ubicacion ? (row.getCell(colMap.ubicacion)?.value || 'General').toString().trim() : 'General';
            const comment = colMap.comentario ? (row.getCell(colMap.comentario)?.value || '').toString().trim() : '';

            // Inicializar Estructura
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

            // Acumular
            s.meses[mIdx].total++;
            s.statsHoras[horaReal].total++;

            if (!s.ubicaciones[ubicName]) s.ubicaciones[ubicName] = { mp:0, p:0, n:0, mn:0, total:0 };
            s.ubicaciones[ubicName].total++;

            // Clasificar (5 y 4 positivos, 3 neutro/pos, 2 y 1 negativos)
            if (rating >= 4) { 
                s.meses[mIdx].mp++; s.ubicaciones[ubicName].mp++; 
            } else if (rating === 3) { 
                s.meses[mIdx].p++; s.ubicaciones[ubicName].p++; 
            } else if (rating === 2) { 
                s.meses[mIdx].n++; s.ubicaciones[ubicName].n++; 
                s.statsHoras[horaReal].neg++; 
            } else if (rating <= 1) { 
                s.meses[mIdx].mn++; s.ubicaciones[ubicName].mn++; 
                s.statsHoras[horaReal].neg++; 
            }

            // Comentarios
            if (esComentarioValido(comment, sectorName)) {
                const info = { texto: comment, meta: `${date.getDate()}/${mIdx+1} ${horaReal}:00hs` };
                if (rating >= 3) { s.comsPos.push(info); s.palabrasPos.push(...extractWords(comment)); }
                else { s.comsNeg.push(info); s.palabrasNeg.push(...extractWords(comment)); }
            }
        });

        console.log(`✅ Filas procesadas: ${filasProcesadas}`);

        if (filasProcesadas === 0) {
            return res.json({ success: false, message: 'Archivo leído pero sin filas válidas. Revisa formatos de Fecha/Rating.' });
        }

        // 3. CÁLCULOS FINALES
        const final = Object.entries(sectores).map(([nombre, data]) => {
            // Datos manuales
            ['enero', 'febrero'].forEach((m, i) => { if (manual[m] && manual[m].total > 0) data.meses[i] = manual[m]; });

            // Sat Anual
            let sumaSat = 0, mesesConDato = 0;
            const mesesFinal = data.meses.map((m) => {
                let val = 0;
                // CSAT: (Positivos / Total) * 100
                if(m.total > 0) val = ((m.mp + m.p) / m.total) * 100;
                if (m.total > 0) { sumaSat += val; mesesConDato++; }
                return { sat: parseFloat(val.toFixed(1)), total: m.total };
            });

            // HORA CRÍTICA (Por Volumen de Negativos)
            let hCritica = "12:00"; 
            let maxNegVol = -1;
            let volNegTotal = 0; 
            let totalEnEsaHora = 0;

            data.statsHoras.forEach((h, i) => {
                if (h.neg > maxNegVol) {
                    maxNegVol = h.neg;
                    volNegTotal = h.neg;
                    totalEnEsaHora = h.total;
                    hCritica = i.toString().padStart(2, '0') + ':00';
                }
            });
            // Porcentaje de rechazo en esa hora
            let porcNegCritica = totalEnEsaHora > 0 ? ((volNegTotal / totalEnEsaHora) * 100).toFixed(1) : "0.0";

            // Métricas Ubicación
            const metricsUbic = Object.entries(data.ubicaciones).map(([uNom, uD]) => {
                const uSat = uD.total > 0 ? (((uD.mp + uD.p) / uD.total) * 100).toFixed(1) : "0.0";
                return { nombre: uNom, totalAnual: uD.total, satProm: uSat, promDiario: (uD.total / 365).toFixed(2) };
            }).sort((a,b) => b.totalAnual - a.totalAnual); // Ordenar por volumen descendente

            // Nubes
            const freq = (arr) => Object.entries(arr.reduce((a,w)=>(a[w]=(a[w]||0)+1,a),{})).sort((a,b)=>b[1]-a[1]).slice(0, 40);

            return {
                nombre, 
                meses: mesesFinal, 
                ubicaciones: metricsUbic,
                comentarios: { pos: data.comsPos.slice(0,5), neg: data.comsNeg.slice(0,5) },
                nubePos: freq(data.palabrasPos), 
                nubeNeg: freq(data.palabrasNeg),
                satAnual: mesesConDato > 0 ? (sumaSat / mesesConDato).toFixed(1) : "0.0",
                infoHora: { hora: hCritica, porcentaje: porcNegCritica, volumenNeg: maxNegVol }
            };
        });

        res.json({ success: true, data: { sectores: final } });
    } catch (e) { 
        console.error("Error Server:", e);
        res.status(500).json({ success: false, message: e.message }); 
    }
});

app.listen(PORT, () => console.log(`Server ON port ${PORT}`));
