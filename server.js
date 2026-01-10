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

// FUNCIÓN CLAVE: Convierte fechas DD/MM/YYYY a objeto Date real
function parsearFecha(valor) {
    if (!valor) return null;
    
    // 1. Si Excel ya lo detectó como fecha
    if (valor instanceof Date) return valor;

    // 2. Si es texto tipo "18/04/2025" o "18-04-2025"
    if (typeof valor === 'string') {
        // Eliminar horas si vienen pegadas "18/04/2025 10:00"
        let soloFecha = valor.split(' ')[0]; 
        
        // Detectar separador
        const separador = soloFecha.includes('/') ? '/' : (soloFecha.includes('-') ? '-' : null);
        
        if (separador) {
            const partes = soloFecha.split(separador);
            // Asumimos formato DD/MM/YYYY
            if (partes.length === 3) {
                const dia = parseInt(partes[0]);
                const mes = parseInt(partes[1]) - 1; // Meses en JS son 0-11
                const anio = parseInt(partes[2]);
                
                const fechaGenerada = new Date(anio, mes, dia);
                if (!isNaN(fechaGenerada.getTime())) return fechaGenerada;
            }
        }
        // Intento final estándar
        const intento = new Date(valor);
        return isNaN(intento.getTime()) ? null : intento;
    }
    
    return null;
}

const CONTEXTO_SECTORES = {
    "atencion": ["atencion", "trato", "amable", "ayudo", "atenta", "chica", "chico", "resolvio", "espera", "tarjeta", "fun", "stand", "personal", "empleado", "gente", "recepcion"],
    "gastro": ["comida", "mozo", "frio", "caliente", "rico", "bebida", "mesa", "pedido", "tardo", "restaurant", "confiteria", "hamburguesa", "menu", "sabor", "cafe", "barra"],
    "cajas": ["pago", "cobro", "fila", "rapido", "dinero", "efectivo", "tarjeta", "atencion", "espera", "cajero", "ticket", "cobrar", "caja"],
    "traslados": ["chofer", "auto", "camioneta", "viaje", "llego", "tarde", "limpio", "conduccion", "carrito", "transporte", "valet", "bus"],
    "baños": ["baño", "limpieza", "olor", "jabon", "papel", "sucio", "higienico", "sanitario", "agua", "inodoro"]
};

const PROHIBIDAS_POR_RUIDO = ["maquina", "paga", "premio", "suerte", "ruleta", "slot", "ganar", "perder", "apuesta", "pozo"];

function esComentarioValido(texto, sector) {
    if (!texto || texto.toString().length < 5) return false; // Bajé el límite de caracteres
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

app.post('/procesar-anual', upload.single('archivoExcel'), async (req, res) => {
    try {
        if (!req.file) return res.status(400).json({ success: false, message: 'Archivo no recibido' });
        
        let manual = {};
        try { manual = JSON.parse(req.body.datosManuales || '{}'); } catch(e) {}

        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(req.file.buffer);
        const worksheet = workbook.worksheets[0];

        console.log(`📂 Procesando hoja: ${worksheet.name}`);

        const colMap = { fecha: null, hora: null, sector: null, ubicacion: null, comentario: null, rating: null };
        let headerRowIndex = 1;

        // 1. BUSCAR ENCABEZADOS (Miramos filas 1 a 5)
        for(let r = 1; r <= 5; r++) {
            const row = worksheet.getRow(r);
            row.eachCell((cell, colNumber) => {
                const val = normalizarTexto(cell.value);
                if (val.includes('fecha')) colMap.fecha = colNumber;
                if (val.includes('hora')) colMap.hora = colNumber;
                if (val.includes('sector')) colMap.sector = colNumber;
                if (val.includes('ubicacion')) colMap.ubicacion = colNumber;
                if (val.includes('comentario')) colMap.comentario = colNumber;
                if (val.includes('calificacion') || val === 'nota' || val.includes('puntos')) colMap.rating = colNumber;
            });
            if (colMap.fecha && colMap.rating) {
                headerRowIndex = r;
                break;
            }
        }

        console.log("📍 Columnas detectadas:", colMap);

        if (!colMap.fecha || !colMap.rating) {
            return res.json({ success: false, message: 'No se encontraron las columnas Fecha y Calificación.' });
        }

        const sectores = {};
        let filasProcesadas = 0;
        let filasErrorFecha = 0;

        // 2. ITERAR FILAS
        worksheet.eachRow((row, rowNum) => {
            if (rowNum <= headerRowIndex) return;

            // RATING
            const ratingVal = row.getCell(colMap.rating)?.value;
            let rating = parseInt(ratingVal);
            if (typeof ratingVal === 'object' && ratingVal?.result) rating = parseInt(ratingVal.result);
            if (isNaN(rating)) return;

            // FECHA (Usamos la nueva función parsearFecha)
            const dateVal = row.getCell(colMap.fecha)?.value;
            let date = parsearFecha(dateVal);
            
            if (!date) {
                filasErrorFecha++;
                // console.log(`Fila ${rowNum} descartada: Fecha inválida (${dateVal})`);
                return;
            }
            
            const mIdx = date.getMonth(); // 0-11

            // HORA
            let horaReal = 12; 
            const hVal = colMap.hora ? row.getCell(colMap.hora)?.value : null;
            
            if (hVal instanceof Date) {
                horaReal = hVal.getHours(); // Si Excel lo detecta como fecha/hora
            } else if (typeof hVal === 'number') {
                // Fracción decimal Excel (0.5 = 12hs)
                horaReal = Math.floor(hVal * 24);
            } else if (typeof hVal === 'string') {
                // Formato texto "23:59:55"
                const partesHora = hVal.trim().split(':');
                if (partesHora.length >= 1) {
                    horaReal = parseInt(partesHora[0]);
                }
            }
            // Asegurar límites
            if (horaReal < 0) horaReal = 0;
            if (horaReal > 23) horaReal = 23;

            // TEXTOS
            const sectorName = colMap.sector ? (row.getCell(colMap.sector)?.value || 'General').toString().trim() : 'General';
            const ubicName = colMap.ubicacion ? (row.getCell(colMap.ubicacion)?.value || 'General').toString().trim() : 'General';
            const comment = colMap.comentario ? (row.getCell(colMap.comentario)?.value || '').toString().trim() : '';

            // INICIALIZAR ESTRUCTURA
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

            // ACUMULAR
            s.meses[mIdx].total++;
            s.statsHoras[horaReal].total++;

            if (!s.ubicaciones[ubicName]) s.ubicaciones[ubicName] = { mp:0, p:0, n:0, mn:0, total:0 };
            s.ubicaciones[ubicName].total++;

            // CLASIFICAR
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

            // COMENTARIOS
            if (esComentarioValido(comment, sectorName)) {
                const info = { texto: comment, meta: `${date.getDate()}/${mIdx+1} ${horaReal}:00hs` };
                if (rating >= 3) { s.comsPos.push(info); s.palabrasPos.push(...extractWords(comment)); }
                else { s.comsNeg.push(info); s.palabrasNeg.push(...extractWords(comment)); }
            }
        });

        console.log(`✅ Filas procesadas OK: ${filasProcesadas}`);
        if(filasErrorFecha > 0) console.log(`⚠️ Filas con fecha inválida: ${filasErrorFecha}`);

        if (filasProcesadas === 0) {
            return res.json({ success: false, message: 'Archivo leído pero sin filas válidas. Probable error en formato de fechas (DD/MM/YYYY) o calificaciones no numéricas.' });
        }

        // 3. FINALIZAR CÁLCULOS
        const final = Object.entries(sectores).map(([nombre, data]) => {
            ['enero', 'febrero'].forEach((m, i) => { if (manual[m] && manual[m].total > 0) data.meses[i] = manual[m]; });

            let sumaSat = 0, mesesConDato = 0;
            const mesesFinal = data.meses.map((m) => {
                let val = 0;
                if(m.total > 0) val = ((m.mp + m.p) / m.total) * 100;
                if (m.total > 0) { sumaSat += val; mesesConDato++; }
                return { sat: parseFloat(val.toFixed(1)), total: m.total };
            });

            // Hora Crítica por Volumen Negativo
            let hCritica = "12:00"; let maxNegVol = -1; let volNegTotal = 0; let totalEnEsaHora = 0;
            data.statsHoras.forEach((h, i) => {
                if (h.neg > maxNegVol) {
                    maxNegVol = h.neg; volNegTotal = h.neg; totalEnEsaHora = h.total;
                    hCritica = i.toString().padStart(2, '0') + ':00';
                }
            });
            let porcNegCritica = totalEnEsaHora > 0 ? ((volNegTotal / totalEnEsaHora) * 100).toFixed(1) : "0.0";

            const metricsUbic = Object.entries(data.ubicaciones).map(([uNom, uD]) => {
                const uSat = uD.total > 0 ? (((uD.mp + uD.p) / uD.total) * 100).toFixed(1) : "0.0";
                return { nombre: uNom, totalAnual: uD.total, satProm: uSat, promDiario: (uD.total / 365).toFixed(2) };
            }).sort((a,b) => b.totalAnual - a.totalAnual);

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
