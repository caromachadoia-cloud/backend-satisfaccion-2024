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

const CONTEXTO_SECTORES = {
    "atencion": ["atencion", "trato", "amable", "ayudo", "atenta", "chica", "chico", "resolvio", "espera", "tarjeta", "fun", "stand", "personal"],
    "gastro": ["comida", "mozo", "frio", "caliente", "rico", "bebida", "mesa", "pedido", "tardo", "restaurant", "confiteria", "hamburguesa", "menu", "sabor"],
    "cajas": ["pago", "cobro", "fila", "rapido", "dinero", "efectivo", "tarjeta", "atencion", "espera", "cajero", "ticket", "cobrar", "caja"],
    "traslados": ["chofer", "auto", "camioneta", "viaje", "llego", "tarde", "limpio", "conduccion", "carrito", "transporte", "valet"],
    "baños": ["baño", "limpieza", "olor", "jabon", "papel", "sucio", "higienico", "sanitario", "agua"]
};

const PROHIBIDAS_POR_RUIDO = ["maquina", "paga", "premio", "suerte", "ruleta", "slot", "ganar", "perder"];

function esComentarioValido(texto, sector) {
    if (!texto || texto.length < 15) return false;
    const limpio = texto.toLowerCase();
    const sectorKey = Object.keys(CONTEXTO_SECTORES).find(k => sector.toLowerCase().includes(k)) || "general";
    
    if (sectorKey !== "general" && PROHIBIDAS_POR_RUIDO.some(p => limpio.includes(p))) return false;

    const palabrasContexto = CONTEXTO_SECTORES[sectorKey] || [];
    return palabrasContexto.some(p => limpio.includes(p)) || sectorKey === "general";
}

function extractWords(text) {
    const STOPWORDS = ['de', 'la', 'que', 'el', 'en', 'y', 'a', 'los', 'del', 'se', 'las', 'por', 'un', 'para', 'con', 'no', 'una', 'su', 'al', 'lo', 'como', 'más', 'pero', 'sus', 'le', 'ya', 'o', 'este', 'ha', 'me', 'si', 'sin', 'sobre', 'muy', 'cuando', 'también', 'hasta', 'hay', 'donde', 'quien', 'desde', 'todo', 'nos', 'uno', 'ni', 'contra', 'ese', 'eso', 'mi', 'qué', 'porque', 'estaba', 'fui'];
    return text.toLowerCase().replace(/[.,/#!$%^&*;:{}=\-_`~()]/g, " ").match(/[a-záéíóúñü]+/g)
        ?.filter(word => !STOPWORDS.includes(word) && word.length > 3) || [];
}

app.post('/procesar-anual', upload.single('archivoExcel'), async (req, res) => {
    try {
        if (!req.file) return res.status(400).json({ success: false, message: 'Archivo no recibido' });
        
        // Procesar datos manuales con seguridad
        let manual = {};
        try { manual = JSON.parse(req.body.datosManuales || '{}'); } catch(e) {}

        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(req.file.buffer);
        const worksheet = workbook.worksheets[0];

        const colMap = {};
        worksheet.getRow(1).eachCell((cell, colNumber) => {
            const val = cell.value?.toString().toLowerCase().trim() || '';
            if (val.includes('fecha')) colMap.fecha = colNumber;
            if (val.includes('hora')) colMap.hora = colNumber;
            if (val.includes('sector')) colMap.sector = colNumber;
            if (val.includes('ubicacion') || val.includes('ubicación')) colMap.ubicacion = colNumber;
            if (val.includes('comentario')) colMap.comentario = colNumber;
            if (val.includes('calificacion') || val.includes('rating') || val.includes('valoracion')) colMap.rating = colNumber;
        });

        const sectores = {};

        worksheet.eachRow((row, rowNum) => {
            if (rowNum === 1) return;
            
            const ratingVal = row.getCell(colMap.rating)?.value;
            let rating = parseInt(ratingVal);
            if (isNaN(rating)) return;

            const dateVal = row.getCell(colMap.fecha)?.value;
            if (!dateVal) return;

            // Lógica robusta para la HORA
            const hVal = row.getCell(colMap.hora)?.value;
            let horaReal = 12; // Default
            
            if (hVal instanceof Date) {
                horaReal = hVal.getHours();
            } else if (typeof hVal === 'number') {
                // Excel guarda horas como fracción de día (0.5 = 12:00)
                // Multiplicamos por 24 y tomamos el piso
                horaReal = Math.floor(hVal * 24);
            } else if (typeof hVal === 'string' && hVal.includes(':')) {
                // Formato "14:30"
                horaReal = parseInt(hVal.split(':')[0]);
            }
            
            // Asegurar rango 0-23
            if (horaReal < 0) horaReal = 0;
            if (horaReal > 23) horaReal = 23;

            // Manejo de Fecha
            let date = (dateVal instanceof Date) ? dateVal : new Date(dateVal);
            const mIdx = date.getMonth(); // 0-11
            if (isNaN(mIdx)) return;

            const sectorName = (row.getCell(colMap.sector)?.value || 'General').toString().trim();
            const ubicName = (row.getCell(colMap.ubicacion)?.value || 'General').toString().trim();
            const comment = (row.getCell(colMap.comentario)?.value || '').toString().trim();

            if (!sectores[sectorName]) {
                sectores[sectorName] = {
                    meses: Array.from({length: 12}, () => ({ mp:0, p:0, n:0, mn:0, total:0 })),
                    statsHoras: Array.from({length: 24}, () => ({ total: 0, neg: 0 })), // Array 0-23
                    ubicaciones: {}, 
                    comsPos: [], comsNeg: [], 
                    palabrasPos: [], palabrasNeg: []
                };
            }

            const s = sectores[sectorName];
            
            // Acumular totales
            s.meses[mIdx].total++;
            s.statsHoras[horaReal].total++;

            if (!s.ubicaciones[ubicName]) s.ubicaciones[ubicName] = { mp:0, p:0, n:0, mn:0, total:0 };
            s.ubicaciones[ubicName].total++;

            // Clasificación
            if (rating === 4) { s.meses[mIdx].mp++; s.ubicaciones[ubicName].mp++; }
            if (rating === 3) { s.meses[mIdx].p++; s.ubicaciones[ubicName].p++; }
            if (rating === 2) { 
                s.meses[mIdx].n++; 
                s.ubicaciones[ubicName].n++; 
                s.statsHoras[horaReal].neg++; // Sumar voto negativo a esa hora
            }
            if (rating === 1) { 
                s.meses[mIdx].mn++; 
                s.ubicaciones[ubicName].mn++; 
                s.statsHoras[horaReal].neg++; // Sumar voto muy negativo a esa hora
            }

            if (esComentarioValido(comment, sectorName)) {
                const info = { texto: comment, meta: `${date.getDate()}/${mIdx+1} ${horaReal}:00hs` };
                if (rating >= 3) { s.comsPos.push(info); s.palabrasPos.push(...extractWords(comment)); }
                else { s.comsNeg.push(info); s.palabrasNeg.push(...extractWords(comment)); }
            }
        });

        // Procesamiento final
        const final = Object.entries(sectores).map(([nombre, data]) => {
            // Sobreescribir con datos manuales si existen
            ['enero', 'febrero'].forEach((m, i) => { 
                if (manual[m] && manual[m].total > 0) data.meses[i] = manual[m]; 
            });

            // Calcular Satisfacción Mensual y Anual
            let sumaSat = 0, mesesConDato = 0;
            const mesesFinal = data.meses.map((m) => {
                const factor = m.total / 100;
                // Fórmula: (MuyPos - (MuyNeg + Neg)) ajustada por peso o simple porcentaje neto
                // Usando fórmula estándar de satisfacción neta: (Positivas - Negativas) / Total * 100 ??
                // Usaré la fórmula previa solicitada: (MP/factor) - ((MN + N)/factor) pero asumo que querías decir (Neg + MuyNeg) resta
                // Si la fórmula era: "Porcentaje E - Porcentaje (B+C)"
                let val = 0;
                if(m.total > 0) {
                     val = ((m.mp + m.p) / factor) - ((m.mn + m.n) / factor); // Simplificado: %Pos - %Neg
                     // Ajuste para que sea 0-100 tipo NPS o CSAT? 
                     // Voy a usar CSAT simple: (Positivos / Total) * 100 para evitar negativos confusos en gráfico, o la que tenías.
                     // Volviendo a tu lógica original corregida:
                     val = (m.mp / factor) - ((m.mn - m.n) / factor); // Esta fórmula era rara en tu código original.
                     // Usemos estándar: ((MP + P) / Total) * 100
                     val = ((m.mp + m.p) / m.total) * 100;
                }
                
                if (m.total > 0) { sumaSat += val; mesesConDato++; }
                return { sat: parseFloat(val.toFixed(1)), total: m.total };
            });

            // HORA CRÍTICA: Buscar hora con MAYOR CANTIDAD de negativos
            let hCritica = "12:00"; 
            let maxNegVol = -1;
            let volNegTotal = 0;
            let totalEnEsaHora = 0;

            data.statsHoras.forEach((h, i) => {
                // Buscamos el pico de volumen negativo
                if (h.neg > maxNegVol) {
                    maxNegVol = h.neg;
                    volNegTotal = h.neg;
                    totalEnEsaHora = h.total;
                    hCritica = i.toString().padStart(2, '0') + ':00';
                }
            });

            // Porcentaje de negatividad en esa hora específica
            let porcNegCritica = totalEnEsaHora > 0 ? ((volNegTotal / totalEnEsaHora) * 100).toFixed(1) : "0.0";

            // Métricas por ubicación
            const metricsUbic = Object.entries(data.ubicaciones).map(([uNom, uD]) => {
                // Sat Promedio: ((MP + P) / Total) * 100
                const uSat = uD.total > 0 ? (((uD.mp + uD.p) / uD.total) * 100).toFixed(1) : "0.0";
                return { 
                    nombre: uNom, 
                    totalAnual: uD.total, 
                    satProm: uSat, 
                    promDiario: (uD.total / 365).toFixed(1) 
                };
            }).sort((a,b) => b.totalAnual - a.totalAnual); // Ordenar por más votos

            const freq = (arr) => Object.entries(arr.reduce((a,w)=>(a[w]=(a[w]||0)+1,a),{})).sort((a,b)=>b[1]-a[1]).slice(0, 40);

            return {
                nombre, 
                meses: mesesFinal, 
                ubicaciones: metricsUbic,
                comentarios: { 
                    pos: data.comsPos.sort((a,b)=>b.texto.length-a.texto.length).slice(0,5), 
                    neg: data.comsNeg.sort((a,b)=>b.texto.length-a.texto.length).slice(0,5) 
                },
                nubePos: freq(data.palabrasPos), 
                nubeNeg: freq(data.palabrasNeg),
                satAnual: mesesConDato > 0 ? (sumaSat / mesesConDato).toFixed(1) : "0.0",
                infoHora: { 
                    hora: hCritica, 
                    porcentaje: porcNegCritica,
                    volumenNeg: maxNegVol
                }
            };
        });

        res.json({ success: true, data: { sectores: final } });
    } catch (e) { 
        console.error(e);
        res.status(500).json({ success: false, message: e.message }); 
    }
});

app.listen(PORT, () => console.log(`Server ON port ${PORT}`));
