const API_URL = 'https://backend-satisfaccion.onrender.com/procesar-anual'; 
const LOGO = 'logo.png'; 
let datosGlobales = null;

document.addEventListener('DOMContentLoaded', () => {
    const uploadForm = document.getElementById('uploadForm');
    const reporteContainer = document.getElementById('reporte-container');
    const actionsDiv = document.getElementById('actions');

    Chart.register(ChartDataLabels);

    uploadForm.addEventListener('submit', async (e) => {
        e.preventDefault();
        document.getElementById('loader').style.display = 'block';
        reporteContainer.innerHTML = '';
        if(actionsDiv) actionsDiv.style.display = 'none';

        const formData = new FormData();
        formData.append('archivoExcel', document.getElementById('archivoExcel').files[0]);
        formData.append('datosManuales', JSON.stringify({
            enero: { 
                total: parseInt(document.getElementById('ene_total').value)||0,
                muy_positivas: parseInt(document.getElementById('ene_mp').value)||0,
                muy_negativas: parseInt(document.getElementById('ene_mn').value)||0,
                negativas: parseInt(document.getElementById('ene_n').value)||0
            },
            febrero: { 
                total: parseInt(document.getElementById('feb_total').value)||0,
                muy_positivas: parseInt(document.getElementById('feb_mp').value)||0,
                muy_negativas: parseInt(document.getElementById('feb_mn').value)||0,
                negativas: parseInt(document.getElementById('feb_n').value)||0
            }
        }));

        try {
            const response = await fetch(API_URL, { method: 'POST', body: formData });
            const result = await response.json();
            
            if (result.success && result.data.sectores.length > 0) {
                datosGlobales = result.data.sectores;
                renderizar(datosGlobales);
                document.getElementById('actions').style.display = 'flex';
            } else {
                alert("No se encontraron datos. Verifique que su Excel tenga columnas llamadas: Fecha, Respuestas, Muy Positivas, Muy Negativas y Negativas.");
            }
        } catch (err) { alert("Error: " + err.message); }
        finally { document.getElementById('loader').style.display = 'none'; }
    });

    function renderizar(sectores) {
        let html = `<div class="page cover-page" style="display:flex; flex-direction:column; justify-content:center; align-items:center; text-align:center;">
            <img src="${LOGO}" style="width:250px; margin-bottom:40px;">
            <h1 style="font-size:38px; color:#004d40; border:none !important;">ANÁLISIS DE SATISFACCIÓN 2025</h1>
            <h3>Hipódromo de Palermo</h3>
            <p>Enero - Diciembre 2025</p>
        </div>`;

        sectores.forEach((s, idx) => {
            html += `<div class="page">
                <div class="header-strip"><h2>${s.nombre} - Satisfacción 2025</h2><img src="${LOGO}"></div>
                <div class="chart-box" style="height:450px;"><canvas id="chart-${idx}"></canvas></div>
                <div style="background:#e0f2f1; padding:20px; border-radius:10px; border-left:6px solid #004d40;">
                    Índice de Satisfacción Anual: <strong>${s.satAnual}</strong>
                </div>
            </div>
            <div class="page">
                <div class="header-strip"><h2>${s.nombre} - Análisis Cualitativo 2025</h2><img src="${LOGO}"></div>
                <div class="two-columns">
                    <div class="card pos"><h4>Puntos Fuertes</h4>${s.comentarios.pos.map(c => `<div class="comment-item"><small>${c.meta}</small><p>"${c.texto}"</p></div>`).join('')}</div>
                    <div class="card neg"><h4>Oportunidades de Mejora</h4>${s.comentarios.neg.map(c => `<div class="comment-item"><small>${c.meta}</small><p>"${c.texto}"</p></div>`).join('')}</div>
                </div>
            </div>`;
        });

        reporteContainer.innerHTML = html;

        setTimeout(() => {
            sectores.forEach((s, idx) => {
                const ctx = document.getElementById(`chart-${idx}`);
                if(!ctx) return;
                new Chart(ctx, {
                    data: {
                        labels: s.meses.map(m => m.nombre),
                        datasets: [
                            { type: 'line', label: 'Satisfacción', data: s.meses.map(m => m.sat), borderColor: '#004d40', borderWidth: 3, yAxisID: 'ySat', datalabels: { display: true, align: 'top', font: {weight:'bold'} } },
                            { type: 'bar', label: 'Volumen', data: s.meses.map(m => m.total), backgroundColor: 'rgba(0,0,0,0.05)', yAxisID: 'yVol', datalabels: { display: false } }
                        ]
                    },
                    options: { responsive: true, maintainAspectRatio: false, scales: { ySat: { min: -100, max: 100 }, yVol: { position: 'right', grid: {drawOnChartArea: false} } } }
                });
            });
        }, 800);
    }

    document.getElementById('downloadPptx').addEventListener('click', () => {
        if(!datosGlobales) return;
        const pptx = new PptxGenJS();
        pptx.layout = 'LAYOUT_WIDE';
        datosGlobales.forEach(s => {
            let slide = pptx.addSlide();
            slide.addText(`${s.nombre} - Informe 2025`, { x:0.5, y:0.5, fontSize:22, color:'004d40', bold:true });
            slide.addText(`Satisfacción Anual: ${s.satAnual}`, { x:0.5, y:1.1, fontSize:14 });
            slide.addText("Puntos Fuertes", { x:0.5, y:1.8, color:'2e7d32', bold:true });
            s.comentarios.pos.forEach((c, i) => slide.addText(`- ${c.texto.substring(0,120)}...`, { x:0.5, y:2.3+(i*0.6), fontSize:10 }));
        });
        pptx.writeFile({ fileName: 'Informe_Satisfaccion_2025.pptx' });
    });

    document.getElementById('downloadPdf').addEventListener('click', () => window.print());
});
