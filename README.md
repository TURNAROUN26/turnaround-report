[index.html](https://github.com/user-attachments/files/26665204/index.html)
# turnaround-report
TurnAround Report - Planta PERU LNG
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0, user-scalable=yes">
    <title>TurnAround Report - Planta PERU LNG | Colaboración en Nube</title>
    <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
    <script src="https://cdn.sheetjs.com/xlsx-0.20.2/package/dist/xlsx.full.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/qrcodejs@1.0.0/qrcode.min.js"></script>
    <style>
        * { box-sizing: border-box; }
        body {
            margin: 0;
            padding: 20px;
            font-family: 'Segoe UI', Roboto, sans-serif;
            background: #f5f7fc;
            color: #1e2a3a;
        }
        .container {
            max-width: 1600px;
            margin: auto;
            background: #ffffff;
            border-radius: 28px;
            padding: 20px 24px;
            box-shadow: 0 12px 30px rgba(0, 0, 0, 0.08);
            border: 1px solid #e2e8f0;
        }
        h1 {
            text-align: center;
            font-size: 1.9rem;
            margin: 0 0 6px;
            background: linear-gradient(135deg, #c96f0e, #e6a157);
            -webkit-background-clip: text;
            background-clip: text;
            color: transparent;
        }
        .sub { text-align: center; color: #4a5b6e; margin-bottom: 20px; }
        .sync-bar {
            background: #eef2ff;
            border-radius: 40px;
            padding: 10px 20px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 20px;
            flex-wrap: wrap;
            gap: 10px;
        }
        button {
            background-color: #e68a2e;
            cursor: pointer;
            transition: 0.2s;
            font-weight: 500;
            color: white;
            border: none;
            padding: 8px 18px;
            border-radius: 40px;
        }
        button:hover { background-color: #c96f0e; transform: translateY(-1px); }
        .btn-sync { background-color: #27ae60; }
        .btn-cloud { background-color: #2980b9; }
        .btn-excel { background-color: #1f7b4d; }
        .btn-force { background-color: #c0392b; }
        .chart-row { display: flex; flex-wrap: wrap; gap: 24px; margin-bottom: 30px; }
        .chart-card {
            flex: 1;
            min-width: 300px;
            background: #ffffff;
            border-radius: 24px;
            padding: 12px;
            border: 1px solid #e2edf2;
            box-shadow: 0 4px 12px rgba(0,0,0,0.03);
        }
        .chart-card h3 { margin: 0 0 10px; font-size: 1.2rem; text-align: center; color: #c96f0e; }
        canvas { width: 100%; max-height: 380px; background: #ffffff; border-radius: 16px; }
        .flex-row { display: flex; gap: 24px; flex-wrap: wrap; margin-top: 25px; }
        .card {
            background: #fefefe;
            padding: 18px;
            border-radius: 24px;
            border: 1px solid #e2e8f0;
            box-shadow: 0 2px 6px rgba(0,0,0,0.03);
            flex: 1;
            min-width: 280px;
        }
        label { font-weight: 600; display: block; margin-top: 12px; margin-bottom: 6px; color: #b45f06; }
        input, select, textarea {
            padding: 10px 14px;
            font-size: 0.9rem;
            border-radius: 40px;
            border: 1px solid #cbd5e1;
            margin-right: 8px;
            margin-bottom: 8px;
            background: #ffffff;
            width: calc(100% - 20px);
        }
        textarea { border-radius: 16px; resize: vertical; min-height: 150px; }
        .proyectos-mosaicos { display: flex; flex-wrap: wrap; gap: 12px; margin-bottom: 20px; justify-content: center; }
        .proyecto-mosaico {
            background: #eef2ff;
            border-radius: 40px;
            padding: 8px 18px;
            font-size: 0.85rem;
            font-weight: bold;
            cursor: pointer;
            transition: all 0.2s;
            border: 2px solid #cbd5e1;
            display: inline-flex;
            align-items: center;
            gap: 6px;
            color: #1e293b;
        }
        .proyecto-mosaico.active { background: #f39c12; color: #ffffff; border-color: #f39c12; }
        .proyecto-mosaico.has-comment.active { background: #f39c12; color: #ffffff; border-color: #f39c12; }
        .proyecto-mosaico.has-comment { border-color: #f1c40f; box-shadow: 0 0 6px rgba(241,196,15,0.6); background: #fff8e7; }
        .categorias-proyecto { display: flex; flex-wrap: wrap; gap: 12px; margin: 15px 0 10px; }
        .cat-btn {
            background: #eef2ff;
            border-radius: 40px;
            padding: 6px 16px;
            font-size: 0.8rem;
            font-weight: bold;
            cursor: pointer;
            border: 1px solid #cbd5e1;
            display: inline-flex;
            align-items: center;
            gap: 6px;
            color: #1e293b;
        }
        .cat-btn.active { background: #f39c12; color: white; border-color: #f39c12; }
        .cat-btn.has-comment.active { background: #f39c12; color: white; border-color: #f39c12; }
        .cat-btn.has-comment { border-color: #f1c40f; box-shadow: 0 0 4px #f1c40f; background: #fff6e0; }
        table { width: 100%; border-collapse: collapse; color: #1e293b; font-size: 0.8rem; }
        th, td { text-align: left; padding: 10px 8px; border-bottom: 1px solid #e2e8f0; }
        th { background-color: #f1f5f9; color: #b45f06; }
        .scrollable-table {
            max-height: 320px;
            overflow-y: auto;
            border-radius: 16px;
            position: relative;
        }
        .scrollable-table table {
            border-collapse: separate;
            border-spacing: 0;
        }
        .scrollable-table th {
            position: sticky;
            top: 0;
            background-color: #f1f5f9;
            z-index: 10;
            box-shadow: 0 1px 0 #e2e8f0;
        }
        .semaphore-table { background: #ffffff; border-radius: 24px; padding: 12px 16px; border: 1px solid #e2edf2; margin-top: 0; width: 100%; }
        .traffic-light { display: inline-block; width: 16px; height: 16px; border-radius: 50%; margin-right: 8px; }
        .green { background-color: #2ecc71; }
        .red { background-color: #e74c3c; }
        .yellow { background-color: #f1c40f; }
        .info-corte { margin-top: 20px; background: #fef9e6; padding: 15px; border-radius: 20px; border-left: 5px solid #f39c12; color: #4a5b6e; }
        .fecha-selector { display: flex; align-items: center; gap: 12px; flex-wrap: wrap; margin-bottom: 16px; background: #f8fafc; padding: 10px 16px; border-radius: 48px; border: 1px solid #e2e8f0; }
        .upload-area { background: #f8fafc; border-radius: 40px; padding: 5px 15px; display: flex; align-items: center; gap: 12px; flex-wrap: wrap; margin-bottom: 15px; border: 1px solid #e2e8f0; }
        .notes-card, .comments-card { background: #ffffff; border-radius: 24px; padding: 18px; margin-top: 20px; border: 1px solid #e2e8f0; }
        .qr-info { color: #2c3e50; font-size: 0.75rem; margin-top: 12px; }
        .warning-badge { background-color: #f39c12; color: #1e2a3a; border-radius: 40px; padding: 4px 12px; font-size: 0.7rem; }
        .reporte-dia-card {
            background: #fef9e6;
            border-radius: 28px;
            padding: 16px 24px;
            margin-bottom: 25px;
            border: 1px solid #f0d9a8;
            box-shadow: 0 4px 10px rgba(0,0,0,0.02);
        }
        .personal-grid {
            display: flex;
            gap: 20px;
            flex-wrap: wrap;
            margin: 15px 0 10px;
        }
        .personal-item {
            flex: 1;
            min-width: 140px;
        }
        .personal-item input {
            width: 100%;
            margin-top: 5px;
        }
        .btn-guardar-personal {
            background-color: #2c7da0;
            margin-top: 8px;
        }
        .comments-card textarea {
            width: 100%;
            min-height: 120px;
            border-radius: 20px;
        }
        .qr-container {
            background: #ffffff;
            border-radius: 24px;
            padding: 20px;
            margin-top: 25px;
            text-align: center;
            border: 1px solid #e2e8f0;
        }
        .qr-wrapper {
            display: inline-block;
            padding: 15px;
            background: white;
            border-radius: 16px;
            border: 1px solid #e2e8f0;
        }
        .btn-qr {
            background-color: #27ae60;
            margin-top: 10px;
        }
        .qr-url-display {
            font-size: 0.7rem;
            color: #64748b;
            margin-top: 10px;
            word-break: break-all;
            background: #f1f5f9;
            padding: 8px;
            border-radius: 12px;
        }
        .qr-note {
            background: #e8f5e9;
            border-radius: 16px;
            padding: 12px;
            margin-top: 15px;
            font-size: 0.75rem;
            color: #2e7d32;
        }
        @media (max-width: 768px) {
            .proyecto-mosaico, .cat-btn { font-size: 0.75rem; padding: 6px 12px; }
            .personal-grid { flex-direction: column; }
        }
        .auto-sync-badge {
            background: #27ae60;
            color: white;
            border-radius: 40px;
            padding: 4px 12px;
            font-size: 0.7rem;
            display: inline-block;
            margin-left: 10px;
        }
        .semaphore-full-width {
            width: 100%;
            margin-bottom: 20px;
        }
        .edit-protection-info {
            background: #eef2fa;
            border-radius: 24px;
            padding: 6px 12px;
            font-size: 0.7rem;
            margin-top: 10px;
            color: #2c3e50;
            display: inline-block;
        }
        .excel-feedback {
            font-weight: bold;
            margin-left: 8px;
        }
        .timer {
            font-family: monospace;
            font-weight: bold;
            background: #2c3e50;
            color: white;
            padding: 4px 12px;
            border-radius: 40px;
            font-size: 0.85rem;
        }
        .unified-date-warning {
            background: #f1c40f20;
            border-radius: 28px;
            padding: 6px 12px;
            font-size: 0.75rem;
            color: #7d5d00;
        }
    </style>
</head>
<body>
<div class="container">
    <h1>📊 TurnAround Report - Planta PERU LNG</h1>
    <div class="sub">📌 Curvas S | Notas por Proyecto | Personal & HSE <span class="auto-sync-badge">⚡ Sincronización automática c/15 min</span></div>
    
    <div class="sync-bar">
        <span id="syncStatus">🔄 Conectando con la nube...</span>
        <div style="display: flex; gap: 8px; align-items: center;">
            <span id="syncTimer" class="timer">⏱️ Próxima sincronización: 15:00</span>
            <button id="btnSync" class="btn-sync">🔄 Sincronizar (fusión)</button>
            <button id="btnForceReload" class="btn-force">⬇️ Forzar recarga completa</button>
            <button id="btnUpload" class="btn-cloud">☁️ Guardar en la nube</button>
        </div>
    </div>

    <div class="upload-area">
        <label style="margin:0;">📂 Cargar archivo Excel (Curva S):</label>
        <input type="file" id="excelUpload" accept=".xlsx, .xls" style="flex:2;">
        <button id="btnCargarExcel" class="btn-excel">🔄 Reemplazar PLAN y REAL en la nube</button>
        <span id="excelStatus" style="font-size:0.75rem; color:#2c5a2e;">(Esperando archivo...)</span>
    </div>
    <div class="qr-info" style="margin-top:-5px; margin-bottom:10px; font-size:0.7rem;">
        📌 Formato requerido: 1ra columna = código proyecto, 2da columna = "PLANNED" o "ACTUAL", desde 3ra columna valores de avance (0 a 1) por fecha.
    </div>

    <!-- FECHA ÚNICA: Controla TODO el reporte -->
    <div class="reporte-dia-card">
        <div class="fecha-selector" style="margin-bottom: 12px; background: #fff4e0; justify-content: space-between;">
            <label style="font-weight: bold; font-size: 1rem;">📅 Fecha ÚNICA del Reporte (controla SEMÁFORO, NOTAS, PERSONAL y COMENTARIOS):</label>
            <select id="fechaUnicaReporte" style="flex: 2; min-width: 240px;"></select>
            <span class="unified-date-warning">🔁 Al cambiar se actualizan TODOS los apartados</span>
        </div>
        
        <h3 style="margin: 8px 0 0 0;">👷 Personal en obra</h3>
        <div class="personal-grid">
            <div class="personal-item">
                <label>👥 Total obra</label>
                <input type="number" id="personalTotal" step="1" value="" placeholder="Ej: 245">
            </div>
            <div class="personal-item">
                <label>🌞 Turno día</label>
                <input type="number" id="personalDia" step="1" value="" placeholder="Ej: 150">
            </div>
            <div class="personal-item">
                <label>🌙 Turno noche</label>
                <input type="number" id="personalNoche" step="1" value="" placeholder="Ej: 95">
            </div>
        </div>
        
        <h3 style="margin: 12px 0 0 0;">⚠️ HSE - Novedad relevante</h3>
        <textarea id="hseNovedad" rows="2" placeholder="Describir cualquier incidente, acto inseguro, o novedad relevante en HSE..."></textarea>
        
        <div style="display: flex; justify-content: flex-end; margin-top: 12px;">
            <button id="btnGuardarPersonalHSE" class="btn-guardar-personal">💾 Guardar Personal / HSE para esta fecha</button>
        </div>
        <div class="qr-info" style="margin-top: 10px;">📌 Los datos de personal y HSE se guardan por cada fecha única y se sincronizan en la nube.</div>
        <div id="personalEditProtectionMsg" class="edit-protection-info">✍️ Mientras edites, la sincronización automática NO borrará tus cambios.</div>
    </div>

    <div class="chart-row">
        <div class="chart-card">
            <h3>📈 Curva S - PORTAFOLIO GENERAL</h3>
            <canvas id="avanceChart" width="800" height="400"></canvas>
        </div>
        <div class="chart-card">
            <h3>📊 Curva S por proyecto</h3>
            <div class="proyecto-selector">
                <label>Seleccionar proyecto:</label>
                <select id="selectProyectoGrafica">
                    <option value="24001">24001 - Delta-V/MD Controller</option>
                    <option value="24005">24005 - Support Replace HVAC</option>
                    <option value="24018">24018 - System Control Flare</option>
                    <option value="25005">25005 - Flame Detector</option>
                    <option value="25011">25011 - Online Monitoring Orbit</option>
                    <option value="24007">24007 - Spring Hangers</option>
                    <option value="25006">25006 - Ground Flare Shield</option>
                    <option value="25010">25010 - UV-114020</option>
                    <option value="25016">25016 - New MCHE valves</option>
                    <option value="25020">25020 - Vibration Support</option>
                    <option value="25024">25024 - Thermal Spray BOG</option>
                    <option value="25031">25031 - Airline MCHE</option>
                </select>
            </div>
            <canvas id="proyectoChart" width="800" height="380"></canvas>
        </div>
    </div>

    <!-- CUADRO SEMÁFORO DE PROYECTOS -->
    <div class="semaphore-full-width">
        <div class="semaphore-table">
            <h3>🚦 Cuadro Semáforo de Proyectos (basado en fecha única)</h3>
            <div style="overflow-x: auto;">
                <table id="tablaSemaforo">
                    <thead>
                        <tr><th>Proyecto</th><th>Plan %</th><th>Real %</th><th>Variación</th><th>SPI</th><th>Estado</th>
                        </tr>
                    </thead>
                    <tbody id="semaforoBody">
                        <tr><td colspan="6">Cargando......</td></tr>
                    </tbody>
                </table>
            </div>
        </div>
    </div>

    <!-- SECCIÓN DE NOTAS / ACTIVIDADES PRINCIPALES -->
    <div class="notes-card">
        <h3>📝 Principales actividades realizadas (fecha única)</h3>
        <div class="proyectos-mosaicos" id="proyectosMosaicos"></div>
        <div id="categoriasProyectoArea" style="margin-top: 12px;"></div>
        <textarea id="comentarioTexto" rows="6" placeholder="Escribe aquí las principales actividades realizadas para este proyecto y categoría..."></textarea>
        <div style="display: flex; gap: 12px; margin-top: 8px;">
            <button id="btnGuardarComentario" style="background:#4a6fa5;">💾 Guardar</button>
            <button id="btnEditarComentario" class="btn-edit">✏️ Editar</button>
        </div>
        <div id="infoNotas" class="info-corte" style="margin-top: 12px; font-size:0.8rem;">✅ Selecciona un proyecto y categoría. Las notas se guardan por la fecha única seleccionada arriba.</div>
    </div>

    <!-- COMENTARIOS GENERALES -->
    <div class="comments-card">
        <h3>💬 Comentarios Generales del día</h3>
        <textarea id="comentariosGenerales" rows="4" placeholder="Escriba aquí comentarios, observaciones, resumen del día, desvíos, etc. (por fecha única)"></textarea>
        <div style="display: flex; justify-content: flex-end; margin-top: 12px;">
            <button id="btnGuardarComentariosGenerales" class="btn-guardar-personal">💾 Guardar Comentarios</button>
        </div>
        <div class="qr-info" style="margin-top: 8px;">📌 Los comentarios generales se guardan para la fecha única seleccionada.</div>
    </div>

    <!-- SECCIÓN QR AL FINAL DEL REPORTE -->
    <div class="qr-container">
        <h3>🔗 QR para Visualización del Reporte</h3>
        <p style="font-size:0.8rem; color:#64748b;">Genera un código QR con la URL actual para compartir la visualización del reporte.</p>
        <div class="qr-wrapper" id="qrCodeDiv">
            <div style="padding: 20px; color: #64748b;">⬅️ Haz clic en "Generar QR"</div>
        </div>
        <div>
            <button id="btnGenerarQR" class="btn-qr">📱 Generar QR del Reporte</button>
            <button id="btnCopiarURL" class="btn-qr" style="background-color:#2980b9; margin-left:10px;">📋 Copiar URL</button>
        </div>
        <div class="qr-url-display" id="qrUrlDisplay"></div>
        <div class="qr-note">
            📌 <strong>Instrucciones:</strong> Escanea el código QR con la cámara de tu celular. Se abrirá automáticamente el reporte completo.<br>
            💡 Si el QR no funciona, copia la URL de arriba y pégala en el navegador de tu celular.
        </div>
    </div>

    <div id="infoCorte" class="info-corte">
        ✅ Los datos se guardan localmente y <strong>se suben automáticamente a la nube</strong> al hacer cualquier cambio.
    </div>
</div>

<script>
    // ========== CONFIGURACIÓN JSONBIN ==========
    const JSONBIN_BIN_ID = "69d2a901856a68218900cf05";
    const JSONBIN_API_KEY = "$2a$10$PtMTY.RZC7wtT0GEsBvc9uMPQ3uWmg9sDTqZCgSgB.vAGILz3NsCm";
    const JSONBIN_URL = `https://api.jsonbin.io/v3/b/${JSONBIN_BIN_ID}`;
    
    // ========== DATOS GLOBALES ==========
    const fechasCorte = [];
    const startDate = new Date(2026, 2, 30, 18, 30, 0);
    for (let i = 0; i < 56; i++) {
        let newDate = new Date(startDate);
        newDate.setHours(startDate.getHours() + (i * 12));
        fechasCorte.push(newDate);
    }
    function formatearFechaCompleta(date) { let d = date.getDate(); let m = date.getMonth() + 1; let h = date.getHours().toString().padStart(2,'0'); let min = date.getMinutes().toString().padStart(2,'0'); return `${d}/${m}/${date.getFullYear()} ${h}:${min}`; }
    function formatearFechaCorta(date) { return `${date.getDate()}/${date.getMonth()+1}`; }
    
    let planes = {};
    const codigosProyecto = ["PORTAFOLIO","24001","24005","24018","25005","25011","24007","25006","25010","25016","25020","25024","25031"];
    for (let cod of codigosProyecto) planes[cod] = new Array(fechasCorte.length).fill(0);
    
    let realData = {};
    let comentarios = {};
    const categorias = ["management", "ingenieria", "procura", "construccion"];
    const nombresCat = { management:"🏢 Management", ingenieria:"⚙️ Ingeniería", procura:"📦 Procura", construccion:"🏗️ Construcción" };
    
    function initReales() { for (let cod of codigosProyecto) realData[cod] = new Array(fechasCorte.length).fill(null); }
    function initComentarios() { for (let cod of codigosProyecto) { comentarios[cod] = []; for (let i = 0; i < fechasCorte.length; i++) comentarios[cod][i] = { management: "", ingenieria: "", procura: "", construccion: "" }; } }
    initReales(); initComentarios();
    
    let infoAdicional = {};
    let comentariosGenerales = {};
    
    const defaultPlanes = { "PORTAFOLIO": [0.0000,0.0000,0.0000,0.0000,0.0103,0.0103,0.0296,0.0296,0.0436,0.0436,0.0582,0.0673,0.1032,0.1094,0.1399,0.1399,0.1700,0.1841,0.2227,0.2422,0.2819,0.2935,0.3338,0.3398,0.3734,0.3904,0.4463,0.4610,0.5017,0.5099,0.5539,0.5725,0.6304,0.6533,0.7015,0.7026,0.7377,0.7377,0.7658,0.7658,0.7938,0.7938,0.8165,0.8165,0.8407,0.8407,0.8697,0.8697,0.9027,0.9027,0.9402,0.9486,0.9736,0.9763,0.9951,0.9951,0.9995,0.9995,1.0000], "24001":[0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0.0779,0.0779,0.1818,0.1818,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1], "24005":[0,0,0,0,0.0763,0.0763,0.1526,0.1526,0.1526,0.1526,0.1526,0.1526,0.1526,0.1526,0.1526,0.1526,0.1526,0.1526,0.1526,0.1526,0.2289,0.2289,0.3051,0.3051,0.378,0.378,0.4543,0.4543,0.5306,0.5306,0.6069,0.6069,0.6832,0.6832,0.756,0.756,0.756,0.756,0.756,0.756,0.7914,0.7914,0.8375,0.8375,0.8876,0.8876,0.9375,0.9375,1,1,1,1,1,1,1,1,1,1,1], "24018":[0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0.0909,0.0909,0.1818,0.1818,0.2727,0.2727,0.3636,0.3636,0.4545,0.4545,0.5455,0.5455,0.6364,0.6364,0.6364,0.6364,0.6364,0.6364,0.6364,0.6364,0.6364,0.6364,0.6364,0.6364,0.8182,0.8182,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1], "25005":[0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0.8908,0.8908,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1], "25011":[0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0.0747,0.0747,0.1547,0.1547,0.2293,0.2293,0.304,0.304,0.3787,0.3787,0.4587,0.4587,0.5333,0.5333,0.6042,0.6042,0.675,0.675,0.7458,0.7458,0.8167,0.8167,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1], "24007":[0,0,0,0,0.0572,0.0572,0.1496,0.1496,0.207,0.207,0.2593,0.2733,0.3372,0.3452,0.3902,0.3902,0.4352,0.4352,0.4971,0.4971,0.5235,0.5235,0.548,0.548,0.5571,0.5571,0.5599,0.5599,0.6113,0.6113,0.6635,0.6635,0.6756,0.6756,0.6779,0.6779,0.6799,0.6799,0.6971,0.6971,0.7516,0.7516,0.7905,0.7905,0.8283,0.8283,0.8682,0.8682,0.9165,0.9165,0.98,0.98,1,1,1,1,1,1,1], "25006":[0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0.0261,0.0872,0.1288,0.2131,0.2929,0.3387,0.3758,0.402,0.4274,0.4731,0.4821,0.5347,0.5659,0.6012,0.6461,0.7263,0.8253,0.9242,0.995,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1], "25010":[0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0.0295,0.0295,0.059,0.059,0.0711,0.0711,0.0833,0.0833,0.0955,0.0955,0.1528,0.1528,0.2047,0.2047,0.2549,0.2549,0.3571,0.3571,0.4615,0.4615,0.5619,0.5619,0.6641,0.6641,0.7645,0.7645,0.8707,0.8707,0.9711,0.9711,0.9855,0.9855,1,1,1,1,1], "25016":[0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0.1168,0.3412,0.4268,0.4551,0.4551,0.4551,0.4551,0.4551,0.4551,0.4551,0.4551,0.4551,0.4551,0.4551,0.4551,0.4551,0.4551,0.4551,0.4551,0.4551,0.4551,0.4551,0.4551,0.4551,0.4551,0.4551,0.4551,0.4551,0.4551,0.4551,0.4551,0.4551,0.4551,0.4551,0.4551,0.4551,0.4551,0.4551,0.4551,0.4551,0.599,0.7425,0.8758,0.9226,0.973,0.973,0.9918,0.9918,1], "25020":[0,0,0,0,0.0062,0.0062,0.0102,0.0102,0.0246,0.0246,0.0401,0.0401,0.0546,0.0546,0.0701,0.0701,0.0846,0.0846,0.1001,0.1001,0.1146,0.1146,0.1474,0.1474,0.1792,0.1792,0.2311,0.2311,0.2907,0.2907,0.3506,0.3506,0.4254,0.4254,0.4989,0.4989,0.5712,0.5712,0.6273,0.6273,0.664,0.664,0.701,0.701,0.7429,0.7429,0.7975,0.7975,0.8593,0.8593,0.906,0.906,0.9455,0.9455,0.9903,0.9903,1,1,1], "25024":[0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0.0753,0.0753,0.1794,0.1794,0.184,0.184,0.2883,0.2883,0.3913,0.4151,0.557,0.557,0.6604,0.8053,0.9419,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1], "25031":[0,0,0,0,0,0,0,0,0,0,0.042,0.042,0.211,0.211,0.475,0.475,0.7265,0.7265,0.8915,0.8915,0.94,0.94,0.9775,0.9775,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1] };
    
    let activeEditingFields = new Set();
    let editingTimeout = null;
    function markEditing(fieldId) { activeEditingFields.add(fieldId); if (editingTimeout) clearTimeout(editingTimeout); editingTimeout = setTimeout(() => { activeEditingFields.clear(); }, 10000); }
    const editableFields = ['personalTotal', 'personalDia', 'personalNoche', 'hseNovedad', 'comentarioTexto', 'comentariosGenerales'];
    editableFields.forEach(id => { const el = document.getElementById(id); if (el) { el.addEventListener('focus', () => markEditing(id)); el.addEventListener('input', () => markEditing(id)); } });
    
    function cargarInfoAdicionalLocal() { const stored = localStorage.getItem("turnaround_info_adicional_v1"); if (stored) { try { infoAdicional = JSON.parse(stored); } catch(e) { infoAdicional = {}; } } else infoAdicional = {}; }
    function guardarInfoAdicionalLocal() { localStorage.setItem("turnaround_info_adicional_v1", JSON.stringify(infoAdicional)); }
    function cargarComentariosGeneralesLocal() { const stored = localStorage.getItem("turnaround_comentarios_generales_v1"); if (stored) { try { comentariosGenerales = JSON.parse(stored); } catch(e) { comentariosGenerales = {}; } } else comentariosGenerales = {}; }
    function guardarComentariosGeneralesLocal() { localStorage.setItem("turnaround_comentarios_generales_v1", JSON.stringify(comentariosGenerales)); }
    function saveRealesToLocal() { localStorage.setItem("turnaround_reales_proyectos_v5", JSON.stringify(realData)); }
    function savePlanesToLocal() { localStorage.setItem("turnaround_planes_proyectos_v1", JSON.stringify(planes)); }
    function saveComentariosToLocal() { localStorage.setItem("turnaround_comentarios_v2", JSON.stringify(comentarios)); }
    function loadFromLocal() { 
        const storedReal=localStorage.getItem("turnaround_reales_proyectos_v5"); if(storedReal) try{ const parsed=JSON.parse(storedReal); for(let cod of codigosProyecto) if(parsed[cod]) realData[cod]=parsed[cod]; }catch(e){} 
        const storedPlanes=localStorage.getItem("turnaround_planes_proyectos_v1"); if(storedPlanes) try{ const parsed=JSON.parse(storedPlanes); for(let cod of codigosProyecto) if(parsed[cod]) planes[cod]=parsed[cod]; else if(defaultPlanes[cod]) planes[cod]=[...defaultPlanes[cod]]; }catch(e){ cargarPlanesDefault(); }
        else cargarPlanesDefault();
        const storedCom=localStorage.getItem("turnaround_comentarios_v2"); if(storedCom) try{ const parsed=JSON.parse(storedCom); for(let cod of codigosProyecto) if(parsed[cod]) comentarios[cod]=parsed[cod]; }catch(e){} 
        cargarInfoAdicionalLocal(); cargarComentariosGeneralesLocal(); 
    }
    function cargarPlanesDefault() { for(let cod of codigosProyecto) if(defaultPlanes[cod]) planes[cod]=[...defaultPlanes[cod]]; else planes[cod]=new Array(fechasCorte.length).fill(0); }
    
    let isSyncing = false;
    async function fetchFromCloud() { 
        const response = await fetch(JSONBIN_URL, { 
            headers: { 
                "X-Master-Key": JSONBIN_API_KEY,
                "X-Bin-Meta": "false"
            } 
        }); 
        if (!response.ok) throw new Error(`HTTP ${response.status}`); 
        return await response.json(); 
    }
    
    function mergeRemoteToLocal(remote) { if (!remote) return; 
        for (let cod of codigosProyecto) { 
            if (remote.planes && remote.planes[cod]) for (let i=0; i<fechasCorte.length; i++) if (planes[cod][i] === 0 && remote.planes[cod][i] > 0) planes[cod][i] = remote.planes[cod][i]; 
            if (remote.realData && remote.realData[cod]) for (let i=0; i<fechasCorte.length; i++) if (realData[cod][i] === null && remote.realData[cod][i] !== null) realData[cod][i] = remote.realData[cod][i]; 
            if (remote.comentarios && remote.comentarios[cod]) { if (!comentarios[cod]) comentarios[cod] = []; for (let i=0; i<fechasCorte.length; i++) { if (!comentarios[cod][i]) comentarios[cod][i] = { management:"", ingenieria:"", procura:"", construccion:"" }; for (let cat of categorias) if (remote.comentarios[cod][i] && remote.comentarios[cod][i][cat] && !comentarios[cod][i][cat]) comentarios[cod][i][cat] = remote.comentarios[cod][i][cat]; } } 
        } 
        if (remote.infoAdicional && !activeEditingFields.has('personalTotal') && !activeEditingFields.has('personalDia') && !activeEditingFields.has('personalNoche') && !activeEditingFields.has('hseNovedad')) for (let key in remote.infoAdicional) { if (!infoAdicional[key]) infoAdicional[key] = {}; for (let prop of ['personalTotal','personalDia','personalNoche','hseNovedad']) if (remote.infoAdicional[key][prop] !== undefined && (infoAdicional[key][prop] === undefined || infoAdicional[key][prop] === null || infoAdicional[key][prop] === "")) infoAdicional[key][prop] = remote.infoAdicional[key][prop]; } 
        if (remote.comentariosGenerales && !activeEditingFields.has('comentariosGenerales')) for (let key in remote.comentariosGenerales) if (remote.comentariosGenerales[key] && (!comentariosGenerales[key] || comentariosGenerales[key] === "")) comentariosGenerales[key] = remote.comentariosGenerales[key]; 
        saveRealesToLocal(); savePlanesToLocal(); saveComentariosToLocal(); guardarInfoAdicionalLocal(); guardarComentariosGeneralesLocal(); 
    }
    
    async function guardarEnNubeConMerge() { 
        if (isSyncing) return; 
        isSyncing = true; 
        const statusDiv = document.getElementById("syncStatus"); 
        statusDiv.innerHTML = "☁️ Sincronizando..."; 
        try { 
            let remoteRecord = null; 
            try { remoteRecord = await fetchFromCloud(); } catch(e) { console.warn("No se pudo obtener remoto"); } 
            const localPayload = { planes, realData, comentarios, infoAdicional, comentariosGenerales, timestamp: new Date().toISOString() }; 
            if (remoteRecord) { 
                for (let cod of codigosProyecto) { 
                    if (remoteRecord.planes && remoteRecord.planes[cod]) for (let i=0; i<fechasCorte.length; i++) if (planes[cod][i] === 0 && remoteRecord.planes[cod][i] > 0) planes[cod][i] = remoteRecord.planes[cod][i]; 
                    if (remoteRecord.realData && remoteRecord.realData[cod]) for (let i=0; i<fechasCorte.length; i++) if (realData[cod][i] === null && remoteRecord.realData[cod][i] !== null) realData[cod][i] = remoteRecord.realData[cod][i]; 
                    if (remoteRecord.comentarios && remoteRecord.comentarios[cod]) { if (!comentarios[cod]) comentarios[cod] = []; for (let i=0; i<fechasCorte.length; i++) { if (!comentarios[cod][i]) comentarios[cod][i] = { management:"", ingenieria:"", procura:"", construccion:"" }; for (let cat of categorias) if (remoteRecord.comentarios[cod][i] && remoteRecord.comentarios[cod][i][cat] && !comentarios[cod][i][cat]) comentarios[cod][i][cat] = remoteRecord.comentarios[cod][i][cat]; } } 
                } 
                if (remoteRecord.infoAdicional) for (let key in remoteRecord.infoAdicional) { if (!infoAdicional[key]) infoAdicional[key] = {}; for (let prop of ['personalTotal','personalDia','personalNoche','hseNovedad']) if (remoteRecord.infoAdicional[key][prop] !== undefined && (infoAdicional[key][prop] === undefined || infoAdicional[key][prop] === null || infoAdicional[key][prop] === "")) infoAdicional[key][prop] = remoteRecord.infoAdicional[key][prop]; } 
                if (remoteRecord.comentariosGenerales) for (let key in remoteRecord.comentariosGenerales) if (remoteRecord.comentariosGenerales[key] && (!comentariosGenerales[key] || comentariosGenerales[key] === "")) comentariosGenerales[key] = remoteRecord.comentariosGenerales[key]; 
            } 
            const finalPayload = { planes, realData, comentarios, infoAdicional, comentariosGenerales, timestamp: new Date().toISOString() }; 
            const sizeMB = JSON.stringify(finalPayload).length / (1024*1024); 
            if (sizeMB > 9.8) { statusDiv.innerHTML = `❌ Excede 10MB (${sizeMB.toFixed(2)} MB).`; isSyncing = false; return; } 
            const response = await fetch(JSONBIN_URL, { 
                method: "PUT", 
                headers: { "Content-Type": "application/json", "X-Master-Key": JSONBIN_API_KEY, "X-Bin-Meta": "false" }, 
                body: JSON.stringify(finalPayload) 
            }); 
            if (!response.ok) throw new Error(`HTTP ${response.status}`); 
            statusDiv.innerHTML = "✅ Datos guardados en la nube."; 
            setTimeout(() => { if (statusDiv.innerHTML.includes("guardados")) statusDiv.innerHTML = "✅ Todo sincronizado."; }, 2500); 
            actualizarTodo(); 
        } catch (err) { statusDiv.innerHTML = `❌ Error: ${err.message}`; } 
        finally { isSyncing = false; } 
    }
    
    async function forceReloadFromCloud() { 
        if (isSyncing) return; 
        if (!confirm("⚠️ Esta acción reemplazará TODOS los datos locales con la versión de la nube. ¿Deseas continuar?")) return; 
        isSyncing = true; 
        const statusDiv = document.getElementById("syncStatus"); 
        statusDiv.innerHTML = "⬇️ Forzando recarga..."; 
        try { 
            const remote = await fetchFromCloud(); 
            if (remote) { 
                if (remote.planes) planes = JSON.parse(JSON.stringify(remote.planes)); 
                if (remote.realData) realData = JSON.parse(JSON.stringify(remote.realData)); 
                if (remote.comentarios) comentarios = JSON.parse(JSON.stringify(remote.comentarios)); 
                if (remote.infoAdicional) infoAdicional = JSON.parse(JSON.stringify(remote.infoAdicional)); 
                if (remote.comentariosGenerales) comentariosGenerales = JSON.parse(JSON.stringify(remote.comentariosGenerales)); 
                saveRealesToLocal(); savePlanesToLocal(); saveComentariosToLocal(); guardarInfoAdicionalLocal(); guardarComentariosGeneralesLocal(); 
                statusDiv.innerHTML = "✅ Recarga completa."; 
                actualizarTodo(); 
            } else { statusDiv.innerHTML = "⚠️ No hay datos."; } 
        } catch (err) { statusDiv.innerHTML = `❌ Error: ${err.message}`; } 
        finally { isSyncing = false; } 
    }
    
    async function reemplazarPlanRealEnNube() { 
        if (isSyncing) return; 
        isSyncing = true; 
        const statusDiv = document.getElementById("syncStatus"); 
        statusDiv.innerHTML = "🔄 Reemplazando curvas S..."; 
        try { 
            let remoteRecord = null; 
            try { remoteRecord = await fetchFromCloud(); } catch(e) {} 
            const newRecord = { 
                planes: JSON.parse(JSON.stringify(planes)), 
                realData: JSON.parse(JSON.stringify(realData)), 
                comentarios: (remoteRecord && remoteRecord.comentarios) ? remoteRecord.comentarios : comentarios, 
                infoAdicional: (remoteRecord && remoteRecord.infoAdicional) ? remoteRecord.infoAdicional : infoAdicional, 
                comentariosGenerales: (remoteRecord && remoteRecord.comentariosGenerales) ? remoteRecord.comentariosGenerales : comentariosGenerales, 
                timestamp: new Date().toISOString() 
            }; 
            const sizeMB = JSON.stringify(newRecord).length / (1024*1024); 
            if (sizeMB > 9.8) { statusDiv.innerHTML = `❌ Excede 10MB (${sizeMB.toFixed(2)} MB).`; isSyncing = false; return; } 
            const response = await fetch(JSONBIN_URL, { 
                method: "PUT", 
                headers: { "Content-Type": "application/json", "X-Master-Key": JSONBIN_API_KEY, "X-Bin-Meta": "false" }, 
                body: JSON.stringify(newRecord) 
            }); 
            if (!response.ok) throw new Error(`HTTP ${response.status}`); 
            statusDiv.innerHTML = "✅ Curvas S reemplazadas."; 
            setTimeout(() => { if (statusDiv.innerHTML.includes("exitosamente")) statusDiv.innerHTML = "✅ Listo."; }, 3000); 
            actualizarTodo(); 
        } catch (err) { statusDiv.innerHTML = `❌ Error: ${err.message}`; } 
        finally { isSyncing = false; } 
    }
    
    async function cargarDesdeNube() { 
        if (isSyncing) return; 
        isSyncing = true; 
        const statusDiv = document.getElementById("syncStatus"); 
        statusDiv.innerHTML = "🔄 Fusionando..."; 
        try { 
            const remote = await fetchFromCloud(); 
            if (remote) { 
                mergeRemoteToLocal(remote); 
                statusDiv.innerHTML = "✅ Fusión completada."; 
                actualizarTodo(); 
            } else { statusDiv.innerHTML = "✅ Sin datos previos."; } 
        } catch (err) { statusDiv.innerHTML = `❌ Error: ${err.message}`; } 
        finally { isSyncing = false; } 
    }
    
    let autoSyncInterval = null, timerInterval = null; let secondsLeft = 900;
    function actualizarTimerDisplay() { const timerSpan = document.getElementById("syncTimer"); if (timerSpan) { const mins = Math.floor(secondsLeft / 60); const secs = secondsLeft % 60; timerSpan.textContent = `⏱️ Próxima sincronización: ${mins}:${secs.toString().padStart(2,'0')}`; } }
    function iniciarAutoSyncConTimer() { if (autoSyncInterval) clearInterval(autoSyncInterval); if (timerInterval) clearInterval(timerInterval); secondsLeft = 900; actualizarTimerDisplay(); timerInterval = setInterval(() => { if (secondsLeft > 0) { secondsLeft--; actualizarTimerDisplay(); } }, 1000); autoSyncInterval = setInterval(async () => { if (!isSyncing) { await cargarDesdeNube(); secondsLeft = 900; actualizarTimerDisplay(); } }, 900000); }
    
    let chartPortafolio, chartProyecto, currentFechaIdx = 0;
    const verticalLinePlugin = { id:'verticalLine', afterDraw(chart,args,options){ const {ctx,chartArea:{top,bottom,left,right},scales}=chart; const fechaIdx=currentFechaIdx; if(fechaIdx===undefined||fechaIdx<0||fechaIdx>=fechasCorte.length) return; const x=scales.x.getPixelForValue(fechaIdx); if(x<left||x>right) return; ctx.save(); ctx.beginPath(); ctx.moveTo(x,top); ctx.lineTo(x,bottom); ctx.lineWidth=2; ctx.strokeStyle='#e74c3c'; ctx.setLineDash([8,6]); ctx.stroke(); ctx.setLineDash([]); let planVal=null,realVal=null; if(chart===chartPortafolio){ planVal=planes["PORTAFOLIO"]?planes["PORTAFOLIO"][fechaIdx]:null; realVal=realData["PORTAFOLIO"][fechaIdx]; } else if(chart===chartProyecto){ const proy=document.getElementById("selectProyectoGrafica").value; planVal=planes[proy]?planes[proy][fechaIdx]:null; realVal=realData[proy]?realData[proy][fechaIdx]:null; } if(planVal!==null){ const yPlan=scales.y.getPixelForValue(planVal); const textPlan=`Plan: ${(planVal*100).toFixed(1)}%`; const metrics=ctx.measureText(textPlan); const rectWidth=metrics.width+8; const rectX=x-rectWidth-6; if(rectX>=left){ ctx.fillStyle="rgba(255,255,240,0.85)"; ctx.fillRect(rectX, yPlan-10, rectWidth, 18); ctx.fillStyle="#1f6392"; ctx.fillText(textPlan, rectX+4, yPlan+4); } } if(realVal!==null){ const yReal=scales.y.getPixelForValue(realVal); const textReal=`Real: ${(realVal*100).toFixed(1)}%`; const metrics=ctx.measureText(textReal); const rectWidth=metrics.width+8; const rectX=x+6; if(rectX+rectWidth<=right){ ctx.fillStyle="rgba(255,255,240,0.85)"; ctx.fillRect(rectX, yReal-10, rectWidth, 18); ctx.fillStyle="#d35400"; ctx.fillText(textReal, rectX+4, yReal+4); } } ctx.restore(); } };
    
    function initCharts(){ const labelsCortas=fechasCorte.map(f=>formatearFechaCorta(f)); chartPortafolio=new Chart(document.getElementById('avanceChart'),{ type:'line', data:{ labels:labelsCortas, datasets:[{ label:'📋 Planificado Portafolio (%)', data:planes["PORTAFOLIO"], borderColor:'#2c7da0', borderWidth:2, pointRadius:0, tension:0, fill:false },{ label:'✅ Real ejecutado Portafolio (%)', data:realData["PORTAFOLIO"].map(v=>v!==null?v:null), borderColor:'#e67e22', borderWidth:2, pointRadius:0, tension:0, fill:false, spanGaps:true }] }, options:{ responsive:true, maintainAspectRatio:true, animation:false, plugins:{ tooltip:{ callbacks:{ label:(ctx)=>`${ctx.dataset.label}: ${ctx.raw!==null?(ctx.raw*100).toFixed(2)+'%':'Sin dato'}`, title:(tooltipItems)=>formatearFechaCompleta(fechasCorte[tooltipItems[0].dataIndex]) } } }, scales:{ x:{ ticks:{ color:'#334155', maxRotation:90, minRotation:90, autoSkip:true, maxTicksLimit:30 } }, y:{ ticks:{ callback:(v)=>(v*100).toFixed(0)+'%' }, min:0, max:1.05 } } }, plugins:[verticalLinePlugin] }); chartProyecto=new Chart(document.getElementById('proyectoChart'),{ type:'line', data:{ labels:labelsCortas, datasets:[{ label:'📋 Planificado (%)', data:[], borderColor:'#2c7da0', borderWidth:2, pointRadius:0, tension:0, fill:false },{ label:'✅ Real ejecutado (%)', data:[], borderColor:'#e67e22', borderWidth:2, pointRadius:0, tension:0, fill:false, spanGaps:true }] }, options:{ responsive:true, maintainAspectRatio:true, animation:false, plugins:{ tooltip:{ callbacks:{ label:(ctx)=>`${ctx.dataset.label}: ${ctx.raw!==null?(ctx.raw*100).toFixed(2)+'%':'Sin dato'}`, title:(tooltipItems)=>formatearFechaCompleta(fechasCorte[tooltipItems[0].dataIndex]) } } }, scales:{ x:{ ticks:{ maxRotation:90, minRotation:90, autoSkip:true } }, y:{ ticks:{ callback:(v)=>(v*100).toFixed(0)+'%' }, min:0, max:1.05 } } }, plugins:[verticalLinePlugin] }); }
    
    function actualizarGraficosConLinea(){ if(!chartPortafolio) return; chartPortafolio.data.datasets[0].data=planes["PORTAFOLIO"]; chartPortafolio.data.datasets[1].data=realData["PORTAFOLIO"].map(v=>v!==null?v:null); chartPortafolio.update('none'); const proyecto=document.getElementById("selectProyectoGrafica").value; chartProyecto.data.datasets[0].data=planes[proyecto]||[]; chartProyecto.data.datasets[1].data=realData[proyecto].map(v=>v!==null?v:null); chartProyecto.update('none'); }
    
    function actualizarSemaforoPorFecha(){ const fechaIdx=currentFechaIdx; if(isNaN(fechaIdx)) return; const tbody=document.getElementById("semaforoBody"); tbody.innerHTML=""; const todosProyectos = ["PORTAFOLIO", ...codigosProyecto.filter(c=>c!=="PORTAFOLIO")]; for(let cod of todosProyectos){ const planVal=planes[cod]?planes[cod][fechaIdx]:0; const realVal=realData[cod][fechaIdx]; let realMostrar=(realVal!==null)?(realVal*100).toFixed(2)+"%":"—"; let variacion=(realVal!==null)?(realVal-planVal).toFixed(4):null; let variacionMostrar=(variacion!==null)?(parseFloat(variacion)>=0?'+':'')+(parseFloat(variacion)*100).toFixed(2)+"%":"N/D"; let spi=(realVal!==null && planVal>0)?(realVal/planVal).toFixed(3):"N/D"; let estadoHtml=""; if(realVal!==null){ const diff=parseFloat(variacion); if(Math.abs(diff)<1e-6){ estadoHtml="<span class='traffic-light green'></span> 🟢 En línea"; } else if(realVal>planVal){ estadoHtml="<span class='traffic-light green'></span> 🟢 Adelanto"; } else{ estadoHtml="<span class='traffic-light red'></span> 🔴 Atraso"; } } else{ estadoHtml="<span class='traffic-light yellow'></span> ⚪ Sin dato real"; } let nombreMostrar = (cod === "PORTAFOLIO") ? "📊 PORTAFOLIO GENERAL" : cod; tbody.insertAdjacentHTML('beforeend',`<tr><td><strong>${nombreMostrar}</strong></td><td>${(planVal*100).toFixed(2)}%</td><td>${realMostrar}</td><td style="color:${realVal!==null?(parseFloat(variacion)>=0?'#2ecc71':'#e74c3c'):'#f1c40f'}">${variacionMostrar}</td><td>${spi}</td><td>${estadoHtml}</td></tr>`); } actualizarGraficosConLinea(); actualizarMosaicos(); actualizarCategoriasProyecto(); cargarComentarioUI(); }
    
    function actualizarSelectoresFechas(){ const selectUnico=document.getElementById("fechaUnicaReporte"); if(!selectUnico) return; selectUnico.innerHTML=""; fechasCorte.forEach((f,idx)=>{ let opt=document.createElement("option"); opt.value=idx; opt.textContent=formatearFechaCompleta(f); selectUnico.appendChild(opt); }); let lastIdx=0; for(let i=realData["PORTAFOLIO"].length-1;i>=0;i--) if(realData["PORTAFOLIO"][i]!==null){ lastIdx=i; break; } selectUnico.value=lastIdx; setFechaGlobal(lastIdx); }
    
    function setFechaGlobal(fechaIdx){ if (isNaN(fechaIdx) || fechaIdx<0 || fechaIdx>=fechasCorte.length) return; const selectUnico=document.getElementById("fechaUnicaReporte"); if(selectUnico) selectUnico.value=fechaIdx; currentFechaIdx = fechaIdx; actualizarSemaforoPorFecha(); const key=fechaIdx.toString(); const data = infoAdicional[key] || {}; if(!activeEditingFields.has('personalTotal')) document.getElementById("personalTotal").value = (data.personalTotal !== undefined && data.personalTotal !== null) ? data.personalTotal : ""; if(!activeEditingFields.has('personalDia')) document.getElementById("personalDia").value = (data.personalDia !== undefined && data.personalDia !== null) ? data.personalDia : ""; if(!activeEditingFields.has('personalNoche')) document.getElementById("personalNoche").value = (data.personalNoche !== undefined && data.personalNoche !== null) ? data.personalNoche : ""; if(!activeEditingFields.has('hseNovedad')) document.getElementById("hseNovedad").value = data.hseNovedad || ""; if(!activeEditingFields.has('comentariosGenerales')) document.getElementById("comentariosGenerales").value = comentariosGenerales[key] || ""; }
    
    function actualizarMosaicos(){ const container=document.getElementById("proyectosMosaicos"); if(!container) return; const fechaSeleccionada=currentFechaIdx; const proyectos=codigosProyecto.filter(c=>c!=="PORTAFOLIO"); container.innerHTML=""; for(let cod of proyectos){ const tieneComentario=comentarios[cod] && comentarios[cod][fechaSeleccionada] && (comentarios[cod][fechaSeleccionada].management || comentarios[cod][fechaSeleccionada].ingenieria || comentarios[cod][fechaSeleccionada].procura || comentarios[cod][fechaSeleccionada].construccion); const div=document.createElement("div"); div.className=`proyecto-mosaico ${tieneComentario?"has-comment":""}`; div.innerHTML=`${tieneComentario?"💬 ":""}${cod}`; div.addEventListener("click",()=>{ document.querySelectorAll(".proyecto-mosaico").forEach(m=>m.classList.remove("active")); div.classList.add("active"); window.proyectoSeleccionado=cod; actualizarCategoriasProyecto(); cargarComentarioUI(); }); container.appendChild(div); } if(window.proyectoSeleccionado && proyectos.includes(window.proyectoSeleccionado)){ document.querySelectorAll(".proyecto-mosaico").forEach(m=>{ if(m.textContent.trim().replace("💬","").trim()===window.proyectoSeleccionado) m.classList.add("active"); }); } else if(proyectos.length>0){ window.proyectoSeleccionado=proyectos[0]; actualizarCategoriasProyecto(); cargarComentarioUI(); } }
    
    function actualizarCategoriasProyecto(){ const container=document.getElementById("categoriasProyectoArea"); if(!container) return; const proyecto=window.proyectoSeleccionado; const fechaIdx=currentFechaIdx; if(!proyecto) return; let html=`<div style="font-weight:bold;">Categorías de ${proyecto}:</div><div class="categorias-proyecto">`; for(let cat of categorias){ const tieneComentario=comentarios[proyecto] && comentarios[proyecto][fechaIdx] && comentarios[proyecto][fechaIdx][cat] && comentarios[proyecto][fechaIdx][cat].trim()!==""; html+=`<div class="cat-btn ${tieneComentario?"has-comment":""}" data-cat="${cat}">${tieneComentario?"💬 ":""}${nombresCat[cat]}</div>`; } html+=`</div>`; container.innerHTML=html; document.querySelectorAll(".cat-btn").forEach(btn=>{ btn.addEventListener("click",()=>{ document.querySelectorAll(".cat-btn").forEach(b=>b.classList.remove("active")); btn.classList.add("active"); window.categoriaSeleccionada=btn.getAttribute("data-cat"); cargarComentarioUI(); }); }); if(!window.categoriaSeleccionada || !categorias.includes(window.categoriaSeleccionada)) window.categoriaSeleccionada="management"; const activeBtn=document.querySelector(`.cat-btn[data-cat="${window.categoriaSeleccionada}"]`); if(activeBtn) activeBtn.classList.add("active"); }
    
    function cargarComentarioUI(){ const proyecto=window.proyectoSeleccionado; const fechaIdx=currentFechaIdx; const categoria=window.categoriaSeleccionada; if(proyecto && categoria && comentarios[proyecto] && comentarios[proyecto][fechaIdx] && comentarios[proyecto][fechaIdx][categoria]!==undefined) document.getElementById("comentarioTexto").value=comentarios[proyecto][fechaIdx][categoria]; else document.getElementById("comentarioTexto").value=""; }
    
    async function guardarComentario(proyecto,fechaIdx,categoria,texto){ if(!comentarios[proyecto]) comentarios[proyecto]=[]; if(!comentarios[proyecto][fechaIdx]) comentarios[proyecto][fechaIdx]={ management:"", ingenieria:"", procura:"", construccion:"" }; comentarios[proyecto][fechaIdx][categoria]=texto; saveComentariosToLocal(); actualizarMosaicos(); actualizarCategoriasProyecto(); await guardarEnNubeConMerge(); }
    async function guardarPersonalHSE(fechaIdx) { const key = fechaIdx.toString(); const total = document.getElementById("personalTotal").value === "" ? null : parseFloat(document.getElementById("personalTotal").value); const dia = document.getElementById("personalDia").value === "" ? null : parseFloat(document.getElementById("personalDia").value); const noche = document.getElementById("personalNoche").value === "" ? null : parseFloat(document.getElementById("personalNoche").value); const hse = document.getElementById("hseNovedad").value; if (!infoAdicional[key]) infoAdicional[key] = {}; infoAdicional[key].personalTotal = total; infoAdicional[key].personalDia = dia; infoAdicional[key].personalNoche = noche; infoAdicional[key].hseNovedad = hse; guardarInfoAdicionalLocal(); await guardarEnNubeConMerge(); }
    async function guardarComentariosGeneralesHandler(fechaIdx) { const key = fechaIdx.toString(); const texto = document.getElementById("comentariosGenerales").value; comentariosGenerales[key] = texto; guardarComentariosGeneralesLocal(); await guardarEnNubeConMerge(); }
    function actualizarTodo(){ actualizarGraficosConLinea(); actualizarSemaforoPorFecha(); actualizarSelectoresFechas(); actualizarMosaicos(); actualizarCategoriasProyecto(); cargarComentarioUI(); }
    
    // ========== FUNCIONES QR CORREGIDAS ==========
    let qrCodeInstance = null;
    
    function generarQR() {
        // Obtener la URL completa actual (esta es la URL del HTML)
        const urlActual = window.location.href;
        
        // Mostrar la URL en el div para referencia
        const urlDisplay = document.getElementById("qrUrlDisplay");
        if (urlDisplay) {
            urlDisplay.innerHTML = `🔗 URL del reporte: <a href="${urlActual}" target="_blank">${urlActual.length > 80 ? urlActual.substring(0, 80) + '...' : urlActual}</a>`;
        }
        
        // Limpiar el contenedor del QR
        const qrDiv = document.getElementById("qrCodeDiv");
        qrDiv.innerHTML = "";
        
        try {
            // Crear el QR usando la librería QRCode.js con la URL completa
            qrCodeInstance = new QRCode(qrDiv, {
                text: urlActual,
                width: 220,
                height: 220,
                colorDark: "#000000",
                colorLight: "#ffffff",
                correctLevel: QRCode.CorrectLevel.H
            });
            
            // Mensaje de éxito
            const successMsg = document.createElement("div");
            successMsg.style.cssText = "margin-top: 10px; font-size: 0.75rem; color: #27ae60;";
            successMsg.innerHTML = "✅ QR generado correctamente. Escanea con tu celular para abrir el reporte.";
            qrDiv.appendChild(successMsg);
            
        } catch(error) {
            console.error("Error generando QR:", error);
            qrDiv.innerHTML = `<div style="padding: 20px; background: #fee2e2; border-radius: 12px; color: #c0392b;">
                ❌ Error al generar QR: ${error.message}<br>
                <small>Intenta copiar la URL de arriba manualmente.</small>
            </div>`;
        }
    }
    
    function copiarURL() {
        const url = window.location.href;
        navigator.clipboard.writeText(url).then(() => {
            const btn = document.getElementById("btnCopiarURL");
            const textoOriginal = btn.textContent;
            btn.textContent = "✅ ¡Copiado!";
            setTimeout(() => {
                btn.textContent = textoOriginal;
            }, 2000);
        }).catch(err => {
            alert("No se pudo copiar la URL: " + err);
        });
    }
    
    // Eventos
    document.getElementById("btnGenerarQR").addEventListener("click", generarQR);
    document.getElementById("btnCopiarURL").addEventListener("click", copiarURL);
    
    // Generar el QR automáticamente al cargar la página
    window.addEventListener("load", function() {
        setTimeout(generarQR, 500);
    });
    
    document.getElementById("btnCargarExcel").addEventListener("click", async () => { const fileInput = document.getElementById("excelUpload"); if (!fileInput.files.length) { document.getElementById("excelStatus").innerHTML = "❌ Selecciona archivo."; return; } const file = fileInput.files[0]; const statusSpan = document.getElementById("excelStatus"); statusSpan.innerHTML = "📂 Procesando..."; const reader = new FileReader(); reader.onload = async (e) => { try { const workbook = XLSX.read(e.target.result, { type: "array" }); let sheetName = workbook.SheetNames.find(n => n.toLowerCase().includes("curva") || n.toLowerCase().includes("planned")); if (!sheetName) sheetName = workbook.SheetNames[0]; const sheet = workbook.Sheets[sheetName]; const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" }); if (!rows || rows.length < 2) throw new Error("Datos insuficientes"); for (let rowIdx = 1; rowIdx < rows.length; rowIdx++) { const row = rows[rowIdx]; if (!row || row.length < 2) continue; const tipo = row[1]?.toString().trim().toUpperCase(); let codigo = row[0]?.toString().trim(); if (!codigo || !codigosProyecto.includes(codigo)) continue; if (tipo === "PLANNED") for (let i=0; i<Math.min(rows[0].length-2, fechasCorte.length); i++) { let val = parseFloat(row[2+i]); if (!isNaN(val) && val>=0 && val<=1) planes[codigo][i] = val; } else if (tipo === "ACTUAL") for (let i=0; i<Math.min(rows[0].length-2, fechasCorte.length); i++) { let val = parseFloat(row[2+i]); if (!isNaN(val) && val>=0 && val<=1) realData[codigo][i] = val; } } saveRealesToLocal(); savePlanesToLocal(); await reemplazarPlanRealEnNube(); statusSpan.innerHTML = "✅ Excel cargado y reemplazado."; actualizarTodo(); } catch(err) { statusSpan.innerHTML = `❌ Error: ${err.message}`; } }; reader.readAsArrayBuffer(file); });
    document.getElementById("btnGuardarComentario").addEventListener("click",async()=>{ const proyecto=window.proyectoSeleccionado; const fechaIdx=currentFechaIdx; const categoria=window.categoriaSeleccionada; const texto=document.getElementById("comentarioTexto").value; if(!proyecto||!categoria) return; await guardarComentario(proyecto,fechaIdx,categoria,texto); });
    document.getElementById("btnEditarComentario").addEventListener("click",async()=>{ const proyecto=window.proyectoSeleccionado; const fechaIdx=currentFechaIdx; const categoria=window.categoriaSeleccionada; const texto=document.getElementById("comentarioTexto").value; if(!proyecto||!categoria) return; await guardarComentario(proyecto,fechaIdx,categoria,texto); });
    document.getElementById("btnSync").addEventListener("click",cargarDesdeNube);
    document.getElementById("btnForceReload").addEventListener("click",forceReloadFromCloud);
    document.getElementById("btnUpload").addEventListener("click",()=>guardarEnNubeConMerge());
    document.getElementById("btnGuardarPersonalHSE").addEventListener("click",()=>{ guardarPersonalHSE(currentFechaIdx); });
    document.getElementById("fechaUnicaReporte").addEventListener("change",(e)=>setFechaGlobal(parseInt(e.target.value)));
    document.getElementById("btnGuardarComentariosGenerales").addEventListener("click",()=>{ guardarComentariosGeneralesHandler(currentFechaIdx); });
    document.getElementById("selectProyectoGrafica").addEventListener("change",()=>actualizarGraficosConLinea());
    
    cargarPlanesDefault();
    loadFromLocal();
    initCharts();
    actualizarSelectoresFechas();
    setFechaGlobal(currentFechaIdx);
    actualizarMosaicos();
    actualizarCategoriasProyecto();
    cargarComentarioUI();
    cargarDesdeNube();
    iniciarAutoSyncConTimer();
</script>
</body>
</html>
