/**
 * Evento disparado automáticamente al abrir la hoja de cálculo.
 * Crea un menú personalizado en la barra superior.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu(' Dashboard Docentes')
    .addItem('Abrir Panel de Control', 'abrirPanelDocentes')
    .addToUi();
}

/**
 * Despliega la interfaz lateral (Sidebar) construida desde el archivo 'PanelDocentes.html'.
 */
function abrirPanelDocentes() {
  var html = HtmlService.createHtmlOutputFromFile('PanelDocentes')
      .setTitle('Gestión y Cumplimiento Docente')
      .setWidth(450);
  SpreadsheetApp.getUi().showSidebar(html);
}

// ==========================================
// 1. CONVERSIÓN Y CLAVE ÚNICA DE FECHA
// ==========================================

/**
 * Convierte un valor de celda (fecha nativa o texto) a un objeto Date estandarizado sin horas.
 * @param {any} valor - Contenido de la celda.
 * @returns {Date|null} Objeto Date parseado o null si es inválido.
 */
function convertirFecha(valor) {
  if (!valor) return null;

  // Caso A: Si ya es un objeto Date válido de JavaScript
  if (Object.prototype.toString.call(valor) === "[object Date]" && !isNaN(valor.getTime())) {
    return new Date(valor.getFullYear(), valor.getMonth(), valor.getDate());
  }

  // Caso B: Cadena de texto (ej. "DD/MM/YYYY", "YYYY-MM-DD", "DD.MM.YYYY")
  var texto = String(valor).trim();
  if (!texto) return null;

  var partes = texto.split(/[\/\-\.]/);
  if (partes.length !== 3) return null;

  var dia, mes, anio;

  // Determina el formato según la longitud del primer segmento (Año primero vs Día primero)
  if (partes[0].length === 4) {
    anio = Number(partes[0]);
    mes  = Number(partes[1]);
    dia  = Number(partes[2]);
  } else {
    dia  = Number(partes[0]);
    mes  = Number(partes[1]);
    anio = Number(partes[2]);
  }

  if (isNaN(dia) || isNaN(mes) || isNaN(anio)) return null;

  // Construye la fecha (los meses en JS van de 0 a 11)
  var fecha = new Date(anio, mes - 1, dia);

  // Valida que los valores ingresados correspondan a una fecha real (ej. evita 31/02)
  if (
    fecha.getFullYear() !== anio ||
    fecha.getMonth() + 1 !== mes ||
    fecha.getDate() !== dia
  ) {
    return null;
  }

  return fecha;
}

/**
 * Convierte una fecha a formato estándar "YYYY-MM-DD" para comparaciones precisas.
 * @param {Date} fecha - Objeto fecha.
 * @returns {string} Cadena formateada.
 */
function claveFecha(fecha) {
  if (!fecha) return "";
  return Utilities.formatDate(fecha, Session.getScriptTimeZone(), "yyyy-MM-dd");
}

/**
 * Recorre la columna A para encontrar la fecha válida más reciente (máximo valor de timestamp).
 * @param {Array<Array>} data - Matriz de datos obtenida del Spreadsheet.
 * @returns {Date|null} La fecha más reciente encontrada.
 */
function obtenerFechaMasReciente(data) {
  var ultimaFecha = null;
  var ultimoTiempo = -Infinity;

  for (var i = 1; i < data.length; i++) {
    var fechaObj = convertirFecha(data[i][0]); // Columna A
    if (!fechaObj) continue;

    var tiempo = fechaObj.getTime();
    if (tiempo > ultimoTiempo) {
      ultimoTiempo = tiempo;
      ultimaFecha = fechaObj;
    }
  }

  return ultimaFecha;
}

/**
 * Filtra la hoja de cálculo devolviendo únicamente las filas cuya fecha coincida con la más reciente.
 * @param {Array<Array>} data - Matriz completa de datos.
 * @returns {Array<Array>} Filas filtradas por la cohorte más reciente.
 */
function obtenerFilasFechaMasReciente(data) {
  var ultimaFecha = obtenerFechaMasReciente(data);
  if (!ultimaFecha) return [];

  var claveMax = claveFecha(ultimaFecha);
  var resultado = [];

  for (var i = 1; i < data.length; i++) {
    var fila = data[i];
    var fechaObj = convertirFecha(fila[0]);
    if (fechaObj && claveFecha(fechaObj) === claveMax) {
      resultado.push(fila);
    }
  }

  return resultado;
}

/**
 * Convierte cualquier representación de porcentaje (ej. "85%", 0.85, " 85 ") a un entero limpio.
 * @param {any} val - Valor raw de la celda.
 * @returns {number} Porcentaje entero normalizado.
 */
function parsearPorcentaje(val) {
  if (val === null || val === undefined) return 0;
  var str = String(val).replace("%", "").trim();
  var num = parseFloat(str);
  if (isNaN(num)) return 0;

  // Ajusta escala si viene formateado en base decimal (ej: 0.85 -> 85)
  if (num <= 1 && num > 0) {
    num = num * 100;
  }
  return Math.round(num);
}

// ==========================================
// 2. CONSULTAS DE DATOS PARA PANEL
// ==========================================

/**
 * Obtiene la lista ordenada de nombres de docentes activos en la cohorte más reciente.
 * @returns {Array<string>} Arreglo con nombres únicos de docentes.
 */
function obtenerListaDocentes() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  
  var filasRecientes = obtenerFilasFechaMasReciente(data);
  var docentes = new Set();

  for (var i = 0; i < filasRecientes.length; i++) {
    var fila = filasRecientes[i];
    var nombreDocente = fila[3]; // Columna D (Docente)
    var nivel = String(fila[10] || '').trim().toLowerCase(); // Columna K (Nivel/Meta)

    // Solo incluye docentes con metas válidas asignadas
    if (nombreDocente && nivel !== 'sin meta asignada') {
      docentes.add(nombreDocente);
    }
  }

  return Array.from(docentes).sort();
}

/**
 * Extrae las asignaturas y cálculo de actividades faltantes de un docente específico en la última cohorte.
 * @param {string} nombreDocente - Nombre del docente a consultar.
 * @returns {Array<Object>} Arreglo de objetos con el detalle de cada materia/clase.
 */
function obtenerDetalleDocente(nombreDocente) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  
  var filasRecientes = obtenerFilasFechaMasReciente(data);
  var resultados = [];

  for (var i = 0; i < filasRecientes.length; i++) {
    var fila = filasRecientes[i];
    var docenteFila = fila[3]; // Columna D
    var nivel = String(fila[10] || '').trim().toLowerCase(); // Columna K

    if (docenteFila === nombreDocente && nivel !== 'sin meta asignada') {
      var porcentaje = parsearPorcentaje(fila[12]); // Columna M (% Cumplimiento)

      var faltaLibro = 0;
      var faltaPropia = 0;
      var faltaMuro = 0;

      // Si no ha llegado al 100%, calcula exactamente cuántas actividades restan
      if (porcentaje < 100) {
        var metaLibro = Number(fila[14]) || 0;   // Columna O (Meta Libro)
        var totalLibro = Number(fila[15]) || 0;  // Columna P (Total Libro)
        faltaLibro = Math.max(0, metaLibro - totalLibro);

        var metaPropia = Number(fila[17]) || 0;  // Columna R (Meta Propia)
        var totalPropia = Number(fila[18]) || 0; // Columna S (Total Propia)
        faltaPropia = Math.max(0, metaPropia - totalPropia);

        var metaMuro = Number(fila[20]) || 0;    // Columna U (Meta Muro)
        var totalMuro = Number(fila[21]) || 0;   // Columna V (Total Muro)
        faltaMuro = Math.max(0, metaMuro - totalMuro);
      }

      resultados.push({
        clase: fila[6] || 'N/A',       // Columna G
        grado: fila[7] || 'N/A',       // Columna H
        grupo: fila[8] || 'N/A',       // Columna I
        materia: fila[4] || 'N/A',     // Columna E
        cumplimiento: porcentaje,
        faltaLibro: faltaLibro,
        faltaPropia: faltaPropia,
        faltaMuro: faltaMuro
      });
    }
  }
  return resultados;
}

/**
 * Envía un correo electrónico HTML al docente.
 * @param {string} destinatario - Correo del docente.
 * @param {string} asunto - Asunto del correo.
 * @param {string} mensaje - Cuerpo del mensaje en formato texto/HTML.
 * @returns {Object} Estado del envío.
 */
function enviarCorreoDocente(destinatario, asunto, mensaje) {
  try {
    MailApp.sendEmail({
      to: destinatario,
      subject: asunto,
      htmlBody: mensaje.replace(/\n/g, '<br>')
    });
    return { exito: true };
  } catch (e) {
    return { exito: false, error: e.toString() };
  }
}

// ==========================================
// 3. GENERACIÓN DE PDF
// ==========================================

/**
 * Genera el reporte oficial en PDF para un docente con sanitización numérica estricta,
 * promedio por materias, diseño HTML/CSS y footer con fecha de cohorte.
 * @param {string} nombreDocente - Nombre del docente.
 * @returns {Object} Objeto conteniendo el PDF codificado en Base64 y su nombre de archivo.
 */
function generarPDFDocente(nombreDocente) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  
  var fechaMasRecienteObj = obtenerFechaMasReciente(data);
  var detalles = obtenerDetalleDocente(nombreDocente);
  
  if (!detalles || detalles.length === 0) {
    throw new Error('No se encontraron registros recientes para el docente.');
  }

  // 1. Determina las fechas a mostrar (Cohorte de la Hoja vs Fecha actual de emisión)
  var fechaCohorte = fechaMasRecienteObj 
    ? Utilities.formatDate(fechaMasRecienteObj, Session.getScriptTimeZone(), "dd/MM/yyyy")
    : Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");

  var fechaHoy = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");

  // 2. Sanitización numérica y promedio ponderado
  var totalMaterias = detalles.length;
  var sumaPorcentajesLimpios = 0;
  var completadas = 0;

  detalles.forEach(function(d) {
    // Extrae exclusivamente caracteres numéricos y puntos decimales
    var strVal = String(d.cumplimiento || '0').replace(/[^0-9.]/g, '').trim();
    var numVal = parseFloat(strVal) || 0;

    // Escala de decimal (0.53) a porcentaje entero (53) si corresponde
    if (numVal <= 1 && numVal > 0) {
      numVal = numVal * 100;
    }

    // Normaliza el rango para evitar valores inconsistentes (0 a 100)
    numVal = Math.min(100, Math.max(0, Math.round(numVal)));

    sumaPorcentajesLimpios += numVal;
    
    if (numVal >= 100) {
      completadas++;
    }

    d.cumplimientoLimpio = numVal;
  });

  // Cálculo directo del promedio: (Suma de % de cada materia) / (Total de materias)
  var promedioCumplimiento = totalMaterias > 0 
    ? Math.round(sumaPorcentajesLimpios / totalMaterias) 
    : 0;

  // URLs de los logos institucionales
  var logo1 = "https://w7.pngwing.com/pngs/829/654/png-transparent-national-police-of-colombia-national-police-corps-army-officer-police-emblem-people-logo-thumbnail.png";
  var logo2 = "https://www.pngitem.com/pimgs/b/355-3559911_nuevo-png.png";

  // 3. Estructura HTML/CSS del reporte PDF
  var html = `
    <!DOCTYPE html>
    <html>
    <head>
      <style>
        body { font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif; color: #2d3748; margin: 25px; }
        .header-table { width: 100%; border-collapse: collapse; margin-bottom: 20px; border-bottom: 3px solid #1a365d; padding-bottom: 15px; }
        .header-table td { border: none; vertical-align: middle; padding: 0; }
        .logo-cell-left { width: 15%; text-align: left; }
        .logo-cell-right { width: 15%; text-align: right; }
        .title-cell { width: 70%; text-align: center; }
        
        .header-logo { height: 60px; width: auto; max-width: 90px; object-fit: contain; }
        .title-cell h1 { font-size: 18px; color: #1a365d; margin: 0 0 6px 0; text-transform: uppercase; letter-spacing: 0.5px; font-weight: bold; }
        .title-cell p { margin: 0; color: #718096; font-size: 11px; }

        .kpi-table { width: 100%; border-collapse: separate; border-spacing: 10px 0; margin-bottom: 25px; }
        .kpi-card { background: #f7fafc; border: 1px solid #e2e8f0; border-radius: 8px; padding: 10px; text-align: center; width: 33.33%; }
        .kpi-val { font-size: 18px; font-weight: bold; color: #2b6cb0; margin-top: 3px; }
        .kpi-lbl { font-size: 10px; text-transform: uppercase; color: #718096; font-weight: 600; }

        .data-table { width: 100%; border-collapse: collapse; margin-top: 10px; font-size: 11px; }
        .data-table th { background-color: #1a365d; color: #ffffff; text-align: left; padding: 8px 10px; font-weight: 600; text-transform: uppercase; font-size: 10px; }
        .data-table td { padding: 8px 10px; border-bottom: 1px solid #e2e8f0; }
        .data-table tr:nth-child(even) { background-color: #f7fafc; }

        .badge { padding: 3px 7px; border-radius: 10px; font-size: 10px; font-weight: bold; text-align: center; display: inline-block; }
        .badge-success { background-color: #c6f6d5; color: #22543d; }
        .badge-warning { background-color: #feebc8; color: #744210; }
        
        .footer { margin-top: 35px; font-size: 10px; color: #718096; text-align: center; border-top: 1px solid #e2e8f0; padding-top: 12px; }
      </style>
    </head>
    <body>
      
      <!-- Encabezado con Logos -->
      <table class="header-table">
        <tr>
          <td class="logo-cell-left">
            <img src="${logo1}" class="header-logo" alt="Logo 1" />
          </td>
          <td class="title-cell">
            <h1>Informe de Cumplimiento Docente</h1>
            <p><strong>Docente:</strong> ${nombreDocente} &nbsp;|&nbsp; <strong>Fecha de Emisión:</strong> ${fechaHoy}</p>
          </td>
          <td class="logo-cell-right">
            <img src="${logo2}" class="header-logo" alt="Logo 2" />
          </td>
        </tr>
      </table>

      <!-- Tarjetas de Resumen KPI -->
      <table class="kpi-table">
        <tr>
          <td class="kpi-card">
            <div class="kpi-lbl">Asignaturas Evaluadas</div>
            <div class="kpi-val">${totalMaterias}</div>
          </td>
          <td class="kpi-card">
            <div class="kpi-lbl">Cumplimiento Global</div>
            <div class="kpi-val">${promedioCumplimiento}%</div>
          </td>
          <td class="kpi-card">
            <div class="kpi-lbl">Metas Alcanzadas</div>
            <div class="kpi-val">${completadas} de ${totalMaterias}</div>
          </td>
        </tr>
      </table>

      <h3 style="font-size: 13px; color: #1a365d; margin-bottom: 8px; margin-top: 0;">Desglose por Materia y Grupo</h3>
      
      <!-- Tabla Principal de Asignaturas -->
      <table class="data-table">
        <thead>
          <tr>
            <th>Materia / Clase</th>
            <th>Grado - Grupo</th>
            <th>% Cumplimiento</th>
            <th>Faltantes (Libro / Propia / Muro)</th>
          </tr>
        </thead>
        <tbody>
  `;

  // Construcción dinámica de filas por asignatura
  detalles.forEach(function(d) {
    var valMostrar = d.cumplimientoLimpio;
    var estadoBadge = valMostrar >= 100 
      ? `<span class="badge badge-success">100% Ok</span>`
      : `<span class="badge badge-warning">${valMostrar}%</span>`;

    var faltantesTexto = valMostrar >= 100 
      ? '<em style="color:#2b6cb0;">Al día</em>'
      : `Libro: <strong>${d.faltaLibro}</strong> | Propia: <strong>${d.faltaPropia}</strong> | Muro: <strong>${d.faltaMuro}</strong>`;

    html += `
      <tr>
        <td><strong>${d.materia}</strong><br><small style="color:#718096">${d.clase}</small></td>
        <td>${d.grado} - ${d.grupo}</td>
        <td>${estadoBadge}</td>
        <td>${faltantesTexto}</td>
      </tr>
    `;
  });

  html += `
        </tbody>
      </table>

      <!-- Pie de página dinámico con la fecha de cohorte -->
      <div class="footer">
        Este documento oficial se ha generado automáticamente con base en los registros con cohorte: <strong>${fechaCohorte}</strong>.
      </div>
    </body>
    </html>
  `;

  // 4. Conversión del HTML a Blob PDF
  var blob = HtmlService.createHtmlOutput(html).getAs('application/pdf');
  blob.setName('Reporte_' + nombreDocente.replace(/\s+/g, '_') + '.pdf');

  // Retorna en Base64 para ser consumido por el cliente HTML/JS
  return {
    base64: Utilities.base64Encode(blob.getBytes()),
    filename: blob.getName()
  };
}
