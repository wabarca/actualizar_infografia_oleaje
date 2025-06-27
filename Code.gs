// Script de Google Apps Script para automatizar infografía de mareas y oleaje
// Requiere Drive API habilitada desde Servicios avanzados de Google

var CARPETA_ID = '1-vHU1nTNQO6DDqSEwAlfyIJ1-mttTFKs';
var CARPETA_DATOS_ENTRADA_ID = '1DKO73DbITEiWB6thpZQB4hFvMzA26t19';
var CARPETA_DATOS_ESTATICOS_ID = '1EuTqSqbeqdSdMzq9BcRk_fAURwV4WJwj';
var CARPETA_INFOGRAFIAS_ID = '1Njsi_ZgLepGW84M25fC288L_k1dBfEkl';

var NOMBRE_XLSX = 'Oleaje_Viento.xlsx';
var NOMBRE_GSHEET_OLEAJE = 'Oleaje_Viento';
var NOMBRE_MAREA = 'Marea';
var NOMBRE_TEMPLATE = 'template';

function actualizarInfografia() {
  var sheetControl = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  try {
    var hoy = new Date();
    var manana = new Date(hoy);
    manana.setDate(hoy.getDate() + 1);
    var fechaTexto = Utilities.formatDate(manana, Session.getScriptTimeZone(), 'dd/MM/yy');

    var carpetaDatosEntrada = DriveApp.getFolderById(CARPETA_DATOS_ENTRADA_ID);
    var hojaOleaje = SpreadsheetApp.openById(carpetaDatosEntrada.getFilesByName(NOMBRE_GSHEET_OLEAJE).next().getId());
    var datos = obtenerDatosDeMareaYOleaje(hojaOleaje, hoy, manana);

    var carpetaEstatica = DriveApp.getFolderById(CARPETA_DATOS_ESTATICOS_ID);
    var template = carpetaEstatica.getFilesByName(NOMBRE_TEMPLATE).next();

    var carpetaSalida = DriveApp.getFolderById(CARPETA_INFOGRAFIAS_ID);
    var copia = template.makeCopy("Marea " + fechaTexto, carpetaSalida);
    var presentacion = SlidesApp.openById(copia.getId());

    var slides = presentacion.getSlides();
    for (var i = 0; i < slides.length; i++) {
      for (var marcador in datos) {
        slides[i].replaceAllText(marcador, datos[marcador]);
      }
    }
    presentacion.saveAndClose();
    Logger.log("Infografía generada para " + fechaTexto);

    sheetControl.getRange("B2").setValue("✅ Infografía actualizada con éxito");
    sheetControl.getRange("B2").setFontColor("green");
    //sheetControl.getRange("B3").setValue(new Date());
    sheetControl.getRange("B3").setValue(Utilities.formatDate(new Date(), "America/El_Salvador", "dd/MM/yyyy hh:mm a").replace("AM", "a. m.").replace("PM", "p. m."));

  } catch (error) {
    Logger.log("❌ Error: " + error);
    sheetControl.getRange("B2").setValue("❌ Error: " + error.message);
    sheetControl.getRange("B2").setFontColor("red");
    sheetControl.getRange("B3").setValue("");
  }
}

function obtenerDatosDeMareaYOleaje(hojaOleaje, hoy, manana) {
  var carpetaDatosEntrada = DriveApp.getFolderById(CARPETA_DATOS_ENTRADA_ID);
  var archivoMarea = carpetaDatosEntrada.getFilesByName(NOMBRE_MAREA).next();
  var hojaMarea = SpreadsheetApp.open(archivoMarea).getSheetByName("mareas");

  var estaciones = [
  { nombre: 'la_union',     hoja: 'GOFO', columnaHora: 4, columnaAltura: 5 },
  { nombre: 'el_triunfo',   hoja: 'PCOR', columnaHora: 6, columnaAltura: 7 },
  { nombre: 'la_libertad',  hoja: 'COBA', columnaHora: 8, columnaAltura: 9 },
  { nombre: 'acajutla',     hoja: 'PCOC', columnaHora: 10, columnaAltura: 11 }
];


  var datos = {};
  var fechaHoy = Utilities.formatDate(hoy, 'America/El_Salvador', 'd MMM');
  var fechaManiana = Utilities.formatDate(manana, 'America/El_Salvador', 'd MMM');
  var fechaManianaCompleta = Utilities.formatDate(manana, 'America/El_Salvador', 'yyyy-MM-dd');

  var dias = ['domingo', 'lunes', 'martes', 'miércoles', 'jueves', 'viernes', 'sábado'];
  var meses = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio',
    'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'
  ];

  datos["{{fecha_hoy}}"] = dias[hoy.getDay()].charAt(0).toUpperCase() + dias[hoy.getDay()].slice(1) + " " + hoy.getDate() + " de " + meses[hoy.getMonth()];
  datos["{{fecha_maniana}}"] = (manana.getDate() + " DE " + meses[manana.getMonth()].toUpperCase());

  var totalFilas = hojaMarea.getRange("B:B").getValues().filter(function (f) {
    return f[0] && f[0].toString().trim() !== '';
  }).length;
  var rango = hojaMarea.getRange(13, 2, totalFilas - 12, 10).getValues();

  var mareas = {
    'Alta': {},
    'Baja': {}
  };
  var horasPorTipo = {
    'Alta': [],
    'Baja': []
  };

  for (var i = 0; i < rango.length; i++) {
    var fila = rango[i];
    var fechaStr = fila[0];
    if (!fechaStr) continue;

    var fechaDate = new Date(fechaStr);
    var fecha = Utilities.formatDate(fechaDate, 'America/El_Salvador', 'yyyy-MM-dd');
    var tipo = fila[1];

    if (fecha === fechaManianaCompleta && (tipo === 'Alta' || tipo === 'Baja')) {
      estaciones.forEach(function (estacion) {
        var filaHoja = 13 + i;
        var celdaHora = hojaMarea.getRange(filaHoja, estacion.columnaHora).getA1Notation();
        var celdaAltura = hojaMarea.getRange(filaHoja, estacion.columnaAltura).getA1Notation();

        var horaTexto = hojaMarea.getRange(celdaHora).getDisplayValue();
        var altura = hojaMarea.getRange(celdaAltura).getValue();
        Logger.log(`Fila ${filaHoja} | Estación: ${estacion.nombre} | ${tipo} | Hora (${celdaHora}): ${horaTexto} | Altura (${celdaAltura}): ${altura}`);

        if (!mareas[tipo][estacion.nombre]) mareas[tipo][estacion.nombre] = [];

        if (horaTexto && altura !== '-' && !isNaN(parseFloat(altura))) {
          var match = horaTexto.match(/(\d+):(\d+)/);
          if (match) {
            var hora = parseInt(match[1]);
            var minutos = parseInt(match[2]);
            var horaDate = new Date();
            horaDate.setHours(hora, minutos, 0, 0);

            var horaFormateada = Utilities.formatDate(horaDate, 'America/El_Salvador', 'h:mm a')
              .replace('AM', 'a.m.').replace('PM', 'p.m.');

            mareas[tipo][estacion.nombre].push({
              hora: horaFormateada,
              altura: parseFloat(altura).toFixed(1)
            });
            horasPorTipo[tipo].push(horaDate);
          }
        }
      });
    }
  }

estaciones.forEach(function (estacion) {
  ['Alta', 'Baja'].forEach(function (tipo) {
    var registros = mareas[tipo][estacion.nombre] || [];

    if (registros.length === 1) {
      const hojaValores = hojaMarea.getRange(13, 2, hojaMarea.getLastRow() - 12, 10).getValues();
      const fechaObjetivo = Utilities.formatDate(manana, 'America/El_Salvador', 'yyyy-MM-dd');

      // Recolectar horas de mareas del tipo y día para esta estación
      let mareasDelDia = [];

      hojaValores.forEach((fila) => {
        const fecha = fila[0];
        const tipoFila = (fila[1] || '').toString().toLowerCase().trim();
        const fechaStr = Utilities.formatDate(new Date(fecha), 'America/El_Salvador', 'yyyy-MM-dd');
        if (fechaStr !== fechaObjetivo || tipoFila !== tipo.toLowerCase()) return;

        const horaTexto = fila[estacion.columnaHora - 2];
        const alturaTexto = fila[estacion.columnaAltura - 2];

        if (horaTexto && alturaTexto !== '-' && !isNaN(parseFloat(alturaTexto))) {
          const horaDate = new Date(manana.toDateString() + ' ' + horaTexto);
          mareasDelDia.push({ hora: horaDate, altura: parseFloat(alturaTexto).toFixed(1) });
        }
      });

      // Ordenar cronológicamente
      mareasDelDia.sort((a, b) => a.hora - b.hora);

      // Comparar con la hora del único registro válido
      const horaUnica = new Date(manana.toDateString() + ' ' + registros[0].hora.replace('a.m.', 'AM').replace('p.m.', 'PM'));
      const esPrimera = (mareasDelDia.length && horaUnica.getTime() === mareasDelDia[0].hora.getTime());

      let respaldoAltura = null;
      let filaReferencia = null;

      if (esPrimera) {
        // Buscar segunda marea del día anterior
        const fechaAnterior = new Date(manana);
        fechaAnterior.setDate(manana.getDate() - 1);
        const fechaAnteriorStr = Utilities.formatDate(fechaAnterior, 'America/El_Salvador', 'yyyy-MM-dd');

        for (let i = hojaValores.length - 1; i >= 0; i--) {
          const fila = hojaValores[i];
          const fecha = fila[0];
          const tipoFila = (fila[1] || '').toString().toLowerCase().trim();
          if (!fecha || tipoFila !== tipo.toLowerCase()) continue;
          const fStr = Utilities.formatDate(new Date(fecha), 'America/El_Salvador', 'yyyy-MM-dd');
          if (fStr === fechaAnteriorStr) {
            const alt = fila[estacion.columnaAltura - 2];
            if (alt !== '-' && !isNaN(parseFloat(alt))) {
              respaldoAltura = parseFloat(alt).toFixed(1);
              filaReferencia = i + 13;
              Logger.log(`ℹ️ Marea faltante (primera ${tipo}) para ${estacion.nombre} completada con fila ${filaReferencia} del día anterior: altura = ${respaldoAltura}`);
              break;
            }
          }
        }
      } else {
        // Buscar primera marea del día siguiente
        const fechaSiguiente = new Date(manana);
        fechaSiguiente.setDate(manana.getDate() + 1);
        const fechaSiguienteStr = Utilities.formatDate(fechaSiguiente, 'America/El_Salvador', 'yyyy-MM-dd');

        for (let i = 0; i < hojaValores.length; i++) {
          const fila = hojaValores[i];
          const fecha = fila[0];
          const tipoFila = (fila[1] || '').toString().toLowerCase().trim();
          if (!fecha || tipoFila !== tipo.toLowerCase()) continue;
          const fStr = Utilities.formatDate(new Date(fecha), 'America/El_Salvador', 'yyyy-MM-dd');
          if (fStr === fechaSiguienteStr) {
            const alt = fila[estacion.columnaAltura - 2];
            if (alt !== '-' && !isNaN(parseFloat(alt))) {
              respaldoAltura = parseFloat(alt).toFixed(1);
              filaReferencia = i + 13;
              Logger.log(`ℹ️ Marea faltante (segunda ${tipo}) para ${estacion.nombre} completada con fila ${filaReferencia} del día siguiente: altura = ${respaldoAltura}`);
              break;
            }
          }
        }
      }

      if (respaldoAltura !== null) {
        registros.push({
          hora: '12:00 a.m.',
          altura: respaldoAltura
        });
      }
    }

    // Asegurar siempre dos registros completos
    for (let i = 0; i < 2; i++) {
      const registro = registros[i];
      const claveHora = `{{${estacion.nombre}_${tipo.toLowerCase()}${i + 1}_hora}}`;
      const claveAltura = `{{${estacion.nombre}_${tipo.toLowerCase()}${i + 1}_altura}}`;
      if (registro) {
        datos[claveHora] = registro.hora;
        datos[claveAltura] = registro.altura;
      }
    }
  });
});


  ['Alta', 'Baja'].forEach(function (tipo) {
    const horas = horasPorTipo[tipo];
    if (horas.length >= 4) {
      const horasOrdenadas = horas.sort((a, b) => a - b);
      const primera = horasOrdenadas.slice(0, 2);
      const segunda = horasOrdenadas.slice(-2);

      const min1 = new Date(Math.min(...primera.map(d => d.getTime())));
      const max2 = new Date(Math.max(...segunda.map(d => d.getTime())));

      function redondearHoraMasCercana(date) {
        const h = date.getHours();
        const m = date.getMinutes();
        return (m >= 30) ? h + 1 : h;
      }

      const horaIni = redondearHoraMasCercana(min1);
      const horaFin = redondearHoraMasCercana(max2);

      const horaIniStr = Utilities.formatDate(new Date(2000, 0, 1, horaIni), 'America/El_Salvador', 'h:00');
      const horaFinStr = Utilities.formatDate(new Date(2000, 0, 1, horaFin), 'America/El_Salvador', 'h:00');

      if (horaIniStr === horaFinStr) {
        datos[`{{rango_marea_${tipo.toLowerCase()}}}`] =
          `Alrededor de ${horaIniStr} a. m. y p. m.`;
      } else {
        datos[`{{rango_marea_${tipo.toLowerCase()}}}`] =
          `Entre ${horaIniStr} y ${horaFinStr} a. m. y p. m.`;
      }
    }
  });

  estaciones.forEach(function (estacion) {
    var hoja = hojaOleaje.getSheetByName(estacion.hoja);
    var datosHoja = hoja.getRange("A2:H" + hoja.getLastRow()).getValues();
    var alturas = datosHoja.filter(function (r) {
      return r[0] && r[0].toString().toLowerCase().trim() === fechaHoy.toLowerCase();
    }).map(function (r) {
      return parseFloat(r[3]);
    });
    var maxAltura = alturas.length ? Math.max.apply(null, alturas).toFixed(1) : '';
    datos[`{{${estacion.nombre}_hoy}}`] = maxAltura;
  });

  estaciones.forEach(function (estacion) {
    var hoja = hojaOleaje.getSheetByName(estacion.hoja);
    var datosHoja = hoja.getRange("A2:H" + hoja.getLastRow()).getValues();
    var filas = datosHoja.filter(function (r) {
      return r[0] && r[0].toString().toLowerCase().trim() === fechaManiana.toLowerCase();
    });
    var alturas = filas.map(function (r) {
      return parseFloat(r[3]);
    });
    var periodos = filas.map(function (r) {
      return parseFloat(r[4]) * 3;
    });
    var maxAltura = alturas.length ? Math.max.apply(null, alturas).toFixed(1) : '';
    var maxRapidez = periodos.length ?
      Math.ceil(Math.max.apply(null, periodos) / 5) * 5 :
      '';
    datos[`{{${estacion.nombre}_maniana_altura}}`] = maxAltura;
    datos[`{{${estacion.nombre}_maniana_rapidez}}`] = maxRapidez.toString();
  });

  Logger.log("Marcadores generados:\n" + JSON.stringify(datos, null, 2));
  return datos;
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Infografía')
    .addItem('🌊 Actualizar Infografía', 'actualizarInfografia')
    .addItem('📂 Ver carpeta de infografías', 'abrirCarpetaInfografias')
    .addItem('📤 Subir archivo de oleaje', 'mostrarFormularioCarga')
    .addToUi();
}

function abrirCarpetaInfografias() {
  var url = "https://drive.google.com/drive/folders/1Njsi_ZgLepGW84M25fC288L_k1dBfEkl"; // ID de tu carpeta
  var html = HtmlService.createHtmlOutput('<script>window.open("' + url + '", "_blank");google.script.host.close();</script>')
    .setWidth(100)
    .setHeight(50);
  SpreadsheetApp.getUi().showModalDialog(html, "Abriendo carpeta...");
}

function mostrarFormularioCarga() {
  var html = HtmlService.createHtmlOutputFromFile('SubirOleaje')
    .setWidth(400)
    .setHeight(200);
  SpreadsheetApp.getUi().showModalDialog(html, 'Subir archivo Oleaje_Viento.xlsx');
}

function doPost(e) {
  try {
    const nombreArchivo = e.parameter.nombreArchivo;
    const contenidoBase64 = e.parameter.contenidoBase64;

    const carpetaEntradaId = '1DKO73DbITEiWB6thpZQB4hFvMzA26t19'; // Datos de entrada
    const carpetaHistoricoId = '18gWBZ9JfpkmmweGkt4fYkugRO3YFzMFv'; // Datos históricos

    // Buscar y mover versiones anteriores
    const archivos = Drive.Files.list({
      q: `'${carpetaEntradaId}' in parents and (title = 'Oleaje_Viento.xlsx' or title = 'Oleaje_Viento') and trashed = false`
    });

    if (archivos.items && archivos.items.length > 0) {
      const ayer = new Date();
      ayer.setDate(ayer.getDate() - 1);
      const fechaAyer = Utilities.formatDate(ayer, Session.getScriptTimeZone(), 'dd-MM-yyyy');

      archivos.items.forEach(file => {
        const nuevoNombre = file.title + ' (antiguo ' + fechaAyer + ')';
        Drive.Files.update({
          title: nuevoNombre,
          parents: [{
            id: carpetaHistoricoId
          }]
        }, file.id);
      });
    }

    // Subir el nuevo archivo y convertirlo a Google Sheets
    const blob = Utilities.newBlob(
      Utilities.base64Decode(contenidoBase64),
      MimeType.MICROSOFT_EXCEL,
      nombreArchivo
    );

    Drive.Files.insert({
      title: 'Oleaje_Viento',
      mimeType: MimeType.GOOGLE_SHEETS,
      parents: [{
        id: carpetaEntradaId
      }]
    }, blob);

    return ContentService.createTextOutput("✅ Archivo subido y versiones anteriores archivadas.");
  } catch (error) {
    return ContentService.createTextOutput("❌ Error: " + error.message);
  }
}

function doGet() {
  return ContentService.createTextOutput("✅ Web App activo");
}