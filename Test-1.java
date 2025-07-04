function onEdit(e) {
  const sheet = e.source.getActiveSheet();
  const editedCell = e.range;
  const editedColumn = editedCell.getColumn();
  const editedRow = e.range.getRow();
  const sheetName = sheet.getName().trim();

  const indiceSheet = e.source.getSheetByName("Indice");
  const expediente = editedCell.getValue();

  // Agregar marca desde celda L1 del índice
  if (sheetName === "Indice" && editedCell.getA1Notation() === "L1") {
    const valores = expediente.split(";").map(x => x.trim());
    if (valores.length < 7) {
      SpreadsheetApp.getUi().alert("Faltan datos. Debes ingresar al menos 7 campos separados por comas.");
      return;
    }

    const nuevaFila = indiceSheet.getLastRow() + 1;
    indiceSheet.getRange(`B${nuevaFila}:H${nuevaFila}`).setValues([valores.slice(0, 7)]);
    indiceSheet.getRange("L1").clearContent();
    return;
  }

  if (!expediente || editedColumn !== 3) return;

  const indiceData = indiceSheet.getRange("C3:I" + indiceSheet.getLastRow()).getValues();
  const matchIndex = indiceData.findIndex(row => row[0] == expediente);
  if (matchIndex === -1) return;

  const [
    exp,
    _registro,
    clase,
    titular,
    direccion,
    cp,
    correo
  ] = indiceData[matchIndex];

  const titulo = indiceSheet.getRange(`B${matchIndex + 3}`).getValue();
  const ui = SpreadsheetApp.getUi();
  const indiceRow = matchIndex + 3;

  const hojas = ["Marcas en proceso", "Marcas Concedidas", "Marcas Impedimento Legal"];
  hojas.forEach(nombre => {
    if (nombre !== sheetName) {
      const hoja = e.source.getSheetByName(nombre);
      if (!hoja) return;
      const datos = hoja.getRange("C3:C" + hoja.getLastRow()).getValues();
      const fila = datos.findIndex(row => row[0] == expediente);
      if (fila !== -1) {
        hoja.deleteRow(fila + 3);
      }
    }
  });

  if (sheetName === "Marcas en proceso") {
    const fechaRegistroPrompt = ui.prompt("Ingresa la fecha de registro (dd/mm/aaaa):");
    if (fechaRegistroPrompt.getSelectedButton() === ui.Button.OK) {
      const fechaRegistro = fechaRegistroPrompt.getResponseText();
      sheet.getRange(editedRow, 2).setValue(titulo);
      sheet.getRange(editedRow, 4).setValue(clase);
      sheet.getRange(editedRow, 5).setValue(fechaRegistro);
      indiceSheet.getRange(`B${indiceRow}:J${indiceRow}`).setBackground('#b4a7d6');
    }
  }

  if (sheetName === "Marcas Concedidas") {
    const registroPrompt = ui.prompt("Ingresa el número de registro:");
    const fechaPrompt = ui.prompt("Ingresa la fecha de concesión (dd/mm/aaaa):");

    if (
      registroPrompt.getSelectedButton() === ui.Button.OK &&
      fechaPrompt.getSelectedButton() === ui.Button.OK
    ) {
      const numRegistro = registroPrompt.getResponseText();
      const fechaConcesion = fechaPrompt.getResponseText();

      sheet.getRange(editedRow, 2).setValue(titulo);
      sheet.getRange(editedRow, 4).setValue(numRegistro);
      sheet.getRange(editedRow, 5).setValue(clase);
      sheet.getRange(editedRow, 6).setValue(fechaConcesion);
      indiceSheet.getRange(`D${indiceRow}`).setValue(numRegistro); // Actualiza número de registro en el índice
      indiceSheet.getRange(`B${indiceRow}:J${indiceRow}`).setBackground('#93c47d');
    }
  }

  if (sheetName === "Marcas Impedimento Legal") {
    const fechaOficioPrompt = ui.prompt("Ingresa la fecha del Término Para Contestar El Oficio:");
    if (fechaOficioPrompt.getSelectedButton() === ui.Button.OK) {
      const fechaOficio = fechaOficioPrompt.getResponseText();
      sheet.getRange(editedRow, 2).setValue(titulo);
      sheet.getRange(editedRow, 4).setValue(clase);
      sheet.getRange(editedRow, 5).setValue(fechaOficio);
      indiceSheet.getRange(`B${indiceRow}:J${indiceRow}`).setBackground('#ffd966');
    }
  }
}

function verificarFechasTodas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  pintarFechasEnHoja(ss, "Marcas en proceso", ["F", "G", "H"]);
  pintarFechasEnHoja(ss, "Marcas Concedidas", ["G", "H"], true);
  pintarFechasEnHoja(ss, "Marcas Impedimento Legal", ["E"], true);
  pintarFechasEnHoja(ss, "Marcas Impedimento Legal", ["G", "H", "I"]);
}

function pintarFechasEnHoja(ss, nombreHoja, columnas, esEspecial = false) {
  const hoja = ss.getSheetByName(nombreHoja);
  if (!hoja) return;

  const hoy = new Date();
  hoy.setHours(0, 0, 0, 0);

  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 3) return;

  columnas.forEach(letra => {
    const rango = hoja.getRange(`${letra}3:${letra}${ultimaFila}`);
    const valores = rango.getValues();
    const backgrounds = rango.getBackgrounds();

    for (let i = 0; i < valores.length; i++) {
      const celda = valores[i][0];
      let color = "#ffffff";

      if (celda instanceof Date) {
        const fecha = new Date(celda);
        fecha.setHours(0, 0, 0, 0);
        const dias = Math.floor((fecha - hoy) / (1000 * 60 * 60 * 24));

        if (esEspecial) {
          if (dias < 0) color = "#93c47d";
          else if (dias === 0) color = "#e06666";
          else if (dias <= 30) color = "#e06666";
          else if (dias <= 60) color = "#ff9900";
          else color = "#ffd966";
        } else {
          if (dias === 0) color = "#e06666";
          else if (dias > 0 && dias <= 7) color = "#ff6d01";
          else if (dias > 7) color = "#ffd966";
          else if (dias < 0) color = "#93c47d";
        }
      }

      backgrounds[i][0] = color;
    }

    rango.setBackgrounds(backgrounds);
  });
}

function enviarReportesMarcas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaIndice = ss.getSheetByName("Indice");
  const hojaProceso = ss.getSheetByName("Marcas en proceso");
  const hojaImpedimento = ss.getSheetByName("Marcas Impedimento Legal");
  const hoy = new Date();
  hoy.setHours(0, 0, 0, 0);

  const docTemplate2M = "1Da6gDo3W04Oa59QHAgPXkbHyOimf8uc6ySvpMTleAlY";
  const docTemplate4M = "1X4hiad4t0MzhTJeK7-WdtBcnXGYQ0PUjz2Dq8QBVHKY";
  const docTemplateImp2M = "1FcVCufjib4FBUIGg_tCKiYwlGTfZJ3dfbYy6w8Rf834";

  const datosIndice = hojaIndice.getDataRange().getValues();

  function reemplazarDatos(plantilla, datos) {
    let body = plantilla.getBody();
    Object.entries(datos).forEach(([clave, valor]) => {
      body.replaceText(clave, valor);
    });
  }

  function formatearFecha(date) {
    const meses = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"];
    const dia = date.getDate();
    const mes = meses[date.getMonth()];
    const año = date.getFullYear();
    return `${dia} de ${mes} de ${año}`;
  }

  function obtenerDatosDesdeIndice(expediente) {
    for (let i = 2; i < datosIndice.length; i++) {
      if (datosIndice[i][2] == expediente) {
        const titular = datosIndice[i][5];
        const direccionCruda = datosIndice[i][6];
        const cp = datosIndice[i][7];
        const registro = datosIndice[i][3];

        let direccionFormateada = direccionCruda.trim();
        direccionFormateada = direccionFormateada.replace(/(\d+[^,]*,)/, '$1\n');
        direccionFormateada += "\nC.P. " + cp + " – MÉXICO";

        const datosImpedimento = hojaImpedimento.getDataRange().getValues();
        let fechaContestacion = "";
        for (let j = 2; j < datosImpedimento.length; j++) {
          if (datosImpedimento[j][2] == expediente) {
            const fecha = datosImpedimento[j][5];
            if (fecha instanceof Date) {
              fechaContestacion = formatearFecha(fecha);
            }
            break;
          }
        }

        return {
          "<<FECHA ACTUAL>>": formatearFecha(hoy),
          "<<TITULAR>>": titular,
          "<<DIRECCIÓN>>": direccionFormateada,
          "<<C.P>>": cp,
          "<<TÍTULO DE MARCA>>": datosIndice[i][1],
          "<<CLASE>>": datosIndice[i][4],
          "<<PRIMER NOMBRE>>": titular.split(" ")[0],
          "<<EXPEDIENTE>>": datosIndice[i][2],
          "<<REGISTRO>>": registro || "",
          "<<FECHA CONTESTACIÓN>>": fechaContestacion
        };
      }
    }
    return null;
  }

  function generarYEnviar(nombre, correo, datos, celdaFecha, plantillaId) {
    const doc = DriveApp.getFileById(plantillaId).makeCopy(nombre);
    const docFile = DocumentApp.openById(doc.getId());
    reemplazarDatos(docFile, datos);
    docFile.saveAndClose();

    const pdf = doc.getAs(MimeType.PDF);
    MailApp.sendEmail({
      to: correo,
      subject: nombre,
      body: "Adjunto encontrarás el reporte de seguimiento de tu marca.",
      attachments: [pdf]
    });

    celdaFecha.setBackground("#93c47d");
    doc.setTrashed(true);
  }

  function verificarFechasYEnviar(hoja, columnas, tipo, plantillaIds) {
    const datos = hoja.getDataRange().getValues();
    for (let i = 2; i < datos.length; i++) {
      columnas.forEach((col, idx) => {
        const fecha = datos[i][col];
        const celda = hoja.getRange(i + 1, col + 1);
        const colorCelda = celda.getBackground();

        if (fecha instanceof Date) {
          fecha.setHours(0, 0, 0, 0);
          if (fecha.getTime() === hoy.getTime() && colorCelda !== "#93c47d") {
            const expediente = datos[i][2];
            const datosDoc = obtenerDatosDesdeIndice(expediente);
            const correo = datosIndice.find(row => row[2] === expediente)?.[8];
            if (datosDoc && correo) {
              const nombre = `${tipo} ${idx === 0 ? "2 Meses" : "4 Meses"}`;
              const plantillaId = plantillaIds[idx];
              generarYEnviar(nombre, correo, datosDoc, celda, plantillaId);
            }
          }
        }
      });
    }
  }

  verificarFechasYEnviar(hojaProceso, [5, 6], "Reporte Seguimiento De Marca", [docTemplate2M, docTemplate4M]);
  verificarFechasYEnviar(hojaImpedimento, [7, 8], "Reporte Seguimiento De Requerimiento De Marca", [docTemplateImp2M, docTemplate4M]);
}
