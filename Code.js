function doGet(e) {

  // Si viene ?pdf=1 → entregar PDF
  if (e && e.parameter && e.parameter.pdf === "1") {
    return entregarPDFRanking();
  }

  // Si no → mostrar la página
  return HtmlService.createTemplateFromFile("ui")
    .evaluate()
    .setTitle("Participantes");
}
function include(file) {
  return HtmlService.createHtmlOutputFromFile(file).getContent();
}

function getPagina(nombre) {
  return HtmlService.createHtmlOutputFromFile(nombre).getContent();
}

function getAllData() {
  const ss = SpreadsheetApp.openById("1hCkFXrG1TEPer80ayQsgHigpDNb0XS9Oj7zSZ4qQjFI");
  const sh = ss.getSheetByName("Base");
  const lastRow = sh.getLastRow();
  if (lastRow < 3) return [];

  const data = sh.getRange(3, 1, lastRow - 2, 8).getValues();
  return data.map(row => ({
    id: String(row[0] || "").trim(),
    apellidos: row[1],
    nombres: row[2],
    grado: row[3],
    ie: row[4],
    categoria: row[5],
    puntaje: row[6],
    hora: row[7]
  }));
}

function agregarRegistro(data) {
  const ss = SpreadsheetApp.openById("1hCkFXrG1TEPer80ayQsgHigpDNb0XS9Oj7zSZ4qQjFI");
  const sh = ss.getSheetByName("Base");

  // Agregar los valores excepto la hora y puntaje
  const fila = sh.getLastRow() + 1;
  sh.getRange(fila, 1).setValue(data.id);
  sh.getRange(fila, 2).setValue(data.apellidos);
  sh.getRange(fila, 3).setValue(data.nombres);
  sh.getRange(fila, 4).setValue(data.grado);
  sh.getRange(fila, 5).setValue(data.ie);
  sh.getRange(fila, 6).setValue(data.categoria);

  // Puntaje: solo si hay valor, si no dejar vacío
  if (data.puntaje && data.puntaje.toString().trim() !== "") {
    sh.getRange(fila, 7).setValue(Number(data.puntaje));
  } else {
    sh.getRange(fila, 7).setValue(""); // celda vacía
  }

  // Guardar la hora como TEXTO
  const rangoHora = sh.getRange(fila, 8); // columna H
  rangoHora.setNumberFormat('@'); // forzar texto
  rangoHora.setValue(data.hora);
}

function editarRegistroSheet(data) {
  const ss = SpreadsheetApp.openById("1hCkFXrG1TEPer80ayQsgHigpDNb0XS9Oj7zSZ4qQjFI");
  const sh = ss.getSheetByName("Base");
  const lastRow = sh.getLastRow();
  if (lastRow < 1) return;

  const valores = sh.getRange(1, 1, lastRow, 8).getValues();
  for (let i = 0; i < valores.length; i++) {
    if (valores[i][0] === data.id) {
      // Puntaje: solo si hay valor
      const puntajeFinal = (data.puntaje && data.puntaje.toString().trim() !== "") ? Number(data.puntaje) : "";

      // Actualizar las columnas excepto la hora
      sh.getRange(i + 1, 2, 1, 6).setValues([[
        data.apellidos, data.nombres, data.grado,
        data.ie, data.categoria, puntajeFinal
      ]]);

      // Actualizar la hora como TEXTO
      const rangoHora = sh.getRange(i + 1, 8); // columna H
      rangoHora.setNumberFormat('@'); // forzar texto
      rangoHora.setValue(data.hora);

      break;
    }
  }
}

function eliminarRegistro(id) {
  const ss = SpreadsheetApp.openById("1hCkFXrG1TEPer80ayQsgHigpDNb0XS9Oj7zSZ4qQjFI");
  const sh = ss.getSheetByName("Base");
  const lastRow = sh.getLastRow();
  if (lastRow < 3) return;

  const valores = sh.getRange(3, 1, lastRow - 2, 1).getValues();
  for (let i = 0; i < valores.length; i++) {
    if (valores[i][0] === id) {
      sh.deleteRow(i + 3); // +3 porque los datos empiezan en la fila 3
      break;
    }
  }
}

function getOpciones() {
  const ss = SpreadsheetApp.openById("1hCkFXrG1TEPer80ayQsgHigpDNb0XS9Oj7zSZ4qQjFI");
  const sh = ss.getSheetByName("Configuracion");
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { categorias: [], grados: [], instituciones: [] };

  const data = sh.getRange(2, 1, lastRow - 1,4).getValues();
  return {
    categorias: [...new Set(data.map(r => r[0]).filter(x => x))].sort(),
    grados: [...new Set(data.map(r => r[1]).filter(x => x))],
    instituciones: [...new Set(data.map(r => r[2]).filter(x => x))].sort(),
    configuraciones: data.map(r => r[3])
  };
}
