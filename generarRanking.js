function validarClaveServidor(clave) {
  return clave === "mamiluz";  // Cambia tu clave aquí
}

// Genera PDF con ranking de alumnos y devuelve URL de Drive
function generarRankingPdfWeb() {

  const diccionarioGrados = {
    "Inicial 3ro": "Inicial 3ro",
    "Inicial 4to": "Inicial 4to",
    "Inicial 5to": "Inicial 5to",
    "1ro de Prim.": "1ro de Primaria",
    "2do de Prim.": "2do de Primaria",
    "3ro de Prim.": "3ro de Primaria",
    "4to de Prim.": "4to de Primaria",
    "5to de Prim.": "5to de Primaria",
    "6to de Prim.": "6to de Primaria",
    "1ro de Sec.": "1ro de Secundaria",
    "2do de Sec.": "2do de Secundaria",
    "3ro de Sec.": "3ro de Secundaria",
    "4to de Sec.": "4to de Secundaria",
    "5to de Sec.": "5to de Secundaria"
  };

  const ss = SpreadsheetApp.openById("1hCkFXrG1TEPer80ayQsgHigpDNb0XS9Oj7zSZ4qQjFI");
  const sh = ss.getSheetByName("Base");
  const data = sh.getRange(3, 1, sh.getLastRow() - 2, sh.getLastColumn()).getValues();

  const opciones = getOpciones();
  const grados = opciones.grados;
  const categorias = opciones.categorias;
  const configuraciones = opciones.configuraciones;
  Logger.log(configuraciones)

  // Map para agrupar por grado y categoría
  const mapResultados = {};
  data.forEach(r => {
    const grado = r[3];
    const categoria = r[5];
    if (!mapResultados[grado]) mapResultados[grado] = {};
    if (!mapResultados[grado][categoria]) mapResultados[grado][categoria] = [];
    mapResultados[grado][categoria].push(r);
  });

  const resultados = [];
  grados.forEach(g => {
    categorias.forEach(c => {
      const subset = (mapResultados[g]?.[c] || []).sort((a, b) => {
        // Ordenar por puntaje (descendente)
        const diffPuntos = b[6] - a[6];
        if (diffPuntos !== 0) return diffPuntos;

        // Si empatan, ordenar por hora de entrega (ascendente → primero el más temprano)
        const tiempoA = horaASegundos(a[7]);
        const tiempoB = horaASegundos(b[7]);
        return tiempoA - tiempoB;
      });
      resultados.push({ grado: g, categoria: c, rows: subset });
    });
  });

  const logo = DriveApp.getFileById('1d3Ph4dSFR7zigftsszW60Cgrg_qoXGCp');
  const base64Logo = Utilities.base64Encode(logo.getBlob().getBytes());
  const imgSrc = `data:image/png;base64,${base64Logo}`;

  let html = `
<html><head><meta charset="UTF-8">
<style>
@page { size: A4; margin: 1cm 1.5cm 1cm 1cm; }
body { font-family: Arial, sans-serif; margin: 0; padding: 0; }
h1 { text-align: center; margin: 0; }
table { width: 100%; border-collapse: collapse; margin-bottom: 12px; }
th { border: 1px solid #000; padding: 1px 2px; font-size: 10pt; font-weight: bold; text-align: center; background-color: #f0f0f0; }
td { border: 1px solid #000; padding: 1px 2px; font-size: 8pt; text-align: center; }
.page-break { page-break-after: always; }
.footer { position: fixed; bottom: 0; left: 0; width: 100%; border-top: 1px solid #000; text-align: right; font-size: 8pt; padding-top: 2px; padding-right: 5px; }
</style>
</head><body>`;

  resultados.forEach((r) => {
    const nombreGrado = diccionarioGrados[r.grado] || r.grado;
    const fechaActual = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");
    const horaActual = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "HH:mm:ss");

    let agregarEncabezado = ` <div style="position: relative; width: 100%; height: 55px; margin-bottom: 5px;">
    <img src="${imgSrc}" alt="Logo Colegio" style="position: absolute; top: 0; left: 0; width: 120px; height: auto; z-index: -1">
    <div style="position: absolute; top: 0; right: 0; text-align: right; font-size: 9pt;">
      <div>Fecha: ${fechaActual}</div>
      <div>Hora: ${horaActual}</div>
    </div>
  </div>

  <h1 style="font-size: 20pt; font-weight: bold; color: #800000; margin-bottom:2px;">Complejo Educativo Cesar's</h1>`

    if (configuraciones[1] === false) {
      agregarEncabezado = ``
    }

    html += generarSaltosLinea(configuraciones[0]);
    html += agregarEncabezado + `
 
  <h1 style="font-size: 14pt; font-weight: normal; color: #000; margin-bottom:2px;">Resultados de ${nombreGrado} - Categoría: ${r.categoria}</h1>

  <div style="border-top: 1px solid #000; margin: 5px 0;"></div>

  <table>
    <tr><th>#</th><th>Nombre</th><th>Grado</th><th>Institución</th><th>Categoría</th><th>Puntaje</th><th>Hora Entrega</th></tr>`;

    r.rows.forEach((row, i) => {
      html += `<tr>
      <td>${i + 1}°</td>
      <td>${row[1]} ${row[2]}</td>
      <td>${row[3]}</td>
      <td>${row[4]}</td>
      <td>${row[5]}</td>
      <td>${row[6]}</td>
      <td>${row[7]}</td>
      </tr>`;
    });

    html += `</table>`;
    html += `<div class="footer"></div><div class="page-break"></div>`;
  });

  html += `</body></html>`;

  // Guardar PDF en Drive
  const archivo = DriveApp.getFileById(ss.getId());
  const carpetaPadre = archivo.getParents().next();
  const subcarpeta = carpetaPadre.getFoldersByName("Rankings Generados").hasNext()
    ? carpetaPadre.getFoldersByName("Rankings Generados").next()
    : carpetaPadre.createFolder("Rankings Generados");

  // Guardar PDF
  const nombrePDF = `Ranking_${new Date().getTime()}.pdf`;
  const pdfBlob = Utilities.newBlob(html, "text/html", "temp.html")
    .getAs("application/pdf")
    .setName(nombrePDF);

  const pdfFile = subcarpeta.createFile(pdfBlob);
  Utilities.sleep(2000); // 2 segundos
  return pdfFile.getUrl();
}

function generarSaltosLinea(cantidad) {
  let saltos = '';
  for (let i = 0; i < cantidad; i++) {
    saltos += '<br>';
  }
  return saltos;
}


function horaASegundos(hora) {
  const [h, m, s] = hora.split(":").map(Number);
  return h * 3600 + m * 60 + s;
}
