const boton = document.querySelector("#escribir");
const estado = document.querySelector("#estado");

function decir(texto) {
  estado.textContent = texto;
}

Office.onReady(function (info) {
  if (!info || info.host !== Office.HostType.Excel) {
    decir("Este panel solo funciona dentro de Excel.");
    return;
  }
  boton.disabled = false;
  decir("Listo. Pulsa el botón para escribir en A1.");
  boton.addEventListener("click", escribir);
});

function escribir() {
  boton.disabled = true;
  decir("Escribiendo en la hoja…");
  Excel.run(function (context) {
    const hoja = context.workbook.worksheets.getActiveWorksheet();
    const celda = hoja.getRange("A1");
    celda.values = [["El panel escribió en la hoja."]];
    celda.format.fill.color = "#FFF2CC";
    celda.format.font.bold = true;
    return context.sync();
  }).then(function () {
    decir("Listo. La celda A1 ya tiene la frase.");
    boton.disabled = false;
  }).catch(function (error) {
    decir("No se pudo escribir en la hoja. " + (error.message || ""));
    boton.disabled = false;
  });
}
