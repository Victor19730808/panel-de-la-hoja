const boton = document.querySelector("#cargar");
const estado = document.querySelector("#estado");

const API =
  "https://concanacopc-c9btascxdzfxaqce.canadacentral-01.azurewebsites.net/api/get_promociones_establecimientos";

const COLUMNAS = [
  "Título",
  "Establecimiento",
  "Categoría",
  "Estado",
  "Tipo",
  "Descripción",
  "Desde",
  "Hasta",
  "Id",
];

function decir(texto) {
  estado.textContent = texto;
}

Office.onReady(function (info) {
  if (!info || info.host !== Office.HostType.Excel) {
    decir("Este panel solo funciona dentro de Excel.");
    return;
  }
  boton.disabled = false;
  decir("Listo. Pulsa el botón para traer las promociones.");
  boton.addEventListener("click", cargar);
});

function cargar() {
  boton.disabled = true;
  decir("Consultando promociones…");
  traerPromociones()
    .then(function (items) {
      decir("Escribiendo " + items.length + " promociones…");
      return escribir(items);
    })
    .then(function (cantidad) {
      decir("Listo. La hoja tiene " + cantidad + " promociones.");
      boton.disabled = false;
    })
    .catch(function (error) {
      decir("No se pudieron cargar las promociones. " + (error.message || ""));
      boton.disabled = false;
    });
}

function traerPromociones() {
  const pageSize = 100;
  let page = 1;
  let totalPages = 1;
  const todas = [];

  function siguiente() {
    const url =
      API +
      "?boundle_id=2&salida=1&page=" +
      page +
      "&pageSize=" +
      pageSize;
    return fetch(url, { headers: { Accept: "application/json" } })
      .then(function (resp) {
        if (!resp.ok) {
          throw new Error("La API respondió " + resp.status);
        }
        return resp.json();
      })
      .then(function (data) {
        const items = data.items || [];
        todas.push.apply(todas, items);
        totalPages = data.totalPages || 1;
        page += 1;
        if (page <= totalPages && page <= 8) {
          return siguiente();
        }
        return todas;
      });
  }

  return siguiente();
}

function escribir(items) {
  const filas = [COLUMNAS];
  items.forEach(function (item) {
    filas.push([
      texto(item.promocion_titulo),
      texto(item.establecimiento_nombre),
      texto(item.categoria_nombre),
      texto(item.direccion_estado),
      texto(item.establecimiento_tipo),
      recortar(item.promocion_descripcion),
      fecha(item.promocion_fi),
      fecha(item.promocion_ff),
      item.promocion_id == null ? "" : String(item.promocion_id),
    ]);
  });

  return Excel.run(function (context) {
    const hoja = context.workbook.worksheets.getActiveWorksheet();
    const rango = hoja.getRangeByIndexes(0, 0, filas.length, COLUMNAS.length);
    rango.values = filas;
    const encabezado = hoja.getRangeByIndexes(0, 0, 1, COLUMNAS.length);
    encabezado.format.fill.color = "#FFF2CC";
    encabezado.format.font.bold = true;
    if (filas.length > 1) {
      const cuerpo = hoja.getRangeByIndexes(1, 0, filas.length - 1, COLUMNAS.length);
      cuerpo.format.fill.color = "#FFFFFF";
      cuerpo.format.font.bold = false;
    }
    hoja.freezePanes.freezeRows(1);
    return context.sync().then(function () {
      return items.length;
    });
  });
}

function texto(valor) {
  if (valor == null) return "";
  return String(valor).replace(/\s+/g, " ").trim();
}

function recortar(valor) {
  const limpio = texto(valor);
  if (limpio.length <= 180) return limpio;
  return limpio.slice(0, 177) + "...";
}

function fecha(valor) {
  const limpio = texto(valor);
  if (limpio.length >= 10) return limpio.slice(0, 10);
  return limpio;
}
