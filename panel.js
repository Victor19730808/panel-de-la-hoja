const combo = document.querySelector("#bancos");
const estado = document.querySelector("#estado");

const API = "https://127.0.0.1:3443/sp/cliente?accion=listar";

function decir(texto) {
  estado.textContent = texto;
}

Office.onReady(function (info) {
  if (!info || info.host !== Office.HostType.Excel) {
    decir("Este panel solo funciona dentro de Excel.");
    return;
  }
  decir("Consultando bancos…");
  traerBancos()
    .then(function (filas) {
      llenar(filas);
      decir("Listo. Hay " + filas.length + " bancos.");
    })
    .catch(function (error) {
      combo.replaceChildren();
      const opcion = document.createElement("option");
      opcion.value = "";
      opcion.textContent = "No se pudieron cargar";
      combo.appendChild(opcion);
      decir("No se pudieron cargar los bancos. " + (error.message || ""));
    });
});

function traerBancos() {
  return fetch(API, { headers: { Accept: "application/json" } }).then(function (resp) {
    if (!resp.ok) {
      throw new Error("La API respondió " + resp.status);
    }
    return resp.json();
  }).then(function (data) {
    if (!data || data.ok === false) {
      throw new Error((data && data.mensaje) || "La API no devolvió filas.");
    }
    return data.filas || [];
  });
}

function llenar(filas) {
  combo.replaceChildren();
  const vacio = document.createElement("option");
  vacio.value = "";
  vacio.textContent = "Elige un banco";
  combo.appendChild(vacio);
  filas.forEach(function (fila) {
    const opcion = document.createElement("option");
    opcion.value = fila.id == null ? "" : String(fila.id);
    opcion.textContent = fila.name == null ? "" : String(fila.name);
    combo.appendChild(opcion);
  });
  combo.disabled = false;
}
