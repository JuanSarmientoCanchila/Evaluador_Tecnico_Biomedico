const entrada = document.getElementById("archivo-csv");
const mensajes = document.getElementById("mensajes");
const resumen = document.getElementById("resumen");
const contenedor = document.getElementById("criterios");

const OPCIONES = [
  { valor: "", texto: "—" },
  { valor: "0", texto: "0 · No cumple" },
  { valor: "1", texto: "1 · Cumple" },
  { valor: "2", texto: "2 · Cumple por encima" },
];

const botonExportar = document.getElementById("exportar");

let equipos = []; // criterios cargados, cada equipo con sus ofertas
let hayCambios = false; // true si hay trabajo sin exportar

// Avisa antes de recargar o cerrar la pestaña si hay trabajo sin exportar
window.addEventListener("beforeunload", (evento) => {
  if (hayCambios) {
    evento.preventDefault();
    evento.returnValue = "";
  }
});

const botonWord = document.getElementById("exportar-word");

// Datos generales del informe (también se guardan en el Excel)
const datosGenerales = { proyecto: "", proceso: "", elaboradoPor: "", fecha: fechaLocal() };
const camposGenerales = {
  proyecto: document.getElementById("dg-proyecto"),
  proceso: document.getElementById("dg-proceso"),
  elaboradoPor: document.getElementById("dg-elaborado"),
  fecha: document.getElementById("dg-fecha"),
};
Object.entries(camposGenerales).forEach(([clave, campo]) => {
  campo.value = datosGenerales[clave];
  campo.addEventListener("input", () => {
    datosGenerales[clave] = campo.value;
    hayCambios = true;
    programarGuardado();
  });
});

function mostrarDatosGenerales() {
  Object.entries(camposGenerales).forEach(([clave, campo]) => (campo.value = datosGenerales[clave]));
}

const entradaXlsx = document.getElementById("archivo-xlsx");

// --- Autoguardado (ver autoguardado.js) ---
const indicadorGuardado = document.getElementById("guardado");
const avisoRecuperar = document.getElementById("recuperar");
const textoRecuperar = document.getElementById("recuperar-texto");
let temporizadorGuardado = null;

function horaActual() {
  return new Date().toLocaleTimeString("es-CO", { hour: "2-digit", minute: "2-digit" });
}

// Guarda ya la copia (si hay equipos cargados)
function guardarEstado() {
  clearTimeout(temporizadorGuardado);
  if (equipos.length === 0) return;
  const ok = guardarCopia(equipos, datosGenerales, hayCambios);
  indicadorGuardado.textContent = ok
    ? `Guardado automático a las ${horaActual()}`
    : "No se pudo guardar automáticamente: exporta a Excel para no perder el trabajo.";
}

// Espera un momento tras el último cambio para no guardar en cada tecla
function programarGuardado() {
  clearTimeout(temporizadorGuardado);
  temporizadorGuardado = setTimeout(guardarEstado, 800);
}

// Si la pestaña se cierra o se oculta, guarda de inmediato
window.addEventListener("pagehide", guardarEstado);
document.addEventListener("visibilitychange", () => {
  if (document.visibilityState === "hidden") guardarEstado();
});

function confirmarReemplazo() {
  if (!avisoRecuperar.hidden) {
    if (!confirm("Hay una evaluación guardada sin recuperar que se reemplazará. ¿Continuar?")) return false;
    avisoRecuperar.hidden = true;
    return true;
  }
  return !hayCambios || confirm("Hay trabajo sin exportar que se perderá. ¿Quieres reemplazarlo?");
}

// Cualquier cambio dentro de la zona de equipos cuenta como trabajo sin exportar
["input", "change"].forEach((tipo) =>
  contenedor.addEventListener(tipo, () => {
    hayCambios = true;
    programarGuardado();
  })
);
contenedor.addEventListener("click", (evento) => {
  if (evento.target.closest("button")) {
    hayCambios = true;
    programarGuardado();
  }
});

botonExportar.addEventListener("click", async () => {
  try {
    await exportarExcel();
    hayCambios = false;
    guardarEstado(); // la copia queda marcada como ya exportada
    mostrarMensajes([]);
  } catch (error) {
    mostrarMensajes([error.message]);
  }
});

botonWord.addEventListener("click", async () => {
  try {
    await exportarInforme();
    mostrarMensajes([]);
  } catch (error) {
    mostrarMensajes([error.message]);
  }
});

function crear(etiqueta, clase, texto) {
  const e = document.createElement(etiqueta);
  if (clase) e.className = clase;
  if (texto !== undefined) e.textContent = texto;
  return e;
}

function campoTexto(placeholder, valor, alCambiar) {
  const input = crear("input");
  input.type = "text";
  input.placeholder = placeholder;
  input.value = valor;
  input.addEventListener("input", () => alCambiar(input.value));
  return input;
}

function casilla(marcada, alCambiar) {
  const input = crear("input");
  input.type = "checkbox";
  input.checked = marcada;
  input.addEventListener("change", () => alCambiar(input.checked));
  return input;
}

function nuevaOferta(equipo) {
  return {
    proveedor: "",
    modelo: "", // marca - modelo
    invima: false,
    alerta: false,
    valores: equipo.requisitos.map(() => null),
  };
}

function mostrarMensajes(lista) {
  mensajes.replaceChildren();
  if (lista.length === 0) return;
  const ul = crear("ul");
  lista.forEach((texto) => ul.appendChild(crear("li", "", texto)));
  mensajes.append(crear("strong", "", "Revisa el archivo:"), ul);
}

// Muestra el comparativo de ofertas de un equipo debajo de su tabla
function dibujarComparativo(zona, equipo) {
  zona.replaceChildren();
  zona.appendChild(crear("h3", "", "Comparativo de ofertas"));
  const { filas, frases } = compararOfertas(equipo);

  if (filas.length > 0) {
    const tabla = crear("table", "tabla-comparativo");
    const cabecera = tabla.createTHead().insertRow();
    ["Posición", "Oferta", "Concepto técnico", "Superan lo exigido", "Peso superado", "Requisitos que superan"]
      .forEach((t) => cabecera.appendChild(crear("th", "", t)));

    const cuerpo = tabla.createTBody();
    filas.forEach((f) => {
      const fila = cuerpo.insertRow();
      fila.insertCell().textContent = f.posicion === null ? "—" : f.posicion + (f.empate ? " (empate)" : "");
      fila.insertCell().textContent = f.nombre;
      const celdaConcepto = fila.insertCell();
      celdaConcepto.textContent = f.r.concepto;
      celdaConcepto.className = "concepto " + f.r.clase;
      fila.insertCell().textContent = f.r.superan.length;
      fila.insertCell().textContent = formatear(f.r.superanPeso);
      fila.insertCell().textContent = f.r.superan.length > 0 ? f.r.superan.join("; ") : "—";
    });
    zona.appendChild(tabla);
  }

  const lista = crear("ul", "conclusion");
  frases.forEach((texto) => lista.appendChild(crear("li", "", texto)));
  zona.appendChild(lista);
}

function dibujarEquipo(equipo) {
  const seccion = crear("section", "equipo");

  // --- Título del equipo y botón de agregar oferta (arriba a la derecha) ---
  const encabezado = crear("div", "encabezado-equipo");
  encabezado.appendChild(
    crear("h2", "", equipo.item ? `${equipo.item}. ${equipo.equipo}` : equipo.equipo)
  );
  const agregar = crear("button", "agregar", "+ Agregar oferta");
  agregar.type = "button";
  agregar.addEventListener("click", () => {
    equipo.ofertas.push(nuevaOferta(equipo));
    seccion.replaceWith(dibujarEquipo(equipo));
  });
  encabezado.appendChild(agregar);
  seccion.appendChild(encabezado);

  const envoltura = crear("div", "tabla-scroll");
  const tabla = crear("table");
  const refs = []; // por oferta: celdas de resultado y selects
  const columnasFijas = 3;

  // --- Encabezado de la tabla: datos de cada oferta ---
  const cabecera = tabla.createTHead().insertRow();
  ["#", "Requisito", "Peso"].forEach((t, i) => {
    cabecera.appendChild(crear("th", i === 2 ? "num" : "", t));
  });
  equipo.ofertas.forEach((oferta, j) => {
    const th = crear("th", "oferta");
    th.appendChild(campoTexto("Proveedor", oferta.proveedor, (v) => (oferta.proveedor = v)));
    th.appendChild(campoTexto("Marca - modelo", oferta.modelo, (v) => (oferta.modelo = v)));
    const quitar = crear("button", "mini", "Quitar");
    quitar.type = "button";
    quitar.addEventListener("click", () => {
      equipo.ofertas.splice(j, 1);
      seccion.replaceWith(dibujarEquipo(equipo));
    });
    th.appendChild(quitar);
    cabecera.appendChild(th);
  });

  const cuerpo = tabla.createTBody();

  // --- Casillas: registro INVIMA y alerta sanitaria ---
  const filaInvima = cuerpo.insertRow();
  const etiquetaInvima = filaInvima.insertCell();
  etiquetaInvima.colSpan = columnasFijas;
  etiquetaInvima.textContent = "Cuenta con registro INVIMA";
  etiquetaInvima.className = "etiqueta-fila";
  const filaAlerta = cuerpo.insertRow();
  const etiquetaAlerta = filaAlerta.insertCell();
  etiquetaAlerta.colSpan = columnasFijas;
  etiquetaAlerta.textContent = "Alerta sanitaria";
  etiquetaAlerta.className = "etiqueta-fila";

  equipo.ofertas.forEach((oferta) => {
    const celdaInvima = filaInvima.insertCell();
    celdaInvima.className = "num";
    celdaInvima.appendChild(casilla(oferta.invima, (v) => {
      oferta.invima = v;
      actualizar();
    }));

    const celdaAlerta = filaAlerta.insertCell();
    celdaAlerta.className = "num";
    celdaAlerta.appendChild(casilla(oferta.alerta, (v) => {
      oferta.alerta = v;
      actualizar();
    }));
  });

  // --- Requisitos con un select por oferta ---
  equipo.ofertas.forEach(() => refs.push({ selects: [] }));

  equipo.requisitos.forEach((req, i) => {
    const fila = cuerpo.insertRow();
    fila.insertCell().textContent = i + 1;
    fila.insertCell().textContent = req.requisito;
    const celdaPeso = fila.insertCell();
    celdaPeso.className = "num";
    celdaPeso.textContent = req.peso;

    equipo.ofertas.forEach((oferta, j) => {
      const select = crear("select", "nota");
      OPCIONES.forEach((op) => {
        const option = crear("option", "", op.texto);
        option.value = op.valor;
        select.appendChild(option);
      });
      const pintar = () => {
        select.value = oferta.valores[i] === null ? "" : String(oferta.valores[i]);
        select.className = "nota " + (select.value === "" ? "" : "v" + select.value);
      };
      pintar();
      select.addEventListener("change", () => {
        oferta.valores[i] = select.value === "" ? null : Number(select.value);
        pintar();
        actualizar();
      });
      refs[j].selects.push(pintar);
      fila.insertCell().appendChild(select);
    });
  });

  // --- Pie: resultados ---
  const pie = tabla.createTFoot();
  const nombresPie = ["Requisitos que cumple", "Cumplimiento", "Concepto técnico", "Observación", ""];
  const filasPie = nombresPie.map((nombre) => {
    const fila = pie.insertRow();
    const celda = fila.insertCell();
    celda.colSpan = columnasFijas;
    celda.textContent = nombre;
    return fila;
  });

  equipo.ofertas.forEach((oferta, j) => {
    refs[j].puntaje = filasPie[0].insertCell();
    refs[j].porcentaje = filasPie[1].insertCell();
    refs[j].concepto = filasPie[2].insertCell();
    refs[j].detalle = filasPie[3].insertCell();
    refs[j].detalle.className = "detalle";

    const vacios = crear("button", "mini", "Vacíos = 0");
    vacios.type = "button";
    vacios.addEventListener("click", () => {
      oferta.valores = oferta.valores.map((v) => (v === null ? 0 : v));
      refs[j].selects.forEach((pintar) => pintar());
      actualizar();
    });
    filasPie[4].insertCell().appendChild(vacios);
  });

  const comparativo = crear("div", "comparativo");

  function actualizar() {
    equipo.ofertas.forEach((oferta, j) => {
      const r = calcularOferta(equipo, oferta);
      refs[j].puntaje.textContent = `${r.cumplidos} de ${r.total}`;
      refs[j].porcentaje.textContent = `${Math.round(r.porcentaje * 100)} %`;
      refs[j].concepto.textContent = r.concepto;
      refs[j].concepto.className = "concepto " + r.clase;
      refs[j].detalle.textContent = r.detalle;
    });
    dibujarComparativo(comparativo, equipo);
  }
  actualizar();

  envoltura.appendChild(tabla);
  seccion.appendChild(envoltura);
  seccion.appendChild(comparativo);
  return seccion;
}

function dibujarTodo() {
  contenedor.replaceChildren(...equipos.map(dibujarEquipo));
}

function actualizarResumen() {
  if (equipos.length === 0) {
    resumen.textContent = "";
    return;
  }
  const totalReq = equipos.reduce((s, e) => s + e.requisitos.length, 0);
  resumen.textContent = `${equipos.length} equipo(s), ${totalReq} requisito(s).`;
}

entradaXlsx.addEventListener("change", async () => {
  const archivo = entradaXlsx.files[0];
  if (!archivo) return;
  try {
    const resultado = await leerEvaluacion(archivo);
    if (resultado.errores.length > 0) {
      mostrarMensajes(resultado.errores);
      return;
    }
    if (!confirmarReemplazo()) return; // solo pregunta si el archivo es válido
    equipos = resultado.equipos;
    Object.assign(datosGenerales, resultado.general);
    mostrarDatosGenerales();
    hayCambios = false;
    botonExportar.disabled = false;
    botonWord.disabled = false;
    mostrarMensajes([]);
    actualizarResumen();
    resumen.textContent += ` Importado de ${archivo.name}.`;
    dibujarTodo();
    guardarEstado();
  } catch (error) {
    mostrarMensajes([error.message]);
  } finally {
    entradaXlsx.value = ""; // permite volver a elegir el mismo archivo
  }
});

entrada.addEventListener("change", async () => {
  const archivo = entrada.files[0];
  if (!archivo) return;
  if (!confirmarReemplazo()) {
    entrada.value = "";
    return;
  }

  const texto = await archivo.text();
  const resultado = leerCriterios(texto);
  const avisos = [...resultado.errores];
  if (texto.includes("�")) {
    avisos.push('Hay caracteres dañados. Guarda el archivo como "CSV UTF-8" y vuelve a cargarlo.');
  }

  equipos = resultado.equipos;
  equipos.forEach((equipo) => {
    equipo.ofertas = [nuevaOferta(equipo), nuevaOferta(equipo)]; // dos ofertas para empezar
  });

  hayCambios = false;
  botonExportar.disabled = equipos.length === 0;
  botonWord.disabled = equipos.length === 0;
  mostrarMensajes(avisos);
  actualizarResumen();
  dibujarTodo();
  guardarEstado();
});

// --- Al abrir la página: ofrecer recuperar la evaluación guardada ---
(function ofrecerRecuperacion() {
  const copia = leerCopia();
  if (!copia) return;
  const ofertas = copia.equipos.reduce((s, e) => s + e.ofertas.length, 0);
  const fecha = new Date(copia.guardado).toLocaleString("es-CO", { dateStyle: "short", timeStyle: "short" });
  textoRecuperar.textContent =
    `Hay una evaluación guardada automáticamente el ${fecha} ` +
    `(${copia.equipos.length} equipo(s), ${ofertas} oferta(s)). ¿Quieres recuperarla?`;
  avisoRecuperar.hidden = false;

  document.getElementById("recuperar-si").addEventListener("click", () => {
    equipos = copia.equipos;
    Object.assign(datosGenerales, copia.general);
    mostrarDatosGenerales();
    hayCambios = copia.sinExportar;
    botonExportar.disabled = false;
    botonWord.disabled = false;
    avisoRecuperar.hidden = true;
    mostrarMensajes([]);
    actualizarResumen();
    resumen.textContent += " Recuperado del autoguardado.";
    dibujarTodo();
    guardarEstado();
  });

  document.getElementById("recuperar-no").addEventListener("click", () => {
    if (!confirm("Se borrará la copia guardada y no se podrá recuperar. ¿Descartar?")) return;
    borrarCopia();
    avisoRecuperar.hidden = true;
  });
})();
