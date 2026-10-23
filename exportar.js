// ---------- Exportar a Excel (usa la librería ExcelJS) ----------

const COLOR_NOTA = { 0: "FFFF0000", 1: "FFFFFF00", 2: "FF92D050" }; // igual a tu Excel
const COLOR_CONCEPTO = {
  si: "FFC6E0B4",
  obs: "FFFFE699",
  no: "FFF8CBAD",
  pendiente: "FFD9D9D9",
};
const BORDE = {
  top: { style: "thin", color: { argb: "FFBFBFBF" } },
  left: { style: "thin", color: { argb: "FFBFBFBF" } },
  bottom: { style: "thin", color: { argb: "FFBFBFBF" } },
  right: { style: "thin", color: { argb: "FFBFBFBF" } },
};

function relleno(argb) {
  return { type: "pattern", pattern: "solid", fgColor: { argb } };
}

function estilizar(celda, opciones = {}) {
  celda.border = BORDE;
  celda.alignment = {
    vertical: "middle",
    horizontal: opciones.izquierda ? "left" : "center",
    wrapText: true,
  };
  if (opciones.negrita) celda.font = { bold: true };
  if (opciones.fondo) celda.fill = relleno(opciones.fondo);
}

function nombreOferta(oferta, j) {
  const partes = [oferta.proveedor.trim(), oferta.modelo.trim()].filter(Boolean);
  return partes.length > 0 ? partes.join("\n") : `Oferta ${j + 1}`;
}

// ---- Hoja 1: la evaluación, lista para leer e imprimir ----
function hojaEvaluacion(libro) {
  const hoja = libro.addWorksheet("Evaluación", {
    pageSetup: { orientation: "landscape", fitToPage: true, fitToWidth: 1, fitToHeight: 0 },
  });
  const maxOfertas = Math.max(1, ...equipos.map((e) => e.ofertas.length));
  hoja.getColumn(1).width = 5;
  hoja.getColumn(2).width = 55;
  hoja.getColumn(3).width = 8;
  for (let j = 0; j < maxOfertas; j++) hoja.getColumn(4 + j).width = 30;

  let fila = 1;
  equipos.forEach((equipo) => {
    const ultimaColumna = 3 + Math.max(1, equipo.ofertas.length);

    // Título del equipo
    hoja.mergeCells(fila, 1, fila, ultimaColumna);
    const titulo = hoja.getCell(fila, 1);
    titulo.value = equipo.item ? `${equipo.item}. ${equipo.equipo}` : equipo.equipo;
    titulo.font = { bold: true, size: 13 };
    titulo.fill = relleno("FFDDEBF7");
    titulo.alignment = { vertical: "middle" };
    fila++;

    // Encabezados
    ["#", "Requisito", "Peso"].forEach((texto, c) => {
      const celda = hoja.getCell(fila, c + 1);
      celda.value = texto;
      estilizar(celda, { negrita: true, fondo: "FFF2F2F2" });
    });
    equipo.ofertas.forEach((oferta, j) => {
      const celda = hoja.getCell(fila, 4 + j);
      celda.value = nombreOferta(oferta, j);
      estilizar(celda, { negrita: true, fondo: "FFF2F2F2" });
    });
    hoja.getRow(fila).height = 34;
    fila++;

    // INVIMA y alerta sanitaria
    [["Cuenta con registro INVIMA", "invima"], ["Alerta sanitaria", "alerta"]].forEach(([texto, campo]) => {
      hoja.mergeCells(fila, 1, fila, 3);
      const etiqueta = hoja.getCell(fila, 1);
      etiqueta.value = texto;
      estilizar(etiqueta, { negrita: true, izquierda: true });
      equipo.ofertas.forEach((oferta, j) => {
        const celda = hoja.getCell(fila, 4 + j);
        celda.value = oferta[campo] ? "Sí" : "No";
        estilizar(celda);
      });
      fila++;
    });

    // Requisitos con su evaluación (0, 1 o 2) coloreada
    equipo.requisitos.forEach((req, i) => {
      hoja.getCell(fila, 1).value = i + 1;
      hoja.getCell(fila, 2).value = req.requisito;
      hoja.getCell(fila, 3).value = req.peso;
      estilizar(hoja.getCell(fila, 1));
      estilizar(hoja.getCell(fila, 2), { izquierda: true });
      estilizar(hoja.getCell(fila, 3));
      equipo.ofertas.forEach((oferta, j) => {
        const celda = hoja.getCell(fila, 4 + j);
        const valor = oferta.valores[i];
        if (valor === 0 || valor === 1 || valor === 2) {
          celda.value = valor;
          estilizar(celda, { fondo: COLOR_NOTA[valor] });
        } else {
          estilizar(celda); // sin evaluar: celda vacía
        }
      });
      fila++;
    });

    // Resultados
    const resultados = equipo.ofertas.map((oferta) => calcularOferta(equipo, oferta));
    const filasResultado = [
      ["Requisitos que cumple", (r) => `${r.cumplidos} de ${r.total}`],
      ["Cumplimiento", (r) => r.porcentaje],
      ["Concepto técnico", (r) => r.concepto],
      ["Observación", (r) => r.detalle],
    ];
    filasResultado.forEach(([texto, valorDe], k) => {
      hoja.mergeCells(fila, 1, fila, 3);
      const etiqueta = hoja.getCell(fila, 1);
      etiqueta.value = texto;
      estilizar(etiqueta, { negrita: true, izquierda: true, fondo: "FFF2F2F2" });
      resultados.forEach((r, j) => {
        const celda = hoja.getCell(fila, 4 + j);
        celda.value = valorDe(r);
        estilizar(celda, { negrita: k === 2, fondo: k === 2 ? COLOR_CONCEPTO[r.clase] : undefined });
        if (k === 1) celda.numFmt = "0%";
      });
      if (k === 3) hoja.getRow(fila).height = 48;
      fila++;
    });

    // Comparativo de ofertas (conclusión)
    if (equipo.ofertas.length > 0) {
      const { frases } = compararOfertas(equipo);
      hoja.mergeCells(fila, 1, fila, 3);
      const etiqueta = hoja.getCell(fila, 1);
      etiqueta.value = "Comparativo de ofertas";
      estilizar(etiqueta, { negrita: true, izquierda: true, fondo: "FFF2F2F2" });
      if (ultimaColumna > 4) hoja.mergeCells(fila, 4, fila, ultimaColumna);
      const celda = hoja.getCell(fila, 4);
      celda.value = frases.join("\n");
      estilizar(celda, { izquierda: true });
      const caracteresPorLinea = 30 * equipo.ofertas.length;
      const lineas = frases.reduce((s, f) => s + Math.ceil(f.length / caracteresPorLinea), 0);
      hoja.getRow(fila).height = Math.max(32, 16 * lineas + 6);
      fila++;
    }

    fila++; // fila en blanco entre equipos
  });
  return hoja;
}

// ---- Hoja 2: los mismos datos en formato plano (para importar después) ----
function hojaDatos(libro) {
  const hoja = libro.addWorksheet("Datos");
  const columnas = ["item", "equipo", "requisito", "peso", "oferta", "proveedor", "modelo", "invima", "alerta", "valor"];
  hoja.addRow(columnas).font = { bold: true };
  const anchos = [8, 40, 55, 8, 8, 28, 28, 9, 9, 8];
  anchos.forEach((ancho, c) => (hoja.getColumn(c + 1).width = ancho));

  equipos.forEach((equipo) => {
    equipo.requisitos.forEach((req, i) => {
      if (equipo.ofertas.length === 0) {
        hoja.addRow([equipo.item, equipo.equipo, req.requisito, req.peso, "", "", "", "", "", ""]);
        return;
      }
      equipo.ofertas.forEach((oferta, j) => {
        const valor = oferta.valores[i];
        hoja.addRow([
          equipo.item, equipo.equipo, req.requisito, req.peso,
          j + 1, oferta.proveedor, oferta.modelo,
          oferta.invima ? "si" : "no", oferta.alerta ? "si" : "no",
          valor === null ? "" : valor,
        ]);
      });
    });
  });
  hoja.views = [{ state: "frozen", ySplit: 1 }];
  return hoja;
}

// ---- Hoja 3: datos generales del informe ----
function hojaGeneral(libro) {
  const hoja = libro.addWorksheet("General");
  hoja.addRow(["campo", "valor"]).font = { bold: true };
  Object.entries(datosGenerales).forEach(([campo, valor]) => hoja.addRow([campo, valor]));
  hoja.getColumn(1).width = 16;
  hoja.getColumn(2).width = 60;
}

async function exportarExcel() {
  if (typeof ExcelJS === "undefined") {
    throw new Error('No se encontró la librería de Excel. Revisa que el archivo "exceljs.min.js" esté en la carpeta.');
  }
  const libro = new ExcelJS.Workbook();
  libro.creator = "Evaluador técnico";
  hojaEvaluacion(libro);
  hojaDatos(libro);
  hojaGeneral(libro);

  const datos = await libro.xlsx.writeBuffer();
  const blob = new Blob([datos], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  descargarArchivo(blob, `evaluacion-${fechaLocal()}.xlsx`);
}

// Descarga un archivo generado en el navegador
function descargarArchivo(blob, nombre) {
  const enlace = document.createElement("a");
  enlace.href = URL.createObjectURL(blob);
  enlace.download = nombre;
  document.body.appendChild(enlace);
  enlace.click();
  enlace.remove();
  URL.revokeObjectURL(enlace.href);
}
