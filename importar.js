// ---------- Importar una evaluación exportada (hoja "Datos") ----------

const COLUMNAS_DATOS = [
  "item", "equipo", "requisito", "peso", "oferta",
  "proveedor", "modelo", "invima", "alerta", "valor",
];
const MAX_ERRORES = 10;
const CAMPOS_GENERALES = ["proyecto", "proceso", "elaboradoPor", "fecha"];

// Lee el contenido de una celda como texto, venga como venga de Excel
function textoCelda(celda) {
  let v = celda.value;
  if (v === null || v === undefined) return "";
  if (typeof v === "object") {
    if (v.richText) return v.richText.map((t) => t.text).join("").trim();
    if ("result" in v) v = v.result; // celda con fórmula
    else if ("text" in v) v = v.text; // hipervínculo
  }
  return v === null || v === undefined ? "" : String(v).trim();
}

// Devuelve { equipos, errores }. Si hay algún error, no se importa nada.
async function leerEvaluacion(archivo) {
  if (typeof ExcelJS === "undefined") {
    throw new Error('No se encontró la librería de Excel. Revisa que el archivo "exceljs.min.js" esté en la carpeta.');
  }

  const libro = new ExcelJS.Workbook();
  try {
    await libro.xlsx.load(await archivo.arrayBuffer());
  } catch (error) {
    return { equipos: [], errores: ["No se pudo abrir el archivo. ¿Es un Excel (.xlsx) exportado desde esta página?"] };
  }

  const hoja = libro.getWorksheet("Datos");
  if (!hoja) {
    return { equipos: [], errores: ['El archivo no tiene la hoja "Datos". Solo se pueden importar archivos exportados desde esta página.'] };
  }

  // Posición de cada columna según el encabezado
  const errores = [];
  const encabezados = [];
  hoja.getRow(1).eachCell({ includeEmpty: true }, (celda, n) => {
    encabezados[n - 1] = normalizar(textoCelda(celda));
  });
  const pos = {};
  COLUMNAS_DATOS.forEach((col) => {
    pos[col] = encabezados.indexOf(col) + 1;
    if (!pos[col]) errores.push(`Falta la columna "${col}" en la hoja Datos.`);
  });
  if (errores.length > 0) return { equipos: [], errores };

  const grupos = new Map();

  hoja.eachRow((fila, numeroFila) => {
    if (numeroFila === 1) return;
    const dato = (col) => textoCelda(fila.getCell(pos[col]));
    const fallo = (mensaje) => errores.push(`Fila ${numeroFila}: ${mensaje}`);

    const item = dato("item");
    const nombreEquipo = dato("equipo");
    const requisito = dato("requisito");
    if (nombreEquipo === "" && requisito === "") return; // fila vacía
    if (nombreEquipo === "" || requisito === "") return fallo("falta el equipo o el requisito.");

    const peso = Number(dato("peso").replace(",", "."));
    if (!(peso > 0)) return fallo(`el peso "${dato("peso")}" no es válido.`);

    let numOferta = 0; // 0 = el equipo no tiene ofertas
    if (dato("oferta") !== "") {
      numOferta = Number(dato("oferta"));
      if (!Number.isInteger(numOferta) || numOferta < 1) return fallo(`el número de oferta "${dato("oferta")}" no es válido.`);
    }

    const invima = normalizar(dato("invima"));
    const alerta = normalizar(dato("alerta"));
    if (!["", "si", "no"].includes(invima)) return fallo('"invima" debe ser si o no.');
    if (!["", "si", "no"].includes(alerta)) return fallo('"alerta" debe ser si o no.');

    let valor = null;
    if (dato("valor") !== "") {
      valor = Number(dato("valor"));
      if (![0, 1, 2].includes(valor)) return fallo(`la evaluación "${dato("valor")}" debe ser 0, 1 o 2.`);
    }

    const clave = `${item}|${nombreEquipo}`;
    if (!grupos.has(clave)) grupos.set(clave, { item, equipo: nombreEquipo, requisitos: [], ofertas: [] });
    const equipo = grupos.get(clave);

    // La fila de la oferta 1 (o sin oferta) inicia un requisito nuevo
    if (numOferta <= 1) equipo.requisitos.push({ requisito, peso });
    const i = equipo.requisitos.length - 1;
    if (i < 0) return fallo("aparece una oferta antes de la oferta 1.");
    if (equipo.requisitos[i].requisito !== requisito) {
      return fallo("el requisito no coincide con el de la oferta 1. ¿Se movieron filas de la hoja Datos?");
    }
    if (numOferta === 0) return;

    if (numOferta > equipo.ofertas.length + 1) return fallo(`la oferta ${numOferta} salta números.`);
    if (numOferta === equipo.ofertas.length + 1) {
      equipo.ofertas.push({
        proveedor: dato("proveedor"),
        modelo: dato("modelo"),
        invima: invima === "si",
        alerta: alerta === "si",
        valores: [],
      });
    }
    equipo.ofertas[numOferta - 1].valores[i] = valor;
  });

  if (grupos.size === 0 && errores.length === 0) errores.push('La hoja "Datos" no tiene filas.');

  if (errores.length > MAX_ERRORES) {
    const resto = errores.length - MAX_ERRORES;
    errores.length = MAX_ERRORES;
    errores.push(`…y ${resto} error(es) más.`);
  }
  if (errores.length > 0) return { equipos: [], errores };

  // Datos generales del informe (la hoja "General" es opcional)
  const general = {};
  const hojaGeneral = libro.getWorksheet("General");
  if (hojaGeneral) {
    hojaGeneral.eachRow((fila, n) => {
      if (n === 1) return;
      const campo = textoCelda(fila.getCell(1));
      if (CAMPOS_GENERALES.includes(campo)) general[campo] = textoCelda(fila.getCell(2));
    });
  }

  // Completa con "sin evaluar" las evaluaciones que falten
  const equipos = [...grupos.values()];
  equipos.forEach((equipo) => {
    equipo.ofertas.forEach((oferta) => {
      oferta.valores = equipo.requisitos.map((_, k) => oferta.valores[k] ?? null);
    });
  });
  return { equipos, errores, general };
}
