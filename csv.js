// ---------- Lectura del archivo de criterios (CSV) ----------

// La columna "obligatorio" ya no se usa; si viene en el archivo, se ignora.
const COLUMNAS = ["item", "equipo", "requisito", "peso"];

// Excel en español suele guardar con ";" y otros programas con ","
function detectarSeparador(texto) {
  const primeraLinea = texto.split(/\r?\n/)[0];
  const comas = (primeraLinea.match(/,/g) || []).length;
  const puntoYComas = (primeraLinea.match(/;/g) || []).length;
  return puntoYComas > comas ? ";" : ",";
}

// Convierte el texto en una lista de filas, respetando las comillas
function parsearCSV(texto, separador) {
  const filas = [];
  let fila = [];
  let campo = "";
  let entreComillas = false;

  for (let i = 0; i < texto.length; i++) {
    const c = texto[i];
    if (entreComillas) {
      if (c === '"') {
        if (texto[i + 1] === '"') {
          campo += '"';
          i++;
        } else {
          entreComillas = false;
        }
      } else {
        campo += c;
      }
    } else if (c === '"') {
      entreComillas = true;
    } else if (c === separador) {
      fila.push(campo);
      campo = "";
    } else if (c === "\n" || c === "\r") {
      if (c === "\r" && texto[i + 1] === "\n") i++;
      fila.push(campo);
      filas.push(fila);
      fila = [];
      campo = "";
    } else {
      campo += c;
    }
  }
  if (campo !== "" || fila.length > 0) {
    fila.push(campo);
    filas.push(fila);
  }
  return filas.filter((f) => f.some((x) => x.trim() !== ""));
}

function normalizar(texto) {
  return texto
    .trim()
    .toLowerCase()
    .normalize("NFD")
    .replace(/[̀-ͯ]/g, "");
}

// Devuelve { equipos, errores }
function leerCriterios(textoOriginal) {
  const errores = [];
  const texto = textoOriginal.replace(/^﻿/, ""); // quita la marca BOM
  const filas = parsearCSV(texto, detectarSeparador(texto));

  if (filas.length < 2) {
    return { equipos: [], errores: ["El archivo está vacío o solo tiene encabezados."] };
  }

  const encabezados = filas[0].map(normalizar);
  const posicion = {};
  for (const col of COLUMNAS) {
    posicion[col] = encabezados.indexOf(col);
    if (posicion[col] === -1) errores.push(`Falta la columna "${col}".`);
  }
  if (errores.length > 0) return { equipos: [], errores };

  const grupos = new Map();

  for (let i = 1; i < filas.length; i++) {
    const numeroLinea = i + 1;
    const f = filas[i];
    const item = (f[posicion.item] || "").trim();
    const equipo = (f[posicion.equipo] || "").trim();
    const requisito = (f[posicion.requisito] || "").trim();
    const textoPeso = (f[posicion.peso] || "").trim().replace(",", ".");

    if (equipo === "" || requisito === "") {
      errores.push(`Línea ${numeroLinea}: falta el equipo o el requisito.`);
      continue;
    }

    let peso = 1; // si el peso está vacío, vale 1
    if (textoPeso !== "") {
      peso = Number(textoPeso);
      if (!Number.isFinite(peso) || peso <= 0) {
        errores.push(`Línea ${numeroLinea}: el peso "${textoPeso}" no es un número válido.`);
        continue;
      }
    }

    const clave = `${item}|${equipo}`;
    if (!grupos.has(clave)) grupos.set(clave, { item, equipo, requisitos: [] });
    grupos.get(clave).requisitos.push({ requisito, peso });
  }

  return { equipos: [...grupos.values()], errores };
}
