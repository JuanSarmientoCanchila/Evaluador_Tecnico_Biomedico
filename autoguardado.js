// Autoguardado en el navegador (localStorage).
// Guarda una copia de la evaluación en curso para recuperarla si se cierra o recarga la pestaña.
// No reemplaza al Excel: la copia vive solo en este navegador y en esta dirección (localhost:5500).

const CLAVE_GUARDADO = "evaluador2:autoguardado";
const VERSION_GUARDADO = 1;

// Comprueba que lo guardado tenga la forma esperada antes de usarlo
function estadoValido(e) {
  if (!e || e.version !== VERSION_GUARDADO) return false;
  if (!Array.isArray(e.equipos) || e.equipos.length === 0) return false;
  if (typeof e.general !== "object" || e.general === null) return false;
  return e.equipos.every(
    (q) =>
      q &&
      typeof q.item === "string" &&
      typeof q.equipo === "string" &&
      Array.isArray(q.requisitos) &&
      q.requisitos.every((r) => typeof r.requisito === "string" && Number.isFinite(r.peso) && r.peso > 0) &&
      Array.isArray(q.ofertas) &&
      q.ofertas.every(
        (o) =>
          typeof o.proveedor === "string" &&
          typeof o.modelo === "string" &&
          typeof o.invima === "boolean" &&
          typeof o.alerta === "boolean" &&
          Array.isArray(o.valores) &&
          o.valores.length === q.requisitos.length &&
          o.valores.every((v) => v === null || v === 0 || v === 1 || v === 2)
      )
  );
}

// Devuelve true si se pudo guardar
function guardarCopia(equipos, general, sinExportar) {
  try {
    const estado = {
      version: VERSION_GUARDADO,
      guardado: Date.now(),
      sinExportar,
      general,
      equipos,
    };
    localStorage.setItem(CLAVE_GUARDADO, JSON.stringify(estado));
    return true;
  } catch (error) {
    return false; // almacenamiento bloqueado, lleno o modo privado
  }
}

// Devuelve el estado guardado, o null si no hay nada o está dañado
function leerCopia() {
  try {
    const texto = localStorage.getItem(CLAVE_GUARDADO);
    if (!texto) return null;
    const estado = JSON.parse(texto);
    return estadoValido(estado) ? estado : null;
  } catch (error) {
    return null;
  }
}

function borrarCopia() {
  try {
    localStorage.removeItem(CLAVE_GUARDADO);
  } catch (error) {
    // sin almacenamiento: nada que borrar
  }
}
