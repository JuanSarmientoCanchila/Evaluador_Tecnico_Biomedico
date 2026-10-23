// ---------- Clasificación de equipos y cifras para el informe ----------

function porcentajeTexto(parte, total) {
  if (total === 0) return "0 %";
  return ((parte / total) * 100).toLocaleString("es-CO", { maximumFractionDigits: 1 }) + " %";
}

// Un equipo queda en un solo grupo, según el mejor resultado de sus ofertas:
// aprobado (alguna Recomendada) > pendiente de INVIMA (alguna con observaciones)
// > incompleto (sin ofertas u ofertas sin evaluar) > no cumple (todas No recomendadas)
function clasificarEquipos(equipos) {
  const grupos = { aprobados: [], pendientesInvima: [], noCumplen: [], incompletos: [] };
  const ofertas = { total: 0, recomendadas: 0, conObservaciones: 0, noRecomendadas: 0, incompletas: 0 };
  let requisitos = 0;

  equipos.forEach((equipo) => {
    const comparativo = compararOfertas(equipo);
    const de = (clase) => comparativo.filas.filter((f) => f.r.clase === clase);
    const registro = {
      equipo, comparativo,
      rec: de("si"), obs: de("obs"), no: de("no"), pend: de("pendiente"),
    };

    requisitos += equipo.requisitos.length;
    ofertas.total += comparativo.filas.length;
    ofertas.recomendadas += registro.rec.length;
    ofertas.conObservaciones += registro.obs.length;
    ofertas.noRecomendadas += registro.no.length;
    ofertas.incompletas += registro.pend.length;

    if (registro.rec.length > 0) grupos.aprobados.push(registro);
    else if (registro.obs.length > 0) grupos.pendientesInvima.push(registro);
    else if (comparativo.filas.length === 0 || registro.pend.length > 0) grupos.incompletos.push(registro);
    else grupos.noCumplen.push(registro);
  });

  return { total: equipos.length, requisitos, grupos, ofertas };
}
