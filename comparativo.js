// ---------- Comparativo de ofertas de un mismo equipo ----------
// Primero manda el concepto (Recomendado > con observaciones > No recomendado).
// Entre las recomendadas, se ordena por cuánto superan lo exigido, ponderado por
// el peso de cada requisito. Es un criterio de desempate orientativo.

const TOLERANCIA = 1e-9;

function etiquetaOferta(oferta, j) {
  const partes = [oferta.proveedor.trim(), oferta.modelo.trim()].filter(Boolean);
  return partes.length > 0 ? partes.join(" · ") : `Oferta ${j + 1}`;
}

function iguales(a, b) {
  return Math.abs(a - b) < TOLERANCIA;
}

function unir(nombres) {
  if (nombres.length <= 1) return nombres.join("");
  return nombres.slice(0, -1).join(", ") + " y " + nombres[nombres.length - 1];
}

function porSuperan(lista) {
  return [...lista].sort((a, b) => b.r.superanPeso - a.r.superanPeso);
}

// Devuelve { filas, frases }: las filas ya ordenadas y las frases de conclusión
function compararOfertas(equipo) {
  const todas = equipo.ofertas.map((oferta, j) => ({
    nombre: etiquetaOferta(oferta, j),
    oferta,
    indice: j,
    r: calcularOferta(equipo, oferta),
    posicion: null,
  }));

  const recomendadas = porSuperan(todas.filter((f) => f.r.clase === "si"));
  const conObservaciones = porSuperan(todas.filter((f) => f.r.clase === "obs"));
  const noRecomendadas = todas.filter((f) => f.r.clase === "no");
  const pendientes = todas.filter((f) => f.r.clase === "pendiente");

  // Posiciones con empates: 1, 1, 3
  recomendadas.forEach((f, k) => {
    const anterior = recomendadas[k - 1];
    f.posicion = k > 0 && iguales(f.r.superanPeso, anterior.r.superanPeso) ? anterior.posicion : k + 1;
  });
  recomendadas.forEach((f) => {
    f.empate = recomendadas.filter((g) => iguales(g.r.superanPeso, f.r.superanPeso)).length > 1;
  });

  const frases = [];
  const resumenSuperan = (f) =>
    `${f.nombre} (supera lo exigido en ${f.r.superan.length} requisito(s), peso ${formatear(f.r.superanPeso)})`;

  if (todas.length === 0) {
    frases.push("Agrega al menos una oferta para ver el comparativo.");
  } else if (pendientes.length === todas.length) {
    frases.push("Termina de evaluar las ofertas para ver el comparativo.");
  } else if (recomendadas.length === 0 && conObservaciones.length === 0) {
    frases.push("Ninguna de las ofertas evaluadas es recomendada.");
  }

  if (recomendadas.length === 1) {
    frases.push(`Única oferta recomendada: ${recomendadas[0].nombre}.`);
  } else if (recomendadas.length > 1) {
    const primeras = recomendadas.filter((f) => f.posicion === 1);
    if (primeras.length === recomendadas.length) {
      frases.push(`Empate técnico entre ${unir(primeras.map((f) => f.nombre))}: según el puntaje técnico, cualquiera puede ser elegida.`);
    } else if (primeras.length > 1) {
      frases.push(`Van primeras, empatadas: ${unir(primeras.map((f) => f.nombre))}. Según el puntaje técnico, cualquiera de ellas puede ser elegida.`);
    } else {
      frases.push(`La que más supera lo exigido es ${resumenSuperan(primeras[0])}.`);
    }
    frases.push("El orden es un criterio de desempate orientativo: superar lo pedido no implica mayor calidad.");
  }

  if (recomendadas.length === 0 && conObservaciones.length > 0) {
    frases.push(`Ninguna oferta está plenamente recomendada. Con observaciones: ${unir(conObservaciones.map((f) => f.nombre))}.`);
  }

  // Qué pasaría si las ofertas con observaciones presentan el INVIMA
  conObservaciones.forEach((f) => {
    if (recomendadas.length === 0) {
      frases.push(`Si ${f.nombre} presenta el INVIMA, pasaría a Recomendada.`);
      return;
    }
    const mejores = recomendadas.filter((g) => g.r.superanPeso > f.r.superanPeso + TOLERANCIA).length;
    const empata = recomendadas.some((g) => iguales(g.r.superanPeso, f.r.superanPeso));
    const posicion = mejores + 1;
    frases.push(
      `Si ${f.nombre} presenta el INVIMA, pasaría a Recomendada y ${empata ? "empataría en" : "quedaría en"} la posición ${posicion} de ${recomendadas.length + 1}.`
    );
  });

  if (pendientes.length > 0 && pendientes.length < todas.length) {
    frases.push(`Sin evaluar por completo (no entran al comparativo): ${unir(pendientes.map((f) => f.nombre))}.`);
  }

  return { filas: [...recomendadas, ...conObservaciones, ...noRecomendadas, ...pendientes], frases };
}
