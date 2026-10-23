// ---------- Cálculo del cumplimiento y del concepto técnico ----------

function formatear(numero) {
  return String(Number(numero.toFixed(2)));
}

// valores: null = sin evaluar, 0 = No cumple, 1 = Cumple, 2 = Cumple por encima
// Para el cumplimiento solo importa si cumple (1 o 2) o no cumple (0).
// El "2" es solo una anotación: no suma al porcentaje.
function calcularOferta(equipo, oferta) {
  let pesoTotal = 0;
  let pesoCumplido = 0;
  let cumplidos = 0;
  let sinEvaluar = 0;
  const incumplidos = [];
  const superan = [];
  let superanPeso = 0; // suma de los pesos de los requisitos que superan lo exigido

  equipo.requisitos.forEach((req, i) => {
    const valor = oferta.valores[i];
    const evaluado = valor === 0 || valor === 1 || valor === 2;
    pesoTotal += req.peso;
    if (!evaluado) {
      sinEvaluar++;
      return;
    }
    if (valor === 0) {
      incumplidos.push(req.requisito);
      return;
    }
    cumplidos++;
    pesoCumplido += req.peso;
    if (valor === 2) {
      superan.push(req.requisito);
      superanPeso += req.peso;
    }
  });

  const total = equipo.requisitos.length;
  const porcentaje = pesoTotal > 0 ? pesoCumplido / pesoTotal : 0;
  const resultado = { cumplidos, total, porcentaje, sinEvaluar, superan, superanPeso };
  const anotacion = superan.length > 0 ? ` Superan lo exigido: ${superan.join("; ")}.` : "";

  if (oferta.alerta) {
    return { ...resultado, concepto: "No recomendado", clase: "no", detalle: "Tiene alerta sanitaria." };
  }
  if (sinEvaluar > 0) {
    return { ...resultado, concepto: "Evaluación incompleta", clase: "pendiente", detalle: `${sinEvaluar} requisito(s) sin evaluar.` };
  }
  if (incumplidos.length > 0) {
    return {
      ...resultado, concepto: "No recomendado", clase: "no",
      detalle: `No cumple: ${incumplidos.join("; ")}.${anotacion}`,
    };
  }

  // Aquí cumple el 100 % de lo técnico
  if (!oferta.invima) {
    return {
      ...resultado, concepto: "Recomendado con observaciones", clase: "obs",
      detalle: `Registro INVIMA no evidenciado.${anotacion}`,
    };
  }
  return {
    ...resultado, concepto: "Recomendado", clase: "si",
    detalle: `Cumple con especificaciones.${anotacion}`,
  };
}

// Fecha de hoy en tu zona horaria, como AAAA-MM-DD
function fechaLocal() {
  const hoy = new Date();
  const dos = (n) => String(n).padStart(2, "0");
  return `${hoy.getFullYear()}-${dos(hoy.getMonth() + 1)}-${dos(hoy.getDate())}`;
}
