// ---------- Informe técnico en Word (usa la librería docx) ----------

const AZUL_TABLA = "0B5C8A";
const RELLENO_CONCEPTO = { si: "C6E0B4", obs: "FFE699", no: "F8CBAD", pendiente: "D9D9D9" };
const BORDE_TABLA = { style: "single", size: 4, color: "BFBFBF" };
const PENDIENTE_TEXTO = "[completar]";

function plural(n, uno, varios) {
  return n === 1 ? uno : varios;
}

// "Proveedor — Marca - modelo"
function nombreInforme(oferta, indice) {
  const partes = [oferta.proveedor.trim(), oferta.modelo.trim()].filter(Boolean);
  return partes.length > 0 ? partes.join(" — ") : `Oferta ${indice + 1}`;
}

function fechaLarga(iso) {
  const [a, m, d] = (iso || "").split("-").map(Number);
  if (!a || !m || !d) return PENDIENTE_TEXTO;
  return new Date(a, m - 1, d).toLocaleDateString("es-CO", { day: "numeric", month: "long", year: "numeric" });
}

// ---- Piezas básicas del documento ----
function corrida(texto, op = {}) {
  const pendiente = /^\[.*\]$/.test(texto); // lo que va entre corchetes se resalta
  return new docx.TextRun({
    text: texto,
    bold: op.negrita,
    italics: op.cursiva,
    color: op.color,
    size: op.tamano,
    highlight: pendiente ? "yellow" : undefined,
  });
}

// "partes" puede ser un texto o una lista de textos / { t, negrita }
function corridas(partes, op = {}) {
  return (Array.isArray(partes) ? partes : [partes]).map((p) =>
    typeof p === "string" ? corrida(p, op) : corrida(p.t, { ...op, negrita: p.negrita }));
}

function parrafo(partes, op = {}) {
  return new docx.Paragraph({
    alignment: op.alinear ?? docx.AlignmentType.JUSTIFIED,
    spacing: { before: op.antes ?? 0, after: op.despues ?? 120 },
    keepNext: op.mantener,
    children: corridas(partes, op),
  });
}

function vineta(partes) {
  return new docx.Paragraph({
    numbering: { reference: "vinetas", level: 0 },
    alignment: docx.AlignmentType.JUSTIFIED,
    spacing: { after: 80 },
    children: corridas(partes),
  });
}

function titulo1(texto, nuevaPagina = false) {
  return new docx.Paragraph({
    heading: docx.HeadingLevel.HEADING_1,
    pageBreakBefore: nuevaPagina,
    children: [new docx.TextRun(texto)],
  });
}

function titulo2(texto) {
  return new docx.Paragraph({ heading: docx.HeadingLevel.HEADING_2, children: [new docx.TextRun(texto)] });
}

function celda(contenido, ancho, op = {}) {
  const lineas = Array.isArray(contenido) ? contenido : [contenido];
  return new docx.TableCell({
    width: { size: ancho, type: docx.WidthType.DXA },
    columnSpan: op.columnas,
    verticalAlign: docx.VerticalAlign.CENTER,
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
    shading: op.relleno ? { type: docx.ShadingType.CLEAR, fill: op.relleno, color: "auto" } : undefined,
    children: lineas.map((t) => new docx.Paragraph({
      alignment: op.alinear ?? docx.AlignmentType.LEFT,
      spacing: { after: 40 },
      children: [corrida(String(t), { negrita: op.negrita, color: op.color, tamano: 20 })],
    })),
  });
}

// Cada celda puede ser un texto, una lista de textos o { t, relleno, negrita, alinear, columnas }
function fila(anchos, celdas) {
  let i = 0;
  return new docx.TableRow({
    cantSplit: true,
    children: celdas.map((c) => {
      const spec = typeof c === "string" || Array.isArray(c) ? { t: c } : c;
      const tramo = spec.columnas ?? 1;
      const ancho = anchos.slice(i, i + tramo).reduce((a, b) => a + b, 0);
      i += tramo;
      return celda(spec.t, ancho, spec);
    }),
  });
}

function tabla(anchos, titulos, filas) {
  const total = anchos.reduce((a, b) => a + b, 0);
  return new docx.Table({
    width: { size: total, type: docx.WidthType.DXA },
    columnWidths: anchos,
    layout: docx.TableLayoutType.FIXED,
    borders: {
      top: BORDE_TABLA, bottom: BORDE_TABLA, left: BORDE_TABLA, right: BORDE_TABLA,
      insideHorizontal: BORDE_TABLA, insideVertical: BORDE_TABLA,
    },
    rows: [
      new docx.TableRow({
        tableHeader: true,
        cantSplit: true,
        children: titulos.map((t, k) => celda(t, anchos[k], { relleno: AZUL_TABLA, color: "FFFFFF", negrita: true })),
      }),
      ...filas,
    ],
  });
}

function espacio() {
  return new docx.Paragraph({ spacing: { after: 120 }, children: [] });
}

// ---- Secciones ----
function tablaDatosGenerales(general) {
  const anchos = [2600, 6760];
  const dato = (valor) => (valor && valor.trim() ? valor.trim() : PENDIENTE_TEXTO);
  const filaDato = (nombre, valor) => fila(anchos, [{ t: nombre, negrita: true, relleno: "F2F2F2" }, valor]);
  return new docx.Table({
    width: { size: 9360, type: docx.WidthType.DXA },
    columnWidths: anchos,
    layout: docx.TableLayoutType.FIXED,
    borders: {
      top: BORDE_TABLA, bottom: BORDE_TABLA, left: BORDE_TABLA, right: BORDE_TABLA,
      insideHorizontal: BORDE_TABLA, insideVertical: BORDE_TABLA,
    },
    rows: [
      filaDato("Proyecto o entidad", dato(general.proyecto)),
      filaDato("Proceso o referencia", dato(general.proceso)),
      filaDato("Fecha", fechaLarga(general.fecha)),
      filaDato("Elaborado por", dato(general.elaboradoPor)),
    ],
  });
}

function seccionObjetivo(res, general) {
  const n = res.total;
  const proyecto = general.proyecto && general.proyecto.trim() ? ` del proyecto ${general.proyecto.trim()}` : "";
  return [
    titulo1("1. OBJETIVO"),
    parrafo(
      `El presente anexo técnico tiene como propósito consolidar los resultados de la evaluación técnica de las ofertas presentadas por los proveedores para ${n} ${plural(n, "equipo biomédico", "equipos biomédicos")}${proyecto}, con el fin de verificar el cumplimiento de las especificaciones técnicas mínimas exigidas y de la documentación regulatoria. El análisis establece cuáles equipos cuentan con al menos una oferta técnicamente aprobada, cuáles se encuentran pendientes de la verificación del registro INVIMA y cuáles no cumplen con los requisitos mínimos.`
    ),
  ];
}

function seccionMetodologia(res) {
  const { total, requisitos, ofertas } = res;
  return [
    titulo1("2. METODOLOGÍA"),
    parrafo(`Se evaluaron ${total} ${plural(total, "equipo", "equipos")}, con un total de ${requisitos} ${plural(requisitos, "requisito técnico", "requisitos técnicos")} y ${ofertas.total} ${plural(ofertas.total, "oferta", "ofertas")} de proveedores. La evaluación se realizó así:`, { mantener: true }),
    vineta([{ t: "Evaluación cuantitativa. ", negrita: true }, "Cada requisito técnico se calificó de forma cuantitativa: 0 (no cumple), 1 (cumple) y 2 (cumple por encima de lo exigido). La calificación 2 se registra como observación y no suma al porcentaje de cumplimiento."]),
    vineta([{ t: "Pesos. ", negrita: true }, "A cada requisito se le asignó un peso. El porcentaje de cumplimiento de cada oferta corresponde a la suma de los pesos de los requisitos cumplidos sobre la suma total de los pesos."]),
    vineta([{ t: "Verificación normativa. ", negrita: true }, "Se verificó, para cada oferta, la evidencia del registro INVIMA y la existencia de alertas sanitarias."]),
    vineta([{ t: "Concepto técnico. ", negrita: true }, "Recomendado: cumple el 100 % de los requisitos técnicos, cuenta con registro INVIMA y no presenta alerta sanitaria. Recomendado con observaciones: cumple el 100 % de los requisitos técnicos, pero no se evidenció el registro INVIMA. No recomendado: incumple al menos un requisito técnico o presenta alerta sanitaria."]),
    vineta([{ t: "Comparación entre ofertas. ", negrita: true }, "Cuando un equipo tiene más de una oferta recomendada, estas se ordenan según el peso de los requisitos que superan lo exigido, como criterio de desempate orientativo. Si el valor es igual, se considera un empate técnico."]),
  ];
}

function seccionResultados(res) {
  const { total, grupos, ofertas } = res;
  const g = grupos;
  const anchos = [5400, 1980, 1980];
  const filaCifra = (nombre, n, op = {}) =>
    fila(anchos, [
      { t: nombre, negrita: op.negrita },
      { t: String(n), alinear: docx.AlignmentType.CENTER, negrita: op.negrita },
      { t: porcentajeTexto(n, op.base ?? total), alinear: docx.AlignmentType.CENTER, negrita: op.negrita },
    ]);
  const subtitulo = (texto) =>
    new docx.Paragraph({
      alignment: docx.AlignmentType.CENTER,
      spacing: { before: 200, after: 100 },
      keepNext: true,
      children: [new docx.TextRun({ text: texto, bold: true, italics: true })],
    });

  return [
    titulo1("3. RESULTADOS DE LA EVALUACIÓN"),
    parrafo("A continuación se presenta el resumen cuantitativo de la evaluación:", { mantener: true }),
    subtitulo("RESUMEN DE CIFRAS POR EQUIPO"),
    tabla(anchos, ["Clasificación", "Cantidad", "% del total"], [
      filaCifra("Equipos aprobados", g.aprobados.length),
      filaCifra("Equipos pendientes por INVIMA", g.pendientesInvima.length),
      filaCifra("Equipos que no cumplen", g.noCumplen.length),
      filaCifra("Equipos con evaluación incompleta", g.incompletos.length),
      filaCifra("TOTAL DE EQUIPOS", total, { negrita: true }),
    ]),
    subtitulo("RESUMEN DE CIFRAS POR OFERTA"),
    tabla(anchos, ["Concepto técnico de la oferta", "Cantidad", "% del total"], [
      filaCifra("Ofertas recomendadas", ofertas.recomendadas, { base: ofertas.total }),
      filaCifra("Ofertas recomendadas con observaciones", ofertas.conObservaciones, { base: ofertas.total }),
      filaCifra("Ofertas no recomendadas", ofertas.noRecomendadas, { base: ofertas.total }),
      filaCifra("Ofertas con evaluación incompleta", ofertas.incompletas, { base: ofertas.total }),
      filaCifra("TOTAL DE OFERTAS", ofertas.total, { negrita: true, base: ofertas.total }),
    ]),
  ];
}

function bloqueClasificacion(titulo, intro, registros, anchos, titulos, filaDe) {
  const partes = [titulo2(titulo)];
  if (registros.length === 0) {
    partes.push(parrafo("No hay equipos en esta categoría."));
    return partes;
  }
  partes.push(parrafo(intro, { mantener: true }));
  partes.push(tabla(anchos, titulos, registros.map((r, k) => fila(anchos, filaDe(r, k)))));
  partes.push(espacio());
  return partes;
}

function seccionClasificacion(res) {
  const g = res.grupos;
  const itemDe = (r, k) => r.equipo.item || String(k + 1);
  const equipoDe = (r) => ({ t: r.equipo.equipo, negrita: true });
  const n = (lista) => lista.length;

  const aprobados = bloqueClasificacion(
    "4.1 Equipos aprobados",
    `Los siguientes ${n(g.aprobados)} ${plural(n(g.aprobados), "equipo cuenta", "equipos cuentan")} con al menos una oferta con concepto técnico Recomendado: cumplen el 100 % de las especificaciones técnicas mínimas exigidas, cuentan con registro INVIMA y no presentan alerta sanitaria.`,
    g.aprobados, [700, 2800, 3500, 2360], ["Ítem", "Equipo", "Oferta(s) recomendada(s)", "Observaciones"],
    (r, k) => {
      const varias = r.rec.length > 1;
      const ofertas = r.rec.map((f) => `${varias ? f.posicion + ".º " : ""}${nombreInforme(f.oferta, f.indice)}${f.empate ? " (empate)" : ""}`);
      const observaciones = [];
      if (varias) observaciones.push("Varias recomendadas; orden por peso superado (orientativo).");
      r.obs.forEach((f) => observaciones.push(`Pendiente INVIMA: ${nombreInforme(f.oferta, f.indice)}`));
      return [itemDe(r, k), equipoDe(r), { t: ofertas, relleno: RELLENO_CONCEPTO.si }, observaciones.length > 0 ? observaciones : "—"];
    }
  );

  const pendientes = bloqueClasificacion(
    "4.2 Equipos pendientes por INVIMA",
    `Los siguientes ${n(g.pendientesInvima)} ${plural(n(g.pendientesInvima), "equipo cumple", "equipos cumplen")} el 100 % de las especificaciones técnicas mínimas exigidas, pero ninguna de sus ofertas evidenció el registro INVIMA. Podrán consolidar su aprobación una vez se verifique el registro.`,
    g.pendientesInvima, [700, 3300, 5360], ["Ítem", "Equipo", "Oferta(s) con registro INVIMA pendiente"],
    (r, k) => [itemDe(r, k), equipoDe(r), { t: r.obs.map((f) => nombreInforme(f.oferta, f.indice)), relleno: RELLENO_CONCEPTO.obs }]
  );

  const noCumplen = bloqueClasificacion(
    "4.3 Equipos que no cumplen",
    `Los siguientes ${n(g.noCumplen)} ${plural(n(g.noCumplen), "equipo no cuenta", "equipos no cuentan")} con ninguna oferta que cumpla las especificaciones técnicas mínimas exigidas, o sus ofertas presentan alerta sanitaria.`,
    g.noCumplen, [700, 2600, 6060], ["Ítem", "Equipo", "Ofertas evaluadas y motivo"],
    (r, k) => [itemDe(r, k), equipoDe(r), { t: r.no.map((f) => `${nombreInforme(f.oferta, f.indice)}: ${f.r.detalle}`), relleno: RELLENO_CONCEPTO.no }]
  );

  const incompletos = bloqueClasificacion(
    "4.4 Equipos con evaluación incompleta",
    `Los siguientes ${n(g.incompletos)} ${plural(n(g.incompletos), "equipo no cuenta", "equipos no cuentan")} con una evaluación completa: no tienen ofertas registradas o tienen ofertas con requisitos sin evaluar.`,
    g.incompletos, [700, 3300, 5360], ["Ítem", "Equipo", "Motivo"],
    (r, k) => [
      itemDe(r, k), equipoDe(r),
      { t: r.comparativo.filas.length === 0 ? "Sin ofertas registradas." : `Ofertas sin evaluar por completo: ${r.pend.map((f) => nombreInforme(f.oferta, f.indice)).join("; ")}.`, relleno: RELLENO_CONCEPTO.pendiente },
    ]
  );

  return [titulo1("4. CLASIFICACIÓN DETALLADA POR EQUIPO"), ...aprobados, ...pendientes, ...noCumplen, ...incompletos];
}

function seccionConclusiones(res) {
  const { total, grupos: g, ofertas } = res;
  const cumplenTecnico = g.aprobados.length + g.pendientesInvima.length;
  const linea = (nombre, cantidad, base = total) => vineta([{ t: `${nombre}: `, negrita: true }, `${cantidad} (${porcentajeTexto(cantidad, base)})`]);
  return [
    titulo1("5. CONCLUSIONES Y RECOMENDACIONES"),
    parrafo("Cifras de la evaluación:", { mantener: true }),
    vineta([{ t: "Equipos evaluados: ", negrita: true }, String(total)]),
    linea("Equipos aprobados", g.aprobados.length),
    linea("Equipos pendientes por INVIMA", g.pendientesInvima.length),
    linea("Equipos con cumplimiento técnico satisfactorio (aprobados y pendientes por INVIMA)", cumplenTecnico),
    linea("Equipos que no cumplen", g.noCumplen.length),
    linea("Equipos con evaluación incompleta", g.incompletos.length),
    vineta([{ t: "Ofertas evaluadas: ", negrita: true }, `${ofertas.total}, de las cuales ${ofertas.recomendadas} (${porcentajeTexto(ofertas.recomendadas, ofertas.total)}) son recomendadas, ${ofertas.conObservaciones} (${porcentajeTexto(ofertas.conObservaciones, ofertas.total)}) son recomendadas con observaciones, ${ofertas.noRecomendadas} (${porcentajeTexto(ofertas.noRecomendadas, ofertas.total)}) no son recomendadas y ${ofertas.incompletas} (${porcentajeTexto(ofertas.incompletas, ofertas.total)}) tienen evaluación incompleta.`]),
    parrafo("[Redactar aquí el texto de las conclusiones y recomendaciones]", { antes: 200 }),
  ];
}

function seccionAnexoA(equipos) {
  const anchos = [700, 3000, 1000, 1500, 900, 2260];
  const centro = docx.AlignmentType.CENTER;
  const filas = [];
  equipos.forEach((equipo) => {
    filas.push(fila(anchos, [{ t: equipo.item ? `${equipo.item}. ${equipo.equipo}` : equipo.equipo, negrita: true, relleno: "DDEBF7", columnas: 6 }]));
    if (equipo.ofertas.length === 0) {
      filas.push(fila(anchos, [{ t: "Sin ofertas registradas.", columnas: 6 }]));
      return;
    }
    compararOfertas(equipo).filas.forEach((f) => {
      filas.push(fila(anchos, [
        { t: f.posicion === null ? "—" : String(f.posicion), alinear: centro },
        nombreInforme(f.oferta, f.indice),
        { t: `${f.r.cumplidos} de ${f.r.total}`, alinear: centro },
        { t: `${Math.round(f.r.porcentaje * 100)} %`, alinear: centro },
        { t: f.oferta.invima ? "Sí" : "No", alinear: centro },
        { t: f.r.concepto, relleno: RELLENO_CONCEPTO[f.r.clase], negrita: true, alinear: centro },
      ]));
    });
  });
  return [
    titulo1("ANEXO A. DETALLE DE LA EVALUACIÓN POR OFERTA", true),
    parrafo("La columna Pos. indica el orden entre las ofertas recomendadas de cada equipo.", { mantener: true }),
    tabla(anchos, ["Pos.", "Oferta (proveedor — marca - modelo)", "Cumple", "Cumplimiento", "INVIMA", "Concepto técnico"], filas),
  ];
}

// ---- Documento completo ----
async function exportarInforme() {
  if (typeof docx === "undefined") {
    throw new Error('No se encontró la librería de Word. Revisa que el archivo "docx.iife.js" esté en la carpeta.');
  }
  const res = clasificarEquipos(equipos);
  const proyecto = datosGenerales.proyecto.trim();

  const documento = new docx.Document({
    creator: "Evaluador técnico",
    title: "Anexo técnico — Evaluación técnica de ofertas",
    styles: {
      default: { document: { run: { font: "Calibri", size: 22 } } },
      paragraphStyles: [
        { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { font: "Calibri", size: 24, bold: true, color: "1F2933" },
          paragraph: { spacing: { before: 320, after: 120 }, keepNext: true } },
        { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { font: "Calibri", size: 22, bold: true, color: AZUL_TABLA },
          paragraph: { spacing: { before: 240, after: 100 }, keepNext: true } },
      ],
    },
    numbering: {
      config: [{
        reference: "vinetas",
        levels: [{ level: 0, format: docx.LevelFormat.BULLET, text: "•", alignment: docx.AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } } }],
      }],
    },
    sections: [{
      properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } },
      headers: {
        default: new docx.Header({ children: [new docx.Paragraph({
          alignment: docx.AlignmentType.RIGHT,
          children: [new docx.TextRun({ text: proyecto ? `Anexo técnico — ${proyecto}` : "Anexo técnico", color: "7B8794", size: 18 })],
        })] }),
      },
      footers: {
        default: new docx.Footer({ children: [new docx.Paragraph({
          alignment: docx.AlignmentType.CENTER,
          children: [new docx.TextRun({ children: ["Página ", docx.PageNumber.CURRENT, " de ", docx.PageNumber.TOTAL_PAGES], color: "7B8794", size: 18 })],
        })] }),
      },
      children: [
        new docx.Paragraph({
          alignment: docx.AlignmentType.CENTER,
          spacing: { after: 240 },
          children: [new docx.TextRun({ text: "ANEXO TÉCNICO — EVALUACIÓN TÉCNICA DE OFERTAS DE EQUIPOS BIOMÉDICOS", bold: true, size: 26 })],
        }),
        tablaDatosGenerales(datosGenerales),
        ...seccionObjetivo(res, datosGenerales),
        ...seccionMetodologia(res),
        ...seccionResultados(res),
        ...seccionClasificacion(res),
        ...seccionConclusiones(res),
        ...seccionAnexoA(equipos),
      ],
    }],
  });

  const blob = await docx.Packer.toBlob(documento);
  descargarArchivo(blob, `informe-tecnico-${fechaLocal()}.docx`);
}
