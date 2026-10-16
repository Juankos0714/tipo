const FILA_ENCABEZADO = 1;
const TOLERANCIA = 1;
const EJEMPLOS_POR_REGLA = 5;

const COLS: { [clave: string]: string } = {
  facturaNota: "Fact./Nota",
  estado: "Estado",
  razon: "Razón social cliente factura",
  bruto: "Valor bruto",
  descuentos: "Valor descuentos",
  subtotal: "Valor subtotal",
  impuestos: "Valor impuestos",
  neto: "Valor neto",
  vendedor: "Nombre vendedor",
  credito: "Valor crédito",
  contado: "Valor contado",
  cajero: "Nombre cajero",
  efectivo: "Efectivo",
  cheque: "Cheque",
  tarjetaCredito: "Tarjeta crédito",
  tarjetaDebito: "Tarjeta débito",
  otros: "Otros",
  facturaElectronica: "Factura electrónica"
};

const R_RAZON = "Razón social vacía";
const R_ELECTRONICA = "Factura electrónica no confirmada";
const R_ESTADO = "Estado distinto de Aprobadas";
const R_SUBTOTAL = "Bruto - descuentos distinto de subtotal";
const R_NETO = "Subtotal + impuestos distinto de valor neto";
const R_CREDCONT = "Contado + crédito distinto de valor neto";
const R_PAGOS = "Formas de pago distintas del valor contado";
const R_REPETIDA = "Fact./Nota repetida";

const SEVERIDAD_REGLA: { [regla: string]: string } = {};
SEVERIDAD_REGLA[R_RAZON] = "ERROR";
SEVERIDAD_REGLA[R_ELECTRONICA] = "ERROR";
SEVERIDAD_REGLA[R_ESTADO] = "SOSPECHOSO";
SEVERIDAD_REGLA[R_SUBTOTAL] = "ERROR";
SEVERIDAD_REGLA[R_NETO] = "ERROR";
SEVERIDAD_REGLA[R_CREDCONT] = "ERROR";
SEVERIDAD_REGLA[R_PAGOS] = "ERROR";
SEVERIDAD_REGLA[R_REPETIDA] = "ERROR";

type Celda = string | number | boolean;

interface Hallazgo {
  severidad: string;
  fila: number;
  factura: string;
  razon: string;
  cajero: string;
  vendedor: string;
  regla: string;
  detalle: string;
  esperado: number | string;
  encontrado: number | string;
}

interface ConteoPersona {
  total: number;
  conError: number;
}

function normalizar(texto: string): string {
  return texto
    .toUpperCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[.,]/g, "")
    .replace(/\s+/g, " ")
    .trim();
}

function numero(v: Celda): number {
  if (typeof v === "number") return v;
  const t = String(v).replace(/[$\s]/g, "").replace(/\./g, "").replace(",", ".");
  const n = parseFloat(t);
  return isNaN(n) ? 0 : n;
}

function distintos(a: number, b: number): boolean {
  return Math.abs(a - b) > TOLERANCIA;
}

function main(workbook: ExcelScript.Workbook) {
  const hoja = workbook.getActiveWorksheet();
  const usado = hoja.getUsedRange();
  const totalFilas = usado.getRowIndex() + usado.getRowCount();
  const totalCols = usado.getColumnIndex() + usado.getColumnCount();
  const datos = hoja.getRangeByIndexes(0, 0, totalFilas, totalCols).getValues();

  const idxEnc = FILA_ENCABEZADO - 1;
  const titulos = datos[idxEnc].map((v) => normalizar(String(v)));
  const c: { [clave: string]: number } = {};
  const faltan: string[] = [];
  for (const clave of Object.keys(COLS)) {
    c[clave] = titulos.indexOf(normalizar(COLS[clave]));
    if (c[clave] < 0) faltan.push(COLS[clave]);
  }
  if (faltan.length > 0) {
    console.log("No encontre estas columnas: " + faltan.join(" | "));
    console.log("Titulos disponibles: " + datos[idxEnc].join(" | "));
    return;
  }

  let siguiente = totalCols;
  let cSev = titulos.indexOf("SEVERIDAD");
  if (cSev < 0) {
    cSev = siguiente++;
    hoja.getCell(idxEnc, cSev).setValue("Severidad");
  }
  let cObs = titulos.indexOf("OBSERVACIONES");
  if (cObs < 0) {
    cObs = siguiente++;
    hoja.getCell(idxEnc, cObs).setValue("Observaciones");
  }

  const n = totalFilas - idxEnc - 1;
  if (n <= 0) {
    console.log("No hay datos debajo del encabezado.");
    return;
  }

  const conteoFactura: { [clave: string]: number } = {};
  for (let i = 0; i < n; i++) {
    const f = String(datos[idxEnc + 1 + i][c.facturaNota]).trim();
    if (f !== "") conteoFactura[f] = (conteoFactura[f] || 0) + 1;
  }

  const hallazgos: Hallazgo[] = [];
  const salidaSev: string[][] = [];
  const salidaObs: string[][] = [];
  const filasError: number[] = [];
  const filasSosp: number[] = [];
  const porCajero: { [nombre: string]: ConteoPersona } = {};
  const porVendedor: { [nombre: string]: ConteoPersona } = {};

  for (let i = 0; i < n; i++) {
    const indice = idxEnc + 1 + i;
    const fila = datos[indice];
    const filaExcel = indice + 1;

    const factura = String(fila[c.facturaNota]).trim();
    const razon = String(fila[c.razon]).trim();
    const cajero = String(fila[c.cajero]).trim() || "(sin cajero)";
    const vendedor = String(fila[c.vendedor]).trim() || "(sin vendedor)";
    const estado = normalizar(String(fila[c.estado]));
    const electronica = String(fila[c.facturaElectronica]).trim();

    const bruto = numero(fila[c.bruto]);
    const descuentos = numero(fila[c.descuentos]);
    const subtotal = numero(fila[c.subtotal]);
    const impuestos = numero(fila[c.impuestos]);
    const neto = numero(fila[c.neto]);
    const credito = numero(fila[c.credito]);
    const contado = numero(fila[c.contado]);
    const efectivo = numero(fila[c.efectivo]);
    const cheque = numero(fila[c.cheque]);
    const tCredito = numero(fila[c.tarjetaCredito]);
    const tDebito = numero(fila[c.tarjetaDebito]);
    const otros = numero(fila[c.otros]);
    const pagos = efectivo + cheque + tCredito + tDebito + otros;

    const reglasFila: string[] = [];
    const agregar = (regla: string, detalle: string, esperado: number | string, encontrado: number | string) => {
      reglasFila.push(regla);
      hallazgos.push({
        severidad: SEVERIDAD_REGLA[regla], fila: filaExcel, factura: factura, razon: razon,
        cajero: cajero, vendedor: vendedor, regla: regla, detalle: detalle,
        esperado: esperado, encontrado: encontrado
      });
    };

    if (razon === "") agregar(R_RAZON, "La columna de razón social está vacía", "", "");
    if (normalizar(electronica) !== "CONFIRMADOS") {
      agregar(R_ELECTRONICA, "Valor en Factura electrónica: " + (electronica || "(vacío)"), "Confirmados", electronica || "(vacío)");
    }
    if (estado !== "APROBADAS") {
      agregar(R_ESTADO, "Valor en Estado: " + String(fila[c.estado]), "Aprobadas", String(fila[c.estado]));
    }
    if (distintos(bruto - descuentos, subtotal)) {
      agregar(R_SUBTOTAL, "Bruto " + bruto + " - descuentos " + descuentos, bruto - descuentos, subtotal);
    }
    if (distintos(subtotal + impuestos, neto)) {
      agregar(R_NETO, "Subtotal " + subtotal + " + impuestos " + impuestos, subtotal + impuestos, neto);
    }
    if (distintos(contado + credito, neto)) {
      agregar(R_CREDCONT, "Contado " + contado + " + crédito " + credito, contado + credito, neto);
    }
    if (distintos(pagos, contado)) {
      agregar(R_PAGOS,
        "Efectivo " + efectivo + " | Cheque " + cheque + " | T.crédito " + tCredito +
        " | T.débito " + tDebito + " | Otros " + otros, contado, pagos);
    }
    if (factura !== "" && conteoFactura[factura] > 1) {
      agregar(R_REPETIDA, "Aparece " + conteoFactura[factura] + " veces", 1, conteoFactura[factura]);
    }

    let sev = "OK";
    for (const r of reglasFila) {
      if (SEVERIDAD_REGLA[r] === "ERROR") { sev = "ERROR"; break; }
      sev = "SOSPECHOSO";
    }
    salidaSev.push([sev]);
    salidaObs.push([reglasFila.length > 0 ? reglasFila.join("; ") : "OK"]);
    if (sev === "ERROR") filasError.push(indice);
    if (sev === "SOSPECHOSO") filasSosp.push(indice);

    if (!porCajero[cajero]) porCajero[cajero] = { total: 0, conError: 0 };
    if (!porVendedor[vendedor]) porVendedor[vendedor] = { total: 0, conError: 0 };
    porCajero[cajero].total++;
    porVendedor[vendedor].total++;
    if (sev !== "OK") {
      porCajero[cajero].conError++;
      porVendedor[vendedor].conError++;
    }
  }

  const rangoSev = hoja.getRangeByIndexes(idxEnc + 1, cSev, n, 1);
  const rangoObs = hoja.getRangeByIndexes(idxEnc + 1, cObs, n, 1);
  rangoSev.getFormat().getFill().clear();
  rangoObs.getFormat().getFill().clear();
  rangoSev.setValues(salidaSev);
  rangoObs.setValues(salidaObs);
  for (const f of filasError) hoja.getRangeByIndexes(f, cSev, 1, 2).getFormat().getFill().setColor("#FFC7CE");
  for (const f of filasSosp) hoja.getRangeByIndexes(f, cSev, 1, 2).getFormat().getFill().setColor("#FFEB9C");
  hoja.getRange("P:P").getFormat().autofitColumns();

  const reglas = Object.keys(SEVERIDAD_REGLA);
  const porRegla: { [regla: string]: Hallazgo[] } = {};
  for (const r of reglas) porRegla[r] = [];
  for (const h of hallazgos) porRegla[h.regla].push(h);

  const nombreDetalle = "Detalle validacion";
  const detalleVieja = workbook.getWorksheet(nombreDetalle);
  if (detalleVieja) detalleVieja.delete();
  const hojaDetalle = workbook.addWorksheet(nombreDetalle);
  const encDetalle = ["Severidad", "Fila Excel", "Fact./Nota", "Razón social", "Cajero", "Vendedor",
    "Regla", "Detalle", "Esperado", "Encontrado", "Diferencia"];
  hojaDetalle.getRangeByIndexes(0, 0, 1, encDetalle.length).setValues([encDetalle]);
  hojaDetalle.getRangeByIndexes(0, 0, 1, encDetalle.length).getFormat().getFont().setBold(true);
  if (hallazgos.length > 0) {
    const filasDetalle: (string | number)[][] = hallazgos.map((h) => [
      h.severidad, h.fila, h.factura, h.razon, h.cajero, h.vendedor, h.regla, h.detalle,
      h.esperado, h.encontrado,
      (typeof h.esperado === "number" && typeof h.encontrado === "number") ? h.encontrado - h.esperado : ""
    ]);
    hojaDetalle.getRangeByIndexes(1, 0, filasDetalle.length, encDetalle.length).setValues(filasDetalle);
  }
  hojaDetalle.getFreezePanes().freezeRows(1);
  hojaDetalle.getUsedRange().getFormat().autofitColumns();

  const nombreResumen = "Resumen validacion";
  const resumenViejo = workbook.getWorksheet(nombreResumen);
  if (resumenViejo) resumenViejo.delete();
  const hojaResumen = workbook.addWorksheet(nombreResumen);

  const filasResumen: (string | number)[][] = [];
  const filasNegrita: number[] = [];
  const nuevaFila = (a: string | number, b: string | number, c2: string | number, d: string | number) => {
    filasResumen.push([a, b, c2, d]);
  };

  filasNegrita.push(filasResumen.length);
  nuevaFila("RESUMEN DE VALIDACIÓN", "", "", "");
  nuevaFila("Filas revisadas", n, "", "");
  nuevaFila("Filas con ERROR", filasError.length, (filasError.length * 100 / n).toFixed(1) + "%", "");
  nuevaFila("Filas SOSPECHOSAS", filasSosp.length, (filasSosp.length * 100 / n).toFixed(1) + "%", "");
  nuevaFila("Filas OK", n - filasError.length - filasSosp.length, "", "");
  nuevaFila("", "", "", "");

  filasNegrita.push(filasResumen.length);
  nuevaFila("Regla", "Severidad", "Cantidad", "% de filas");
  for (const r of reglas) {
    const cant = porRegla[r].length;
    nuevaFila(r, SEVERIDAD_REGLA[r], cant, (cant * 100 / n).toFixed(1) + "%");
  }
  nuevaFila("", "", "", "");

  const agregarPersonas = (titulo: string, mapa: { [nombre: string]: ConteoPersona }) => {
    filasNegrita.push(filasResumen.length);
    nuevaFila(titulo, "Facturas", "Con hallazgos", "% con hallazgos");
    const lista = Object.keys(mapa).sort((a, b) => mapa[b].conError - mapa[a].conError);
    for (const nombre of lista) {
      const p = mapa[nombre];
      nuevaFila(nombre, p.total, p.conError, (p.conError * 100 / p.total).toFixed(1) + "%");
    }
    nuevaFila("", "", "", "");
  };
  agregarPersonas("Por cajero", porCajero);
  agregarPersonas("Por vendedor", porVendedor);

  hojaResumen.getRangeByIndexes(0, 0, filasResumen.length, 4).setValues(filasResumen);
  for (const f of filasNegrita) hojaResumen.getRangeByIndexes(f, 0, 1, 4).getFormat().getFont().setBold(true);
  hojaResumen.getUsedRange().getFormat().autofitColumns();
  hojaResumen.activate();

  console.log("Filas revisadas: " + n + " | Errores: " + filasError.length + " | Sospechosos: " + filasSosp.length);
  console.log("Total de hallazgos: " + hallazgos.length);
  console.log("---- Por regla ----");
  for (const r of reglas) {
    const lista = porRegla[r];
    console.log(r + " [" + SEVERIDAD_REGLA[r] + "]: " + lista.length);
    for (let k = 0; k < Math.min(EJEMPLOS_POR_REGLA, lista.length); k++) {
      const h = lista[k];
      console.log("   fila " + h.fila + " | " + h.factura + " | esperado " + h.esperado + " | encontrado " + h.encontrado + " | " + h.detalle);
    }
  }
}
