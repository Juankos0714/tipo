// ===== CONFIGURACION =====
const FILA_ENCABEZADO = 1;   // Fila donde estan los titulos
const TOLERANCIA = 1;        // Diferencia permitida en pesos (redondeos)

// Titulos de las columnas (no importan mayusculas, tildes ni puntos)
const COLS: { [clave: string]: string } = {
  facturaNota: "Fact./Nota",
  estado: "Estado",
  razon: "Razón social cliente factura",
  bruto: "Valor bruto",
  descuentos: "Valor descuentos",
  subtotal: "Valor subtotal",
  impuestos: "Valor impuestos",
  neto: "Valor neto",
  credito: "Valor crédito",
  contado: "Valor contado",
  efectivo: "Efectivo",
  cheque: "Cheque",
  tarjetaCredito: "Tarjeta crédito",
  tarjetaDebito: "Tarjeta débito",
  otros: "Otros",
  facturaElectronica: "Factura electrónica"
};

type Celda = string | number | boolean;

function normalizar(texto: string): string {
  return texto
    .toUpperCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[.,]/g, "")
    .replace(/\s+/g, " ")
    .trim();
}

// Convierte una celda a numero (acepta "$55.100,00" si llegara como texto)
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

  // Ubicar columnas por titulo
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

  // Ubicar o crear las columnas de resultado
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

  // Contar Fact./Nota para detectar repetidas
  const conteo: { [clave: string]: number } = {};
  for (let i = 0; i < n; i++) {
    const f = String(datos[idxEnc + 1 + i][c.facturaNota]).trim();
    if (f !== "") conteo[f] = (conteo[f] || 0) + 1;
  }

  // Evaluar cada fila
  const salidaSev: string[][] = [];
  const salidaObs: string[][] = [];
  const filasError: number[] = [];
  const filasSosp: number[] = [];

  for (let i = 0; i < n; i++) {
    const fila = datos[idxEnc + 1 + i];
    const errores: string[] = [];
    const sospechas: string[] = [];

    const factura = String(fila[c.facturaNota]).trim();
    const razon = String(fila[c.razon]).trim();
    const estado = normalizar(String(fila[c.estado]));
    const electronica = String(fila[c.facturaElectronica]).trim();

    const bruto = numero(fila[c.bruto]);
    const descuentos = numero(fila[c.descuentos]);
    const subtotal = numero(fila[c.subtotal]);
    const impuestos = numero(fila[c.impuestos]);
    const neto = numero(fila[c.neto]);
    const credito = numero(fila[c.credito]);
    const contado = numero(fila[c.contado]);
    const pagos = numero(fila[c.efectivo]) + numero(fila[c.cheque]) +
      numero(fila[c.tarjetaCredito]) + numero(fila[c.tarjetaDebito]) + numero(fila[c.otros]);

    if (razon === "") errores.push("Razon social del cliente vacia");
    if (normalizar(electronica) !== "CONFIRMADOS") {
      errores.push("Factura electronica no confirmada (" + (electronica || "vacio") + ")");
    }
    if (estado !== "APROBADAS") sospechas.push("Estado distinto de Aprobadas (" + String(fila[c.estado]) + ")");
    if (distintos(bruto - descuentos, subtotal)) errores.push("Bruto - descuentos no coincide con subtotal");
    if (distintos(subtotal + impuestos, neto)) errores.push("Subtotal + impuestos no coincide con valor neto");
    if (distintos(contado + credito, neto)) errores.push("Contado + credito no coincide con valor neto");
    if (distintos(pagos, contado)) errores.push("Formas de pago (" + pagos + ") no suman el valor contado (" + contado + ")");
    if (factura !== "" && conteo[factura] > 1) errores.push("Fact./Nota repetida");

    const todas = errores.concat(sospechas);
    const sev = errores.length > 0 ? "ERROR" : sospechas.length > 0 ? "SOSPECHOSO" : "OK";
    salidaSev.push([sev]);
    salidaObs.push([todas.length > 0 ? todas.join("; ") : "OK"]);
    if (sev === "ERROR") filasError.push(idxEnc + 1 + i);
    if (sev === "SOSPECHOSO") filasSosp.push(idxEnc + 1 + i);
  }

  // Escribir resultados de una sola vez
  const rangoSev = hoja.getRangeByIndexes(idxEnc + 1, cSev, n, 1);
  const rangoObs = hoja.getRangeByIndexes(idxEnc + 1, cObs, n, 1);
  rangoSev.getFormat().getFill().clear();
  rangoObs.getFormat().getFill().clear();
  rangoSev.setValues(salidaSev);
  rangoObs.setValues(salidaObs);

  // Colorear: rojo = error, amarillo = sospechoso
  for (const f of filasError) {
    hoja.getRangeByIndexes(f, cSev, 1, 2).getFormat().getFill().setColor("#FFC7CE");
  }
  for (const f of filasSosp) {
    hoja.getRangeByIndexes(f, cSev, 1, 2).getFormat().getFill().setColor("#FFEB9C");
  }

  // Ensanchar columnas (soluciona los ##### en Efectivo si lo ejecutas con la hoja activa)
  hoja.getRange("P:P").getFormat().autofitColumns();

  console.log("Filas revisadas: " + n + " | Errores: " + filasError.length + " | Sospechosos: " + filasSosp.length);
}
