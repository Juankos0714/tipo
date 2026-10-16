import re
import unicodedata

df = xl("Sheet1!A1:W2000", headers=True).dropna(how="all")
TOLERANCIA = 1

COLS = {
    "factura": "Fact./Nota",
    "estado": "Estado",
    "razon": "Razón social cliente factura",
    "bruto": "Valor bruto",
    "descuentos": "Valor descuentos",
    "subtotal": "Valor subtotal",
    "impuestos": "Valor impuestos",
    "neto": "Valor neto",
    "vendedor": "Nombre vendedor",
    "credito": "Valor crédito",
    "contado": "Valor contado",
    "cajero": "Nombre cajero",
    "efectivo": "Efectivo",
    "cheque": "Cheque",
    "tcredito": "Tarjeta crédito",
    "tdebito": "Tarjeta débito",
    "otros": "Otros",
    "electronica": "Factura electrónica",
}

R_RAZON = "Razón social vacía"
R_ELEC = "Factura electrónica no confirmada"
R_ESTADO = "Estado distinto de Aprobadas"
R_SUBTOTAL = "Bruto - descuentos distinto de subtotal"
R_NETO = "Subtotal + impuestos distinto de valor neto"
R_CREDCONT = "Contado + crédito distinto de valor neto"
R_PAGOS = "Formas de pago distintas del valor contado"
R_REPETIDA = "Fact./Nota repetida"

REGLAS = [
    (R_RAZON, "ERROR"),
    (R_ELEC, "ERROR"),
    (R_ESTADO, "SOSPECHOSO"),
    (R_SUBTOTAL, "ERROR"),
    (R_NETO, "ERROR"),
    (R_CREDCONT, "ERROR"),
    (R_PAGOS, "ERROR"),
    (R_REPETIDA, "ERROR"),
]
SEV = dict(REGLAS)


def norm(t):
    t = unicodedata.normalize("NFD", str(t))
    t = "".join(ch for ch in t if unicodedata.category(ch) != "Mn")
    return re.sub(r"\s+", " ", re.sub(r"[.,]", "", t)).strip().upper()


def txt(v):
    return "" if pd.isna(v) else str(v).strip()


def num(v):
    if pd.isna(v):
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    t = re.sub(r"[$\s]", "", str(v))
    if "," in t:
        t = t.replace(".", "").replace(",", ".")
    try:
        return float(t)
    except ValueError:
        return 0.0


def fmt(x):
    return f"{x:,.2f}".rstrip("0").rstrip(".")


mapa = {norm(col): col for col in df.columns}
faltan = [v for v in COLS.values() if norm(v) not in mapa]
if faltan:
    raise ValueError(f"Faltan columnas: {faltan}. Disponibles: {list(df.columns)}")
c = {k: mapa[norm(v)] for k, v in COLS.items()}

facturas = df[c["factura"]].map(txt)
repeticiones = facturas.map(facturas[facturas != ""].value_counts()).fillna(0).astype(int)


def evaluar(r, veces):
    t = {k: txt(r[c[k]]) for k in ("factura", "estado", "razon", "electronica", "cajero", "vendedor")}
    v = {k: num(r[c[k]]) for k in ("bruto", "descuentos", "subtotal", "impuestos", "neto",
                                   "credito", "contado", "efectivo", "cheque", "tcredito",
                                   "tdebito", "otros")}
    pagos = v["efectivo"] + v["cheque"] + v["tcredito"] + v["tdebito"] + v["otros"]
    salida = []

    def add(regla, detalle, esperado, encontrado):
        salida.append((regla, SEV[regla], detalle, esperado, encontrado))

    if t["razon"] == "":
        add(R_RAZON, "La columna de razón social está vacía", "", "")
    if norm(t["electronica"]) != "CONFIRMADOS":
        valor = t["electronica"] or "(vacío)"
        add(R_ELEC, f"Valor en Factura electrónica: {valor}", "Confirmados", valor)
    if norm(t["estado"]) != "APROBADAS":
        valor = t["estado"] or "(vacío)"
        add(R_ESTADO, f"Valor en Estado: {valor}", "Aprobadas", valor)
    esperado = v["bruto"] - v["descuentos"]
    if abs(esperado - v["subtotal"]) > TOLERANCIA:
        add(R_SUBTOTAL, f"Bruto {fmt(v['bruto'])} - descuentos {fmt(v['descuentos'])}", esperado, v["subtotal"])
    esperado = v["subtotal"] + v["impuestos"]
    if abs(esperado - v["neto"]) > TOLERANCIA:
        add(R_NETO, f"Subtotal {fmt(v['subtotal'])} + impuestos {fmt(v['impuestos'])}", esperado, v["neto"])
    esperado = v["contado"] + v["credito"]
    if abs(esperado - v["neto"]) > TOLERANCIA:
        add(R_CREDCONT, f"Contado {fmt(v['contado'])} + crédito {fmt(v['credito'])}", esperado, v["neto"])
    if abs(pagos - v["contado"]) > TOLERANCIA:
        detalle = (f"Efectivo {fmt(v['efectivo'])} | Cheque {fmt(v['cheque'])} | "
                   f"T.crédito {fmt(v['tcredito'])} | T.débito {fmt(v['tdebito'])} | Otros {fmt(v['otros'])}")
        add(R_PAGOS, detalle, v["contado"], pagos)
    if t["factura"] != "" and veces > 1:
        add(R_REPETIDA, f"Aparece {veces} veces", 1, veces)
    return t, salida


filas = []
severidades = []
observaciones = []
for i, fila in df.iterrows():
    t, hallazgos = evaluar(fila, int(repeticiones.loc[i]))
    if not hallazgos:
        severidades.append("OK")
    elif any(h[1] == "ERROR" for h in hallazgos):
        severidades.append("ERROR")
    else:
        severidades.append("SOSPECHOSO")
    observaciones.append("; ".join(h[0] for h in hallazgos) or "OK")
    for regla, sev, detalle, esperado, encontrado in hallazgos:
        numerico = isinstance(esperado, (int, float)) and isinstance(encontrado, (int, float))
        filas.append({
            "Severidad": sev,
            "Fila Excel": int(i) + 2,
            "Fact./Nota": t["factura"],
            "Razón social": t["razon"],
            "Cajero": t["cajero"] or "(sin cajero)",
            "Vendedor": t["vendedor"] or "(sin vendedor)",
            "Regla": regla,
            "Detalle": detalle,
            "Esperado": esperado,
            "Encontrado": encontrado,
            "Diferencia": encontrado - esperado if numerico else "",
        })

n = len(df)
marcadas = df.copy()
marcadas["Severidad"] = severidades
marcadas["Observaciones"] = observaciones
detalle = pd.DataFrame(filas, columns=["Severidad", "Fila Excel", "Fact./Nota", "Razón social",
                                       "Cajero", "Vendedor", "Regla", "Detalle",
                                       "Esperado", "Encontrado", "Diferencia"])

por_regla = pd.DataFrame(REGLAS, columns=["Regla", "Severidad"])
por_regla["Cantidad"] = por_regla["Regla"].map(detalle["Regla"].value_counts()).fillna(0).astype(int)
por_regla["% de filas"] = (por_regla["Cantidad"] * 100 / n).round(1)

n_err = severidades.count("ERROR")
n_sos = severidades.count("SOSPECHOSO")
general = pd.DataFrame({
    "Indicador": ["Filas revisadas", "Filas con ERROR", "Filas SOSPECHOSAS", "Filas OK"],
    "Cantidad": [n, n_err, n_sos, n - n_err - n_sos],
    "% de filas": [100.0, round(n_err * 100 / n, 1), round(n_sos * 100 / n, 1),
                   round((n - n_err - n_sos) * 100 / n, 1)],
})


def por_persona(columna, etiqueta):
    nombres = df[c[columna]].map(txt).replace("", f"(sin {etiqueta.lower()})")
    base = pd.DataFrame({etiqueta: nombres, "Con hallazgos": [s != "OK" for s in severidades]})
    r = base.groupby(etiqueta)["Con hallazgos"].agg(Facturas="count", ConHallazgos="sum").reset_index()
    r = r.rename(columns={"ConHallazgos": "Con hallazgos"})
    r["% con hallazgos"] = (r["Con hallazgos"] * 100 / r["Facturas"]).round(1)
    return r.sort_values("Con hallazgos", ascending=False).reset_index(drop=True)


por_cajero = por_persona("cajero", "Cajero")
por_vendedor = por_persona("vendedor", "Vendedor")

por_regla
