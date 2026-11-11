"""
transformador.py
----------------
Lógica pura (sin interfaz) para convertir el export de Monday.com
al formato de importación de sub-tareas de Browix.

Se mantiene separada de la interfaz (app.py) para poder probarla
y reutilizarla fácilmente (por ejemplo, desde un script o un cron).
"""

from __future__ import annotations

import io
import unicodedata
from datetime import date, datetime, timedelta

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# 1. CONFIGURACIÓN
# ---------------------------------------------------------------------------

# Columnas EXACTAS de la plantilla de Browix (hoja "Sub tarea").
# Ojo: "Detalle " lleva un espacio al final, tal cual figura en la plantilla.
COLUMNAS_BROWIX = [
    "Tarea",
    "Agrupador Tarea",
    "Fecha",
    "Dni responsable",
    "Subtarea",
    "Detalle Subtarea",
    "Detalle ",
    "Observaciones",
]
HOJA_BROWIX = "Sub tarea"

# Tabla de transportistas integrada (copiada de la pestaña "TRANSPORTISTAS").
# Se puede ampliar desde la app o subiendo la plantilla de Browix.
TRANSPORTISTAS_DEFAULT = {
    "DAMIAN FARIAS": 26817879,
    "CHRISTIAN LEDESMA": 30473183,
}

# Nombres posibles de cada columna del export de Monday.
# Se comparan "normalizados" (minúsculas, sin acentos ni espacios extra),
# así la app sigue funcionando si Monday cambia levemente un encabezado.
ALIAS_MONDAY = {
    "obra": ["name", "obra", "nombre"],
    "proveedor": ["proveedor elegido", "proveedor"],
    "oc": ["numero oc", "nro oc", "n oc", "n° oc", "nº oc", "orden de compra", "oc"],
    "item": ["item / producto", "item/producto", "item", "producto", "descripcion"],
    "faltante": ["faltante"],
    "estado": ["estado"],
    # Columna de la que SIEMPRE se toma la "Fecha" para Browix.
    # Monday a veces corta el encabezado ("...S/Proveedo"), por eso van ambas variantes.
    "fecha_entrega": [
        "fecha estimada entrega s/proveedor",
        "fecha estimada entrega s/proveedo",
        "fecha estimada entrega s/ proveedor",
        "fecha estimada entrega",
    ],
    # Columnas donde puede venir escrita la persona / DNI asignado
    "dni": ["dni", "dni responsable"],
    "persona": ["nombre apellido", "logistica", "transportista", "responsable"],
}

# Columnas de fecha candidatas, en orden de preferencia.
FECHAS_CANDIDATAS = [
    "fecha tarea browix",
    "fecha estimada entrega s/proveedo",
    "fecha estimada entrega s/proveedor",
    "fecha",
    "fecha transito",
    "fecha carga estado entrega logistica",
]

FORMATO_FECHA = "%d/%m/%Y"  # La columna Fecha de la plantilla es de tipo TEXTO.


# ---------------------------------------------------------------------------
# 2. UTILIDADES
# ---------------------------------------------------------------------------

def normalizar(texto) -> str:
    """Minúsculas, sin acentos y con espacios simples. Sirve para comparar."""
    if texto is None or (isinstance(texto, float) and pd.isna(texto)):
        return ""
    texto = unicodedata.normalize("NFKD", str(texto))
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    return " ".join(texto.lower().split())


def buscar_columna(df: pd.DataFrame, clave: str) -> str | None:
    """Devuelve el nombre real de la columna del df que corresponde a `clave`."""
    reales = {normalizar(c): c for c in df.columns}
    for alias in ALIAS_MONDAY[clave]:
        if normalizar(alias) in reales:
            return reales[normalizar(alias)]
    return None


def columnas_de_fecha(df: pd.DataFrame) -> list[str]:
    """Lista las columnas que parecen fechas (las candidatas primero)."""
    reales = {normalizar(c): c for c in df.columns}
    orden = [reales[n] for n in FECHAS_CANDIDATAS if n in reales]
    otras = [c for c in df.columns if "fecha" in normalizar(c) and c not in orden]
    return orden + otras


def fecha_por_defecto(df: pd.DataFrame) -> str | None:
    """Primera columna de fecha candidata que tenga al menos un valor."""
    for col in columnas_de_fecha(df):
        if df[col].notna().any() and (df[col].astype(str).str.strip() != "").any():
            return col
    cols = columnas_de_fecha(df)
    return cols[0] if cols else None


def a_fecha(valor) -> date | None:
    """Convierte cualquier formato de fecha (datetime, serial de Excel, texto)."""
    if valor is None or (isinstance(valor, float) and pd.isna(valor)) or valor is pd.NaT:
        return None
    if isinstance(valor, (datetime, pd.Timestamp)):
        return valor.date()
    if isinstance(valor, date):
        return valor
    if isinstance(valor, (int, float)):
        # Número de serie de Excel (días desde 30/12/1899)
        if 20000 < valor < 80000:
            return (datetime(1899, 12, 30) + timedelta(days=float(valor))).date()
        return None
    texto = str(valor).strip()
    if not texto:
        return None
    if texto.replace(".", "", 1).isdigit():
        return a_fecha(float(texto))
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%Y/%m/%d", "%d/%m/%y"):
        try:
            return datetime.strptime(texto[:10], fmt).date()
        except ValueError:
            continue
    try:
        return pd.to_datetime(texto, dayfirst=True).date()
    except (ValueError, TypeError):
        return None


def a_entero_o_texto(valor):
    """14.0 -> 14 ; '15028' -> 15028 ; texto se deja como está ; vacío -> None."""
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return None
    if isinstance(valor, (int, float)):
        return int(valor) if float(valor).is_integer() else valor
    texto = str(valor).strip()
    if not texto:
        return None
    try:
        num = float(texto.replace(",", "."))
        return int(num) if num.is_integer() else num
    except ValueError:
        return texto


def a_dni(valor) -> int | None:
    """Limpia un DNI ('26.817.879', 26817879.0, ...) y lo devuelve como entero."""
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return None
    if isinstance(valor, float):
        valor = int(valor)  # 26817879.0 -> 26817879
    digitos = "".join(ch for ch in str(valor) if ch.isdigit())
    return int(digitos) if 6 <= len(digitos) <= 9 else None


# ---------------------------------------------------------------------------
# 3. LECTURA DE ARCHIVOS
# ---------------------------------------------------------------------------

def leer_monday(archivo, hoja: str | int = 0) -> pd.DataFrame:
    """
    Lee el export de Monday. Detecta la fila de encabezados aunque el export
    traiga filas de título arriba (nombre del tablero, grupo, etc.).
    """
    crudo = pd.read_excel(archivo, sheet_name=hoja, header=None, dtype=object)

    # Busca, en las primeras 15 filas, la que tenga "Name" y alguna columna de OC.
    fila_encabezado = 0
    for i in range(min(15, len(crudo))):
        valores = {normalizar(v) for v in crudo.iloc[i].tolist()}
        if "name" in valores and any(normalizar(a) in valores for a in ALIAS_MONDAY["oc"]):
            fila_encabezado = i
            break

    encabezados = [
        str(v).strip() if pd.notna(v) else f"Columna_{j + 1}"
        for j, v in enumerate(crudo.iloc[fila_encabezado].tolist())
    ]
    df = crudo.iloc[fila_encabezado + 1:].copy()
    df.columns = encabezados
    return df.reset_index(drop=True)


def hojas_del_excel(archivo) -> list[str]:
    return pd.ExcelFile(archivo).sheet_names


def leer_transportistas(archivo) -> dict[str, int]:
    """Lee la pestaña TRANSPORTISTAS (columnas NOMBRE / DNI) de un Excel."""
    xls = pd.ExcelFile(archivo)
    hoja = next((h for h in xls.sheet_names if normalizar(h) == "transportistas"), None)
    if hoja is None:
        return {}
    df = pd.read_excel(xls, sheet_name=hoja, dtype=object)
    df.columns = [normalizar(c) for c in df.columns]
    if "nombre" not in df.columns or "dni" not in df.columns:
        return {}
    resultado = {}
    for _, fila in df.iterrows():
        nombre = str(fila["nombre"]).strip() if pd.notna(fila["nombre"]) else ""
        dni = a_dni(fila["dni"])
        if nombre and dni:
            resultado[nombre.upper()] = dni
    return resultado


# ---------------------------------------------------------------------------
# 4. LIMPIEZA Y CRUCE CON TRANSPORTISTAS
# ---------------------------------------------------------------------------

def limpiar_monday(df: pd.DataFrame) -> pd.DataFrame:
    """
    Quita filas que no son ítems reales:
    - fila de totales que Monday agrega al final (sin Name ni Item),
    - encabezados repetidos y títulos de grupo.
    """
    col_obra = buscar_columna(df, "obra")
    col_item = buscar_columna(df, "item")
    df = df.copy()

    def vacio(serie):
        return serie.isna() | (serie.astype(str).str.strip() == "")

    mascara = ~vacio(df[col_obra]) & ~vacio(df[col_item])
    mascara &= df[col_obra].astype(str).map(normalizar) != "name"  # encabezado repetido
    return df[mascara].reset_index(drop=True)


def buscar_transportista(texto, transportistas: dict[str, int]) -> str | None:
    """
    Busca un transportista dentro de un texto libre de Monday.
    Ej.: 'Dami Gómez' -> 'DAMI'. Prioriza coincidencias exactas.
    """
    t = normalizar(texto)
    if not t:
        return None
    for nombre in transportistas:
        if normalizar(nombre) == t:
            return nombre
    palabras = set(t.replace(",", " ").replace("-", " ").split())
    for nombre in transportistas:
        if normalizar(nombre) in palabras:
            return nombre
    return None


def transportista_sugerido(fila: pd.Series, df_cols: dict, transportistas: dict[str, int]) -> str:
    """
    Sugiere el transportista de una fila de Monday:
    1) si la columna DNI ya trae un DNI conocido, usa ese transportista;
    2) si alguna columna de persona (NOMBRE APELLIDO, LOGISTICA...) lo nombra.
    """
    if df_cols.get("dni"):
        dni = a_dni(fila[df_cols["dni"]])
        for nombre, d in transportistas.items():
            if dni and d == dni:
                return nombre
    for col in df_cols.get("persona", []):
        encontrado = buscar_transportista(fila[col], transportistas)
        if encontrado:
            return encontrado
    return ""


# ---------------------------------------------------------------------------
# 5. TRANSFORMACIÓN PRINCIPAL
# ---------------------------------------------------------------------------

def preparar_items(
    df_monday: pd.DataFrame,
    transportistas: dict[str, int],
    separador_tarea: str = " - ",
    columna_fecha: str | None = None,
    fecha_fija: date | None = None,
) -> pd.DataFrame:
    """
    Arma una tabla intermedia (una fila por ítem) con todo lo necesario para
    Browix + la columna 'Transportista', que el usuario puede corregir en la app.
    """
    df = limpiar_monday(df_monday)

    cols = {k: buscar_columna(df, k)
            for k in ("obra", "proveedor", "oc", "item", "faltante", "dni", "fecha_entrega")}

    # La Fecha sale SIEMPRE de "Fecha estimada Entrega S/Proveedor",
    # salvo que se pase explícitamente otra columna o una fecha fija.
    if columna_fecha is None and fecha_fija is None:
        columna_fecha = cols["fecha_entrega"]

    obligatorias = ["obra", "proveedor", "oc", "item"]
    if fecha_fija is None and columna_fecha is None:
        obligatorias.append("fecha_entrega")
    faltan = [k for k in obligatorias if cols[k] is None]
    if faltan:
        raise ValueError(
            "No se encontraron estas columnas en el Excel de Monday: "
            + ", ".join(faltan)
            + ". Columnas disponibles: " + ", ".join(map(str, df.columns))
        )
    reales = {normalizar(c): c for c in df.columns}
    cols["persona"] = [reales[normalizar(a)] for a in ALIAS_MONDAY["persona"] if normalizar(a) in reales]

    def texto(v):
        return "" if v is None or (isinstance(v, float) and pd.isna(v)) else str(v).strip()

    filas = []
    for _, f in df.iterrows():
        proveedor, obra = texto(f[cols["proveedor"]]), texto(f[cols["obra"]])
        # Tarea = PROVEEDOR ELEGIDO + NAME (obra)
        tarea = separador_tarea.join(p for p in (proveedor, obra) if p)

        if fecha_fija is not None:
            fecha = fecha_fija
        elif columna_fecha:
            fecha = a_fecha(f[columna_fecha])
        else:
            fecha = None

        filas.append({
            "Tarea": tarea,
            # El N° de OC NO es el agrupador de Browix: se guarda aparte y
            # termina en la columna "Observaciones" (ver construir_browix).
            "OC": a_entero_o_texto(f[cols["oc"]]),
            "Fecha": fecha.strftime(FORMATO_FECHA) if fecha else "",
            "Transportista": transportista_sugerido(f, cols, transportistas),
            "Subtarea": texto(f[cols["item"]]),
            "Detalle Subtarea": a_entero_o_texto(f[cols["faltante"]]) if cols["faltante"] else None,
        })
    return pd.DataFrame(filas)


def asignaciones_por_oc(items: pd.DataFrame) -> pd.DataFrame:
    """
    Una fila por OC (en el orden en que aparecen en Monday) con:
    Tarea, cantidad de ítems y transportista sugerido (el más frecuente).
    """
    def primero_no_vacio(s):
        s = [x for x in s if x]
        return max(set(s), key=s.count) if s else ""

    return (
        items.groupby("OC", sort=False, dropna=False)
        .agg(Tarea=("Tarea", "first"),
             Items=("Subtarea", "count"),
             Transportista=("Transportista", primero_no_vacio))
        .reset_index()
    )


def numerar_secuencial(ocs, inicio: int) -> dict:
    """
    Asigna números de tarea consecutivos a cada OC, empezando en `inicio`.
    Ej.: ocs=[15028, 14751, 14736], inicio=500 -> {15028: 500, 14751: 501, 14736: 502}
    """
    return {oc: int(inicio) + i for i, oc in enumerate(ocs)}


def a_numero_tarea(valor) -> int | None:
    """Limpia el N° de tarea ingresado por el usuario (vacío o inválido -> None)."""
    if valor is None or valor is pd.NA or (isinstance(valor, float) and pd.isna(valor)):
        return None
    try:
        numero = float(valor)
    except (TypeError, ValueError):
        return None
    return int(numero) if numero.is_integer() and numero > 0 else None


def construir_browix(
    items: pd.DataFrame,
    transportistas: dict[str, int],
    asignacion_oc: dict | None = None,
    numeros_tarea: dict | None = None,
) -> pd.DataFrame:
    """
    Devuelve el DataFrame final con las columnas exactas de Browix.

    - `asignacion_oc` = {numero_oc: nombre_transportista} elegido en la app;
      tiene prioridad sobre lo sugerido automáticamente.
    - `numeros_tarea` = {numero_oc: numero_de_tarea_browix}. Ese número va en
      "Agrupador Tarea": todas las sub-tareas de una misma OC quedan en la
      misma tarea de Browix.
    - El N° de OC va en "Observaciones".
    """
    salida = items.copy()
    if asignacion_oc:
        salida["Transportista"] = [
            asignacion_oc.get(oc) or t
            for oc, t in zip(salida["OC"], salida["Transportista"])
        ]
    numeros_tarea = numeros_tarea or {}
    salida["Agrupador Tarea"] = salida["OC"].map(
        lambda oc: a_numero_tarea(numeros_tarea.get(oc))
    ).astype("Int64")
    salida["Dni responsable"] = salida["Transportista"].map(
        lambda n: transportistas.get(n) if n else None
    ).astype("Int64")
    salida["Detalle "] = ""
    salida["Observaciones"] = salida["OC"]
    return salida[COLUMNAS_BROWIX]


# ---------------------------------------------------------------------------
# 6. EXPORTACIÓN A EXCEL
# ---------------------------------------------------------------------------

def exportar_excel(df_browix: pd.DataFrame) -> bytes:
    """
    Genera el .xlsx listo para importar en Browix: una única hoja "Sub tarea".
    La tabla de transportistas NO se exporta (solo vive en la app).
    """
    wb = Workbook()
    ws = wb.active
    ws.title = HOJA_BROWIX

    ws.append(COLUMNAS_BROWIX)
    for celda in ws[1]:
        celda.font = Font(bold=True)
        celda.fill = PatternFill("solid", fgColor="D9E1F2")

    for fila in df_browix.itertuples(index=False):
        # Celdas vacías (NaN / NA / "") se dejan realmente vacías en el Excel.
        valores = [None if (v is None or v is pd.NA or v == "" or (isinstance(v, float) and pd.isna(v)))
                   else (int(v) if hasattr(v, "item") and float(v).is_integer() else v)
                   for v in fila]
        ws.append(valores)

    # La columna Fecha (C) de la plantilla está formateada como TEXTO ("@").
    for (celda,) in ws.iter_rows(min_row=1, min_col=3, max_col=3):
        celda.number_format = "@"

    anchos = {"A": 45, "B": 16, "C": 12, "D": 16, "E": 60, "F": 16, "G": 14, "H": 18}
    for col, ancho in anchos.items():
        ws.column_dimensions[col].width = ancho
    ws.freeze_panes = "A2"

    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()
