"""
app.py — Conversor Monday.com ➜ Browix (importación de Sub-tareas)
--------------------------------------------------------------------
Ejecutar con:   streamlit run app.py

Flujo:
  1. Subir el Excel exportado de Monday.
  2. Asignar a cada Orden de Compra (OC) su N° de tarea de Browix
     (manual o secuencial) y su transportista.
  3. Verificar la vista previa con el formato exacto de Browix.
  4. Descargar el Excel listo para importar.
"""

from datetime import date
from pathlib import Path

import pandas as pd
import streamlit as st

import transformador as tf

# Logo de la empresa (se guarda en la carpeta assets/ del repositorio)
LOGO = Path(__file__).parent / "assets" / "logo_grupo_canter.png"

st.set_page_config(
    page_title="Grupo Canter | Monday ➜ Browix",
    page_icon=str(LOGO) if LOGO.exists() else "🚚",
    layout="wide",
)


# ---------------------------------------------------------------------------
# Estado inicial: tabla de transportistas integrada
# ---------------------------------------------------------------------------
if "transportistas_df" not in st.session_state:
    st.session_state.transportistas_df = pd.DataFrame(
        {"NOMBRE": list(tf.TRANSPORTISTAS_DEFAULT.keys()),
         "DNI": list(tf.TRANSPORTISTAS_DEFAULT.values())}
    )
    st.session_state.version_transp = 0  # cambia cuando se recarga la tabla


def tabla_a_dict(df: pd.DataFrame) -> dict[str, int]:
    """Convierte la tabla editable NOMBRE/DNI en un diccionario limpio."""
    resultado = {}
    for _, fila in df.iterrows():
        nombre = str(fila.get("NOMBRE") or "").strip().upper()
        dni = tf.a_dni(fila.get("DNI"))
        if nombre and nombre != "NAN" and dni:
            resultado[nombre] = dni
    return resultado


# ---------------------------------------------------------------------------
# BARRA LATERAL: transportistas y opciones
# ---------------------------------------------------------------------------
with st.sidebar:
    st.header("🚚 Transportistas")
    st.caption("Tabla NOMBRE ➜ DNI usada para completar *Dni responsable*. "
               "Podés editarla o agregar filas acá mismo (fila vacía al final). "
               "Esta tabla **no** se incluye en el Excel exportado.")

    plantilla = st.file_uploader(
        "Cargar desde plantilla Browix (pestaña TRANSPORTISTAS)",
        type=["xlsx"], key="plantilla",
    )
    if plantilla is not None and st.session_state.get("plantilla_cargada") != plantilla.name:
        leidos = tf.leer_transportistas(plantilla)
        if leidos:
            st.session_state.transportistas_df = pd.DataFrame(
                {"NOMBRE": list(leidos.keys()), "DNI": list(leidos.values())}
            )
            st.session_state.version_transp += 1
            st.session_state.plantilla_cargada = plantilla.name
            st.success(f"Se cargaron {len(leidos)} transportistas.")
        else:
            st.warning("No se encontró la pestaña TRANSPORTISTAS (columnas NOMBRE y DNI).")

    tabla_editada = st.data_editor(
        st.session_state.transportistas_df,
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        key=f"editor_transp_{st.session_state.version_transp}",
        column_config={
            "NOMBRE": st.column_config.TextColumn("NOMBRE", required=True),
            "DNI": st.column_config.NumberColumn("DNI", format="%d", step=1),
        },
    )
    TRANSPORTISTAS = tabla_a_dict(tabla_editada)

    st.divider()
    st.header("⚙️ Opciones")
    separador = st.text_input("Separador en *Tarea* (Proveedor + Obra)", value=" - ")
    solo_faltante = st.checkbox("Solo ítems con FALTANTE mayor a 0", value=False)


# ---------------------------------------------------------------------------
# PASO 1: Subir archivo de Monday
# ---------------------------------------------------------------------------
if LOGO.exists():
    st.image(str(LOGO), width=240)
st.title("Conversor Monday.com ➜ Browix")
st.write("Transforma el export de Monday en la plantilla de importación de **Sub tareas** de Browix.")

st.subheader("1. Cargar el Excel de Monday")
archivo = st.file_uploader("Export de Monday (.xlsx)", type=["xlsx", "xls"], key="monday")

if archivo is None:
    st.info("Subí el archivo exportado de Monday para comenzar.")
    st.stop()

# Si el Excel tiene varias pestañas, se deja elegir cuál procesar.
try:
    hojas = tf.hojas_del_excel(archivo)
except Exception as error:  # archivo corrupto o formato inesperado
    st.error(f"No se pudo abrir el archivo: {error}")
    st.stop()
hoja = st.selectbox("Pestaña a procesar", hojas) if len(hojas) > 1 else hojas[0]

df_monday = tf.leer_monday(archivo, hoja)

with st.expander(f"Ver datos originales de Monday ({len(df_monday)} filas)"):
    st.dataframe(df_monday, use_container_width=True, hide_index=True)

# --- Opciones que dependen del archivo: fecha y estado -----------------------
col_a, col_b = st.columns(2)
with col_a:
    # La Fecha de Browix se toma SIEMPRE de "Fecha estimada Entrega S/Proveedor".
    columna_fecha = tf.buscar_columna(df_monday, "fecha_entrega")
    fecha_fija = None
    if columna_fecha:
        st.info(f"📅 *Fecha* tomada de la columna **{columna_fecha}** (formato dd/mm/aaaa).")
    else:
        st.error("No se encontró la columna **Fecha estimada Entrega S/Proveedor** "
                 "en el Excel de Monday. Verificá el export.")
        st.stop()

with col_b:
    col_estado = tf.buscar_columna(df_monday, "estado")
    if col_estado:
        estados = sorted(df_monday[col_estado].dropna().astype(str).str.strip().unique())
        estados_elegidos = st.multiselect("Filtrar por *Estado* (vacío = todos)", estados)
        if estados_elegidos:
            df_monday = df_monday[df_monday[col_estado].astype(str).str.strip().isin(estados_elegidos)]

# --- Transformación base ------------------------------------------------------
try:
    items = tf.preparar_items(df_monday, TRANSPORTISTAS, separador, columna_fecha, fecha_fija)
except ValueError as error:
    st.error(str(error))
    st.stop()

if solo_faltante:
    faltante_num = pd.to_numeric(items["Detalle Subtarea"], errors="coerce")
    items = items[faltante_num > 0].reset_index(drop=True)

if items.empty:
    st.warning("No quedaron ítems para procesar con los filtros actuales.")
    st.stop()


# ---------------------------------------------------------------------------
# PASO 2: Asignar transportista por OC
# ---------------------------------------------------------------------------
st.subheader("2. Asignar N° de tarea y transportista por Orden de Compra")
st.caption(
    "Cada OC es **una tarea** en Browix y sus ítems son las sub-tareas. "
    "El **N° de tarea** (va en *Agrupador Tarea*) debe continuar la numeración "
    "de la última tarea cargada en Browix; el **N° de OC** va en *Observaciones*. "
    "El transportista se sugiere cuando Monday ya trae el DNI o el nombre."
)

nombres = list(TRANSPORTISTAS.keys())
asignacion = tf.asignaciones_por_oc(items)
asignacion["Transportista"] = asignacion["Transportista"].where(asignacion["Transportista"] != "", None)

# --- Numeración de tareas -----------------------------------------------------
c1, c2, c3 = st.columns([1.2, 1.3, 1.5])
with c1:
    numero_inicial = st.number_input(
        "N° de tarea para la primera OC",
        min_value=1, step=1, value=None, placeholder="Ej.: 1250",
        help="Número de tarea de Browix que corresponde a la primera OC de la lista.",
    )
with c2:
    secuencial = st.toggle(
        "Numerar las siguientes OC en forma secuencial",
        value=True,
        help="Si está activo, cada OC siguiente recibe el número anterior + 1. "
             "Igual podés corregir cualquier número a mano en la tabla.",
    )
with c3:
    para_todas = st.selectbox(
        "Asignar transportista a todas las OC sin transportista",
        ["(no asignar)"] + nombres,
    )

# Precarga de la columna "N° Tarea" según las opciones elegidas
asignacion.insert(1, "N° Tarea", pd.array([None] * len(asignacion), dtype="Int64"))
if numero_inicial is not None:
    if secuencial:
        numeros = tf.numerar_secuencial(asignacion["OC"], int(numero_inicial))
        asignacion["N° Tarea"] = pd.array(asignacion["OC"].map(numeros), dtype="Int64")
    else:
        asignacion.loc[0, "N° Tarea"] = int(numero_inicial)  # solo la primera OC

if para_todas != "(no asignar)":
    asignacion["Transportista"] = asignacion["Transportista"].fillna(para_todas)

# La clave depende del archivo y de las opciones: si cambian, la tabla se recalcula.
clave_editor = (f"oc_{archivo.name}_{archivo.size}_{hoja}_{len(asignacion)}_"
                f"{numero_inicial}_{secuencial}_{para_todas}")
asignacion_editada = st.data_editor(
    asignacion,
    hide_index=True,
    use_container_width=True,
    key=clave_editor,
    disabled=["OC", "Tarea", "Items"],
    column_order=["OC", "N° Tarea", "Tarea", "Items", "Transportista"],
    column_config={
        "OC": st.column_config.Column("N° OC (➜ Observaciones)"),
        "N° Tarea": st.column_config.NumberColumn(
            "N° Tarea (➜ Agrupador)", min_value=1, step=1, format="%d",
            help="Editable: número de tarea de Browix para esta OC."),
        "Items": st.column_config.NumberColumn("Ítems", format="%d"),
        "Transportista": st.column_config.SelectboxColumn(
            "Transportista", options=nombres, required=False),
    },
)
mapa_oc = {
    oc: ("" if pd.isna(t) else t)
    for oc, t in zip(asignacion_editada["OC"], asignacion_editada["Transportista"])
}
mapa_tareas = dict(zip(asignacion_editada["OC"], asignacion_editada["N° Tarea"]))

# Controles sobre la numeración
ocs_sin_numero = [oc for oc, n in mapa_tareas.items() if tf.a_numero_tarea(n) is None]
numeros_validos = [tf.a_numero_tarea(n) for n in mapa_tareas.values() if tf.a_numero_tarea(n)]
repetidos = sorted({n for n in numeros_validos if numeros_validos.count(n) > 1})


# ---------------------------------------------------------------------------
# PASO 3: Vista previa con formato Browix
# ---------------------------------------------------------------------------
st.subheader("3. Vista previa (formato de importación Browix)")
resultado = tf.construir_browix(items, TRANSPORTISTAS, mapa_oc, mapa_tareas)

sin_dni = int(resultado["Dni responsable"].isna().sum())
sin_fecha = int((resultado["Fecha"] == "").sum())
m1, m2, m3, m4, m5 = st.columns(5)
m1.metric("Sub-tareas", len(resultado))
m2.metric("Órdenes de compra", resultado["Observaciones"].nunique())
m3.metric("OC sin N° de tarea", len(ocs_sin_numero))
m4.metric("Sin DNI responsable", sin_dni)
m5.metric("Sin fecha", sin_fecha)

if ocs_sin_numero:
    st.error(f"{len(ocs_sin_numero)} OC sin N° de tarea: "
             + ", ".join(map(str, ocs_sin_numero))
             + ". Asignalo en el paso 2 para poder descargar.")
if repetidos:
    st.warning("Hay N° de tarea repetidos en distintas OC: " + ", ".join(map(str, repetidos))
               + ". Esas OC quedarían agrupadas en una misma tarea de Browix.")
if sin_dni:
    st.warning(f"{sin_dni} fila(s) no tienen *Dni responsable*. Asigná un transportista en el paso 2.")
if sin_fecha:
    st.warning(f"{sin_fecha} fila(s) no tienen fecha en la columna elegida.")

st.dataframe(
    resultado,
    use_container_width=True,
    hide_index=True,
    column_config={
        "Agrupador Tarea": st.column_config.NumberColumn(format="%d"),
        "Dni responsable": st.column_config.NumberColumn(format="%d"),
        "Observaciones": st.column_config.Column("Observaciones (N° OC)"),
    },
)


# ---------------------------------------------------------------------------
# PASO 4: Descargar
# ---------------------------------------------------------------------------
st.subheader("4. Descargar")
nombre_salida = f"Importacion_Browix_{date.today():%Y-%m-%d}.xlsx"
st.download_button(
    "⬇️ Descargar Excel para Browix",
    data=tf.exportar_excel(resultado),
    file_name=nombre_salida,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary",
    use_container_width=True,
    # Sin N° de tarea no hay agrupador: no se permite descargar.
    disabled=bool(ocs_sin_numero),
)
if ocs_sin_numero:
    st.caption("⚠️ La descarga se habilita cuando todas las OC tienen N° de tarea.")
if sin_dni:
    st.caption("⚠️ Se puede descargar igual, pero Browix podría rechazar filas sin DNI.")
