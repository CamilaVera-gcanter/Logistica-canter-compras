"""
app.py — Conversor Monday.com ➜ Browix (importación de Sub-tareas)
--------------------------------------------------------------------
Ejecutar con:   streamlit run app.py

Flujo:
  1. Subir el Excel exportado de Monday.
  2. Revisar / asignar el transportista de cada Orden de Compra (OC).
  3. Verificar la vista previa con el formato exacto de Browix.
  4. Descargar el Excel listo para importar.
"""

from datetime import date

import pandas as pd
import streamlit as st

import transformador as tf

st.set_page_config(page_title="Monday ➜ Browix", page_icon="🚚", layout="wide")


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
               "Podés editarla o agregar filas acá mismo.")

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
    modo_fecha = st.radio("Origen de la columna *Fecha*",
                          ["Columna del Excel", "Fecha fija para todas"], horizontal=True)
    columna_fecha, fecha_fija = None, None
    if modo_fecha == "Columna del Excel":
        opciones_fecha = tf.columnas_de_fecha(df_monday)
        if opciones_fecha:
            por_defecto = tf.fecha_por_defecto(df_monday)
            columna_fecha = st.selectbox("Columna de fecha", opciones_fecha,
                                         index=opciones_fecha.index(por_defecto))
        else:
            st.warning("El archivo no tiene columnas de fecha; usá una fecha fija.")
    else:
        fecha_fija = st.date_input("Fecha", value=date.today(), format="DD/MM/YYYY")

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
st.subheader("2. Asignar transportista (DNI responsable)")
st.caption("Se sugiere automáticamente cuando Monday ya trae el DNI o el nombre "
           "(columnas DNI / NOMBRE APELLIDO / LOGISTICA). Elegí el transportista "
           "de cada OC; se aplica a todos sus ítems.")

nombres = list(TRANSPORTISTAS.keys())
asignacion = tf.asignaciones_por_oc(items)
asignacion["Transportista"] = asignacion["Transportista"].where(asignacion["Transportista"] != "", None)

para_todas = st.selectbox(
    "Asignar rápidamente a todas las OC sin transportista",
    ["(no asignar)"] + nombres,
)
if para_todas != "(no asignar)":
    asignacion["Transportista"] = asignacion["Transportista"].fillna(para_todas)

# La clave depende del archivo y de las opciones para reiniciar el editor si cambian.
clave_editor = f"oc_{archivo.name}_{archivo.size}_{hoja}_{para_todas}_{len(asignacion)}"
asignacion_editada = st.data_editor(
    asignacion,
    hide_index=True,
    use_container_width=True,
    key=clave_editor,
    disabled=["Agrupador Tarea", "Tarea", "Items"],
    column_config={
        "Agrupador Tarea": st.column_config.Column("N° OC"),
        "Items": st.column_config.NumberColumn("Ítems", format="%d"),
        "Transportista": st.column_config.SelectboxColumn(
            "Transportista", options=nombres, required=False),
    },
)
mapa_oc = {
    oc: ("" if pd.isna(t) else t)
    for oc, t in zip(asignacion_editada["Agrupador Tarea"], asignacion_editada["Transportista"])
}


# ---------------------------------------------------------------------------
# PASO 3: Vista previa con formato Browix
# ---------------------------------------------------------------------------
st.subheader("3. Vista previa (formato de importación Browix)")
resultado = tf.construir_browix(items, TRANSPORTISTAS, mapa_oc)

sin_dni = int(resultado["Dni responsable"].isna().sum())
sin_fecha = int((resultado["Fecha"] == "").sum())
m1, m2, m3, m4 = st.columns(4)
m1.metric("Sub-tareas", len(resultado))
m2.metric("Órdenes de compra", resultado["Agrupador Tarea"].nunique())
m3.metric("Sin DNI responsable", sin_dni)
m4.metric("Sin fecha", sin_fecha)

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
    },
)


# ---------------------------------------------------------------------------
# PASO 4: Descargar
# ---------------------------------------------------------------------------
st.subheader("4. Descargar")
nombre_salida = f"Importacion_Browix_{date.today():%Y-%m-%d}.xlsx"
st.download_button(
    "⬇️ Descargar Excel para Browix",
    data=tf.exportar_excel(resultado, TRANSPORTISTAS),
    file_name=nombre_salida,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary",
    use_container_width=True,
)
if sin_dni:
    st.caption("⚠️ Se puede descargar igual, pero Browix podría rechazar filas sin DNI.")
