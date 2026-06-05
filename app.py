import streamlit as st
import pandas as pd
import io
import zipfile
from datetime import datetime
from styles import BANNER_STYLE, FOOTER_STYLE

# --------------------------------------------------
# Configuracion
# --------------------------------------------------

st.set_page_config(
    page_title="Limpiador de Datos — CompuGamer",
    page_icon=None,
    layout="wide"
)

# --------------------------------------------------
# Transformaciones
# --------------------------------------------------

TRANSFORMACIONES = [
    {"columnName": "Categoria", "expression": "toTitlecase()"},
    {"columnName": "Marca",     "expression": "toTitlecase()"},
    {"columnName": "Precio",    "expression": "replace_currency"},
    {"columnName": "Precio",    "expression": "toNumber()"},
]

# --------------------------------------------------
# Funciones de limpieza
# --------------------------------------------------

def aplicar_titlecase(valor):
    if pd.isna(valor) or not isinstance(valor, str):
        return valor
    return valor.strip().title()


def limpiar_precio(valor):
    if pd.isna(valor):
        return None
    limpio = (
        str(valor)
        .strip()
        .replace("$", "")
        .replace("MXN", "")
        .replace(",", "")
        .strip()
    )
    try:
        return float(limpio)
    except ValueError:
        return None


def limpiar_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    for op in TRANSFORMACIONES:
        col = op["columnName"]
        expr = op["expression"]
        if col not in df.columns:
            continue
        if expr == "toTitlecase()":
            df[col] = df[col].apply(aplicar_titlecase)
        elif expr == "replace_currency":
            df[col] = df[col].apply(limpiar_precio)
        elif expr == "toNumber()":
            df[col] = pd.to_numeric(df[col], errors="coerce")
    return df


def contar_cambios(original: pd.DataFrame, limpio: pd.DataFrame) -> dict:
    cambios = {}
    for col in ["Categoria", "Marca"]:
        if col in original.columns and col in limpio.columns:
            n = (original[col].astype(str) != limpio[col].astype(str)).sum()
            cambios[col] = n
    return cambios


def procesar_archivo(archivo):
    try:
        df_original = pd.read_excel(archivo)
        df_limpio = limpiar_dataframe(df_original)
        stats = {
            "nombre": archivo.name,
            "filas": len(df_original),
            "columnas": len(df_original.columns),
            "cambios": contar_cambios(df_original, df_limpio),
            "exito": True,
            "error": None,
        }
        return df_limpio, stats
    except Exception as e:
        return None, {
            "nombre": archivo.name,
            "exito": False,
            "error": str(e),
        }


def df_a_excel_bytes(df: pd.DataFrame) -> bytes:
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Datos_Limpios")
    return buffer.getvalue()


def construir_zip(resultados: dict) -> bytes:
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for nombre, df in resultados.items():
            zf.writestr(f"limpio_{nombre}", df_a_excel_bytes(df))
    return buffer.getvalue()

# --------------------------------------------------
# Interfaz
# --------------------------------------------------

st.title("Limpiador Automatico de Datos — CompuGamer")
st.markdown("---")
st.markdown(BANNER_STYLE, unsafe_allow_html=True)

with st.expander("Que hace esta herramienta", expanded=False):
    st.markdown("""
    La herramienta aplica las siguientes transformaciones automaticamente:

    - **Categoria**: convierte a formato titulo (`electronica` → `Electronica`)
    - **Marca**: convierte a formato titulo (`logitech` → `Logitech`)
    - **Precio**: elimina `$` y `MXN`, convierte a numero (`$450` → `450.0`)

    Puedes subir hasta 10 archivos a la vez y descargarlos limpios de forma
    individual o agrupados en un ZIP.
    """)

st.markdown("---")

archivos = st.file_uploader(
    "Sube tus archivos Excel (multiples permitidos)",
    type=["xlsx", "xls"],
    accept_multiple_files=True,
)

if not archivos:
    st.info("Sube uno o varios archivos Excel para comenzar.")

    st.markdown("### Formato esperado")
    ejemplo = pd.DataFrame({
        "ID":        [1, 2, 3],
        "Producto":  ["Mouse RGB", "Teclado Gamer", 'Monitor 24"'],
        "Categoria": ["electronica", "ELECTRONICA", "Electronica"],
        "Marca":     ["logitech", "LOGITECH", "Samsung"],
        "Precio":    ["$450", "MXN 1200", "$3,500"],
    })
    st.dataframe(ejemplo, use_container_width=True)
    st.caption(
        "No importa si los datos estan en mayusculas, minusculas o incluyen $ / MXN; "
        "la herramienta los normaliza automaticamente."
    )

    with st.expander("Como usar multiples archivos", expanded=False):
        st.markdown("""
        1. Prepara tus archivos con la misma estructura de columnas.
        2. Selecciona varios archivos manteniendo Ctrl (Windows) o Cmd (Mac).
        3. Haz clic en **Limpiar todos los archivos**.
        4. Descarga los resultados individualmente o como ZIP.

        Bloques sugeridos para CompuGamer:
        - `CompuGamer_Tienda.xlsx` — ventas fisicas
        - `CompuGamer_Redes.xlsx` — redes sociales
        - `CompuGamer_Marketplace.xlsx` — marketplace
        - `CompuGamer_WhatsApp.xlsx` — ventas por WhatsApp
        """)

else:
    st.subheader(f"Archivos cargados: {len(archivos)}")

    resumen = [
        {
            "Nombre": f.name,
            "Tamano": f"{len(f.getvalue()) / 1024:.1f} KB",
            "Estado": "Pendiente",
        }
        for f in archivos
    ]
    st.dataframe(pd.DataFrame(resumen), use_container_width=True)
    st.markdown("---")

    if st.button("Limpiar todos los archivos", type="primary", use_container_width=True):

        barra = st.progress(0)
        estado = st.empty()

        resultados = {}
        stats_lista = []

        for i, archivo in enumerate(archivos):
            estado.text(f"Procesando: {archivo.name}…")
            df_limpio, stats = procesar_archivo(archivo)
            if df_limpio is not None:
                resultados[archivo.name] = df_limpio
            stats_lista.append(stats)
            barra.progress((i + 1) / len(archivos))

        estado.text("Todos los archivos procesados.")

        exitosos = sum(1 for s in stats_lista if s["exito"])
        st.success(f"Se procesaron {exitosos} de {len(stats_lista)} archivos correctamente.")

        st.subheader("Resumen de procesamiento")
        resumen_final = []
        for s in stats_lista:
            if s["exito"]:
                cambios_str = ", ".join(f"{k}: {v}" for k, v in s["cambios"].items())
                resumen_final.append({
                    "Archivo": s["nombre"],
                    "Estado":  "Exito",
                    "Filas":   s["filas"],
                    "Cambios": cambios_str or "Sin cambios detectados",
                })
            else:
                resumen_final.append({
                    "Archivo": s["nombre"],
                    "Estado":  "Error",
                    "Filas":   "-",
                    "Cambios": f"Error: {s['error'][:60]}",
                })
        st.dataframe(pd.DataFrame(resumen_final), use_container_width=True)

        # Descarga masiva
        if len(resultados) > 1:
            st.markdown("---")
            st.subheader("Descarga masiva")
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            st.download_button(
                label="Descargar todos como ZIP",
                data=construir_zip(resultados),
                file_name=f"datos_limpios_compugamer_{timestamp}.zip",
                mime="application/zip",
                use_container_width=True,
            )

        # Descargas individuales
        st.markdown("---")
        st.subheader("Descargas individuales")
        cols = st.columns(min(3, len(resultados)))
        for idx, (nombre, df) in enumerate(resultados.items()):
            with cols[idx % 3]:
                st.download_button(
                    label=nombre,
                    data=df_a_excel_bytes(df),
                    file_name=f"limpio_{nombre}",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

        # Vista previa
        if resultados:
            st.markdown("---")
            st.subheader("Vista previa — primer archivo limpio")
            primer_nombre = next(iter(resultados))
            st.caption(f"Mostrando primeras filas de: {primer_nombre}")
            st.dataframe(resultados[primer_nombre].head(10), use_container_width=True)

# --------------------------------------------------
# Pie de pagina
# --------------------------------------------------

st.markdown("---")
st.markdown(FOOTER_STYLE, unsafe_allow_html=True)