
import tempfile
from pathlib import Path
import pandas as pd
import streamlit as st

from motor import (
    procesar_archivos,
    validar_stock,
    validar_sabores,
    validar_data,
)

st.set_page_config(
    page_title="GridoPlanner",
    page_icon="🍦",
    layout="wide"
)

APP_DIR = Path(__file__).resolve().parent
DEFAULT_MAESTRO = APP_DIR / "maestro" / "Maestro_Productos_Grido.xlsx"

st.title("🍦 GridoPlanner")
st.caption("Herramienta experimental para generar sugerencias de pedido a partir de los exports de Grido.")

modo = st.sidebar.radio("Modo", ["Usuario", "Administrador"])

st.sidebar.markdown("---")
st.sidebar.info("Versión de prueba. Subí cada archivo en su campo correspondiente.")


def detectar_grupos_powerbi(file_path):
    """
    Lee de forma liviana el archivo data.xlsx y devuelve los grupos que el sistema intentará leer.
    Esto ayuda a detectar si el usuario no desplegó bien los + en Power BI antes de exportar.
    """
    try:
        df = pd.read_excel(file_path)
        if df.shape[1] < 13:
            return [], "El archivo tiene menos columnas de las esperadas."

        df.columns = [
            "Categoria", "SubCategoria", "Grupo", "Producto",
            "PL_Cant", "PL_Fact", "PL_Kilos", "PL_Porc",
            "Promo_Cant", "Promo_Fact", "Promo_Kilos", "Promo_Porc",
            "Total_Cantidad", "Total_Fact", "Total_Kilos", "Total_Porc"
        ]

        df = df.iloc[1:].copy()

        for c in ["Categoria", "SubCategoria", "Grupo"]:
            df[c] = df[c].ffill()

        mask = (
            ((df["Categoria"] == "Heladería") & (df["SubCategoria"] == "Impulsivos"))
            |
            ((df["Categoria"] == "Congelados") & (df["Grupo"].isin(["Congelados Multimarca", "Frizzio"])))
        )

        df = df[
            mask
            & df["Producto"].notna()
            & (~df["Producto"].astype(str).str.lower().eq("total"))
        ].copy()

        grupos = sorted([str(g) for g in df["Grupo"].dropna().unique()])
        return grupos, ""

    except Exception as e:
        return [], f"No pude leer los grupos del Power BI. Error: {e}"


def estado_archivo(label, ok, msg, detalles=None, warning=None):
    if ok:
        st.success(f"✅ {label} OK")
        if warning:
            st.warning(warning)
        if detalles:
            with st.expander(f"Ver detalles de {label}", expanded=False):
                if isinstance(detalles, list):
                    for item in detalles:
                        st.write(f"- {item}")
                else:
                    st.write(detalles)
    else:
        st.error(f"❌ {label}: {msg}")
        if detalles:
            with st.expander(f"Ver detalle del error", expanded=True):
                st.write(detalles)


if modo == "Administrador":
    st.header("Administrador")
    st.write("Desde acá podés descargar el maestro actual. En esta primera versión, para actualizarlo hay que reemplazar el archivo en GitHub o en la carpeta `maestro`.")

    if DEFAULT_MAESTRO.exists():
        st.download_button(
            "Descargar Maestro actual",
            data=DEFAULT_MAESTRO.read_bytes(),
            file_name="Maestro_Productos_Grido.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        st.error("No se encontró el maestro en la carpeta maestro/.")

    st.warning("La carga web de un nuevo maestro la vamos a agregar en una próxima versión para evitar sobrescribir por error.")
    st.stop()


st.header("Generar pedido")

st.markdown(
    """
    Subí cada archivo en su lugar.  
    El nombre del archivo puede variar, pero el contenido tiene que corresponder al tipo indicado.
    """
)

col1, col2, col3 = st.columns(3)

with col1:
    stock_file = st.file_uploader(
        "1. Archivo de STOCK",
        type=["csv"],
        help="Exportación de stock. Debe contener Grupo, Rubro, SubRubro, Item, Stock y Tránsito."
    )

with col2:
    sabores_file = st.file_uploader(
        "2. Ventas de SABORES",
        type=["xlsx"],
        help="Archivo Cajas por Sabor."
    )

with col3:
    data_file = st.file_uploader(
        "3. Ventas Power BI",
        type=["xlsx"],
        help="Exportación del Power BI, sección Mix de Ventas, con Total Cantidad."
    )

st.subheader("Configuración")

st.info(
    "Importante: antes de generar el pedido, verificá que los días analizados coincidan con el período exportado en Power BI. "
    "Si este dato está mal, el promedio semanal y la sugerencia de compra pueden salir mal."
)

cfg1, cfg2, cfg3 = st.columns(3)

with cfg1:
    semanas_objetivo = st.number_input("Semanas objetivo", min_value=0.5, max_value=12.0, value=4.0, step=0.5)
    st.caption("Cantidad de semanas de stock que querés tener disponibles cuando llegue el pedido.")

with cfg2:
    tiempo_reposicion = st.number_input("Tiempo de reposición", min_value=0.0, max_value=8.0, value=1.0, step=0.5)
    st.caption("Cantidad de semanas que tarda en llegar el pedido desde que lo hacés. Ejemplo: de jueves a jueves = 1.")

with cfg3:
    dias_analizados = st.number_input("Días analizados", min_value=1, max_value=60, value=14, step=1)
    st.caption("Debe coincidir con los días del filtro usado en Power BI, sección Mix de Ventas.")

st.markdown("---")
st.subheader("Estado de archivos")

if not DEFAULT_MAESTRO.exists():
    st.error("No se encontró el maestro de productos. Falta `maestro/Maestro_Productos_Grido.xlsx`.")
    st.stop()

ready = True

with tempfile.TemporaryDirectory() as tmp_dir:
    tmp = Path(tmp_dir)
    paths = {}

    if stock_file:
        p = tmp / "stock.csv"
        p.write_bytes(stock_file.getvalue())
        ok, msg = validar_stock(p)
        estado_archivo("Stock", ok, msg)
        ready = ready and ok
        paths["stock"] = p
    else:
        ready = False
        st.info("📄 Stock: pendiente")

    if sabores_file:
        p = tmp / "cajas_por_sabor.xlsx"
        p.write_bytes(sabores_file.getvalue())
        ok, msg = validar_sabores(p)
        estado_archivo("Sabores", ok, msg)
        ready = ready and ok
        paths["sabores"] = p
    else:
        ready = False
        st.info("📄 Sabores: pendiente")

    if data_file:
        p = tmp / "data.xlsx"
        p.write_bytes(data_file.getvalue())
        ok, msg = validar_data(p)

        grupos, grupo_error = detectar_grupos_powerbi(p)

        warning = (
            "Revisá que estén todos los grupos esperados. "
            "Si en Power BI no desplegaste el signo '+', algunos productos pueden no aparecer en la exportación."
        )

        detalles = []
        if grupos:
            detalles.append("Grupos detectados en la exportación:")
            detalles.extend(grupos)
        if grupo_error:
            detalles.append(grupo_error)

        estado_archivo(
            "Power BI",
            ok,
            msg,
            detalles=detalles if detalles else None,
            warning=warning if ok else None
        )

        ready = ready and ok
        paths["data"] = p
    else:
        ready = False
        st.info("📄 Power BI: pendiente")

    st.markdown("---")

    if not ready:
        st.info("Cuando los 3 archivos estén en OK, se habilita la generación del pedido.")
        st.stop()

    st.success("✅ Todo listo para generar el pedido.")

    if st.button("GENERAR PEDIDO", type="primary", use_container_width=True):
        output = tmp / "Pedido_Final.xlsx"

        try:
            result = procesar_archivos(
                stock_file=paths["stock"],
                sabores_file=paths["sabores"],
                data_file=paths["data"],
                maestro_file=DEFAULT_MAESTRO,
                output_file=output,
                overrides={
                    "Semanas objetivo": semanas_objetivo,
                    "Tiempo de reposición": tiempo_reposicion,
                    "Días analizados": dias_analizados,
                }
            )

            pedido = result["pedido"]
            sin_clasificar = result["sin_clasificar"]

            st.success("Pedido generado correctamente.")

            m1, m2, m3 = st.columns(3)
            m1.metric("Productos en pedido", len(pedido))
            m2.metric("Productos sin clasificar", len(sin_clasificar))
            m3.metric("Packs totales sugeridos", int(pedido["Packs a Comprar"].fillna(0).sum()))

            st.download_button(
                "DESCARGAR PEDIDO FINAL",
                data=output.read_bytes(),
                file_name="Pedido_Final.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

            with st.expander("Ver vista previa del pedido", expanded=False):
                st.dataframe(pedido.head(100), use_container_width=True)

            if len(sin_clasificar) > 0:
                with st.expander("Ver productos sin clasificar", expanded=True):
                    st.warning("Hay productos sin clasificar. Revisá esta hoja en el Excel y actualizá el Maestro.")
                    st.dataframe(sin_clasificar, use_container_width=True)

        except Exception as e:
            st.error("Ocurrió un error al generar el pedido.")
            st.exception(e)
