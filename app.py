
import tempfile
from pathlib import Path
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
st.sidebar.info("Versión inicial de prueba. Subí cada archivo en su campo correspondiente.")

if modo == "Administrador":
    st.header("Administrador")
    st.write("Desde acá podés descargar el maestro actual. En esta primera versión, para actualizarlo hay que reemplazar el archivo en el repositorio o en la carpeta `maestro`.")

    if DEFAULT_MAESTRO.exists():
        st.download_button(
            "Descargar Maestro actual",
            data=DEFAULT_MAESTRO.read_bytes(),
            file_name="Maestro_Productos_Grido.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        st.error("No se encontró el maestro en la carpeta maestro/.")

    st.warning("La carga web de un nuevo maestro la vamos a agregar en la próxima versión para evitar sobrescribir por error.")
    st.stop()

st.header("Generar pedido")

col1, col2, col3 = st.columns(3)

with col1:
    stock_file = st.file_uploader(
        "1. Subí el archivo de STOCK",
        type=["csv"],
        help="Exportación de stock. Debe tener columnas Grupo, Rubro, SubRubro, Item, Stock y Tránsito."
    )

with col2:
    sabores_file = st.file_uploader(
        "2. Subí CAJAS POR SABOR",
        type=["xlsx"],
        help="Archivo de ventas de sabores 7,800 kg."
    )

with col3:
    data_file = st.file_uploader(
        "3. Subí DATA Power BI",
        type=["xlsx"],
        help="Exportación Power BI con columna Total Cantidad."
    )

st.subheader("Configuración")
cfg1, cfg2, cfg3 = st.columns(3)

with cfg1:
    semanas_objetivo = st.number_input("Semanas objetivo", min_value=0.5, max_value=12.0, value=4.0, step=0.5)

with cfg2:
    tiempo_reposicion = st.number_input("Tiempo de reposición", min_value=0.0, max_value=8.0, value=1.0, step=0.5)

with cfg3:
    dias_analizados = st.number_input("Días analizados", min_value=1, max_value=60, value=14, step=1)

st.markdown("---")

if not DEFAULT_MAESTRO.exists():
    st.error("No se encontró el maestro de productos. Falta `maestro/Maestro_Productos_Grido.xlsx`.")
    st.stop()

ready = True

with tempfile.TemporaryDirectory() as tmp:
    tmp = Path(tmp)

    paths = {}

    if stock_file:
        p = tmp / "stock.csv"
        p.write_bytes(stock_file.getvalue())
        ok, msg = validar_stock(p)
        st.success(msg) if ok else st.error(msg)
        ready = ready and ok
        paths["stock"] = p
    else:
        ready = False

    if sabores_file:
        p = tmp / "cajas_por_sabor.xlsx"
        p.write_bytes(sabores_file.getvalue())
        ok, msg = validar_sabores(p)
        st.success(msg) if ok else st.error(msg)
        ready = ready and ok
        paths["sabores"] = p
    else:
        ready = False

    if data_file:
        p = tmp / "data.xlsx"
        p.write_bytes(data_file.getvalue())
        ok, msg = validar_data(p)
        st.success(msg) if ok else st.error(msg)
        ready = ready and ok
        paths["data"] = p
    else:
        ready = False

    if not ready:
        st.info("Subí los 3 archivos correctos para habilitar la generación del pedido.")
        st.stop()

    if st.button("Generar pedido", type="primary"):
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

            st.subheader("Resumen")
            m1, m2, m3 = st.columns(3)
            m1.metric("Productos en pedido", len(pedido))
            m2.metric("Productos sin clasificar", len(sin_clasificar))
            m3.metric("Packs totales sugeridos", int(pedido["Packs a Comprar"].fillna(0).sum()))

            st.subheader("Vista previa")
            st.dataframe(pedido.head(100), use_container_width=True)

            if len(sin_clasificar) > 0:
                st.warning("Hay productos sin clasificar. Revisá la hoja correspondiente en el Excel.")
                st.dataframe(sin_clasificar, use_container_width=True)

            st.download_button(
                "Descargar Pedido Final",
                data=output.read_bytes(),
                file_name="Pedido_Final.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error("Ocurrió un error al generar el pedido.")
            st.exception(e)
