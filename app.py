
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

APP_VERSION = "Mark XVIII (v1.7-beta)"
ADMIN_PIN = "2468"  # Cambiar este PIN si querés otro.

APP_DIR = Path(__file__).resolve().parent
DEFAULT_MAESTRO = APP_DIR / "maestro" / "Maestro_Productos_Grido.xlsx"
DEFAULT_CARRITO = APP_DIR / "datos" / "Modelo_de_Carrito.xlsx"


def mostrar_screenshot(nombre_archivo, caption):
    img_path = APP_DIR / "assets" / "screenshots" / nombre_archivo
    if img_path.exists():
        st.image(str(img_path), caption=caption, use_container_width=True)
    else:
        st.error(f"No se encontró la imagen: {nombre_archivo}")


def formato_moneda_ar(valor):
    try:
        valor = int(float(valor))
    except Exception:
        valor = 0

    signo = "-" if valor < 0 else ""
    entero_txt = f"{abs(valor):,}".replace(",", ".")
    return f"$ {signo}{entero_txt}"


st.title("🍦 GridoPlanner")
st.caption(f"Herramienta experimental para generar sugerencias de pedido a partir de los exports de Grido. Versión {APP_VERSION}")

modo = st.sidebar.radio("Modo", ["Usuario", "Administrador"])

st.sidebar.markdown("---")
st.sidebar.info(f"Versión {APP_VERSION}. Subí cada archivo en su campo correspondiente.")


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



def _norm_col(c):
    return str(c).strip().lower().replace("á", "a").replace("é", "e").replace("í", "i").replace("ó", "o").replace("ú", "u")


def validar_maestro(maestro_path):
    """
    Valida que el maestro subido tenga la estructura mínima esperada.
    No valida que cada producto sea correcto, solo que el archivo sea usable.
    """
    errores = []
    advertencias = []

    try:
        xls = pd.ExcelFile(maestro_path)
    except Exception as e:
        return False, [f"No pude abrir el Excel del maestro. Error: {e}"], []

    hojas_requeridas = ["Productos", "Aliases", "Exclusiones", "Configuración"]
    hojas = set(xls.sheet_names)

    for hoja in hojas_requeridas:
        if hoja not in hojas:
            errores.append(f"Falta la hoja obligatoria: {hoja}")

    if errores:
        return False, errores, advertencias

    try:
        productos = pd.read_excel(maestro_path, sheet_name="Productos")
        cols = {_norm_col(c): c for c in productos.columns}

        requeridas_productos = [
            "producto base",
            "codigo compra",
            "producto compra",
            "compra minima",
        ]

        for req in requeridas_productos:
            if req not in cols:
                errores.append(f"Hoja Productos: falta columna '{req}'.")

        if "producto base" in cols:
            vacios = productos[cols["producto base"]].isna().sum()
            if vacios > 0:
                advertencias.append(f"Hoja Productos: hay {vacios} filas sin Producto Base.")

        if "compra minima" in cols:
            compra_min = pd.to_numeric(productos[cols["compra minima"]], errors="coerce")
            invalidos = compra_min.isna().sum()
            if invalidos > 0:
                advertencias.append(f"Hoja Productos: hay {invalidos} compras mínimas vacías o no numéricas.")

    except Exception as e:
        errores.append(f"No pude validar la hoja Productos. Error: {e}")

    try:
        aliases = pd.read_excel(maestro_path, sheet_name="Aliases")
        cols = {_norm_col(c): c for c in aliases.columns}
        if not ("alias detectado" in cols or "alias" in cols or "nombre original" in cols):
            errores.append("Hoja Aliases: falta columna de alias.")
        if "producto base" not in cols:
            errores.append("Hoja Aliases: falta columna 'Producto Base'.")
    except Exception as e:
        errores.append(f"No pude validar la hoja Aliases. Error: {e}")

    try:
        exclusiones = pd.read_excel(maestro_path, sheet_name="Exclusiones")
        if exclusiones.shape[1] < 1:
            errores.append("Hoja Exclusiones: debe tener al menos una columna con productos a excluir.")
    except Exception as e:
        errores.append(f"No pude validar la hoja Exclusiones. Error: {e}")

    try:
        config = pd.read_excel(maestro_path, sheet_name="Configuración")
        cols = {_norm_col(c): c for c in config.columns}
        if "parametro" not in cols or "valor" not in cols:
            errores.append("Hoja Configuración: debe tener columnas Parámetro y Valor.")
    except Exception as e:
        errores.append(f"No pude validar la hoja Configuración. Error: {e}")

    return len(errores) == 0, errores, advertencias



def validar_carrito_precios(carrito_path):
    errores = []
    advertencias = []
    try:
        df = pd.read_excel(carrito_path)
    except Exception as e:
        return False, [f"No pude abrir el Excel de carrito/precios. Error: {e}"], []

    if df.shape[1] < 9:
        return False, ["El archivo debe tener al menos 9 columnas. Se espera código en columna B y precio en columna I."], []

    codigo_col = df.columns[1]
    precio_col = df.columns[8]

    if len(df[codigo_col].dropna()) == 0:
        errores.append("No se encontraron códigos en la columna B.")

    precios = pd.to_numeric(df[precio_col], errors="coerce")
    precios_validos = precios.dropna()

    if len(precios_validos) == 0:
        errores.append("No se encontraron precios numéricos en la columna I.")
    elif (precios_validos <= 0).any():
        advertencias.append("Algunos precios de la columna I son cero o negativos.")

    advertencias.append(f"Columna B detectada como código: {codigo_col}")
    advertencias.append(f"Columna I detectada como precio: {precio_col}")
    advertencias.append(f"Filas con precios válidos: {len(precios_validos)}")

    return len(errores) == 0, errores, advertencias


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
    st.caption(f"Versión {APP_VERSION}")

    st.subheader("Maestro de productos")
    if DEFAULT_MAESTRO.exists():
        st.success("Maestro actual encontrado.")
        st.download_button(
            "Descargar Maestro actual",
            data=DEFAULT_MAESTRO.read_bytes(),
            file_name="Maestro_Productos_Grido.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    else:
        st.error("No se encontró el maestro en la carpeta maestro/.")

    st.markdown("---")
    st.subheader("Archivo de precios / carrito")
    if DEFAULT_CARRITO.exists():
        st.success("Archivo de precios actual encontrado.")
        st.download_button(
            "Descargar carrito/precios actual",
            data=DEFAULT_CARRITO.read_bytes(),
            file_name="Modelo_de_Carrito.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    else:
        st.warning("No se encontró el archivo de precios en la carpeta datos/.")

    st.markdown("---")
    st.subheader("Subir archivos de administración")
    st.info("Ingresá el PIN de administrador para subir Maestro o Carrito/precios.")

    pin = st.text_input("PIN de administrador", type="password")

    if pin != ADMIN_PIN:
        st.warning("Ingresá el PIN correcto para habilitar la carga.")
    else:
        st.success("PIN correcto. Carga habilitada.")
        tab_maestro, tab_carrito = st.tabs(["Actualizar maestro", "Actualizar carrito/precios"])

        with tab_maestro:
            nuevo_maestro = st.file_uploader(
                "Nuevo Maestro_Productos_Grido.xlsx",
                type=["xlsx"],
                help="Debe tener las hojas Productos, Aliases, Exclusiones y Configuración.",
                key="nuevo_maestro_upload"
            )
            if nuevo_maestro:
                with tempfile.TemporaryDirectory() as tmp_dir:
                    tmp_path = Path(tmp_dir) / "Maestro_Productos_Grido.xlsx"
                    tmp_path.write_bytes(nuevo_maestro.getvalue())
                    ok, errores, advertencias = validar_maestro(tmp_path)

                    if ok:
                        st.success("✅ Maestro validado correctamente.")
                        if advertencias:
                            with st.expander("Ver advertencias del maestro", expanded=False):
                                for adv in advertencias:
                                    st.warning(adv)
                        st.warning("En Streamlit Cloud este reemplazo puede perderse si la app se reinicia. Para hacerlo permanente, subilo a GitHub.")
                        if st.button("Reemplazar maestro en esta sesión", type="primary", use_container_width=True):
                            DEFAULT_MAESTRO.parent.mkdir(exist_ok=True)
                            DEFAULT_MAESTRO.write_bytes(tmp_path.read_bytes())
                            st.success("Maestro reemplazado en esta sesión.")
                    else:
                        st.error("❌ El maestro no pasó la validación.")
                        for err in errores:
                            st.error(err)

        with tab_carrito:
            nuevo_carrito = st.file_uploader(
                "Nuevo Modelo_de_Carrito.xlsx",
                type=["xlsx"],
                help="Debe tener códigos en columna B y precios en columna I.",
                key="nuevo_carrito_upload"
            )
            if nuevo_carrito:
                with tempfile.TemporaryDirectory() as tmp_dir:
                    tmp_path = Path(tmp_dir) / "Modelo_de_Carrito.xlsx"
                    tmp_path.write_bytes(nuevo_carrito.getvalue())
                    ok, errores, advertencias = validar_carrito_precios(tmp_path)

                    if ok:
                        st.success("✅ Carrito/precios validado correctamente.")
                        if advertencias:
                            with st.expander("Ver detalles del carrito/precios", expanded=False):
                                for adv in advertencias:
                                    st.info(adv)
                        st.warning("En Streamlit Cloud este reemplazo puede perderse si la app se reinicia. Para hacerlo permanente, subilo a GitHub en datos/.")
                        if st.button("Reemplazar carrito/precios en esta sesión", type="primary", use_container_width=True):
                            DEFAULT_CARRITO.parent.mkdir(exist_ok=True)
                            DEFAULT_CARRITO.write_bytes(tmp_path.read_bytes())
                            st.success("Carrito/precios reemplazado en esta sesión.")
                    else:
                        st.error("❌ El archivo de carrito/precios no pasó la validación.")
                        for err in errores:
                            st.error(err)

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
            posibles_faltantes = result.get("posibles_faltantes")
            stock_negativo = result.get("stock_negativo")

            st.success("Pedido generado correctamente.")

            m1, m2, m3, m4, m5 = st.columns(5)
            m1.metric("Productos en pedido", len(pedido))
            m2.metric("Stock negativo", len(stock_negativo) if stock_negativo is not None else 0)
            m3.metric("Sin clasificar", len(sin_clasificar))
            m4.metric("Posibles faltantes", len(posibles_faltantes) if posibles_faltantes is not None else 0)
            m5.metric("Packs sugeridos", int(pedido["Packs a Comprar"].fillna(0).sum()))

            valor_stock_total = result.get("valor_stock_total", 0)
            valor_pedido_total = result.get("valor_pedido_total", 0)

            v1, v2 = st.columns(2)
            v1.metric("Valor stock actual", formato_moneda_ar(valor_stock_total))
            v2.metric("Valor pedido sugerido", formato_moneda_ar(valor_pedido_total))

            valorizacion_categoria = result.get("valorizacion_categoria")
            valorizacion_producto = result.get("valorizacion_producto")

            if valorizacion_categoria is not None:
                with st.expander("Ver valorización por categoría", expanded=False):
                    st.dataframe(valorizacion_categoria, use_container_width=True)

            if valorizacion_producto is not None:
                with st.expander("Ver valorización por producto", expanded=False):
                    st.dataframe(valorizacion_producto, use_container_width=True)

            if stock_negativo is not None and len(stock_negativo) > 0:
                st.error(
                    "⚠️ Se detectaron productos con stock negativo. "
                    "Esto puede indicar errores de inventario o descarga."
                )

            if posibles_faltantes is not None and len(posibles_faltantes) > 0:
                st.warning(
                    "Se detectaron posibles faltantes: productos sin ventas y con stock bajo. "
                    "No modifican automáticamente el pedido; revisalos manualmente."
                )

            st.download_button(
                "DESCARGAR PEDIDO FINAL",
                data=output.read_bytes(),
                file_name="Pedido_Final.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

            with st.expander("Ver vista previa del pedido", expanded=False):
                st.dataframe(pedido.head(100), use_container_width=True)

            if stock_negativo is not None and len(stock_negativo) > 0:
                with st.expander("Ver stock negativo", expanded=True):
                    st.error("Productos con stock negativo detectado.")
                    st.dataframe(stock_negativo, use_container_width=True)

            if posibles_faltantes is not None and len(posibles_faltantes) > 0:
                with st.expander("Ver posibles faltantes", expanded=True):
                    st.warning("Estos productos no tuvieron ventas y tienen stock bajo. Revisar manualmente.")
                    st.dataframe(posibles_faltantes, use_container_width=True)

            if len(sin_clasificar) > 0:
                with st.expander("Ver productos sin clasificar", expanded=True):
                    st.warning("Hay productos sin clasificar. Revisá esta hoja en el Excel y actualizá el Maestro.")
                    st.dataframe(sin_clasificar, use_container_width=True)

        except Exception as e:
            st.error("Ocurrió un error al generar el pedido.")
            st.exception(e)
