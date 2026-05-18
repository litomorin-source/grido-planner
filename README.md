
# GridoPlanner Web v1

App web experimental para generar pedidos de Grido a partir de tres exports:

1. Stock CSV
2. Cajas por Sabor XLSX
3. Data Power BI XLSX

## Archivos incluidos

- `app.py`: interfaz web Streamlit
- `motor.py`: lógica de procesamiento
- `requirements.txt`: dependencias
- `maestro/Maestro_Productos_Grido.xlsx`: maestro central

## Cómo probar localmente

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Cómo subir a Streamlit Cloud

1. Subir esta carpeta a GitHub.
2. Entrar a Streamlit Community Cloud.
3. Crear nueva app.
4. Elegir el repositorio.
5. Main file path: `app.py`.
6. Deploy.

## Modo usuario

Permite subir los 3 archivos, configurar semanas objetivo y tiempo de reposición, y descargar el pedido.

## Modo administrador

Por ahora permite descargar el maestro actual. En una próxima versión se agregará carga segura de maestro.


## Cambios v2

- Estados de carga compactos.
- Explicación breve de Semanas objetivo, Tiempo de reposición y Días analizados.
- Resumen especial de grupos detectados en Power BI.
- Advertencia sobre desplegar el signo `+` en Power BI antes de exportar.
- Vista previa del pedido y productos sin clasificar dentro de desplegables.


## Cambios v3

- Nueva detección de `Posibles faltantes`.
- Regla: ventas = 0 y stock bajo.
- Helado a granel: stock bajo menor a 1.
- Otros productos: stock bajo menor a 5.
- Se agrega hoja `Posibles faltantes` al Excel.
- La app muestra contador y alerta visual de posibles faltantes.
- Esta regla no modifica automáticamente la compra sugerida.


## Cambios v4

- Versión visible en la app.
- Modo Administrador mejorado.
- Descarga del maestro actual.
- Carga de nuevo maestro desde la web.
- Validación del maestro antes de reemplazarlo.
- Advertencia: en Streamlit Cloud el reemplazo desde la app es temporal si la app se reinicia; para hacerlo permanente hay que subir el maestro a GitHub.


## Cambios v5

- Se incorpora el maestro validado por Gabriel como maestro oficial incluido.
- Se agrega PIN simple para habilitar la carga/reemplazo del maestro desde modo administrador.
- PIN inicial: `2468`.
- El PIN se puede cambiar editando `ADMIN_PIN` en `app.py`.


## Cambios Mark VIII / v0.8-beta

- Tutorial visual dentro de la app.
- Screenshots reales de Stock, Cajas por Sabor y Mix de Ventas.
- Títulos del tutorial más compactos.
- Carpeta interna y ZIP con nombre consistente.
- Ancho fijo para Código Compra en hoja Carrito: 15.


## Cambios Mark IX / v0.9-beta

- Corrección definitiva de branding y versión.
- Toda la app muestra Mark IX (v0.9-beta).
- Corrección de renderizado de screenshots en Streamlit Cloud.


## Cambios Mark X / v1.0-beta

- Corrección real del tutorial visual.
- El tutorial ahora llama efectivamente a los screenshots incluidos.
- Versión visible: Mark X (v1.0-beta).
