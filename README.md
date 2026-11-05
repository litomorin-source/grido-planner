# GridoPlanner Stable 2.4.0

Migración completa desde la versión probada de GridoPlanner Labs (respaldo del 05/11/2026), fusionada con producción.

Incluye motor de pedidos, informes, costos según Mix de Ventas, importación de Carrito, gráficos, descarga persistente y omisiones de futuros pedidos.

**Datos:** Maestro y Carrito vigentes de Labs; historial de stock fusionado por ID sin descartar registros distintos; lista de omisiones de Labs; archivos auxiliares originales de producción conservados.

**Instalación:** extraer ZIP y subir el contenido del directorio al raíz del repositorio `grido-planner`, manteniendo subcarpetas (`ui/`, `core/`, `data/`, etc.). No subir el ZIP como un único archivo. No sobrescribir los secretos configurados en Streamlit Cloud.

**Verificar luego de subir:** acceso, maestro/carrito, cálculos, omisiones, historial, Excel Pedido Final y Carrito importable. No se efectuó prueba de la app en Streamlit Cloud.
