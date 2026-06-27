# GridoPlanner

## Versión

Stable 1.9.0

## Cambios principales

- Maestro y Carrito se administran desde la app.
- La app usa GitHub como fuente persistente de datos.
- Usuario solo carga:
  - Stock CSV
  - Cajas por Sabor XLSX
  - Mix de Ventas XLSX
- Administrador con PIN para:
  - Centro de Datos
  - Actualizar Maestro
  - Actualizar Carrito
- Validación reforzada de Maestro.
- Validación reforzada de Carrito por estructura real del Modelo de Carrito.
- Tutorial no incluido por ahora.

## Secrets necesarios en Streamlit

```toml
GITHUB_TOKEN = "tu_token"
GITHUB_REPO = "litomorin-source/grido-planner"
GITHUB_BRANCH = "main"
```

## Archivos persistentes esperados en GitHub

```text
data/Maestro_Productos_Grido.xlsx
data/Modelo_de_Carrito.xlsx
```
