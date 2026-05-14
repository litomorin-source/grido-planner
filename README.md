
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
