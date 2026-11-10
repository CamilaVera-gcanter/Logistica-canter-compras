# Conversor Monday.com ➜ Browix

Aplicación web (Streamlit) que convierte el Excel exportado de Monday en la
plantilla de importación de **Sub tareas** de Browix.

## Archivos
| Archivo | Para qué sirve |
|---|---|
| `app.py` | Interfaz web (carga, asignación de transportistas, vista previa, descarga) |
| `transformador.py` | Lógica de transformación (lectura, limpieza, mapeo, exportación) |
| `requirements.txt` | Librerías necesarias |

## Mapeo de columnas
| Browix | Origen en Monday |
|---|---|
| Tarea | `Proveedor Elegido` + " - " + `Name` (obra) |
| Agrupador Tarea | `Numero OC` |
| Fecha | Columna de fecha elegida (por defecto `Fecha Tarea Browix` si tiene datos, si no `Fecha estimada Entrega S/Proveedo`), como texto dd/mm/aaaa |
| Dni responsable | DNI del transportista asignado (tabla TRANSPORTISTAS) |
| Subtarea | `Item / Producto` |
| Detalle Subtarea | `FALTANTE` |
| Detalle / Observaciones | Vacías |

## Instalación (una sola vez)
1. Instalar Python 3.10 o superior desde https://www.python.org (en Windows, tildar **"Add Python to PATH"**).
2. Abrir una terminal en la carpeta del proyecto y ejecutar:
   ```bash
   pip install -r requirements.txt
   ```

## Ejecución
```bash
streamlit run app.py
```
Se abre el navegador en http://localhost:8501.

## Agregar transportistas de forma permanente
Editar el diccionario `TRANSPORTISTAS_DEFAULT` al inicio de `transformador.py`.
