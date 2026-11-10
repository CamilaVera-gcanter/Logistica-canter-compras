# Conversor Monday.com ➜ Browix

Aplicación web (Streamlit) que convierte el Excel exportado de Monday en la
plantilla de importación de **Sub tareas** de Browix.

## Archivos
| Archivo | Para qué sirve |
|---|---|
| `app.py` | Interfaz web (carga, asignación de transportistas, vista previa, descarga) |
| `transformador.py` | Lógica de transformación (lectura, limpieza, mapeo, exportación) |
| `requirements.txt` | Librerías necesarias |
| `assets/logo_grupo_canter.png` | Logo de Grupo Canter que se muestra en la app |

## Mapeo de columnas
| Browix | Origen en Monday |
|---|---|
| Tarea | `Proveedor Elegido` + " - " + `Name` (obra) |
| Agrupador Tarea | **N° de tarea de Browix** asignado a cada OC en la app (manual o secuencial). Todas las sub-tareas de una OC llevan el mismo número |
| Fecha | Siempre `Fecha estimada Entrega S/Proveedor`, como texto dd/mm/aaaa |
| Dni responsable | DNI del transportista asignado (tabla de la barra lateral de la app) |
| Subtarea | `Item / Producto` |
| Detalle Subtarea | `FALTANTE` |
| Detalle | Vacía |
| Observaciones | `Numero OC` |

El Excel exportado tiene **una sola hoja** (`Sub tarea`). La tabla de transportistas solo vive en la app.

## Numeración de tareas
En el paso 2 de la app se carga el **N° de tarea para la primera OC** (continuando la numeración de la última tarea cargada en Browix). Con la opción *Numerar las siguientes OC en forma secuencial* activa, cada OC siguiente recibe +1. Cualquier número se puede corregir a mano en la tabla. La descarga se habilita cuando todas las OC tienen número.

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
