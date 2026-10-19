# Radar IA-Empleo

Página estática del Observatorio de las Ocupaciones (SEPE) que cruza la exposición de las ocupaciones (CNO) a la IA
con la evolución de la contratación y del paro registrado. Fuentes de exposición: OIT, JRC y Anthropic.
La columna OCDE es **ficticia** y solo sirve para probar la maqueta.

## Ficheros

| Fichero | Qué es |
|---|---|
| `radar-ia-empleo-produccion.html` | La página. Carga `datos_radar.json` y `analisis_final.md` (necesita servidor web). |
| `radar-ia-empleo-embebido.html` | La misma página con datos y análisis incrustados; se abre con doble clic. |
| `datos_radar.json` | Cifras del mes (lo genera el paso 1). |
| `analisis_borrador.md` | Borrador del análisis escrito por reglas (lo genera el paso 1). |
| `analisis_final.md` | Texto revisado a mano que muestra la página. El paso 1 nunca lo toca. |
| `JRC_correspondencias_y_niveles.xlsx` | Correspondencias y niveles de exposición JRC. |
| `design/` | Identidad visual del Observatorio (`design.md`) y logotipos SVG. |

## Flujo mensual

1. Copiar el Excel mensual como `Datos radar SEPE original.xlsx` en la carpeta (no se versiona, ver `.gitignore`).
2. `Crear_datos_json_y_borrador_analisis.bat` (o `python Crear_datos_json_y_borrador_analisis.py`) → `datos_radar.json` y `analisis_borrador.md`.
3. Revisar el borrador a mano y guardarlo como `analisis_final.md`.
4. `python comprobar_analisis_final.py` → avisa si el periodo o las cifras del texto no cuadran con los datos.
5. `python embeber_datos.py` → `radar-ia-empleo-embebido.html`.
6. Commit y push.

Requisitos: Python 3 y `openpyxl`.
