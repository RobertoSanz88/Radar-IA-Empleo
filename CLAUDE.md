# Radar IA-Empleo

Página estática del Observatorio de las Ocupaciones (SEPE) que cruza la exposición de las ocupaciones (CNO) a la IA
con la evolución de la contratación y del paro registrado. Flujo mensual y tabla de ficheros: `README.md` (no repetirlos aquí).
Imagen del Observatorio: seguir `~/.claude/design/design.md`.

## Reglas que no se deducen del código

- **La columna OCDE es ficticia** (solo prueba la maqueta). No tocarla ni "mejorarla" hasta que el usuario pase la
  lista real de ocupaciones con su nivel de exposición. Fuentes reales: OIT, JRC y Anthropic.
- **`analisis_final.md` es texto revisado a mano.** El paso 1 (`Crear_datos_json_y_borrador_analisis.py`) nunca lo
  toca y no hay que hacerlo tampoco al regenerar datos: si el mes cambia, se escribe uno nuevo a partir de
  `analisis_borrador.md` y se valida con `python comprobar_analisis_final.py` (periodo y cifras con decimal).
- **Los cambios porcentuales se calculan con totales absolutos** (periodo 1 frente a periodo 2), nunca como media de
  porcentajes. El Excel trae las proporciones en tanto por uno y el JSON se escribe multiplicado por 100.
- **Entrada mensual:** `Datos radar SEPE original.xlsx` (4 pestañas: parados y contratos por CNO, año y mes). No se
  versiona (`.gitignore`). Ya no existe la plantilla manual (`datos_radar_plantilla.xlsx`): no recrearla.
- **`radar-ia-empleo-embebido.html` se genera**, no se edita a mano: lo crea `embeber_datos.py` sustituyendo las dos
  zonas entre marcas `<<EMBED-CABECERA-INICIO/FIN>>` y `<<EMBED-INIT-INICIO/FIN>>` del HTML de producción. Cambiar
  el HTML solo fuera de esas marcas, o el script deja de funcionar. El embebido es el entregable que se manda por
  correo; el de producción necesita servidor web (carga el JSON y el `.md` por `fetch`).
- **`radar-ia-empleo-produccion_old_negro.html`** es solo referencia del estilo antiguo: no modificarla.
- **Dónde se publica la página:** sin decidir. No asumir ningún servidor ni proponer subir nada.

## Trabajar con el código

- `Crear_datos_json_y_borrador_analisis.py` tiene más de 3.000 líneas y casi todas son diccionarios de mapeo
  (CNO → ISCO → nivel JRC/Anthropic, líneas ~650-2260). No leerlo entero: la lógica está en `build_json` (~2668) y
  en las funciones de redacción del borrador (~3060 en adelante).
- Requisitos: Python 3 y `openpyxl`. Se trabaja desde el PC de casa (usuario `rsanz`, sin Netskope): `pip` normal,
  sin `--trusted-host`.
- Los `.bat` usan `chcp 65001` y `PYTHONIOENCODING=utf-8`: mantenerlo, o los acentos salen rotos en consola.
- Cualquier cambio en el cálculo se prueba regenerando `datos_radar.json` con el Excel del mes y comparándolo con el
  anterior, no solo comprobando que el script arranca.
