"""
Genera una versión de Radar IA-Empleo con los datos y los textos de análisis
INCRUSTADOS dentro del propio HTML, en vez de cargarlos por fetch de
datos_radar.json y analisis_final.md. Resultado: un único archivo que cualquiera
puede abrir con doble clic (por ejemplo, recibido por correo) sin servidor local.

Uso:
    python embeber_datos.py [pagina_html] [datos_json] [salida_html] [analisis_md]

Por defecto:
    pagina_html = radar-ia-empleo-produccion.html
    datos_json  = datos_radar.json
    salida_html = radar-ia-empleo-embebido.html
    analisis_md = analisis_final.md   (si no existe, el Análisis sale vacío)

Cómo funciona: el HTML de producción lleva dos zonas delimitadas por marcas
    // <<EMBED-CABECERA-INICIO>> ... // <<EMBED-CABECERA-FIN>>
    // <<EMBED-INIT-INICIO>>     ... // <<EMBED-INIT-FIN>>
Este script sustituye el contenido de esas dos zonas por la versión con datos
incrustados. Lo demás del HTML se copia tal cual. Así ya no hace falta mantener
a mano una copia exacta del código del HTML aquí dentro: si cambias el HTML fuera
de las marcas, este script sigue funcionando.

Antes de embeber, comprueba que el análisis cuadra con los datos (periodo y
cifras) y AVISA si no (no bloquea).
"""
import json
import os
import re
import sys

CLAVES = ("oit", "ocde", "jrc", "anthropic")


# ---------------------------------------------------------------------------
# Comprobación del análisis antes de embeber (misma lógica que comprobar_analisis_final.py)
# ---------------------------------------------------------------------------
def leer_md(texto):
    """-> {"periodo": str|None, "secciones": {clave: {"subtitulo": str|None, "parrafos": [str]}}}
    (el mismo formato que interpreta el JavaScript de la página)."""
    m = re.search(r"<!--\s*periodo:\s*(.+?)\s*-->", texto)
    periodo = m.group(1).strip() if m else None
    texto = re.sub(r"<!--.*?-->", "", texto, flags=re.S)
    secciones, clave, bloque = {}, None, []

    def cerrar():
        if clave in CLAVES:
            sub, pars, cur = None, [], []
            for ln in bloque + [""]:
                if ln.startswith("### "):
                    sub = ln[4:].strip()
                elif ln.strip():
                    cur.append(ln.strip())
                elif cur:
                    pars.append(" ".join(cur))
                    cur = []
            secciones[clave] = {"subtitulo": sub, "parrafos": pars}

    for ln in texto.splitlines():
        h = re.match(r"^##\s+([\w-]+)\s*$", ln)
        if h:
            cerrar()
            clave, bloque = h.group(1).lower(), []
        elif clave is not None:
            bloque.append(ln)
    cerrar()
    return {"periodo": periodo, "secciones": secciones}


def leer_md_fichero(path):
    with open(path, encoding="utf-8-sig") as f:
        return leer_md(f.read())


# ---------------------------------------------------------------------------
# Comprobación de cifras y periodo
# ---------------------------------------------------------------------------
def _pool_indice(data, ix):
    """Todas las cifras (variación % 'mes') que un texto de ese índice puede citar."""
    vals = set()

    def add(v):
        if v is not None:
            vals.add(round(float(v), 1))

    for met in ("contratos", "desempleo"):
        b = data["breakdowns"][met]
        add(b["national"]["mes"])
        idx = b["indices"][ix]
        for t in idx["tiers"]:
            add(t["mes"])
            for code in (t.get("occupations") or []):
                o = b["occupations"].get(code)
                if o:
                    add(o["mes"])
                    for v in (o["sexo"]["mes"] + o["edad"]["mes"]):
                        add(v)
        for g in idx["resumen"]["edad"] + idx["resumen"]["sexo"]:
            add(g["mes"])
    return vals


def comprobar(md, data):
    """-> lista de avisos (vacía = todo cuadra). md es el resultado de leer_md()."""
    avisos = []
    esperado = data.get("compare_text", {}).get("mes")
    if md["periodo"] is None:
        avisos.append("El .md no indica el periodo (falta <!-- periodo: ... -->).")
    elif esperado and md["periodo"].strip().lower() != esperado.strip().lower():
        avisos.append(f"PERIODO: el texto es de «{md['periodo']}» pero los datos son de «{esperado}».")
    for ix in CLAVES:
        sec = md["secciones"].get(ix)
        if not sec:
            continue
        if ix not in data["breakdowns"]["contratos"]["indices"]:
            continue
        pool = _pool_indice(data, ix)
        pool_abs = {abs(v) for v in pool}
        texto = " ".join(([sec["subtitulo"]] if sec["subtitulo"] else []) + sec["parrafos"])
        for tok in sorted(set(re.findall(r"(?<![\d,])([+\-]?\d+,\d)%", texto))):
            v = round(float(tok.replace(",", ".")), 1)
            firmado = tok[0] in "+-"
            ok = (v in pool) if firmado else (abs(v) in pool_abs)
            if not ok:
                avisos.append(f"[{ix}] la cifra {tok}% no coincide con ningún dato actual de ese índice.")
    return avisos



CABECERA_NEW = '''// <<EMBED-CABECERA-INICIO>>
/* =========================================================================
   Versión con los datos y el análisis INCRUSTADOS (no hace fetch a
   datos_radar.json ni a analisis_final.md). Se abre con doble clic, sin
   servidor local: pensada para enviar por correo o compartir como archivo
   suelto. Generada con embeber_datos.py a partir de
   radar-ia-empleo-produccion.html + datos_radar.json + analisis_final.md --
   para actualizar, vuelve a ejecutar ese script, no edites este archivo a mano.
   ========================================================================= */
const EMBEDDED_DATA = {json_data};
const EMBEDDED_ANALISIS_MD = {md_data};
// <<EMBED-CABECERA-FIN>>'''

INIT_NEW = '''// <<EMBED-INIT-INICIO>>
/* ---------- carga de datos (incrustados) ---------- */
async function cargarAnalisis(){
  try { ANALISIS = parseAnalisisMd(EMBEDDED_ANALISIS_MD); }
  catch (e) { console.warn('Análisis incrustado no válido:', e); }
}

async function init(){
  try {
    await cargarAnalisis();
    buildFromData(EMBEDDED_DATA);
    renderHero();
    renderTierList();
    renderSummaryCards();
    renderAnalysisBox();
  } catch (err){
    console.error(err);
    document.getElementById('tierList').innerHTML =
      '<div class="load-error">No se han podido pintar los datos incrustados (' + err.message + ').</div>';
  }
}
// <<EMBED-INIT-FIN>>'''


def _zona(nombre):
    return re.compile(r"// <<EMBED-%s-INICIO>>.*?// <<EMBED-%s-FIN>>" % (nombre, nombre), re.S)


def main():
    html_path = sys.argv[1] if len(sys.argv) > 1 else "radar-ia-empleo-produccion.html"
    json_path = sys.argv[2] if len(sys.argv) > 2 else "datos_radar.json"
    out_path = sys.argv[3] if len(sys.argv) > 3 else "radar-ia-empleo-embebido.html"
    md_path = sys.argv[4] if len(sys.argv) > 4 else "analisis_final.md"

    with open(html_path, encoding="utf-8") as f:
        html = f.read()
    with open(json_path, encoding="utf-8") as f:
        data = json.load(f)

    re_cab, re_init = _zona("CABECERA"), _zona("INIT")
    if not (re_cab.search(html) and re_init.search(html)):
        raise SystemExit(
            "No encuentro las marcas <<EMBED-CABECERA-...>> / <<EMBED-INIT-...>> en el HTML. "
            "¿Es la versión de producción actual (radar-ia-empleo-produccion.html)?"
        )

    md_text = ""
    if os.path.exists(md_path):
        with open(md_path, encoding="utf-8-sig") as f:
            md_text = f.read()
        avisos = comprobar(leer_md(md_text), data)
        if avisos:
            print(f"AVISO: {md_path} no cuadra con {json_path}:")
            for a in avisos:
                print("  -", a)
        else:
            print(f"Análisis comprobado: {md_path} cuadra con {json_path}.")
    else:
        print(f"AVISO: no existe {md_path}; el embebido saldrá con el Análisis vacío.")

    # json.dumps -> cadena JS válida; "</" se escapa para que nada cierre el <script> antes de tiempo
    seguro = lambda t: t.replace("</", "<\\/")
    cab = CABECERA_NEW.replace("{json_data}", seguro(json.dumps(data, ensure_ascii=False))) \
                      .replace("{md_data}", seguro(json.dumps(md_text, ensure_ascii=False)))
    html = re_cab.sub(lambda m: cab, html, count=1)
    html = re_init.sub(lambda m: INIT_NEW, html, count=1)

    with open(out_path, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"OK -> {out_path}  ({len(html):,} caracteres)")


if __name__ == "__main__":
    main()
