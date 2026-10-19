"""
comprobar_analisis_final.py -- comprueba que analisis_final.md cuadra con datos_radar.json.

Avisa si:
  - el PERIODO del texto (<!-- periodo: ... -->) no es el de los datos, o
  - alguna CIFRA con decimal del texto (p. ej. -7,2%) no coincide con ningún dato
    actual de su índice.
No comprueba porcentajes enteros ni otras cifras sueltas (esas se revisan a mano).

Uso:
    python comprobar_analisis_final.py [analisis_final.md] [datos_radar.json]

Por defecto usa analisis_final.md y datos_radar.json de esta misma carpeta.
Código de salida: 0 = todo cuadra, 1 = hay avisos, 2 = falta algún fichero.
"""
import json
import os
import re
import sys

CLAVES = ("oit", "ocde", "jrc", "anthropic")


# ---------------------------------------------------------------------------
# Lectura del .md y comprobación de cifras y periodo
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





def main(argv):
    md_path = argv[1] if len(argv) > 1 else "analisis_final.md"
    js_path = argv[2] if len(argv) > 2 else "datos_radar.json"
    for p in (md_path, js_path):
        if not os.path.exists(p):
            print(f"No encuentro {p}. Ejecuta este script desde la carpeta del Radar.")
            return 2
    with open(js_path, encoding="utf-8") as f:
        data = json.load(f)
    avisos = comprobar(leer_md_fichero(md_path), data)
    print(f"Comprobando {md_path} contra {js_path} ({data.get('compare_text', {}).get('mes', '?')})")
    if avisos:
        print("AVISOS:")
        for a in avisos:
            print("  -", a)
        print("Revisa el texto (o regenera el borrador) antes de publicar.")
        return 1
    print("OK: el periodo y las cifras con decimal del texto cuadran con los datos.")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
