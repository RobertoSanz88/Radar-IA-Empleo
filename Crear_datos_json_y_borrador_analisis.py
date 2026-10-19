#!/usr/bin/env python3
"""
Crear_datos_json_y_borrador_analisis.py — Radar IA-Empleo
==========================================================

RESUMEN (paso 1 de cada mes; se lanza con Crear_datos_json_y_borrador_analisis.bat):
  ENTRADA : "Datos radar SEPE original.xlsx"  (el Excel mensual)
  SALIDAS : datos_radar.json      -> las cifras que carga la Web
            analisis_borrador.md  -> borrador del análisis, escrito por reglas
  Después, el comentarista humano (opcionalmente ayudado por IA) revisa el borrador
  y lo guarda como analisis_final.md, que es el texto que carga la Web. Este script
  NUNCA toca analisis_final.md. (Para comprobar que cuadra con los datos: comprobar_analisis_final.py)

Nueva versión (oct-2026): construye datos_radar.json directamente desde el
export en bruto de la base de datos de Roberto ("Datos radar SEPE original.xlsx",
4 pestañas: Parados por CNO año/mes, P Contratadas por CNO año/mes), SIN pasar ya por
la plantilla manual (datos_radar_plantilla.xlsx / Cabecera / Niveles_Exposicion...).

Qué hace:
  1. Lee las 4 pestañas tal cual vienen (cifras absolutas, nunca formulas).
  2. Calcula en Python (nunca como media de porcentajes, siempre
     total absoluto periodo 1 vs total absoluto periodo 2 -> % de cambio):
       - "Gran total" (todas las ocupaciones)       -> Total nacional
       - Un subtotal por cada uno de los 6 niveles de exposición
         (agrupando por la columna "Exposure" del original)
       - Para las 48 ocupaciones de máxima + significativa: % total,
         % por sexo y % por edad
       - El resumen (Cajas_Resumen) de sexo/edad agregado de las 10
         ocupaciones de máxima exposición
  3. Multiplica todo x100 al escribir el JSON (el Excel lo tiene en tanto por uno).
  4. Actividades económicas (CNAE): de momento vacío, como se ha pedido.
  5. (Opcional, --enriquecido) escribe también un Excel tipo
     "Datos_radar_SEPE_final" con 7 filas nuevas (Gran total + 6 niveles)
     y fórmulas reales de Excel, solo para auditar a ojo. El JSON nunca
     se construye releyendo esas fórmulas (una fórmula recién escrita con
     openpyxl no tiene valor cacheado hasta que Excel la recalcula).

Uso:
  python Crear_datos_json_y_borrador_analisis.py [origen.xlsx] [salida.json] [--enriquecido [salida.xlsx]]

Códigos CNO de las 9 filas con Exposure "=#N/A" (sin clasificar OIT) se
excluyen de todos los niveles pero SÍ cuentan en el Gran total / Total nacional.
"""

import sys
import re
import json
import copy
from pathlib import Path

import openpyxl
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# 1. Catálogo fijo: ya no depende de la plantilla. Si la OIT amplía la lista
#    de ocupaciones de máxima/significativa exposición, se actualiza aquí.
# ---------------------------------------------------------------------------

# Clasificación OIT por CNO (4 dígitos) — VERSIÓN V3 corregida por Roberto (oct-2026):
# fichero "Índice Exposición a IA por CNO a 4 dig desde CIUO de OIT v3.xlsx", hoja
# "Index por CNO 4 dig". Respecto a la V2 cambian 6 CNO: 4309 pasa a Gradiente 4;
# 2159 pasa a Mínima; 3331, 5120, 5910 y 5921 pasan a Sin exposición. Esta tabla
# MANDA sobre la columna "Exposure" del Excel en bruto (que sigue trayendo la V2).
OIT_EXPOSURE_V3 = {
    "1111": "Not Exposed",
    "1112": "Minimal Exposure",
    "1113": "Minimal Exposure",
    "1120": "Minimal Exposure",
    "1211": "Minimal Exposure",
    "1212": "Minimal Exposure",
    "1219": "Minimal Exposure",
    "1221": "Gradient 2",
    "1222": "Minimal Exposure",
    "1223": "Minimal Exposure",
    "1311": "Minimal Exposure",
    "1312": "Not Exposed",
    "1313": "Minimal Exposure",
    "1314": "Minimal Exposure",
    "1315": "Minimal Exposure",
    "1316": "Minimal Exposure",
    "1321": "Gradient 2",
    "1322": "Minimal Exposure",
    "1323": "Minimal Exposure",
    "1324": "Minimal Exposure",
    "1325": "Minimal Exposure",
    "1326": "Minimal Exposure",
    "1327": "Gradient 2",
    "1329": "Minimal Exposure",
    "1411": "Gradient 1",
    "1419": "Gradient 1",
    "1421": "Minimal Exposure",
    "1422": "Minimal Exposure",
    "1429": "Minimal Exposure",
    "1431": "Gradient 2",
    "1432": "Gradient 2",
    "1501": "Minimal Exposure",
    "2111": "Not Exposed",
    "2112": "Not Exposed",
    "2121": "Not Exposed",
    "2122": "Not Exposed",
    "2123": "Minimal Exposure",
    "2130": "Not Exposed",
    "2140": "Minimal Exposure",
    "2151": "Not Exposed",
    "2152": "Minimal Exposure",
    "2153": "Minimal Exposure",
    "2154": "Not Exposed",
    "2155": "Not Exposed",
    "2156": "Not Exposed",
    "2157": "Not Exposed",
    "2158": "Not Exposed",
    "2159": "Minimal Exposure",
    "2210": "Minimal Exposure",
    "2220": "Minimal Exposure",
    "2230": "Not Exposed",
    "2240": "Not Exposed",
    "2251": "Not Exposed",
    "2252": "Not Exposed",
    "2311": "Not Exposed",
    "2312": "Not Exposed",
    "2321": "Gradient 1",
    "2322": "Minimal Exposure",
    "2323": "Minimal Exposure",
    "2324": "Minimal Exposure",
    "2325": "Gradient 2",
    "2326": "Minimal Exposure",
    "2329": "Minimal Exposure",
    "2411": "Gradient 1",
    "2412": "Gradient 3",
    "2413": "Minimal Exposure",
    "2414": "Minimal Exposure",
    "2415": "Gradient 3",
    "2416": "Gradient 3",
    "2421": "Gradient 2",
    "2422": "Not Exposed",
    "2423": "Not Exposed",
    "2424": "Not Exposed",
    "2425": "Not Exposed",
    "2426": "Minimal Exposure",
    "2427": "Not Exposed",
    "2431": "Minimal Exposure",
    "2432": "Not Exposed",
    "2433": "Not Exposed",
    "2434": "Not Exposed",
    "2435": "Minimal Exposure",
    "2436": "Not Exposed",
    "2437": "Minimal Exposure",
    "2439": "Not Exposed",
    "2441": "Not Exposed",
    "2442": "Gradient 2",
    "2443": "Gradient 2",
    "2451": "Minimal Exposure",
    "2452": "Minimal Exposure",
    "2453": "Minimal Exposure",
    "2454": "Gradient 2",
    "2461": "Minimal Exposure",
    "2462": "Not Exposed",
    "2463": "Not Exposed",
    "2464": "Not Exposed",
    "2465": "Minimal Exposure",
    "2466": "Not Exposed",
    "2469": "Not Exposed",
    "2471": "Not Exposed",
    "2472": "Gradient 2",
    "2473": "Gradient 2",
    "2481": "Minimal Exposure",
    "2482": "Minimal Exposure",
    "2483": "Gradient 2",
    "2484": "Gradient 2",
    "2511": "Minimal Exposure",
    "2512": "Minimal Exposure",
    "2513": "Not Exposed",
    "2591": "Minimal Exposure",
    "2592": "Minimal Exposure",
    "2599": "Minimal Exposure",
    "2611": "Gradient 3",
    "2612": "Gradient 3",
    "2613": "Gradient 4",
    "2621": "Gradient 2",
    "2622": "Minimal Exposure",
    "2623": "Minimal Exposure",
    "2624": "Gradient 2",
    "2625": "Gradient 2",
    "2630": "Minimal Exposure",
    "2640": "Gradient 3",
    "2651": "Gradient 3",
    "2652": "Gradient 2",
    "2653": "Gradient 3",
    "2711": "Gradient 2",
    "2712": "Gradient 3",
    "2713": "Gradient 4",
    "2719": "Gradient 3",
    "2721": "Gradient 3",
    "2722": "Gradient 3",
    "2723": "Gradient 3",
    "2729": "Gradient 2",
    "2810": "Gradient 3",
    "2821": "Gradient 2",
    "2822": "Gradient 2",
    "2823": "Gradient 1",
    "2824": "Not Exposed",
    "2825": "Not Exposed",
    "2830": "Not Exposed",
    "2911": "Gradient 2",
    "2912": "Gradient 3",
    "2921": "Gradient 3",
    "2922": "Gradient 3",
    "2923": "Gradient 3",
    "2931": "Not Exposed",
    "2932": "Not Exposed",
    "2933": "Not Exposed",
    "2934": "Minimal Exposure",
    "2935": "Minimal Exposure",
    "2936": "Gradient 2",
    "2937": "Not Exposed",
    "2939": "Not Exposed",
    "3110": "Minimal Exposure",
    "3121": "Not Exposed",
    "3122": "Not Exposed",
    "3123": "Not Exposed",
    "3124": "Minimal Exposure",
    "3125": "Minimal Exposure",
    "3126": "Not Exposed",
    "3127": "Not Exposed",
    "3128": "Not Exposed",
    "3129": "Not Exposed",
    "3131": "Not Exposed",
    "3132": "Not Exposed",
    "3133": "Not Exposed",
    "3134": "Not Exposed",
    "3135": "Minimal Exposure",
    "3141": "Gradient 1",
    "3142": "Not Exposed",
    "3143": "Not Exposed",
    "3151": "Not Exposed",
    "3152": "Not Exposed",
    "3153": "Not Exposed",
    "3154": "Minimal Exposure",
    "3155": "Minimal Exposure",
    "3160": "Not Exposed",
    "3201": "Not Exposed",
    "3202": "Not Exposed",
    "3203": "Minimal Exposure",
    "3204": "Minimal Exposure",
    "3205": "Minimal Exposure",
    "3206": "Minimal Exposure",
    "3207": "Minimal Exposure",
    "3209": "Minimal Exposure",
    "3311": "Not Exposed",
    "3312": "Not Exposed",
    "3313": "Minimal Exposure",
    "3314": "Minimal Exposure",
    "3315": "Not Exposed",
    "3316": "Not Exposed",
    "3317": "Not Exposed",
    "3321": "Not Exposed",
    "3322": "Gradient 3",
    "3323": "Minimal Exposure",
    "3324": "Not Exposed",
    "3325": "Not Exposed",
    "3326": "Not Exposed",
    "3327": "Not Exposed",
    "3329": "Minimal Exposure",
    "3331": "Not Exposed",
    "3339": "Not Exposed",
    "3401": "Gradient 4",
    "3402": "Gradient 4",
    "3403": "Gradient 2",
    "3404": "Gradient 3",
    "3405": "Gradient 2",
    "3510": "Gradient 2",
    "3521": "Gradient 3",
    "3522": "Minimal Exposure",
    "3523": "Gradient 2",
    "3531": "Gradient 3",
    "3532": "Gradient 2",
    "3533": "Minimal Exposure",
    "3534": "Minimal Exposure",
    "3535": "Gradient 2",
    "3539": "Gradient 2",
    "3611": "Gradient 2",
    "3612": "Gradient 3",
    "3613": "Gradient 3",
    "3614": "Gradient 3",
    "3621": "Gradient 2",
    "3622": "Gradient 2",
    "3623": "Gradient 2",
    "3629": "Minimal Exposure",
    "3631": "Not Exposed",
    "3632": "Not Exposed",
    "3711": "Gradient 1",
    "3712": "Gradient 1",
    "3713": "Minimal Exposure",
    "3714": "Minimal Exposure",
    "3715": "Minimal Exposure",
    "3716": "Not Exposed",
    "3721": "Not Exposed",
    "3722": "Minimal Exposure",
    "3723": "Not Exposed",
    "3724": "Not Exposed",
    "3731": "Gradient 1",
    "3732": "Minimal Exposure",
    "3733": "Gradient 1",
    "3734": "Not Exposed",
    "3811": "Gradient 2",
    "3812": "Gradient 2",
    "3813": "Gradient 2",
    "3814": "Gradient 3",
    "3820": "Gradient 2",
    "3831": "Minimal Exposure",
    "3832": "Minimal Exposure",
    "3833": "Minimal Exposure",
    "4111": "Gradient 4",
    "4112": "Gradient 4",
    "4113": "Gradient 4",
    "4121": "Minimal Exposure",
    "4122": "Gradient 2",
    "4123": "Gradient 3",
    "4210": "Gradient 2",
    "4221": "Gradient 2",
    "4222": "Gradient 3",
    "4223": "Gradient 4",
    "4301": "Gradient 4",
    "4309": "Gradient 4",
    "4411": "Gradient 3",
    "4412": "Gradient 3",
    "4421": "Gradient 3",
    "4422": "Gradient 3",
    "4423": "Gradient 3",
    "4424": "Gradient 3",
    "4430": "Gradient 3",
    "4441": "Gradient 3",
    "4442": "Gradient 2",
    "4443": "Gradient 2",
    "4444": "Gradient 2",
    "4445": "Gradient 2",
    "4446": "Gradient 3",
    "4500": "Gradient 2",
    "5000": "Not Exposed",
    "5110": "Not Exposed",
    "5120": "Not Exposed",
    "5210": "Minimal Exposure",
    "5220": "Gradient 1",
    "5300": "Gradient 2",
    "5411": "Gradient 1",
    "5412": "Gradient 1",
    "5420": "Gradient 4",
    "5430": "Minimal Exposure",
    "5491": "Gradient 2",
    "5492": "Gradient 2",
    "5493": "Not Exposed",
    "5500": "Gradient 1",
    "5611": "Not Exposed",
    "5612": "Not Exposed",
    "5621": "Minimal Exposure",
    "5622": "Not Exposed",
    "5629": "Not Exposed",
    "5710": "Minimal Exposure",
    "5721": "Not Exposed",
    "5722": "Not Exposed",
    "5811": "Not Exposed",
    "5812": "Not Exposed",
    "5821": "Not Exposed",
    "5822": "Minimal Exposure",
    "5823": "Minimal Exposure",
    "5824": "Minimal Exposure",
    "5825": "Minimal Exposure",
    "5831": "Not Exposed",
    "5832": "Not Exposed",
    "5833": "Minimal Exposure",
    "5840": "Not Exposed",
    "5891": "Not Exposed",
    "5892": "Not Exposed",
    "5893": "Not Exposed",
    "5894": "Not Exposed",
    "5895": "Minimal Exposure",
    "5899": "Not Exposed",
    "5910": "Not Exposed",
    "5921": "Not Exposed",
    "5922": "Not Exposed",
    "5923": "Not Exposed",
    "5931": "Not Exposed",
    "5932": "Not Exposed",
    "5941": "Not Exposed",
    "5942": "Not Exposed",
    "5991": "Not Exposed",
    "5992": "Not Exposed",
    "5993": "Not Exposed",
    "5999": "Not Exposed",
    "6110": "Not Exposed",
    "6120": "Not Exposed",
    "6201": "Not Exposed",
    "6202": "Not Exposed",
    "6203": "Not Exposed",
    "6204": "Not Exposed",
    "6205": "Not Exposed",
    "6209": "Not Exposed",
    "6300": "Not Exposed",
    "6410": "Not Exposed",
    "6421": "Not Exposed",
    "6422": "Not Exposed",
    "6423": "Not Exposed",
    "6430": "Not Exposed",
    "7111": "Not Exposed",
    "7112": "Not Exposed",
    "7121": "Not Exposed",
    "7122": "Not Exposed",
    "7131": "Not Exposed",
    "7132": "Not Exposed",
    "7191": "Not Exposed",
    "7192": "Not Exposed",
    "7193": "Not Exposed",
    "7199": "Not Exposed",
    "7211": "Not Exposed",
    "7212": "Not Exposed",
    "7221": "Not Exposed",
    "7222": "Not Exposed",
    "7223": "Not Exposed",
    "7231": "Not Exposed",
    "7232": "Not Exposed",
    "7240": "Not Exposed",
    "7250": "Not Exposed",
    "7291": "Not Exposed",
    "7292": "Not Exposed",
    "7293": "Not Exposed",
    "7294": "Not Exposed",
    "7295": "Not Exposed",
    "7311": "Not Exposed",
    "7312": "Not Exposed",
    "7313": "Not Exposed",
    "7314": "Not Exposed",
    "7315": "Not Exposed",
    "7321": "Not Exposed",
    "7322": "Not Exposed",
    "7323": "Not Exposed",
    "7324": "Not Exposed",
    "7401": "Not Exposed",
    "7402": "Not Exposed",
    "7403": "Not Exposed",
    "7404": "Not Exposed",
    "7405": "Not Exposed",
    "7510": "Not Exposed",
    "7521": "Not Exposed",
    "7522": "Not Exposed",
    "7531": "Not Exposed",
    "7532": "Not Exposed",
    "7533": "Not Exposed",
    "7611": "Not Exposed",
    "7612": "Not Exposed",
    "7613": "Not Exposed",
    "7614": "Not Exposed",
    "7615": "Not Exposed",
    "7616": "Not Exposed",
    "7617": "Not Exposed",
    "7618": "Not Exposed",
    "7621": "Gradient 1",
    "7622": "Not Exposed",
    "7623": "Not Exposed",
    "7701": "Not Exposed",
    "7702": "Not Exposed",
    "7703": "Not Exposed",
    "7704": "Not Exposed",
    "7705": "Not Exposed",
    "7706": "Not Exposed",
    "7707": "Not Exposed",
    "7708": "Not Exposed",
    "7709": "Minimal Exposure",
    "7811": "Not Exposed",
    "7812": "Not Exposed",
    "7820": "Not Exposed",
    "7831": "Not Exposed",
    "7832": "Not Exposed",
    "7833": "Not Exposed",
    "7834": "Not Exposed",
    "7835": "Not Exposed",
    "7836": "Not Exposed",
    "7837": "Not Exposed",
    "7891": "Not Exposed",
    "7892": "Not Exposed",
    "7893": "Minimal Exposure",
    "7894": "Not Exposed",
    "7899": "Not Exposed",
    "8111": "Not Exposed",
    "8112": "Not Exposed",
    "8113": "Not Exposed",
    "8114": "Not Exposed",
    "8121": "Not Exposed",
    "8122": "Not Exposed",
    "8131": "Not Exposed",
    "8132": "Not Exposed",
    "8133": "Not Exposed",
    "8141": "Not Exposed",
    "8142": "Not Exposed",
    "8143": "Not Exposed",
    "8144": "Not Exposed",
    "8145": "Minimal Exposure",
    "8151": "Not Exposed",
    "8152": "Not Exposed",
    "8153": "Not Exposed",
    "8154": "Not Exposed",
    "8155": "Not Exposed",
    "8156": "Not Exposed",
    "8159": "Not Exposed",
    "8160": "Not Exposed",
    "8170": "Not Exposed",
    "8191": "Not Exposed",
    "8192": "Not Exposed",
    "8193": "Not Exposed",
    "8201": "Minimal Exposure",
    "8202": "Minimal Exposure",
    "8209": "Minimal Exposure",
    "8311": "Not Exposed",
    "8312": "Not Exposed",
    "8321": "Not Exposed",
    "8322": "Not Exposed",
    "8331": "Not Exposed",
    "8332": "Not Exposed",
    "8333": "Not Exposed",
    "8340": "Not Exposed",
    "8411": "Gradient 1",
    "8412": "Gradient 1",
    "8420": "Not Exposed",
    "8431": "Minimal Exposure",
    "8432": "Minimal Exposure",
    "8440": "Minimal Exposure",
    "9100": "Not Exposed",
    "9210": "Not Exposed",
    "9221": "Not Exposed",
    "9222": "Not Exposed",
    "9223": "Not Exposed",
    "9229": "Not Exposed",
    "9310": "Not Exposed",
    "9320": "Not Exposed",
    "9410": "Not Exposed",
    "9420": "Not Exposed",
    "9431": "Gradient 1",
    "9432": "Gradient 1",
    "9433": "Gradient 1",
    "9434": "Gradient 1",
    "9441": "Not Exposed",
    "9442": "Not Exposed",
    "9443": "Not Exposed",
    "9490": "Not Exposed",
    "9511": "Not Exposed",
    "9512": "Not Exposed",
    "9520": "Not Exposed",
    "9530": "Not Exposed",
    "9541": "Not Exposed",
    "9542": "Not Exposed",
    "9543": "Not Exposed",
    "9601": "Not Exposed",
    "9602": "Not Exposed",
    "9603": "Not Exposed",
    "9700": "Not Exposed",
    "9811": "Not Exposed",
    "9812": "Not Exposed",
    "9820": "Not Exposed",
}

OCUPACIONES_ORDEN = [
    ("5420", "Operadores de telemarketing", "maxima"),
    ("4112", "Empleados de control de personal y nóminas", "maxima"),
    ("4223", "Empleados de servicio de personal", "maxima"),
    ("4301", "Grabadores de datos", "maxima"),
    ("4111", "Empleados de contabilidad", "maxima"),
    ("4113", "Empleados de oficina de servicios estadísticos, financieros y bancarios", "maxima"),
    ("2713", "Analistas, programadores y diseñadores web y multimedia", "maxima"),
    ("3402", "Comerciales de préstamos y créditos", "maxima"),
    ("3401", "Profesionales de apoyo e intermediarios de cambio, bolsa y finanzas", "maxima"),
    ("2613", "Analistas financieros", "maxima"),
    ("4309", "Empleados administrativos sin tareas de atención al público no clasificados bajo otros epígrafes", "maxima"),
    ("3531", "Representantes de aduanas", "significativa"),
    ("2412", "Meteorólogos", "significativa"),
    ("2416", "Estadísticos", "significativa"),
    ("2921", "Escritores", "significativa"),
    ("3814", "Técnicos de la web", "significativa"),
    ("4421", "Empleados de agencias de viajes", "significativa"),
    ("4423", "Telefonistas", "significativa"),
    ("4222", "Codificadores y correctores de imprenta", "significativa"),
    ("4424", "Teleoperadores", "significativa"),
    ("4123", "Empleados de logística y transporte de pasajeros y mercancías", "significativa"),
    ("4430", "Agentes de encuestas", "significativa"),
    ("2810", "Economistas", "significativa"),
    ("2719", "Analistas y diseñadores de software y multimedia no clasificados bajo otros epígrafes", "significativa"),
    ("2721", "Diseñadores y administradores de bases de datos", "significativa"),
    ("2912", "Bibliotecarios, documentalistas y afines", "significativa"),
    ("2922", "Periodistas", "significativa"),
    ("4422", "Recepcionistas de hoteles", "significativa"),
    ("2415", "Matemáticos y actuarios", "significativa"),
    ("3613", "Asistentes de dirección y administrativos", "significativa"),
    ("3521", "Mediadores y agentes de seguros", "significativa"),
    ("3322", "Técnicos superiores en documentación sanitaria", "significativa"),
    ("4441", "Cajeros de bancos y afines", "significativa"),
    ("2611", "Especialistas en contabilidad", "significativa"),
    ("4412", "Recepcionistas (excepto de hoteles)", "significativa"),
    ("3612", "Asistentes jurídico-legales", "significativa"),
    ("2651", "Profesionales de la publicidad y la comercialización", "significativa"),
    ("2640", "Profesionales de ventas técnicas y médicas (excepto las TIC)", "significativa"),
    ("4411", "Empleados de información al usuario", "significativa"),
    ("2923", "Filólogos, intérpretes y traductores", "significativa"),
    ("2612", "Asesores financieros y en inversiones", "significativa"),
    ("3404", "Profesionales de apoyo en servicios estadísticos, matemáticos y afines", "significativa"),
    ("2712", "Analistas y diseñadores de software", "significativa"),
    ("2723", "Analistas de redes informáticas", "significativa"),
    ("2722", "Administradores de sistemas y redes", "significativa"),
    ("4446", "Empleados de mostrador de correos", "significativa"),
    ("2653", "Profesionales de la venta de tecnologías de la información y las comunicaciones", "significativa"),
    ("3614", "Secretarios de centros médicos o clínicas", "significativa"),
]

TIER_META = [
    # (key, label, color, short_label-para-frases-en-linea)
    ("maxima", "Exposición máxima (Gradiente 4 de la OIT)", "#FF5A45", "exposición máxima"),
    ("significativa", "Exposición significativa (Gradiente 3 de la OIT)", "#FF8A4C", "exposición significativa"),
    ("moderada", "Exposición moderada", "#F2B84C", "exposición moderada"),
    ("baja", "Exposición baja", "#B9D24C", "exposición baja"),
    ("minima", "Exposición mínima", "#6FC9A0", "exposición mínima"),
    ("sinexp", "Sin exposición", "#33C2A6", "sin exposición"),
]

# Texto literal de la columna "Exposure" del original -> clave de nivel
EXPOSURE_TO_TIER = {
    "Gradient 4": "maxima",
    "Gradient 3": "significativa",
    "Gradient 2": "moderada",
    "Gradient 1": "baja",
    "Minimal Exposure": "minima",
    "Not Exposed": "sinexp",
}

# ---------------------------------------------------------------------------
# 1bis. Índice OCDE — DE PRUEBA. Roberto todavía no tiene las categorías/columna
#   reales de la OCDE ("aun no los conozco, ya te los diré"). Mientras tanto,
#   para poder construir y probar de verdad el selector de índice de exposición
#   (esquema de datos + web), se genera aquí una clasificación ALEATORIA de las
#   502 ocupaciones en 5 niveles (semilla fija -> mismo reparto cada vez que se
#   regenera el JSON). SUSTITUIR build_ocde_nivel_ficticio() por la lectura de
#   la columna/categoría real en cuanto Roberto la facilite.
# ---------------------------------------------------------------------------

OCDE_TIER_META = [
    # Reparto pedido: 15 en el nivel 5 (máxima), 50 en el 4, y cada vez más
    # ocupaciones cuanto menor es el nivel de exposición.
    ("nivel5", "Nivel 5 de exposición OCDE (ficticio, de prueba)", "#FF5A45", "nivel 5 de la OCDE (ficticio)"),
    ("nivel4", "Nivel 4 de exposición OCDE (ficticio, de prueba)", "#FF8A4C", "nivel 4 de la OCDE (ficticio)"),
    ("nivel3", "Nivel 3 de exposición OCDE (ficticio, de prueba)", "#F2B84C", "nivel 3 de la OCDE (ficticio)"),
    ("nivel2", "Nivel 2 de exposición OCDE (ficticio, de prueba)", "#B9D24C", "nivel 2 de la OCDE (ficticio)"),
    ("nivel1", "Nivel 1 de exposición OCDE (ficticio, de prueba)", "#33C2A6", "nivel 1 de la OCDE (ficticio)"),
]
OCDE_NIVEL_COUNTS = [("nivel5", 15), ("nivel4", 50), ("nivel3", 90), ("nivel2", 140), ("nivel1", 207)]
OCDE_SEMILLA_FICTICIA = 20261003  # fecha de generación: reparto reproducible, no cambia entre ejecuciones


def build_ocde_nivel_ficticio(all_codes):
    """Reparte aleatoriamente (semilla fija) todas las ocupaciones recibidas
    entre los 5 niveles ficticios de OCDE_NIVEL_COUNTS. Puramente temporal."""
    import random
    codes = sorted(all_codes)
    rnd = random.Random(OCDE_SEMILLA_FICTICIA)
    rnd.shuffle(codes)
    niveles = {}
    i = 0
    for nivel_key, n in OCDE_NIVEL_COUNTS:
        for code in codes[i:i + n]:
            niveles[code] = nivel_key
        i += n
    return niveles


# ---------------------------------------------------------------------------
# 1ter. Correspondencia CNO11 (4 dígitos) -> ISCO/CIUO08 (4 dígitos), para
#   cuando se implemente JRC (que clasifica a 3 dígitos de ISCO). Todavía NO
#   se usa en ningún cálculo de este script (JRC sigue inerte en la web) —
#   se deja aquí guardada, ya resuelta, para no perder el trabajo de validarla.
#   8 ocupaciones no tenían una correspondencia 1:1 real entre ambas
#   clasificaciones (CNO más agregado que ISCO, o al revés); resueltas junto
#   con Roberto:
#     4309 -> 4419 (441), 2159 -> 2269 (226), 3331 -> 2230 (223),
#     5000 -> 5131 (513), 5120 -> 5131 (513), 5910 -> 5412 (541),
#     5921 -> 5412 (541). (El resto de ambigüedades compartían el mismo
#     prefijo ISCO a 3 dígitos y no necesitaban decisión.)
#   CNO "ZZZZ" ("Indeterminado", fila corrupta de P Contratadas por CNO mes) no tiene
#   entrada: no es una ocupación real.
# ---------------------------------------------------------------------------

CNO_TO_ISCO4 = {
    "0011": "0110",
    "0012": "0210",
    "0020": "0310",
    "1111": "1111",
    "1112": "1112",
    "1113": "1114",
    "1120": "1120",
    "1211": "1211",
    "1212": "1212",
    "1219": "1213",
    "1221": "1221",
    "1222": "1222",
    "1223": "1223",
    "1311": "1311",
    "1312": "1312",
    "1313": "1321",
    "1314": "1322",
    "1315": "1324",
    "1316": "1323",
    "1321": "1330",
    "1322": "1341",
    "1323": "1342",
    "1324": "1343",
    "1325": "1344",
    "1326": "1345",
    "1327": "1346",
    "1329": "1349",
    "1411": "1411",
    "1419": "1411",
    "1421": "1412",
    "1422": "1412",
    "1429": "1412",
    "1431": "1420",
    "1432": "1420",
    "1501": "1431",
    "1509": "1439",
    "2111": "2211",
    "2112": "2212",
    "2121": "2221",
    "2122": "2221",
    "2123": "2222",
    "2130": "2250",
    "2140": "2262",
    "2151": "2261",
    "2152": "2264",
    "2153": "2265",
    "2154": "2266",
    "2155": "2267",
    "2156": "2269",
    "2157": "2269",
    "2158": "2263",
    "2159": "2269",
    "2210": "2310",
    "2220": "2320",
    "2230": "2330",
    "2240": "2341",
    "2251": "2342",
    "2252": "2342",
    "2311": "2352",
    "2312": "2352",
    "2321": "2351",
    "2322": "2353",
    "2323": "2354",
    "2324": "2355",
    "2325": "2356",
    "2326": "2359",
    "2329": "2359",
    "2411": "2111",
    "2412": "2112",
    "2413": "2113",
    "2414": "2114",
    "2415": "2120",
    "2416": "2120",
    "2421": "2131",
    "2422": "2132",
    "2423": "2132",
    "2424": "2132",
    "2425": "2132",
    "2426": "2133",
    "2427": "2132",
    "2431": "2141",
    "2432": "2142",
    "2433": "2144",
    "2434": "2144",
    "2435": "2145",
    "2436": "2146",
    "2437": "2143",
    "2439": "2149",
    "2441": "2151",
    "2442": "2152",
    "2443": "2153",
    "2451": "2161",
    "2452": "2162",
    "2453": "2164",
    "2454": "2165",
    "2461": "2141",
    "2462": "2142",
    "2463": "2144",
    "2464": "2144",
    "2465": "2145",
    "2466": "2146",
    "2469": "2149",
    "2471": "2151",
    "2472": "2152",
    "2473": "2153",
    "2481": "2161",
    "2482": "2163",
    "2483": "2165",
    "2484": "2166",
    "2511": "2611",
    "2512": "2611",
    "2513": "2612",
    "2591": "2619",
    "2592": "2619",
    "2599": "2619",
    "2611": "2411",
    "2612": "2412",
    "2613": "2413",
    "2621": "2421",
    "2622": "2422",
    "2623": "2422",
    "2624": "2423",
    "2625": "2424",
    "2630": "2422",
    "2640": "2433",
    "2651": "2431",
    "2652": "2432",
    "2653": "2434",
    "2711": "2511",
    "2712": "2512",
    "2713": "2513",
    "2719": "2519",
    "2721": "2521",
    "2722": "2522",
    "2723": "2523",
    "2729": "2529",
    "2810": "2631",
    "2821": "2632",
    "2822": "2633",
    "2823": "2634",
    "2824": "2635",
    "2825": "2635",
    "2830": "2636",
    "2911": "2621",
    "2912": "2622",
    "2921": "2641",
    "2922": "2642",
    "2923": "2643",
    "2931": "2651",
    "2932": "2652",
    "2933": "2653",
    "2934": "2654",
    "2935": "2655",
    "2936": "2656",
    "2937": "2659",
    "2939": "2659",
    "3110": "3118",
    "3121": "3111",
    "3122": "3112",
    "3123": "3113",
    "3124": "3114",
    "3125": "3114",
    "3126": "3115",
    "3127": "3116",
    "3128": "3117",
    "3129": "3119",
    "3131": "3131",
    "3132": "3132",
    "3133": "3133",
    "3134": "3134",
    "3135": "3135",
    "3139": "3139",
    "3141": "3141",
    "3142": "3142",
    "3143": "3143",
    "3151": "3151",
    "3152": "3152",
    "3153": "3153",
    "3154": "3154",
    "3155": "3155",
    "3160": "3111",
    "3201": "3121",
    "3202": "3123",
    "3203": "3122",
    "3204": "3122",
    "3205": "3122",
    "3206": "3122",
    "3207": "3122",
    "3209": "3122",
    "3311": "3211",
    "3312": "3211",
    "3313": "3212",
    "3314": "3212",
    "3315": "3214",
    "3316": "3214",
    "3317": "3214",
    "3321": "3251",
    "3322": "3252",
    "3323": "3259",
    "3324": "3254",
    "3325": "3255",
    "3326": "3257",
    "3327": "3240",
    "3329": "3259",
    "3331": "2230",
    "3339": "3230",
    "3401": "3311",
    "3402": "3312",
    "3403": "3313",
    "3404": "3314",
    "3405": "3315",
    "3510": "3322",
    "3521": "3321",
    "3522": "3323",
    "3523": "3324",
    "3531": "3331",
    "3532": "3332",
    "3533": "3333",
    "3534": "3334",
    "3535": "3339",
    "3539": "3339",
    "3611": "3341",
    "3612": "3342",
    "3613": "3343",
    "3614": "3344",
    "3621": "3352",
    "3622": "3353",
    "3623": "3354",
    "3629": "3359",
    "3631": "3355",
    "3632": "3355",
    "3711": "3411",
    "3712": "3411",
    "3713": "3412",
    "3714": "3412",
    "3715": "3412",
    "3716": "3413",
    "3721": "3421",
    "3722": "3422",
    "3723": "3423",
    "3724": "3423",
    "3731": "3431",
    "3732": "3432",
    "3733": "3433",
    "3734": "3434",
    "3739": "3435",
    "3811": "3511",
    "3812": "3512",
    "3813": "3513",
    "3814": "3514",
    "3820": "3512",
    "3831": "3521",
    "3832": "3521",
    "3833": "3522",
    "4111": "4311",
    "4112": "4313",
    "4113": "4312",
    "4121": "4321",
    "4122": "4322",
    "4123": "4323",
    "4210": "4411",
    "4221": "4412",
    "4222": "4413",
    "4223": "4416",
    "4301": "4131",
    "4309": "4419",
    "4411": "4225",
    "4412": "4226",
    "4421": "4221",
    "4422": "4224",
    "4423": "4223",
    "4424": "4222",
    "4430": "4227",
    "4441": "4211",
    "4442": "4212",
    "4443": "4212",
    "4444": "4213",
    "4445": "4214",
    "4446": "4211",
    "4500": "4229",
    "5000": "5131",
    "5110": "5120",
    "5120": "5131",
    "5210": "5222",
    "5220": "5223",
    "5300": "5221",
    "5411": "5211",
    "5412": "5211",
    "5420": "5244",
    "5430": "5245",
    "5491": "5243",
    "5492": "5242",
    "5493": "5241",
    "5499": "5249",
    "5500": "5230",
    "5611": "5321",
    "5612": "5321",
    "5621": "3213",
    "5622": "3258",
    "5629": "5329",
    "5710": "5322",
    "5721": "5311",
    "5722": "5311",
    "5811": "5141",
    "5812": "5142",
    "5821": "5111",
    "5822": "5112",
    "5823": "5113",
    "5824": "5113",
    "5825": "5113",
    "5831": "5151",
    "5832": "5152",
    "5833": "5153",
    "5840": "5152",
    "5891": "5162",
    "5892": "5163",
    "5893": "5164",
    "5894": "5165",
    "5895": "5161",
    "5899": "5169",
    "5910": "5412",
    "5921": "5412",
    "5922": "5412",
    "5923": "5412",
    "5931": "5411",
    "5932": "5411",
    "5941": "5414",
    "5942": "5414",
    "5991": "5413",
    "5992": "5419",
    "5993": "5419",
    "5999": "5419",
    "6110": "6111",
    "6120": "6113",
    "6201": "6121",
    "6202": "6121",
    "6203": "6121",
    "6204": "6123",
    "6205": "6122",
    "6209": "6121",
    "6300": "6130",
    "6410": "6210",
    "6421": "6221",
    "6422": "6222",
    "6423": "6223",
    "6430": "6224",
    "7111": "7114",
    "7112": "7114",
    "7121": "7111",
    "7122": "7113",
    "7131": "7115",
    "7132": "7115",
    "7191": "7119",
    "7192": "7119",
    "7193": "7119",
    "7199": "7119",
    "7211": "7123",
    "7212": "7123",
    "7221": "7126",
    "7222": "7126",
    "7223": "7126",
    "7231": "7131",
    "7232": "7131",
    "7240": "7122",
    "7250": "7127",
    "7291": "7121",
    "7292": "7124",
    "7293": "7125",
    "7294": "7126",
    "7295": "7133",
    "7311": "7211",
    "7312": "7212",
    "7313": "7213",
    "7314": "7214",
    "7315": "7215",
    "7321": "7221",
    "7322": "7222",
    "7323": "7223",
    "7324": "7224",
    "7401": "7231",
    "7402": "7232",
    "7403": "7233",
    "7404": "7233",
    "7405": "7234",
    "7510": "7411",
    "7521": "7412",
    "7522": "7413",
    "7531": "7421",
    "7532": "7421",
    "7533": "7422",
    "7611": "7311",
    "7612": "7312",
    "7613": "7313",
    "7614": "7314",
    "7615": "7315",
    "7616": "7316",
    "7617": "7317",
    "7618": "7318",
    "7619": "7319",
    "7621": "7321",
    "7622": "7322",
    "7623": "7323",
    "7701": "7511",
    "7702": "7511",
    "7703": "7512",
    "7704": "7513",
    "7705": "7514",
    "7706": "7514",
    "7707": "7514",
    "7708": "7516",
    "7709": "7515",
    "7811": "7521",
    "7812": "7523",
    "7820": "7522",
    "7831": "7531",
    "7832": "7532",
    "7833": "7532",
    "7834": "7533",
    "7835": "7534",
    "7836": "7535",
    "7837": "7536",
    "7891": "7541",
    "7892": "7542",
    "7893": "7543",
    "7894": "7544",
    "7899": "7549",
    "8111": "8111",
    "8112": "8112",
    "8113": "8113",
    "8114": "8114",
    "8121": "8121",
    "8122": "8122",
    "8131": "8131",
    "8132": "8131",
    "8133": "8132",
    "8141": "8141",
    "8142": "8142",
    "8143": "8143",
    "8144": "8172",
    "8145": "8171",
    "8151": "8151",
    "8152": "8152",
    "8153": "8153",
    "8154": "8154",
    "8155": "8155",
    "8156": "8156",
    "8159": "8159",
    "8160": "8160",
    "8170": "8157",
    "8191": "8181",
    "8192": "8182",
    "8193": "8183",
    "8199": "8189",
    "8201": "8211",
    "8202": "8212",
    "8209": "8219",
    "8311": "8311",
    "8312": "8312",
    "8321": "8341",
    "8322": "8341",
    "8331": "8342",
    "8332": "8343",
    "8333": "8344",
    "8340": "8350",
    "8411": "8322",
    "8412": "8322",
    "8420": "8331",
    "8431": "8332",
    "8432": "8332",
    "8440": "8321",
    "9100": "9111",
    "9210": "9112",
    "9221": "9121",
    "9222": "9122",
    "9223": "9123",
    "9229": "9129",
    "9310": "9412",
    "9320": "9411",
    "9410": "9520",
    "9420": "9510",
    "9431": "9621",
    "9432": "9621",
    "9433": "9621",
    "9434": "9623",
    "9441": "9611",
    "9442": "9612",
    "9443": "9613",
    "9490": "9622",
    "9511": "9211",
    "9512": "9214",
    "9520": "9212",
    "9530": "9213",
    "9541": "9216",
    "9542": "9216",
    "9543": "9215",
    "9601": "9312",
    "9602": "9313",
    "9603": "9311",
    "9700": "9321",
    "9811": "9331",
    "9812": "9331",
    "9820": "9334",
}

# ---------------------------------------------------------------------------
# 1quater. Índice JRC — REAL (del docx que pasó Roberto). Clasifica a 3
#   dígitos de ISCO, en 3 categorías (Alta/Media/Baja), sin relación con los
#   6 niveles de la OIT ni los 5 (ficticios) de la OCDE: cada índice tiene su
#   propia estructura, no hay que hacer ninguna correspondencia entre ellas.
#   Solo "Alta" se despliega (las otras dos, como el resto de niveles no
#   desplegables de los otros índices, solo muestran su dato agregado).
#   Para clasificar una ocupación CNO: CNO_TO_ISCO4[cno][:3] (3 primeros
#   dígitos del ISCO de 4) -> JRC_ISCO3_TO_NIVEL. Los códigos de las fuerzas
#   armadas (0011/0012/0020) no tienen puntuación JRC (no están en la tabla
#   del docx) y quedan sin clasificar, igual que los 9 "=#N/A" de la OIT.
# ---------------------------------------------------------------------------

JRC_TIER_META = [
    ("alta", "Alta exposición (JRC)", "#FF5A45", "alta exposición (JRC)"),
    ("media", "Media exposición (JRC)", "#F2B84C", "media exposición (JRC)"),
    ("baja", "Baja exposición (JRC)", "#33C2A6", "baja exposición (JRC)"),
]

JRC_ISCO3_TO_NIVEL = {
    "111": "alta", "112": "alta", "121": "alta", "122": "alta", "131": "alta",
    "132": "alta", "133": "alta", "134": "alta", "141": "media", "142": "media",
    "143": "alta", "211": "media", "212": "alta", "213": "alta", "214": "alta",
    "215": "alta", "216": "alta", "221": "alta", "222": "alta", "223": "alta",
    "224": "alta", "225": "alta", "226": "alta", "231": "alta", "232": "media",
    "233": "media", "234": "media", "235": "media", "241": "alta", "242": "alta",
    "243": "alta", "251": "alta", "252": "media", "261": "media", "262": "alta",
    "263": "alta", "264": "media", "265": "media", "311": "media", "312": "media",
    "313": "media", "314": "media", "315": "alta", "321": "media", "322": "media",
    "323": "media", "324": "media", "325": "media", "331": "alta", "332": "alta",
    "333": "alta", "334": "alta", "335": "alta", "341": "media", "342": "media",
    "343": "media", "351": "media", "352": "media", "411": "media", "412": "media",
    "413": "media", "421": "media", "422": "media", "431": "media", "432": "media",
    "441": "baja", "511": "media", "512": "baja", "513": "baja", "514": "baja",
    "515": "media", "516": "baja", "521": "baja", "522": "baja", "523": "baja",
    "524": "baja", "531": "media", "532": "baja", "541": "media", "611": "media",
    "612": "baja", "613": "media", "621": "baja", "622": "baja", "631": "baja",
    "632": "baja", "633": "baja", "634": "baja", "711": "baja", "712": "baja",
    "713": "baja", "721": "baja", "722": "baja", "723": "media", "731": "baja",
    "732": "baja", "741": "media", "742": "media", "751": "baja", "752": "baja",
    "753": "baja", "754": "baja", "811": "baja", "812": "baja", "813": "baja",
    "814": "baja", "815": "baja", "816": "baja", "817": "baja", "818": "baja",
    "821": "baja", "831": "baja", "832": "media", "833": "baja", "834": "baja",
    "835": "baja", "911": "baja", "912": "baja", "921": "baja", "931": "baja",
    "932": "baja", "933": "baja", "941": "baja", "951": "baja", "952": "baja",
    "961": "baja", "962": "baja",
}


def build_cno_to_jrc_nivel(cno_to_isco4):
    """CNO -> 'alta'/'media'/'baja' a partir de CNO_TO_ISCO4[:3] + la tabla
    JRC (a 3 dígitos). Las ocupaciones cuyo ISCO no está en la tabla JRC
    (las 3 de fuerzas armadas) quedan fuera -> sin clasificar."""
    out = {}
    for cno, isco4 in cno_to_isco4.items():
        nivel = JRC_ISCO3_TO_NIVEL.get(isco4[:3])
        if nivel:
            out[cno] = nivel
    return out


# Nombre en INGLÉS (tal cual viene del docx de JRC) de cada grupo ISCO a 3 dígitos.
# Se conserva para el futuro selector ES/EN de la página.
JRC_ISCO3_NAME_EN = {
    "111": "Legislators and senior officials",
    "112": "Managing directors and chief executives",
    "121": "Business services and administration managers",
    "122": "Sales, marketing and development managers",
    "131": "Production managers in agriculture, forestry and fisheries",
    "132": "Manufacturing, mining, construction, and distribution managers",
    "133": "Information and communications technology service managers",
    "134": "Professional services managers",
    "141": "Hotel and restaurant managers",
    "142": "Retail and wholesale trade managers",
    "143": "Other services managers",
    "211": "Physical and earth science professionals",
    "212": "Mathematicians, actuaries and statisticians",
    "213": "Life science professionals",
    "214": "Engineering professionals (excluding electrotechnology)",
    "215": "Electrotechnology engineers",
    "216": "Architects, planners, surveyors and designers",
    "221": "Medical doctors",
    "222": "Nursing and midwifery professionals",
    "223": "Traditional and complementary medicine professionals",
    "224": "Paramedical practitioners",
    "225": "Veterinarians",
    "226": "Other health professionals",
    "231": "University and higher education teachers",
    "232": "Vocational education teachers",
    "233": "Secondary education teachers",
    "234": "Primary school and early childhood teachers",
    "235": "Other teaching professionals",
    "241": "Finance professionals",
    "242": "Administration professionals",
    "243": "Sales, marketing and public relations professionals",
    "251": "Software and applications developers and analysts",
    "252": "Database and network professionals",
    "261": "Legal professionals",
    "262": "Librarians, archivists and curators",
    "263": "Social and religious professionals",
    "264": "Authors, journalists and linguists",
    "265": "Creative and performing artists",
    "311": "Physical and engineering science technicians",
    "312": "Mining, manufacturing and construction supervisors",
    "313": "Process control technicians",
    "314": "Life science technicians and related associate professionals",
    "315": "Ship and aircraft controllers and technicians",
    "321": "Medical and pharmaceutical technicians",
    "322": "Nursing and midwifery associate professionals",
    "323": "Traditional and complementary medicine associate professionals",
    "324": "Veterinary technicians and assistants",
    "325": "Other health associate professionals",
    "331": "Financial and mathematical associate professionals",
    "332": "Sales and purchasing agents and brokers",
    "333": "Business services agents",
    "334": "Administrative and specialised secretaries",
    "335": "Regulatory government associate professionals",
    "341": "Legal, social and religious associate professionals",
    "342": "Sports and fitness workers",
    "343": "Artistic, cultural and culinary associate professionals",
    "351": "ICT operations and user support technicians",
    "352": "Telecommunications and broadcasting technicians",
    "411": "General office clerks",
    "412": "Secretaries (general)",
    "413": "Keyboard operators",
    "421": "Tellers, money collectors and related clerks",
    "422": "Client information workers",
    "431": "Numerical clerks",
    "432": "Material-recording and transport clerks",
    "441": "Other clerical support workers",
    "511": "Travel attendants, conductors and guides",
    "512": "Cooks",
    "513": "Waiters and bartenders",
    "514": "Hairdressers, beauticians and related workers",
    "515": "Building and housekeeping supervisors",
    "516": "Other personal services workers",
    "521": "Street and market salespersons",
    "522": "Shop salespersons",
    "523": "Cashiers and ticket clerks",
    "524": "Other sales workers",
    "531": "Child care workers and teachers’ aides",
    "532": "Personal care workers in health services",
    "541": "Protective services workers",
    "611": "Market gardeners and crop growers",
    "612": "Animal producers",
    "613": "Mixed crop and animal producers",
    "621": "Forestry and related workers",
    "622": "Fishery workers, hunters and trappers",
    "631": "Subsistence crop farmers",
    "632": "Subsistence livestock farmers",
    "633": "Subsistence mixed crop and livestock farmers",
    "634": "Subsistence fishers, hunters, trappers and gatherers",
    "711": "Building frame and related trades workers",
    "712": "Building finishers and related trades workers",
    "713": "Painters, building structure cleaners and related trades workers",
    "721": "Sheet and structural metal workers, moulders and welders",
    "722": "Blacksmiths, toolmakers and related trades workers",
    "723": "Machinery mechanics and repairers",
    "731": "Handicraft workers",
    "732": "Printing trades workers",
    "741": "Electrical equipment installers and repairers",
    "742": "Electronics and telecommunications installers and repairers",
    "751": "Food processing and related trades workers",
    "752": "Wood treaters, cabinet-makers and related trades workers",
    "753": "Garment and related trades workers",
    "754": "Other craft and related workers",
    "811": "Mining and mineral processing plant operators",
    "812": "Metal processing and finishing plant operators",
    "813": "Chemical and photographic products plant and machine operators",
    "814": "Rubber, plastic and paper products machine operators",
    "815": "Textile, fur and leather products machine operators",
    "816": "Food and related products machine operators",
    "817": "Wood processing and papermaking plant operators",
    "818": "Other stationary plant and machine operators",
    "821": "Assemblers",
    "831": "Locomotive engine drivers and related workers",
    "832": "Car, van and motorcycle drivers",
    "833": "Heavy truck and bus drivers",
    "834": "Mobile plant operators",
    "835": "Ships’ deck crews and related workers",
    "911": "Domestic, hotel and office cleaners and helpers",
    "912": "Vehicle, window, laundry and other hand cleaning workers",
    "921": "Agricultural, forestry and fishery labourers",
    "931": "Mining and construction labourers",
    "932": "Manufacturing labourers",
    "933": "Transport and storage labourers",
    "941": "Food preparation assistants",
    "951": "Street and related service workers",
    "952": "Street vendors (excluding food)",
    "961": "Refuse workers",
    "962": "Other elementary workers",
}


# Nombre en ESPAÑOL (CIUO-08, fichero aportado por Roberto) de cada grupo ISCO a
# 3 dígitos. Se usa para mostrar las filas del desplegable de JRC: a diferencia de
# la OIT/OCDE (ocupaciones CNO de 4 dígitos), JRC clasifica a 3 dígitos de ISCO,
# así que se agregan TODAS las CNO cuyo ISCO comparte esos 3 dígitos y se muestra
# una fila por grupo ISCO, no por CNO.
JRC_ISCO3_NAME = {
    "111": "Miembros del poder ejecutivo y de los cuerpos legislativos",
    "112": "Directores generales y gerentes generales",
    "121": "Directores de administración y servicios",
    "122": "Directores de ventas, comercialización y desarrollo",
    "131": "Directores de producción agropecuaria, silvicultura y pesca",
    "132": "Directores de industrias manufactureras, de minería, construcción y distribución",
    "133": "Directores de servicios de tecnología de la información y las comunicaciones",
    "134": "Directores y gerentes de servicios profesionales",
    "141": "Gerentes de hoteles y restaurantes",
    "142": "Gerentes de comercios al por mayor y al por menor",
    "143": "Otros gerentes de servicios",
    "211": "Físicos, químicos y afines",
    "212": "Matemáticos, actuarios y estadísticos",
    "213": "Profesionales en ciencias biológicas",
    "214": "Ingenieros (excluyendo electrotecnólogos)",
    "215": "Ingenieros en electrotecnología",
    "216": "Arquitectos, urbanistas, agrimensores y diseñadores",
    "221": "Médicos",
    "222": "Profesionales de enfermería y partería",
    "223": "Profesionales de medicina tradicional y alternativa",
    "224": "Practicantes paramédicos",
    "225": "Veterinarios",
    "226": "Otros profesionales de la salud",
    "231": "Profesores de universidades y de la enseñanza superior",
    "232": "Profesores de formación profesional",
    "233": "Profesores de enseñanza secundaria",
    "234": "Maestros de enseñanza primaria y maestros preescolares",
    "235": "Otros profesionales de la enseñanza",
    "241": "Especialistas en finanzas",
    "242": "Especialistas en organización de administración",
    "243": "Profesionales de las ventas, la comercialización y las relaciones públicas",
    "251": "Desarrolladores y analistas de software y multimedia",
    "252": "Especialistas en bases de datos y en redes de computadores",
    "261": "Profesionales en derecho",
    "262": "Archivistas, bibliotecarios, curadores y afines",
    "263": "Especialistas en ciencias sociales y teología",
    "264": "Autores, periodistas y lingüistas",
    "265": "Artistas creativos e interpretativos",
    "311": "Técnicos en ciencias físicas y en ingeniería",
    "312": "Supervisores en ingeniería de minas, de industrias manufactureras y de la construcción",
    "313": "Técnicos en control de procesos",
    "314": "Técnicos y profesionales de nivel medio en ciencias biológicas y afines",
    "315": "Técnicos y controladores en navegación marítima y aeronáutica",
    "321": "Técnicos médicos y farmacéuticos",
    "322": "Profesionales de nivel medio de enfermería y partería",
    "323": "Profesionales de nivel medio de medicina tradicional y alternativa",
    "324": "Técnicos y asistentes veterinarios",
    "325": "Otros profesionales de nivel medio de la salud",
    "331": "Profesionales de nivel medio en finanzas y matemáticas",
    "332": "Agentes comerciales y corredores",
    "333": "Agentes de servicios comerciales",
    "334": "Secretarios administrativos y especializados",
    "335": "Agentes de la administración pública para la aplicación de la ley y afines",
    "341": "Profesionales de nivel medio, de servicios jurídicos, sociales y religiosos",
    "342": "Entrenadores de deportes y aptitud física",
    "343": "Profesionales de nivel medio en actividades culturales, artísticas y culinarias",
    "351": "Técnicos en operaciones de tecnología de la información y las comunicaciones y asistencia al usuario",
    "352": "Técnicos en telecomunicaciones y radiodifusión",
    "411": "Oficinistas generales",
    "412": "Secretarios (general)",
    "413": "Operadores de máquinas de oficina",
    "421": "Pagadores y cobradores de ventanilla y afines",
    "422": "Empleados de servicios de información al cliente",
    "431": "Auxiliares contables y financieros",
    "432": "Empleados encargados del registro de materiales y de transportes",
    "441": "Otro personal de apoyo administrativo",
    "511": "Personal al servicio directo de los pasajeros",
    "512": "Cocineros",
    "513": "Camareros",
    "514": "Peluqueros, especialistas en tratamientos de belleza y afines",
    "515": "Supervisores de mantenimiento y limpieza de edificios",
    "516": "Otros trabajadores de servicios personales",
    "521": "Vendedores callejeros y de puestos de mercado",
    "522": "Comerciantes y vendedores de tiendas y almacenes",
    "523": "Cajeros y expendedores de billetes",
    "524": "Otros vendedores",
    "531": "Cuidadores de niños y auxiliares de maestros",
    "532": "Trabajadores de los cuidados personales en servicios de salud",
    "541": "Personal de los servicios de protección",
    "611": "Agricultores y trabajadores calificados de jardines y de cultivos para el mercado",
    "612": "Criadores y trabajadores pecuarios calificados de la cría de animales para el mercado y afines",
    "613": "Productores y trabajadores calificados de explotaciones agropecuarias mixtas cuya producción se destina al mercado",
    "621": "Trabajadores forestales calificados y afines",
    "622": "Pescadores, cazadores y tramperos",
    "631": "Trabajadores agrícolas de subsistencia",
    "632": "Trabajadores pecuarios de subsistencia",
    "633": "Trabajadores agropecuarios de subsistencia",
    "634": "Pescadores, cazadores, tramperos y recolectores de subsistencia",
    "711": "Oficiales y operarios de la construcción (obra gruesa) y afines",
    "712": "Oficiales y operarios de la construcción (trabajos de acabado) y afines",
    "713": "Pintores, limpiadores de fachadas y afines",
    "721": "Moldeadores, soldadores, chapistas, caldereros, montadores de estructuras metálicas y afines",
    "722": "Herreros, herramentistas y afines",
    "723": "Mecánicos y reparadores de máquinas",
    "731": "Artesanos",
    "732": "Oficiales y operarios de las artes gráficas",
    "741": "Instaladores y reparadores de equipos eléctricos",
    "742": "Instaladores y reparadores de equipos electrónicos y de telecomunicaciones",
    "751": "Oficiales y operarios de procesamiento de alimentos y afines",
    "752": "Oficiales y operarios del tratamiento de la madera, ebanistas y afines",
    "753": "Oficiales y operarios de la confección y afines",
    "754": "Otros oficiales, operarios y artesanos de artes mecánicas y de otros oficios",
    "811": "Operadores de instalaciones mineras y de extracción y procesamiento de minerales",
    "812": "Operadores de instalaciones de procesamiento y recubridoras de metales",
    "813": "Operadores de instalaciones y máquinas de productos químicos y fotográficos",
    "814": "Operadores de máquinas para fabricar productos de caucho, de papel y de material plástico",
    "815": "Operadores de máquinas para fabricar productos textiles y artículos de piel y cuero",
    "816": "Operadores de máquinas para elaborar alimentos y productos afines",
    "817": "Operadores de instalaciones para la preparación de papel y de procesamiento de la madera",
    "818": "Otros operadores de máquinas y de instalaciones fijas",
    "821": "Ensambladores",
    "831": "Maquinistas de locomotoras y afines",
    "832": "Conductores de automóviles, camionetas y motocicletas",
    "833": "Conductores de camiones pesados y autobuses",
    "834": "Operadores de equipos pesados móviles",
    "835": "Marineros de cubierta y afines",
    "911": "Limpiadores y asistentes domésticos de hoteles y oficinas",
    "912": "Limpiadores de vehículos, ventanas, ropa y otra limpieza a mano",
    "921": "Peones agropecuarios, pesqueros y forestales",
    "931": "Peones de la minería y la construcción",
    "932": "Peones de la industria manufacturera",
    "933": "Peones del transporte y almacenamiento",
    "941": "Ayudantes de preparación de alimentos",
    "951": "Trabajadores ambulantes de servicios y afines",
    "952": "Vendedores ambulantes (excluyendo de comida)",
    "961": "Recolectores de desechos",
    "962": "Otras ocupaciones elementales",
}


# ---------------------------------------------------------------------------
# Índice ANTHROPIC (exposición observada a la IA, continuo 0-1) por CNO a 4 dígitos.
#   Fuente: "indice_anthropic_por_CNO_via_CIUO.xlsx" (hoja Indice_Anthropic_CNO),
#   elaborado por Roberto: CNO -> CIUO (tabla V3) -> SOC/O*NET, media simple. Valor =
#   proporción aproximada del tiempo de trabajo con tareas en las que se observa
#   uso de IA (índice de Massenkoff et al. 2026 / Anthropic Economic Index).
#   3 CNO quedan "Sin dato" (5891, 5894, 7704): no entran en ningún tramo, pero
#   sí en el total nacional (igual que las no clasificadas de la OIT).
#   Al ser numérico se parte en 4 tramos con cortes fijos (elegidos por Roberto):
#   >= 0,30 | 0,20-0,30 | 0,10-0,20 | < 0,10.  Se desplegan los 2 tramos superiores.
# ---------------------------------------------------------------------------

ANTHROPIC_TIER_META = [
    ("muyalta", "Exposición muy alta (índice ≥ 0,30)", "#FF5A45", "exposición muy alta (Anthropic)"),
    ("alta", "Exposición alta (índice 0,20–0,30)", "#FF8A4C", "exposición alta (Anthropic)"),
    ("media", "Exposición media (índice 0,10–0,20)", "#F2B84C", "exposición media (Anthropic)"),
    ("baja", "Exposición baja (índice < 0,10)", "#33C2A6", "exposición baja (Anthropic)"),
]
# (clave, umbral inferior inclusivo), de mayor a menor
ANTHROPIC_CORTES = [("muyalta", 0.30), ("alta", 0.20), ("media", 0.10), ("baja", 0.0)]

ANTHROPIC_CNO_VALOR = {
    "0011": 0.0321,
    "0012": 0.023,
    "0020": 0.0168,
    "1111": 0.0167,
    "1112": 0.0438,
    "1113": 0,
    "1120": 0.0972,
    "1211": 0.1978,
    "1212": 0.1285,
    "1219": 0.1349,
    "1221": 0.229,
    "1222": 0.2859,
    "1223": 0.1125,
    "1311": 0.0025,
    "1312": 0,
    "1313": 0.0577,
    "1314": 0.0496,
    "1315": 0.094,
    "1316": 0.1035,
    "1321": 0.2975,
    "1322": 0.0371,
    "1323": 0.2772,
    "1324": 0.0673,
    "1325": 0.1808,
    "1326": 0.1158,
    "1327": 0.1293,
    "1329": 0.1226,
    "1411": 0.0623,
    "1419": 0.0623,
    "1421": 0.0188,
    "1422": 0.0188,
    "1429": 0.0188,
    "1431": 0.0653,
    "1432": 0.0653,
    "1501": 0.0609,
    "1509": 0.0672,
    "2111": 0.0416,
    "2112": 0.044,
    "2121": 0.0343,
    "2122": 0.0343,
    "2123": 0.0412,
    "2130": 0.0309,
    "2140": 0.0727,
    "2151": 0.0039,
    "2152": 0.0179,
    "2153": 0.1288,
    "2154": 0.0069,
    "2155": 0.0053,
    "2156": 0.0326,
    "2157": 0.0326,
    "2158": 0.0309,
    "2159": 0.0464,
    "2210": 0.1337,
    "2220": 0.0884,
    "2230": 0.1666,
    "2240": 0.1381,
    "2251": 0.0242,
    "2252": 0.0242,
    "2311": 0.1282,
    "2312": 0.1282,
    "2321": 0.1022,
    "2322": 0.1305,
    "2323": 0.1753,
    "2324": 0.1364,
    "2325": 0.2603,
    "2326": 0.1384,
    "2329": 0.1384,
    "2411": 0.1758,
    "2412": 0.0379,
    "2413": 0.1155,
    "2414": 0.0754,
    "2415": 0.3056,
    "2416": 0.3056,
    "2421": 0.1195,
    "2422": 0.0117,
    "2423": 0.0117,
    "2424": 0.0117,
    "2425": 0.0117,
    "2426": 0.0542,
    "2427": 0.0117,
    "2431": 0.0777,
    "2432": 0.0606,
    "2433": 0.0283,
    "2434": 0.0283,
    "2435": 0.0995,
    "2436": 0.0446,
    "2437": 0.0679,
    "2439": 0.0573,
    "2441": 0.0595,
    "2442": 0.0986,
    "2443": 0.1031,
    "2451": 0.0266,
    "2452": 0.0311,
    "2453": 0.0273,
    "2454": 0.0877,
    "2461": 0.0777,
    "2462": 0.0606,
    "2463": 0.0283,
    "2464": 0.0283,
    "2465": 0.0995,
    "2466": 0.0446,
    "2469": 0.0573,
    "2471": 0.0595,
    "2472": 0.0986,
    "2473": 0.1031,
    "2481": 0.027,
    "2482": 0.0196,
    "2483": 0.0877,
    "2484": 0.247,
    "2511": 0.1799,
    "2512": 0.1799,
    "2513": 0.1882,
    "2591": 0.1628,
    "2592": 0.1628,
    "2599": 0.1628,
    "2611": 0.1479,
    "2612": 0.1987,
    "2613": 0.1995,
    "2621": 0.3204,
    "2622": 0.1426,
    "2623": 0.1426,
    "2624": 0.1674,
    "2625": 0.3285,
    "2630": 0.1426,
    "2640": 0.1706,
    "2651": 0.2586,
    "2652": 0.2752,
    "2653": 0.3224,
    "2711": 0.3182,
    "2712": 0.336,
    "2713": 0.4535,
    "2719": 0.2897,
    "2721": 0.3772,
    "2722": 0.3135,
    "2723": 0.2833,
    "2729": 0.2691,
    "2810": 0.1728,
    "2821": 0.1673,
    "2822": 0.1971,
    "2823": 0.0603,
    "2824": 0.0328,
    "2825": 0.0328,
    "2830": 0.1323,
    "2911": 0.1314,
    "2912": 0.2018,
    "2921": 0.2843,
    "2922": 0.2621,
    "2923": 0.1944,
    "2931": 0.1142,
    "2932": 0.0123,
    "2933": 0.0401,
    "2934": 0.1523,
    "2935": 0.1011,
    "2936": 0.0634,
    "2937": 0.0637,
    "2939": 0.0637,
    "3110": 0.0402,
    "3121": 0.0798,
    "3122": 0.0397,
    "3123": 0.0143,
    "3124": 0.049,
    "3125": 0.049,
    "3126": 0.03,
    "3127": 0.0756,
    "3128": 0.0774,
    "3129": 0.0492,
    "3131": 0.0135,
    "3132": 0.004,
    "3133": 0.0448,
    "3134": 0.0141,
    "3135": 0,
    "3139": 0.022,
    "3141": 0.1054,
    "3142": 0.0446,
    "3143": 0.002,
    "3151": 0.0799,
    "3152": 0,
    "3153": 0.0108,
    "3154": 0.0119,
    "3155": 0,
    "3160": 0.0519,
    "3201": 0.0059,
    "3202": 0.0104,
    "3203": 0.0306,
    "3204": 0.0306,
    "3205": 0.0306,
    "3206": 0.0306,
    "3207": 0.0306,
    "3209": 0.0306,
    "3311": 0.0386,
    "3312": 0.0386,
    "3313": 0.061,
    "3314": 0.061,
    "3315": 0.0021,
    "3316": 0.0021,
    "3317": 0.0021,
    "3321": 0.0044,
    "3322": 0.2824,
    "3323": 0.0298,
    "3324": 0.0122,
    "3325": 0.0173,
    "3326": 0.0622,
    "3327": 0.0221,
    "3329": 0.0298,
    "3331": 0.0136,
    "3339": 0,
    "3401": 0.2865,
    "3402": 0.1698,
    "3403": 0.1679,
    "3404": 0.2627,
    "3405": 0.0695,
    "3510": 0.4493,
    "3521": 0.1327,
    "3522": 0.1075,
    "3523": 0.3291,
    "3531": 0.1027,
    "3532": 0.0551,
    "3533": 0.0216,
    "3534": 0.1553,
    "3535": 0.1603,
    "3539": 0.1603,
    "3611": 0.2261,
    "3612": 0.1841,
    "3613": 0.2129,
    "3614": 0.3327,
    "3621": 0.0957,
    "3622": 0.0707,
    "3623": 0.1646,
    "3629": 0.0203,
    "3631": 0.0558,
    "3632": 0.0558,
    "3711": 0.1269,
    "3712": 0.1269,
    "3713": 0.0373,
    "3714": 0.0373,
    "3715": 0.0373,
    "3716": 0.0882,
    "3721": 0,
    "3722": 0.0227,
    "3723": 0.0178,
    "3724": 0.0178,
    "3731": 0.0686,
    "3732": 0.0204,
    "3733": 0.1516,
    "3734": 0.0019,
    "3739": 0.0474,
    "3811": 0.312,
    "3812": 0.3115,
    "3813": 0.2446,
    "3814": 0.3466,
    "3820": 0.3115,
    "3831": 0.113,
    "3832": 0.113,
    "3833": 0.1037,
    "4111": 0.185,
    "4112": 0.1197,
    "4113": 0.1454,
    "4121": 0.0191,
    "4122": 0.0934,
    "4123": 0.0722,
    "4210": 0.2229,
    "4221": 0.1284,
    "4222": 0.2911,
    "4223": 0.2142,
    "4301": 0.3329,
    "4309": 0.2333,
    "4411": 0.2636,
    "4412": 0.2723,
    "4421": 0.1404,
    "4422": 0.188,
    "4423": 0.1795,
    "4424": 0.355,
    "4430": 0.3172,
    "4441": 0.0503,
    "4442": 0.0891,
    "4443": 0.0891,
    "4444": 0.0231,
    "4445": 0.1246,
    "4446": 0.0503,
    "4500": 0.1364,
    "5000": 0.0369,
    "5110": 0.0432,
    "5120": 0.0353,
    "5210": 0.1696,
    "5220": 0.0562,
    "5300": 0.0686,
    "5411": 0.0848,
    "5412": 0.0848,
    "5420": 0.2238,
    "5430": 0,
    "5491": 0.1758,
    "5492": 0.0849,
    "5493": 0,
    "5499": 0.1234,
    "5500": 0.0716,
    "5611": 0.0053,
    "5612": 0.0053,
    "5621": 0.0617,
    "5622": 0,
    "5629": 0.0084,
    "5710": 0,
    "5721": 0.048,
    "5722": 0.0086,
    "5811": 0.0076,
    "5812": 0.0515,
    "5821": 0.083,
    "5822": 0.1169,
    "5823": 0.0338,
    "5824": 0.0338,
    "5825": 0.0338,
    "5831": 0.0281,
    "5832": 0.0514,
    "5833": 0.0576,
    "5840": 0.0514,
    "5892": 0.0264,
    "5893": 0.0026,
    "5895": 0.1722,
    "5899": 0.0143,
    "5910": 0.0434,
    "5921": 0.0434,
    "5922": 0.0259,
    "5923": 0.0259,
    "5931": 0,
    "5932": 0,
    "5941": 0.0494,
    "5942": 0.0494,
    "5991": 0.0444,
    "5992": 0.0168,
    "5993": 0.0168,
    "5999": 0.0168,
    "6110": 0.0085,
    "6120": 0.0143,
    "6201": 0,
    "6202": 0,
    "6203": 0,
    "6204": 0,
    "6205": 0,
    "6209": 0,
    "6300": 0,
    "6410": 0.0032,
    "6421": 0.0014,
    "6422": 0,
    "6423": 0,
    "6430": 0,
    "7111": 0.0026,
    "7112": 0.0026,
    "7121": 0,
    "7122": 0,
    "7131": 0.0016,
    "7132": 0.0016,
    "7191": 0.0035,
    "7192": 0.0035,
    "7193": 0.0035,
    "7199": 0.0035,
    "7211": 0,
    "7212": 0,
    "7221": 0.0148,
    "7222": 0.0148,
    "7223": 0.0148,
    "7231": 0,
    "7232": 0,
    "7240": 0,
    "7250": 0.0191,
    "7291": 0.0052,
    "7292": 0.011,
    "7293": 0,
    "7294": 0.0148,
    "7295": 0.0092,
    "7311": 0,
    "7312": 0.0057,
    "7313": 0.0147,
    "7314": 0.0082,
    "7315": 0.0044,
    "7321": 0,
    "7322": 0.0142,
    "7323": 0.0083,
    "7324": 0,
    "7401": 0.006,
    "7402": 0.0215,
    "7403": 0.0026,
    "7404": 0.0026,
    "7405": 0,
    "7510": 0.0098,
    "7521": 0.004,
    "7522": 0.0074,
    "7531": 0.0384,
    "7532": 0.0384,
    "7533": 0.0525,
    "7611": 0.0062,
    "7612": 0.0178,
    "7613": 0,
    "7614": 0,
    "7615": 0,
    "7616": 0.0586,
    "7617": 0.012,
    "7618": 0.0215,
    "7619": 0.0082,
    "7621": 0.2096,
    "7622": 0.139,
    "7623": 0.0312,
    "7701": 0,
    "7702": 0,
    "7703": 0,
    "7705": 0,
    "7706": 0,
    "7707": 0,
    "7708": 0,
    "7709": 0.051,
    "7811": 0,
    "7812": 0,
    "7820": 0.006,
    "7831": 0.0092,
    "7832": 0.0144,
    "7833": 0.0144,
    "7834": 0.0176,
    "7835": 0.0099,
    "7836": 0.0324,
    "7837": 0.0104,
    "7891": 0,
    "7892": 0,
    "7893": 0.0262,
    "7894": 0.0309,
    "7899": 0,
    "8111": 0,
    "8112": 0.0047,
    "8113": 0.0081,
    "8114": 0.0027,
    "8121": 0,
    "8122": 0.0053,
    "8131": 0.0103,
    "8132": 0.0103,
    "8133": 0.1059,
    "8141": 0.0064,
    "8142": 0,
    "8143": 0.0064,
    "8144": 0.0036,
    "8145": 0,
    "8151": 0,
    "8152": 0.0157,
    "8153": 0.0071,
    "8154": 0.0211,
    "8155": 0.0106,
    "8156": 0.0119,
    "8159": 0.008,
    "8160": 0.0045,
    "8170": 0.0109,
    "8191": 0,
    "8192": 0,
    "8193": 0.0066,
    "8199": 0.0053,
    "8201": 0.0116,
    "8202": 0.0065,
    "8209": 0.003,
    "8311": 0,
    "8312": 0.068,
    "8321": 0,
    "8322": 0,
    "8331": 0.002,
    "8332": 0,
    "8333": 0,
    "8340": 0,
    "8411": 0.0053,
    "8412": 0.0053,
    "8420": 0,
    "8431": 0.0019,
    "8432": 0.0019,
    "8440": 0.0173,
    "9100": 0,
    "9210": 0.0048,
    "9221": 0,
    "9222": 0,
    "9223": 0,
    "9229": 0,
    "9310": 0,
    "9320": 0.0019,
    "9410": 0.0848,
    "9420": 0.1758,
    "9431": 0.0375,
    "9432": 0.0375,
    "9433": 0.0375,
    "9434": 0.0481,
    "9441": 0,
    "9442": 0,
    "9443": 0,
    "9490": 0.009,
    "9511": 0.0101,
    "9512": 0.011,
    "9520": 0,
    "9530": 0.0144,
    "9541": 0,
    "9542": 0,
    "9543": 0.0025,
    "9601": 0.0023,
    "9602": 0.007,
    "9603": 0,
    "9700": 0.0039,
    "9811": 0.0159,
    "9812": 0.0069,
    "9820": 0,
}


def build_cno_to_anthropic_nivel(cno_valor=None):
    """CNO (4 díg.) -> clave de tramo Anthropic, según ANTHROPIC_CORTES."""
    cno_valor = ANTHROPIC_CNO_VALOR if cno_valor is None else cno_valor
    out = {}
    for cno, v in cno_valor.items():
        for key, lo in ANTHROPIC_CORTES:
            if v >= lo:
                out[cno] = key
                break
    return out


def build_isco3_to_cnos(cno_to_isco4):
    """Agrupa CNO_TO_ISCO4 al revés: ISCO a 3 dígitos -> lista de CNO cuyo
    ISCO de 4 dígitos empieza por esos 3 (ej. '251' -> todas las CNO con
    ISCO4 '251X')."""
    out = {}
    for cno, isco4 in cno_to_isco4.items():
        out.setdefault(isco4[:3], []).append(cno)
    return out


CNO3_NAME = {  # nombres oficiales CNO-11 a 3 dígitos (fichero CNO2011 aportado por Roberto, oct-2026)
    '001': 'Oficiales y suboficiales de las fuerzas armadas',
    '002': 'Tropa y marinería de las fuerzas armadas',
    '111': 'Miembros del poder ejecutivo y de los cuerpos legislativos; directivos de la Administración Pública y organizaciones de interés social',
    '112': 'Directores generales y presidentes ejecutivos',
    '121': 'Directores de departamentos administrativos',
    '122': 'Directores comerciales, de publicidad, relaciones públicas y de investigación y desarrollo',
    '131': 'Directores de producción de explotaciones agropecuarias, forestales y pesqueras, y de industrias manufactureras, de minería, construcción y distribución',
    '132': 'Directores de servicios de tecnologías de la información y las comunicaciones (TIC) y de empresas de servicios profesionales',
    '141': 'Directores y gerentes de empresas de alojamiento',
    '142': 'Directores y gerentes de empresas de restauración',
    '143': 'Directores y gerentes de empresas de comercio al por mayor y al por menor',
    '150': 'Directores y gerentes de otras empresas de servicios no clasificados bajo otros epígrafes',
    '211': 'Médicos',
    '212': 'Profesionales de enfermería y partería',
    '213': 'Veterinarios',
    '214': 'Farmacéuticos',
    '215': 'Otros profesionales de la salud',
    '221': 'Profesores de universidades y otra enseñanza superior (excepto formación profesional)',
    '222': 'Profesores de formación profesional (materias específicas)',
    '223': 'Profesores de enseñanza secundaria (excepto materias específicas de formación profesional)',
    '224': 'Profesores de enseñanza primaria',
    '225': 'Maestros y educadores de enseñanza infantil',
    '231': 'Profesores y técnicos de educación especial',
    '232': 'Otros profesores y profesionales de la enseñanza',
    '241': 'Físicos, químicos, matemáticos y afines',
    '242': 'Profesionales en ciencias naturales',
    '243': 'Ingenieros (excepto ingenieros agrónomos, de montes, eléctricos, electrónicos y TIC)',
    '244': 'Ingenieros eléctricos, electrónicos y de telecomunicaciones',
    '245': 'Arquitectos, urbanistas e ingenieros geógrafos',
    '246': 'Ingenieros técnicos (excepto agrícolas, forestales, eléctricos, electrónicos y TIC)',
    '247': 'Ingenieros técnicos en electricidad, electrónica y telecomunicaciones',
    '248': 'Arquitectos técnicos, topógrafos y diseñadores',
    '251': 'Jueces, magistrados, abogados y fiscales',
    '259': 'Otros profesionales del derecho',
    '261': 'Especialistas en finanzas',
    '262': 'Especialistas en organización y administración',
    '263': 'Técnicos de empresas y actividades turísticas',
    '264': 'Profesionales de ventas técnicas y médicas (excepto las TIC)',
    '265': 'Otros profesionales de las ventas, la comercialización, la publicidad y las relaciones públicas',
    '271': 'Analistas y diseñadores de software y multimedia',
    '272': 'Especialistas en bases de datos y en redes informáticas',
    '281': 'Economistas',
    '282': 'Sociólogos, historiadores, psicólogos y otros profesionales en ciencias sociales',
    '283': 'Sacerdotes de las distintas religiones',
    '291': 'Archivistas, bibliotecarios, conservadores y afines',
    '292': 'Escritores, periodistas y lingüistas',
    '293': 'Artistas creativos e interpretativos',
    '311': 'Delineantes y dibujantes técnicos',
    '312': 'Técnicos de las ciencias físicas, químicas, medioambientales y de las ingenierías',
    '313': 'Técnicos en control de procesos',
    '314': 'Técnicos de las ciencias naturales y profesionales auxiliares afines',
    '315': 'Profesionales en navegación marítima y aeronáutica',
    '316': 'Técnicos de control de calidad de las ciencias físicas, químicas y de las ingenierías',
    '320': 'Supervisores en ingeniería de minas, de industrias manufactureras y de la construcción',
    '331': 'Técnicos sanitarios de laboratorio, pruebas diagnósticas y prótesis',
    '332': 'Otros técnicos sanitarios',
    '333': 'Profesionales de las terapias alternativas',
    '340': 'Profesionales de apoyo en finanzas y matemáticas',
    '351': 'Agentes y representantes comerciales',
    '352': 'Otros agentes comerciales',
    '353': 'Agentes inmobiliarios y otros agentes',
    '361': 'Asistentes administrativos y especializados',
    '362': 'Agentes de aduanas, tributos y afines que trabajan en tareas propias de la Administración Pública',
    '363': 'Técnicos de las fuerzas y cuerpos de seguridad',
    '371': 'Profesionales de apoyo de servicios jurídicos y sociales',
    '372': 'Deportistas, entrenadores, instructores de actividades deportivas; monitores de actividades recreativas',
    '373': 'Técnicos y profesionales de apoyo de actividades culturales, artísticas y culinarias',
    '381': 'Técnicos en operaciones de tecnologías de la información y asistencia al usuario',
    '382': 'Programadores informáticos',
    '383': 'Técnicos en grabación audiovisual, radiodifusión y telecomunicaciones',
    '411': 'Empleados contables y financieros',
    '412': 'Empleados de registro de materiales, de servicios de apoyo a la producción y al transporte',
    '421': 'Empleados de bibliotecas y archivos',
    '422': 'Empleados de servicios de correos, codificadores, correctores y servicios de personal',
    '430': 'Otros empleados administrativos sin tareas de atención al público',
    '441': 'Empleados de información y recepcionistas (excepto de hoteles)',
    '442': 'Empleados de agencias de viajes, recepcionistas de hoteles y telefonistas',
    '443': 'Agentes de encuestas',
    '444': 'Empleados de ventanilla y afines (excepto taquilleros)',
    '450': 'Empleados administrativos con tareas de atención al público no clasificados bajo otros epígrafes',
    '500': 'Camareros y cocineros propietarios',
    '511': 'Cocineros asalariados',
    '512': 'Camareros asalariados',
    '521': 'Jefes de sección de tiendas y almacenes',
    '522': 'Vendedores en tiendas y almacenes',
    '530': 'Comerciantes propietarios de tiendas',
    '541': 'Vendedores en quioscos o en mercadillos',
    '542': 'Operadores de telemarketing',
    '543': 'Expendedores de gasolineras',
    '549': 'Otros vendedores',
    '550': 'Cajeros y taquilleros (excepto bancos)',
    '561': 'Auxiliares de enfermería',
    '562': 'Técnicos auxiliares de farmacia y emergencias sanitarias y otros trabajadores de los cuidados a las personas en servicios de salud',
    '571': 'Trabajadores de los cuidados personales a domicilio (excepto cuidadores de niños)',
    '572': 'Cuidadores de niños',
    '581': 'Peluqueros y especialistas en tratamientos de estética, bienestar y afines',
    '582': 'Trabajadores que atienden a viajeros, guías turísticos y afines',
    '583': 'Supervisores de mantenimiento y limpieza de edificios, conserjes y mayordomos domésticos',
    '584': 'Trabajadores propietarios de pequeños alojamientos',
    '589': 'Otros trabajadores de servicios personales',
    '591': 'Guardias civiles',
    '592': 'Policías',
    '593': 'Bomberos',
    '594': 'Personal de seguridad privado',
    '599': 'Otros trabajadores de los servicios de protección y seguridad',
    '611': 'Trabajadores cualificados en actividades agrícolas (excepto en huertas, invernaderos, viveros y jardines)',
    '612': 'Trabajadores cualificados en huertas, invernaderos, viveros y jardines',
    '620': 'Trabajadores cualificados en actividades ganaderas (incluidas avícolas, apícolas y similares)',
    '630': 'Trabajadores cualificados en actividades agropecuarias mixtas',
    '641': 'Trabajadores cualificados en actividades forestales y del medio natural',
    '642': 'Trabajadores cualificados en actividades pesqueras y acuicultura',
    '643': 'Trabajadores cualificados en actividades cinegéticas',
    '711': 'Trabajadores en hormigón, encofradores, ferrallistas y afines',
    '712': 'Albañiles, canteros, tronzadores, labrantes y grabadores de piedras',
    '713': 'Carpinteros (excepto ebanistas y montadores de estructuras metálicas)',
    '719': 'Otros trabajadores de las obras estructurales de construcción',
    '721': 'Escayolistas y aplicadores de revestimientos de pasta y mortero',
    '722': 'Fontaneros e instaladores de tuberías',
    '723': 'Pintores, empapeladores y afines',
    '724': 'Soladores, colocadores de parquet y afines',
    '725': 'Mecánicos-instaladores de refrigeración y climatización',
    '729': 'Otros trabajadores de acabado en la construcción, instalaciones (excepto electricistas) y afines',
    '731': 'Moldeadores, soldadores, chapistas, montadores de estructuras metálicas y trabajadores afines',
    '732': 'Herreros y trabajadores de la fabricación de herramientas y afines',
    '740': 'Mecánicos y ajustadores de maquinaria',
    '751': 'Electricistas de la construcción y afines',
    '752': 'Otros instaladores y reparadores de equipos eléctricos',
    '753': 'Instaladores y reparadores de equipos electrónicos y de telecomunicaciones',
    '761': 'Mecánicos de precisión en metales, ceramistas, vidrieros y artesanos',
    '762': 'Oficiales y operarios de las artes gráficas',
    '770': 'Trabajadores de la industria de la alimentación, bebidas y tabaco',
    '781': 'Trabajadores que tratan la madera y afines',
    '782': 'Ebanistas y trabajadores afines',
    '783': 'Trabajadores del textil, confección, piel, cuero y calzado',
    '789': 'Pegadores, buceadores, probadores de productos y otros operarios y artesanos diversos',
    '811': 'Operadores en instalaciones de la extracción y explotación de minerales',
    '812': 'Operadores en instalaciones para el tratamiento de metales',
    '813': 'Operadores de instalaciones y máquinas de productos químicos, farmacéuticos y materiales fotosensibles',
    '814': 'Operadores en instalaciones para el tratamiento y transformación de la madera, la fabricación de papel, productos de papel y caucho o materias plásticas',
    '815': 'Operadores de máquinas para fabricar productos textiles y artículos de piel y de cuero',
    '816': 'Operadores de máquinas para elaborar productos alimenticios, bebidas y tabaco',
    '817': 'Operadores de máquinas de lavandería y tintorería',
    '819': 'Otros operadores de instalaciones y maquinaria fijas',
    '820': 'Montadores y ensambladores en fábricas',
    '831': 'Maquinistas de locomotoras y afines',
    '832': 'Operadores de maquinaria agrícola y forestal móvil',
    '833': 'Operadores de otras máquinas móviles',
    '834': 'Marineros de puente, marineros de máquinas y afines',
    '841': 'Conductores de automóviles, taxis y furgonetas',
    '842': 'Conductores de autobuses y tranvías',
    '843': 'Conductores de camiones',
    '844': 'Conductores de motocicletas y ciclomotores',
    '910': 'Empleados domésticos',
    '921': 'Personal de limpieza de oficinas, hoteles y otros establecimientos similares',
    '922': 'Limpiadores de vehículos, ventanas y personal de limpieza a mano',
    '931': 'Ayudantes de cocina',
    '932': 'Preparadores de comidas rápidas',
    '941': 'Vendedores callejeros',
    '942': 'Repartidores de publicidad, limpiabotas y otros trabajadores de oficios callejeros',
    '943': 'Ordenanzas, mozos de equipaje, repartidores a pie y afines',
    '944': 'Recogedores de residuos, clasificadores de desechos, barrenderos y afines',
    '949': 'Otras ocupaciones elementales',
    '951': 'Peones agrícolas',
    '952': 'Peones ganaderos',
    '953': 'Peones agropecuarios',
    '954': 'Peones de la pesca, la acuicultura, forestales y de la caza',
    '960': 'Peones de la construcción y de la minería',
    '970': 'Peones de las industrias manufactureras',
    '981': 'Peones del transporte, descargadores y afines',
    '982': 'Reponedores',
}

# Aclaración (entre paréntesis) de cada parte de un grupo CNO-3 "mixto": grupos
# cuyas CNO caen en niveles JRC distintos porque CNO-11 e ISCO-08 no coinciden a
# 3 dígitos. Decisión de Roberto (oct-2026): una fila por (grupo, nivel).
JRC_PARCIAL_ACLARACION = {
    ("241", "alta"): "matemáticos y estadísticos",
    ("241", "media"): "físicos, meteorólogos, químicos y geólogos",
    ("333", "alta"): "acupuntura, naturopatía, homeopatía, medicina tradicional china y ayurveda",
    ("333", "media"): "otras terapias alternativas",
    ("430", "media"): "grabadores de datos",
    ("430", "baja"): "resto de empleados administrativos sin atención al público",
    ("562", "media"): "técnicos auxiliares de farmacia y de emergencias sanitarias",
    ("562", "baja"): "otros trabajadores de los cuidados en servicios de salud",
}


def build_jrc_cno3_groups(cno_to_jrc_nivel):
    """Agrupa las CNO con nivel JRC por grupo CNO de 3 dígitos.
    Devuelve {clave: {"code", "name", "nivel", "cnos", "parcial"}}.
    Si todas las CNO del grupo tienen el mismo nivel -> clave = '241'.
    Si el grupo es mixto -> una fila por nivel, clave = '241-alta', con el
    código '241' y la aclaración entre paréntesis en el nombre."""
    por_grupo = {}
    for cno, nivel in cno_to_jrc_nivel.items():
        por_grupo.setdefault(cno[:3], {}).setdefault(nivel, []).append(cno)
    out = {}
    for g, niveles in por_grupo.items():
        base = CNO3_NAME.get(g, g)
        if len(niveles) == 1:
            (nivel, cnos), = niveles.items()
            out[g] = {"code": g, "name": base, "nivel": nivel, "cnos": sorted(cnos), "parcial": False}
        else:
            for nivel, cnos in niveles.items():
                aclar = JRC_PARCIAL_ACLARACION[(g, nivel)]  # KeyError = grupo mixto nuevo sin aclarar
                out[f"{g}-{nivel}"] = {"code": g, "name": f"{base} ({nivel.capitalize()}: {aclar})",
                                       "nivel": nivel, "cnos": sorted(cnos), "parcial": True}
    return out


def aggregate_occ_block(rows):
    """Como occupation_block, pero para un GRUPO de filas en vez de una sola
    (se usa para agregar varias ocupaciones CNO bajo un mismo ISCO a 3
    dígitos): suma primero total/sexo/edad de TODO el grupo y solo entonces
    calcula el % de cambio."""
    t1, t2 = sum_group(rows, "total1", "total2")
    h1, h2 = sum_group(rows, "hombre1", "hombre2")
    m1, m2 = sum_group(rows, "mujer1", "mujer2")
    e30_1, e30_2 = sum_group(rows, "e1_30", "e2_30")
    e3045_1, e3045_2 = sum_group(rows, "e1_3045", "e2_3045")
    e45_1, e45_2 = sum_group(rows, "e1_45", "e2_45")
    return dict(
        total=pct(t1, t2), hombres=pct(h1, h2), mujeres=pct(m1, m2),
        menos30=pct(e30_1, e30_2), entre3045=pct(e3045_1, e3045_2), mas45=pct(e45_1, e45_2),
    )


CNAE_SECTIONS = [
    ["A", "Agricultura, ganadería y pesca"], ["B", "Industrias extractivas"],
    ["C", "Industria manufacturera"], ["D", "Energía eléctrica y gas"],
    ["E", "Agua y residuos"], ["F", "Construcción"],
    ["G", "Comercio y reparación de vehículos"], ["H", "Transporte y almacenamiento"],
    ["I", "Hostelería"], ["J", "Información y comunicaciones"],
    ["K", "Actividades financieras y de seguros"], ["L", "Actividades inmobiliarias"],
    ["M", "Actividades profesionales y técnicas"], ["N", "Actividades administrativas y auxiliares"],
    ["O", "Administración pública y defensa"], ["P", "Educación"],
    ["Q", "Actividades sanitarias y sociales"], ["R", "Actividades artísticas y recreativas"],
    ["S", "Otros servicios"], ["T", "Actividades de los hogares"],
    ["U", "Organismos extraterritoriales"],
]

MAXIMA_CODES = {c for c, _, t in OCUPACIONES_ORDEN if t == "maxima"}
SIGNIFICATIVA_CODES = {c for c, _, t in OCUPACIONES_ORDEN if t == "significativa"}

# ---------------------------------------------------------------------------
# 2. Posiciones de columna EN BRUTO (tal cual el original, sin fórmulas).
#    Son fijas mientras Roberto no cambie el export de su base de datos.
# ---------------------------------------------------------------------------

RAW_COLS_PARADOS = dict(
    cno=1, desc=2, exposure=4, total1=5, total2=6,
    hombre1=7, mujer1=8, hombre2=9, mujer2=10,
    e1_30=11, e1_3045=12, e1_45=13, e2_30=14, e2_3045=15, e2_45=16,
)

RAW_COLS_CONTRATADAS = dict(
    cno=1, desc=2, exposure=4, total1=5, total2=6,
    hombre1=8, mujer1=9, hombre2=10, mujer2=11,
    e1_30=14, e1_3045=15, e1_45=16, e2_30=17, e2_3045=18, e2_45=19,
)

SHEETS = {
    ("desempleo", "mes"): ("Parados por CNO mes", RAW_COLS_PARADOS),
    ("desempleo", "anio"): ("Parados por CNO año", RAW_COLS_PARADOS),
    ("contratos", "mes"): ("P Contratadas por CNO mes", RAW_COLS_CONTRATADAS),
    ("contratos", "anio"): ("P Contratadas por CNO año", RAW_COLS_CONTRATADAS),
}

METRIC_TITULOS = {
    "contratos": "Personas contratadas",
    "desempleo": "Parados",
}


# ---------------------------------------------------------------------------
# Utilidades
# ---------------------------------------------------------------------------

def num(v):
    """Convierte a float cualquier celda; None/texto/vacío -> 0.0 (para sumar)."""
    if isinstance(v, (int, float)):
        return float(v)
    return 0.0


def pct(v1, v2, decimals=1):
    """% de cambio entre dos totales ABSOLUTOS (nunca media de porcentajes)."""
    if v1 is None or v2 is None or v1 == 0:
        return None
    return round((v2 - v1) / v1 * 100, decimals)


def parse_period_label(texto):
    """'PC TOTAL AGOSTO DE 2025' -> 'Agosto 2025'; 'PC TOTAL 2024' -> '2024'."""
    if texto is None:
        return ""
    texto = str(texto).strip()
    m = re.search(r"([A-ZÁÉÍÓÚÑ]+)\s+DE\s+(\d{4})", texto, re.IGNORECASE)
    if m:
        return f"{m.group(1).capitalize()} {m.group(2)}"
    m2 = re.search(r"(\d{4})", texto)
    if m2:
        return m2.group(1)
    return texto


# ---------------------------------------------------------------------------
# Lectura del original
# ---------------------------------------------------------------------------

def detectar_cols(ws, cols_por_defecto):
    """Localiza las columnas de la hoja POR SU CABECERA (fila 1), no por su
    posición, así que da igual que Roberto quite o añada columnas (Exposure,
    columnas de % como "% TOTAL", etc.). Reglas, sobre el texto de la cabecera:
      - cno / desc: cabeceras "CNO ID" y "CNO DESC" (si no, columnas 1 y 2).
      - exposure: cabecera con "Exposure" (opcional; la clasificación OIT va en el script).
      - total1 / total2: las dos primeras columnas tras las de identificación
        (cabeceras que empiezan por "CNO" o "Exposure") — los totales de los dos periodos.
      - hombre1 / hombre2, mujer1 / mujer2: las dos primeras cabeceras con
        "Hombre" / "Mujer" (1.ª = periodo anterior, 2.ª = periodo actual).
      - edad: cabeceras con "<30", "30 a 45" / "30-45" y ">45" / "Mayores de 45",
        igual: 1.ª aparición = periodo anterior, 2.ª = actual.
      Se ignoran las columnas de % o de variación ("%...", "...25vs24").
    Si falta algo, se para con un mensaje que lista las cabeceras leídas."""
    def norm(v):
        return "" if v is None else " ".join(str(v).lower().split())

    heads = {c: norm(ws.cell(row=1, column=c).value) for c in range(1, ws.max_column + 1)}
    es_base = lambda h: h.startswith("cno") or h.startswith("exposure")  # id/descripciones/exposure

    out = {"cno": None, "desc": None, "exposure": None}
    for c, h in heads.items():
        if h == "cno id" and out["cno"] is None:
            out["cno"] = c
        elif h == "cno desc" and out["desc"] is None:
            out["desc"] = c
        elif "exposure" in h and out["exposure"] is None:
            out["exposure"] = c
    if out["cno"] is None:
        out["cno"] = cols_por_defecto["cno"]
    if out["desc"] is None:
        out["desc"] = cols_por_defecto["desc"]

    # totales: las dos primeras columnas que no son identificadores
    tot = [c for c, h in heads.items() if h and not es_base(h) and not h.startswith("%")][:2]
    if len(tot) == 2:
        out["total1"], out["total2"] = tot

    grupos = {"hombre": [], "mujer": [], "e30": [], "e3045": [], "e45": []}
    for c, h in heads.items():
        if not h or h.startswith("%") or "vs" in h or "sin empleo" in h:
            continue
        hs = h.replace(" ", "")
        if "hombre" in hs:
            grupos["hombre"].append(c)
        elif "mujer" in hs:
            grupos["mujer"].append(c)
        elif "<30" in hs:
            grupos["e30"].append(c)
        elif ">45" in hs or "mayoresde45" in hs:
            grupos["e45"].append(c)
        elif re.search(r"30\D{0,3}45", hs):  # "30-45", "30 a 45", "30.-45"...
            grupos["e3045"].append(c)
    for nombre, (k1, k2) in {"hombre": ("hombre1", "hombre2"), "mujer": ("mujer1", "mujer2"),
                             "e30": ("e1_30", "e2_30"), "e3045": ("e1_3045", "e2_3045"),
                             "e45": ("e1_45", "e2_45")}.items():
        if len(grupos[nombre]) >= 2:
            out[k1], out[k2] = grupos[nombre][0], grupos[nombre][1]

    obligatorias = ["total1", "total2", "hombre1", "hombre2", "mujer1", "mujer2",
                    "e1_30", "e1_3045", "e1_45", "e2_30", "e2_3045", "e2_45"]
    faltan = [k for k in obligatorias if out.get(k) is None]
    if faltan:
        cabeceras = [(c, ws.cell(row=1, column=c).value) for c in range(1, min(ws.max_column, 26) + 1)]
        raise SystemExit(
            f"En la hoja '{ws.title}' no encuentro por su cabecera estas columnas: {faltan}. "
            f"Cabeceras leídas (columna, texto): {cabeceras}"
        )
    return out


# Nombres antiguos de pestaña que también se aceptan (por si algún Excel viejo los trae)
ALIAS_HOJAS = {
    "P Contratadas por CNO mes": ["P Contratadas mes"],
    "P Contratadas por CNO año": ["P Contratadas año"],
}


def buscar_hoja(wb, nombre):
    """Devuelve la hoja `nombre` tolerando pequeños cambios (mayúsculas, espacios
    de más, guiones bajos, acentos o la ñ de "año"). Si no la encuentra, para con
    un mensaje que lista las hojas que SÍ hay en el Excel."""
    import unicodedata

    def norm(t):
        t = unicodedata.normalize("NFKD", str(t)).encode("ascii", "ignore").decode()
        return " ".join(t.lower().replace("_", " ").split())

    if nombre in wb.sheetnames:
        return wb[nombre]
    objetivo = norm(nombre)
    candidatos = {objetivo} | {norm(a) for a in ALIAS_HOJAS.get(nombre, [])}
    for n in wb.sheetnames:
        if norm(n) in candidatos:
            return wb[n]
    raise SystemExit(
        f"No encuentro la hoja '{nombre}' en el Excel. Hojas que hay: {wb.sheetnames}. "
        f"Deben llamarse (salvo mayúsculas/espacios): 'Parados por CNO año', "
        f"'Parados por CNO mes', 'P Contratadas por CNO año' y 'P Contratadas por CNO mes'."
    )


def read_sheet_rows(ws, cols):
    """Devuelve la lista de filas (una por ocupación) en bruto, y el texto
    de cabecera de las columnas total1/total2 (para compare_text)."""
    cols = detectar_cols(ws, cols)
    header1 = ws.cell(row=1, column=cols["total1"]).value
    header2 = ws.cell(row=1, column=cols["total2"]).value

    rows = []
    for r in range(2, ws.max_row + 1):
        cno = ws.cell(row=r, column=cols["cno"]).value
        if cno is None or str(cno).strip() == "":
            continue
        rows.append(dict(
            fila_excel=r,
            cno=str(cno).strip(),
            desc=ws.cell(row=r, column=cols["desc"]).value,
            exposure=(ws.cell(row=r, column=cols["exposure"]).value if cols["exposure"] else None),
            total1=num(ws.cell(row=r, column=cols["total1"]).value),
            total2=num(ws.cell(row=r, column=cols["total2"]).value),
            hombre1=num(ws.cell(row=r, column=cols["hombre1"]).value),
            mujer1=num(ws.cell(row=r, column=cols["mujer1"]).value),
            hombre2=num(ws.cell(row=r, column=cols["hombre2"]).value),
            mujer2=num(ws.cell(row=r, column=cols["mujer2"]).value),
            e1_30=num(ws.cell(row=r, column=cols["e1_30"]).value),
            e1_3045=num(ws.cell(row=r, column=cols["e1_3045"]).value),
            e1_45=num(ws.cell(row=r, column=cols["e1_45"]).value),
            e2_30=num(ws.cell(row=r, column=cols["e2_30"]).value),
            e2_3045=num(ws.cell(row=r, column=cols["e2_3045"]).value),
            e2_45=num(ws.cell(row=r, column=cols["e2_45"]).value),
        ))
    return rows, header1, header2


def sum_group(rows, field1, field2):
    return sum(r[field1] for r in rows), sum(r[field2] for r in rows)


def aggregate_by_map(rows, cno_to_key):
    """Generaliza el cálculo de un nivel/tier: agrupa las filas por
    cno_to_key.get(cno) y para cada grupo suma total1/total2 por separado y
    SOLO ENTONCES calcula el % de cambio (nunca media de porcentajes). Vale
    tanto para los 6 niveles de la OIT como para los 5 niveles ficticios de
    la OCDE: el método de cálculo es el mismo, solo cambia el mapa."""
    sums = {}
    for r in rows:
        key = cno_to_key.get(r["cno"])
        if key is None:
            continue
        s = sums.setdefault(key, [0.0, 0.0])
        s[0] += r["total1"]
        s[1] += r["total2"]
    return {key: pct(v[0], v[1]) for key, v in sums.items()}


def bases_by_map(rows, cno_to_key):
    """Como aggregate_by_map pero devuelve las BASES absolutas [total1, total2]
    de cada grupo (para que los textos de análisis puedan avisar de bases pequeñas)."""
    sums = {}
    for r in rows:
        key = cno_to_key.get(r["cno"])
        if key is None:
            continue
        s = sums.setdefault(key, [0.0, 0.0])
        s[0] += r["total1"]
        s[1] += r["total2"]
    return {key: [int(round(v[0])), int(round(v[1]))] for key, v in sums.items()}


def base_of_rows(rows):
    """[total1, total2] sumados de un conjunto de filas."""
    return [int(round(sum(r["total1"] for r in rows))), int(round(sum(r["total2"] for r in rows)))]


def compute_resumen(rows, codes_set):
    """Resumen sexo/edad (sum-then-ratio) de las filas cuyo CNO está en
    codes_set. Se usa para el nivel superior de CADA índice (10 ocupaciones
    de máxima en la OIT, 15 del nivel 5 ficticio en la OCDE...)."""
    sel = [r for r in rows if r["cno"] in codes_set]
    h1, h2 = sum_group(sel, "hombre1", "hombre2")
    m1, m2 = sum_group(sel, "mujer1", "mujer2")
    e30_1, e30_2 = sum_group(sel, "e1_30", "e2_30")
    e3045_1, e3045_2 = sum_group(sel, "e1_3045", "e2_3045")
    e45_1, e45_2 = sum_group(sel, "e1_45", "e2_45")
    return dict(
        sexo=dict(hombres=pct(h1, h2), mujeres=pct(m1, m2)),
        edad=dict(menos30=pct(e30_1, e30_2), entre3045=pct(e3045_1, e3045_2), mas45=pct(e45_1, e45_2)),
    )


def resumen_out_dict(resumen_mes, resumen_anio):
    """Da forma de salida (lista con label/mes/anio) a un par de resumenes
    (uno de mes, uno de año) con la forma que devuelve compute_resumen."""
    return {
        "edad": [
            {"label": "Menos de 30", "mes": resumen_mes["edad"]["menos30"], "anio": resumen_anio["edad"]["menos30"]},
            {"label": "30–45", "mes": resumen_mes["edad"]["entre3045"], "anio": resumen_anio["edad"]["entre3045"]},
            {"label": "Más de 45", "mes": resumen_mes["edad"]["mas45"], "anio": resumen_anio["edad"]["mas45"]},
        ],
        "sexo": [
            {"label": "Hombres", "mes": resumen_mes["sexo"]["hombres"], "anio": resumen_anio["sexo"]["hombres"]},
            {"label": "Mujeres", "mes": resumen_mes["sexo"]["mujeres"], "anio": resumen_anio["sexo"]["mujeres"]},
        ],
        "actividad_economica": [],  # pendiente: se deja vacío de momento
    }


def build_exposure_map(wb_formulas):
    """Mapa CNO -> texto de Exposure, construido a partir de las celdas que
    son texto literal (no fórmula). Una de las 4 pestañas del original
    ("P Contratadas por CNO mes") trae esa columna como un VLOOKUP roto (da
    #VALUE!, el rango de la fórmula solo tiene 4 columnas pero pide la 5ª)
    así que no nos fiamos columna a columna: se construye UNA tabla única
    y se aplica igual a las 4 pestañas. Se necesita el libro cargado SIN
    data_only (data_only=False) para poder distinguir fórmula de texto."""
    exposure_map = {}
    for sheet_name, cols in [SHEETS[("desempleo", "anio")], SHEETS[("desempleo", "mes")],
                              SHEETS[("contratos", "anio")], SHEETS[("contratos", "mes")]]:
        ws = buscar_hoja(wb_formulas, sheet_name)
        for r in range(2, ws.max_row + 1):
            cno = ws.cell(row=r, column=cols["cno"]).value
            if cno is None or str(cno).strip() == "":
                continue
            cno = str(cno).strip()
            if cno in exposure_map:
                continue
            exp = ws.cell(row=r, column=cols["exposure"]).value
            if isinstance(exp, str) and not exp.startswith("="):
                exposure_map[cno] = exp
    return exposure_map


def build_period(ws, cols, exposure_map):
    """Calcula todo lo necesario para una pestaña (una combinación
    metrica x periodo): nacional, 6 niveles, resumen máxima, y un
    diccionario por código CNO con sus cifras en bruto (para sacar
    luego el % de cada una de las 48 ocupaciones)."""
    rows, header1, header2 = read_sheet_rows(ws, cols)
    for r in rows:
        r["exposure"] = exposure_map.get(r["cno"], r["exposure"])

    t1, t2 = sum_group(rows, "total1", "total2")
    national = pct(t1, t2)

    tiers = {}
    for tier_key in dict.fromkeys(EXPOSURE_TO_TIER.values()):
        trows = [r for r in rows if EXPOSURE_TO_TIER.get(r["exposure"]) == tier_key]
        gt1, gt2 = sum_group(trows, "total1", "total2")
        tiers[tier_key] = pct(gt1, gt2)

    # Resumen sexo/edad agregado de las 10 ocupaciones de máxima exposición
    maxima_rows = [r for r in rows if r["cno"] in MAXIMA_CODES]
    h1, h2 = sum_group(maxima_rows, "hombre1", "hombre2")
    m1, m2 = sum_group(maxima_rows, "mujer1", "mujer2")
    e30_1, e30_2 = sum_group(maxima_rows, "e1_30", "e2_30")
    e3045_1, e3045_2 = sum_group(maxima_rows, "e1_3045", "e2_3045")
    e45_1, e45_2 = sum_group(maxima_rows, "e1_45", "e2_45")

    resumen = dict(
        sexo=dict(hombres=pct(h1, h2), mujeres=pct(m1, m2)),
        edad=dict(menos30=pct(e30_1, e30_2), entre3045=pct(e3045_1, e3045_2), mas45=pct(e45_1, e45_2)),
    )

    by_cno = {r["cno"]: r for r in rows}

    return dict(
        national=national, tiers=tiers, resumen=resumen, by_cno=by_cno,
        header1=header1, header2=header2,
    )


def occupation_block(row):
    """% total/sexo/edad de UNA ocupación a partir de su fila en bruto."""
    if row is None:
        return dict(total=None, hombres=None, mujeres=None, menos30=None, entre3045=None, mas45=None)
    return dict(
        total=pct(row["total1"], row["total2"]),
        hombres=pct(row["hombre1"], row["hombre2"]),
        mujeres=pct(row["mujer1"], row["mujer2"]),
        menos30=pct(row["e1_30"], row["e2_30"]),
        entre3045=pct(row["e1_3045"], row["e2_3045"]),
        mas45=pct(row["e1_45"], row["e2_45"]),
    )


# ---------------------------------------------------------------------------
# Construcción del JSON
# ---------------------------------------------------------------------------

def build_json(wb_values, wb_formulas):
    exposure_map = dict(OIT_EXPOSURE_V3)  # V3 corregida; manda sobre la columna Exposure del Excel
    computed = {}  # (metric, period) -> build_period(...)
    for (metric, period), (sheet_name, cols) in SHEETS.items():
        ws = buscar_hoja(wb_values, sheet_name)
        computed[(metric, period)] = build_period(ws, cols, exposure_map)

    # compare_text (se recalcula solo a partir de las cabeceras del original)
    c_mes = computed[("contratos", "mes")]
    c_anio = computed[("contratos", "anio")]
    compare_text = {
        "mes": f"{parse_period_label(c_mes['header2'])} vs {parse_period_label(c_mes['header1'])}",
        "anio": f"{parse_period_label(c_anio['header2'])} vs {parse_period_label(c_anio['header1'])}",
    }

    metrics = {}
    for metric in ("contratos", "desempleo"):
        cm, ca = computed[(metric, "mes")], computed[(metric, "anio")]
        metrics[metric] = {
            "titulo_linea1": METRIC_TITULOS[metric],
            "mes": {"national": cm["national"]},
            "anio": {"national": ca["national"]},
        }
    # Afiliación / Uso de IA / Productividad: todavía no tenemos fuente de datos
    for key, titulo in (
        ("afiliados", "Afiliados SS"),
        ("uso_ia", "Uso IA"),
        ("productividad", "Productividad BdE"),
        ("desempleo_epa", "Desempleo EPA"),
        ("ocupados_epa", "Ocupados EPA"),
        ("salarios_bde", "Salarios BdE"),
    ):
        metrics[key] = {"titulo_linea1": titulo,
                         "mes": {"national": None},
                         "anio": {"national": None}}

    # Universo completo de CNO (para el reparto ficticio de la OCDE): unión de
    # las 4 hojas, excluyendo "ZZZZ" (fila corrupta "Indeterminado").
    all_codes = set()
    for c in computed.values():
        all_codes.update(c["by_cno"].keys())
    all_codes.discard("ZZZZ")
    ocde_nivel = build_ocde_nivel_ficticio(all_codes)
    ocde_codes_por_nivel = {key: sorted(c for c, n in ocde_nivel.items() if n == key)
                             for key, _ in OCDE_NIVEL_COUNTS}

    jrc_nivel = build_cno_to_jrc_nivel(CNO_TO_ISCO4)
    jrc_alta_codes = sorted(c for c, n in jrc_nivel.items() if n == "alta")  # CNO, para el resumen
    isco3_to_cnos = build_isco3_to_cnos(CNO_TO_ISCO4)
    # ISCO a 3 dígitos que se listan en el desplegable "Alta" (solo los que
    # de verdad tienen alguna ocupación CNO presente en el Excel de Roberto)
    jrc_alta_isco3 = sorted(isco3 for isco3, nivel in JRC_ISCO3_TO_NIVEL.items()
                             if nivel == "alta" and isco3 in isco3_to_cnos)
    # Versión española: grupos CNO de 3 dígitos (con filas parciales si el grupo es mixto)
    jrc_cno3 = build_jrc_cno3_groups(jrc_nivel)
    jrc_alta_cno3 = sorted(k for k, g in jrc_cno3.items() if g["nivel"] == "alta")

    anth_nivel = build_cno_to_anthropic_nivel()
    anth_codes_por_nivel = {key: sorted(c for c, n in anth_nivel.items() if n == key)
                             for key, _ in ANTHROPIC_CORTES}

    oit_name_by_code = {c: n for c, n, _ in OCUPACIONES_ORDEN}

    breakdowns = {}
    for metric in ("contratos", "desempleo"):
        cm, ca = computed[(metric, "mes")], computed[(metric, "anio")]
        cm_rows = list(cm["by_cno"].values())
        ca_rows = list(ca["by_cno"].values())
        oit_map_m = {r["cno"]: EXPOSURE_TO_TIER.get(r["exposure"]) for r in cm_rows}
        oit_map_a = {r["cno"]: EXPOSURE_TO_TIER.get(r["exposure"]) for r in ca_rows}
        oit_base_m, oit_base_a = bases_by_map(cm_rows, oit_map_m), bases_by_map(ca_rows, oit_map_a)

        # ---------------- índice OIT (igual que antes) ----------------
        oit_tiers_out = []
        for key, label, color, short_label in TIER_META:
            occ_codes = None
            if key == "maxima":
                occ_codes = [c for c, _, t in OCUPACIONES_ORDEN if t == "maxima"]
            elif key == "significativa":
                occ_codes = [c for c, _, t in OCUPACIONES_ORDEN if t == "significativa"]
            oit_tiers_out.append({
                "key": key, "label": label, "color": color, "short_label": short_label,
                "mes": cm["tiers"][key], "anio": ca["tiers"][key],
                "base": {"mes": oit_base_m.get(key), "anio": oit_base_a.get(key)},
                "occupations": occ_codes,
            })
        oit_resumen = resumen_out_dict(cm["resumen"], ca["resumen"])

        # ---------------- índice OCDE (ficticio, de prueba) ----------------
        ocde_mes_por_nivel = aggregate_by_map(cm_rows, ocde_nivel)
        ocde_anio_por_nivel = aggregate_by_map(ca_rows, ocde_nivel)
        ocde_tiers_out = []
        for key, label, color, short_label in OCDE_TIER_META:
            occ_codes = ocde_codes_por_nivel[key] if key in ("nivel5", "nivel4") else None
            ocde_tiers_out.append({
                "key": key, "label": label, "color": color, "short_label": short_label,
                "mes": ocde_mes_por_nivel.get(key), "anio": ocde_anio_por_nivel.get(key),
                "base": {"mes": bases_by_map(cm_rows, ocde_nivel).get(key), "anio": bases_by_map(ca_rows, ocde_nivel).get(key)},
                "occupations": occ_codes,
            })
        ocde_resumen = resumen_out_dict(
            compute_resumen(cm_rows, set(ocde_codes_por_nivel["nivel5"])),
            compute_resumen(ca_rows, set(ocde_codes_por_nivel["nivel5"])),
        )

        # ---------------- índice JRC (real) ----------------
        jrc_mes_por_nivel = aggregate_by_map(cm_rows, jrc_nivel)
        jrc_anio_por_nivel = aggregate_by_map(ca_rows, jrc_nivel)
        jrc_tiers_out = []
        for key, label, color, short_label in JRC_TIER_META:
            # "Alta" lista códigos ISCO a 3 dígitos (agregando todas las CNO
            # que caen en cada uno), NO códigos CNO de 4 dígitos: JRC
            # clasifica a ese nivel de agregación, no al de la OIT/OCDE.
            occ_codes = jrc_alta_cno3 if key == "alta" else None
            jrc_tiers_out.append({
                "key": key, "label": label, "color": color, "short_label": short_label,
                "mes": jrc_mes_por_nivel.get(key), "anio": jrc_anio_por_nivel.get(key),
                "base": {"mes": bases_by_map(cm_rows, jrc_nivel).get(key), "anio": bases_by_map(ca_rows, jrc_nivel).get(key)},
                "occupations": occ_codes,
            })
        jrc_resumen = resumen_out_dict(
            compute_resumen(cm_rows, set(jrc_alta_codes)),
            compute_resumen(ca_rows, set(jrc_alta_codes)),
        )

        # ---------------- índice Anthropic (real, numérico en 4 tramos) ----------------
        anth_mes_por_nivel = aggregate_by_map(cm_rows, anth_nivel)
        anth_anio_por_nivel = aggregate_by_map(ca_rows, anth_nivel)
        anth_tiers_out = []
        for key, label, color, short_label in ANTHROPIC_TIER_META:
            occ_codes = anth_codes_por_nivel[key] if key in ("muyalta", "alta") else None
            anth_tiers_out.append({
                "key": key, "label": label, "color": color, "short_label": short_label,
                "mes": anth_mes_por_nivel.get(key), "anio": anth_anio_por_nivel.get(key),
                "base": {"mes": bases_by_map(cm_rows, anth_nivel).get(key), "anio": bases_by_map(ca_rows, anth_nivel).get(key)},
                "occupations": occ_codes,
            })
        anth_resumen = resumen_out_dict(
            compute_resumen(cm_rows, set(anth_codes_por_nivel["muyalta"])),
            compute_resumen(ca_rows, set(anth_codes_por_nivel["muyalta"])),
        )

        # ---------------- occupations: compartidas entre índices ----------------
        # Códigos CNO (4 dígitos): los que algún tier expandible de OIT/OCDE
        # vaya a mostrar (las 48 de la OIT + las 65 del nivel 5/4 ficticio de
        # OCDE). JRC NO se lista aquí: no hay colisión de claves porque sus
        # códigos son ISCO a 3 dígitos (ver bloque siguiente), nunca CNO de 4.
        needed_codes = ({c for c, _, _ in OCUPACIONES_ORDEN}
                         | set(ocde_codes_por_nivel["nivel5"])
                         | set(ocde_codes_por_nivel["nivel4"])
                         | set(anth_codes_por_nivel["muyalta"])
                         | set(anth_codes_por_nivel["alta"]))
        occupations_out = {}
        for code in sorted(needed_codes):  # orden fijo: el JSON sale igual en cada ejecución
            row_mes = cm["by_cno"].get(code)
            row_anio = ca["by_cno"].get(code)
            name = oit_name_by_code.get(code)
            if not name:
                src = row_mes or row_anio
                name = (src["desc"] if src and src["desc"] else code)
            bm = occupation_block(row_mes)
            ba = occupation_block(row_anio)
            occupations_out[code] = {
                "name": name,
                "mes": bm["total"], "anio": ba["total"],
                "base": {"mes": base_of_rows([row_mes]) if row_mes else None,
                         "anio": base_of_rows([row_anio]) if row_anio else None},
                "sexo": {"mes": [bm["hombres"], bm["mujeres"]], "anio": [ba["hombres"], ba["mujeres"]]},
                "edad": {"mes": [bm["menos30"], bm["entre3045"], bm["mas45"]],
                         "anio": [ba["menos30"], ba["entre3045"], ba["mas45"]]},
                "cnae": {"mes": [None] * 21, "anio": [None] * 21},
            }

        # Grupos CNO (3 dígitos) de "Alta" en JRC: una fila por grupo (o por parte
        # del grupo si es mixto), sumando los totales de sus CNO (suma -> %, nunca media).
        for key3 in jrc_alta_cno3:
            g = jrc_cno3[key3]
            rows_mes = [cm["by_cno"][c] for c in g["cnos"] if c in cm["by_cno"]]
            rows_anio = [ca["by_cno"][c] for c in g["cnos"] if c in ca["by_cno"]]
            bm = aggregate_occ_block(rows_mes)
            ba = aggregate_occ_block(rows_anio)
            occupations_out[key3] = {
                "code": g["code"],
                "cnos": list(g["cnos"]),  # CNO (4 díg.) que componen el grupo: sirve para cruzar índices
                "name": g["name"],
                "mes": bm["total"], "anio": ba["total"],
                "base": {"mes": base_of_rows(rows_mes), "anio": base_of_rows(rows_anio)},
                "sexo": {"mes": [bm["hombres"], bm["mujeres"]], "anio": [ba["hombres"], ba["mujeres"]]},
                "edad": {"mes": [bm["menos30"], bm["entre3045"], bm["mas45"]],
                         "anio": [ba["menos30"], ba["entre3045"], ba["mas45"]]},
                "cnae": {"mes": [None] * 21, "anio": [None] * 21},
            }

        breakdowns[metric] = {
            "national": {"mes": cm["national"], "anio": ca["national"]},
            "national_base": {"mes": base_of_rows(cm_rows), "anio": base_of_rows(ca_rows)},
            "occupations": occupations_out,
            "indices": {
                "oit": {"tiers": oit_tiers_out, "resumen": oit_resumen},
                "ocde": {"tiers": ocde_tiers_out, "resumen": ocde_resumen},
                "jrc": {"tiers": jrc_tiers_out, "resumen": jrc_resumen},
                "anthropic": {"tiers": anth_tiers_out, "resumen": anth_resumen},
            },
        }

    return {
        "metrics": metrics,
        "compare_text": compare_text,
        "breakdowns": breakdowns,
        "cnae_sections": CNAE_SECTIONS,
    }, computed


# ---------------------------------------------------------------------------
# (Opcional) Excel enriquecido para auditar a ojo
# ---------------------------------------------------------------------------

TIER_ROW_LABELS = [
    ("Gran total", None),
    ("Total Gradiente 4", "Gradient 4"),
    ("Total Gradiente 3", "Gradient 3"),
    ("Total Gradiente 2", "Gradient 2"),
    ("Total Gradiente 1", "Gradient 1"),
    ("Total Minimal Exposure", "Minimal Exposure"),
    ("Total Not Exposed", "Not Exposed"),
]


def write_enriched_workbook(src_path, out_path):
    """Genera una copia del original con 7 filas nuevas bajo la cabecera
    (Gran total + los 6 niveles) y columnas % con fórmulas de Excel de
    verdad, para que Roberto lo pueda abrir y revisar a ojo. Este fichero
    NO se vuelve a leer por el script: el JSON sale siempre de los valores
    calculados en Python a partir del original."""
    wb = openpyxl.load_workbook(src_path, data_only=False)
    n_new = len(TIER_ROW_LABELS)  # filas nuevas insertadas

    for (metric, period), (sheet_name, cols) in SHEETS.items():
        ws = buscar_hoja(wb, sheet_name)
        last_col = ws.max_column
        last_row_original = ws.max_row

        ws.insert_rows(2, amount=n_new)

        # Cabeceras de columnas %: se insertan justo a la derecha de cada
        # bloque (total, sexo, edad) -> igual que en "Datos radar SEPE final"
        is_parados = cols is RAW_COLS_PARADOS
        # columnas (letra) antes de insertar nada de % (bloques ya fijos)
        # Para Contratadas ya existían columnas % en el original -> no se
        # duplican, se limita a rellenar las 7 filas nuevas con fórmulas.
        c = cols
        first_data_row = 2 + n_new
        last_data_row = last_row_original + n_new

        def L(col_idx):
            return get_column_letter(col_idx)

        # Fila "Gran total" y una fila por nivel de exposición
        for i, (label, exposure_value) in enumerate(TIER_ROW_LABELS):
            r = 2 + i
            ws.cell(row=r, column=c["cno"], value=label)
            ws.cell(row=r, column=c["desc"], value=label)
            ws.cell(row=r, column=c["exposure"], value=label)

            if exposure_value is None:
                rango = f"{first_data_row}:{last_data_row}"
                cond = None
            else:
                # suma condicional por nivel de exposición
                cond = exposure_value

            def sumif(field):
                col = L(c[field])
                total_col = L(c["exposure"])
                if cond is None:
                    return f"=SUM({col}{first_data_row}:{col}{last_data_row})"
                return f'=SUMIF({total_col}{first_data_row}:{total_col}{last_data_row},"{cond}",{col}{first_data_row}:{col}{last_data_row})'

            ws.cell(row=r, column=c["total1"], value=sumif("total1"))
            ws.cell(row=r, column=c["total2"], value=sumif("total2"))
            ws.cell(row=r, column=c["hombre1"], value=sumif("hombre1"))
            ws.cell(row=r, column=c["mujer1"], value=sumif("mujer1"))
            ws.cell(row=r, column=c["hombre2"], value=sumif("hombre2"))
            ws.cell(row=r, column=c["mujer2"], value=sumif("mujer2"))
            ws.cell(row=r, column=c["e1_30"], value=sumif("e1_30"))
            ws.cell(row=r, column=c["e1_3045"], value=sumif("e1_3045"))
            ws.cell(row=r, column=c["e1_45"], value=sumif("e1_45"))
            ws.cell(row=r, column=c["e2_30"], value=sumif("e2_30"))
            ws.cell(row=r, column=c["e2_3045"], value=sumif("e2_3045"))
            ws.cell(row=r, column=c["e2_45"], value=sumif("e2_45"))

        # % por fila, para TODAS las filas (las 7 nuevas + cada ocupación)
        for r in range(2, last_data_row + 1):
            t1, t2 = L(c["total1"]), L(c["total2"])
            h1, h2 = L(c["hombre1"]), L(c["hombre2"])
            m1, m2 = L(c["mujer1"]), L(c["mujer2"])
            e30_1, e30_2 = L(c["e1_30"]), L(c["e2_30"])
            e3045_1, e3045_2 = L(c["e1_3045"]), L(c["e2_3045"])
            e45_1, e45_2 = L(c["e1_45"]), L(c["e2_45"])
            # columnas % al final de la fila (no pisan nada del original)
            col_pct_total = last_col + 1
            ws.cell(row=r, column=col_pct_total, value=f"=IFERROR({t2}{r}/{t1}{r}-1,\"\")")
            ws.cell(row=r, column=col_pct_total + 1, value=f"=IFERROR({h2}{r}/{h1}{r}-1,\"\")")
            ws.cell(row=r, column=col_pct_total + 2, value=f"=IFERROR({m2}{r}/{m1}{r}-1,\"\")")
            ws.cell(row=r, column=col_pct_total + 3, value=f"=IFERROR({e30_2}{r}/{e30_1}{r}-1,\"\")")
            ws.cell(row=r, column=col_pct_total + 4, value=f"=IFERROR({e3045_2}{r}/{e3045_1}{r}-1,\"\")")
            ws.cell(row=r, column=col_pct_total + 5, value=f"=IFERROR({e45_2}{r}/{e45_1}{r}-1,\"\")")
        ws.cell(row=1, column=last_col + 1, value="% TOTAL")
        ws.cell(row=1, column=last_col + 2, value="% HOMBRES")
        ws.cell(row=1, column=last_col + 3, value="% MUJERES")
        ws.cell(row=1, column=last_col + 4, value="% <30 años")
        ws.cell(row=1, column=last_col + 5, value="% 30-45 años")
        ws.cell(row=1, column=last_col + 6, value="% >45 años")

    wb.save(out_path)


# ---------------------------------------------------------------------------
# main
# ---------------------------------------------------------------------------

# ---------------------------------------------------------------------------
# BORRADOR DEL ANÁLISIS (analisis_borrador.md)
# Escribe, por reglas, un borrador de los textos de "Análisis" a partir de datos_radar.json,
# más una ficha de hechos al final para revisarlo. NO es el texto final: el comentarista
# humano (opcionalmente con IA) lo revisa y lo guarda como analisis_final.md.
# Este script NUNCA toca analisis_final.md.
#
# ENFOQUE (mientras solo hay contratos y paro, sin afiliados): el comentario se centra en las
# ocupaciones con DOBLE SEÑAL NEGATIVA = menos contratación Y más paro, siempre que sean
# relevantes (bases mínimas y variaciones de al menos SENAL_MIN). Las que la tienen en los
# tres índices (OIT, JRC y Anthropic) van primero; después las de dos; dentro de cada grupo,
# por PERSONAS AFECTADAS (contratos perdidos + parados añadidos, en cifras absolutas).
# Para cruzar índices se usan las CNO de 4 dígitos: un grupo JRC (CNO a 3 dígitos) coincide
# con una ocupación OIT/Anthropic si esa ocupación es una de las CNO que lo componen.
# ---------------------------------------------------------------------------

# Parámetros del borrador (editables)
BASE_MIN = {"contratos": 100, "desempleo": 250}  # bases mínimas de partida para citar una ocupación
SENAL_MIN = 2.0          # variación mínima (%) para hablar de "menos contratación" / "más paro"
DOBLE_MAX = 5            # máximo de ocupaciones con doble señal que se listan por índice
PESO_AVISO = 0.40        # si una ocupación pesa/explica >= 40% del grupo, se avisa
BASE_PEQUENA = 1000      # contratos de partida por debajo de los que se muestra la base entre paréntesis
BAJADAS_MAX = 2          # máximo de "mayores caídas de contratación" (por contratos perdidos) que se citan además de la doble señal
SEGUNDO_MAX = 3          # máximo de ocupaciones con doble señal que se citan del segundo tramo (OIT Gradiente 3, Anthropic alta), más breve que el primero
DIF_SEXO = 2.0           # puntos de diferencia entre hombres y mujeres a partir de los cuales se comenta el sexo
DIF_EDAD = 2.0           # puntos entre el tramo de edad más alto y el más bajo a partir de los cuales se comenta la edad
EMPATE = 0.3             # diferencia (en puntos) por debajo de la cual dos tramos de edad se dan por iguales


def mostrar_parametros():
    """Imprime por pantalla los parámetros con los que se genera el borrador."""
    f = lambda x: f"{x:g}".replace(".", ",")
    f1 = lambda x: f"{x:.1f}".replace(".", ",")
    print("  Parámetros del borrador (se editan en el bloque «Parámetros del borrador» de este script):")
    print(f"    BASE_MIN               base mínima de partida: {BASE_MIN['contratos']} contratos y "
          f"{BASE_MIN['desempleo']} parados")
    print(f"    SENAL_MIN              variación mínima: {f1(SENAL_MIN)}% "
          f"(contratación <= -{f1(SENAL_MIN)}% y paro >= +{f1(SENAL_MIN)}%)")
    print(f"    DOBLE_MAX              máximo de ocupaciones con doble señal por índice: {DOBLE_MAX}")
    print(f"    PESO_AVISO             aviso de concentración si una ocupación pesa/explica >= {PESO_AVISO:.0%}".replace("%", " %"))
    print(f"    BASE_PEQUENA           se anota la base si hay menos de {BASE_PEQUENA} contratos de partida")
    print(f"    BAJADAS_MAX            máximo de mayores caídas de contratación citadas (por contratos perdidos): {BAJADAS_MAX}")
    print(f"    SEGUNDO_MAX            máximo de ocupaciones citadas del segundo tramo (OIT Gradiente 3, Anthropic alta): {SEGUNDO_MAX}")
    print(f"    DIF_SEXO               se comenta el sexo si hombres y mujeres difieren >= {f1(DIF_SEXO)} puntos")
    print(f"    DIF_EDAD               se comenta la edad si los tramos difieren >= {f1(DIF_EDAD)} puntos")
    print(f"    EMPATE                 dos tramos de edad se agrupan si difieren menos de {f(EMPATE)} puntos")

# Tramo que se analiza de cada índice (el de mayor exposición) y cómo se nombra.
INDICES = {
    "oit": dict(tier="maxima", noun="ocupaciones", art="las", uno="una", nombre="la OIT",
                descr="máxima exposición de la OIT"),
    # Segundo tramo de la OIT y de Anthropic: se analizan aparte (y más breve) dentro de su sección
    "oit3": dict(indice="oit", tier="significativa", noun="ocupaciones", art="las", uno="una", nombre="la OIT",
                 descr="exposición significativa de la OIT (Gradiente 3)",
                 intro="el Gradiente 3 de la OIT (exposición significativa, {n} ocupaciones)",
                 ficha="Gradiente 3 de la OIT"),
    "anthropic2": dict(indice="anthropic", tier="alta", noun="ocupaciones", art="las", uno="una", nombre="Anthropic",
                       descr="exposición alta según Anthropic (índice 0,20–0,30)",
                       intro="la exposición alta según Anthropic (índice 0,20–0,30; {n} ocupaciones)",
                       ficha="Exposición alta según Anthropic"),
    "jrc": dict(tier="alta", noun="grupos", art="los", uno="uno", nombre="el JRC",
                descr="alta exposición del JRC (ocupaciones CNO a 3 dígitos)"),
    "anthropic": dict(tier="muyalta", noun="ocupaciones", art="las", uno="una", nombre="Anthropic",
                      descr="exposición muy alta según Anthropic (índice ≥ 0,30)"),
}

GRUPOS_EDAD = {"Menos de 30": "los **menores de 30**", "30–45": "los de **30 a 45 años**",
               "Más de 45": "los **mayores de 45**"}

# "no clasificados bajo otros epígrafes" -> abreviatura oficial de la CNO: n.c.o.p.
_NCOP = re.compile(r"\s+no clasificad[oa]s?\s+bajo\s+otros\s+epígrafes", re.I)


def _nombre(n):
    return _NCOP.sub(" n.c.o.p.", n)


def abreviar_ncop_en_json(data):
    """Escribe 'n.c.o.p.' en lugar de 'no clasificados bajo otros epígrafes' en los nombres de las
    ocupaciones del JSON (así lo ven igual la web y el borrador). Devuelve cuántos nombres cambia."""
    cambiados = 0
    for met in data.get("breakdowns", {}).values():
        for o in (met.get("occupations") or {}).values():
            nuevo = _nombre(o.get("name", ""))
            if nuevo != o.get("name"):
                o["name"] = nuevo
                cambiados += 1
    return cambiados


# ---------------------------------------------------------------------------
# Formato de cifras (coma decimal, signo "-" normal, miles con punto)
# ---------------------------------------------------------------------------
def fp(v, signo=True):
    s = f"{abs(v):.1f}".replace(".", ",")
    if not signo:
        return s + "%"
    return ("+" if v > 0 else "-" if v < 0 else "") + s + "%"


def fn(n):
    return f"{int(round(n)):,}".replace(",", ".")


def _lista(items):
    if len(items) <= 1:
        return "".join(items)
    return ", ".join(items[:-1]) + " y " + items[-1]


# ---------------------------------------------------------------------------
# Acceso a los datos de un índice
# ---------------------------------------------------------------------------
def _filas_indice(data, ix, met):
    """Filas (código, nombre, % mes, bases, CNO que lo componen) del tramo analizado de un índice."""
    b = data["breakdowns"][met]
    tier = next(t for t in b["indices"][INDICES[ix].get("indice", ix)]["tiers"] if t["key"] == INDICES[ix]["tier"])
    filas = []
    for code in tier["occupations"] or []:
        o = b["occupations"].get(code)
        if o and o.get("mes") is not None:
            base = (o.get("base") or {}).get("mes") or [0, 0]
            filas.append(dict(code=o.get("code", code), key=code, name=o["name"], pct=o["mes"],
                              b1=base[0], b2=base[1], cnos=list(o.get("cnos") or [code])))
    return tier, filas


def _mes_etiqueta(data):
    ct = data.get("compare_text", {}).get("mes", "")
    return (ct.split()[0].lower() if ct else "el mes")


# ---------------------------------------------------------------------------
# Doble señal negativa (menos contratación y más paro) y cruce entre índices
# ---------------------------------------------------------------------------
def _doble_senal(data, ix):
    """{clave: {c: fila contratos, d: fila paro, cnos: set, sev: personas afectadas}} de las ocupaciones
    relevantes con contratación <= -SENAL_MIN y paro >= +SENAL_MIN."""
    _, fc = _filas_indice(data, ix, "contratos")
    _, fd = _filas_indice(data, ix, "desempleo")
    pd = {r["key"]: r for r in fd}
    out = {}
    for c in fc:
        d = pd.get(c["key"])
        if (d and c["pct"] <= -SENAL_MIN and d["pct"] >= SENAL_MIN
                and c["b1"] >= BASE_MIN["contratos"] and d["b1"] >= BASE_MIN["desempleo"]):
            out[c["key"]] = dict(c=c, d=d, cnos=set(c["cnos"]), sev=(c["b1"] - c["b2"]) + (d["b2"] - d["b1"]))
    return out


def _cruzar(dobles):
    """Añade a cada ocupación con doble señal en qué OTROS índices también la tiene ('otros')
    y en cuántos índices en total aparece ('n'). El segundo tramo ('oit3', 'anthropic2') cuenta como el mismo índice que el primero."""
    canon = lambda k: INDICES[k].get("indice", k)
    for ix, items in dobles.items():
        for it in items.values():
            otros = []
            for oy, its in dobles.items():
                if canon(oy) != canon(ix) and canon(oy) not in otros \
                        and any(it["cnos"] & o["cnos"] for o in its.values()):
                    otros.append(canon(oy))
            it["otros"] = otros
            it["n"] = 1 + len(otros)


def _item_doble(it):
    c, d = it["c"], it["d"]
    cont = f"contratación {fp(c['pct'])}"
    if c["b1"] < BASE_PEQUENA:
        cont += f", de {fn(c['b1'])} a {fn(c['b2'])} contratos"
    txt = f"**{_nombre(c['name'])}** ({cont}; paro {fp(d['pct'])})"
    if it["n"] == 3:
        txt += ", que aparece con esta doble señal en los tres índices"
    elif it["n"] == 2:
        txt += f", que también la tiene en {INDICES[it['otros'][0]]['nombre']}"
    return txt


def _ordenar(items):
    return sorted(items.values(), key=lambda it: (-it["n"], -it["sev"]))


# ---------------------------------------------------------------------------
# Sexo y edad
# ---------------------------------------------------------------------------
def _comparar(sust, a_lab, a, b_lab, b, caer, subir):
    """Frase sobre dos grupos (hombres/mujeres)."""
    if (a < 0) != (b < 0) and a != 0 and b != 0:
        neg, pos = ((a_lab, a), (b_lab, b)) if a < 0 else ((b_lab, b), (a_lab, a))
        return f"{sust} {caer} entre {neg[0]} ({fp(neg[1])}) y {subir} entre {pos[0]} ({fp(pos[1])})"
    if abs(a - b) < 0.5:
        return f"{sust} {caer if a < 0 else subir} igual entre {a_lab} ({fp(a)}) y {b_lab} ({fp(b)})"
    mayor, menor = ((a_lab, a), (b_lab, b)) if abs(a) > abs(b) else ((b_lab, b), (a_lab, a))
    return (f"{sust} {caer if a < 0 else subir} más entre {mayor[0]} ({fp(mayor[1])}) "
            f"que entre {menor[0]} ({fp(menor[1])})")


def _par(v1, v2):
    return f"{fp(v1)} en ambos" if round(v1, 1) == round(v2, 1) else f"{fp(v1)} y {fp(v2)}"


def _frase_edad(sust, grupos, caer, subir):
    """Frase sobre tres tramos de edad. grupos: [(etiqueta, valor)]. Los empates se agrupan."""
    pos = [g for g in grupos if g[1] > 0]
    neg = [g for g in grupos if g[1] < 0]
    if pos and neg:
        a = _lista([f"{l} ({fp(v)})" for l, v in sorted(pos, key=lambda g: -g[1])])
        b = _lista([f"{l} ({fp(v)})" for l, v in sorted(neg, key=lambda g: g[1])])
        return f"{sust} {subir} entre {a}; {caer} entre {b}"
    verbo = caer if neg else subir
    a, b, c = sorted(grupos, key=lambda g: -abs(g[1]))
    t_ab = abs(abs(a[1]) - abs(b[1])) < EMPATE
    t_bc = abs(abs(b[1]) - abs(c[1])) < EMPATE
    if t_ab and t_bc:
        vs = sorted(g[1] for g in grupos)
        return f"{sust} {verbo} de forma parecida en los tres tramos (entre {fp(vs[0])} y {fp(vs[-1])})"
    if t_bc:
        return f"{sust} {verbo} más entre {a[0]} ({fp(a[1])}) que entre {b[0]} y {c[0]} ({_par(b[1], c[1])})"
    if t_ab:
        return (f"{sust} {verbo} más entre {a[0]} y {b[0]} ({_par(a[1], b[1])}) "
                f"que entre {c[0]} ({fp(c[1])})")
    return (f"{sust} {verbo} más entre {a[0]} ({fp(a[1])}), seguidos de "
            f"{b[0]} ({fp(b[1])}) y {c[0]} ({fp(c[1])})")


def _etiqueta_edad(label):
    return GRUPOS_EDAD.get(label, label)


# ---------------------------------------------------------------------------
# Borrador de un índice
# ---------------------------------------------------------------------------
def _aviso_concentracion(tc, td, fc, fd):
    """Si una ocupación domina el grupo (por peso en la base o por aportación al cambio)."""
    base_c = (tc.get("base") or {}).get("mes")
    base_d = (td.get("base") or {}).get("mes")
    if not (base_c and base_d):
        return None, None
    pd = {r["key"]: r for r in fd}
    # a) peso en la base
    for r in sorted(fc, key=lambda r: -r["b1"]):
        rd = pd.get(r["key"])
        peso_c = r["b1"] / base_c[0] if base_c[0] else 0
        peso_d = (rd["b1"] / base_d[0]) if (rd and base_d[0]) else 0
        if max(peso_c, peso_d) >= PESO_AVISO:
            txt = (f"Conviene leer el total con cautela: **{_nombre(r['name'])}** concentra el "
                   f"{round(peso_c * 100)}% de la contratación y el {round(peso_d * 100)}% del paro del grupo")
            if rd:
                txt += (f", así que el resultado conjunto depende sobre todo de ella "
                        f"({fp(r['pct'])} en contratación y {fp(rd['pct'])} en paro).")
            else:
                txt += "."
            return txt, r["key"]
    # b) aportación al cambio del grupo (solo si el grupo cambia de forma apreciable: si no, el reparto no significa nada)
    perdida = (base_c[0] - base_c[1]) if tc["mes"] <= -SENAL_MIN else 0
    aumento = (base_d[1] - base_d[0]) if td["mes"] >= SENAL_MIN else 0
    mejor = None
    for r in fc:
        rd = pd.get(r["key"])
        sc = (r["b1"] - r["b2"]) / perdida if perdida > 0 and r["b2"] < r["b1"] else 0
        sd = (rd["b2"] - rd["b1"]) / aumento if (rd and aumento > 0 and rd["b2"] > rd["b1"]) else 0
        if sc > 1 or sd > 1:   # otras ocupaciones compensan: "explica el 150%" no se lee bien
            continue
        if max(sc, sd) >= PESO_AVISO and (mejor is None or max(sc, sd) > mejor[0]):
            mejor = (max(sc, sd), r, rd, sc, sd)
    if mejor:
        _, r, rd, sc, sd = mejor
        partes = []
        if sc >= PESO_AVISO:
            partes.append(f"el {round(sc * 100)}% de los contratos perdidos por el grupo "
                          f"({fn(r['b1'] - r['b2'])} de {fn(perdida)})")
        if sd >= PESO_AVISO:
            partes.append(f"el {round(sd * 100)}% del aumento del paro ({fn(rd['b2'] - rd['b1'])} de {fn(aumento)})")
        return (f"Conviene leer el total con cautela: **{_nombre(r['name'])}** explica por sí sola "
                + " y ".join(partes) + "."), r["key"]
    return None, None


def _borrador_indice(data, ix, dobles):
    cfg = INDICES[ix]
    mes = _mes_etiqueta(data)
    tc, fc = _filas_indice(data, ix, "contratos")
    td, fd = _filas_indice(data, ix, "desempleo")
    nac_c = data["breakdowns"]["contratos"]["national"]["mes"]
    nac_d = data["breakdowns"]["desempleo"]["national"]["mes"]
    c, d = tc["mes"], td["mes"]
    n = len(tc["occupations"] or [])
    pd = {r["key"]: r for r in fd}

    sub = (f"En {cfg['descr'].split(' (')[0]}, la contratación {'cae' if c < 0 else 'sube'} un {fp(c, False)} "
           f"y el paro {'baja' if d < 0 else 'sube'} un {fp(d, False)}")

    # --- párrafo 1: titular frente al total nacional
    p1 = (f"En {mes}, {cfg['art']} {n} {cfg['noun']} de {cfg['descr']} registran una variación interanual "
          f"del {fp(c)} en contratación (total nacional: {fp(nac_c)}) y del {fp(d)} en paro (total: {fp(nac_d)}).")
    pars = [p1]

    # --- párrafo 2: doble señal negativa (menos contratación y más paro)
    items = _ordenar(dobles[ix])
    k = len(items)
    cond = f"con variaciones de al menos un {SENAL_MIN:g}% y bases relevantes"
    if k == 0:
        p2 = (f"Ninguno de {cfg['art']} {n} {cfg['noun']} reúne a la vez menos contratación y más paro "
              f"({cond}).")
    else:
        if k == 1:
            intro = (f"Solo {cfg['uno']} de {cfg['art']} {n} {cfg['noun']} reúne las dos señales negativas a la vez, "
                     f"menos contratación y más paro ({cond}): ")
        else:
            intro = (f"Reúnen las dos señales negativas a la vez, menos contratación y más paro ({cond}), "
                     f"{k} de {cfg['art']} {n} {cfg['noun']}: ")
        p2 = intro + "; ".join(_item_doble(it) for it in items[:DOBLE_MAX]) + "."
        if k > DOBLE_MAX:
            p2 += f" Otros {k - DOBLE_MAX} tienen la misma combinación, con menor intensidad."
    pars.append(p2)
    citadas = {it["c"]["key"] for it in items[:DOBLE_MAX]}

    # --- párrafo 3: otros hechos destacables (sin repetir lo ya citado)
    extra = []
    por_d = sorted((r for r in fd if r["b1"] >= BASE_MIN["desempleo"]), key=lambda r: r["pct"])
    if por_d and por_d[-1]["pct"] > 0 and por_d[-1]["key"] not in citadas:
        r = por_d[-1]
        rc = next((x for x in fc if x["key"] == r["key"]), None)
        s = f"El mayor aumento del paro es el de **{_nombre(r['name'])}** ({fp(r['pct'])})"
        if rc and rc["pct"] >= SENAL_MIN:
            s += f", aunque su contratación sube ({fp(rc['pct'])})"
        elif rc and abs(rc["pct"]) < SENAL_MIN:
            s += f", con la contratación estable ({fp(rc['pct'])})"
        extra.append(s + ".")
    # mayores caídas de contratación (por contratos perdidos), sin repetir doble señal ni la ocupación del aviso
    aviso, clave_aviso = _aviso_concentracion(tc, td, fc, fd)
    bajadas = sorted((r for r in fc if r["b1"] >= BASE_MIN["contratos"] and r["pct"] <= -SENAL_MIN
                      and r["key"] not in dobles[ix] and r["key"] != clave_aviso),
                     key=lambda r: -(r["b1"] - r["b2"]))[:BAJADAS_MAX]
    if bajadas:
        txts = []
        for r in bajadas:
            t = f"**{_nombre(r['name'])}** ({fp(r['pct'])}, de {fn(r['b1'])} a {fn(r['b2'])} contratos"
            rd = pd.get(r["key"])
            if rd and abs(rd["pct"]) >= SENAL_MIN:
                t += f"; paro {fp(rd['pct'])}"
            txts.append(t + ")")
        extra.append("Caen con más fuerza en contratación, en contratos perdidos, " + _lista(txts) + ".")
    if extra:
        pars.append(" ".join(extra))

    # --- aviso de concentración
    if aviso:
        pars.append(aviso)

    # --- sexo y edad
    rc = data["breakdowns"]["contratos"]["indices"][ix]["resumen"]
    rd_ = data["breakdowns"]["desempleo"]["indices"][ix]["resumen"]
    sx_c = {g["label"]: g["mes"] for g in rc["sexo"]}
    sx_d = {g["label"]: g["mes"] for g in rd_["sexo"]}
    if (sx_c.get("Hombres") is not None and sx_d.get("Hombres") is not None
            and sx_c.get("Mujeres") is not None and sx_d.get("Mujeres") is not None):
        fr = []
        if abs(sx_c["Hombres"] - sx_c["Mujeres"]) >= DIF_SEXO:
            fr.append(_comparar("la contratación", "los **hombres**", sx_c["Hombres"], "las **mujeres**",
                                sx_c["Mujeres"], "cae", "sube"))
        if abs(sx_d["Hombres"] - sx_d["Mujeres"]) >= DIF_SEXO:
            fr.append(_comparar("el paro", "los hombres", sx_d["Hombres"], "las mujeres", sx_d["Mujeres"],
                                "baja", "sube"))
        if fr:
            pars.append("Por sexo, " + "; ".join(fr) + ".")
    ed_c = [(_etiqueta_edad(g["label"]), g["mes"]) for g in rc["edad"] if g["mes"] is not None]
    ed_d = [(_etiqueta_edad(g["label"]).replace("**", ""), g["mes"]) for g in rd_["edad"] if g["mes"] is not None]
    fr = []
    if len(ed_c) == 3 and max(v for _, v in ed_c) - min(v for _, v in ed_c) >= DIF_EDAD:
        fr.append(_frase_edad("la contratación", ed_c, "cae", "sube") + ".")
    if len(ed_d) == 3 and max(v for _, v in ed_d) - min(v for _, v in ed_d) >= DIF_EDAD:
        t = _frase_edad("el paro", ed_d, "baja", "sube").replace("**", "") + "."
        fr.append(t[0].upper() + t[1:] if fr else t)
    if fr:
        pars.append("Por edad, " + " ".join(fr))
    return sub, pars


SEGUNDO_TRAMO = {"oit": "oit3", "anthropic": "anthropic2"}


def _parrafo_segundo(data, dobles, clave):
    """Párrafo breve sobre el segundo tramo de un índice (el primero es el principal)."""
    cfg = INDICES[clave]
    tc, _ = _filas_indice(data, clave, "contratos")
    td, _ = _filas_indice(data, clave, "desempleo")
    nac_c = data["breakdowns"]["contratos"]["national"]["mes"]
    nac_d = data["breakdowns"]["desempleo"]["national"]["mes"]
    n = len(tc["occupations"] or [])
    txt = (f"En {cfg['intro'].format(n=n)}, la variación interanual es del "
           f"{fp(tc['mes'])} en contratación (total nacional: {fp(nac_c)}) y del {fp(td['mes'])} en paro "
           f"(total: {fp(nac_d)}).")
    items = _ordenar(dobles[clave])
    k = len(items)
    if k == 0:
        txt += " Ninguna reúne a la vez menos contratación y más paro."
    else:
        txt += (f" Reúnen las dos señales negativas a la vez {k} de ellas; "
                + ("destaca " if min(k, SEGUNDO_MAX) == 1 else "destacan ")
                + "; ".join(_item_doble(it) for it in items[:SEGUNDO_MAX]) + ".")
        if k > SEGUNDO_MAX:
            txt += f" Otras {k - SEGUNDO_MAX} comparten esa combinación."
    return txt


def _ficha(data, ix, dobles):
    """Ficha de hechos para quien revise: doble señal, y todas las filas con % y bases."""
    out = []
    items = _ordenar(dobles[ix])
    out.append(f"**Doble señal negativa** (contratación <= -{SENAL_MIN:g}% y paro >= +{SENAL_MIN:g}%, "
               f"bases >= {BASE_MIN['contratos']} contratos y {BASE_MIN['desempleo']} parados): {len(items)}")
    for it in items:
        c, d = it["c"], it["d"]
        ixs = ", ".join(["esta"] + [INDICES[o]["nombre"] for o in it["otros"]]) if it["otros"] else "solo este índice"
        out.append(f"- {c['code']} {c['name']}: contratación {fp(c['pct'])} ({fn(c['b1'])} → {fn(c['b2'])}), "
                   f"paro {fp(d['pct'])} ({fn(d['b1'])} → {fn(d['b2'])}) — personas afectadas: {fn(it['sev'])} "
                   f"— en {it['n']} índice(s): {ixs}")
    out.append("")
    for met in ("contratos", "desempleo"):
        tier, filas = _filas_indice(data, ix, met)
        base = (tier.get("base") or {}).get("mes")
        out.append(f"**{met}** — tramo: {fp(tier['mes'])}" + (f" (base {fn(base[0])} → {fn(base[1])})" if base else "")
                   + f"; total nacional: {fp(data['breakdowns'][met]['national']['mes'])}")
        for r in sorted(filas, key=lambda r: r["pct"]):
            marca = "  ⚠ base pequeña" if r["b1"] < BASE_MIN[met] else ""
            out.append(f"- {r['code']} {r['name']}: {fp(r['pct'])} ({fn(r['b1'])} → {fn(r['b2'])}){marca}")
        out.append("")
    return "\n".join(out)


def borrador(data):
    """-> texto Markdown con el borrador de los tres índices + ficha de hechos."""
    ct = data.get("compare_text", {}).get("mes", "")
    L = ["# Análisis del Radar IA-Empleo — BORRADOR generado por el script",
         "",
         "<!-- Revisar, retocar y guardar como analisis_final.md (este borrador se regenera en cada ejecución; analisis_final.md, nunca). -->",
         f"<!-- periodo: {ct} -->", ""]
    dobles = {ix: _doble_senal(data, ix) for ix in INDICES}
    _cruzar(dobles)
    for ix in ("oit", "jrc", "anthropic"):
        sub, pars = _borrador_indice(data, ix, dobles)
        if ix in SEGUNDO_TRAMO:
            pars.append(_parrafo_segundo(data, dobles, SEGUNDO_TRAMO[ix]))
        L += [f"## {ix}", f"### {sub}", ""]
        for p in pars:
            L += [p, ""]
    L += ["## ocde", "", "<!-- OCDE: en pruebas, sin análisis -->", ""]
    L += ["---", "", "# FICHA DE HECHOS (apoyo para revisar; el parser ignora estas secciones)", ""]
    for ix in ("oit", "jrc", "anthropic"):
        L += [f"## hechos-{ix}", "", _ficha(data, ix, dobles)]
        if ix in SEGUNDO_TRAMO:
            L += [f"### {INDICES[SEGUNDO_TRAMO[ix]]['ficha']}", "", _ficha(data, SEGUNDO_TRAMO[ix], dobles)]
    return "\n".join(L)


def generar_borrador(json_path, salida):
    with open(json_path, encoding="utf-8") as f:
        data = json.load(f)
    with open(salida, "w", encoding="utf-8") as f:
        f.write(borrador(data))
    return salida


def main():
    args = sys.argv[1:]
    enriquecido = "--enriquecido" in args
    if enriquecido:
        args.remove("--enriquecido")

    # Por defecto: "Datos radar SEPE original.xlsx" (con espacios); si no existe,
    # prueba también la versión con guiones bajos.
    if len(args) >= 1:
        origen = args[0]
    else:
        origen = "Datos radar SEPE original.xlsx"
        if not Path(origen).exists() and Path("Datos_radar_SEPE_original.xlsx").exists():
            origen = "Datos_radar_SEPE_original.xlsx"
    salida_json = args[1] if len(args) >= 2 else "datos_radar.json"

    if not Path(origen).exists():
        raise SystemExit(f"No encuentro el fichero de origen: {origen}")

    # Dos lecturas del mismo fichero: una con los valores calculados por
    # Excel (data_only=True, para todas las cifras) y otra con las fórmulas
    # tal cual (data_only=False, solo para distinguir texto de fórmula en
    # la columna Exposure — ver build_exposure_map).
    wb_values = openpyxl.load_workbook(origen, data_only=True)
    wb_formulas = openpyxl.load_workbook(origen, data_only=False)
    data, computed = build_json(wb_values, wb_formulas)
    n_ncop = abreviar_ncop_en_json(data)
    print(f"Nombres con «no clasificados bajo otros epígrafes» abreviados a n.c.o.p.: {n_ncop}")

    with open(salida_json, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    def tier_val(metric, idx, key, period):
        for t in data["breakdowns"][metric]["indices"][idx]["tiers"]:
            if t["key"] == key:
                return t[period]
        return None

    print(f"OK -> {salida_json}")
    print(f"  Total nacional contratos: mes={data['metrics']['contratos']['mes']['national']}  "
          f"año={data['metrics']['contratos']['anio']['national']}")
    print(f"  OIT Gradiente 4 contratos:  mes={tier_val('contratos','oit','maxima','mes')}  "
          f"año={tier_val('contratos','oit','maxima','anio')}")
    print(f"  OCDE Nivel 5 contratos (ficticio): mes={tier_val('contratos','ocde','nivel5','mes')}  "
          f"año={tier_val('contratos','ocde','nivel5','anio')}")
    print(f"  Total nacional desempleo: mes={data['metrics']['desempleo']['mes']['national']}  "
          f"año={data['metrics']['desempleo']['anio']['national']}")
    print(f"  OIT Gradiente 4 desempleo:  mes={tier_val('desempleo','oit','maxima','mes')}  "
          f"año={tier_val('desempleo','oit','maxima','anio')}")
    print(f"  OCDE Nivel 5 desempleo (ficticio): mes={tier_val('desempleo','ocde','nivel5','mes')}  "
          f"año={tier_val('desempleo','ocde','nivel5','anio')}")
    print(f"  Anthropic muy alta contratos: mes={tier_val('contratos','anthropic','muyalta','mes')}  "
          f"año={tier_val('contratos','anthropic','muyalta','anio')}  | desempleo: "
          f"mes={tier_val('desempleo','anthropic','muyalta','mes')}  año={tier_val('desempleo','anthropic','muyalta','anio')}")
    _an = build_cno_to_anthropic_nivel()
    print("  Anthropic: CNO por tramo -> " + ", ".join(
        f"{k}={sum(1 for n in _an.values() if n == k)}" for k, _ in ANTHROPIC_CORTES))
    print(f"  compare_text: {data['compare_text']}")

    # Borrador del análisis (analisis_final.md no se toca nunca)
    borrador_md = Path(salida_json).resolve().parent / "analisis_borrador.md"
    mostrar_parametros()
    generar_borrador(salida_json, str(borrador_md))
    print(f"OK -> {borrador_md} (borrador del análisis; revisar y guardar como analisis_final.md)")
    print("  Recuerda: analisis_final.md NO se ha tocado; si es de otro mes, hay que sustituirlo.")
    print("  Para comprobar que cuadra con estos datos: python comprobar_analisis_final.py")

    if enriquecido:
        salida_xlsx = args[2] if len(args) >= 3 else "Datos_radar_SEPE_final_generado.xlsx"
        write_enriched_workbook(origen, salida_xlsx)
        print(f"OK -> {salida_xlsx} (fichero de auditoría, no se usa para el JSON)")


if __name__ == "__main__":
    main()
