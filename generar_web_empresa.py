# -*- coding: utf-8 -*-
"""
generar_web.py
Genera todos los HTML de la Porra Mundial 2026 a partir del Excel maestro.

Páginas generadas:
  - index.html          → Clasificación general
  - resultados.html     → Calendario y resultados de partidos
  - participantes.html  → Lista de participantes con enlaces
  - fer.html, pedro.html, etc. → Página individual de cada participante
"""

import pandas as pd
from datetime import datetime
import re
import warnings
warnings.filterwarnings("ignore")

# ─────────────────────────────────────────────
# CONFIGURACIÓN
# ─────────────────────────────────────────────

ARCHIVO_EXCEL = "Porra Mundial 2026.xlsx"

# Hojas individuales de participantes que existen en el Excel
HOJAS_PARTICIPANTES = [
    "CAROLINA",
    "COKE",
    "FER_LOBO",
    "FRAN",
    "GEN",
    "JOAN",
    "JOSE",
    "ROBERTO"
]

# Nombre para mostrar y nombre de archivo HTML para cada hoja
PARTICIPANTES_INFO = {
    "CAROLINA": {"display": "Carolina", "archivo": "carolina.html"},
    "COKE":     {"display": "Coke",     "archivo": "coke.html"},
    "FER_LOBO": {"display": "Fer Lobo", "archivo": "fer_lobo.html"},
    "FRAN":     {"display": "Fran",     "archivo": "fran.html"},
    "GEN":      {"display": "Gen",      "archivo": "gen.html"},
    "JOAN":     {"display": "Joan",     "archivo": "joan.html"},
    "JOSE":     {"display": "Jose",     "archivo": "jose.html"},
    "ROBERTO":  {"display": "Roberto",  "archivo": "roberto.html"},
}
    # # Participantes sin hoja propia (aún no han enviado su porra)
    # "anto":   {"display": "Anto",   "archivo": None},
    # "trebo":  {"display": "Trebo",  "archivo": None},
    # "Santi": {"display": "Santi", "archivo": None},
    #}

# ─────────────────────────────────────────────
# CSS COMPARTIDO (tema oscuro, estilo Mundial)
# ─────────────────────────────────────────────

CSS = """
  * { box-sizing: border-box; margin: 0; padding: 0; }

  body {
    background-color: #0d1b2a;
    color: #e2e8f0;
    font-family: 'Segoe UI', Arial, sans-serif;
    min-height: 100vh;
    display: flex;
    flex-direction: column;
  }

  /* ── HEADER ── */
  header {
    background: linear-gradient(135deg, #7f1d1d 0%, #991b1b 50%, #7c2d12 100%);
    padding: 24px 20px 18px;
    text-align: center;
    border-bottom: 3px solid #f59e0b;
    box-shadow: 0 4px 20px rgba(0,0,0,0.5);
  }
  header h1 {
    font-size: clamp(1.4rem, 4vw, 2.2rem);
    color: #fbbf24;
    letter-spacing: 2px;
    text-shadow: 0 2px 8px rgba(0,0,0,0.6);
    font-weight: 900;
  }
  header p.subtitulo {
    color: #fca5a5;
    font-size: 0.85rem;
    margin-top: 4px;
    letter-spacing: 1px;
  }

  /* ── NAV ── */
  nav {
    background-color: #1e293b;
    padding: 0;
    display: flex;
    justify-content: center;
    flex-wrap: wrap;
    border-bottom: 2px solid #334155;
    position: sticky;
    top: 0;
    z-index: 100;
    box-shadow: 0 2px 10px rgba(0,0,0,0.4);
  }
  nav a {
    color: #cbd5e1;
    text-decoration: none;
    padding: 14px 22px;
    font-weight: 600;
    font-size: 0.9rem;
    letter-spacing: 0.5px;
    transition: background 0.2s, color 0.2s;
    border-bottom: 3px solid transparent;
  }
  nav a:hover, nav a.activo {
    color: #fbbf24;
    background-color: #0f172a;
    border-bottom-color: #f59e0b;
  }

  /* ── MAIN ── */
  main {
    max-width: 1000px;
    margin: 0 auto;
    padding: 32px 16px;
    flex: 1;
    width: 100%;
  }

  h2 {
    font-size: 1.4rem;
    color: #fbbf24;
    letter-spacing: 1px;
    margin-bottom: 20px;
    padding-bottom: 10px;
    border-bottom: 2px solid #334155;
    text-align: center;
    text-transform: uppercase;
  }

  /* ── TABLAS ── */
  .tabla-wrapper {
    overflow-x: auto;
    border-radius: 10px;
    box-shadow: 0 4px 20px rgba(0,0,0,0.4);
  }
  table {
    width: 100%;
    border-collapse: collapse;
    background-color: #1e293b;
    font-size: 0.92rem;
  }
  thead tr {
    background: linear-gradient(90deg, #991b1b, #b91c1c);
  }
  th {
    color: #fff;
    padding: 13px 12px;
    text-align: center;
    font-size: 0.8rem;
    letter-spacing: 0.5px;
    text-transform: uppercase;
    white-space: nowrap;
  }
  th.col-nombre { text-align: left; }
  td {
    padding: 11px 12px;
    border-bottom: 1px solid #2d3f55;
    text-align: center;
    vertical-align: middle;
  }
  td.col-nombre { text-align: left; font-weight: 600; }
  tr:hover td { background-color: #263347; }
  tr:last-child td { border-bottom: none; }

  /* Fila del TOP 3 */
  tr.pos-1 td { background-color: rgba(251,191,36,0.12); }
  tr.pos-2 td { background-color: rgba(148,163,184,0.10); }
  tr.pos-3 td { background-color: rgba(180,120,80,0.10); }
  tr.pos-1:hover td { background-color: rgba(251,191,36,0.18); }
  tr.pos-2:hover td { background-color: rgba(148,163,184,0.15); }
  tr.pos-3:hover td { background-color: rgba(180,120,80,0.15); }

  /* Celda TOTAL destacada */
  td.total { font-weight: 800; font-size: 1.05rem; color: #fbbf24; }
  td.total-cero { color: #64748b; }

  /* ── TARJETAS PARTICIPANTE ── */
  .grid-participantes {
    display: grid;
    grid-template-columns: repeat(auto-fill, minmax(200px, 1fr));
    gap: 16px;
    margin-top: 10px;
  }
  .card-participante {
    background: linear-gradient(135deg, #1e293b, #263347);
    border: 1px solid #334155;
    border-radius: 12px;
    padding: 20px 16px;
    text-align: center;
    transition: transform 0.2s, box-shadow 0.2s;
  }
  .card-participante:hover { transform: translateY(-3px); box-shadow: 0 8px 24px rgba(0,0,0,0.4); }
  .card-participante a {
    color: #fbbf24;
    text-decoration: none;
    font-size: 1.1rem;
    font-weight: 700;
  }
  .card-participante.sin-porra { opacity: 0.55; }
  .card-participante.sin-porra span { color: #94a3b8; font-size: 1.05rem; font-weight: 600; }
  .card-participante .puntos-card {
    font-size: 2rem;
    font-weight: 900;
    color: #fbbf24;
    line-height: 1;
    margin: 6px 0 2px;
  }
  .card-participante .label-pts { font-size: 0.7rem; color: #64748b; letter-spacing: 1px; text-transform: uppercase; }
  .card-participante .badge-pos {
    display: inline-block;
    font-size: 1.3rem;
    margin-bottom: 4px;
  }
  .tag-pendiente {
    display: inline-block;
    background-color: #334155;
    color: #94a3b8;
    font-size: 0.7rem;
    border-radius: 4px;
    padding: 2px 6px;
    margin-top: 6px;
    letter-spacing: 0.5px;
  }

  /* ── SECCIÓN PORRA INDIVIDUAL ── */
  .seccion-porra {
    margin-bottom: 36px;
  }
  .seccion-porra h3 {
    font-size: 0.85rem;
    text-transform: uppercase;
    letter-spacing: 1.5px;
    color: #f59e0b;
    padding: 8px 14px;
    background-color: #1a2a3a;
    border-left: 4px solid #f59e0b;
    border-radius: 0 6px 6px 0;
    margin-bottom: 12px;
  }
  .tabla-partidos th { font-size: 0.72rem; }
  .tabla-partidos td { font-size: 0.85rem; padding: 8px 10px; }
  .resultado-1  { color: #86efac; font-weight: 700; }
  .resultado-x  { color: #fbbf24; font-weight: 700; }
  .resultado-2  { color: #f87171; font-weight: 700; }
  .grupo-badge {
    display: inline-block;
    font-size: 0.7rem;
    background: #334155;
    border-radius: 4px;
    padding: 1px 5px;
    color: #94a3b8;
  }

  .tabla-fase th, .tabla-fase td { padding: 9px 12px; }
  .tabla-cuestiones .respuesta-si  { color: #86efac; font-weight: 700; }
  .tabla-cuestiones .respuesta-no  { color: #f87171; font-weight: 700; }

  .tabla-varios td:first-child { text-align: left; color: #94a3b8; font-size: 0.82rem; }
  .tabla-varios td.valor { color: #e2e8f0; font-weight: 600; text-align: left; }

  /* ── PARTIDOS (resultados.html) ── */
  .filtro-grupo {
    display: flex;
    flex-wrap: wrap;
    gap: 8px;
    margin-bottom: 20px;
    justify-content: center;
  }
  .btn-filtro {
    background: #1e293b;
    border: 1px solid #334155;
    color: #94a3b8;
    padding: 6px 14px;
    border-radius: 20px;
    cursor: pointer;
    font-size: 0.8rem;
    font-family: inherit;
    transition: all 0.2s;
  }
  .btn-filtro:hover, .btn-filtro.activo {
    background: #991b1b;
    border-color: #f59e0b;
    color: #fbbf24;
  }
  .partido-pendiente { color: #475569; font-style: italic; }
  .partido-jugado { color: #86efac; font-weight: 600; }
  .resultado-partido { font-size: 1rem; font-weight: 800; color: #fbbf24; letter-spacing: 1px; }
  .fase-header {
    background: linear-gradient(90deg, #1a2a3a, #1e293b);
    color: #f59e0b;
    font-size: 0.8rem;
    letter-spacing: 2px;
    text-transform: uppercase;
    padding: 8px 14px;
    border-left: 3px solid #f59e0b;
    margin: 24px 0 10px;
    border-radius: 0 6px 6px 0;
  }

  /* ── FOOTER ── */
  footer {
    background-color: #0f172a;
    border-top: 1px solid #1e293b;
    text-align: center;
    padding: 16px;
    color: #475569;
    font-size: 0.8rem;
  }

  /* ── RESPONSIVE ── */
  @media (max-width: 640px) {
    nav a { padding: 12px 14px; font-size: 0.82rem; }
    table { font-size: 0.82rem; }
    th, td { padding: 8px 6px; }
    main { padding: 20px 10px; }
  }
"""

# ─────────────────────────────────────────────
# PLANTILLA BASE HTML
# ─────────────────────────────────────────────

def pagina_base(titulo, pagina_activa, contenido, fecha):
    nav_links = [
        ("index.html",        "🏆 Clasificación"),
        ("resultados.html",   "⚽ Resultados"),
        ("participantes.html","👥 Participantes"),
        ("normas.html",       "📋 Normas"),
    ]
    nav_html = "\n".join(
        f'<a href="{href}" class="{"activo" if href == pagina_activa else ""}">{label}</a>'
        for href, label in nav_links
    )
    return f"""<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>{titulo} — Porra Mundial 2026</title>
  <style>{CSS}</style>
</head>
<body>
  <header>
    <h1>🏆 PORRA MUNDIAL 2026 🏆</h1>
    <p class="subtitulo">USA · CANADA · MEXICO</p>
  </header>
  <nav>
    {nav_html}
  </nav>
  <main>
    {contenido}
  </main>
  <footer>
    <p>Última actualización: {fecha}</p>
  </footer>
</body>
</html>"""


# ─────────────────────────────────────────────
# LECTURA DEL EXCEL
# ─────────────────────────────────────────────

def leer_clasificacion():
    df = pd.read_excel(ARCHIVO_EXCEL, sheet_name="Clasificación", header=None)
    # Fila 1 = cabeceras, filas 2-8 = datos
    cabeceras = ["POSICIÓN", "NOMBRE", "1ª FASE", "DIECISEISAVOS", "OCTAVOS",
                 "CUARTOS", "SEMIFINAL", "FINAL", "CUESTIONES", "VARIOS", "TOTAL"]
    datos = []
    for i in range(2, 10):
        fila = df.iloc[i]
        nombre = fila[1]
        if pd.isna(nombre):
            continue
        puntos = [fila[j] if not pd.isna(fila[j]) else None for j in range(2, 11)]
        datos.append({"nombre": str(nombre).strip(), "puntos": puntos})
    # Ordenar por TOTAL (col index 8 dentro de puntos)
    datos.sort(key=lambda x: (x["puntos"][8] if x["puntos"][8] is not None else -1), reverse=True)
    return datos, cabeceras[2:]  # devuelve datos + nombres de columnas de puntos


def leer_partidos_porra():
    """Lee la hoja Porra y devuelve partidos de la 1ª fase."""
    df = pd.read_excel(ARCHIVO_EXCEL, sheet_name="Porra", header=None)
    partidos = []
    # Filas 2-73: partidos de la 1ª fase (cols 0-8)
    for i in range(2, 74):
        fila = df.iloc[i]
        fecha_str = str(fila[0]).strip() if not pd.isna(fila[0]) else ""
        estadio   = str(fila[1]).strip() if not pd.isna(fila[1]) else ""
        grupo     = str(fila[2]).strip() if not pd.isna(fila[2]) else ""
        local     = str(fila[4]).strip() if not pd.isna(fila[4]) else ""
        visitante = str(fila[6]).strip() if not pd.isna(fila[6]) else ""
        resultado = str(fila[7]).strip() if not pd.isna(fila[7]) else ""
        puntos    = str(fila[8]).strip() if not pd.isna(fila[8]) else ""
        if not local or local == "nan":
            continue
        partidos.append({
            "fecha": fecha_str, "estadio": estadio, "grupo": grupo,
            "local": local, "visitante": visitante,
            "resultado": resultado if resultado not in ["nan", ""] else None,
            "puntos": puntos
        })
    return partidos


def leer_cuestiones_y_varios_porra():
    """Lee la sección de Cuestiones y Varios de la hoja Porra (resultados reales)."""
    df = pd.read_excel(ARCHIVO_EXCEL, sheet_name="Porra", header=None)

    # Cuestiones: col 10 = pregunta, col 14 = resultado real, filas 45-61
    cuestiones = []
    for i in range(45, 62):
        fila = df.iloc[i]
        pregunta  = str(fila[10]).strip() if not pd.isna(fila[10]) else ""
        resultado = str(fila[14]).strip() if not pd.isna(fila[14]) else ""
        if pregunta and pregunta not in ["nan", "CUESTIONES", "Sub-Total Cuestiones"]:
            cuestiones.append({
                "pregunta": pregunta,
                "resultado": resultado if resultado not in ["nan", ""] else None
            })

    # Varios: misma estructura que en hojas de participante pero en la hoja Porra
    varios = {}
    labels_varios = {
        3:  ("campeón_1",      "Campeón 1ª opción"),
        4:  ("campeón_2",      "Campeón 2ª opción"),
        6:  ("pichichi_1",     "Pichichi 1ª opción"),
        7:  ("pichichi_2",     "Pichichi 2ª opción"),
        9:  ("mejor_1",        "Mejor jugador 1ª opción"),
        10: ("mejor_2",        "Mejor jugador 2ª opción"),
        12: ("guante_1",       "Guante de oro 1ª opción"),
        13: ("guante_2",       "Guante de oro 2ª opción"),
        15: ("pichichi_esp_1", "Pichichi español 1ª opción"),
        16: ("pichichi_esp_2", "Pichichi español 2ª opción"),
        18: ("final_res_1",    "Resultado final 1ª opción"),
        19: ("final_res_2",    "Resultado final 2ª opción"),
        21: ("esp_caboverde",  "España vs Cabo Verde"),
        22: ("esp_arabia",     "España vs Arabia Saudí"),
        23: ("uru_esp",        "Uruguay vs España"),
        25: ("max_goles_partido", "Máx. goles en un partido"),
        27: ("goles_esp",         "Total goles España"),
        29: ("goles_enc_esp",     "Goles encajados España"),
        31: ("max_goleada",       "Mayor goleada"),
    }
    for row_idx, (key, label) in labels_varios.items():
        fila = df.iloc[row_idx]
        val_18 = str(fila[18]).strip() if not pd.isna(fila[18]) else ""
        valor = val_18 if val_18 and val_18 != "nan" else ""
        if key in ("final_res_1", "final_res_2"):
            v_a = str(fila[18]).strip() if not pd.isna(fila[18]) else ""
            v_b = str(fila[19]).strip() if not pd.isna(fila[19]) else ""
            if v_a not in ["nan", ""] and v_b not in ["nan", ""]:
                try:
                    valor = f"{int(float(v_a))}-{int(float(v_b))}"
                except:
                    valor = f"{v_a}-{v_b}"
            else:
                valor = None
        varios[key] = {"label": label, "valor": valor if valor and valor != "nan" else None}

    return cuestiones, varios


def leer_porra_participante(hoja):
    """Devuelve la porra completa de un participante."""
    df = pd.read_excel(ARCHIVO_EXCEL, sheet_name=hoja, header=None)

    # ── 1ª FASE (col 0-8, filas 2-73)
    fase1 = []
    for i in range(2, 74):
        fila = df.iloc[i]
        local     = str(fila[4]).strip() if not pd.isna(fila[4]) else ""
        visitante = str(fila[6]).strip() if not pd.isna(fila[6]) else ""
        pronostico= str(fila[7]).strip() if not pd.isna(fila[7]) else ""
        grupo     = str(fila[2]).strip() if not pd.isna(fila[2]) else ""
        if not local or local == "nan":
            continue
        fase1.append({"local": local, "visitante": visitante,
                      "pron": pronostico, "grupo": grupo})

    # ── DIECISEISAVOS (cols 10-14, filas 2-17)
    doceeavos = []
    for i in range(2, 18):
        fila = df.iloc[i]
        cruce = str(fila[12]).strip() if not pd.isna(fila[12]) else ""
        eq1   = str(fila[13]).strip() if not pd.isna(fila[13]) else ""
        eq2   = str(fila[14]).strip() if not pd.isna(fila[14]) else ""
        if cruce and cruce != "nan" and eq1 and eq1 != "nan":
            doceeavos.append({"cruce": cruce, "eq1": eq1, "eq2": eq2 if eq2 != "nan" else ""})

    # ── OCTAVOS (cols 10-14, filas 21-29)
    octavos = []
    for i in range(21, 30):
        fila = df.iloc[i]
        cruce = str(fila[12]).strip() if not pd.isna(fila[12]) else ""
        eq1   = str(fila[13]).strip() if not pd.isna(fila[13]) else ""
        eq2   = str(fila[14]).strip() if not pd.isna(fila[14]) else ""
        if cruce and cruce != "nan" and eq1 and eq1 != "nan":
            octavos.append({"cruce": cruce, "eq1": eq1, "eq2": eq2 if eq2 != "nan" else ""})

    # ── CUARTOS (cols 10-14, filas 31-35)
    cuartos = []
    for i in range(31, 36):
        fila = df.iloc[i]
        cruce = str(fila[12]).strip() if not pd.isna(fila[12]) else ""
        eq1   = str(fila[13]).strip() if not pd.isna(fila[13]) else ""
        eq2   = str(fila[14]).strip() if not pd.isna(fila[14]) else ""
        if cruce and cruce != "nan" and eq1 and eq1 != "nan":
            cuartos.append({"cruce": cruce, "eq1": eq1, "eq2": eq2 if eq2 != "nan" else ""})

    # ── SEMIFINALES (cols 10-14, filas 37-39)
    semis = []
    for i in range(37, 40):
        fila = df.iloc[i]
        cruce = str(fila[12]).strip() if not pd.isna(fila[12]) else ""
        eq1   = str(fila[13]).strip() if not pd.isna(fila[13]) else ""
        eq2   = str(fila[14]).strip() if not pd.isna(fila[14]) else ""
        if cruce and cruce != "nan" and eq1 and eq1 != "nan":
            semis.append({"cruce": cruce, "eq1": eq1, "eq2": eq2 if eq2 != "nan" else ""})

    # ── FINAL (cols 10-14, fila 41-42)
    final = []
    for i in range(41, 43):
        fila = df.iloc[i]
        cruce = str(fila[12]).strip() if not pd.isna(fila[12]) else ""
        eq1   = str(fila[13]).strip() if not pd.isna(fila[13]) else ""
        eq2   = str(fila[14]).strip() if not pd.isna(fila[14]) else ""
        if cruce and cruce != "nan" and eq1 and eq1 != "nan":
            final.append({"cruce": cruce, "eq1": eq1, "eq2": eq2 if eq2 != "nan" else ""})

    # ── CUESTIONES (col 10 = pregunta, col 14 = SI/NO, filas 45-61)
    cuestiones = []
    for i in range(45, 62):
        fila = df.iloc[i]
        pregunta = str(fila[10]).strip() if not pd.isna(fila[10]) else ""
        respuesta= str(fila[14]).strip() if not pd.isna(fila[14]) else ""
        if pregunta and pregunta not in ["nan", "CUESTIONES", "Sub-Total Cuestiones"]:
            cuestiones.append({"pregunta": pregunta, "respuesta": respuesta if respuesta != "nan" else "—"})

    # ── VARIOS (cols 17-20, filas 2-31)
    varios = {}
    labels_varios = {
        3:  ("campeón_1", "Campeón 1ª opción"),
        4:  ("campeón_2", "Campeón 2ª opción"),
        6:  ("pichichi_1", "Pichichi 1ª opción"),
        7:  ("pichichi_2", "Pichichi 2ª opción"),
        9:  ("mejor_1", "Mejor jugador 1ª opción"),
        10: ("mejor_2", "Mejor jugador 2ª opción"),
        12: ("guante_1", "Guante de oro 1ª opción"),
        13: ("guante_2", "Guante de oro 2ª opción"),
        15: ("pichichi_esp_1", "Pichichi español 1ª opción"),
        16: ("pichichi_esp_2", "Pichichi español 2ª opción"),
        18: ("final_res_1", "Resultado final 1ª opción"),
        19: ("final_res_2", "Resultado final 2ª opción"),
        21: ("esp_caboverde", "España vs Cabo Verde"),
        22: ("esp_arabia",    "España vs Arabia Saudí"),
        23: ("uru_esp",       "Uruguay vs España"),
        25: ("max_goles_partido", "Máx. goles en un partido"),
        27: ("goles_esp",         "Total goles España"),
        29: ("goles_enc_esp",     "Goles encajados España"),
        31: ("max_goleada",       "Mayor goleada"),
    }
    for row_idx, (key, label) in labels_varios.items():
        fila = df.iloc[row_idx]
        val_18 = str(fila[18]).strip() if not pd.isna(fila[18]) else ""
        val_17 = str(fila[17]).strip() if not pd.isna(fila[17]) else ""
        valor = val_18 if val_18 and val_18 != "nan" else val_17
        # Para resultado final (numérico)
        if key in ("final_res_1", "final_res_2"):
            v_a = str(fila[18]).strip() if not pd.isna(fila[18]) else ""
            v_b = str(fila[19]).strip() if not pd.isna(fila[19]) else ""
            if v_a not in ["nan",""] and v_b not in ["nan",""]:
                try:
                    valor = f"{int(float(v_a))}-{int(float(v_b))}"
                except:
                    valor = f"{v_a}-{v_b}"
            else:
                valor = "—"
        varios[key] = {"label": label, "valor": valor if valor and valor != "nan" else "—"}

    return fase1, doceeavos, octavos, cuartos, semis, final, cuestiones, varios


# ─────────────────────────────────────────────
# GENERADORES DE HTML
# ─────────────────────────────────────────────

def medalla(pos):
    if pos == 1: return "🥇"
    if pos == 2: return "🥈"
    if pos == 3: return "🥉"
    return str(pos)


def generar_index(datos, col_nombres, fecha):
    filas = []
    pos_real = 0
    prev_total = None
    for i, p in enumerate(datos):
        total = p["puntos"][8]
        if total != prev_total:
            pos_real = i + 1
        prev_total = total

        clase_fila = f"pos-{pos_real}" if pos_real <= 3 else ""
        nombre = p["nombre"]
        # Enlace si tiene hoja propia
        info = PARTICIPANTES_INFO.get(nombre, PARTICIPANTES_INFO.get(nombre.upper(), PARTICIPANTES_INFO.get(nombre.lower())))
        if info and info.get("archivo"):
            nombre_html = f'<a href="{info["archivo"]}" style="color:inherit;text-decoration:none;">{info["display"]}</a>'
        else:
            nombre_html = nombre

        celdas = []
        for j, pts in enumerate(p["puntos"]):
            if j == 8:  # TOTAL
                clase = "total" if pts and pts != 0 else "total total-cero"
                celdas.append(f'<td class="{clase}">{int(pts) if pts is not None else 0}</td>')
            else:
                val = int(pts) if pts is not None else 0
                celdas.append(f"<td>{val if val != 0 else '—'}</td>")
        celdas_str = "\n          ".join(celdas)

        filas.append(f"""
        <tr class="{clase_fila}">
          <td><strong>{medalla(pos_real)}</strong></td>
          <td class="col-nombre">{nombre_html}</td>
          {celdas_str}
        </tr>""")

    filas_str = "\n".join(filas)
    cabeceras_html = "\n          ".join(f'<th>{c}</th>' for c in col_nombres)

    contenido = f"""
    <h2>Clasificación General</h2>
    <div class="tabla-wrapper">
      <table>
        <thead>
          <tr>
            <th>POS</th>
            <th class="col-nombre">NOMBRE</th>
            {cabeceras_html}
          </tr>
        </thead>
        <tbody>
          {filas_str}
        </tbody>
      </table>
    </div>
    <p style="text-align:center; color:#e2e8f0; font-size:0.95rem; margin-top:18px; letter-spacing:0.3px;">
      Haz clic en un nombre para ver su porra completa
    </p>
    """
    return pagina_base("Clasificación", "index.html", contenido, fecha)


def generar_resultados(partidos, cuestiones_porra, varios_porra, fecha):
    grupos = sorted(set(p["grupo"] for p in partidos if p["grupo"]))
    botones = '<button class="btn-filtro activo" onclick="filtrar(\'todos\', this)">Todos</button>\n'
    botones += "\n".join(
        f'<button class="btn-filtro" onclick="filtrar(\'{g.replace(" ","_")}\', this)">{g}</button>'
        for g in grupos
    )

    # ── Tab 1: Partidos
    filas = []
    grupo_actual = None
    for p in partidos:
        if p["grupo"] != grupo_actual:
            grupo_actual = p["grupo"]
            filas.append(f'<tr class="grupo-sep" data-grupo="{grupo_actual.replace(" ","_")}"><td colspan="5" class="fase-header">{grupo_actual}</td></tr>')
        res_html = f'<span class="resultado-partido">{p["resultado"]}</span>' if p["resultado"] else '<span class="partido-pendiente">Por jugar</span>'
        filas.append(f"""
        <tr data-grupo="{p['grupo'].replace(' ','_')}">
          <td style="color:#64748b;font-size:0.8rem;white-space:nowrap;">{p['fecha']}</td>
          <td class="col-nombre">{p['local']}</td>
          <td>{res_html}</td>
          <td class="col-nombre">{p['visitante']}</td>
          <td style="color:#64748b;font-size:0.75rem;">{p['estadio']}</td>
        </tr>""")
    filas_str = "\n".join(filas)

    # ── Tab 2: Cuestiones
    def icono_cuest(r):
        if r == "SI":  return '<span style="color:#86efac;font-weight:700;">✔ SI</span>'
        if r == "NO":  return '<span style="color:#f87171;font-weight:700;">✘ NO</span>'
        return '<span style="color:#475569;font-style:italic;">Pendiente</span>'

    filas_c = "\n".join(
        f"<tr><td style='text-align:left;padding:10px 14px;'>{c['pregunta']}</td>"
        f"<td style='text-align:center;padding:10px 14px;'>{icono_cuest(c['resultado'])}</td></tr>"
        for c in cuestiones_porra
    )
    tabla_cuestiones = f"""
    <div class="tabla-wrapper">
      <table class="tabla-cuestiones tabla-fase">
        <thead><tr>
          <th class="col-nombre" style="width:85%;">Pregunta</th>
          <th style="width:15%;">Resultado</th>
        </tr></thead>
        <tbody>{filas_c}</tbody>
      </table>
    </div>"""

    # ── Tab 3: Varios
    def celda_varios(valor):
        if valor:
            return f'<span style="color:#fbbf24;font-weight:600;">{valor}</span>'
        return '<span style="color:#475569;font-style:italic;">Pendiente</span>'

    # Group the varios items visually — in results page only show the real outcome (no "2ª opción")
    secciones_varios = [
        ("Campeón del Mundial",         ["campeón_1"]),
        ("Pichichi",                    ["pichichi_1"]),
        ("Mejor jugador",               ["mejor_1"]),
        ("Guante de oro",               ["guante_1"]),
        ("Pichichi español",            ["pichichi_esp_1"]),
        ("Resultado de la final",       ["final_res_1"]),
        ("Resultados España 1ª F.",     ["esp_caboverde", "esp_arabia", "uru_esp"]),
        ("Estadísticas de goles",       ["max_goles_partido", "goles_esp", "goles_enc_esp", "max_goleada"]),
    ]
    bloques_varios = []
    for titulo_sec, claves in secciones_varios:
        filas_v = "\n".join(
            f"<tr><td style='text-align:left;color:#94a3b8;font-size:0.85rem;padding:9px 14px;'>"
            f"{varios_porra[k]['label']}</td>"
            f"<td style='text-align:center;padding:9px 14px;'>{celda_varios(varios_porra[k]['valor'])}</td></tr>"
            for k in claves if k in varios_porra
        )
        bloques_varios.append(f"""
        <div style="margin-bottom:20px;">
          <div class="fase-header" style="margin:0 0 8px 0;">{titulo_sec}</div>
          <div class="tabla-wrapper">
            <table class="tabla-varios tabla-fase">
              <thead><tr>
                <th class="col-nombre" style="width:70%;">Predicción</th>
                <th style="width:30%;">Resultado</th>
              </tr></thead>
              <tbody>{filas_v}</tbody>
            </table>
          </div>
        </div>""")
    tabla_varios = "\n".join(bloques_varios)

    contenido = f"""
    <h2>Resultados del Mundial</h2>

    <!-- PESTAÑAS -->
    <div style="display:flex;flex-wrap:wrap;gap:0;margin-bottom:20px;
                border-bottom:2px solid #334155;overflow-x:auto;">
      <button class="tab-btn tab-activo" onclick="cambiarTab('partidos', this)"
              style="padding:10px 16px;background:none;border:none;color:#fbbf24;
                     font-weight:700;font-size:0.85rem;cursor:pointer;font-family:inherit;
                     border-bottom:3px solid #f59e0b;margin-bottom:-2px;white-space:nowrap;">
        ⚽ 1ª Fase
      </button>
      <button class="tab-btn" onclick="cambiarTab('doce', this)"
              style="padding:10px 16px;background:none;border:none;color:#94a3b8;
                     font-weight:700;font-size:0.85rem;cursor:pointer;font-family:inherit;
                     border-bottom:3px solid transparent;margin-bottom:-2px;white-space:nowrap;">
        🔟 1/16
      </button>
      <button class="tab-btn" onclick="cambiarTab('octavos', this)"
              style="padding:10px 16px;background:none;border:none;color:#94a3b8;
                     font-weight:700;font-size:0.85rem;cursor:pointer;font-family:inherit;
                     border-bottom:3px solid transparent;margin-bottom:-2px;white-space:nowrap;">
        8️⃣ Octavos
      </button>
      <button class="tab-btn" onclick="cambiarTab('cuartos', this)"
              style="padding:10px 16px;background:none;border:none;color:#94a3b8;
                     font-weight:700;font-size:0.85rem;cursor:pointer;font-family:inherit;
                     border-bottom:3px solid transparent;margin-bottom:-2px;white-space:nowrap;">
        4️⃣ Cuartos
      </button>
      <button class="tab-btn" onclick="cambiarTab('semis', this)"
              style="padding:10px 16px;background:none;border:none;color:#94a3b8;
                     font-weight:700;font-size:0.85rem;cursor:pointer;font-family:inherit;
                     border-bottom:3px solid transparent;margin-bottom:-2px;white-space:nowrap;">
        2️⃣ Semis
      </button>
      <button class="tab-btn" onclick="cambiarTab('final', this)"
              style="padding:10px 16px;background:none;border:none;color:#94a3b8;
                     font-weight:700;font-size:0.85rem;cursor:pointer;font-family:inherit;
                     border-bottom:3px solid transparent;margin-bottom:-2px;white-space:nowrap;">
        🏆 Final
      </button>
      <button class="tab-btn" onclick="cambiarTab('cuestiones', this)"
              style="padding:10px 16px;background:none;border:none;color:#94a3b8;
                     font-weight:700;font-size:0.85rem;cursor:pointer;font-family:inherit;
                     border-bottom:3px solid transparent;margin-bottom:-2px;white-space:nowrap;">
        ❓ Cuestiones
      </button>
      <button class="tab-btn" onclick="cambiarTab('varios', this)"
              style="padding:10px 16px;background:none;border:none;color:#94a3b8;
                     font-weight:700;font-size:0.85rem;cursor:pointer;font-family:inherit;
                     border-bottom:3px solid transparent;margin-bottom:-2px;white-space:nowrap;">
        🎯 Varios
      </button>
    </div>

    <!-- TAB: 1ª FASE -->
    <div id="tab-partidos">
      <div class="filtro-grupo">{botones}</div>
      <div class="tabla-wrapper">
        <table class="tabla-partidos">
          <thead>
            <tr>
              <th>Fecha</th>
              <th class="col-nombre">Local</th>
              <th>Resultado</th>
              <th class="col-nombre">Visitante</th>
              <th>Estadio</th>
            </tr>
          </thead>
          <tbody id="cuerpo-partidos">{filas_str}</tbody>
        </table>
      </div>
    </div>

    <!-- TABs FASES ELIMINATORIAS (pendientes de datos) -->
    <div id="tab-doce" style="display:none;"><p style="color:#475569;text-align:center;padding:40px 0;font-style:italic;">Pendiente de disputarse</p></div>
    <div id="tab-octavos" style="display:none;"><p style="color:#475569;text-align:center;padding:40px 0;font-style:italic;">Pendiente de disputarse</p></div>
    <div id="tab-cuartos" style="display:none;"><p style="color:#475569;text-align:center;padding:40px 0;font-style:italic;">Pendiente de disputarse</p></div>
    <div id="tab-semis" style="display:none;"><p style="color:#475569;text-align:center;padding:40px 0;font-style:italic;">Pendiente de disputarse</p></div>
    <div id="tab-final" style="display:none;"><p style="color:#475569;text-align:center;padding:40px 0;font-style:italic;">Pendiente de disputarse</p></div>

    <!-- TAB: CUESTIONES -->
    <div id="tab-cuestiones" style="display:none;">
      <p style="color:#64748b;font-size:0.82rem;margin-bottom:14px;text-align:center;">
        Preguntas puntuales sobre España y el Mundial (3 pts cada acierto)
      </p>
      {tabla_cuestiones}
    </div>

    <!-- TAB: VARIOS -->
    <div id="tab-varios" style="display:none;">
      <p style="color:#64748b;font-size:0.82rem;margin-bottom:14px;text-align:center;">
        Predicciones especiales con puntuación variable
      </p>
      {tabla_varios}
    </div>

    <script>
    const ALL_TABS = ['partidos','doce','octavos','cuartos','semis','final','cuestiones','varios'];
    function filtrar(grupo, btn) {{
      document.querySelectorAll('.btn-filtro').forEach(b => b.classList.remove('activo'));
      btn.classList.add('activo');
      document.querySelectorAll('#cuerpo-partidos tr').forEach(tr => {{
        if (grupo === 'todos' || tr.dataset.grupo === grupo) tr.style.display = '';
        else tr.style.display = 'none';
      }});
    }}
    function cambiarTab(id, btn) {{
      ALL_TABS.forEach(t => {{
        document.getElementById('tab-' + t).style.display = 'none';
      }});
      document.querySelectorAll('.tab-btn').forEach(b => {{
        b.style.color = '#94a3b8';
        b.style.borderBottomColor = 'transparent';
      }});
      document.getElementById('tab-' + id).style.display = '';
      btn.style.color = '#fbbf24';
      btn.style.borderBottomColor = '#f59e0b';
    }}
    </script>
    """
    return pagina_base("Resultados", "resultados.html", contenido, fecha)


def generar_participantes(datos, fecha):
    tarjetas = []
    pos_real = 0
    prev_total = None
    for i, p in enumerate(datos):
        total = p["puntos"][8]
        if total != prev_total:
            pos_real = i + 1
        prev_total = total

        nombre = p["nombre"]
        info = PARTICIPANTES_INFO.get(nombre, PARTICIPANTES_INFO.get(nombre.upper(), PARTICIPANTES_INFO.get(nombre.lower())))
        display = info["display"] if info else nombre
        archivo = info.get("archivo") if info else None
        total_str = int(total) if total is not None else 0

        if archivo:
            tarjetas.append(f"""
            <a href="{archivo}" style="text-decoration:none;color:inherit;display:block;">
              <div class="card-participante">
                <div class="badge-pos">{medalla(pos_real)}</div>
                <!-- imagen: descomenta y ajusta cuando tengas el archivo de imagen
                <img src="img/{archivo.replace('.html','')}.jpg"
                     alt="{display}"
                     style="width:70px;height:70px;border-radius:50%;object-fit:cover;
                            border:2px solid #334155;margin:6px auto 4px;display:block;">
                -->
                <div class="puntos-card">{total_str}</div>
                <div class="label-pts">puntos</div>
                <div style="color:#fbbf24;font-size:1.1rem;font-weight:700;margin-top:6px;">{display}</div>
              </div>
            </a>""")
        else:
            tarjetas.append(f"""
            <div class="card-participante sin-porra">
              <div class="badge-pos">—</div>
              <div class="puntos-card" style="color:#475569;">—</div>
              <div class="label-pts">puntos</div>
              <span>{display}</span>
              <div><span class="tag-pendiente">Porra pendiente</span></div>
            </div>""")

    tarjetas_str = "\n".join(tarjetas)
    contenido = f"""
    <h2>Participantes</h2>
    <div class="grid-participantes">
      {tarjetas_str}
    </div>
    """
    return pagina_base("Participantes", "participantes.html", contenido, fecha)


def tabla_fases(titulo, fases):
    """Genera una tabla HTML para dieciseisavos, octavos, etc."""
    if not fases:
        return ""
    filas = "\n".join(
        f"<tr><td style='color:#64748b;font-size:0.75rem;'>{f['cruce']}</td>"
        f"<td>{f['eq1']}</td>"
        f"<td style='color:#64748b;'>vs</td>"
        f"<td>{f['eq2'] or '—'}</td></tr>"
        for f in fases
    )
    return f"""
    <div class="seccion-porra">
      <h3>{titulo}</h3>
      <div class="tabla-wrapper">
        <table class="tabla-fase">
          <thead><tr>
            <th>Cruce</th><th>Clasificado A</th><th></th><th>Clasificado B</th>
          </tr></thead>
          <tbody>{filas}</tbody>
        </table>
      </div>
    </div>"""


def generar_pagina_participante(hoja, datos_clasificacion, fecha):
    info = PARTICIPANTES_INFO.get(hoja, {"display": hoja, "archivo": f"{hoja.lower()}.html"})
    display = info["display"]

    # Puntos desde clasificación
    puntos_data = next((p for p in datos_clasificacion if p["nombre"] == hoja), None)
    pts = puntos_data["puntos"] if puntos_data else [None]*9
    def fmt_pts(v): return int(v) if v is not None else 0
    total_pts   = fmt_pts(pts[8])
    pts_fase1   = fmt_pts(pts[0])
    pts_doce    = fmt_pts(pts[1])
    pts_octavos = fmt_pts(pts[2])
    pts_cuartos = fmt_pts(pts[3])
    pts_semis   = fmt_pts(pts[4])
    pts_final   = fmt_pts(pts[5])
    pts_cuest   = fmt_pts(pts[6])
    pts_varios  = fmt_pts(pts[7])

    fase1, doceeavos, octavos, cuartos, semis, final, cuestiones, varios = leer_porra_participante(hoja)

    # ── helper: recuadro de puntos con enlace a sección
    def celda_pts(label, val, tab_id):
        color = "#fbbf24" if val > 0 else "#475569"
        return (
            f'<a href="javascript:void(0)" onclick="cambiarTabPorra(\'{tab_id}\')" '
            f'style="text-decoration:none;" title="Ver {label}">'
            f'<div style="text-align:center;padding:8px 10px;background:#0f172a;border-radius:8px;'
            f'min-width:62px;cursor:pointer;transition:background 0.15s;" '
            f'onmouseover="this.style.background=\'#1e293b\'" onmouseout="this.style.background=\'#0f172a\'">'
            f'<div style="font-size:1.3rem;font-weight:800;color:{color};line-height:1;">{val}</div>'
            f'<div style="font-size:0.62rem;color:#475569;letter-spacing:0.5px;'
            f'text-transform:uppercase;margin-top:2px;">{label}</div>'
            f'</div></a>'
        )

    # ── cabecera con puntos clicables
    cabecera = f"""
    <h2 style="margin-bottom:20px;">⚽ Porra de {display}</h2>
    <div style="background:#1e293b;border:1px solid #334155;border-radius:12px;
                padding:16px 20px;margin-bottom:24px;">
      <div style="display:flex;align-items:center;flex-wrap:wrap;gap:10px;justify-content:center;">
        <div style="text-align:center;padding:10px 16px;background:#991b1b;
                    border-radius:10px;min-width:72px;">
          <div style="font-size:2rem;font-weight:900;color:#fbbf24;line-height:1;">{total_pts}</div>
          <div style="font-size:0.65rem;color:#fca5a5;letter-spacing:1px;
                      text-transform:uppercase;margin-top:3px;">TOTAL</div>
        </div>
        {celda_pts("1ª Fase",   pts_fase1,   "tab-p-fase1")}
        {celda_pts("1/16",      pts_doce,    "tab-p-doce")}
        {celda_pts("Octavos",   pts_octavos, "tab-p-octavos")}
        {celda_pts("Cuartos",   pts_cuartos, "tab-p-cuartos")}
        {celda_pts("Semis",     pts_semis,   "tab-p-semis")}
        {celda_pts("Final",     pts_final,   "tab-p-final")}
        {celda_pts("Cuestiones",pts_cuest,   "tab-p-cuestiones")}
        {celda_pts("Varios",    pts_varios,  "tab-p-varios")}
      </div>
      <p style="text-align:center;color:#475569;font-size:0.75rem;margin-top:10px;">
        Toca un apartado para ir directamente a sus predicciones
      </p>
    </div>"""

    # ── pestañas de la porra individual
    TABS_PORRA = [
        ("tab-p-fase1",     "⚽ 1ª Fase"),
        ("tab-p-doce",      "🔟 1/16"),
        ("tab-p-octavos",   "8️⃣ Octavos"),
        ("tab-p-cuartos",   "4️⃣ Cuartos"),
        ("tab-p-semis",     "2️⃣ Semis"),
        ("tab-p-final",     "🏆 Final"),
        ("tab-p-cuestiones","❓ Cuestiones"),
        ("tab-p-varios",    "🎯 Varios"),
    ]
    botones_tabs = "\n".join(
        f'<button class="tab-btn-p {"tab-p-activo" if i==0 else ""}" '
        f'onclick="cambiarTabPorra(\'{tid}\')" id="btn-{tid}" '
        f'style="padding:9px 14px;background:none;border:none;'
        f'color:{"#fbbf24" if i==0 else "#94a3b8"};font-weight:700;font-size:0.82rem;'
        f'cursor:pointer;font-family:inherit;border-bottom:3px solid {"#f59e0b" if i==0 else "transparent"};'
        f'margin-bottom:-2px;white-space:nowrap;">{label}</button>'
        for i, (tid, label) in enumerate(TABS_PORRA)
    )
    barra_tabs = f"""
    <div style="display:flex;flex-wrap:wrap;gap:0;margin-bottom:20px;
                border-bottom:2px solid #334155;overflow-x:auto;">
      {botones_tabs}
    </div>"""

    # ── contenido de cada pestaña
    def clase_res(r):
        if r == "1": return "resultado-1"
        if r in ("X","x"): return "resultado-x"
        if r == "2": return "resultado-2"
        return ""

    filas_f1 = "\n".join(
        f"<tr><td>{p['local']}</td>"
        f"<td class='{clase_res(p['pron'])}'>{p['pron'] or '—'}</td>"
        f"<td>{p['visitante']}</td>"
        f"<td><span class='grupo-badge'>{p['grupo']}</span></td></tr>"
        for p in fase1
    )
    tab_fase1 = f"""
    <div id="tab-p-fase1">
      <div class="seccion-porra">
        <h3>1ª Fase — Pronósticos (1 = local, X = empate, 2 = visitante)</h3>
        <div class="tabla-wrapper">
          <table class="tabla-partidos tabla-fase">
            <thead><tr>
              <th class="col-nombre">Local</th><th>Pronóstico</th>
              <th class="col-nombre">Visitante</th><th>Grupo</th>
            </tr></thead>
            <tbody>{filas_f1}</tbody>
          </table>
        </div>
      </div>
    </div>"""

    def tab_fase_eliminatoria(tab_id, titulo, datos):
        if not datos:
            return f'<div id="{tab_id}" style="display:none;"><p style="color:#475569;text-align:center;padding:40px 0;">Pendiente de disputarse</p></div>'
        filas = "\n".join(
            f"<tr><td style='color:#64748b;font-size:0.75rem;'>{f['cruce']}</td>"
            f"<td>{f['eq1']}</td><td style='color:#64748b;'>vs</td>"
            f"<td>{f['eq2'] or '—'}</td></tr>"
            for f in datos
        )
        return f"""
        <div id="{tab_id}" style="display:none;">
          <div class="seccion-porra">
            <h3>{titulo}</h3>
            <div class="tabla-wrapper">
              <table class="tabla-fase">
                <thead><tr><th>Cruce</th><th>Clasificado A</th><th></th><th>Clasificado B</th></tr></thead>
                <tbody>{filas}</tbody>
              </table>
            </div>
          </div>
        </div>"""

    filas_c = "\n".join(
        f"<tr><td style='text-align:left'>{c['pregunta']}</td>"
        f"<td class='{'respuesta-si' if c['respuesta']=='SI' else 'respuesta-no'}'>{c['respuesta']}</td></tr>"
        for c in cuestiones
    )
    tab_cuestiones = f"""
    <div id="tab-p-cuestiones" style="display:none;">
      <div class="seccion-porra">
        <h3>Cuestiones</h3>
        <div class="tabla-wrapper">
          <table class="tabla-cuestiones tabla-fase">
            <thead><tr><th class="col-nombre">Pregunta</th><th>Respuesta</th></tr></thead>
            <tbody>{filas_c}</tbody>
          </table>
        </div>
      </div>
    </div>""" if cuestiones else f'<div id="tab-p-cuestiones" style="display:none;"><p style="color:#475569;text-align:center;padding:40px 0;">Sin datos</p></div>'

    # Varios: agrupar con las 2 opciones (esto es la porra del participante, no resultados reales)
    secciones_v = [
        ("Campeón del Mundial",        ["campeón_1", "campeón_2"],       "🥇"),
        ("Pichichi",                   ["pichichi_1", "pichichi_2"],     "👟"),
        ("Mejor jugador",              ["mejor_1", "mejor_2"],           "⭐"),
        ("Guante de oro",              ["guante_1", "guante_2"],         "🧤"),
        ("Pichichi español",           ["pichichi_esp_1","pichichi_esp_2"],"🇪🇸"),
        ("Resultado de la final",      ["final_res_1", "final_res_2"],   "📋"),
        ("Resultados España 1ª Fase",  ["esp_caboverde","esp_arabia","uru_esp"], "🛡️"),
        ("Estadísticas de goles",      ["max_goles_partido","goles_esp","goles_enc_esp","max_goleada"], "🔢"),
    ]
    bloques_v = []
    for titulo_v, claves_v, emoji_v in secciones_v:
        filas_v = "\n".join(
            f"<tr><td style='text-align:left;color:#94a3b8;font-size:0.85rem;'>{varios[k]['label']}</td>"
            f"<td style='color:#e2e8f0;font-weight:600;'>{varios[k]['valor']}</td></tr>"
            for k in claves_v if k in varios
        )
        bloques_v.append(f"""
        <div style="margin-bottom:18px;">
          <div class="fase-header" style="margin:0 0 8px 0;">{emoji_v} {titulo_v}</div>
          <div class="tabla-wrapper">
            <table class="tabla-varios tabla-fase">
              <thead><tr><th class="col-nombre">Predicción</th><th class="col-nombre">Valor</th></tr></thead>
              <tbody>{filas_v}</tbody>
            </table>
          </div>
        </div>""")
    tab_varios = f"""
    <div id="tab-p-varios" style="display:none;">
      <div class="seccion-porra">
        <h3>Varios</h3>
        {"".join(bloques_v)}
      </div>
    </div>"""

    ALL_TABS_P = [t for t, _ in TABS_PORRA]
    all_tabs_js = str(ALL_TABS_P).replace("'", '"')

    contenido = f"""
    {cabecera}
    {barra_tabs}
    {tab_fase1}
    {tab_fase_eliminatoria("tab-p-doce",    "Dieciseisavos de Final", doceeavos)}
    {tab_fase_eliminatoria("tab-p-octavos", "Octavos de Final",       octavos)}
    {tab_fase_eliminatoria("tab-p-cuartos", "Cuartos de Final",       cuartos)}
    {tab_fase_eliminatoria("tab-p-semis",   "Semifinales",            semis)}
    {tab_fase_eliminatoria("tab-p-final",   "Final",                  final)}
    {tab_cuestiones}
    {tab_varios}

    <p style="text-align:center;margin-top:28px;">
      <a href="participantes.html" style="color:#f59e0b;font-size:0.85rem;">← Volver a participantes</a>
    </p>

    <script>
    const ALL_TABS_P = {all_tabs_js};
    function cambiarTabPorra(id) {{
      ALL_TABS_P.forEach(t => {{
        const el = document.getElementById(t);
        if (el) el.style.display = 'none';
        const btn = document.getElementById('btn-' + t);
        if (btn) {{ btn.style.color = '#94a3b8'; btn.style.borderBottomColor = 'transparent'; }}
      }});
      const target = document.getElementById(id);
      if (target) target.style.display = '';
      const btn = document.getElementById('btn-' + id);
      if (btn) {{ btn.style.color = '#fbbf24'; btn.style.borderBottomColor = '#f59e0b'; }}
    }}
    </script>
    """
    return pagina_base(f"Porra de {display}", info["archivo"], contenido, fecha)


def leer_normas():
    """Extrae las normas del Excel y las devuelve como lista de secciones estructuradas."""
    df = pd.read_excel(ARCHIVO_EXCEL, sheet_name="Normas", header=None)

    def txt(fila, col):
        v = fila[col] if col < len(fila) else None
        return str(v).strip() if v is not None and not pd.isna(v) else ""

    # Construimos las secciones manualmente a partir de la estructura conocida del Excel.
    # Cada sección: {"titulo": str, "emoji": str, "items": [{"desc": str, "puntos": str}]}
    secciones = [
        {
            "titulo": "1ª Fase",
            "emoji": "⚽",
            "items": [
                {"desc": "Un resultado por partido (1, X ó 2).", "puntos": "2 pts por acierto"},
            ]
        },
        {
            "titulo": "Dieciseisavos de Final",
            "emoji": "🔟",
            "items": [
                {"desc": "Pronosticar las selecciones clasificadas.", "puntos": "1 pt por selección acertada"},
                {"desc": "Bonus si la aciertas en la posición correcta.", "puntos": "+4 ó +6 pts"},
            ]
        },
        {
            "titulo": "Octavos de Final",
            "emoji": "8️⃣",
            "items": [
                {"desc": "Pronosticar las selecciones clasificadas.", "puntos": "5 pts por selección acertada"},
                {"desc": "Bonus si la aciertas en la posición correcta.", "puntos": "+2 pts"},
            ]
        },
        {
            "titulo": "Cuartos de Final",
            "emoji": "4️⃣",
            "items": [
                {"desc": "Pronosticar las selecciones clasificadas.", "puntos": "6 pts por selección acertada"},
                {"desc": "Bonus si la aciertas en la posición correcta.", "puntos": "+3 pts"},
            ]
        },
        {
            "titulo": "Semifinales",
            "emoji": "2️⃣",
            "items": [
                {"desc": "Pronosticar las selecciones clasificadas.", "puntos": "7 pts por selección acertada"},
                {"desc": "Bonus si la aciertas en la posición correcta.", "puntos": "+4 pts"},
            ]
        },
        {
            "titulo": "Final",
            "emoji": "🏆",
            "items": [
                {"desc": "Pronosticar las selecciones clasificadas.", "puntos": "10 pts por selección acertada"},
                {"desc": "Bonus si la aciertas en la posición correcta.", "puntos": "+5 pts"},
            ]
        },
        {
            "titulo": "Cuestiones",
            "emoji": "❓",
            "items": [
                {"desc": "Marcar Sí o No a cada pregunta.", "puntos": "3 pts por respuesta correcta"},
            ]
        },
        {
            "titulo": "Campeón del Mundial",
            "emoji": "🥇",
            "items": [
                {"desc": "Escribir la selección ganadora. Tienes 2 opciones.", "puntos": ""},
                {"desc": "Si aciertas en la 1ª opción.", "puntos": "20 pts"},
                {"desc": "Si aciertas en la 2ª opción.", "puntos": "10 pts"},
            ]
        },
        {
            "titulo": "Pichichi",
            "emoji": "👟",
            "items": [
                {"desc": "Jugador que la FIFA nombrará pichichi. Tienes 2 opciones.", "puntos": ""},
                {"desc": "Si aciertas en la 1ª opción.", "puntos": "16 pts"},
                {"desc": "Si aciertas en la 2ª opción.", "puntos": "8 pts"},
            ]
        },
        {
            "titulo": "Mejor Jugador",
            "emoji": "⭐",
            "items": [
                {"desc": "Jugador que la FIFA nombrará mejor jugador. Tienes 2 opciones.", "puntos": ""},
                {"desc": "Si aciertas en la 1ª opción.", "puntos": "16 pts"},
                {"desc": "Si aciertas en la 2ª opción.", "puntos": "8 pts"},
            ]
        },
        {
            "titulo": "Guante de Oro",
            "emoji": "🧤",
            "items": [
                {"desc": "Portero que la FIFA nombrará mejor portero. Tienes 2 opciones.", "puntos": ""},
                {"desc": "Si aciertas en la 1ª opción.", "puntos": "14 pts"},
                {"desc": "Si aciertas en la 2ª opción.", "puntos": "7 pts"},
            ]
        },
        {
            "titulo": "Pichichi Español",
            "emoji": "🇪🇸",
            "items": [
                {"desc": "Máximo goleador/es de España (mínimo 2 goles). Tienes 2 opciones.", "puntos": ""},
                {"desc": "Si aciertas en la 1ª opción.", "puntos": "12 pts"},
                {"desc": "Si aciertas en la 2ª opción.", "puntos": "6 pts"},
            ]
        },
        {
            "titulo": "Resultado de la Final",
            "emoji": "📋",
            "items": [
                {"desc": "Resultado exacto. El orden importa (el cuadro oficial marca quién va primero). Tienes 2 opciones.", "puntos": ""},
                {"desc": "Si aciertas en la 1ª opción.", "puntos": "8 pts"},
                {"desc": "Si aciertas en la 2ª opción.", "puntos": "4 pts"},
            ]
        },
        {
            "titulo": "Resultados de España en 1ª Fase",
            "emoji": "🛡️",
            "items": [
                {"desc": "Resultado exacto de cada partido de España (el orden importa).", "puntos": "6 pts por partido"},
            ]
        },
        {
            "titulo": "Mayor nº de goles en un partido",
            "emoji": "🔢",
            "items": [
                {"desc": "Suma de goles de ambos equipos en un único partido, sin penaltis.", "puntos": "5 pts por acierto"},
            ]
        },
        {
            "titulo": "Nº de goles de España",
            "emoji": "🎯",
            "items": [
                {"desc": "Total de goles marcados por España en todo el Mundial, sin penaltis.", "puntos": "5 pts por acierto"},
            ]
        },
        {
            "titulo": "Nº de goles encajados por España",
            "emoji": "🚫",
            "items": [
                {"desc": "Total de goles encajados por España en todo el Mundial, sin penaltis.", "puntos": "5 pts por acierto"},
            ]
        },
        {
            "titulo": "Mayor goleada",
            "emoji": "💥",
            "items": [
                {"desc": "Máximo de goles que un equipo marcará en un solo partido, sin penaltis.", "puntos": "5 pts por acierto"},
            ]
        },
    ]

    premios = [
        ("🥇 1er clasificado", "60% de lo recaudado"),
        ("🥈 2º clasificado",  "25% de lo recaudado"),
        ("🥉 3er clasificado", "15% de lo recaudado"),
    ]
    nota_empate = ("En caso de empate a puntos, el importe se reparte entre los "
                   "empatados sumando los premios de los puestos que ocupan.")
    return secciones, premios, nota_empate


def generar_normas(secciones, premios, nota_empate, fecha):
    # Build category cards
    cards = []
    for s in secciones:
        items_html = "\n".join(
            f"""<div style="display:flex;justify-content:space-between;align-items:baseline;
                           padding:6px 0;border-bottom:1px solid #1e293b;gap:12px;">
              <span style="color:#cbd5e1;font-size:0.88rem;">{it['desc']}</span>
              {'<span style="color:#fbbf24;font-weight:700;font-size:0.9rem;white-space:nowrap;">' + it['puntos'] + '</span>' if it['puntos'] else ''}
            </div>"""
            for it in s["items"]
        )
        cards.append(f"""
        <div style="background:#1e293b;border:1px solid #334155;border-radius:12px;
                    padding:18px 20px;break-inside:avoid;">
          <div style="display:flex;align-items:center;gap:10px;margin-bottom:12px;">
            <span style="font-size:1.5rem;">{s['emoji']}</span>
            <h3 style="color:#fbbf24;font-size:0.95rem;font-weight:700;
                       text-transform:uppercase;letter-spacing:0.5px;">{s['titulo']}</h3>
          </div>
          {items_html}
        </div>""")

    cards_html = "\n".join(cards)

    premios_html = "\n".join(
        f"""<div style="display:flex;justify-content:space-between;padding:10px 0;
                        border-bottom:1px solid #334155;">
          <span style="font-size:1rem;">{p[0]}</span>
          <span style="color:#fbbf24;font-weight:700;">{p[1]}</span>
        </div>"""
        for p in premios
    )

    contenido = f"""
    <h2>Normas de la Porra</h2>
    <p style="text-align:center;color:#94a3b8;font-size:0.88rem;margin-bottom:28px;">
      Sistema de puntuación detallado para cada sección del torneo
    </p>

    <!-- Grid de secciones -->
    <div style="display:grid;grid-template-columns:repeat(auto-fill,minmax(280px,1fr));
                gap:16px;margin-bottom:36px;">
      {cards_html}
    </div>

    <!-- Reparto de premios -->
    <div style="background:#1e293b;border:1px solid #334155;border-radius:12px;
                padding:20px 24px;max-width:500px;margin:0 auto 16px;">
      <h3 style="color:#fbbf24;font-size:1rem;font-weight:700;text-transform:uppercase;
                 letter-spacing:1px;margin-bottom:14px;">💰 Reparto de Premios</h3>
      {premios_html}
      <p style="color:#64748b;font-size:0.78rem;margin-top:12px;line-height:1.5;">
        {nota_empate}
      </p>
    </div>
    """
    return pagina_base("Normas", "normas.html", contenido, fecha)


# ─────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────

def main():
    fecha = datetime.now().strftime("%d/%m/%Y %H:%M")
    print(f"Generando web Porra Mundial 2026 — {fecha}")

    datos_clas, col_nombres = leer_clasificacion()
    partidos = leer_partidos_porra()
    cuestiones_porra, varios_porra = leer_cuestiones_y_varios_porra()
    secciones_normas, premios_normas, nota_empate = leer_normas()

    # index.html
    with open("index.html", "w", encoding="utf-8") as f:
        f.write(generar_index(datos_clas, col_nombres, fecha))
    print("  ✓ index.html")

    # resultados.html
    with open("resultados.html", "w", encoding="utf-8") as f:
        f.write(generar_resultados(partidos, cuestiones_porra, varios_porra, fecha))
    print("  ✓ resultados.html")

    # participantes.html
    with open("participantes.html", "w", encoding="utf-8") as f:
        f.write(generar_participantes(datos_clas, fecha))
    print("  ✓ participantes.html")

    # normas.html
    with open("normas.html", "w", encoding="utf-8") as f:
        f.write(generar_normas(secciones_normas, premios_normas, nota_empate, fecha))
    print("  ✓ normas.html")

    # Páginas individuales
    for hoja in HOJAS_PARTICIPANTES:
        info = PARTICIPANTES_INFO.get(hoja)
        if not info or not info.get("archivo"):
            continue
        with open(info["archivo"], "w", encoding="utf-8") as f:
            f.write(generar_pagina_participante(hoja, datos_clas, fecha))
        print(f"  ✓ {info['archivo']}")

    print("\n¡Web generada correctamente!")
    print(f"Archivos: index.html, resultados.html, participantes.html + {len(HOJAS_PARTICIPANTES)} porras individuales")


if __name__ == "__main__":
    main()
