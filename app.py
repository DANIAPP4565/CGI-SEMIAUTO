# app_CGI_ZLOGIC_ICG_MORFOLOGIA_SEMIAUTOMATICA.py
# ============================================================
# APP CGI / Z-LOGIC - INFORME HEMODINAMICO INTEGRADO
# + ANALISIS MORFOLOGICO SEMIAUTOMATICO DE SEÑAL ICG
# Autor: Ricardo Daniel Olano
# ============================================================
# Cambios principales de esta version:
# - Mantiene carga de PDF / Excel / CSV.
# - Extrae variables hemodinamicas con validacion fisiologica.
# - Prioriza ESTUDIO BASAL / ACOSTADO / CINTA / SPOT para patron principal.
# - Usa PARADO solo para comportamiento ortostatico si existe.
# - Agrega modulo semiautomatico de digitalizacion de curva ICG promedio.
# - Detecta y permite corregir cursores QRS/Q, B, C y X.
# - Calcula PEP aproximado, LVET B-X, B-C, C-X y tiempo al pico C.
# - Agrega seccion: "Analisis morfologico semiautomatico de señal ICG".
# - Exporta PDF con grafico, tabla de cursores e interpretacion.
# - Integra calculo de Ea, Ees y acoplamiento ventriculo-arterial (Ea/Ees) por metodo ICG/Capan-Chen.
# - Agrega variables de contractilidad derivadas/importadas y contractilidad morfologica desde la curva digitalizada.
# - Digitaliza automaticamente el panel superior derecho del PDF Z-Logic: dZ/dt + ECG medio + fonocardiograma inferior.
# - Completa BMI/BSA y recalcula IC/IRV por superficie corporal.
# - Agrega barra de aprendizaje por validaciones semiautomaticas de cursores.
# ============================================================

from __future__ import annotations

import io
import os
import re
import zipfile
import tempfile
import textwrap
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import numpy as np
import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt
from PIL import Image, ImageOps


def trapecio_np(y: np.ndarray, x: np.ndarray) -> float:
    """Integracion por regla del trapecio compatible con NumPy 2.x.
    np.trapz fue removido/deprecado; np.trapezoid es la funcion actual.
    """
    try:
        return float(np.trapezoid(y, x))
    except AttributeError:
        return float(np.trapz(y, x))


# Dependencias opcionales robustas
try:
    import pdfplumber
except Exception:  # pragma: no cover
    pdfplumber = None

try:
    from pypdf import PdfReader
except Exception:  # pragma: no cover
    PdfReader = None

try:
    import fitz  # PyMuPDF
except Exception:  # pragma: no cover
    fitz = None

try:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import cm
    from reportlab.platypus import (
        SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage, PageBreak
    )
except Exception:  # pragma: no cover
    colors = None
    A4 = None


# ============================================================
# CONFIGURACION STREAMLIT
# ============================================================

st.set_page_config(
    page_title="APP CGI - Z-Logic + Morfologia ICG",
    page_icon="🫀",
    layout="wide",
)

AUTOR_APP = "Dr. Ricardo Daniel Olano - Cardiologia / Hipertension Arterial"
TITULO_APP = "APP CGI - INFORME HEMODINAMICO INTEGRADO"
TITULO_MODULO_NO_EMBARAZADA = "MODULO DE EVALUACION HEMODINAMICA NO INVASIVA POR CARDIOGRAFIA DE IMPEDANCIA"


# ============================================================
# ESTILO VISUAL
# ============================================================

def aplicar_estilos() -> None:
    st.markdown(
        """
        <style>
        :root{
            --azul:#082F49; --azul2:#0B4F7A; --celeste:#EAF6FF; --line:#D7E3EE;
            --txt:#102033; --muted:#5B6B7D; --ok:#0F766E; --warn:#B45309; --bad:#B91C1C;
        }
        .stApp{background:linear-gradient(180deg,#F2F7FB,#FFFFFF)!important;}
        .block-container{max-width:1320px;padding-top:1.1rem;padding-bottom:2.5rem;}
        h1,h2,h3{color:var(--azul)!important;font-weight:800!important;}
        p,li,label,span,div{color:var(--txt);}
        .hero{background:linear-gradient(90deg,#082F49,#0B4F7A);color:#fff;border-radius:18px;padding:20px 24px;margin-bottom:18px;box-shadow:0 12px 28px rgba(8,47,73,.18)}
        .hero h1{color:#fff!important;margin:0;font-size:1.55rem}.hero p{color:#E0F2FE!important;margin:4px 0 0 0}
        .card{background:#fff;border:1px solid var(--line);border-radius:16px;padding:18px 20px;margin-bottom:16px;box-shadow:0 4px 14px rgba(15,23,42,.06)}
        .pill{display:inline-block;padding:4px 11px;border-radius:999px;font-weight:700;font-size:.80rem;margin-right:6px;}
        .pill-ok{background:#ECFDF5;color:#065F46;border:1px solid #99F6E4}.pill-warn{background:#FFF7ED;color:#9A3412;border:1px solid #FED7AA}.pill-bad{background:#FEF2F2;color:#991B1B;border:1px solid #FECACA}.pill-info{background:#EAF6FF;color:#075985;border:1px solid #BAE6FD}
        .small-muted{color:var(--muted)!important;font-size:.88rem;}
        div[data-testid="stFileUploader"] section{background:#fff!important;border:1.5px dashed #38BDF8!important;border-radius:14px!important;}
        .stButton>button,.stDownloadButton>button{background:#0B4F7A!important;color:#fff!important;border-radius:10px!important;border:1px solid #082F49!important;font-weight:800!important;}
        .stButton>button *,.stDownloadButton>button *{color:#fff!important;}
        </style>
        """,
        unsafe_allow_html=True,
    )

aplicar_estilos()
st.markdown(f"<div class='hero'><h1>{TITULO_APP}</h1><p>{AUTOR_APP}</p></div>", unsafe_allow_html=True)


# ============================================================
# UTILIDADES GENERALES
# ============================================================

def normalizar_txt(s: Any) -> str:
    s = str(s or "").lower().strip()
    for a, b in {
        "á": "a", "é": "e", "í": "i", "ó": "o", "ú": "u", "ü": "u", "ñ": "n",
        "Á": "a", "É": "e", "Í": "i", "Ó": "o", "Ú": "u", "Ñ": "n",
    }.items():
        s = s.replace(a, b)
    s = re.sub(r"\s+", " ", s)
    return s


def limpiar_numero(x: Any) -> Optional[float]:
    if x is None:
        return None
    if isinstance(x, (int, float, np.number)):
        try:
            if pd.isna(x):
                return None
        except Exception:
            pass
        return float(x)
    s = str(x).strip()
    if not s or normalizar_txt(s) in {"nan", "none", "null", "no disponible", "sd", "-"}:
        return None
    s = s.replace(" ", "")
    s = re.sub(r"[^0-9,\.\-+]", "", s)
    if not s:
        return None
    if "," in s and "." in s:
        # 1.234,5 -> 1234.5 / 1,234.5 -> 1234.5
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    else:
        s = s.replace(",", ".")
    try:
        return float(s)
    except Exception:
        return None


def fmt(x: Any, dec: int = 2, sufijo: str = "") -> str:
    v = limpiar_numero(x)
    if v is None:
        return "No disponible"
    return f"{v:.{dec}f}{sufijo}".replace(".", ",")


# ============================================================
# FORMATO Y COHERENCIA - REFERENCIA Z-LOGIC
# ============================================================
# El archivo Excel de ejemplo se usa solo para calibrar presentacion:
# cantidad de decimales y controles de coherencia fisiologica/matematica.
# No se incorpora como dato del paciente ni como fuente clinica individual.

DECIMALES_ZLOGIC = {
    "fc": 0,
    "pas": 0,
    "pad": 0,
    "pam": 0,
    "ds": 1,
    "sv": 1,
    "ids": 1,
    "si": 1,
    "vm": 1,
    "co": 1,
    "ic": 1,
    "rvs": 0,
    "svr": 0,
    "irv": 0,
    "svri": 0,
    "ca": 2,
    "iv": 0,
    "ih": 4,
    "iac": 0,
    "cts": 2,
    "str": 2,
    "itc": 1,
    "cft": 1,
    "tfc": 1,
    "cfti": 1,
    "tfci": 1,
    "z0": 1,
    "dzdt": 2,
    "pep": 0,
    "ppe": 0,
    "lvet": 0,
    "pe": 0,
    "rr": 0,
    "bsa": 1,
    "superficie": 1,
    "bmi": 1,
    "imc": 1,
    "fe": 1,
    "ea": 2,
    "ees": 2,
    "ac": 2,
    "vac": 2,
    "end": 3,
    "map": 0,
}

RANGOS_COHERENCIA_ZLOGIC = {
    "FC": (30, 180),
    "PAS": (70, 260),
    "PAD": (35, 150),
    "PAM": (45, 180),
    "DS": (15, 160),
    "IDS": (10, 90),
    "VM": (1.5, 15),
    "CO": (1.5, 15),
    "IC": (1.0, 8.0),
    "RVS": (400, 4000),
    "IRV": (700, 6500),
    "CA": (0.2, 6.0),
    "IV": (5, 150),
    "IAC": (10, 300),
    "ITC": (0.5, 8.0),
    "CFT": (15, 90),
    "Z0": (8, 80),
    "PEP": (40, 180),
    "LVET": (180, 450),
    "FE": (20, 90),
    "Ea": (0.2, 6.0),
    "Ees": (0.2, 8.0),
    "AC": (0.2, 3.5),
}

def clave_formato_metrica(nombre: Any) -> str:
    n = normalizar_txt(nombre)
    if "frecuencia" in n or n == "fc": return "fc"
    if "presion arterial media" in n or "pam" in n or n == "map": return "map"
    if "pas" in n: return "pas"
    if "pad" in n: return "pad"
    if "descarga sistolica" in n or " ds" in f" {n}" or n == "ds": return "ds"
    if "indice de descarga" in n or "si/ids" in n or n == "ids": return "ids"
    if "volumen minuto" in n or "vm" in n or "co" == n or " co" in f" {n}": return "co"
    if "indice cardiaco" in n or n.startswith("ci") or " ci" in f" {n}": return "ic"
    if "resistencia vascular sistemica" in n or "rvs" in n or "svr" in n: return "rvs"
    if "indice de resistencia" in n or "irv" in n or "svri" in n: return "irv"
    if "complacencia" in n or n == "ca": return "ca"
    if "indice de velocidad" in n or n.startswith("iv"): return "iv"
    if "heather" in n or n.startswith("ih"): return "ih"
    if "aceleracion" in n or n.startswith("iac"): return "iac"
    if "tiempo sistolico" in n or "str" in n or "cts" in n or "pep/lvet" in n: return "cts"
    if "trabajo cardiaco" in n or "itc" in n: return "itc"
    if "fluido" in n or "cft" in n or "tfc" in n: return "cft"
    if "z0" in n: return "z0"
    if "dz/dt" in n or "dzdt" in n: return "dzdt"
    if "pep" in n or "ppe" in n: return "pep"
    if "lvet" in n or " pe" in f" {n}": return "lvet"
    if "rr" in n: return "rr"
    if "bsa" in n or "superficie" in n: return "bsa"
    if "bmi" in n or "imc" in n: return "bmi"
    if "fe capan" in n or n.startswith("fe"): return "fe"
    if n.startswith("ea ") or "ea recalculada" in n: return "ea"
    if "ees" in n: return "ees"
    if "acoplamiento" in n or "ea/ees" in n or "vac" in n or n == "ac": return "ac"
    if "end(" in n: return "end"
    return "general"


def decimales_metrica(nombre: Any) -> int:
    return DECIMALES_ZLOGIC.get(clave_formato_metrica(nombre), 2)


def fmt_metrica(nombre: Any, valor: Any, sufijo: str = "") -> str:
    v = limpiar_numero(valor)
    if v is None:
        return "No disponible"
    dec = decimales_metrica(nombre)
    return f"{v:.{dec}f}{sufijo}".replace(".", ",")


def estado_rango_simple(nombre: str, valor: Any) -> Optional[str]:
    v = limpiar_numero(valor)
    if v is None:
        return None
    key = None
    cn = clave_formato_metrica(nombre)
    mapa = {
        "fc": "FC", "pas": "PAS", "pad": "PAD", "map": "PAM", "ds": "DS", "ids": "IDS",
        "co": "CO", "ic": "IC", "rvs": "RVS", "irv": "IRV", "ca": "CA", "iv": "IV",
        "iac": "IAC", "itc": "ITC", "cft": "CFT", "z0": "Z0", "pep": "PEP",
        "lvet": "LVET", "fe": "FE", "ea": "Ea", "ees": "Ees", "ac": "AC"
    }
    key = mapa.get(cn)
    if key not in RANGOS_COHERENCIA_ZLOGIC:
        return None
    lo, hi = RANGOS_COHERENCIA_ZLOGIC[key]
    if lo <= v <= hi:
        return "coherente"
    return f"revisar: fuera de rango operativo esperado ({lo}-{hi})"


def aplicar_formato_y_coherencia(df: pd.DataFrame) -> pd.DataFrame:
    """Agrega columnas de visualizacion y coherencia sin modificar el valor numerico crudo."""
    if df is None or df.empty:
        return df
    out = df.copy()
    if "Variable" in out.columns and "Valor recalculado" in out.columns:
        out["Valor mostrado"] = out.apply(lambda r: fmt_metrica(r.get("Variable"), r.get("Valor recalculado")), axis=1)
        def _coh(r):
            base = estado_rango_simple(str(r.get("Variable", "")), r.get("Valor recalculado"))
            return base or "sin regla de coherencia especifica"
        out["Control de coherencia"] = out.apply(_coh, axis=1)
    return out


# ============================================================
# VISUALIZACION DE ESTADO POR BARRA DE NORMALIDAD
# ============================================================
# La columna Estado se usa como barra horizontal: zona normal + punto del paciente.
# Las ecuaciones quedan fuera de la tabla visible para mejorar legibilidad.

NORMALIDAD_BARRAS = {
    "fc": (58, 86, 30, 180, "lpm"),
    "pas": (90, 135, 60, 220, "mmHg"),
    "pad": (60, 85, 30, 140, "mmHg"),
    "map": (70, 100, 45, 160, "mmHg"),
    "ds": (60, 110, 20, 160, "mL/lat"),
    "ids": (35, 65, 10, 100, "mL/lat/m²"),
    "co": (4, 8, 1, 14, "L/min"),
    "ic": (2.5, 4.4, 0.8, 7.0, "L/min/m²"),
    "rvs": (800, 1400, 300, 4000, "dyn.s.cm⁻5"),
    "irv": (1300, 2500, 700, 6500, "dyn.s.cm⁻5.m²"),
    "ca": (1.3, 2.8, 0.2, 5.0, "mL/mmHg"),
    "iv": (35, 65, 0, 120, "/1000/s"),
    "ih": (0.05, 0.15, 0, 0.35, "rel"),
    "iac": (70, 150, 0, 260, "/100/s²"),
    "cts": (0.30, 0.50, 0.05, 0.90, "ratio"),
    "itc": (3.0, 5.5, 0.5, 8.0, "kg.m/m²"),
    "cft": (41, 56, 15, 90, "1/kOhm"),
    "z0": (15, 30, 5, 60, "Ohm"),
    "pep": (60, 140, 30, 220, "ms"),
    "lvet": (200, 400, 120, 550, "ms"),
    "bmi": (18.5, 25, 10, 50, "kg/m²"),
    "fe": (0.50, 0.75, 0.20, 0.90, "fracción"),
    "ea": (0.8, 2.5, 0.1, 6.0, "mmHg/mL"),
    "ees": (1.0, 3.5, 0.1, 8.0, "mmHg/mL"),
    "ac": (0.0, 1.0, 0.0, 2.5, "Ea/Ees"),
}


def _html_escape(x: Any) -> str:
    import html
    return html.escape(str(x if x is not None else ""))


def _estado_clinico_barra(variable: Any, valor: Any) -> str:
    """Estado textual corto para acompañar la barra."""
    v = limpiar_numero(valor)
    if v is None:
        return "Sin dato"
    key = clave_formato_metrica(variable)
    spec = NORMALIDAD_BARRAS.get(key)
    if not spec:
        return "Sin límites definidos"
    lo, hi, _, _, _ = spec
    n = normalizar_txt(variable)
    if key == "ac":
        if v < 1.0:
            return "Óptimo"
        if v <= 1.3:
            return "Subóptimo"
        return "Desacoplamiento"
    if key == "fe":
        return "Conservada" if v >= 0.50 else "Disminuida"
    if key == "cft":
        if v < lo:
            return "Hipovolemia relativa"
        if v > hi:
            return "Hipervolemia relativa"
        return "Normovolemia"
    if v < lo:
        return "Bajo"
    if v > hi:
        return "Alto"
    return "Normal"


def barra_estado_html(variable: Any, valor: Any) -> str:
    """Genera barra horizontal con límites de normalidad y punto de medición."""
    v = limpiar_numero(valor)
    if v is None:
        return "<span style='color:#64748B'>Sin dato numérico</span>"
    key = clave_formato_metrica(variable)
    spec = NORMALIDAD_BARRAS.get(key)
    if not spec:
        return "<span style='color:#64748B'>Sin límites de normalidad definidos</span>"
    lo, hi, axis_min, axis_max, unit = spec
    if axis_max <= axis_min:
        return "<span style='color:#64748B'>Rango no disponible</span>"
    pos = max(0, min(100, (v - axis_min) / (axis_max - axis_min) * 100))
    lo_pct = max(0, min(100, (lo - axis_min) / (axis_max - axis_min) * 100))
    hi_pct = max(0, min(100, (hi - axis_min) / (axis_max - axis_min) * 100))
    estado = _estado_clinico_barra(variable, v)
    dec = decimales_metrica(variable)
    valor_txt = f"{v:.{dec}f}".replace('.', ',')
    lo_txt = f"{lo:.{dec}f}".replace('.', ',')
    hi_txt = f"{hi:.{dec}f}".replace('.', ',')
    color_estado = "#0F766E" if estado in {"Normal", "Óptimo", "Conservada", "Normovolemia"} else "#B45309" if estado in {"Subóptimo", "Bajo", "Alto", "Hipovolemia relativa", "Hipervolemia relativa"} else "#B91C1C"
    return f"""
    <div style='min-width:260px;max-width:360px;'>
      <div style='position:relative;height:14px;border-radius:999px;border:1px solid #CBD5E1;
                  background:linear-gradient(to right,#FEE2E2 0%,#FEE2E2 {lo_pct:.1f}%,#DCFCE7 {lo_pct:.1f}%,#DCFCE7 {hi_pct:.1f}%,#FEE2E2 {hi_pct:.1f}%,#FEE2E2 100%);'>
        <div title='Valor {valor_txt} {unit}' style='position:absolute;left:calc({pos:.1f}% - 4px);top:-4px;width:8px;height:22px;background:#0B4F7A;border-radius:3px;border:1px solid white;box-shadow:0 0 0 1px #0B4F7A;'></div>
      </div>
      <div style='display:flex;justify-content:space-between;font-size:10px;color:#64748B;line-height:1.15;margin-top:2px;'>
        <span>{axis_min:g}</span><span>Normal {lo_txt}-{hi_txt}</span><span>{axis_max:g}</span>
      </div>
      <div style='font-size:11px;line-height:1.15;margin-top:2px;color:{color_estado};font-weight:700;'>
        {estado}: {valor_txt} {unit}
      </div>
    </div>
    """


def preparar_metricas_para_visualizar(df: pd.DataFrame) -> pd.DataFrame:
    """Retira Ecuacion y convierte Estado en barra de normalidad para la tabla visible."""
    if df is None or df.empty:
        return pd.DataFrame()
    out = aplicar_formato_y_coherencia(df.copy())
    if "Ecuacion" in out.columns:
        out = out.drop(columns=["Ecuacion"])
    if "Variable" in out.columns and "Valor recalculado" in out.columns:
        out["Estado"] = out.apply(lambda r: barra_estado_html(r.get("Variable"), r.get("Valor recalculado")), axis=1)
    # Estado queda como quinta columna, según pedido: Variable | Valor recalculado | Unidad | Valor mostrado | Estado | Control...
    orden = [c for c in ["Variable", "Valor recalculado", "Unidad", "Valor mostrado", "Estado", "Control de coherencia"] if c in out.columns]
    resto = [c for c in out.columns if c not in orden]
    return out[orden + resto]


def mostrar_tabla_metricas_con_barras(df: pd.DataFrame, height: int = 520) -> None:
    """Muestra métricas en HTML para permitir barras dentro de la columna Estado."""
    if df is None or df.empty:
        st.info("No hay métricas para mostrar.")
        return
    vis = preparar_metricas_para_visualizar(df)
    # Escapar todas las columnas salvo Estado, que contiene HTML controlado.
    safe = vis.copy()
    for c in safe.columns:
        if c != "Estado":
            safe[c] = safe[c].map(_html_escape)
    html_table = safe.to_html(index=False, escape=False, border=0)
    st.markdown(
        f"""
        <div style='max-height:{height}px;overflow:auto;border:1px solid #D7E3EE;border-radius:12px;background:white;'>
        <style>
        table.metricas-barra{{border-collapse:collapse;width:100%;font-size:13px;}}
        table.metricas-barra th{{position:sticky;top:0;background:#0B4F7A;color:white;text-align:left;padding:8px;border:1px solid #D7E3EE;z-index:1;}}
        table.metricas-barra td{{padding:7px 8px;border:1px solid #E5E7EB;vertical-align:middle;}}
        table.metricas-barra tr:nth-child(even){{background:#F8FAFC;}}
        </style>
        {html_table.replace('<table border="0" class="dataframe">', '<table class="metricas-barra">')}
        </div>
        """,
        unsafe_allow_html=True,
    )


def metricas_excel_sin_ecuacion(df: Optional[pd.DataFrame]) -> pd.DataFrame:
    """Exporta métricas sin columna Ecuacion, manteniendo estado textual y coherencia."""
    if df is None or df.empty:
        return pd.DataFrame()
    out = aplicar_formato_y_coherencia(df.copy())
    if "Ecuacion" in out.columns:
        out = out.drop(columns=["Ecuacion"])
    if "Variable" in out.columns and "Valor recalculado" in out.columns:
        out["Estado"] = out.apply(lambda r: _estado_clinico_barra(r.get("Variable"), r.get("Valor recalculado")), axis=1)
    orden = [c for c in ["Variable", "Valor recalculado", "Unidad", "Valor mostrado", "Estado", "Control de coherencia"] if c in out.columns]
    resto = [c for c in out.columns if c not in orden]
    return out[orden + resto]


def control_consistencia_basica(datos: Dict[str, Any], df_der: Optional[pd.DataFrame]) -> List[str]:
    """Controles cruzados con tolerancias amplias de equipo Z-Logic."""
    avisos: List[str] = []
    if datos is None:
        datos = {}
    peso = limpiar_numero(datos.get("peso"))
    talla = limpiar_numero(datos.get("talla"))
    bsa = limpiar_numero(datos.get("superficie_corporal")) or calcular_bsa_mosteller(talla, peso)
    fc = limpiar_numero(datos.get("fc"))
    ds = limpiar_numero(datos.get("ds"))
    co = limpiar_numero(datos.get("gc"))
    ic_imp = limpiar_numero(datos.get("ic"))
    rvs = limpiar_numero(datos.get("rvs"))
    irv_imp = limpiar_numero(datos.get("irv"))
    ids_imp = limpiar_numero(datos.get("ids"))
    if peso and talla:
        bmi = peso / ((talla/100.0)**2)
        bmi_imp = limpiar_numero(datos.get("bmi"))
        if bmi_imp is not None and abs(bmi - bmi_imp) > 0.6:
            avisos.append(f"BMI informado difiere del recalculado: {fmt_metrica('BMI', bmi_imp)} vs {fmt_metrica('BMI', bmi)}")
    if ds is not None and fc is not None:
        co_calc = ds * fc / 1000.0
        if co is not None and abs(co_calc - co) > 0.35:
            avisos.append(f"CO/VM informado difiere de DS×FC/1000: {fmt_metrica('CO', co)} vs {fmt_metrica('CO', co_calc)}")
    if co is not None and bsa is not None:
        ic_calc = co / bsa
        if ic_imp is not None and abs(ic_calc - ic_imp) > 0.25:
            avisos.append(f"IC informado difiere de CO/BSA: {fmt_metrica('IC', ic_imp)} vs {fmt_metrica('IC', ic_calc)}")
    if rvs is not None and bsa is not None:
        irv_calc = rvs * bsa
        if irv_imp is not None and abs(irv_calc - irv_imp) > 250:
            avisos.append(f"IRV informado difiere de RVS×BSA: {fmt_metrica('IRV', irv_imp)} vs {fmt_metrica('IRV', irv_calc)}")
    if ds is not None and bsa is not None:
        ids_calc = ds / bsa
        if ids_imp is not None and abs(ids_calc - ids_imp) > 2.0:
            avisos.append(f"IDS informado difiere de DS/BSA: {fmt_metrica('IDS', ids_imp)} vs {fmt_metrica('IDS', ids_calc)}")
    return avisos


def numeros_en_texto(texto: Any) -> List[float]:
    s = str(texto or "")
    vals = []
    for m in re.findall(r"[-+]?\d+(?:[\.,]\d+)?", s):
        v = limpiar_numero(m)
        if v is not None:
            vals.append(v)
    return vals


def es_valor_util(x: Any) -> bool:
    if x is None:
        return False
    try:
        if pd.isna(x):
            return False
    except Exception:
        pass
    s = str(x).strip()
    return bool(s and normalizar_txt(s) not in {"nan", "none", "null", "sd", "no disponible", "-"})


def limpiar_nombre_archivo(txt: Any) -> str:
    s = str(txt or "").strip()
    s = re.sub(r"[\\/:*?\"<>|]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip(" .-_")
    return s or "SIN_DATO"


def safe_pdf_text(txt: Any) -> str:
    s = str(txt or "")
    reempl = {
        "—": "-", "–": "-", "“": '"', "”": '"', "‘": "'", "’": "'",
        "≥": ">=", "≤": "<=", "·": ".", "²": "2", "⁻": "-",
    }
    for a, b in reempl.items():
        s = s.replace(a, b)
    return s


# ============================================================
# MAPEO Y RANGOS CGI / Z-LOGIC
# ============================================================

VARIABLES_CGI: Dict[str, List[str]] = {
    "paciente": ["paciente", "apellido y nombre", "nombre y apellido", "patient", "nombre"],
    "dni": ["dni", "documento", "doc", "id"],
    "edad": ["edad", "age", "anos", "años"],
    "sexo": ["sexo", "sex"],
    "peso": ["peso", "weight", "kg"],
    "talla": ["talla", "altura", "height", "cm"],
    "imc": ["imc", "bmi", "indice de masa corporal"],
    "superficie_corporal": ["superficie corporal", "bsa", "sc"],
    "fecha_estudio": ["fecha estudio", "fecha del estudio", "fecha informe", "date"],
    "diagnostico": ["diagnostico", "diagnóstico", "diagnosis"],
    "medicacion": ["medicacion", "medicación", "tratamiento", "farmacos", "fármacos"],
    "posicion": ["posicion", "posición", "situacion", "situación", "condicion", "condición"],
    "pas": ["pas", "sistolica", "sistólica", "systolic"],
    "pad": ["pad", "diastolica", "diastólica", "diastolic"],
    "fc": ["frecuencia cardiaca", "frecuencia cardíaca", "heart rate", "fc", "lpm"],
    "ic": ["indice cardiaco", "índice cardíaco", "cardiac index", "ic", "ci"],
    "gc": ["gasto cardiaco", "gasto cardíaco", "cardiac output", "co", "gc", "vm"],
    "irv": ["indice de resistencia vascular", "índice de resistencia vascular", "irv", "svri", "tpvri"],
    "rvs": ["resistencia vascular sistemica", "resistencia vascular sistémica", "rvs", "svr", "tpvr"],
    "ca": ["complacencia arterial", "arterial compliance", "ca"],
    "cft": ["contenido de fluidos toracicos", "contenido de fluidos torácicos", "cft", "tfc", "thoracic fluid"],
    "cftnr": ["cftnr", "cft nr", "cft normalizado", "tfc index", "thoracic fluid content index"],
    "iv": ["indice de velocidad", "índice de velocidad", "velocity index", "iv"],
    "iac": ["indice de aceleracion", "índice de aceleración", "acceleration index", "iac", "aci"],
    "ih": ["indice de heather", "índice de heather", "heather", "ih", "hi"],
    "cts": ["cts", "pep/lvet", "pep / lvet", "systolic time ratio", "str"],
    "ea": ["elastancia arterial", "arterial elastance", "ea"],
    "ees": ["elastancia de fin de sistole", "elastancia ventricular", "ees"],
    "ava": ["acoplamiento ventriculo arterial", "acoplamiento ventriculo-arterial", "ea/ees", "ava", "ac capan"],
    "ds": ["descarga sistolica", "volumen sistolico", "stroke volume", "sv", "ds"],
    "ids": ["indice de descarga sistolica", "stroke index", "si", "ids"],
    "pep": ["pep", "periodo preeyectivo", "periodo pre-eyectivo", "pre ejection period", "preejection period", "q-b", "qrs-b"],
    "lvet": ["lvet", "tevi", "tiempo de eyeccion", "tiempo de eyección", "b-x"],
    "z0": ["z0", "impedancia basal"],
    "rr": ["rr", "intervalo rr"],
    "dzdt_max": ["dz/dt max", "dzdt max", "dzdtmax", "dz/dt", "dz dt"],
}

RANGOS: Dict[str, Tuple[float, float]] = {
    "edad": (0, 120), "peso": (25, 250), "talla": (90, 230), "imc": (10, 80),
    "pas": (60, 260), "pad": (30, 160), "fc": (35, 190),
    "ic": (0.8, 8.0), "gc": (1.0, 25.0), "irv": (700, 7000), "rvs": (300, 4500),
    "ca": (0.1, 10.0), "cft": (5, 120), "cftnr": (1, 220),
    "iv": (0, 200), "iac": (0, 80), "ih": (0, 80), "cts": (0.05, 1.5),
    "ea": (0.1, 10), "ees": (0.1, 25), "ava": (0.1, 5),
    "ds": (10, 250), "ids": (5, 150), "pep": (40, 180), "lvet": (150, 500), "z0": (5, 80),
    "rr": (300, 2000), "dzdt_max": (0.01, 20),
}

FUENTES_IC_PROHIBIDAS = re.compile(
    r"\b(dz\s*/?\s*dt|dzdt|itc|trabajo|iv|iac|aci|ih|heather|cft|tfc|irv|rvs|svr|ca|complacencia|ds|ids|stroke|z0|fc)\b",
    re.IGNORECASE,
)


def rango_plausible(clave: str, valor: Any) -> bool:
    v = limpiar_numero(valor)
    if v is None:
        return False
    if clave == "cts" and v > 2:
        v = v / 100.0
    if clave not in RANGOS:
        return True
    bajo, alto = RANGOS[clave]
    return bajo <= v <= alto


def contiene_sinonimo_seguro(nombre: Any, sinonimo: str) -> bool:
    n = re.sub(r"[^a-z0-9/]+", " ", normalizar_txt(nombre)).strip()
    s = re.sub(r"[^a-z0-9/]+", " ", normalizar_txt(sinonimo)).strip()
    if not s:
        return False
    if len(s) <= 3 or "/" in s:
        return bool(re.search(rf"(?<![a-z0-9]){re.escape(s)}(?![a-z0-9])", n))
    return s in n


def clave_por_texto(texto: Any) -> Optional[str]:
    t = normalizar_txt(texto)
    # Orden especial: evitar que IC se detecte dentro de otra variable
    for clave in ["cftnr", "irv", "rvs", "iac", "ih", "cts", "ava", "ic", "gc", "ca", "cft", "iv", "ea", "ees", "ds", "ids", "pas", "pad", "fc", "pep", "lvet", "z0", "rr", "dzdt_max", "paciente", "dni", "edad", "sexo", "peso", "talla", "imc", "superficie_corporal", "fecha_estudio", "diagnostico", "posicion"]:
        for s in VARIABLES_CGI.get(clave, []):
            if contiene_sinonimo_seguro(t, s):
                if clave == "ic" and FUENTES_IC_PROHIBIDAS.search(t):
                    continue
                return clave
    return None


def extraer_valor_de_linea(linea: str, clave: str) -> Any:
    if clave in {"paciente", "sexo", "diagnostico", "posicion", "fecha_estudio", "dni"}:
        if ":" in linea:
            return linea.split(":", 1)[1].strip()
        # quitar etiquetas conocidas
        out = linea
        for s in VARIABLES_CGI.get(clave, []):
            out = re.sub(re.escape(s), "", out, flags=re.IGNORECASE).strip(" -–—:\t")
        return out.strip() or None
    nums = numeros_en_texto(linea)
    if clave == "cts":
        vals = [v / 100 if v > 2 else v for v in nums]
    else:
        vals = nums
    plausibles = [v for v in vals if rango_plausible(clave, v)]
    return plausibles[-1] if plausibles else (vals[-1] if vals else None)



def aplicar_extraccion_zlogic_estructurada(datos: Dict[str, Any], lineas: List[str]) -> Dict[str, Any]:
    """
    Corrección robusta específica para informes Z-Logic de una hoja.

    Motivo: el texto del PDF se extrae en varias columnas y el parser genérico
    puede mezclar Edad/Sexo/Peso/Cuello o tomar PAM como PAS. Esta rutina usa
    patrones anclados a las etiquetas clínicas reales y pisa valores erróneos.
    """
    txt = "\n".join(lineas)
    flat = re.sub(r"\s+", " ", txt)

    def set_num(clave: str, valor: Any, force: bool = True):
        v = limpiar_numero(valor)
        if v is None:
            return
        if clave in RANGOS and not rango_plausible(clave, v):
            return
        if force or not es_valor_util(datos.get(clave)):
            datos[clave] = v

    def set_txt(clave: str, valor: Any, force: bool = True):
        if valor is None:
            return
        v = str(valor).strip(" :-\t")
        if not v:
            return
        if force or not es_valor_util(datos.get(clave)):
            datos[clave] = v

    # Paciente: detener antes de H.C. o parámetros si el OCR/lector unió columnas.
    m = re.search(r"(?:^|\n)\s*Paciente\s+(.+?)(?=\n\s*(?:H\.C\.|PAR[ÁA]METRO|Observaciones|ECG|Fecha)\b|$)", txt, flags=re.I)
    if m:
        pac = re.sub(r"\s+", " ", m.group(1)).strip()
        if paciente_valido(pac):
            set_txt("paciente", pac)

    # Datos demográficos y antropométricos.
    for pat, clave in [
        (r"\bEdad\s+([0-9]{1,3})\b", "edad"),
        (r"\bAltura\s+([0-9]{2,3}(?:[\.,][0-9]+)?)\s*cm", "talla"),
        (r"\bPeso\s+([0-9]{2,3}(?:[\.,][0-9]+)?)\s*kg", "peso"),
        (r"\bBSA\s+([0-9]+(?:[\.,][0-9]+)?)\s*m", "superficie_corporal"),
        (r"\bBMI\s+([0-9]+(?:[\.,][0-9]+)?)\b", "imc"),
    ]:
        m = re.search(pat, flat, flags=re.I)
        if m:
            set_num(clave, m.group(1))

    m = re.search(r"\bSexo\s+([MF])\b", flat, flags=re.I)
    if m:
        set_txt("sexo", m.group(1).upper())
    m = re.search(r"\bFecha\s+([0-9]{1,2}/[0-9]{1,2}/[0-9]{2,4})\b", flat, flags=re.I)
    if m:
        set_txt("fecha_estudio", m.group(1))
    m = re.search(r"\bSituaci[oó]n\s+([A-Za-zÁÉÍÓÚÜÑáéíóúüñ ]+?)(?=\s+Fecha\s+de\s+nac|\s+Peso\b|\n|$)", flat, flags=re.I)
    if m:
        set_txt("posicion", m.group(1).strip().upper())
    m = re.search(r"\bDiagn[oó]stico\s+(.+?)(?=\s+(?:FC\s+Frecuencia|Medicación|Medicaci[oó]n|PA\s+Sist)|\n|$)", flat, flags=re.I)
    if m:
        set_txt("diagnostico", re.sub(r"\s+", " ", m.group(1)).strip())

    # Medicación/tratamiento: se usa solo para redactar recomendaciones de optimización, no como métrica hemodinámica.
    m = re.search(r"\bMedicaci[oó]n\s+(.+?)(?=\s+(?:Dist\.?\s*e/\s*electrodos|Observaciones|PAR[ÁA]METRO|ECG|Fecha|Sistema)|$)", flat, flags=re.I)
    if m:
        med = re.sub(r"\s+", " ", m.group(1)).strip(" :-")
        set_txt("medicacion", med[:260])

    # Variables principales del informe. Usar etiquetas ancladas para evitar confundir
    # rangos de normalidad o valores de la columna derecha.
    patrones = [
        ("fc", r"\bFC\s+Frecuencia\s+Card[íi]aca\s+([0-9]+(?:[\.,][0-9]+)?)\s+pulsos"),
        ("ds", r"\bDS\s+Descarga\s+Sist[oó]lica\s+([0-9]+(?:[\.,][0-9]+)?)\s+ml/pulso"),
        ("ids", r"\bIDS\s+Indice\s+de\s+Descarga\s+Sist[oó]lica\s+([0-9]+(?:[\.,][0-9]+)?)\s+ml/pulso/m"),
        ("gc", r"\bVM\s+Volumen\s+Minuto\s+([0-9]+(?:[\.,][0-9]+)?)\s+L/min"),
        ("ic", r"\bIC\s+Indice\s+Card[íi]aco\s+([0-9]+(?:[\.,][0-9]+)?)\s+L/min/m"),
        ("rvs", r"\bRVS\s+Resistencia\s+Vascular\s+Sist[eé]mica\s+([0-9]+(?:[\.,][0-9]+)?)\s+dyn"),
        ("irv", r"\bIRV\s+Indice\s+de\s+Resistencia\s+Vascular\s+([0-9]+(?:[\.,][0-9]+)?)\s+dyn"),
        ("ca", r"\bCA\s+Complacencia\s+Arterial\s+([0-9]+(?:[\.,][0-9]+)?)\s+ml/mmHg"),
        ("iv", r"\bIV\s+Indice\s+de\s+Velocidad\s+([0-9]+(?:[\.,][0-9]+)?)\s+/1000"),
        ("iac", r"\bIAC\s+Indice\s+de\s+Aceleraci[oó]n\s+Card[íi]aca\s+([0-9]+(?:[\.,][0-9]+)?)\s+/100"),
        ("itc", r"\bITC\s+Indice\s+de\s+Trabajo\s+Card[íi]aco\s+([0-9]+(?:[\.,][0-9]+)?)\s+Kg"),
        ("cft", r"\bCFT\s+Contenido\s+de\s+Fluidos\s+Tor[áa]cicos\s+([0-9]+(?:[\.,][0-9]+)?)\s+kohms"),
        ("z0", r"\bZ0\s+([0-9]+(?:[\.,][0-9]+)?)\b"),
        ("dzdt_max", r"dz\s*/\s*dt\s*\|?\s*max\s+([0-9]+(?:[\.,][0-9]+)?)"),
    ]
    for clave, pat in patrones:
        m = re.search(pat, flat, flags=re.I)
        if m:
            set_num(clave, m.group(1))

    m = re.search(r"\bPA\s+Sist[oó]lica/Diast[oó]lica\s*\(Media\)\s*([0-9]{2,3})\s*/\s*([0-9]{2,3})\s*\(([0-9]{2,3})\)", flat, flags=re.I)
    if m:
        set_num("pas", m.group(1))
        set_num("pad", m.group(2))
        datos["pam_importada"] = limpiar_numero(m.group(3))

    m = re.search(r"\bCTS\s+Cociente\s+de\s+Tiempo\s+Sist[oó]lico\s*\(PPE/PE\)\s*([0-9]+(?:[\.,][0-9]+)?)\s*%\s*\(([0-9]+)\s*/\s*([0-9]+)\)", flat, flags=re.I)
    if m:
        set_num("cts", limpiar_numero(m.group(1)) / 100.0)
        set_num("pep", m.group(2))
        set_num("lvet", m.group(3))

    # PE/PPE también pueden aparecer junto a IDS/VM por el solapamiento de columnas.
    for clave, pat in [("lvet", r"\bPE\s+([0-9]{2,4})\b"), ("pep", r"\bPPE\s+([0-9]{2,4})\b")]:
        m = re.search(pat, flat, flags=re.I)
        if m:
            set_num(clave, m.group(1), force=False)

    # RR puede no extraerse por la medicación superpuesta; si falta, estimar por FC.
    m = re.search(r"\bRR\s+([0-9]{3,4})\b", flat, flags=re.I)
    if m:
        set_num("rr", m.group(1))
    elif limpiar_numero(datos.get("fc")):
        rr_est = 60000.0 / float(limpiar_numero(datos.get("fc")))
        if 300 <= rr_est <= 2000:
            datos["rr"] = round(rr_est, 1)

    # Recalcular BSA/BMI si el informe no los aporta o si quedaron incoherentes.
    peso = limpiar_numero(datos.get("peso"))
    talla = limpiar_numero(datos.get("talla"))
    bsa_calc = calcular_bsa_mosteller(talla, peso) if "calcular_bsa_mosteller" in globals() else None
    if bsa_calc is not None:
        bsa_v = limpiar_numero(datos.get("superficie_corporal"))
        if bsa_v is None or not (0.8 <= bsa_v <= 3.2):
            datos["superficie_corporal"] = round(bsa_calc, 3)
    if peso is not None and talla is not None and talla > 0:
        bmi_calc = peso / ((talla/100.0) ** 2)
        if 10 <= bmi_calc <= 80:
            imc_v = limpiar_numero(datos.get("imc"))
            if imc_v is None or abs(imc_v - bmi_calc) > 1.5:
                datos["imc"] = round(bmi_calc, 1)

    return datos


# ============================================================
# LECTURA DE ARCHIVOS
# ============================================================

def leer_pdf_texto(uploaded_file) -> Tuple[List[str], bytes]:
    pdf_bytes = uploaded_file.read()
    uploaded_file.seek(0)
    lineas: List[str] = []
    if pdfplumber is not None:
        try:
            with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
                for page in pdf.pages:
                    texto = page.extract_text() or ""
                    for lin in texto.splitlines():
                        lin = lin.strip()
                        if lin:
                            lineas.append(lin)
        except Exception:
            pass
    if not lineas and PdfReader is not None:
        try:
            reader = PdfReader(io.BytesIO(pdf_bytes))
            for page in reader.pages:
                texto = page.extract_text() or ""
                for lin in texto.splitlines():
                    lin = lin.strip()
                    if lin:
                        lineas.append(lin)
        except Exception:
            pass
    return lineas, pdf_bytes


def extraer_datos_desde_lineas(lineas: List[str], nombre_archivo: str = "") -> pd.DataFrame:
    registros: List[Dict[str, Any]] = []
    datos_globales: Dict[str, Any] = {"archivo_origen": nombre_archivo}

    # 1) Extraccion etiqueta: valor
    for lin in lineas:
        clave = clave_por_texto(lin)
        if not clave:
            # Presion arterial tipo 120/80
            mpa = re.search(r"(\d{2,3})\s*/\s*(\d{2,3})", lin)
            if mpa and ("pres" in normalizar_txt(lin) or "mmhg" in normalizar_txt(lin)):
                pas, pad = float(mpa.group(1)), float(mpa.group(2))
                if rango_plausible("pas", pas): datos_globales["pas"] = pas
                if rango_plausible("pad", pad): datos_globales["pad"] = pad
            continue
        valor = extraer_valor_de_linea(lin, clave)
        if clave in RANGOS and not rango_plausible(clave, valor):
            continue
        if clave == "cts" and limpiar_numero(valor) is not None and limpiar_numero(valor) > 2:
            valor = limpiar_numero(valor) / 100.0
        # no pisar paciente con etiquetas raras
        if clave == "paciente" and not paciente_valido(valor):
            continue
        datos_globales[clave] = valor

    # 2) Detectar posicion
    texto_full = " \n".join(lineas)
    nfull = normalizar_txt(texto_full)

    # 2a) Bloque vertical caracteristico del informe Z-Logic:
    # RR / PE / PPE / dz/dt / Z0 seguido por sus valores.
    try:
        mblock = re.search(
            r"\bRR\b\s*\n\s*PE\s*\n\s*PPE\s*\n\s*dz\s*/?\s*dt\s*\n\s*Z0\s*\n\s*"
            r"([0-9]+(?:[\.,][0-9]+)?)\s*\n\s*"
            r"([0-9]+(?:[\.,][0-9]+)?)\s*\n\s*"
            r"([0-9]+(?:[\.,][0-9]+)?)\s*\n\s*"
            r"([0-9]+(?:[\.,][0-9]+)?)\s*\n\s*"
            r"([0-9]+(?:[\.,][0-9]+)?)",
            texto_full, flags=re.IGNORECASE
        )
        if mblock:
            rr_v, pe_v, ppe_v, dz_v, z0_v = [limpiar_numero(x) for x in mblock.groups()]
            if rango_plausible("rr", rr_v): datos_globales["rr"] = rr_v
            if rango_plausible("lvet", pe_v): datos_globales["lvet"] = pe_v
            if rango_plausible("pep", ppe_v): datos_globales["pep"] = ppe_v
            if rango_plausible("dzdt_max", dz_v): datos_globales["dzdt_max"] = dz_v
            if rango_plausible("z0", z0_v): datos_globales["z0"] = z0_v
    except Exception:
        pass
    if "parado" in nfull or "bipedest" in nfull or "de pie" in nfull:
        datos_globales.setdefault("posicion", "PARADO")
    if any(x in nfull for x in ["acostado", "decubito", "supino", "cinta", "spot", "estudio basal"]):
        # si no es claramente solo parado, marcar basal
        datos_globales.setdefault("posicion", "ESTUDIO BASAL / ACOSTADO / CINTA / SPOT")

    if not datos_globales.get("posicion"):
        datos_globales["posicion"] = "ESTUDIO BASAL"

    # Corrección estructurada para Z-Logic: evita Edad/Sexo/Peso/PA/IC mal ubicados por lectura multicolumna.
    datos_globales = aplicar_extraccion_zlogic_estructurada(datos_globales, lineas)

    registros.append(datos_globales)
    return pd.DataFrame(registros)


def paciente_valido(valor: Any) -> bool:
    if not es_valor_util(valor):
        return False
    s = str(valor).strip()
    n = normalizar_txt(s)
    if re.search(r"\b(hc|dni|documento|fecha|edad|obra social|metodo|cinta|spot|ecg|ekg|mmhg|ohm|z logic|cardiograf|impedancia|indice|vascular|sistol|diastol)\b", n):
        return False
    letras = re.findall(r"[A-Za-zÁÉÍÓÚÜÑáéíóúüñ]", s)
    return len(letras) >= 3


def leer_excel_csv(uploaded_file) -> pd.DataFrame:
    name = uploaded_file.name.lower()
    if name.endswith(".csv"):
        try:
            return pd.read_csv(uploaded_file)
        except Exception:
            uploaded_file.seek(0)
            return pd.read_csv(uploaded_file, sep=";")
    else:
        return pd.read_excel(uploaded_file, sheet_name=None)


def normalizar_dataframe_tabular(df: pd.DataFrame, nombre_archivo: str = "") -> pd.DataFrame:
    # Caso 1: columnas ya son variables
    rows: List[Dict[str, Any]] = []
    if df is None or df.empty:
        return pd.DataFrame()
    tmp = df.copy()
    tmp.columns = [str(c).strip() for c in tmp.columns]

    mapa_cols: Dict[str, str] = {}
    for c in tmp.columns:
        clave = clave_por_texto(c)
        if clave and clave not in mapa_cols:
            mapa_cols[clave] = c

    if len(mapa_cols) >= 2:
        for _, r in tmp.iterrows():
            out: Dict[str, Any] = {"archivo_origen": nombre_archivo}
            for clave, col in mapa_cols.items():
                val = r.get(col)
                if clave in RANGOS:
                    val2 = limpiar_numero(val)
                    if rango_plausible(clave, val2):
                        out[clave] = val2
                else:
                    if es_valor_util(val):
                        out[clave] = val
            if len(out) > 1:
                rows.append(out)
        return pd.DataFrame(rows)

    # Caso 2: tabla variable/valor
    for _, r in tmp.iterrows():
        vals = [x for x in r.tolist() if es_valor_util(x)]
        if len(vals) < 2:
            continue
        etiqueta = str(vals[0])
        valor = vals[1]
        clave = clave_por_texto(etiqueta)
        if not clave:
            continue
        if not rows:
            rows.append({"archivo_origen": nombre_archivo})
        if clave in RANGOS:
            v = limpiar_numero(valor)
            if rango_plausible(clave, v):
                rows[0][clave] = v
        else:
            rows[0][clave] = valor
    return pd.DataFrame(rows)


def leer_archivo(uploaded_file) -> pd.DataFrame:
    nombre = uploaded_file.name
    low = nombre.lower()
    if low.endswith(".pdf"):
        lineas, _ = leer_pdf_texto(uploaded_file)
        return extraer_datos_desde_lineas(lineas, nombre)
    if low.endswith((".xlsx", ".xls", ".csv")):
        obj = leer_excel_csv(uploaded_file)
        if isinstance(obj, dict):
            partes = []
            for sheet, dfx in obj.items():
                dfn = normalizar_dataframe_tabular(dfx, f"{nombre} | {sheet}")
                if not dfn.empty:
                    dfn["hoja_origen"] = sheet
                    partes.append(dfn)
            return pd.concat(partes, ignore_index=True) if partes else pd.DataFrame()
        return normalizar_dataframe_tabular(obj, nombre)
    return pd.DataFrame()


# ============================================================
# CLASIFICACION CLINICA
# ============================================================

def es_registro_basal(row: Dict[str, Any]) -> bool:
    pos = normalizar_txt(row.get("posicion") or row.get("situacion") or row.get("condicion") or "")
    origen = normalizar_txt(row.get("archivo_origen") or "")
    txt = f"{pos} {origen}"
    if any(x in txt for x in ["parado", "de pie", "bipedest"]):
        return False
    return any(x in txt for x in ["basal", "acostado", "decubito", "supino", "cinta", "spot"]) or True


def seleccionar_basal_y_parado(df: pd.DataFrame) -> Tuple[Optional[Dict[str, Any]], Optional[Dict[str, Any]]]:
    if df is None or df.empty:
        return None, None
    registros = df.to_dict("records")
    basal = None
    parado = None
    for r in registros:
        txt = normalizar_txt(" ".join([str(r.get(k, "")) for k in ["posicion", "archivo_origen", "hoja_origen"]]))
        if any(x in txt for x in ["parado", "de pie", "bipedest"]):
            parado = r
        elif basal is None:
            basal = r
    if basal is None and registros:
        basal = registros[0]
    return basal, parado


def clasificar_dinamia(r: Optional[Dict[str, Any]]) -> str:
    if not r:
        return "Datos insuficientes"
    ic = limpiar_numero(r.get("ic"))
    irv = limpiar_numero(r.get("irv")) or limpiar_numero(r.get("rvs"))
    if ic is None or irv is None:
        return "Datos insuficientes"
    # Umbrales operativos ajustables
    ic_bajo, ic_alto = 2.5, 4.0
    irv_alto = 2400 if irv > 700 else 1600
    irv_bajo = 1600 if irv > 700 else 800
    if ic < ic_bajo and irv >= irv_alto:
        return "Hipodinamia vasoconstrictiva"
    if ic > ic_alto and irv <= irv_bajo:
        return "Hiperdinamia vasodilatada"
    if ic_bajo <= ic <= ic_alto and irv_bajo < irv < irv_alto:
        return "Normodinamia"
    if ic < ic_bajo:
        return "Hipodinamia"
    if ic > ic_alto:
        return "Hiperdinamia"
    if irv >= irv_alto:
        return "Predominio vasoconstrictor"
    if irv <= irv_bajo:
        return "Predominio vasodilatador"
    return "Patron hemodinamico definido con datos parciales"


def generar_informe_texto(df: pd.DataFrame) -> str:
    basal, parado = seleccionar_basal_y_parado(df)
    if basal is None:
        return "No se encontraron datos suficientes para generar informe."
    dinamia = clasificar_dinamia(basal)
    lineas = []
    lineas.append(TITULO_MODULO_NO_EMBARAZADA)
    lineas.append("")
    lineas.append("PATRON HEMODINAMICO PRINCIPAL: se informa exclusivamente con el registro basal/acostado/cinta/spot.")
    lineas.append(f"Fenotipo circulatorio basal: {dinamia}.")
    lineas.append(f"Presion arterial basal: PAS {fmt(basal.get('pas'),0)} mmHg / PAD {fmt(basal.get('pad'),0)} mmHg. FC {fmt(basal.get('fc'),0)} lpm.")
    lineas.append(f"Indice cardiaco: {fmt(basal.get('ic'),2)} L/min/m2. Resistencia vascular: IRV/RVS {fmt(basal.get('irv') or basal.get('rvs'),0)}.")
    lineas.append(f"Complacencia arterial: {fmt(basal.get('ca'),2)}. CFT: {fmt(basal.get('cft'),2)}. CFTnr: {fmt(basal.get('cftnr'),2)}.")
    lineas.append(f"Contractilidad: IV {fmt(basal.get('iv'),2)}, IAC {fmt(basal.get('iac'),2)}, IH {fmt(basal.get('ih'),2)}, CTS {fmt(basal.get('cts'),2)}.")
    lineas.append(f"Acoplamiento ventriculo-arterial: EA {fmt(basal.get('ea'),2)}, EES {fmt(basal.get('ees'),2)}, EA/EES {fmt(basal.get('ava'),2)}.")
    if parado is not None:
        lineas.append("")
        lineas.append("COMPORTAMIENTO ORTOSTATICO: el registro parado/de pie se utiliza solo como comparacion funcional y no reemplaza el patron basal.")
        lineas.append(f"Parado/de pie: IC {fmt(parado.get('ic'),2)}, IRV/RVS {fmt(parado.get('irv') or parado.get('rvs'),0)}, FC {fmt(parado.get('fc'),0)}.")
    lineas.append("")
    lineas.append("Interpretacion: integrar los resultados con clinica, medicacion, presion arterial seriada, laboratorio, ecocardiograma y contexto vascular del paciente.")
    return "\n".join(lineas)


# ============================================================
# BASE DE CONOCIMIENTO Y CALCULO DE ECUACIONES ICG
# ============================================================

ECUACIONES_ICG: Dict[str, Dict[str, str]] = {
    "BSA_MOSTELLER": {
        "dominio": "Antropometria",
        "ecuacion": "BSA = sqrt((talla_cm * peso_kg) / 3600)",
        "descripcion": "Superficie corporal por Mosteller.",
        "tipo": "recalculable",
    },
    "MAP": {
        "dominio": "Presion arterial",
        "ecuacion": "MAP = (PAS + 2 * PAD) / 3",
        "descripcion": "Presion arterial media estimada.",
        "tipo": "recalculable",
    },
    "CO": {
        "dominio": "Gasto cardiaco",
        "ecuacion": "CO = SV * FC / 1000",
        "descripcion": "Gasto cardiaco si SV esta en mL y FC en lpm.",
        "tipo": "recalculable si hay SV/DS y FC",
    },
    "CI": {
        "dominio": "Gasto indexado",
        "ecuacion": "CI = CO / BSA",
        "descripcion": "Indice cardiaco.",
        "tipo": "recalculable si hay CO y BSA",
    },
    "SI": {
        "dominio": "Volumen sistolico indexado",
        "ecuacion": "SI = SV / BSA",
        "descripcion": "Indice sistolico.",
        "tipo": "recalculable si hay SV/DS y BSA",
    },
    "SVR": {
        "dominio": "Resistencia vascular",
        "ecuacion": "SVR = ((MAP - CVP) / CO) * 80",
        "descripcion": "Resistencia vascular sistemica. Si no hay CVP se usa CVP=0 como aproximacion.",
        "tipo": "aproximada si no hay CVP",
    },
    "SVRI_IRV": {
        "dominio": "Resistencia vascular indexada",
        "ecuacion": "SVRI/IRV = ((MAP - CVP) / CI) * 80",
        "descripcion": "Indice de resistencia vascular sistemica. Si no hay CVP se usa CVP=0 como aproximacion.",
        "tipo": "aproximada si no hay CVP",
    },
    "CFT_TFC": {
        "dominio": "Fluido toracico",
        "ecuacion": "CFT/TFC = 1000 / Z0",
        "descripcion": "Contenido de fluido toracico derivado de impedancia basal.",
        "tipo": "recalculable si hay Z0",
    },
    "CFTI_TFCI": {
        "dominio": "Fluido toracico indexado",
        "ecuacion": "CFTi/TFCi = (1000 / Z0) / BSA",
        "descripcion": "Contenido de fluido toracico indexado por superficie corporal.",
        "tipo": "recalculable si hay Z0 y BSA",
    },
    "PEP": {
        "dominio": "Tiempo sistolico",
        "ecuacion": "PEP = tB - tQRS",
        "descripcion": "Intervalo preeyectivo aproximado desde cursores morfologicos.",
        "tipo": "morfologia digitalizada",
    },
    "LVET": {
        "dominio": "Tiempo sistolico",
        "ecuacion": "LVET = tX - tB",
        "descripcion": "Tiempo de eyeccion ventricular izquierda desde B hasta X.",
        "tipo": "morfologia digitalizada",
    },
    "STR": {
        "dominio": "Tiempo sistolico",
        "ecuacion": "STR = PEP / LVET",
        "descripcion": "Relacion PEP/LVET.",
        "tipo": "morfologia digitalizada",
    },
    "EF_CAPAN": {
        "dominio": "Funcion sistolica",
        "ecuacion": "EF = 0.84 - 0.64 * (PEP / LVET)",
        "descripcion": "Fraccion de eyeccion estimada por metodo Capan a partir de tiempos sistolicos.",
        "tipo": "recalculable con PEP y LVET",
    },
    "EA": {
        "dominio": "Acoplamiento ventriculo-arterial",
        "ecuacion": "Ea = (PAS * 0.9) / SV",
        "descripcion": "Elastancia arterial efectiva estimada como presion telesistolica aproximada sobre volumen sistolico.",
        "tipo": "recalculable con PAS y SV/DS",
    },
    "END_AVG_CHEN": {
        "dominio": "Acoplamiento ventriculo-arterial",
        "ecuacion": "End(avg)=0.35695 - 7.2266*tNd + 74.249*tNd^2 - 307.39*tNd^3 + 684.54*tNd^4 - 856.92*tNd^5 + 571.95*tNd^6 - 159.1*tNd^7",
        "descripcion": "Elastancia ventricular normalizada promedio del metodo single-beat de Chen.",
        "tipo": "recalculable con tNd",
    },
    "END_EST_CHEN": {
        "dominio": "Acoplamiento ventriculo-arterial",
        "ecuacion": "End(est)=0.0275 - 0.165*EF + 0.3656*(PAD/(PAS*0.9)) + 0.515*End(avg)",
        "descripcion": "Estimacion de elastancia ventricular normalizada al inicio de eyeccion.",
        "tipo": "recalculable con EF, PAS, PAD y End(avg)",
    },
    "EES_CHEN_ICG": {
        "dominio": "Acoplamiento ventriculo-arterial",
        "ecuacion": "Ees = (PAD - End(est) * PAS * 0.9) / (End(est) * SV)",
        "descripcion": "Elastancia telesistolica ventricular por metodo single-beat de Chen usando datos de ICG.",
        "tipo": "recalculable con PAS, PAD, SV, PEP y LVET",
    },
    "VAC_EA_EES": {
        "dominio": "Acoplamiento ventriculo-arterial",
        "ecuacion": "VAC = Ea / Ees",
        "descripcion": "Acoplamiento ventriculo-arterial como relacion Ea/Ees.",
        "tipo": "recalculable si Ea y Ees son validas",
    },
    "VI": {
        "dominio": "Contractilidad",
        "ecuacion": "VI = 1000 * (dZ/dt)max / Z0",
        "descripcion": "Indice de velocidad. Requiere amplitud absoluta calibrada de dZ/dt.",
        "tipo": "requiere calibracion absoluta",
    },
    "HI": {
        "dominio": "Contractilidad",
        "ecuacion": "HI = (dZ/dt)max / (tC - tQRS)",
        "descripcion": "Indice de Heather. Requiere amplitud absoluta calibrada para equivalencia con equipo.",
        "tipo": "requiere calibracion absoluta",
    },
    "DZDT_MAX_REL": {
        "dominio": "Contractilidad morfologica",
        "ecuacion": "dZdt_max_rel = Amp(C) - Amp(B)",
        "descripcion": "Amplitud sistolica relativa de la onda dZ/dt digitalizada. No reemplaza al dZ/dt maximo absoluto del equipo.",
        "tipo": "morfologia digitalizada relativa",
    },
    "HI_REL": {
        "dominio": "Contractilidad morfologica",
        "ecuacion": "HI_rel = dZdt_max_rel / (tC - tQRS)",
        "descripcion": "Indice de Heather relativo derivado de la curva digitalizada; solo comparable dentro del mismo metodo/escala.",
        "tipo": "morfologia digitalizada relativa",
    },
    "VI_REL": {
        "dominio": "Contractilidad morfologica",
        "ecuacion": "VI_rel = 1000 * dZdt_max_rel / Z0",
        "descripcion": "Indice de velocidad relativo si existe Z0; no equivale al VI propietario sin calibracion absoluta.",
        "tipo": "morfologia digitalizada relativa",
    },
    "ACI_REL": {
        "dominio": "Contractilidad morfologica",
        "ecuacion": "ACI_rel = max(d(dZ/dt)/dt)",
        "descripcion": "Aceleracion maxima relativa de la curva dZ/dt digitalizada.",
        "tipo": "morfologia digitalizada relativa",
    },
    "PENDIENTE_ASC": {
        "dominio": "Contractilidad morfologica",
        "ecuacion": "Pendiente_ascendente = (AmpC - AmpB) / (tC - tB)",
        "descripcion": "Pendiente de ascenso sistolico B-C; marcador relativo de inotropismo/eyeccion temprana.",
        "tipo": "morfologia digitalizada relativa",
    },
    "PENDIENTE_DESC": {
        "dominio": "Contractilidad morfologica",
        "ecuacion": "Pendiente_descendente = (AmpX - AmpC) / (tX - tC)",
        "descripcion": "Pendiente de caida C-X; describe desaceleracion sistolica relativa.",
        "tipo": "morfologia digitalizada relativa",
    },
    "AREA_SISTOLICA_REL": {
        "dominio": "Contractilidad morfologica",
        "ecuacion": "Area_sistolica_rel = integral_B^X dZdt_rel dt",
        "descripcion": "Area sistolica relativa bajo la curva digitalizada entre B y X.",
        "tipo": "morfologia digitalizada relativa",
    },
    "SIMETRIA_SISTOLICA": {
        "dominio": "Contractilidad morfologica",
        "ecuacion": "Simetria = (tC - tB) / (tX - tC)",
        "descripcion": "Relacion entre ascenso y descenso sistolico de la curva dZ/dt.",
        "tipo": "morfologia digitalizada relativa",
    },
    "SV_KUBICEK": {
        "dominio": "Volumen sistolico",
        "ecuacion": "SV = rho * (L / Z0)^2 * LVET * (dZ/dt)max",
        "descripcion": "Formula clasica de Kubicek; requiere rho, distancia L, Z0, LVET y dZ/dt maximo absoluto.",
        "tipo": "requiere calibracion absoluta",
    },
    "SV_SRAMEK_BERNSTEIN": {
        "dominio": "Volumen sistolico",
        "ecuacion": "SV = delta * ((0.17 * H)^3 / (4.25 * Z0)) * (dZ/dt)max * LVET",
        "descripcion": "Formula de Sramek-Bernstein; requiere H, Z0, LVET, dZ/dt maximo absoluto y factor delta.",
        "tipo": "requiere calibracion absoluta",
    },
    "LSWI": {
        "dominio": "Trabajo ventricular",
        "ecuacion": "LSWI = (MAP - PCWP) * SI * 0.0136",
        "descripcion": "Trabajo sistolico ventricular izquierdo indexado. Si no hay PCWP, solo aproximacion.",
        "tipo": "aproximada si no hay PCWP",
    },
    "LCWI": {
        "dominio": "Trabajo ventricular",
        "ecuacion": "LCWI = (MAP - PCWP) * CI * 0.0144",
        "descripcion": "Trabajo cardiaco izquierdo indexado. Si no hay PCWP, solo aproximacion.",
        "tipo": "aproximada si no hay PCWP",
    },
}


def tabla_ecuaciones_icg() -> pd.DataFrame:
    return pd.DataFrame([
        {
            "Variable": k,
            "Dominio": v["dominio"],
            "Ecuacion": v["ecuacion"],
            "Uso en la app": v["tipo"],
            "Descripcion": v["descripcion"],
        }
        for k, v in ECUACIONES_ICG.items()
    ])


def calcular_bsa_mosteller(talla_cm: Optional[float], peso_kg: Optional[float]) -> Optional[float]:
    if talla_cm is None or peso_kg is None:
        return None
    if not (90 <= talla_cm <= 230 and 25 <= peso_kg <= 250):
        return None
    return float(np.sqrt((talla_cm * peso_kg) / 3600.0))


def calcular_ef_capan(pep_ms: Optional[float], lvet_ms: Optional[float]) -> Optional[float]:
    """Fraccion de eyeccion estimada por Capan: EF = 0.84 - 0.64*(PEP/LVET)."""
    pep = limpiar_numero(pep_ms)
    lvet = limpiar_numero(lvet_ms)
    if pep is None or lvet is None or lvet <= 0:
        return None
    if not (30 <= pep <= 220 and 120 <= lvet <= 600):
        return None
    ef = 0.84 - 0.64 * (pep / lvet)
    if not (0.10 <= ef <= 0.90):
        return None
    return float(ef)


def calcular_end_avg_chen(tnd: Optional[float]) -> Optional[float]:
    """Polinomio End(avg) del metodo single-beat de Chen."""
    t = limpiar_numero(tnd)
    if t is None or not (0.05 <= t <= 0.60):
        return None
    end_avg = (
        0.35695
        - 7.2266 * t
        + 74.249 * (t ** 2)
        - 307.39 * (t ** 3)
        + 684.54 * (t ** 4)
        - 856.92 * (t ** 5)
        + 571.95 * (t ** 6)
        - 159.1 * (t ** 7)
    )
    if not (0.01 <= end_avg <= 2.0):
        return None
    return float(end_avg)


def calcular_end_est_chen(ef: Optional[float], pas: Optional[float], pad: Optional[float], end_avg: Optional[float]) -> Optional[float]:
    """End(est) para calculo de Ees. Usa Pes aproximada = PAS*0.9."""
    ef = limpiar_numero(ef)
    pas = limpiar_numero(pas)
    pad = limpiar_numero(pad)
    end_avg = limpiar_numero(end_avg)
    if ef is None or pas is None or pad is None or end_avg is None:
        return None
    pes = pas * 0.9
    if pes <= 0 or not (0.10 <= ef <= 0.90):
        return None
    end_est = 0.0275 - 0.165 * ef + 0.3656 * (pad / pes) + 0.515 * end_avg
    if not (0.01 <= end_est <= 2.0):
        return None
    return float(end_est)


def calcular_acoplamiento_capan_chen(
    pas: Optional[float],
    pad: Optional[float],
    sv_ml: Optional[float],
    pep_ms: Optional[float],
    lvet_ms: Optional[float],
) -> Dict[str, Optional[float]]:
    """
    Calcula Ea, Ees y VAC(Ea/Ees) con datos ICG.
    Requiere PAS, PAD, SV/DS, PEP y LVET.
    PEP y LVET pueden venir de la digitalizacion semiautomatica o del informe si se extrajeron.
    """
    pas = limpiar_numero(pas)
    pad = limpiar_numero(pad)
    sv = limpiar_numero(sv_ml)
    pep = limpiar_numero(pep_ms)
    lvet = limpiar_numero(lvet_ms)
    out = {
        "ea": None, "ees": None, "vac": None, "ef_capan": None,
        "end_avg": None, "end_est": None, "tnd": None, "pes": None,
    }
    if pas is None or pad is None or sv is None or pep is None or lvet is None:
        return out
    if not (60 <= pas <= 260 and 30 <= pad <= 160 and 10 <= sv <= 250 and 30 <= pep <= 220 and 120 <= lvet <= 600):
        return out
    if lvet <= 0 or sv <= 0:
        return out
    pes = pas * 0.9
    ea = pes / sv
    ef = calcular_ef_capan(pep, lvet)
    total_sistolico = pep + lvet
    tnd = pep / total_sistolico if total_sistolico > 0 else None
    end_avg = calcular_end_avg_chen(tnd)
    end_est = calcular_end_est_chen(ef, pas, pad, end_avg)
    ees = None
    vac = None
    if end_est is not None and end_est > 0:
        ees = (pad - (end_est * pes)) / (end_est * sv)
        # En casos raros puede resultar no fisiologico por datos inconsistentes.
        if ees is not None and ees > 0:
            vac = ea / ees
        else:
            ees = None
    out.update({
        "ea": float(ea) if 0.05 <= ea <= 10 else None,
        "ees": float(ees) if ees is not None and 0.05 <= ees <= 25 else None,
        "vac": float(vac) if vac is not None and 0.05 <= vac <= 5 else None,
        "ef_capan": ef,
        "end_avg": end_avg,
        "end_est": end_est,
        "tnd": float(tnd) if tnd is not None else None,
        "pes": float(pes),
    })
    return out


def interpretar_acoplamiento_vac(vac: Optional[float]) -> str:
    v = limpiar_numero(vac)
    if v is None:
        return "No interpretable por datos insuficientes o no fisiologicos."
    if v < 0.8:
        return "Acoplamiento bajo/ventriculo relativamente dominante; interpretar con contractilidad y carga arterial."
    if 0.8 <= v <= 1.0:
        return "Acoplamiento ventriculo-arterial conservado/optimo."
    if 1.0 < v <= 1.3:
        return "Acoplamiento suboptimo: la carga arterial se aproxima o supera a la elastancia ventricular."
    return "Desacoplamiento ventriculo-arterial: predominio relativo de carga arterial o reserva ventricular insuficiente."



def calcular_contractilidad_morfologica(
    df_curva: Optional[pd.DataFrame],
    puntos: Optional[Dict[str, Any]],
    z0: Optional[float] = None,
) -> Dict[str, Optional[float]]:
    """
    Calcula variables de contractilidad relativas desde la curva dZ/dt digitalizada.

    Importante: al provenir de imagen, la amplitud esta en unidades relativas. Estas metricas
    sirven para comparacion intra-paciente o control morfologico, pero no reemplazan los indices
    absolutos/proprietarios del equipo si no existe calibracion vertical real de dZ/dt.
    """
    out = {
        "dzdt_max_rel": None,
        "hi_rel": None,
        "vi_rel": None,
        "aci_rel": None,
        "pendiente_ascendente": None,
        "pendiente_descendente": None,
        "area_sistolica_rel": None,
        "area_ascenso_rel": None,
        "area_descenso_rel": None,
        "simetria_sistolica": None,
        "ratio_area_asc_desc": None,
    }
    if df_curva is None or puntos is None or df_curva.empty:
        return out
    try:
        t = df_curva["tiempo_ms"].to_numpy(dtype=float)
        y = df_curva["amplitud_relativa"].to_numpy(dtype=float)
        i_q = int(puntos.get("idx_q"))
        i_b = int(puntos.get("idx_b"))
        i_c = int(puntos.get("idx_c"))
        i_x = int(puntos.get("idx_x"))
    except Exception:
        return out

    n = len(t)
    if n < 5 or not (0 <= i_q < i_b < i_c < i_x < n):
        return out

    tq, tb, tc, tx = t[i_q], t[i_b], t[i_c], t[i_x]
    yb, yc, yx = y[i_b], y[i_c], y[i_x]
    pep = tb - tq
    bc = tc - tb
    cx = tx - tc

    if bc <= 0 or cx <= 0 or pep <= 0:
        return out

    dzdt_max_rel = yc - yb
    pendiente_asc = dzdt_max_rel / bc
    pendiente_desc = (yx - yc) / cx

    seg_bx = slice(i_b, i_x + 1)
    seg_bc = slice(i_b, i_c + 1)
    seg_cx = slice(i_c, i_x + 1)
    area_sis = trapecio_np(y[seg_bx] - yb, t[seg_bx])
    area_asc = trapecio_np(np.maximum(y[seg_bc] - yb, 0), t[seg_bc])
    area_desc = trapecio_np(np.maximum(y[seg_cx] - yx, 0), t[seg_cx])

    dy_dt = np.gradient(y, t)
    aci_rel = float(np.nanmax(dy_dt[seg_bc])) if len(dy_dt[seg_bc]) else None

    z0v = limpiar_numero(z0)
    vi_rel = None
    if z0v is not None and z0v > 0:
        vi_rel = 1000.0 * dzdt_max_rel / z0v

    out.update({
        "dzdt_max_rel": float(dzdt_max_rel),
        "hi_rel": float(dzdt_max_rel / pep),
        "vi_rel": float(vi_rel) if vi_rel is not None else None,
        "aci_rel": aci_rel,
        "pendiente_ascendente": float(pendiente_asc),
        "pendiente_descendente": float(pendiente_desc),
        "area_sistolica_rel": area_sis,
        "area_ascenso_rel": area_asc,
        "area_descenso_rel": area_desc,
        "simetria_sistolica": float(bc / cx),
        "ratio_area_asc_desc": float(area_asc / area_desc) if area_desc and area_desc != 0 else None,
    })
    return out


def interpretar_contractilidad_morfologica(vars_ct: Dict[str, Optional[float]]) -> str:
    """Interpretacion prudente de contractilidad relativa derivada de imagen."""
    if not vars_ct or vars_ct.get("dzdt_max_rel") is None:
        return "No interpretable por falta de curva digitalizada o cursores validos."
    pendiente = vars_ct.get("pendiente_ascendente")
    sim = vars_ct.get("simetria_sistolica")
    partes = ["Contractilidad morfologica relativa derivada de la pendiente B-C, amplitud relativa C-B y tiempo Q-C."]
    if pendiente is not None:
        if pendiente > 0.8:
            partes.append("Ascenso sistolico rapido en la escala digitalizada.")
        elif pendiente < 0.15:
            partes.append("Ascenso sistolico lento en la escala digitalizada; revisar calidad de curva y cursores.")
        else:
            partes.append("Ascenso sistolico intermedio en la escala digitalizada.")
    if sim is not None:
        if sim < 0.5:
            partes.append("Morfologia con ascenso temprano y descenso prolongado.")
        elif sim > 1.2:
            partes.append("Morfologia con ascenso relativamente prolongado respecto del descenso.")
        else:
            partes.append("Morfologia sistolica relativamente balanceada.")
    partes.append("Estas metricas son relativas y no sustituyen IV/IAC/IH absolutos del equipo sin calibracion vertical de dZ/dt.")
    return " ".join(partes)

def calcular_variables_icg_derivadas(datos: Dict[str, Any], morfo: Optional[Dict[str, Any]] = None) -> pd.DataFrame:
    """
    Calcula variables derivadas solo cuando hay datos suficientes.
    Mantiene separados los valores importados del equipo y los recalculados por la app.
    """
    rows: List[Dict[str, Any]] = []

    peso = limpiar_numero(datos.get("peso"))
    talla = limpiar_numero(datos.get("talla"))
    bsa_importada = limpiar_numero(datos.get("superficie_corporal"))
    bsa_calc = calcular_bsa_mosteller(talla, peso)
    bsa = bsa_importada or bsa_calc

    pas = limpiar_numero(datos.get("pas"))
    pad = limpiar_numero(datos.get("pad"))
    fc = limpiar_numero(datos.get("fc"))
    sv = limpiar_numero(datos.get("ds"))
    co_importado = limpiar_numero(datos.get("gc"))
    ci_importado = limpiar_numero(datos.get("ic"))
    z0 = limpiar_numero(datos.get("z0"))
    pep_importado = limpiar_numero(datos.get("pep"))
    lvet_importado = limpiar_numero(datos.get("lvet"))

    # Variables de contractilidad importadas desde el equipo, si estan disponibles.
    # Se mantienen separadas de las variables morfologicas recalculadas por imagen.
    for clave_imp, nombre_imp, unidad_imp in [
        ("iv", "IV importado - indice de velocidad", "u. equipo"),
        ("iac", "IAC importado - indice de aceleracion", "u. equipo"),
        ("ih", "IH importado - indice de Heather", "u. equipo"),
        ("cts", "CTS/STR importado", "ratio"),
    ]:
        v_imp = limpiar_numero(datos.get(clave_imp))
        if v_imp is not None:
            rows.append({
                "Variable": nombre_imp,
                "Valor recalculado": v_imp,
                "Unidad": unidad_imp,
                "Ecuacion": "valor importado del equipo",
                "Estado": "contractilidad importada; no recalculada",
            })

    map_calc = None
    if pas is not None and pad is not None:
        map_calc = (pas + 2 * pad) / 3.0
        rows.append({"Variable": "MAP", "Valor recalculado": map_calc, "Unidad": "mmHg", "Ecuacion": ECUACIONES_ICG["MAP"]["ecuacion"], "Estado": "calculado por PAS/PAD"})

    if peso is not None and talla is not None and talla > 0:
        talla_m = talla / 100.0
        bmi_calc = peso / (talla_m ** 2)
        if 10 <= bmi_calc <= 80:
            rows.append({"Variable": "BMI/IMC", "Valor recalculado": bmi_calc, "Unidad": "kg/m2", "Ecuacion": "BMI = peso_kg / talla_m^2", "Estado": "calculado por talla/peso"})

    if bsa_calc is not None:
        rows.append({"Variable": "BSA Mosteller", "Valor recalculado": bsa_calc, "Unidad": "m2", "Ecuacion": ECUACIONES_ICG["BSA_MOSTELLER"]["ecuacion"], "Estado": "calculado por talla/peso"})

    co_calc = None
    if sv is not None and fc is not None:
        co_calc = sv * fc / 1000.0
        rows.append({"Variable": "CO", "Valor recalculado": co_calc, "Unidad": "L/min", "Ecuacion": ECUACIONES_ICG["CO"]["ecuacion"], "Estado": "calculado por SV/DS y FC"})

    if co_importado is not None and bsa is not None:
        ci_calc = co_importado / bsa
        rows.append({"Variable": "CI desde CO importado", "Valor recalculado": ci_calc, "Unidad": "L/min/m2", "Ecuacion": ECUACIONES_ICG["CI"]["ecuacion"], "Estado": "control de consistencia"})
    elif co_calc is not None and bsa is not None:
        ci_calc = co_calc / bsa
        rows.append({"Variable": "CI", "Valor recalculado": ci_calc, "Unidad": "L/min/m2", "Ecuacion": ECUACIONES_ICG["CI"]["ecuacion"], "Estado": "calculado por CO y BSA"})

    if sv is not None and bsa is not None:
        si_calc = sv / bsa
        rows.append({"Variable": "SI/IDS", "Valor recalculado": si_calc, "Unidad": "mL/lat/m2", "Ecuacion": ECUACIONES_ICG["SI"]["ecuacion"], "Estado": "calculado por SV/DS y BSA"})

    co_base = co_importado or co_calc
    if map_calc is not None and co_base is not None and co_base > 0:
        svr_calc = (map_calc / co_base) * 80.0
        rows.append({"Variable": "SVR/RVS aproximada", "Valor recalculado": svr_calc, "Unidad": "dyn.s.cm-5", "Ecuacion": ECUACIONES_ICG["SVR"]["ecuacion"], "Estado": "CVP no disponible: asumida 0"})

    ci_base = ci_importado
    if ci_base is None and co_base is not None and bsa is not None:
        ci_base = co_base / bsa
    if map_calc is not None and ci_base is not None and ci_base > 0:
        svri_calc = (map_calc / ci_base) * 80.0
        rows.append({"Variable": "SVRI/IRV aproximada", "Valor recalculado": svri_calc, "Unidad": "dyn.s.cm-5.m2", "Ecuacion": ECUACIONES_ICG["SVRI_IRV"]["ecuacion"], "Estado": "CVP no disponible: asumida 0"})

    if map_calc is not None and ci_base is not None and ci_base > 0:
        itc_calc = map_calc * ci_base * 0.0144
        rows.append({"Variable": "ITC aproximado", "Valor recalculado": itc_calc, "Unidad": "kg.m/m2 aprox", "Ecuacion": "ITC = MAP * IC * 0.0144; PCWP no disponible", "Estado": "calculado desde MAP e IC; PCWP asumida 0"})

    if z0 is not None and z0 > 0:
        cft_calc = 1000.0 / z0
        rows.append({"Variable": "CFT/TFC", "Valor recalculado": cft_calc, "Unidad": "1/kOhm aprox", "Ecuacion": ECUACIONES_ICG["CFT_TFC"]["ecuacion"], "Estado": "calculado por Z0"})
        if bsa is not None:
            rows.append({"Variable": "CFTi/TFCi", "Valor recalculado": cft_calc / bsa, "Unidad": "indexado", "Ecuacion": ECUACIONES_ICG["CFTI_TFCI"]["ecuacion"], "Estado": "calculado por Z0 y BSA"})

    # Calculo de Ea/Ees por metodo ICG/Capan-Chen.
    pep_para_vac = pep_importado
    lvet_para_vac = lvet_importado
    fuente_tiempos_vac = "desde tiempos importados del informe"
    if morfo is not None and morfo.get("puntos") is not None:
        pm = morfo["puntos"]
        pep_m = limpiar_numero(pm.get("pep_ms"))
        lvet_m = limpiar_numero(pm.get("lvet_ms"))
        if pep_m is not None and lvet_m is not None:
            pep_para_vac = pep_m
            lvet_para_vac = lvet_m
            fuente_tiempos_vac = "desde cursores morfologicos QRS-B y B-X"

    vac_calc = calcular_acoplamiento_capan_chen(pas, pad, sv, pep_para_vac, lvet_para_vac)
    if vac_calc.get("ef_capan") is not None:
        rows.append({"Variable": "FE Capan", "Valor recalculado": vac_calc["ef_capan"], "Unidad": "fraccion", "Ecuacion": ECUACIONES_ICG["EF_CAPAN"]["ecuacion"], "Estado": fuente_tiempos_vac})
    if vac_calc.get("ea") is not None:
        rows.append({"Variable": "Ea recalculada", "Valor recalculado": vac_calc["ea"], "Unidad": "mmHg/mL", "Ecuacion": ECUACIONES_ICG["EA"]["ecuacion"], "Estado": "PAS*0.9 y SV/DS"})
    if vac_calc.get("end_avg") is not None:
        rows.append({"Variable": "End(avg) Chen", "Valor recalculado": vac_calc["end_avg"], "Unidad": "adimensional", "Ecuacion": ECUACIONES_ICG["END_AVG_CHEN"]["ecuacion"], "Estado": fuente_tiempos_vac})
    if vac_calc.get("end_est") is not None:
        rows.append({"Variable": "End(est) Chen", "Valor recalculado": vac_calc["end_est"], "Unidad": "adimensional", "Ecuacion": ECUACIONES_ICG["END_EST_CHEN"]["ecuacion"], "Estado": "EF Capan + PA + End(avg)"})
    if vac_calc.get("ees") is not None:
        rows.append({"Variable": "Ees recalculada", "Valor recalculado": vac_calc["ees"], "Unidad": "mmHg/mL", "Ecuacion": ECUACIONES_ICG["EES_CHEN_ICG"]["ecuacion"], "Estado": "metodo single-beat Chen aplicado a ICG"})
    if vac_calc.get("vac") is not None:
        rows.append({"Variable": "Acoplamiento VA Ea/Ees", "Valor recalculado": vac_calc["vac"], "Unidad": "ratio", "Ecuacion": ECUACIONES_ICG["VAC_EA_EES"]["ecuacion"], "Estado": interpretar_acoplamiento_vac(vac_calc["vac"])})

    if morfo is not None and morfo.get("puntos") is not None:
        p = morfo["puntos"]
        pep = limpiar_numero(p.get("pep_ms"))
        lvet = limpiar_numero(p.get("lvet_ms"))
        if pep is not None:
            rows.append({"Variable": "PEP morfologico", "Valor recalculado": pep, "Unidad": "ms", "Ecuacion": ECUACIONES_ICG["PEP"]["ecuacion"], "Estado": "desde cursores QRS-B"})
        if lvet is not None:
            rows.append({"Variable": "LVET morfologico", "Valor recalculado": lvet, "Unidad": "ms", "Ecuacion": ECUACIONES_ICG["LVET"]["ecuacion"], "Estado": "desde cursores B-X"})
        if pep is not None and lvet is not None and lvet > 0:
            rows.append({"Variable": "STR/CTS morfologico", "Valor recalculado": pep / lvet, "Unidad": "ratio", "Ecuacion": ECUACIONES_ICG["STR"]["ecuacion"], "Estado": "desde PEP/LVET"})

        df_curva_m = morfo.get("df_curva")
        vars_ct = calcular_contractilidad_morfologica(df_curva_m, p, z0=z0)
        estado_ct = "desde curva dZ/dt digitalizada con cursores QRS-B-C-X"
        mapa_ct = [
            ("dZ/dt maximo relativo", vars_ct.get("dzdt_max_rel"), "u. relativa", "DZDT_MAX_REL"),
            ("Indice de Heather relativo", vars_ct.get("hi_rel"), "u.rel/ms", "HI_REL"),
            ("Indice de velocidad relativo", vars_ct.get("vi_rel"), "u.rel/Z0", "VI_REL"),
            ("Indice de aceleracion relativo", vars_ct.get("aci_rel"), "u.rel/ms", "ACI_REL"),
            ("Pendiente ascendente B-C", vars_ct.get("pendiente_ascendente"), "u.rel/ms", "PENDIENTE_ASC"),
            ("Pendiente descendente C-X", vars_ct.get("pendiente_descendente"), "u.rel/ms", "PENDIENTE_DESC"),
            ("Area sistolica relativa B-X", vars_ct.get("area_sistolica_rel"), "u.rel*ms", "AREA_SISTOLICA_REL"),
            ("Simetria sistolica B-C/C-X", vars_ct.get("simetria_sistolica"), "ratio", "SIMETRIA_SISTOLICA"),
        ]
        for nom, val, uni, eqk in mapa_ct:
            if val is not None and np.isfinite(val):
                rows.append({"Variable": nom, "Valor recalculado": val, "Unidad": uni, "Ecuacion": ECUACIONES_ICG[eqk]["ecuacion"], "Estado": estado_ct})
        interp_ct = interpretar_contractilidad_morfologica(vars_ct)
        if vars_ct.get("dzdt_max_rel") is not None:
            rows.append({"Variable": "Interpretacion contractilidad morfologica", "Valor recalculado": interp_ct, "Unidad": "texto", "Ecuacion": "integracion morfologica", "Estado": "interpretacion clinica prudente"})

    df = pd.DataFrame(rows)
    if df.empty:
        return pd.DataFrame(columns=["Variable", "Valor recalculado", "Unidad", "Ecuacion", "Estado"])
    df["Valor recalculado"] = df["Valor recalculado"].apply(lambda x: round(float(x), 6) if limpiar_numero(x) is not None else x)
    df = aplicar_formato_y_coherencia(df)
    return df


def texto_ecuaciones_derivadas(df_der: pd.DataFrame) -> str:
    if df_der is None or df_der.empty:
        return "No se recalcularon variables por falta de datos suficientes."
    partes = []
    for _, r in df_der.iterrows():
        val = r.get("Valor recalculado")
        unidad = r.get("Unidad", "")
        partes.append(f"{r.get('Variable')}: {val} {unidad}")
    return "; ".join(partes) + "."


def _valor_derivado(df_der: pd.DataFrame, nombre: str) -> Optional[float]:
    """Busca una variable en la tabla de derivadas y devuelve valor numerico."""
    if df_der is None or df_der.empty:
        return None
    nombre_n = normalizar_txt(nombre)
    # Primero coincidencia exacta para evitar confusiones: IC no debe coincidir con IAC/CFTi.
    for _, r in df_der.iterrows():
        vname = normalizar_txt(r.get("Variable", ""))
        if vname == nombre_n:
            return limpiar_numero(r.get("Valor recalculado"))
    # Luego coincidencia parcial solo con nombres suficientemente específicos.
    if len(nombre_n) >= 4:
        for _, r in df_der.iterrows():
            vname = normalizar_txt(r.get("Variable", ""))
            if nombre_n in vname:
                return limpiar_numero(r.get("Valor recalculado"))
    return None


def _filtrar_derivadas_por_palabras(df_der: pd.DataFrame, palabras: List[str]) -> pd.DataFrame:
    if df_der is None or df_der.empty:
        return pd.DataFrame(columns=["Variable", "Valor recalculado", "Unidad", "Ecuacion", "Estado"])
    pats = [normalizar_txt(p) for p in palabras]
    mask = []
    for _, r in df_der.iterrows():
        texto = normalizar_txt(str(r.get("Variable", "")) + " " + str(r.get("Estado", "")) + " " + str(r.get("Ecuacion", "")) + " " + str(r.get("Control de coherencia", "")))
        mask.append(any(p in texto for p in pats))
    return df_der.loc[mask].copy()


def mostrar_metricas_hemodinamicas_digitalizacion(df_der: pd.DataFrame, puntos: Dict[str, Any]) -> None:
    """
    Panel visible y completo de metricas hemodinamicas calculadas a partir de la
    curva dZ/dt corregida semiautomatica y de las variables basales del informe.
    Incluye FE Capan en forma destacada.
    """
    if df_der is None or df_der.empty:
        st.info("No hay metricas hemodinamicas recalculadas para mostrar.")
        return

    st.markdown("### Panel completo de metricas hemodinamicas derivadas de la curva corregida")
    st.caption(
        "Estas metricas combinan la digitalizacion semiautomatica QRS-B-C-X con los datos basales del informe. "
        "Las variables morfologicas dependen de la confirmacion visual de los cursores."
    )

    fe = _valor_derivado(df_der, "FE Capan")
    pep = _valor_derivado(df_der, "PEP morfologico")
    lvet = _valor_derivado(df_der, "LVET morfologico")
    cts = _valor_derivado(df_der, "STR/CTS morfologico")
    ea = _valor_derivado(df_der, "Ea recalculada")
    ees = _valor_derivado(df_der, "Ees recalculada")
    vac = _valor_derivado(df_der, "Acoplamiento VA")
    ci = _valor_derivado(df_der, "CI desde CO importado") or _valor_derivado(df_der, "CI")
    irv = _valor_derivado(df_der, "SVRI/IRV aproximada")
    dzdt_rel = _valor_derivado(df_der, "dZ/dt maximo relativo")
    hi_rel = _valor_derivado(df_der, "Indice de Heather relativo")

    m1, m2, m3, m4 = st.columns(4)
    with m1:
        st.metric("FE Capan", "No disponible" if fe is None else f"{fe*100:.1f}%")
    with m2:
        st.metric("PEP QRS-B", "No disponible" if pep is None else f"{pep:.0f} ms")
    with m3:
        st.metric("LVET B-X", "No disponible" if lvet is None else f"{lvet:.0f} ms")
    with m4:
        st.metric("CTS/STR", "No disponible" if cts is None else f"{cts:.3f}")

    m5, m6, m7, m8 = st.columns(4)
    with m5:
        st.metric("Ea", "No disponible" if ea is None else f"{ea:.3f} mmHg/mL")
    with m6:
        st.metric("Ees", "No disponible" if ees is None else f"{ees:.3f} mmHg/mL")
    with m7:
        st.metric("Ea/Ees", "No disponible" if vac is None else f"{vac:.3f}")
    with m8:
        st.metric("IC recalculado", "No disponible" if ci is None else f"{ci:.2f} L/min/m²")

    m9, m10, m11, m12 = st.columns(4)
    with m9:
        st.metric("IRV recalculado", "No disponible" if irv is None else f"{irv:.0f} dyn·s·cm⁻5·m²")
    with m10:
        st.metric("dZ/dt máx relativo", "No disponible" if dzdt_rel is None else f"{dzdt_rel:.3f}")
    with m11:
        st.metric("Heather relativo", "No disponible" if hi_rel is None else f"{hi_rel:.5f}")
    with m12:
        st.metric("Validación", "Confirmada" if bool(puntos.get("validado_operador", False)) else "Preliminar")

    tabs = st.tabs([
        "Todas las métricas",
        "Tiempos + FE Capan",
        "Ea/Ees",
        "IC/IRV/BSA",
        "Contractilidad morfológica",
    ])
    with tabs[0]:
        mostrar_tabla_metricas_con_barras(df_der, height=560)
    with tabs[1]:
        df_t = _filtrar_derivadas_por_palabras(df_der, ["pep", "lvet", "str", "cts", "fe capan", "end(avg)", "end(est)"])
        mostrar_tabla_metricas_con_barras(df_t, height=360)
    with tabs[2]:
        df_v = _filtrar_derivadas_por_palabras(df_der, ["ea", "ees", "acoplamiento", "chen", "vac"])
        mostrar_tabla_metricas_con_barras(df_v, height=360)
    with tabs[3]:
        df_i = _filtrar_derivadas_por_palabras(df_der, ["bsa", "bmi", "map", "co", "ci", "si/ids", "svr", "rvs", "svri", "irv", "cft"])
        mostrar_tabla_metricas_con_barras(df_i, height=420)
    with tabs[4]:
        df_c = _filtrar_derivadas_por_palabras(df_der, ["contractilidad", "heather", "velocidad", "aceleracion", "pendiente", "area sistolica", "simetria", "dz/dt"])
        mostrar_tabla_metricas_con_barras(df_c, height=460)


# ============================================================
# MODULO MORFOLOGICO SEMIAUTOMATICO DE SEÑAL ICG
# ============================================================

def obtener_fc_desde_df(df: pd.DataFrame) -> Optional[float]:
    if df is None or df.empty:
        return None
    for _, r in df.iterrows():
        v = limpiar_numero(r.get("fc"))
        if v is not None and 35 <= v <= 190:
            return v
    return None



def obtener_valor_basal_desde_df(df: pd.DataFrame, clave: str) -> Optional[float]:
    """Obtiene un valor numerico de la fila basal/acostado/cinta/spot si existe."""
    if df is None or df.empty:
        return None
    try:
        basal, _ = seleccionar_basal_y_parado(df)
        if basal:
            v = limpiar_numero(basal.get(clave))
            if v is not None:
                return v
    except Exception:
        pass
    for _, r in df.iterrows():
        v = limpiar_numero(r.get(clave))
        if v is not None:
            return v
    return None


def obtener_rr_desde_df(df: pd.DataFrame) -> Optional[float]:
    """Recupera RR en ms desde el informe Z-Logic si fue extraido."""
    if df is None or df.empty:
        return None
    for clave in ["rr", "RR", "intervalo_rr"]:
        v = obtener_valor_basal_desde_df(df, clave)
        if v is not None and 300 <= v <= 2000:
            return float(v)
    # búsqueda textual de respaldo
    for _, r in df.iterrows():
        txt = " ".join(str(x) for x in r.values)
        m = re.search(r"\bRR\s*[:=]?\s*(\d{3,4})\b", txt, re.I)
        if m:
            val = limpiar_numero(m.group(1))
            if val is not None and 300 <= val <= 2000:
                return float(val)
    return None


def calcular_bmi_desde_talla_peso(talla_cm: Optional[float], peso_kg: Optional[float]) -> Optional[float]:
    if talla_cm is None or peso_kg is None:
        return None
    if not (90 <= talla_cm <= 230 and 25 <= peso_kg <= 250):
        return None
    m = talla_cm / 100.0
    if m <= 0:
        return None
    return float(peso_kg / (m * m))


def completar_antropometria_e_indices(df: pd.DataFrame) -> pd.DataFrame:
    """
    Completa BMI/IMC, BSA, IC e IRV cuando sea posible.
    Mantiene separados los valores importados de los recalculados en la tabla de ecuaciones,
    pero aquí rellena campos vacíos para que el informe quede completo.
    """
    if df is None or df.empty:
        return df
    out = df.copy()
    for i, row in out.iterrows():
        peso = limpiar_numero(row.get("peso"))
        talla = limpiar_numero(row.get("talla"))
        bsa = limpiar_numero(row.get("superficie_corporal"))
        bmi = limpiar_numero(row.get("imc"))
        if bsa is None:
            bsa = calcular_bsa_mosteller(talla, peso)
            if bsa is not None:
                out.at[i, "superficie_corporal"] = round(bsa, 3)
        if bmi is None:
            bmi = calcular_bmi_desde_talla_peso(talla, peso)
            if bmi is not None:
                out.at[i, "imc"] = round(bmi, 1)
        co = limpiar_numero(row.get("gc"))
        ci = limpiar_numero(row.get("ic"))
        rvs = limpiar_numero(row.get("rvs"))
        irv = limpiar_numero(row.get("irv"))
        if ci is None and co is not None and bsa is not None and bsa > 0:
            out.at[i, "ic"] = round(co / bsa, 3)
        if irv is None and rvs is not None and bsa is not None and bsa > 0:
            out.at[i, "irv"] = round(rvs * bsa, 0)
        ds = limpiar_numero(row.get("ds"))
        ids = limpiar_numero(row.get("ids"))
        if ids is None and ds is not None and bsa is not None and bsa > 0:
            out.at[i, "ids"] = round(ds / bsa, 2)
    return out


def renderizar_pdf_a_imagen(pdf_bytes: bytes, pagina: int = 0, zoom: float = 2.5) -> Image.Image:
    """Renderiza una página PDF a PIL.Image usando PyMuPDF."""
    if fitz is None:
        raise RuntimeError("PyMuPDF no esta instalado. Agregar PyMuPDF a requirements.txt")
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    if len(doc) == 0:
        raise ValueError("PDF sin paginas renderizables")
    page = doc[min(max(0, pagina), len(doc)-1)]
    mat = fitz.Matrix(float(zoom), float(zoom))
    pix = page.get_pixmap(matrix=mat, alpha=False)
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    doc.close()
    return img


def crop_relativo(imagen: Image.Image, x1: float, y1: float, x2: float, y2: float) -> Image.Image:
    w, h = imagen.size
    return imagen.crop((int(x1*w), int(y1*h), int(x2*w), int(y2*h)))



def _digitalizar_fono_inferior_rescate(
    roi_rgb: np.ndarray,
    y_ini_frac: float,
    y_fin_frac: float,
    duracion_ms: float,
    suavizado: int = 1,
    nombre_canal: str = "fono_inferior",
) -> pd.DataFrame:
    """
    Rescate robusto para el fonocardiograma inferior Z-Logic.

    En algunos PDF el fono no se renderiza como trazo azul puro sino como trazo
    morado/oscuro de muy baja amplitud, superpuesto a grilla gris y a una franja
    amarilla. Este método NO exige azul: reconstruye la señal desde los píxeles
    más oscuros por columna, excluyendo amarillo, bordes y textos evidentes.
    Nunca debe detener la app: si la señal es insuficiente devuelve una línea
    basal con advertencia en la columna calidad.
    """
    h, w = roi_rgb.shape[:2]
    y1 = int(max(0, min(h - 1, y_ini_frac * h)))
    y2 = int(max(y1 + 6, min(h, y_fin_frac * h)))
    banda = roi_rgb[y1:y2, :, :].astype(float)
    bh, bw = banda.shape[:2]

    r = banda[:, :, 0]
    g = banda[:, :, 1]
    b = banda[:, :, 2]
    gris = (r + g + b) / 3.0
    maxc = np.maximum.reduce([r, g, b])
    minc = np.minimum.reduce([r, g, b])
    sat = maxc - minc

    # Excluir franja amarilla típica del canal inferior.
    amarillo = (r > 135) & (g > 110) & (b < 155) & (sat > 25)

    # Excluir texto negro muy denso en extremo derecho/bajo si aparece.
    # El trazo del fono es fino; los textos tienen columnas/filas muy densas.
    oscuro = (gris < 180) & (~amarillo)
    if oscuro.sum() < 10:
        oscuro = (gris < 210) & (~amarillo)

    # Quitar filas/columnas densas: suelen ser letras, bordes o grilla fuerte.
    mask = oscuro.copy()
    if mask.sum() > 0:
        row_density = mask.mean(axis=1)
        col_density = mask.mean(axis=0)
        mask[row_density > 0.45, :] = False
        mask[:, col_density > 0.55] = False

    # Evitar bordes de la banda y últimos píxeles donde suele estar texto.
    by = max(1, int(0.04 * bh))
    bx = max(1, int(0.02 * bw))
    mask[:by, :] = False
    mask[-by:, :] = False
    mask[:, :bx] = False
    mask[:, -bx:] = False

    # Score: preferir píxeles oscuros y ligeramente azulados/morados.
    blueness = b - 0.5 * (r + g)
    score = (255.0 - gris) + np.maximum(blueness, 0) * 0.35
    score[~mask] = -1e9

    # Basal estimada: fila con mayor continuidad de candidatos en zona izquierda/central.
    central_cols = slice(int(0.03 * bw), int(0.72 * bw))
    dens_rows = mask[:, central_cols].mean(axis=1) if bw > 20 else mask.mean(axis=1)
    if np.nanmax(dens_rows) > 0:
        baseline_y = float(np.argmax(dens_rows))
    else:
        baseline_y = bh / 2.0

    xs = []
    ys = []
    prev_y = baseline_y
    max_jump = max(2.0, 0.12 * bh)

    for x in range(bw):
        cand = np.where(mask[:, x])[0]
        if len(cand) == 0:
            # mantener continuidad; no cortar la señal.
            xs.append(float(x)); ys.append(float(prev_y))
            continue
        dist = np.abs(cand.astype(float) - prev_y)
        # No permitir saltos a letras o marcas lejanas salvo que el score sea muy alto.
        local = score[cand, x] - 2.2 * dist
        y_sel = float(cand[int(np.argmax(local))])
        if abs(y_sel - prev_y) > max_jump:
            # elegir candidato más cercano si el mejor salta demasiado.
            y_sel = float(cand[int(np.argmin(dist))])
        xs.append(float(x)); ys.append(y_sel)
        prev_y = 0.82 * prev_y + 0.18 * y_sel

    yinterp = np.asarray(ys, dtype=float)
    # En imagen y crece hacia abajo; invertir para que deflexión superior sea positiva.
    amp_px = -(yinterp - np.nanmedian(yinterp))
    amp_px = suavizar_senal(amp_px, max(1, min(int(suavizado), 3)))

    # Si la amplitud quedó prácticamente nula, dejar trazado basal pero no bloquear.
    calidad = "rescate_fono"
    if not np.isfinite(amp_px).all() or float(np.nanstd(amp_px)) < 0.05:
        amp_px = np.zeros(bw, dtype=float)
        calidad = "fono_no_confiable_linea_basal"

    t = np.linspace(0, float(duracion_ms), bw)
    return pd.DataFrame({
        "tiempo_ms": t,
        "amplitud": amp_px,
        "amplitud_px": amp_px,
        "unidad_y": "px_relativos",
        "calidad": calidad,
    })



def _digitalizar_trazo_rescate_no_bloqueante(
    roi_rgb: np.ndarray,
    y_ini_frac: float,
    y_fin_frac: float,
    duracion_ms: float,
    nombre_canal: str = "dzdt",
    suavizado: int = 3,
    invertir: bool = False,
    amplitud_mm_por_px: Optional[float] = None,
    escala_y_por_mm: Optional[float] = None,
) -> pd.DataFrame:
    """
    Rescate robusto y no bloqueante para trazos Z-Logic.

    Se usa cuando el detector principal no encuentra suficientes píxeles de curva.
    A diferencia del método principal, no exige que el trazo sea claramente azul:
    combina azul/morado + oscuridad + continuidad, penalizando grilla, texto y
    relleno amarillo. Si aun así no logra seguir una curva confiable, devuelve
    una línea basal marcada como baja confianza para que la app no se detenga y
    el operador pueda usar el modo semiautomático/manual.
    """
    h, w = roi_rgb.shape[:2]
    y1 = int(max(0, min(h - 1, y_ini_frac * h)))
    y2 = int(max(y1 + 4, min(h, y_fin_frac * h)))
    banda = roi_rgb[y1:y2, :, :].astype(float)
    bh, bw = banda.shape[:2]
    r = banda[:, :, 0]
    g = banda[:, :, 1]
    b = banda[:, :, 2]
    gris = (r + g + b) / 3.0
    maxc = np.maximum.reduce([r, g, b])
    minc = np.minimum.reduce([r, g, b])
    sat = maxc - minc

    # Excluir amarillos y grilla muy clara. Mantener trazos azul oscuro/morados.
    mask_amarillo = (r > 145) & (g > 120) & (b < 145) & (sat > 25)
    mask_blanco_grilla = (gris > 232) | ((sat < 7) & (gris > 170))

    # Score continuo: favorece curva azul/morada y trazos oscuros no amarillos.
    blueness = b - 0.5 * (r + g)
    darkness = 255.0 - gris
    score = 1.7 * np.maximum(blueness, 0) + 0.55 * darkness + 0.20 * sat
    score[mask_amarillo | mask_blanco_grilla] = -1e9

    # Evitar bordes y zonas densas típicas de texto/ejes.
    by = max(1, int(0.03 * bh))
    bx = max(1, int(0.01 * bw))
    score[:by, :] = -1e9
    score[-by:, :] = -1e9
    score[:, :bx] = -1e9
    score[:, -bx:] = -1e9

    # Umbral adaptativo por percentil: más flexible para dZ/dt bajo contraste.
    valid = np.isfinite(score) & (score > -1e8)
    if valid.sum() < 20:
        amp_px = np.zeros(bw, dtype=float)
        t = np.linspace(0, float(duracion_ms), bw)
        return pd.DataFrame({"tiempo_ms": t, "amplitud": amp_px, "amplitud_px": amp_px,
                             "unidad_y": "px_relativos", "calidad": f"{nombre_canal}_no_confiable_linea_basal"})

    vals = score[valid]
    thr = np.nanpercentile(vals, 97.0)
    if not np.isfinite(thr):
        thr = np.nanmax(vals) - 1

    mask = score >= thr
    # Si queda demasiado poco, relajar umbral.
    if mask.sum() < max(20, int(0.004 * bh * bw)):
        thr = np.nanpercentile(vals, 94.0)
        mask = score >= thr

    xs: List[float] = []
    ys: List[float] = []
    prev_y: Optional[float] = None
    max_jump = max(5, int(0.16 * bh)) if "dz" in nombre_canal.lower() else max(3, int(0.10 * bh))

    for x in range(bw):
        cand = np.where(mask[:, x])[0]
        if len(cand) == 0:
            continue
        if prev_y is None:
            y_sel = float(cand[int(np.argmax(score[cand, x]))])
        else:
            dist = np.abs(cand.astype(float) - prev_y)
            local = score[cand, x] - 1.8 * dist
            y_sel = float(cand[int(np.argmax(local))])
            if abs(y_sel - prev_y) > max_jump:
                # no saltar a texto; usar el más cercano si sigue teniendo score aceptable
                close = cand[np.argmin(dist)]
                if score[close, x] >= thr * 0.75:
                    y_sel = float(close)
                else:
                    continue
        xs.append(float(x))
        ys.append(y_sel)
        prev_y = y_sel if prev_y is None else 0.72 * prev_y + 0.28 * y_sel

    # Si no alcanza, tomar centroide de los mejores píxeles por columna.
    if len(xs) < 12:
        xs, ys = [], []
        for x in range(bw):
            col = score[:, x]
            ok = np.where(col > np.nanpercentile(vals, 92.0))[0]
            if len(ok):
                weights = col[ok] - np.nanmin(col[ok]) + 1.0
                xs.append(float(x))
                ys.append(float(np.average(ok, weights=weights)))

    if len(xs) < 8:
        amp_px = np.zeros(bw, dtype=float)
        calidad = f"{nombre_canal}_no_confiable_linea_basal"
    else:
        xs_arr = np.asarray(xs, dtype=float)
        ys_arr = np.asarray(ys, dtype=float)
        xfull = np.arange(bw, dtype=float)
        yinterp = np.interp(xfull, xs_arr, ys_arr)
        amp_px = -yinterp
        if invertir:
            amp_px = -amp_px
        amp_px = amp_px - np.nanmedian(amp_px)
        amp_px = suavizar_senal(amp_px, max(1, min(int(suavizado), 5)))
        calidad = f"{nombre_canal}_rescate_adaptativo"

    amp = amp_px.copy()
    unidad = "px_relativos"
    if amplitud_mm_por_px is not None and escala_y_por_mm is not None:
        amp = amp_px * float(amplitud_mm_por_px) * float(escala_y_por_mm)
        unidad = "calibrada"

    t = np.linspace(0, float(duracion_ms), len(amp))
    return pd.DataFrame({"tiempo_ms": t, "amplitud": amp, "amplitud_px": amp_px,
                         "unidad_y": unidad, "calidad": calidad})

def _digitalizar_banda_trazo(
    roi_rgb: np.ndarray,
    y_ini_frac: float,
    y_fin_frac: float,
    duracion_ms: float,
    umbral_oscuridad: int = 150,
    suavizado: int = 5,
    invertir: bool = False,
    max_salto_px: Optional[int] = None,
    amplitud_mm_por_px: Optional[float] = None,
    escala_y_por_mm: Optional[float] = None,
    preferir_trazo_azul: bool = True,
    excluir_relleno_amarillo: bool = True,
    nombre_canal: str = "",
) -> pd.DataFrame:
    """
    Digitaliza una banda horizontal del panel Z-Logic siguiendo SOLO el trazo azul
    cuando este existe. Esta corrección evita que el ECG y el fonocardiograma se
    transformen en escalones por haber capturado grilla, texto, números o la
    franja amarilla inferior.

    Principio:
    - La curva real de Z-Logic es azul/morada.
    - La grilla es gris: R≈G≈B.
    - Los textos son negros: R≈G≈B bajos.
    - El resaltado amarillo no debe participar.
    """
    h, w = roi_rgb.shape[:2]
    y1 = int(max(0, min(h - 1, y_ini_frac * h)))
    y2 = int(max(y1 + 3, min(h, y_fin_frac * h)))
    banda = roi_rgb[y1:y2, :, :].astype(float)
    bh, bw = banda.shape[:2]

    r = banda[:, :, 0]
    g = banda[:, :, 1]
    b = banda[:, :, 2]
    gris = (r + g + b) / 3.0
    maxc = np.maximum.reduce([r, g, b])
    minc = np.minimum.reduce([r, g, b])
    sat = maxc - minc

    # Excluir relleno amarillo del canal inferior y textos muy negros.
    mask_amarillo = (r > 145) & (g > 125) & (b < 140) & (sat > 35)

    # Máscara azul/morada estricta: captura el trazo real y descarta grilla/textos.
    # Incluye antialiasing claro de la curva: azul con diferencia respecto a R/G.
    blueness = (b - 0.5 * (r + g))
    mask_azul_estricta = (
        (b >= 95) &
        (blueness >= 18) &
        (b >= r + 18) &
        (b >= g + 10) &
        (gris < 245)
    )
    if excluir_relleno_amarillo:
        mask_azul_estricta &= ~mask_amarillo

    # Máscara de rescate para trazos azul oscuro de baja saturación.
    mask_azul_rescate = (
        (b >= 80) &
        (b >= r + 12) &
        (b >= g + 6) &
        (gris < 235)
    )
    if excluir_relleno_amarillo:
        mask_azul_rescate &= ~mask_amarillo

    mask = mask_azul_estricta if mask_azul_estricta.sum() >= 20 else mask_azul_rescate

    # Si no hay trazo azul suficiente, usar rescate por trazo oscuro.
    # Importante: en algunos PDF Z-Logic el fonocardiograma inferior no queda azul,
    # sino casi negro/morado sobre la grilla; exigir saturación lo borra y produce
    # el error: "No se detectó trazo azul suficiente en canal fono_inferior".
    if mask.sum() < 20:
        es_fono = "fono" in str(nombre_canal).lower()
        if es_fono:
            # Rescate específico del canal inferior: permitir trazo oscuro aunque
            # tenga baja saturación. Mantener fuera el relleno amarillo y la grilla clara.
            mask_oscura = (gris < max(150, umbral_oscuridad))
        else:
            # Para dZ/dt y ECG mantener cierta saturación para no tomar grilla/texto.
            mask_oscura = (gris < umbral_oscuridad) & ((sat > 8) | (b >= r + 8))
        if excluir_relleno_amarillo:
            mask_oscura &= ~mask_amarillo
        mask = mask_oscura

    # Eliminar líneas verticales/horizontales extensas de grilla o marcas, pero con umbrales
    # menos agresivos para no borrar un fono de baja amplitud.
    mask = np.asarray(mask, dtype=bool).copy()
    if mask.size == 0 or mask.sum() < 10:
        # Último rescate no fatal para fono: seleccionar los pixeles más oscuros por columna
        # dentro de la banda inferior y dejar que la interpolación reconstruya una línea.
        if "fono" in str(nombre_canal).lower():
            p_dark = np.nanpercentile(gris, 1.2)
            mask = (gris <= p_dark + 8)
            if excluir_relleno_amarillo:
                mask &= ~mask_amarillo
        if mask.size == 0 or mask.sum() < 10:
            return _digitalizar_trazo_rescate_no_bloqueante(
                roi_rgb=roi_rgb,
                y_ini_frac=y_ini_frac,
                y_fin_frac=y_fin_frac,
                duracion_ms=duracion_ms,
                nombre_canal=nombre_canal,
                suavizado=suavizado,
                invertir=invertir,
                amplitud_mm_por_px=amplitud_mm_por_px,
                escala_y_por_mm=escala_y_por_mm,
            )

    row_density = mask.mean(axis=1)
    col_density = mask.mean(axis=0)
    mask[row_density > 0.60, :] = False
    mask[:, col_density > 0.70] = False
    by = max(1, int(0.005 * bh))
    bx = max(1, int(0.005 * bw))
    mask[:by, :] = False
    mask[-by:, :] = False
    mask[:, :bx] = False
    mask[:, -bx:] = False

    # Construcción de trayectoria por columna usando score de azul. Esto conserva
    # picos angostos de ECG/fono mejor que la mediana de todos los pixeles.
    xs: List[float] = []
    ys: List[float] = []
    prev_y: Optional[float] = None
    if max_salto_px is None:
        max_salto_px = max(4, int(0.08 * bh))

    score = blueness.copy()
    score[~mask] = -1e9

    for x in range(bw):
        cand = np.where(mask[:, x])[0]
        if len(cand) == 0:
            continue
        if prev_y is None:
            # Usar el pixel más azul de la columna, no la mediana.
            y_sel = float(cand[int(np.argmax(score[cand, x]))])
        else:
            # Combinar continuidad con intensidad azul; evita saltar a texto o grilla.
            dist = np.abs(cand.astype(float) - prev_y)
            local_score = score[cand, x]
            local_score = local_score - 3.0 * dist
            y_sel = float(cand[int(np.argmax(local_score))])
            if abs(y_sel - prev_y) > max_salto_px:
                # Permitir picos agudos breves si son muy azules; si no, dejar hueco.
                y_best = float(cand[int(np.argmax(score[cand, x]))])
                if abs(y_best - prev_y) <= max_salto_px * 2.2:
                    y_sel = y_best
                else:
                    continue
        xs.append(float(x))
        ys.append(y_sel)
        prev_y = y_sel if prev_y is None else (0.65 * prev_y + 0.35 * y_sel)

    if len(xs) < 15:
        # Fallback: para fono de muy baja amplitud, usar centroide azul por columna.
        xs = []
        ys = []
        for x in range(bw):
            cand = np.where(mask[:, x])[0]
            if len(cand) > 0:
                sc = np.maximum(score[cand, x], 1.0)
                if np.all(sc <= 0):
                    y = float(np.median(cand))
                else:
                    y = float(np.average(cand, weights=sc - np.min(sc) + 1.0))
                xs.append(float(x)); ys.append(y)

    if len(xs) < 15:
        if "fono" in str(nombre_canal).lower():
            return _digitalizar_fono_inferior_rescate(
                roi_rgb=roi_rgb,
                y_ini_frac=y_ini_frac,
                y_fin_frac=y_fin_frac,
                duracion_ms=duracion_ms,
                suavizado=1,
                nombre_canal=nombre_canal,
            )
        return _digitalizar_trazo_rescate_no_bloqueante(
            roi_rgb=roi_rgb,
            y_ini_frac=y_ini_frac,
            y_fin_frac=y_fin_frac,
            duracion_ms=duracion_ms,
            nombre_canal=nombre_canal,
            suavizado=suavizado,
            invertir=invertir,
            amplitud_mm_por_px=amplitud_mm_por_px,
            escala_y_por_mm=escala_y_por_mm,
        )

    xs_arr = np.asarray(xs, dtype=float)
    ys_arr = np.asarray(ys, dtype=float)
    xfull = np.arange(bw, dtype=float)
    yinterp = np.interp(xfull, xs_arr, ys_arr)

    amp_px = -yinterp
    if invertir:
        amp_px = -amp_px
    amp_px = amp_px - np.nanmedian(amp_px)

    # Para ECG/fono no sobre-suavizar: se perderían QRS/S1/S2.
    suav = int(suavizado)
    if "ecg" in nombre_canal.lower() or "fono" in nombre_canal.lower():
        suav = min(suav, 3)
    amp_px = suavizar_senal(amp_px, suav)

    amp = amp_px.copy()
    unidad = "px_relativos"
    if amplitud_mm_por_px is not None and escala_y_por_mm is not None:
        amp = amp_px * float(amplitud_mm_por_px) * float(escala_y_por_mm)
        unidad = "calibrada"

    t = np.linspace(0, float(duracion_ms), len(amp))
    return pd.DataFrame({"tiempo_ms": t, "amplitud": amp, "amplitud_px": amp_px, "unidad_y": unidad})

def estimar_espaciado_grilla_px(imagen_roi: Image.Image) -> Optional[float]:
    """Estimación simple del espaciado de grilla en pixeles. Devuelve None si no es confiable."""
    arr = np.asarray(imagen_roi.convert("RGB")).astype(float)
    gris = np.mean(arr, axis=2)
    # La grilla suele ser gris clara: mucha densidad de pixeles intermedios por fila/columna.
    edges_v = ((gris > 135) & (gris < 235)).mean(axis=0)
    peaks = np.where(edges_v > np.percentile(edges_v, 88))[0]
    if len(peaks) < 5:
        return None
    # agrupar picos contiguos
    grupos = []
    ini = peaks[0]
    prev = peaks[0]
    for p in peaks[1:]:
        if p <= prev + 1:
            prev = p
        else:
            grupos.append((ini, prev))
            ini = prev = p
    grupos.append((ini, prev))
    centers = np.array([(a+b)/2 for a,b in grupos], dtype=float)
    diffs = np.diff(centers)
    diffs = diffs[(diffs >= 2) & (diffs <= 60)]
    if len(diffs) == 0:
        return None
    return float(np.median(diffs))


def digitalizar_pdf_zlogic_sector_superior_derecho(
    pdf_bytes: bytes,
    rr_ms: Optional[float] = None,
    pagina: int = 0,
    zoom: float = 2.8,
    crop_rel: Tuple[float, float, float, float] = (0.735, 0.292, 0.985, 0.675),
    suavizado: int = 5,
    invertir_dzdt: bool = False,
) -> Dict[str, Any]:
    """
    Digitaliza automaticamente el panel superior derecho del informe Z-Logic.
    Devuelve curva dZ/dt, ECG y Fono alineados temporalmente.
    """
    img = renderizar_pdf_a_imagen(pdf_bytes, pagina=pagina, zoom=zoom)
    rr = float(rr_ms) if rr_ms is not None and 300 <= float(rr_ms) <= 2000 else 800.0
    panel = crop_relativo(img, *crop_rel)
    arr = np.asarray(panel.convert("RGB"))
    h, w = arr.shape[:2]

    # Estimar px/mm de grilla si es posible. En papel estándar 25 mm/s: 20 ms = 0.5 mm.
    grid_px = estimar_espaciado_grilla_px(panel)
    # Si la grilla detectada parece corresponder a 1 mm, usar 1/grid_px mm/px.
    mm_por_px = (1.0 / grid_px) if grid_px and grid_px > 0 else None
    # Z-Logic informa en este ejemplo 0.5 ohm/seg/mm para dZ/dt.
    escala_dzdt = 0.5

    # Bandas observadas en el panel superior derecho Z-Logic:
    #   superior: dZ/dt
    #   media: ECG
    #   inferior: fonocardiograma
    # Corrección: el fono NO es la banda media; es la banda más inferior, por debajo del ECG.
    # dZ/dt: intento principal + rescate adaptativo ampliando banda. Nunca debe bloquear
    # el modo semiautomático/manual: si falla, se entrega una señal basal de baja confianza.
    try:
        df_dz = _digitalizar_banda_trazo(
            arr, 0.02, 0.38, rr, umbral_oscuridad=170, suavizado=suavizado,
            invertir=invertir_dzdt, max_salto_px=max(8, int(h*0.09)),
            amplitud_mm_por_px=mm_por_px, escala_y_por_mm=escala_dzdt,
            preferir_trazo_azul=True, excluir_relleno_amarillo=True, nombre_canal="dzdt"
        )
    except Exception:
        df_dz = _digitalizar_trazo_rescate_no_bloqueante(
            roi_rgb=arr, y_ini_frac=0.00, y_fin_frac=0.42, duracion_ms=rr,
            nombre_canal="dzdt", suavizado=max(1, min(int(suavizado), 5)),
            invertir=invertir_dzdt, amplitud_mm_por_px=mm_por_px,
            escala_y_por_mm=escala_dzdt,
        )
    df_ecg = _digitalizar_banda_trazo(
        arr, 0.41, 0.68, rr, umbral_oscuridad=155, suavizado=max(3, suavizado),
        invertir=False, max_salto_px=max(6, int(h*0.06)),
        preferir_trazo_azul=True, excluir_relleno_amarillo=True, nombre_canal="ecg"
    )
    try:
        df_fono = _digitalizar_banda_trazo(
            arr, 0.735, 0.930, rr, umbral_oscuridad=175, suavizado=1,
            invertir=False, max_salto_px=max(3, int(h*0.045)),
            preferir_trazo_azul=False, excluir_relleno_amarillo=True, nombre_canal="fono_inferior"
        )
    except Exception:
        df_fono = _digitalizar_fono_inferior_rescate(
            roi_rgb=arr, y_ini_frac=0.735, y_fin_frac=0.930,
            duracion_ms=rr, suavizado=1, nombre_canal="fono_inferior"
        )

    # Adaptar dZ/dt al formato del módulo existente.
    df_curva = pd.DataFrame({
        "tiempo_ms": df_dz["tiempo_ms"],
        "amplitud_relativa": df_dz["amplitud"],
    })

    return {
        "imagen_pagina": img,
        "panel": panel,
        "rr_ms": rr,
        "grid_px": grid_px,
        "mm_por_px": mm_por_px,
        "df_curva": df_curva,
        "df_dzdt": df_dz,
        "df_ecg": df_ecg,
        "df_fono": df_fono,
        "crop_rel": crop_rel,
    }



def detectar_ruidos_fonocardiograma(
    df_fono: pd.DataFrame,
    puntos: Optional[Dict[str, Any]] = None,
    rr_ms: Optional[float] = None,
) -> Dict[str, Any]:
    """
    Detecta eventos principales del fonocardiograma digitalizado.

    Uso clínico/visual solicitado:
    - No se dibuja la curva completa del fono.
    - Se colocan barras verticales en:
        1) inicio del primer ruido cardíaco (S1_inicio)
        2) fin del primer ruido cardíaco (S1_fin)
        3) inicio del segundo ruido cardíaco (S2_inicio)

    La detección es semiautomática: si el fono no es confiable, se usan referencias
    fisiológicas aproximadas basadas en QRS/B/X para no bloquear el informe.
    """
    out = {
        "s1_inicio_ms": None,
        "s1_fin_ms": None,
        "s2_inicio_ms": None,
        "metodo": "no_disponible",
        "confianza": "Baja",
        "alertas": [],
    }
    try:
        if df_fono is None or df_fono.empty or "tiempo_ms" not in df_fono.columns:
            raise ValueError("df_fono vacío")
        amp_col = "amplitud" if "amplitud" in df_fono.columns else df_fono.columns[-1]
        t = df_fono["tiempo_ms"].to_numpy(dtype=float)
        y = df_fono[amp_col].to_numpy(dtype=float)
        if len(t) < 20:
            raise ValueError("fono con pocos puntos")
        rr = float(rr_ms or np.nanmax(t) or 800.0)
        y0 = y - np.nanmedian(y)
        # Si el canal vino como línea basal de rescate, no forzar detección falsa.
        if not np.isfinite(np.nanstd(y0)) or np.nanstd(y0) < 1e-6:
            raise ValueError("fono plano/no confiable")

        env = np.abs(y0)
        # Suavizado mínimo de envolvente para agrupar vibraciones S1/S2 sin borrar picos.
        win = max(3, int(round(len(env) * 0.008)))
        if win % 2 == 0:
            win += 1
        if win > 3:
            kernel = np.ones(win) / win
            env_s = np.convolve(np.pad(env, (win//2, win//2), mode="edge"), kernel, mode="valid")
        else:
            env_s = env

        med = float(np.nanmedian(env_s))
        mad = float(np.nanmedian(np.abs(env_s - med))) + 1e-9
        thr = max(float(np.nanpercentile(env_s, 88)), med + 3.0 * mad)
        mask = env_s >= thr

        # Convertir máscara a clusters temporales.
        clusters = []
        in_cl = False
        start = 0
        for i, val in enumerate(mask):
            if val and not in_cl:
                in_cl = True
                start = i
            if in_cl and ((not val) or i == len(mask) - 1):
                end = i if val and i == len(mask) - 1 else i - 1
                if end > start:
                    # Expandir un poco para abarcar inicio/fin real del ruido.
                    pad = max(1, int(round(0.006 * len(mask))))
                    a = max(0, start - pad)
                    b = min(len(mask) - 1, end + pad)
                    dur = float(t[b] - t[a])
                    if 4 <= dur <= 160:
                        seg = env_s[a:b+1]
                        peak_i = a + int(np.nanargmax(seg))
                        clusters.append({
                            "inicio": float(t[a]),
                            "fin": float(t[b]),
                            "pico": float(t[peak_i]),
                            "energia": float(np.nansum(seg)),
                            "duracion": dur,
                        })
                in_cl = False

        # Referencias fisiológicas desde cursores.
        q_ms = float(puntos.get("q_ms")) if puntos and puntos.get("q_ms") is not None else 0.0
        b_ms = float(puntos.get("b_ms")) if puntos and puntos.get("b_ms") is not None else min(90.0, rr * 0.12)
        x_ms = float(puntos.get("x_ms")) if puntos and puntos.get("x_ms") is not None else min(rr - 80.0, 360.0)

        def elegir_cluster(ini, fin, ref):
            cand = [c for c in clusters if ini <= c["pico"] <= fin]
            if not cand:
                return None
            # Priorizar cercanía a referencia y energía.
            return sorted(cand, key=lambda c: (abs(c["pico"] - ref), -c["energia"]))[0]

        s1 = elegir_cluster(max(0, q_ms - 45), min(rr, b_ms + 100), b_ms)
        if s1 is None:
            s1 = elegir_cluster(0, min(rr, 180), b_ms)
        s2 = elegir_cluster(max(0, x_ms - 110), min(rr, x_ms + 140), x_ms)
        if s2 is None:
            s2 = elegir_cluster(max(0, b_ms + 160), min(rr, rr - 20), x_ms)

        if s1 is not None:
            out["s1_inicio_ms"] = s1["inicio"]
            out["s1_fin_ms"] = s1["fin"]
        if s2 is not None:
            out["s2_inicio_ms"] = s2["inicio"]

        if s1 is not None and s2 is not None:
            out["metodo"] = "envolvente_fonocardiograma"
            out["confianza"] = "Intermedia/Alta"
        elif s1 is not None or s2 is not None:
            out["metodo"] = "envolvente_parcial"
            out["confianza"] = "Intermedia"
            out["alertas"].append("Detección parcial de ruidos cardíacos; revisar visualmente.")
        else:
            raise ValueError("no se encontraron clusters de fono")
        return out
    except Exception as exc:
        # Fallback fisiológico no bloqueante: permite ubicar las barras como guía.
        if puntos:
            q_ms = float(puntos.get("q_ms") or 0.0)
            b_ms = float(puntos.get("b_ms") or q_ms + 70.0)
            x_ms = float(puntos.get("x_ms") or b_ms + 280.0)
            out.update({
                "s1_inicio_ms": q_ms,
                "s1_fin_ms": b_ms,
                "s2_inicio_ms": x_ms,
                "metodo": "referencia_QRS_B_X_por_fono_no_confiable",
                "confianza": "Baja/guía visual",
                "alertas": [f"Fono no confiable para detección automática ({exc}). Barras ubicadas por QRS-B-X."],
            })
        else:
            out["alertas"] = [f"No se pudo detectar fono ni usar cursores: {exc}"]
        return out


def graficar_canales_zlogic(
    canales: Dict[str, Any],
    puntos: Optional[Dict[str, Any]] = None,
    amplificar_fono: float = 4.00,
    amplificar_ecg: float = 1.80,
    expandir_vertical: bool = True,
) -> plt.Figure:
    """
    Grafica dZ/dt, ECG y fonocardiograma con grilla vertical cada 20 ms.

    Convención visual solicitada:
    - dZ/dt: azul
    - ECG: verde
    - Fonocardiograma: naranja

    La figura se expande verticalmente para separar los canales y aumentar la
    visibilidad del fonocardiograma, que suele tener menor amplitud aparente.
    """
    df_dz = canales["df_dzdt"]
    df_ecg = canales["df_ecg"]
    df_fono = canales["df_fono"]
    rr = float(canales.get("rr_ms") or df_dz["tiempo_ms"].max())

    def norm(y, escala=1.0):
        y = np.asarray(y, dtype=float)
        y = y - np.nanmedian(y)
        m = np.nanmax(np.abs(y))
        return (y / m * escala) if m and m > 0 else y

    # Canales separados verticalmente para evitar superposición.
    # Se amplifica fono en forma controlada para que no quede achatado.
    t = df_dz["tiempo_ms"].to_numpy(float)
    base_dz = 3.10 if expandir_vertical else 2.20
    base_ecg = 1.05 if expandir_vertical else 0.70
    base_fono = -1.65 if expandir_vertical else -0.70

    y_dz = norm(df_dz["amplitud"].to_numpy(float), 0.95) + base_dz
    y_ecg = norm(df_ecg["amplitud"].to_numpy(float), 0.75 * float(amplificar_ecg)) + base_ecg
    # El fonocardiograma NO se dibuja como curva completa. Se usa solo para
    # detectar/mostrar barras verticales de S1 inicio, S1 fin y S2 inicio.
    eventos_fono = detectar_ruidos_fonocardiograma(df_fono, puntos=puntos, rr_ms=rr)
    canales["eventos_fono"] = eventos_fono

    fig_h = 7.2 if expandir_vertical else 5.2
    fig, ax = plt.subplots(figsize=(13.5, fig_h))

    ax.plot(t, y_dz, linewidth=2.1, color="#1f77b4", label="dZ/dt digitalizada")
    ax.plot(df_ecg["tiempo_ms"], y_ecg, linewidth=1.6, color="#2ca02c", label="ECG digitalizado")
    ax.plot([], [], color="#ff7f0e", linewidth=2.4, label="Fono: barras S1/S2")

    # Grillado cada 20 ms.
    ticks20 = np.arange(0, rr + 20, 20)
    for x in ticks20:
        ax.axvline(x, color="#D6DEE8", linewidth=0.55, alpha=0.70, zorder=0)
    for x in np.arange(0, rr + 100, 100):
        ax.axvline(x, color="#8EA8C3", linewidth=1.05, alpha=0.90, zorder=0)

    # Líneas basales por canal.
    ax.axhline(base_dz, color="#9CA3AF", linewidth=0.8, alpha=0.55)
    ax.axhline(base_ecg, color="#9CA3AF", linewidth=0.8, alpha=0.55)
    ax.axhline(base_fono, color="#9CA3AF", linewidth=0.8, alpha=0.55)

    ax.text(-0.025 * rr, base_dz, "dZ/dt", ha="right", va="center", fontsize=10, fontweight="bold", color="#1f77b4")
    ax.text(-0.025 * rr, base_ecg, "ECG", ha="right", va="center", fontsize=10, fontweight="bold", color="#2ca02c")
    ax.text(-0.025 * rr, base_fono, "Fono", ha="right", va="center", fontsize=10, fontweight="bold", color="#ff7f0e")

    # Barras verticales solicitadas para fonocardiograma.
    barra_y0 = base_fono - 0.62
    barra_y1 = base_fono + 0.62
    eventos_fono = canales.get("eventos_fono") or {}
    marcas_fono = [
        ("S1 inicio", eventos_fono.get("s1_inicio_ms")),
        ("S1 fin", eventos_fono.get("s1_fin_ms")),
        ("S2 inicio", eventos_fono.get("s2_inicio_ms")),
    ]
    for etiqueta, xev in marcas_fono:
        if xev is None:
            continue
        try:
            xev = float(xev)
        except Exception:
            continue
        if 0 <= xev <= rr:
            ax.vlines(xev, barra_y0, barra_y1, color="#ff7f0e", linewidth=2.7, alpha=0.95, zorder=4)
            ax.annotate(etiqueta, xy=(xev, barra_y1), xytext=(4, 5), textcoords="offset points", fontsize=9, color="#ff7f0e", fontweight="bold", rotation=90, va="bottom")
    if eventos_fono.get("alertas"):
        ax.text(0.01 * rr, base_fono - 1.02, "Fono: " + str(eventos_fono.get("confianza")), color="#ff7f0e", fontsize=9, ha="left", va="center")

    if puntos is not None:
        dfc = canales["df_curva"]
        yd = dfc["amplitud_relativa"].to_numpy(float)
        yd_norm = norm(yd, 0.95) + base_dz
        tt = dfc["tiempo_ms"].to_numpy(float)
        for lab, key in [("QRS", "idx_q"), ("B", "idx_b"), ("C", "idx_c"), ("X", "idx_x")]:
            if key in puntos:
                idx = int(puntos[key])
                idx = max(0, min(idx, len(tt) - 1))
                ax.scatter(tt[idx], yd_norm[idx], s=78, zorder=5)
                ax.annotate(lab, (tt[idx], yd_norm[idx]), xytext=(5, 9), textcoords="offset points", fontsize=11, fontweight="bold")
                ax.axvline(tt[idx], linestyle="--", linewidth=1.15, alpha=0.75)

    ax.set_xlim(0, rr)
    if expandir_vertical:
        ax.set_ylim(base_fono - 1.70, base_dz + 1.45)
    ax.set_yticks([])
    ax.set_xlabel("Tiempo (ms) - grillado vertical cada 20 ms")
    ax.set_title("Digitalización automática: dZ/dt + ECG verde + barras fonocardiograma S1/S2")
    ax.grid(axis="x", alpha=0.15)
    ax.legend(loc="upper right")
    fig.tight_layout()
    return fig

def actualizar_aprendizaje_cursores(pauto: Dict[str, Any], pfinal: Dict[str, Any], validado: bool) -> Dict[str, Any]:
    """
    Registra aprendizaje semiautomático por validación del operador.

    La app no aprende de forma autónoma ni cambia cursores sin autorización;
    documenta cuántas veces el operador corrigió/confirmó QRS-B-C-X y cuál fue
    la diferencia media entre la propuesta automática y la posición final.
    """
    estado = st.session_state.get(
        "aprendizaje_cursores",
        {
            "n": 0,
            "error_ms_acum": 0.0,
            "error_ms_prom": None,
            "historial_error_ms": [],
            "historial_q_ms": [],
            "historial_b_ms": [],
            "historial_c_ms": [],
            "historial_x_ms": [],
        },
    )
    if validado:
        diffs = []
        diffs_por_cursor = {}
        for etiqueta, k in [("q", "q_ms"), ("b", "b_ms"), ("c", "c_ms"), ("x", "x_ms")]:
            if k in pauto and k in pfinal:
                d = abs(float(pfinal[k]) - float(pauto[k]))
                diffs.append(d)
                diffs_por_cursor[etiqueta] = d

        err = float(np.mean(diffs)) if diffs else 0.0
        estado["n"] = int(estado.get("n", 0)) + 1
        estado["error_ms_acum"] = float(estado.get("error_ms_acum", 0.0)) + err
        estado["error_ms_prom"] = estado["error_ms_acum"] / max(1, estado["n"])

        estado.setdefault("historial_error_ms", []).append(err)
        estado.setdefault("historial_q_ms", []).append(float(diffs_por_cursor.get("q", np.nan)))
        estado.setdefault("historial_b_ms", []).append(float(diffs_por_cursor.get("b", np.nan)))
        estado.setdefault("historial_c_ms", []).append(float(diffs_por_cursor.get("c", np.nan)))
        estado.setdefault("historial_x_ms", []).append(float(diffs_por_cursor.get("x", np.nan)))
        estado["ultimo_error_ms"] = err
        estado["ultimo_error_por_cursor"] = diffs_por_cursor
        st.session_state["aprendizaje_cursores"] = estado
    return estado


def graficar_aprendizaje_cursores(estado: Dict[str, Any]) -> Optional[plt.Figure]:
    """Grafico simple del avance de aprendizaje operativo por validaciones."""
    hist = estado.get("historial_error_ms") or []
    if len(hist) == 0:
        return None

    x = np.arange(1, len(hist) + 1)
    y = np.asarray(hist, dtype=float)
    # Media movil corta para ver tendencia sin requerir librerias externas.
    if len(y) >= 3:
        kernel = np.ones(min(5, len(y))) / min(5, len(y))
        ypad = np.pad(y, (len(kernel)//2, len(kernel)//2), mode="edge")
        trend = np.convolve(ypad, kernel, mode="valid")[:len(y)]
    else:
        trend = y

    fig, ax = plt.subplots(figsize=(9.5, 2.8))
    ax.plot(x, y, marker="o", linewidth=1.3, label="Corrección media por validación")
    ax.plot(x, trend, linewidth=2.2, label="Tendencia")
    ax.axhline(10, linestyle="--", linewidth=1.0, alpha=0.60, label="Meta <10 ms")
    ax.set_xlabel("Validación del operador")
    ax.set_ylabel("Corrección media (ms)")
    ax.set_title("Avance de aprendizaje semiautomático de cursores QRS-B-C-X")
    ax.grid(True, alpha=0.25)
    ax.legend(loc="upper right", fontsize=8)
    fig.tight_layout()
    return fig


def mostrar_barra_aprendizaje() -> None:
    estado = st.session_state.get("aprendizaje_cursores", {"n": 0, "error_ms_prom": None, "historial_error_ms": []})
    n = int(estado.get("n", 0))
    progreso = min(1.0, n / 30.0)
    st.progress(progreso, text=f"Aprendizaje semiautomático de cursores: {n}/30 validaciones del operador")

    c1, c2, c3 = st.columns(3)
    c1.metric("Validaciones", n)
    if estado.get("error_ms_prom") is not None:
        c2.metric("Corrección media histórica", f"{estado['error_ms_prom']:.1f} ms")
    else:
        c2.metric("Corrección media histórica", "Sin datos")
    if estado.get("ultimo_error_ms") is not None:
        c3.metric("Última corrección", f"{estado['ultimo_error_ms']:.1f} ms")
    else:
        c3.metric("Última corrección", "Pendiente")

    fig_ap = graficar_aprendizaje_cursores(estado)
    if fig_ap is not None:
        st.pyplot(fig_ap, use_container_width=True)
    st.caption(
        "La barra y el gráfico registran el aprendizaje operativo basado en correcciones manuales o semiautomáticas. "
        "La app no modifica cursores automáticamente sin confirmación visual del médico."
    )

def ciclo_ms_desde_fc(fc: Optional[float], default: float = 800.0) -> float:
    if fc is None or fc <= 0:
        return default
    return float(60000.0 / fc)


def suavizar_senal(y: np.ndarray, ventana: int = 9) -> np.ndarray:
    y = np.asarray(y, dtype=float)
    if len(y) < 5:
        return y
    ventana = int(max(3, ventana))
    if ventana % 2 == 0:
        ventana += 1
    ventana = min(ventana, len(y) - 1 if (len(y) - 1) % 2 else len(y) - 2)
    if ventana < 3:
        return y
    kernel = np.ones(ventana) / ventana
    ypad = np.pad(y, (ventana // 2, ventana // 2), mode="edge")
    return np.convolve(ypad, kernel, mode="valid")


def _limpiar_mascara_grafica(mask: np.ndarray) -> np.ndarray:
    """
    Elimina líneas de grilla/ejes cuando ocupan gran parte de una fila o columna.
    Mantiene trazos curvos focales. Implementado sin OpenCV para Streamlit Cloud.
    """
    mask = np.asarray(mask, dtype=bool).copy()
    if mask.size == 0:
        return mask
    h, w = mask.shape
    # Grilla horizontal/vertical: muchas columnas/filas activas.
    row_density = mask.mean(axis=1)
    col_density = mask.mean(axis=0)
    mask[row_density > 0.45, :] = False
    mask[:, col_density > 0.55] = False
    # Quitar bordes del recorte, que suelen ser ejes/cuadros.
    borde_y = max(1, int(0.01 * h))
    borde_x = max(1, int(0.01 * w))
    mask[:borde_y, :] = False
    mask[-borde_y:, :] = False
    mask[:, :borde_x] = False
    mask[:, -borde_x:] = False
    return mask


def _seguir_trazo_por_continuidad(mask: np.ndarray, max_salto_px: Optional[int] = None) -> Tuple[np.ndarray, np.ndarray]:
    """
    Sigue una curva por continuidad columna a columna.
    Evita usar la mediana global si existen varias señales o texto dentro del recorte.
    Devuelve xs, ys en coordenadas de la ROI.
    """
    h, w = mask.shape
    if max_salto_px is None:
        max_salto_px = max(6, int(0.10 * h))

    columnas = []
    candidatos_col = []
    for x in range(w):
        ys = np.where(mask[:, x])[0]
        if len(ys) > 0:
            columnas.append(x)
            candidatos_col.append(ys.astype(float))

    if len(columnas) < 20:
        raise ValueError("No se detectaron suficientes columnas con curva. Ajustar recorte, umbral o activar deteccion de color.")

    xs_final: List[float] = []
    ys_final: List[float] = []

    # Semilla: primera zona con suficientes candidatos, usando mediana local.
    prev_y: Optional[float] = None
    for x, ys in zip(columnas, candidatos_col):
        if len(ys) > 0:
            prev_y = float(np.median(ys))
            xs_final.append(float(x))
            ys_final.append(prev_y)
            break

    if prev_y is None:
        raise ValueError("No se pudo iniciar el seguimiento de la curva.")

    started = False
    for x, ys in zip(columnas, candidatos_col):
        if not started:
            if int(x) == int(xs_final[0]):
                started = True
            continue
        # Elegir el píxel más cercano a la trayectoria previa.
        dist = np.abs(ys - prev_y)
        j = int(np.argmin(dist))
        y_sel = float(ys[j])
        # Si el salto es excesivo, se ignora esa columna para interpolar luego.
        if abs(y_sel - prev_y) <= max_salto_px:
            xs_final.append(float(x))
            ys_final.append(y_sel)
            prev_y = 0.80 * prev_y + 0.20 * y_sel

    if len(xs_final) < 20:
        # Fallback: mediana por columna si el seguimiento fue demasiado estricto.
        xs_final = [float(x) for x in columnas]
        ys_final = [float(np.median(ys)) for ys in candidatos_col]

    return np.asarray(xs_final, dtype=float), np.asarray(ys_final, dtype=float)


def digitalizar_curva_icg_desde_imagen(
    imagen_pil: Image.Image,
    crop_x1: int,
    crop_x2: int,
    crop_y1: int,
    crop_y2: int,
    duracion_ms: float,
    umbral_oscuridad: int = 120,
    invertir_polaridad: bool = False,
    suavizado: int = 9,
    detectar_color: bool = True,
    max_salto_px: Optional[int] = None,
) -> pd.DataFrame:
    """
    Digitalizacion corregida para señal ICG promedio.

    Mejoras respecto de la version previa:
    - Limpieza de grilla/ejes por densidad de filas/columnas.
    - Deteccion de trazos oscuros y, opcionalmente, trazos coloreados.
    - Seguimiento de una curva por continuidad, en vez de mediana global por columna.
    - Interpolacion de huecos y suavizado leve.

    Requisito clinico-operativo: el recorte debe contener solo la curva dZ/dt.
    """
    if crop_x2 <= crop_x1 or crop_y2 <= crop_y1:
        raise ValueError("Recorte no valido.")

    img = imagen_pil.convert("RGB")
    arr = np.asarray(img)
    h, w = arr.shape[:2]
    crop_x1 = max(0, min(int(crop_x1), w - 1))
    crop_x2 = max(1, min(int(crop_x2), w))
    crop_y1 = max(0, min(int(crop_y1), h - 1))
    crop_y2 = max(1, min(int(crop_y2), h))
    roi = arr[crop_y1:crop_y2, crop_x1:crop_x2, :]
    if roi.size == 0:
        raise ValueError("La region recortada esta vacia.")

    roi_f = roi.astype(float)
    gris = np.mean(roi_f, axis=2)
    r, g, b = roi_f[:, :, 0], roi_f[:, :, 1], roi_f[:, :, 2]
    maxc = np.max(roi_f, axis=2)
    minc = np.min(roi_f, axis=2)
    saturacion = maxc - minc

    # Trazos oscuros: negro/gris/azul oscuro.
    mask_oscura = gris < int(umbral_oscuridad)

    # Trazos coloreados: útil si el informe exporta la curva en azul/verde/rojo.
    # Evita grilla clara: requiere saturacion y brillo no blanco.
    mask_color = (saturacion > 35) & (gris < 245)

    mask = mask_oscura | (mask_color if detectar_color else False)
    mask = _limpiar_mascara_grafica(mask)

    # Si el filtrado removio demasiado, usar máscara menos agresiva, pero siempre limpiando grilla.
    if mask.sum() < 20:
        mask = _limpiar_mascara_grafica(mask_oscura)

    xs, ys = _seguir_trazo_por_continuidad(mask, max_salto_px=max_salto_px)

    if len(xs) < 20:
        raise ValueError("No se detectaron suficientes puntos de curva luego de limpiar la imagen.")

    xfull = np.arange(mask.shape[1], dtype=float)
    yinterp = np.interp(xfull, xs, ys)

    # En imagen, y crece hacia abajo; se invierte para que el ascenso sistolico sea positivo.
    amp = -yinterp
    if invertir_polaridad:
        amp = -amp

    amp = amp - np.nanmedian(amp)
    amp = suavizar_senal(amp, suavizado)

    # Recentrar y escalar solo en amplitud relativa para facilitar lectura.
    if np.nanstd(amp) > 0:
        # Mantener escala visual en pixeles relativos, no z-score, para no alterar morfologia.
        amp = amp - np.nanmedian(amp)

    t = np.linspace(0, float(duracion_ms), len(amp))
    return pd.DataFrame({"tiempo_ms": t, "amplitud_relativa": amp})

def idx_cercano(t: np.ndarray, ms: float) -> int:
    return int(np.argmin(np.abs(np.asarray(t) - float(ms))))



def ajustar_puntos_con_tiempos_importados(
    df_curva: pd.DataFrame,
    puntos: Dict[str, Any],
    datos: Optional[Dict[str, Any]],
) -> Dict[str, Any]:
    """
    Inicializa/ajusta QRS-B-X usando los tiempos del informe Z-Logic cuando existen.

    En Z-Logic: PPE = PEP = QRS-B y PE = LVET = B-X. Si la digitalización
    automática detecta cursores demasiado cortos o desplazados, estos tiempos del
    informe se usan como ancla inicial semiautomática. El operador sigue pudiendo
    corregir manualmente los cursores con sliders.
    """
    if df_curva is None or df_curva.empty or not datos:
        return puntos
    pep = limpiar_numero(datos.get("pep"))
    lvet = limpiar_numero(datos.get("lvet"))
    if pep is None or lvet is None or not (40 <= pep <= 180) or not (150 <= lvet <= 500):
        return puntos
    try:
        t = df_curva["tiempo_ms"].to_numpy(dtype=float)
        y = df_curva["amplitud_relativa"].to_numpy(dtype=float)
        tmin, tmax = float(np.nanmin(t)), float(np.nanmax(t))
        q_ms = float(puntos.get("q_ms", tmin))
        # Si no entra el intervalo completo, mover QRS hacia el inicio útil.
        if q_ms + pep + lvet > tmax:
            q_ms = max(tmin, tmax - pep - lvet - 5)
        b_ms = q_ms + pep
        x_ms = b_ms + lvet
        i_q = indice_mas_cercano(t, q_ms)
        i_b = indice_mas_cercano(t, b_ms)
        i_x = indice_mas_cercano(t, x_ms)
        if not (0 <= i_q < i_b < i_x < len(t)):
            return puntos
        # C = máximo de dZ/dt dentro de la ventana B-X; si no, conservar el previo.
        seg = np.arange(i_b, max(i_b + 1, i_x + 1))
        if len(seg) > 3:
            i_c = int(seg[np.nanargmax(y[seg])])
        else:
            i_c = int(puntos.get("idx_c", i_b + 1))
        if not (i_b < i_c < i_x):
            i_c = int(puntos.get("idx_c", i_b + max(1, (i_x - i_b)//3)))
            i_c = max(i_b + 1, min(i_c, i_x - 1))
        out = dict(puntos)
        out.update({
            "idx_q": int(i_q), "idx_b": int(i_b), "idx_c": int(i_c), "idx_x": int(i_x),
            "q_ms": float(t[i_q]), "b_ms": float(t[i_b]), "c_ms": float(t[i_c]), "x_ms": float(t[i_x]),
            "pep_ms": float(t[i_b] - t[i_q]),
            "lvet_ms": float(t[i_x] - t[i_b]),
            "bc_ms": float(t[i_c] - t[i_b]),
            "cx_ms": float(t[i_x] - t[i_c]),
            "tiempo_pico_c_ms": float(t[i_c] - t[i_q]),
            "anclado_a_tiempos_informe": True,
            "fuente_anclaje": "PPE/PE importados del informe Z-Logic",
        })
        alertas = []
        if not (out["q_ms"] <= out["b_ms"] < out["c_ms"] < out["x_ms"]):
            alertas.append("Orden QRS-B-C-X no fisiológico.")
        if not (40 <= out["pep_ms"] <= 180):
            alertas.append("PEP fuera de rango orientativo.")
        if not (150 <= out["lvet_ms"] <= 500):
            alertas.append("LVET fuera de rango orientativo.")
        out["alertas"] = alertas
        out["confianza"] = "Anclada a tiempos del informe; confirmar visualmente" if not alertas else "Intermedia/Baja"
        return out
    except Exception:
        return puntos

def detectar_puntos_qbcx_icg(df_curva: pd.DataFrame, fc: Optional[float] = None) -> Dict[str, Any]:
    """
    Deteccion corregida con ventanas fisiologicas.

    - C: pico sistolico principal preferentemente entre 40-250 ms.
    - B: inicio de ascenso antes de C, dentro de una ventana fisiologica previa.
    - X: nadir post-C, pero restringido a LVET plausible desde B.

    Importante: el resultado es preliminar y debe validarse visualmente con cursores.
    """
    if df_curva is None or df_curva.empty or len(df_curva) < 30:
        raise ValueError("Curva insuficiente para detectar cursores.")

    t = df_curva["tiempo_ms"].to_numpy(float)
    y = df_curva["amplitud_relativa"].to_numpy(float)
    n = len(y)
    dur = float(t[-1] - t[0]) if len(t) > 1 else 800.0

    # Si la señal parece invertida, advertir en el resultado pero no invertir automaticamente.
    polaridad_sospechosa = abs(float(np.nanmin(y))) > 1.30 * abs(float(np.nanmax(y)))

    # C: maximo positivo en ventana sistolica inicial. Si no hay ventana suficiente, usar zona amplia.
    c_ini_ms = 35.0
    c_fin_ms = min(280.0, max(180.0, 0.42 * dur))
    mask_c = (t >= c_ini_ms) & (t <= c_fin_ms)
    if mask_c.sum() >= 8:
        idxs_c = np.where(mask_c)[0]
        ic = int(idxs_c[np.argmax(y[idxs_c])])
    else:
        i0, i1 = max(1, int(.04 * n)), min(n - 2, int(.70 * n))
        ic = int(i0 + np.argmax(y[i0:i1]))

    dy = np.gradient(y, t)
    iq = 0

    # B: buscar antes de C entre 40 ms y C-10 ms, maximo 180 ms antes de C.
    b_ini_ms = max(35.0, float(t[ic]) - 180.0)
    b_fin_ms = max(b_ini_ms + 5.0, float(t[ic]) - 8.0)
    mask_b = (t >= b_ini_ms) & (t <= b_fin_ms)
    idxs_b = np.where(mask_b)[0]

    if len(idxs_b) >= 8:
        ypre = y[idxs_b]
        dypre = dy[idxs_b]
        base = np.percentile(ypre, 20)
        pico = y[ic]
        amp = max(1e-6, pico - base)
        # inicio del ascenso sistolico: primer punto que supera 10-12% de la amplitud C-base.
        candidatos = idxs_b[ypre > base + 0.10 * amp]
        if len(candidatos) > 0:
            ib_umbral = int(candidatos[0])
        else:
            ib_umbral = int(idxs_b[np.argmax(dypre)])
        ib_pend = int(idxs_b[np.argmax(dypre)])
        # tomar el mas precoz de ambos para no correr B hacia la pendiente maxima.
        ib = min(ib_umbral, ib_pend)
        ib = max(1, min(ib, ic - 1))
    else:
        ib = max(1, idx_cercano(t, max(50.0, float(t[ic]) - 90.0)))
        ib = min(ib, ic - 1)

    # Restringir B segun reglas operativas, sin forzar si C esta muy precoz.
    b_min = 35.0
    b_max = min(float(t[ic]) - 5.0, max(120.0, 0.30 * dur))
    if t[ib] < b_min:
        ib = idx_cercano(t, b_min)
    if t[ib] > b_max:
        ib = idx_cercano(t, b_max)
    ib = max(1, min(ib, ic - 1))

    # X: no usar minimo global tardio. Buscar en ventana de LVET fisiologico desde B.
    # A FC alta, acortar ventana; a FC baja, permitir algo mas larga.
    ciclo_ms = ciclo_ms_desde_fc(fc, dur)
    lvet_min = 180.0
    lvet_max = min(450.0, max(260.0, 0.58 * ciclo_ms))
    x_ini_ms = float(t[ib]) + lvet_min
    x_fin_ms = min(float(t[ib]) + lvet_max, dur)
    mask_x = (t >= x_ini_ms) & (t <= x_fin_ms)
    idxs_x = np.where(mask_x)[0]

    if len(idxs_x) >= 8:
        # Nadir dentro de ventana fisiologica. Se evita escoger valles tardios fuera de LVET.
        ix = int(idxs_x[np.argmin(y[idxs_x])])
    else:
        # Fallback: ventana post-C corta.
        post_ini = min(n - 2, ic + max(2, int(.05 * n)))
        post_fin = min(n - 1, ic + max(8, int(.35 * n)))
        ix = post_ini + int(np.argmin(y[post_ini:post_fin])) if post_fin > post_ini else min(n - 1, ic + int(.25 * n))

    res = armar_resultado_puntos(df_curva, iq, ib, ic, ix)
    if polaridad_sospechosa:
        res.setdefault("alertas", []).append("La polaridad parece invertida o el valle domina al pico C; probar 'Invertir polaridad' y redigitalizar.")
        res["confianza"] = "Intermedia/Baja"
    res["deteccion_preliminar"] = True
    res["validado_operador"] = False
    return res

def armar_resultado_puntos(df_curva: pd.DataFrame, iq: int, ib: int, ic: int, ix: int) -> Dict[str, Any]:
    t = df_curva["tiempo_ms"].to_numpy(float)
    n = len(t)
    iq = max(0, min(iq, n-1)); ib = max(0, min(ib, n-1)); ic = max(0, min(ic, n-1)); ix = max(0, min(ix, n-1))
    pep = float(t[ib] - t[iq])
    lvet = float(t[ix] - t[ib])
    alertas: List[str] = []
    if not (iq < ib < ic < ix):
        alertas.append("Orden temporal QRS-B-C-X no fisiologico.")
    if not (50 <= pep <= 180):
        alertas.append("PEP/QRS-B fuera del rango operativo esperado.")
    if not (180 <= lvet <= 450):
        alertas.append("LVET B-X fuera del rango fisiologico orientativo; revisar cursores.")
    # rango CMV mas estricto para FC 60-80 se informa como alerta leve
    if 60 <= pep <= 100:
        b_estado = "B dentro del rango operativo 60-100 ms post-QRS."
    else:
        b_estado = "B fuera del rango 60-100 ms post-QRS; validar visualmente."
    if 280 <= lvet <= 340:
        x_estado = "LVET dentro del rango orientativo 280-340 ms a FC 60-80 lpm."
    else:
        x_estado = "LVET fuera del rango 280-340 ms a FC 60-80 lpm; interpretar segun FC y clinica."
    if len(alertas) >= 2:
        confianza = "Baja"
    elif len(alertas) == 1:
        confianza = "Intermedia"
    else:
        confianza = "Alta / corregida por operador"
    return {
        "idx_q": iq, "idx_b": ib, "idx_c": ic, "idx_x": ix,
        "q_ms": float(t[iq]), "b_ms": float(t[ib]), "c_ms": float(t[ic]), "x_ms": float(t[ix]),
        "pep_ms": pep, "lvet_ms": lvet,
        "bc_ms": float(t[ic] - t[ib]), "cx_ms": float(t[ix] - t[ic]),
        "tiempo_pico_c_ms": float(t[ic] - t[iq]),
        "b_estado": b_estado, "x_estado": x_estado,
        "confianza": confianza, "alertas": alertas,
    }


def recalcular_puntos_manuales(df_curva: pd.DataFrame, q_ms: float, b_ms: float, c_ms: float, x_ms: float) -> Dict[str, Any]:
    t = df_curva["tiempo_ms"].to_numpy(float)
    return armar_resultado_puntos(df_curva, idx_cercano(t, q_ms), idx_cercano(t, b_ms), idx_cercano(t, c_ms), idx_cercano(t, x_ms))


def tabla_resultados_qbcx(p: Dict[str, Any]) -> pd.DataFrame:
    filas = [
        ["QRS/Q", f"{p['q_ms']:.1f} ms", "Referencia electrica inicial"],
        ["Punto B", f"{p['b_ms']:.1f} ms", "Inflexion ascendente inicial de dZ/dt / apertura aortica"],
        ["Punto C", f"{p['c_ms']:.1f} ms", "Pico maximo de dZ/dt"],
        ["Punto X", f"{p['x_ms']:.1f} ms", "Nadir post-sistolico / cierre aortico aproximado"],
        ["PEP aproximado QRS-B", f"{p['pep_ms']:.1f} ms", p.get("b_estado", "")],
        ["LVET B-X", f"{p['lvet_ms']:.1f} ms", p.get("x_estado", "")],
        ["B-C", f"{p['bc_ms']:.1f} ms", "Ascenso mecanico hasta el pico C"],
        ["C-X", f"{p['cx_ms']:.1f} ms", "Descenso post-pico hasta cierre aortico"],
        ["Tiempo al pico C", f"{p['tiempo_pico_c_ms']:.1f} ms", "QRS-C aproximado"],
        ["Confianza", p.get("confianza", "No disponible"), "Calidad operativa de la deteccion"],
    ]
    if p.get("alertas"):
        filas.append(["Alertas", " | ".join(p["alertas"]), "Revisar manualmente"])
    return pd.DataFrame(filas, columns=["Metrica", "Valor", "Interpretacion"])


def texto_interpretacion_qbcx(p: Dict[str, Any]) -> str:
    txt = (
        "Se realizo analisis morfologico semiautomatico de la señal promedio de cardiografia de impedancia. "
        f"El cursor B se ubico a {p['b_ms']:.1f} ms post-QRS y el cursor X a {p['x_ms']:.1f} ms. "
        f"El PEP aproximado QRS-B fue {p['pep_ms']:.1f} ms y el LVET B-X fue {p['lvet_ms']:.1f} ms. "
        f"El pico C se identifico a {p['c_ms']:.1f} ms, con tiempo B-C de {p['bc_ms']:.1f} ms y C-X de {p['cx_ms']:.1f} ms. "
        f"Confianza operativa: {p.get('confianza','No disponible')}. "
    )
    if p.get("alertas"):
        txt += "Advertencias: " + "; ".join(p["alertas"]) + ". Se recomienda validacion visual por el operador antes de cerrar el informe."
    else:
        txt += "La relacion temporal QRS-B-C-X fue compatible con una deteccion fisiologica en esta digitalizacion."
    return txt


def graficar_curva_icg_con_puntos(df_curva: pd.DataFrame, p: Dict[str, Any]) -> plt.Figure:
    t = df_curva["tiempo_ms"].to_numpy(float)
    y = df_curva["amplitud_relativa"].to_numpy(float)
    fig, ax = plt.subplots(figsize=(11, 4.6))
    ax.plot(t, y, linewidth=2.2, label="dZ/dt digitalizada")
    for lab, key in [("QRS", "idx_q"), ("B", "idx_b"), ("C", "idx_c"), ("X", "idx_x")]:
        idx = int(p[key])
        ax.scatter(t[idx], y[idx], s=85, zorder=5)
        ax.axvline(t[idx], linestyle="--", alpha=.55, linewidth=1)
        ax.annotate(lab, (t[idx], y[idx]), xytext=(7, 10), textcoords="offset points", fontsize=12, fontweight="bold")
    ax.set_title("Analisis morfologico semiautomatico de señal ICG")
    ax.set_xlabel("Tiempo (ms)")
    ax.set_ylabel("Amplitud relativa dZ/dt")
    ax.grid(True, alpha=.25)
    ax.legend(loc="best")
    fig.tight_layout()
    return fig


def fig_to_png_bytes(fig: plt.Figure) -> bytes:
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=180, bbox_inches="tight")
    buf.seek(0)
    return buf.getvalue()


def excel_metricas_hemodinamicas_bytes(
    df_metricas: Optional[pd.DataFrame] = None,
    df_qbcx: Optional[pd.DataFrame] = None,
    df_curva: Optional[pd.DataFrame] = None,
    df_canales: Optional[pd.DataFrame] = None,
    df_final: Optional[pd.DataFrame] = None,
    aprendizaje: Optional[Dict[str, Any]] = None,
) -> bytes:
    """Exporta a Excel todas las metricas hemodinamicas y tablas soporte."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        if df_metricas is not None and not df_metricas.empty:
            metricas_excel_sin_ecuacion(df_metricas).to_excel(writer, sheet_name="Metricas_hemodinamicas", index=False)
        else:
            pd.DataFrame([{"Estado": "Sin metricas recalculadas disponibles"}]).to_excel(writer, sheet_name="Metricas_hemodinamicas", index=False)
        if df_qbcx is not None and not df_qbcx.empty:
            df_qbcx.to_excel(writer, sheet_name="Cursores_QBCX", index=False)
        if df_curva is not None and not df_curva.empty:
            df_curva.to_excel(writer, sheet_name="Curva_dZdt", index=False)
        if df_canales is not None and not df_canales.empty:
            df_canales.to_excel(writer, sheet_name="Canales_dZdt_ECG_Fono", index=False)
        if df_final is not None and not df_final.empty:
            df_final.to_excel(writer, sheet_name="Datos_ZLogic_importados", index=False)
        if aprendizaje:
            pd.DataFrame([aprendizaje]).to_excel(writer, sheet_name="Aprendizaje_cursores", index=False)
        pd.DataFrame({
            "Contenido": [
                "Metricas hemodinamicas calculadas desde curva de impedancia corregida semiautomaticamente",
                "Formato de decimales calibrado con archivo de ejemplo Z-Logic: FC/PA/RVS/IRV enteros; DS/IDS/VM/IC/ITC/CFT/BSA/BMI con 1 decimal; CA/Ea/Ees/AC con 2 decimales; tiempos en ms sin decimales.",
                "Controles de coherencia cruzan BMI, CO=DS*FC/1000, IC=CO/BSA, IRV=RVS*BSA e IDS=DS/BSA.",
                "Cursores QRS-B-C-X",
                "Curva dZ/dt digitalizada",
                "Canales dZ/dt, ECG y fonocardiograma si fueron extraidos del PDF",
                "Datos estructurados importados del Z-Logic",
                "Estado del aprendizaje por confirmacion/correccion manual de cursores",
            ]
        }).to_excel(writer, sheet_name="README", index=False)
    output.seek(0)
    return output.getvalue()



# ============================================================
# DIAGNOSTICO MORFOLOGICO VALIDADO Y RECOMENDACION TERAPEUTICA
# ============================================================

def _metric_value(df_der: Optional[pd.DataFrame], nombre: str) -> Optional[float]:
    if df_der is None or df_der.empty:
        return None
    try:
        m = df_der[df_der["Variable"].astype(str).str.lower() == nombre.lower()]
        if m.empty:
            return None
        return limpiar_numero(m.iloc[0].get("Valor recalculado"))
    except Exception:
        return None


def _metric_contains(df_der: Optional[pd.DataFrame], texto: str) -> Optional[float]:
    if df_der is None or df_der.empty:
        return None
    try:
        m = df_der[df_der["Variable"].astype(str).str.lower().str.contains(texto.lower(), na=False)]
        if m.empty:
            return None
        return limpiar_numero(m.iloc[0].get("Valor recalculado"))
    except Exception:
        return None


def filtrar_metricas_pdf_morfologicas(df_der: Optional[pd.DataFrame]) -> pd.DataFrame:
    """PDF restringido a variables pedidas para el analisis morfologico validado."""
    if df_der is None or df_der.empty:
        return pd.DataFrame(columns=["Variable", "Valor recalculado", "Unidad", "Ecuacion", "Estado"])
    permitidas = [
        "CI", "CI desde CO importado", "SI/IDS", "SVRI/IRV aproximada",
        "FE Capan", "Indice de velocidad relativo", "Indice de Heather relativo",
        "Indice de aceleracion relativo", "ITC aproximado", "STR/CTS morfologico",
        "Acoplamiento VA Ea/Ees", "CFT/TFC", "PEP morfologico", "LVET morfologico",
    ]
    out_rows = []
    for _, r in df_der.iterrows():
        var = str(r.get("Variable", ""))
        if any(normalizar_txt(var) == normalizar_txt(p) for p in permitidas):
            out_rows.append(r.to_dict())
    if not out_rows:
        return pd.DataFrame(columns=df_der.columns)
    return pd.DataFrame(out_rows).reset_index(drop=True)


def _mv(df_der: Optional[pd.DataFrame], *nombres: str) -> Optional[float]:
    if df_der is None or df_der.empty:
        return None
    for nombre in nombres:
        v = _metric_value(df_der, nombre)
        if v is not None:
            return v
    for nombre in nombres:
        v = _metric_contains(df_der, nombre)
        if v is not None:
            return v
    return None


def clasificar_flujo_ic(ci: Optional[float]) -> str:
    if ci is None:
        return "No clasificable"
    if ci > 4.2:
        return "Hiperdinámico"
    if ci < 2.5:
        return "Hipodinámico"
    return "Normodinámico"


def clasificar_ids(ids: Optional[float]) -> str:
    if ids is None:
        return "No clasificable"
    if ids < 35:
        return "IDS bajo"
    if ids > 65:
        return "IDS alto"
    return "IDS normal"


def clasificar_irv(irv: Optional[float]) -> str:
    if irv is None:
        return "No clasificable"
    if irv > 2580:
        return "Vasoconstricción / poscarga elevada"
    if irv < 1300:
        return "IRV baja / vasodilatación relativa"
    return "IRV normal"


def clasificar_fe_capan(fe: Optional[float]) -> str:
    if fe is None:
        return "No clasificable"
    fe_pct = fe * 100 if fe <= 1.5 else fe
    if fe_pct < 50:
        return "Función cardíaca disminuida"
    return "Función cardíaca conservada"


def clasificar_acoplamiento(ac: Optional[float]) -> str:
    if ac is None:
        return "No clasificable"
    if ac < 1.0:
        return "Óptimo"
    if ac <= 1.3:
        return "Subóptimo"
    return "Desacoplamiento ventrículo-arterial"


def clasificar_volemia_cft(cft: Optional[float]) -> str:
    if cft is None:
        return "No clasificable"
    if cft < 41:
        return "Hipovolemia relativa"
    if cft > 56:
        return "Hipervolemia / retención de fluidos"
    return "Normovolemia"


def construir_dominios_hemodinamicos(df_der: Optional[pd.DataFrame], puntos: Dict[str, Any]) -> Dict[str, Any]:
    ci = _mv(df_der, "CI desde CO importado", "CI")
    ids = _mv(df_der, "SI/IDS")
    irv = _mv(df_der, "SVRI/IRV aproximada")
    fe = _mv(df_der, "FE Capan")
    iv = _mv(df_der, "Indice de velocidad relativo")
    ih = _mv(df_der, "Indice de Heather relativo")
    iac = _mv(df_der, "Indice de aceleracion relativo")
    itc = _mv(df_der, "ITC aproximado")
    rts = _mv(df_der, "STR/CTS morfologico")
    ac = _mv(df_der, "Acoplamiento VA Ea/Ees")
    cft = _mv(df_der, "CFT/TFC")
    return {
        "ci": ci, "ids": ids, "irv": irv, "fe": fe, "iv": iv, "ih": ih,
        "iac": iac, "itc": itc, "rts": rts, "ac": ac, "cft": cft,
        "flujo_ic": clasificar_flujo_ic(ci),
        "flujo_ids": clasificar_ids(ids),
        "flujo_irv": clasificar_irv(irv),
        "funcion": clasificar_fe_capan(fe),
        "rendimiento": clasificar_acoplamiento(ac),
        "volemia": clasificar_volemia_cft(cft),
    }


def tabla_dominios_hemodinamicos(df_der: Optional[pd.DataFrame], puntos: Dict[str, Any]) -> pd.DataFrame:
    d = construir_dominios_hemodinamicos(df_der, puntos)
    fe_pct = None if d["fe"] is None else (d["fe"] * 100 if d["fe"] <= 1.5 else d["fe"])
    return pd.DataFrame([
        {"Dominio": "Flujo", "Valores": f"IC {fmt_metrica('IC', d['ci'])} L/min/m²; IDS {fmt_metrica('IDS', d['ids'])} mL/lat/m²; IRV {fmt_metrica('IRV', d['irv'])}", "Interpretación": f"{d['flujo_ic']}; {d['flujo_ids']}; {d['flujo_irv']}"},
        {"Dominio": "Función cardíaca", "Valores": f"FE Capan {fmt_metrica('FE Capan', fe_pct)} %", "Interpretación": d["funcion"]},
        {"Dominio": "Contractilidad", "Valores": f"IV {fmt_metrica('IV', d['iv'])}; IH {fmt_metrica('IH', d['ih'])}; IAC {fmt_metrica('IAC', d['iac'])}; ITC {fmt_metrica('ITC', d['itc'])}; PEP/LVET {fmt_metrica('CTS', d['rts'])}", "Interpretación": "Perfil contráctil morfológico derivado de la onda validada"},
        {"Dominio": "Rendimiento cardiovascular", "Valores": f"AC {fmt_metrica('AC', d['ac'])}", "Interpretación": d["rendimiento"]},
        {"Dominio": "Volemia", "Valores": f"CFT {fmt_metrica('CFT', d['cft'])} 1/kOhm", "Interpretación": d["volemia"]},
    ])


def diagnostico_hemodinamico_desde_morfologia(df_der: Optional[pd.DataFrame], puntos: Dict[str, Any]) -> str:
    """Conclusión cualitativa sin mezclar valores numéricos.

    Los valores de las métricas quedan exclusivamente en las tablas. Esta
    conclusión resume el diagnóstico por dominios para evitar duplicación y
    mejorar legibilidad del PDF.
    """
    d = construir_dominios_hemodinamicos(df_der, puntos)
    return " ".join([
        f"Dominio de flujo: {d['flujo_ic']}; {d['flujo_ids']}; {d['flujo_irv']}.",
        f"Función cardíaca: {d['funcion']} por FE Capan.",
        "Contractilidad: interpretación basada en IV, IH, IAC, ITC y relación de tiempos sistólicos de la onda validada.",
        f"Rendimiento cardiovascular: {d['rendimiento']} según acoplamiento ventrículo-arterial.",
        f"Volemia: {d['volemia']} por CFT/TFC.",
    ])


def recomendacion_terapeutica_desde_morfologia(datos: Dict[str, Any], df_der: Optional[pd.DataFrame], puntos: Dict[str, Any]) -> str:
    medicacion = str(datos.get("medicacion") or "").strip()
    if not medicacion:
        return "No se identificó tratamiento farmacológico informado; la sugerencia se limita al patrón hemodinámico y no propone optimización farmacológica específica."
    # La recomendacion se genera siempre con los datos morfologicos disponibles.
    # La validacion del operador alimenta la curva de aprendizaje, pero no bloquea la sugerencia.
    d = construir_dominios_hemodinamicos(df_der, puntos)
    recomendaciones = [f"Tratamiento informado: {medicacion}."]
    med_low = medicacion.lower()
    if d["ci"] is not None and d["ci"] > 4.2:
        recomendaciones.append("Patrón hiperdinámico: según el algoritmo ICG, considerar agregar o aumentar beta-bloqueante o calcioantagonista no dihidropiridínico si no existen contraindicaciones.")
    elif d["ci"] is not None and d["ci"] < 2.5:
        recomendaciones.append("Patrón hipodinámico: considerar reducir beta-bloqueo si estuviera presente, salvo indicación comórbida obligatoria, y reevaluar función ventricular.")
    if d["irv"] is not None and d["irv"] > 2580:
        recomendaciones.append("IRV elevada/vasoconstricción: considerar intensificar IECA/ARA-II, calcioantagonista dihidropiridínico o vasodilatador directo según tolerancia, función renal y potasio.")
    elif d["irv"] is not None and d["irv"] < 1300:
        recomendaciones.append("IRV baja: evitar vasodilatación excesiva sin correlación con presión arterial, síntomas y ortostatismo.")
    if d["cft"] is not None and d["cft"] > 56:
        recomendaciones.append("CFT elevado/hipervolemia: considerar agregar o aumentar diurético si la evaluación clínica confirma retención de fluidos.")
    elif d["cft"] is not None and d["cft"] < 41:
        recomendaciones.append("CFT bajo/hipovolemia relativa: evitar intensificación diurética y revisar volumen efectivo, síntomas ortostáticos y función renal.")
    if d["ac"] is not None and d["ac"] > 1.3:
        recomendaciones.append("Desacoplamiento ventrículo-arterial: priorizar reducción de carga arterial y reevaluar acoplamiento tras optimización.")
    if d["fe"] is not None and (d["fe"] * 100 if d["fe"] <= 1.5 else d["fe"]) < 50:
        recomendaciones.append("FE Capan disminuida: correlacionar con ecocardiograma antes de cambios farmacológicos mayores.")
    if any(x in med_low for x in ["losartan", "valsartan", "telmisartan", "irbesartan", "enalapril"]):
        recomendaciones.append("Ya figura bloqueo del sistema renina-angiotensina; revisar dosis, adherencia, creatinina y potasio.")
    if any(x in med_low for x in ["nifedipina", "amlodipina", "lercanidipina"]):
        recomendaciones.append("Ya figura calcioantagonista dihidropiridínico; revisar edema, respuesta tensional y horario/dosis.")
    if any(x in med_low for x in ["doxazosina", "prazosina", "terazosina"]):
        recomendaciones.append("Ya figura alfa-bloqueo; controlar hipotensión ortostática.")
    recomendaciones.append("Sugerencia orientativa: integrar con PA clínica/MAPA, laboratorio, función renal, potasio, síntomas, ECG y ecocardiograma.")
    return " ".join(recomendaciones)




def ramas_terapeuticas_desde_morfologia(df_der: Optional[pd.DataFrame], puntos: Dict[str, Any]) -> Dict[str, bool]:
    """Ramas terapéuticas a resaltar según algoritmo ICG y dominios solicitados."""
    d = construir_dominios_hemodinamicos(df_der, puntos)
    ci = d.get("ci")
    irv = d.get("irv")
    cft = d.get("cft")
    return {
        "hiperdinamico": bool(ci is not None and ci > 4.2),
        "hipodinamico": bool(ci is not None and ci < 2.5),
        "vasoconstrictor": bool(irv is not None and irv > 2500),
        "retencion_fluido": bool(cft is not None and cft > 56),
        "hipovolemia": bool(cft is not None and cft < 41),
    }


def construir_grafico_terapeutico_png(datos: Dict[str, Any], df_der: Optional[pd.DataFrame], puntos: Dict[str, Any]) -> Optional[bytes]:
    """Crea un gráfico terapéutico tipo algoritmo, sin texto largo superpuesto.

    El gráfico solo muestra el flujo decisional y resalta la rama activa.
    La recomendación terapéutica completa queda fuera del gráfico como conclusión aparte.
    """
    try:
        ramas = ramas_terapeuticas_desde_morfologia(df_der, puntos)

        fig, ax = plt.subplots(figsize=(12.0, 6.2))
        ax.set_xlim(0, 1)
        ax.set_ylim(0, 1)
        ax.axis("off")
        fig.patch.set_facecolor("white")
        ax.set_facecolor("white")

        navy = "#0B1F3A"
        blue = "#DFF2FA"
        active = "#FFE3A3"
        active_edge = "#B45309"
        inactive_edge = "#94A3B8"
        teal = "#1D7F78"
        burg = "#9F173D"

        ax.text(
            0.04, 0.955,
            "Algoritmo de optimización terapéutica guiado por ICG",
            fontsize=19, fontweight="bold", color=navy, va="center"
        )
        ax.text(
            0.04, 0.915,
            "El gráfico señala la rama activa; la conclusión terapéutica se informa aparte para evitar superposición de texto.",
            fontsize=9.2, color="#475569", va="center"
        )

        headers = [
            (0.05, 0.82, 0.22, 0.075, "Evaluación"),
            (0.34, 0.82, 0.19, 0.075, "Perfil\nhemodinámico"),
            (0.59, 0.82, 0.18, 0.075, "Implicancia\ndiagnóstica"),
            (0.82, 0.82, 0.14, 0.075, "Rama\nactiva"),
        ]
        for x, y, w, h, txt in headers:
            ax.add_patch(plt.Rectangle(
                (x, y), w, h,
                facecolor=blue if x < 0.80 else "#FFD28A",
                edgecolor=navy, linewidth=1.1
            ))
            ax.text(x+w/2, y+h/2, txt, ha="center", va="center", fontsize=10.0,
                    color=navy, fontweight="bold", linespacing=1.08)

        ax.add_patch(plt.Rectangle((0.05, 0.49), 0.22, 0.16,
                                   facecolor="#F2FAFC", edgecolor=inactive_edge, linewidth=1.0))
        ax.text(0.16, 0.57, "Historia clínica\nPA / MAPA\nlaboratorio\n+ BioZ / CGI",
                ha="center", va="center", fontsize=9.6, color=navy, linespacing=1.13)

        rows = [
            {"key": "hiperdinamico", "y": 0.67, "perfil": "Hiperdinámico", "diag": "Flujo alto", "rama": "Control\nadrenérgico", "color": teal},
            {"key": "hipodinamico", "y": 0.51, "perfil": "Hipodinámico", "diag": "Bajo flujo", "rama": "Revisar\nbeta-bloqueo", "color": burg},
            {"key": "vasoconstrictor", "y": 0.35, "perfil": "Vasoconstrictor", "diag": "Poscarga elevada", "rama": "Reducir\nposcarga", "color": teal},
            {"key": "retencion_fluido", "y": 0.19, "perfil": "Hipervolemia", "diag": "Volumen alto", "rama": "Diurético", "color": "#0F766E"},
        ]
        if ramas.get("hipovolemia") and not ramas.get("retencion_fluido"):
            rows[-1] = {"key": "hipovolemia", "y": 0.19, "perfil": "Hipovolemia", "diag": "Volumen bajo", "rama": "Evitar\ndiurético", "color": "#0369A1"}

        def draw_box(x, y, w, h, text, active_flag=False, color=navy, fontsize=9.4, bold=False, center=False):
            fc = active if active_flag else "#FFFFFF"
            ec = active_edge if active_flag else inactive_edge
            lw = 2.4 if active_flag else 0.9
            ax.add_patch(plt.Rectangle((x, y), w, h, facecolor=fc, edgecolor=ec, linewidth=lw))
            ax.text(x + (w/2 if center else 0.012), y + h/2, str(text),
                    ha="center" if center else "left", va="center", fontsize=fontsize,
                    color=color if active_flag else "#111827",
                    fontweight="bold" if bold else "normal", linespacing=1.08)

        for row in rows:
            y = row["y"]
            is_active = bool(ramas.get(row["key"], False))
            ax.annotate("", xy=(0.34, y+0.045), xytext=(0.27, 0.57), arrowprops=dict(arrowstyle="->", color="#B6C3CC", lw=1.1))
            ax.annotate("", xy=(0.59, y+0.045), xytext=(0.53, y+0.045), arrowprops=dict(arrowstyle="->", color="#B6C3CC", lw=1.1))
            ax.annotate("", xy=(0.82, y+0.045), xytext=(0.77, y+0.045), arrowprops=dict(arrowstyle="->", color=active_edge if is_active else "#B6C3CC", lw=1.7 if is_active else 1.1))
            draw_box(0.34, y, 0.19, 0.09, row["perfil"], is_active, row["color"], fontsize=10.0, bold=True)
            draw_box(0.59, y, 0.18, 0.09, row["diag"], is_active, navy, fontsize=9.4)
            draw_box(0.82, y, 0.14, 0.09, row["rama"], is_active, navy, fontsize=8.8, bold=is_active, center=True)
            if is_active:
                ax.text(0.795, y+0.045, "➜", fontsize=18, color=active_edge, ha="center", va="center", fontweight="bold")

        ax.text(0.05, 0.07,
                "La recomendación farmacológica detallada se redacta fuera del gráfico, en la sección de conclusión terapéutica.",
                fontsize=8.8, color="#475569")

        bio = io.BytesIO()
        fig.savefig(bio, format="png", dpi=190, bbox_inches="tight", facecolor="white")
        plt.close(fig)
        bio.seek(0)
        return bio.getvalue()
    except Exception:
        try:
            plt.close("all")
        except Exception:
            pass
        return None

def perfil_hemodinamico_primario(df_der: Optional[pd.DataFrame], puntos: Dict[str, Any]) -> Dict[str, Any]:
    """Clasifica el perfil primario para la matriz visual del PDF.

    La matriz usa solo las variables solicitadas para el informe:
    IC, IDS e IRV. El umbral principal de flujo es IC > 4.2 para
    hiperdinamia e IC < 2.5 para hipodinamia. La poscarga se categoriza
    con IRV/SVRI > 2500 dyn·s·cm-5·m2, coherente con el algoritmo clínico.
    """
    d = construir_dominios_hemodinamicos(df_der, puntos)
    ci = d.get("ci")
    ids = d.get("ids")
    irv = d.get("irv")
    perfil = "No clasificable"

    active_key = perfil
    comentario_perfil = ""

    if ci is not None and irv is not None and ci > 4.2 and irv > 2500:
        perfil = "Patrón mixto / severo"
        active_key = "Patrón mixto / severo"
        comentario_perfil = "alto flujo + elevada poscarga"
    elif ci is not None and irv is not None and ci < 2.5 and irv > 2500:
        perfil = "Patrón vasoconstrictor con bajo flujo"
        active_key = "Patrón vasoconstrictor"
        comentario_perfil = "bajo flujo + elevada poscarga"
    elif ci is not None and ci > 4.2:
        perfil = "Patrón hiperdinámico"
        active_key = "Patrón hiperdinámico"
        comentario_perfil = "alto flujo predominante"
    elif ci is not None and ci < 2.5:
        perfil = "Patrón hipodinámico"
        active_key = "Patrón hipodinámico"
        comentario_perfil = "bajo flujo predominante"
    elif irv is not None and irv > 2500:
        perfil = "Patrón vasoconstrictor"
        active_key = "Patrón vasoconstrictor"
        comentario_perfil = "poscarga elevada con flujo no hiperdinámico"
    elif ci is not None and 2.5 <= ci <= 4.2 and (irv is None or irv <= 2500):
        perfil = "Patrón compensado / no severo"
        active_key = "Patrón compensado / no severo"
        comentario_perfil = "IC e IRV sin criterios de patrón primario severo"

    return {"perfil": perfil, "active_key": active_key, "comentario_perfil": comentario_perfil, "ci": ci, "ids": ids, "irv": irv, "dominios": d}


def construir_matriz_perfiles_png(df_der: Optional[pd.DataFrame], puntos: Dict[str, Any]) -> Optional[bytes]:
    """Genera una matriz 2x2 tipo infografía para el PDF integrado.

    No agrega variables fuera de las solicitadas: usa IC, IDS e IRV para
    diagnosticar el perfil primario y resalta el cuadrante correspondiente.
    """
    try:
        info = perfil_hemodinamico_primario(df_der, puntos)
        perfil_activo = info.get("perfil", "No clasificable")
        active_key = info.get("active_key", perfil_activo)
        comentario_perfil = info.get("comentario_perfil", "")
        ci = info.get("ci")
        ids = info.get("ids")
        irv = info.get("irv")

        fig, ax = plt.subplots(figsize=(11.2, 6.2))
        ax.set_xlim(0, 1)
        ax.set_ylim(0, 1)
        ax.axis("off")
        fig.patch.set_facecolor("white")
        ax.set_facecolor("white")

        navy = "#0B1F3A"
        teal = "#1D7F78"
        burg = "#9F173D"
        gray_border = "#B6C3CC"
        light_blue = "#F4FAFD"
        light_gray = "#F7F8FA"
        light_peach = "#FFF4E0"
        light_red = "#FFF1F3"
        center = "#2C8C89"

        ax.text(0.035, 0.94, "Matriz de perfiles hemodinámicos primarios",
                fontsize=24, fontweight="bold", color=navy, va="center")
        subtitulo_perfil = f"Perfil desde onda dZ/dt validada: {perfil_activo}"
        if comentario_perfil:
            subtitulo_perfil += f" ({comentario_perfil})"
        subtitulo_perfil += f"  |  IC {fmt(ci,2)} L/min/m²  |  IDS {fmt(ids,1)} mL/lat/m²  |  IRV {fmt(irv,0)}"
        ax.text(0.035, 0.895, subtitulo_perfil,
                fontsize=10.5, color="#475569", va="center")

        cards = [
            {
                "key": "Patrón hiperdinámico",
                "xy": (0.055, 0.52), "wh": (0.41, 0.28),
                "title": "Patrón hiperdinámico",
                "rule": "IC > 4.2",
                "body": "Flujo exageradamente alto.\nPredominio de descarga elevada;\nvalorar tono adrenérgico, FC y poscarga.",
                "rule_color": teal, "fill": light_blue,
            },
            {
                "key": "Patrón mixto / severo",
                "xy": (0.535, 0.52), "wh": (0.41, 0.28),
                "title": "Patrón mixto / severo",
                "rule": "IC > 4.2 + IRV > 2500",
                "body": "Flujo elevado contra poscarga alta.\nCorazón sobreexigido; riesgo\nhemodinámico global aumentado.",
                "rule_color": burg, "fill": light_red,
            },
            {
                "key": "Patrón hipodinámico",
                "xy": (0.055, 0.15), "wh": (0.41, 0.28),
                "title": "Patrón hipodinámico",
                "rule": "IC < 2.5",
                "body": "Flujo bajo. Evaluar reserva\nventricular, volumen efectivo\ny exceso de bloqueo cronotrópico.",
                "rule_color": burg, "fill": light_gray,
            },
            {
                "key": "Patrón vasoconstrictor",
                "xy": (0.535, 0.15), "wh": (0.41, 0.28),
                "title": "Patrón vasoconstrictor",
                "rule": "IRV > 2500",
                "body": "Poscarga/fricción elevada, con\nflujo normal o bajo. Fenotipo\nfrecuente en HTA establecida.",
                "rule_color": teal, "fill": light_blue,
            },
        ]

        # Ejes centrales estilo matriz
        ax.plot([0.5, 0.5], [0.14, 0.83], color=center, linewidth=1.4, alpha=0.7)
        ax.plot([0.055, 0.945], [0.47, 0.47], color=center, linewidth=1.4, alpha=0.7)
        ax.scatter([0.5], [0.47], s=90, color=center, zorder=5)

        for card in cards:
            x, y = card["xy"]
            w, h = card["wh"]
            active = (active_key == card["key"])
            edge = card["rule_color"] if active else navy
            lw = 3.2 if active else 1.6
            fill = card["fill"] if active else "#FFFFFF"
            rect = plt.Rectangle((x, y), w, h, facecolor=fill, edgecolor=edge, linewidth=lw, joinstyle="round")
            ax.add_patch(rect)

            # esquina de estado
            corner_color = card["rule_color"] if active else "#CBD5E1"
            tri = plt.Polygon([[x+w-0.045, y+h], [x+w, y+h], [x+w, y+h-0.045]], closed=True,
                              facecolor=corner_color, edgecolor=corner_color)
            ax.add_patch(tri)
            if active:
                ax.text(x+w-0.018, y+h-0.018, "✓", ha="center", va="center", fontsize=12,
                        color="white", fontweight="bold")

            ax.text(x+0.025, y+h-0.06, card["title"], fontsize=17, fontweight="bold", color=navy, va="top")
            ax.text(x+0.025, y+h-0.112, card["rule"], fontsize=12.5, fontweight="bold", color=card["rule_color"], va="top")
            ax.text(x+0.025, y+h-0.17, card["body"], fontsize=10.8, color="#111827", va="top", linespacing=1.18)

        # Punto individual del paciente dentro de la matriz.
        # Eje X: IRV/SVRI. Eje Y: IC. Se limita al área visible para evitar que el punto quede fuera del diagrama.
        def _clip(v, vmin, vmax):
            try:
                return max(vmin, min(vmax, float(v)))
            except Exception:
                return None

        if ci is not None and irv is not None:
            irv_min, irv_max = 1000.0, 4000.0
            ci_min, ci_max = 1.5, 5.5
            plot_x_min, plot_x_max = 0.075, 0.925
            plot_y_min, plot_y_max = 0.175, 0.795

            irv_c = _clip(irv, irv_min, irv_max)
            ci_c = _clip(ci, ci_min, ci_max)
            if irv_c is not None and ci_c is not None:
                xp = plot_x_min + (irv_c - irv_min) / (irv_max - irv_min) * (plot_x_max - plot_x_min)
                yp = plot_y_min + (ci_c - ci_min) / (ci_max - ci_min) * (plot_y_max - plot_y_min)

                # Líneas de referencia del punto.
                ax.plot([xp, xp], [plot_y_min, yp], linestyle=":", linewidth=1.4, color="#0B1F3A", alpha=0.55, zorder=6)
                ax.plot([plot_x_min, xp], [yp, yp], linestyle=":", linewidth=1.4, color="#0B1F3A", alpha=0.55, zorder=6)
                ax.scatter([xp], [yp], s=330, facecolor="#FFD166", edgecolor="#0B1F3A", linewidth=2.4, zorder=10)
                ax.scatter([xp], [yp], s=75, facecolor="#0B1F3A", edgecolor="#0B1F3A", zorder=11)

                label_x = xp + 0.018 if xp < 0.76 else xp - 0.22
                label_y = yp + 0.045 if yp < 0.72 else yp - 0.075
                ax.text(label_x, label_y,
                        f"Paciente\nIC {fmt(ci,2)} | IRV {fmt(irv,0)}",
                        fontsize=9.8, color=navy, fontweight="bold",
                        bbox=dict(boxstyle="round,pad=0.35", facecolor="#FFF7CC", edgecolor="#0B1F3A", linewidth=1.2),
                        zorder=12)

                # Umbrales clínicos usados para ubicar el punto.
                x_thr = plot_x_min + (2500.0 - irv_min) / (irv_max - irv_min) * (plot_x_max - plot_x_min)
                y_thr_low = plot_y_min + (2.5 - ci_min) / (ci_max - ci_min) * (plot_y_max - plot_y_min)
                y_thr_high = plot_y_min + (4.2 - ci_min) / (ci_max - ci_min) * (plot_y_max - plot_y_min)
                ax.axvline(x_thr, ymin=0.13, ymax=0.84, linestyle="--", linewidth=0.9, color="#64748B", alpha=0.45)
                ax.axhline(y_thr_low, xmin=0.055, xmax=0.945, linestyle="--", linewidth=0.9, color="#64748B", alpha=0.35)
                ax.axhline(y_thr_high, xmin=0.055, xmax=0.945, linestyle="--", linewidth=0.9, color="#64748B", alpha=0.35)
                ax.text(x_thr+0.006, 0.835, "IRV 2500", fontsize=8.5, color="#475569", rotation=90, va="top")
                ax.text(0.935, y_thr_low+0.006, "IC 2.5", fontsize=8.5, color="#475569", ha="right")
                ax.text(0.935, y_thr_high+0.006, "IC 4.2", fontsize=8.5, color="#475569", ha="right")

        ax.text(0.055, 0.07,
                "La matriz se integra al diagnóstico por dominios: flujo (IC, IDS, IRV), función cardíaca, contractilidad, acoplamiento VA y volemia.",
                fontsize=9.5, color="#475569")

        bio = io.BytesIO()
        fig.savefig(bio, format="png", dpi=190, bbox_inches="tight", facecolor="white")
        plt.close(fig)
        bio.seek(0)
        return bio.getvalue()
    except Exception:
        try:
            plt.close("all")
        except Exception:
            pass
        return None

# ============================================================
# PDF
# ============================================================

def generar_pdf_informe(df: pd.DataFrame, morfo: Optional[Dict[str, Any]] = None) -> bytes:
    """PDF exclusivo del análisis morfológico validado por dominios."""
    if A4 is None:
        raise RuntimeError("ReportLab no esta instalado. Agregar reportlab a requirements.txt")
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=1.4*cm, leftMargin=1.4*cm, topMargin=1.3*cm, bottomMargin=1.2*cm)
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name="TituloAzul", parent=styles["Title"], fontSize=16, leading=20, textColor=colors.HexColor("#082F49"), spaceAfter=8))
    styles.add(ParagraphStyle(name="Sub", parent=styles["Heading2"], fontSize=12, leading=15, textColor=colors.HexColor("#0B4F7A"), spaceBefore=8, spaceAfter=6))
    styles.add(ParagraphStyle(name="NormalSmall", parent=styles["BodyText"], fontSize=8.2, leading=10.4))
    styles.add(ParagraphStyle(name="Advertencia", parent=styles["BodyText"], fontSize=8.2, leading=10.4, textColor=colors.HexColor("#7C2D12")))
    elems: List[Any] = []
    basal_pdf, _ = seleccionar_basal_y_parado(df)
    datos = basal_pdf or {}
    paciente = datos.get("paciente", "No disponible")
    elems.append(Paragraph("Análisis morfológico de onda de impedancia dZ/dt con cursores QRS-B-C-X", styles["TituloAzul"]))
    elems.append(Paragraph(safe_pdf_text(AUTOR_APP), styles["NormalSmall"]))
    elems.append(Spacer(1, 8))
    if morfo is None or morfo.get("df_curva") is None or morfo.get("puntos") is None:
        elems.append(Paragraph("No se adjuntó una curva de impedancia digitalizada con cursores QRS-B-C-X.", styles["Advertencia"]))
        doc.build(elems); buffer.seek(0); return buffer.getvalue()
    puntos = morfo.get("puntos") or {}
    validado = bool(puntos.get("validado_operador", False))
    tabla_datos = [["Paciente", safe_pdf_text(paciente), "Estado", "Análisis morfológico generado"], ["Base analítica", "Onda dZ/dt con QRS-B-C-X", "Archivo", safe_pdf_text(datos.get("archivo_origen", ""))]]
    table = Table(tabla_datos, colWidths=[3.0*cm, 6.0*cm, 2.5*cm, 5.5*cm])
    table.setStyle(TableStyle([("BACKGROUND", (0,0), (-1,0), colors.HexColor("#EAF6FF")), ("GRID", (0,0), (-1,-1), .35, colors.HexColor("#D7E3EE")), ("FONTNAME", (0,0), (-1,-1), "Helvetica"), ("FONTSIZE", (0,0), (-1,-1), 8), ("VALIGN", (0,0), (-1,-1), "TOP")]))
    elems.append(table); elems.append(Spacer(1, 6))
    png_multi = morfo.get("png_canales")
    if png_multi:
        elems.append(Paragraph("Registro de apoyo temporal: dZ/dt + ECG + marcas fonocardiográficas", styles["Sub"]))
        elems.append(Paragraph("El fonocardiograma se informa solo como barras temporales de S1/S2; no se usa como curva cuantitativa.", styles["NormalSmall"]))
        elems.append(RLImage(io.BytesIO(png_multi), width=17.5*cm, height=7.2*cm)); elems.append(Spacer(1, 8))
    png = morfo.get("png")
    if png:
        elems.append(Paragraph("Onda dZ/dt con cursores QRS-B-C-X", styles["Sub"]))
        elems.append(RLImage(io.BytesIO(png), width=17.5*cm, height=7.2*cm)); elems.append(Spacer(1, 8))
    elems.append(Paragraph("Cursores y tiempos sistólicos", styles["Sub"]))
    df_tab = tabla_resultados_qbcx(puntos)
    # Compatibilidad: algunas versiones de la app generan la columna sin tilde (Metrica)
    # y otras con tilde (Métrica). Evita KeyError al construir el PDF.
    col_metrica = "Métrica" if "Métrica" in df_tab.columns else ("Metrica" if "Metrica" in df_tab.columns else df_tab.columns[0])
    keep_rows = df_tab[df_tab[col_metrica].astype(str).str.contains("Q|B|C|X|PEP|LVET|Confianza", case=False, na=False)].copy()
    if keep_rows.empty:
        keep_rows = df_tab.copy()
    data = [keep_rows.columns.tolist()] + keep_rows.values.tolist(); data = [[safe_pdf_text(c) for c in row] for row in data]
    tt = Table(data, colWidths=[4.2*cm, 3.4*cm, 9.4*cm])
    tt.setStyle(TableStyle([("BACKGROUND", (0,0), (-1,0), colors.HexColor("#082F49")), ("TEXTCOLOR", (0,0), (-1,0), colors.white), ("GRID", (0,0), (-1,-1), .3, colors.HexColor("#D7E3EE")), ("VALIGN", (0,0), (-1,-1), "TOP"), ("FONTSIZE", (0,0), (-1,-1), 7.2)]))
    elems.append(tt); elems.append(Spacer(1, 8))
    df_der = calcular_variables_icg_derivadas(datos, morfo)
    df_der_pdf = filtrar_metricas_pdf_morfologicas(df_der)
    elems.append(Paragraph("Diagnóstico por dominios hemodinámicos", styles["Sub"]))
    df_dom = tabla_dominios_hemodinamicos(df_der_pdf, puntos)
    data_dom = [df_dom.columns.tolist()] + df_dom.values.tolist(); data_dom = [[safe_pdf_text(c) for c in row] for row in data_dom]
    tdom = Table(data_dom, colWidths=[4.0*cm, 6.2*cm, 6.8*cm], repeatRows=1)
    tdom.setStyle(TableStyle([("BACKGROUND", (0,0), (-1,0), colors.HexColor("#0B4F7A")), ("TEXTCOLOR", (0,0), (-1,0), colors.white), ("GRID", (0,0), (-1,-1), .3, colors.HexColor("#D7E3EE")), ("VALIGN", (0,0), (-1,-1), "TOP"), ("FONTSIZE", (0,0), (-1,-1), 7.0)]))
    elems.append(tdom); elems.append(Spacer(1, 8))
    matriz_png = construir_matriz_perfiles_png(df_der_pdf, puntos)
    if matriz_png:
        elems.append(Paragraph("Matriz visual de perfiles hemodinámicos primarios", styles["Sub"]))
        elems.append(RLImage(io.BytesIO(matriz_png), width=17.5*cm, height=9.6*cm))
        elems.append(Spacer(1, 8))
    elems.append(Paragraph("Variables incluidas en el informe", styles["Sub"]))
    if not df_der_pdf.empty:
        data_m = [["Variable", "Valor", "Unidad", "Interpretación/estado"]]
        for _, rr_m in df_der_pdf.iterrows():
            valor_vis = rr_m.get("Valor mostrado", rr_m.get("Valor recalculado", ""))
            estado_vis = rr_m.get("Control de coherencia", rr_m.get("Estado", ""))
            data_m.append([safe_pdf_text(rr_m.get("Variable", "")), safe_pdf_text(valor_vis), safe_pdf_text(rr_m.get("Unidad", "")), safe_pdf_text(estado_vis)])
        tmet = Table(data_m, colWidths=[4.2*cm, 2.4*cm, 3.0*cm, 7.2*cm], repeatRows=1)
        tmet.setStyle(TableStyle([("BACKGROUND", (0,0), (-1,0), colors.HexColor("#0B4F7A")), ("TEXTCOLOR", (0,0), (-1,0), colors.white), ("GRID", (0,0), (-1,-1), .3, colors.HexColor("#D7E3EE")), ("VALIGN", (0,0), (-1,-1), "TOP"), ("FONTSIZE", (0,0), (-1,-1), 7.0)]))
        elems.append(tmet)
    else:
        elems.append(Paragraph("No se calcularon métricas suficientes desde la curva validada.", styles["NormalSmall"]))
    avisos_coh = control_consistencia_basica(datos, df_der_pdf)
    if avisos_coh:
        elems.append(Spacer(1, 6))
        elems.append(Paragraph("Control de coherencia de métricas", styles["Sub"]))
        for av in avisos_coh[:6]:
            elems.append(Paragraph(safe_pdf_text("• " + av), styles["Advertencia"]))
    elems.append(Spacer(1, 8))
    elems.append(Paragraph("Conclusión integrada por dominios", styles["Sub"]))
    elems.append(Paragraph(
        "La conclusión se expresa en forma cualitativa. Los valores numéricos quedan separados en las tablas de métricas y dominios para evitar duplicación.",
        styles["NormalSmall"]
    ))
    elems.append(Paragraph(safe_pdf_text(diagnostico_hemodinamico_desde_morfologia(df_der_pdf, puntos)), styles["NormalSmall"]))
    elems.append(Spacer(1, 8))

    grafico_tx = construir_grafico_terapeutico_png(datos, df_der_pdf, puntos)
    elems.append(Paragraph("Algoritmo gráfico de orientación terapéutica", styles["Sub"]))
    if grafico_tx:
        elems.append(RLImage(io.BytesIO(grafico_tx), width=17.5*cm, height=8.7*cm))
        elems.append(Spacer(1, 8))
    elems.append(Paragraph("Conclusión terapéutica orientativa", styles["Sub"]))
    elems.append(Paragraph(safe_pdf_text(recomendacion_terapeutica_desde_morfologia(datos, df_der_pdf, puntos)), styles["NormalSmall"]))
    elems.append(Spacer(1, 6))
    doc.build(elems); buffer.seek(0); return buffer.getvalue()

st.write("Subir PDF, Excel o CSV. El patron hemodinamico principal se toma del registro basal/acostado/cinta/spot; el registro parado se informa solo como comportamiento ortostatico.")
archivos = st.file_uploader(
    "Archivos CGI / Z-Logic",
    type=["pdf", "xlsx", "xls", "csv"],
    accept_multiple_files=True,
)
st.markdown("</div>", unsafe_allow_html=True)

if "df_final" not in st.session_state:
    st.session_state["df_final"] = pd.DataFrame()

if archivos:
    partes = []
    errores = []
    for f in archivos:
        try:
            # conservar bytes del primer PDF para digitalizacion automatica del panel superior derecho
            if str(getattr(f, "name", "")).lower().endswith(".pdf") and "pdf_zlogic_auto_bytes" not in st.session_state:
                try:
                    st.session_state["pdf_zlogic_auto_bytes"] = f.getvalue()
                    st.session_state["pdf_zlogic_auto_name"] = getattr(f, "name", "informe_zlogic.pdf")
                except Exception:
                    pass
            dfx = leer_archivo(f)
            if not dfx.empty:
                partes.append(dfx)
            else:
                errores.append(f"{f.name}: sin datos estructurados detectables")
        except Exception as e:
            errores.append(f"{f.name}: {e}")
    if partes:
        st.session_state["df_final"] = completar_antropometria_e_indices(pd.concat(partes, ignore_index=True))
    if errores:
        st.warning("\n".join(errores))

df_final = st.session_state.get("df_final", pd.DataFrame())

# Variables globales de contexto para que el módulo morfológico pueda ejecutarse
# incluso si el usuario usa solo el PDF/captura para digitalizar la curva.
basal = {}
parado = None
if isinstance(df_final, pd.DataFrame) and not df_final.empty:
    try:
        basal, parado = seleccionar_basal_y_parado(df_final)
        basal = basal or {}
    except Exception:
        basal = {}
        parado = None

if not df_final.empty:
    st.markdown("<div class='card'>", unsafe_allow_html=True)
    st.subheader("2. Datos integrados estructurados")
    st.dataframe(df_final, use_container_width=True)
    basal, parado = seleccionar_basal_y_parado(df_final)
    cols = st.columns(4)
    cols[0].metric("Patron basal", clasificar_dinamia(basal))
    cols[1].metric("IC basal", fmt((basal or {}).get("ic"), 2))
    cols[2].metric("IRV/RVS basal", fmt((basal or {}).get("irv") or (basal or {}).get("rvs"), 0))
    cols[3].metric("Registro parado", "Si" if parado is not None else "No")
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown("<div class='card'>", unsafe_allow_html=True)
    st.subheader("3. Informe medico integrado")
    informe = generar_informe_texto(df_final)
    st.text_area("Texto del informe", informe, height=300)
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown("<div class='card'>", unsafe_allow_html=True)
    st.subheader("4. Base de ecuaciones, contractilidad y variables recalculadas ICG")
    basal_tmp, _ = seleccionar_basal_y_parado(df_final)
    morfo_tmp = None
    if "puntos_qbcx_final" in st.session_state:
        morfo_tmp = {"puntos": st.session_state["puntos_qbcx_final"], "df_curva": st.session_state.get("df_curva_icg")}
    df_derivadas = calcular_variables_icg_derivadas(basal_tmp or {}, morfo_tmp)
    st.markdown("**Variables importadas/recalculadas disponibles con los datos actuales, incluyendo contractilidad**")
    if not df_derivadas.empty:
        mostrar_tabla_metricas_con_barras(df_derivadas, height=520)
    else:
        st.info("No se pudieron recalcular variables porque faltan datos como PAS/PAD, FC, talla/peso, SV/DS, CO/GC, CI/IC o Z0.")
    with st.expander("Ver base de conocimiento de ecuaciones ICG"):
        st.dataframe(tabla_ecuaciones_icg(), use_container_width=True)
        st.caption("Ea/Ees se recalcula con PAS, PAD, SV/DS y tiempos PEP/LVET. Si se corrigen cursores en la curva digitalizada, esos tiempos tienen prioridad para el cálculo de FE Capan, Ees y acoplamiento ventriculo-arterial. Las variables morfologicas de contractilidad son relativas si provienen de imagen. Las ecuaciones de Kubicek y Sramek-Bernstein requieren señal dZ/dt calibrada en amplitud absoluta.")
    st.markdown("</div>", unsafe_allow_html=True)
else:
    st.info("Cargar al menos un archivo para generar el informe. Tambien puede usarse solo el modulo morfologico con una imagen de curva ICG.")

# ============================================================
# SECCION NUEVA: ANALISIS MORFOLOGICO SEMIAUTOMATICO
# ============================================================

st.markdown("<div class='card'>", unsafe_allow_html=True)
st.subheader("5. Analisis morfologico semiautomatico de señal ICG")
st.write(
    "La app puede digitalizar automaticamente el sector superior derecho del PDF Z-Logic "
    "(dZ/dt, ECG y fonocardiograma) o trabajar con una captura manual. "
    "Luego propone QRS-B-C-X, permite correccion manual y recalcula PEP, LVET, contractilidad, Ea, Ees y acoplamiento VA."
)

mostrar_barra_aprendizaje()

modo_morfo = st.radio(
    "Fuente para digitalizacion morfologica",
    ["PDF Z-Logic automatico - sector superior derecho", "Imagen/captura manual"],
    horizontal=True,
)

fc_det = obtener_fc_desde_df(df_final)
rr_det = obtener_rr_desde_df(df_final)
if rr_det is None:
    rr_det = ciclo_ms_desde_fc(fc_det, 800.0)

if modo_morfo.startswith("PDF"):
    st.info(
        "El recorte automatico toma el panel superior derecho del informe: banda superior dZ/dt, banda media ECG verde y banda mas inferior fonocardiograma naranja. "
        "El eje X se calibra por RR del informe si está disponible; el grillado del grafico generado es cada 20 ms."
    )
    pdf_bytes_auto = st.session_state.get("pdf_zlogic_auto_bytes")
    pdf_auto_file = st.file_uploader(
        "PDF Z-Logic para digitalizacion automatica del sector superior derecho",
        type=["pdf"],
        key="pdf_auto_morfo",
        help="Si ya subio un PDF en la seccion 1, se usa automaticamente. Este cargador permite reemplazarlo."
    )
    if pdf_auto_file is not None:
        pdf_bytes_auto = pdf_auto_file.getvalue()
        st.session_state["pdf_zlogic_auto_bytes"] = pdf_bytes_auto
        st.session_state["pdf_zlogic_auto_name"] = getattr(pdf_auto_file, "name", "informe_zlogic.pdf")

    # Boton de captura automatica colocado arriba, inmediatamente despues de cargar el PDF/captura.
    # Usa recorte Z-Logic por defecto; los controles finos permanecen debajo.
    if pdf_bytes_auto is not None:
        st.markdown("#### Captura automatica inmediata")
        if st.button("Captura automatica del panel superior derecho del PDF", key="btn_pdf_auto_top"):
            try:
                canales = digitalizar_pdf_zlogic_sector_superior_derecho(
                    pdf_bytes_auto,
                    rr_ms=float(rr_det),
                    crop_rel=(0.735, 0.292, 0.985, 0.675),
                    suavizado=5,
                    invertir_dzdt=False,
                )
                puntos = detectar_puntos_qbcx_icg(canales["df_curva"], fc_det)
                puntos = ajustar_puntos_con_tiempos_importados(canales["df_curva"], puntos, basal or {})
                st.session_state["canales_zlogic_pdf"] = canales
                st.session_state["df_curva_icg"] = canales["df_curva"]
                st.session_state["puntos_qbcx_auto"] = puntos
                st.session_state["puntos_qbcx_final"] = puntos
                st.session_state["qbcx_validado_operador"] = False
                st.success("Captura automatica realizada. Debajo puede ajustar cursores, amplificacion y recorte si hace falta.")
            except Exception as e:
                st.error(f"No se pudo realizar la captura automatica inicial: {e}")

    colauto1, colauto2, colauto3 = st.columns(3)
    with colauto1:
        rr_ms_user = st.number_input("RR/ciclo para eje X (ms)", min_value=300.0, max_value=2000.0, value=float(round(rr_det, 1)), step=1.0)
    with colauto2:
        suav_pdf = st.slider("Suavizado digitalizacion PDF", 3, 21, 5, 2)
    with colauto3:
        invertir_pdf = st.checkbox("Invertir dZ/dt del PDF", value=False)

    colamp1, colamp2 = st.columns(2)
    with colamp1:
        amplificar_ecg = st.slider("Amplificación visual ECG verde", 0.7, 4.0, 1.80, 0.05)
    with colamp2:
        amplificar_fono = st.slider("Amplificación visual fonocardiograma naranja inferior", 1.0, 8.0, 4.00, 0.05)
    st.session_state["amplificar_ecg_visual"] = float(amplificar_ecg)
    st.session_state["amplificar_fono_visual"] = float(amplificar_fono)

    with st.expander("Ajuste fino del recorte automatico superior derecho"):
        st.caption("Valores relativos a la pagina. En el informe ejemplo el panel superior derecho se encuentra aproximadamente entre X 0,735-0,985 e Y 0,292-0,675.")
        ca, cb, cc, cd = st.columns(4)
        with ca:
            rel_x1 = st.slider("X inicial", 0.50, 0.90, 0.735, 0.005)
        with cb:
            rel_y1 = st.slider("Y inicial", 0.10, 0.50, 0.292, 0.005)
        with cc:
            rel_x2 = st.slider("X final", 0.75, 1.00, 0.985, 0.005)
        with cd:
            rel_y2 = st.slider("Y final", 0.40, 0.90, 0.675, 0.005)

    if pdf_bytes_auto is not None:
        if st.button("Digitalizar automaticamente PDF: dZ/dt + ECG + Fono", key="btn_pdf_auto"):
            try:
                canales = digitalizar_pdf_zlogic_sector_superior_derecho(
                    pdf_bytes_auto,
                    rr_ms=float(rr_ms_user),
                    crop_rel=(rel_x1, rel_y1, rel_x2, rel_y2),
                    suavizado=int(suav_pdf),
                    invertir_dzdt=bool(invertir_pdf),
                )
                puntos = detectar_puntos_qbcx_icg(canales["df_curva"], fc_det)
                puntos = ajustar_puntos_con_tiempos_importados(canales["df_curva"], puntos, basal or {})
                st.session_state["canales_zlogic_pdf"] = canales
                st.session_state["df_curva_icg"] = canales["df_curva"]
                st.session_state["puntos_qbcx_auto"] = puntos
                st.session_state["puntos_qbcx_final"] = puntos
                st.session_state["qbcx_validado_operador"] = False
                st.success("Sector superior derecho digitalizado. Revisar panel y confirmar/corregir cursores.")
            except Exception as e:
                st.error(f"No se pudo digitalizar automaticamente el PDF: {e}")

        if "canales_zlogic_pdf" in st.session_state:
            canales = st.session_state["canales_zlogic_pdf"]
            st.image(canales["panel"], caption="Panel superior derecho recortado automaticamente", use_container_width=True)
            st.info("Canales detectados: dZ/dt en banda superior, ECG en banda media y FONO en la banda inferior resaltada. La app excluye la franja amarilla para no confundirla con la señal.")
            st.caption(
                f"RR usado: {canales.get('rr_ms'):.1f} ms. Espaciado de grilla estimado: "
                f"{fmt(canales.get('grid_px'), 2)} px. Si el recorte incluye textos, ajustar los límites relativos."
            )
            fig_canales_pre = graficar_canales_zlogic(canales, st.session_state.get("puntos_qbcx_final"), amplificar_fono=float(amplificar_fono), amplificar_ecg=float(amplificar_ecg), expandir_vertical=True)
            st.pyplot(fig_canales_pre, use_container_width=True)
            st.session_state["canales_zlogic_png"] = fig_to_png_bytes(fig_canales_pre)

    else:
        st.warning("Cargar un PDF Z-Logic en esta sección o en la sección 1 para habilitar la digitalización automática.")

else:
    img_file = st.file_uploader("Imagen/captura de curva promedio ICG", type=["png", "jpg", "jpeg"], key="img_icg")
    if img_file is not None:
        imagen = Image.open(img_file).convert("RGB")
        ancho, alto = imagen.size
        st.image(imagen, caption="Imagen original", use_container_width=True)

        st.warning(
            "Para evitar curvas deformadas, recortar solo la señal dZ/dt promedio. "
            "No incluir ECG, fonocardiograma, carotidograma, texto, números, bordes ni más de una curva."
        )

        ciclo_sug = ciclo_ms_desde_fc(fc_det, 800.0)
        c1, c2, c3 = st.columns(3)
        with c1:
            duracion_ms = st.number_input("Duracion del ciclo (ms)", min_value=300.0, max_value=2000.0, value=float(round(ciclo_sug, 1)), step=10.0)
        with c2:
            umbral = st.slider("Umbral de oscuridad", 20, 245, 120, 5)
        with c3:
            invertir = st.checkbox("Invertir polaridad", value=False, help="Activar si el pico C queda hacia abajo o aparece una alerta de polaridad invertida.")

        st.markdown("**Recorte de la zona util de la curva**")
        r1, r2 = st.columns(2)
        with r1:
            x1 = st.slider("X1 izquierda", 0, ancho-1, 0)
            x2 = st.slider("X2 derecha", 1, ancho, ancho)
        with r2:
            y1 = st.slider("Y1 superior", 0, alto-1, 0)
            y2 = st.slider("Y2 inferior", 1, alto, alto)

        suav = st.slider("Suavizado", 3, 31, 9, 2)
        detectar_color = st.checkbox(
            "Detectar trazos coloreados ademas de oscuros",
            value=True,
            help="Mantener activado si la curva original es azul, verde o roja. Desactivar si capta textos o marcas de color."
        )
        max_salto_px = st.slider(
            "Continuidad maxima entre puntos de curva (px)",
            4,
            max(8, int(max(10, (y2 - y1) * 0.25))),
            max(6, int(max(8, (y2 - y1) * 0.10))),
            1,
            help="Si la curva queda cortada, aumentar. Si salta a otra señal, disminuir."
        )

        if x2 > x1 and y2 > y1:
            st.image(imagen.crop((x1, y1, x2, y2)), caption="Vista previa del recorte usado para digitalizar", use_container_width=True)

        if st.button("Digitalizar curva y proponer QRS-B-C-X"):
            try:
                df_curva = digitalizar_curva_icg_desde_imagen(
                    imagen, x1, x2, y1, y2, duracion_ms,
                    umbral_oscuridad=umbral,
                    invertir_polaridad=invertir,
                    suavizado=suav,
                    detectar_color=detectar_color,
                    max_salto_px=max_salto_px,
                )
                puntos = detectar_puntos_qbcx_icg(df_curva, fc_det)
                st.session_state["df_curva_icg"] = df_curva
                st.session_state["puntos_qbcx_auto"] = puntos
                st.session_state["puntos_qbcx_final"] = puntos
                st.session_state["qbcx_validado_operador"] = False
                st.success("Curva digitalizada y cursores propuestos. Validar/corregir manualmente debajo antes de usar Ea/Ees.")
            except Exception as e:
                st.error(f"No se pudo procesar la curva: {e}")

if "df_curva_icg" in st.session_state and "puntos_qbcx_auto" in st.session_state:
    df_curva = st.session_state["df_curva_icg"]
    pauto = st.session_state["puntos_qbcx_auto"]
    tmin = float(df_curva["tiempo_ms"].min())
    tmax = float(df_curva["tiempo_ms"].max())
    st.markdown("### Correccion manual de cursores")
    qcol, bcol, ccol, xcol = st.columns(4)
    with qcol:
        qms = st.slider("QRS/Q (ms)", tmin, tmax, float(pauto["q_ms"]), 1.0)
    with bcol:
        bms = st.slider("B (ms)", tmin, tmax, float(pauto["b_ms"]), 1.0)
    with ccol:
        cms = st.slider("C (ms)", tmin, tmax, float(pauto["c_ms"]), 1.0)
    with xcol:
        xms = st.slider("X (ms)", tmin, tmax, float(pauto["x_ms"]), 1.0)

    st.markdown("#### Validación y aprendizaje de cursores")
    st.caption(
        "La app cuenta una validación cuando, después de corregir los cursores, se presiona "
        "el botón **Registrar validación de cursores**. Cada pulsación guarda la diferencia "
        "absoluta en ms entre la propuesta automática y la posición final de QRS, B, C y X. "
        "Ese historial alimenta la barra y la gráfica de aprendizaje. La validación no bloquea "
        "las gráficas, el diagnóstico ni las sugerencias terapéuticas."
    )

    pfinal = recalcular_puntos_manuales(df_curva, qms, bms, cms, xms)

    col_val1, col_val2 = st.columns([1, 2])
    with col_val1:
        if st.button("Registrar validación de cursores", key="btn_registrar_validacion_qbcx"):
            estado_ap = actualizar_aprendizaje_cursores(pauto, pfinal, True)
            st.session_state["qbcx_validado_operador"] = True
            st.session_state["ultima_validacion_qbcx"] = {
                "q_ms": float(pfinal.get("q_ms", np.nan)),
                "b_ms": float(pfinal.get("b_ms", np.nan)),
                "c_ms": float(pfinal.get("c_ms", np.nan)),
                "x_ms": float(pfinal.get("x_ms", np.nan)),
                "n": int(estado_ap.get("n", 0)),
            }
            st.success(f"Validación registrada. Total acumulado: {int(estado_ap.get('n', 0))}.")
            # Forzar recarga para que la barra de aprendizaje ubicada arriba
            # refleje inmediatamente la nueva validación en la misma pantalla.
            st.rerun()

    with col_val2:
        if st.button("Reiniciar contador de aprendizaje", key="btn_reiniciar_aprendizaje_qbcx"):
            st.session_state["aprendizaje_cursores"] = {
                "n": 0,
                "error_ms_acum": 0.0,
                "error_ms_prom": None,
                "historial_error_ms": [],
                "historial_q_ms": [],
                "historial_b_ms": [],
                "historial_c_ms": [],
                "historial_x_ms": [],
            }
            st.session_state["qbcx_validado_operador"] = False
            st.session_state.pop("ultima_validacion_qbcx", None)
            st.info("Contador de aprendizaje reiniciado.")

    pfinal["validado_operador"] = bool(st.session_state.get("qbcx_validado_operador", False))

    # La ausencia de confirmacion no bloquea graficas, diagnostico ni sugerencias.
    # El aprendizaje se registra explícitamente con el botón de validación.
    st.session_state["puntos_qbcx_final"] = pfinal

    fig = graficar_curva_icg_con_puntos(df_curva, pfinal)
    st.pyplot(fig, use_container_width=True)
    png_bytes = fig_to_png_bytes(fig)
    st.session_state["qbcx_png"] = png_bytes

    if "canales_zlogic_pdf" in st.session_state:
        fig_canales = graficar_canales_zlogic(st.session_state["canales_zlogic_pdf"], pfinal, amplificar_fono=float(st.session_state.get("amplificar_fono_visual", 4.00)), amplificar_ecg=float(st.session_state.get("amplificar_ecg_visual", 1.80)), expandir_vertical=True)
        st.markdown("### Registro multicanal digitalizado con grilla cada 20 ms")
        st.caption("Convención visual corregida: ECG en verde en el canal medio y fonocardiograma en naranja en el canal inferior. La gráfica está expandida verticalmente para resaltar el fonocardiograma.")
        st.pyplot(fig_canales, use_container_width=True)
        st.session_state["canales_zlogic_png"] = fig_to_png_bytes(fig_canales)

    tab_qbcx = tabla_resultados_qbcx(pfinal)
    st.dataframe(tab_qbcx, use_container_width=True)
    interpretacion = texto_interpretacion_qbcx(pfinal)
    st.info(interpretacion)
    st.session_state["interpretacion_qbcx"] = interpretacion

    if not df_final.empty:
        basal_para_morfo, _ = seleccionar_basal_y_parado(df_final)
        df_der_morfo = calcular_variables_icg_derivadas(basal_para_morfo or {}, {"puntos": pfinal, "df_curva": df_curva})
        st.markdown("### Variables recalculadas integrando cursores, antropometria, indices y contractilidad morfologica")
        if not df_der_morfo.empty:
            mostrar_metricas_hemodinamicas_digitalizacion(df_der_morfo, pfinal)
            st.session_state["df_metricas_hemodinamicas_digitalizacion"] = df_der_morfo
        else:
            st.info("No se recalcularon variables adicionales por falta de datos suficientes.")

    df_multi_export = None
    if "canales_zlogic_pdf" in st.session_state:
        canales = st.session_state["canales_zlogic_pdf"]
        df_multi_export = pd.DataFrame({
            "tiempo_ms": canales["df_dzdt"]["tiempo_ms"],
            "dzdt": canales["df_dzdt"]["amplitud"],
            "ecg": np.interp(canales["df_dzdt"]["tiempo_ms"], canales["df_ecg"]["tiempo_ms"], canales["df_ecg"]["amplitud"]),
            "fono": np.interp(canales["df_dzdt"]["tiempo_ms"], canales["df_fono"]["tiempo_ms"], canales["df_fono"]["amplitud"]),
        })

    df_metricas_export = st.session_state.get("df_metricas_hemodinamicas_digitalizacion", pd.DataFrame())
    aprendizaje_export = st.session_state.get("aprendizaje_cursores", {})
    excel_metricas = excel_metricas_hemodinamicas_bytes(
        df_metricas=df_metricas_export,
        df_qbcx=tab_qbcx,
        df_curva=df_curva,
        df_canales=df_multi_export,
        df_final=df_final,
        aprendizaje=aprendizaje_export,
    )

    cols_dl = st.columns(5)
    cols_dl[0].download_button("Descargar dZ/dt CSV", df_curva.to_csv(index=False).encode("utf-8"), "curva_dzdt_icg_digitalizada.csv", "text/csv")
    cols_dl[1].download_button("Descargar cursores CSV", tab_qbcx.to_csv(index=False).encode("utf-8"), "cursores_qbcx_icg.csv", "text/csv")
    cols_dl[2].download_button("Descargar grafico QBCX PNG", png_bytes, "grafico_qbcx_icg.png", "image/png")
    if df_multi_export is not None:
        cols_dl[3].download_button("Descargar dZ/dt+ECG+Fono CSV", df_multi_export.to_csv(index=False).encode("utf-8"), "canales_zlogic_digitalizados.csv", "text/csv")
    cols_dl[4].download_button(
        "Exportar metricas hemodinamicas Excel",
        data=excel_metricas,
        file_name="metricas_hemodinamicas_curva_impedancia.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

st.markdown("</div>", unsafe_allow_html=True)


# ============================================================
# DESCARGA PDF FINAL
# ============================================================

st.markdown("<div class='card'>", unsafe_allow_html=True)
st.subheader("6. Exportación de análisis morfológico validado")
if not df_final.empty:
    morfo_payload = None
    if "df_curva_icg" in st.session_state and "puntos_qbcx_final" in st.session_state:
        morfo_payload = {
            "df_curva": st.session_state["df_curva_icg"],
            "puntos": st.session_state["puntos_qbcx_final"],
            "interpretacion": st.session_state.get("interpretacion_qbcx", ""),
            "png": st.session_state.get("qbcx_png"),
            "png_canales": st.session_state.get("canales_zlogic_png"),
        }
    try:
        pdf = generar_pdf_informe(df_final, morfo_payload)
        basal, _ = seleccionar_basal_y_parado(df_final)
        nombre_paciente = limpiar_nombre_archivo((basal or {}).get("paciente", "PACIENTE"))
        st.download_button(
            "Descargar PDF morfologico validado",
            data=pdf,
            file_name=f"{nombre_paciente} ANALISIS MORFOLOGICO ICG VALIDADO.pdf",
            mime="application/pdf",
        )
        if "df_metricas_hemodinamicas_digitalizacion" in st.session_state:
            df_multi_export_final = None
            if "canales_zlogic_pdf" in st.session_state:
                canales_final = st.session_state["canales_zlogic_pdf"]
                df_multi_export_final = pd.DataFrame({
                    "tiempo_ms": canales_final["df_dzdt"]["tiempo_ms"],
                    "dzdt": canales_final["df_dzdt"]["amplitud"],
                    "ecg": np.interp(canales_final["df_dzdt"]["tiempo_ms"], canales_final["df_ecg"]["tiempo_ms"], canales_final["df_ecg"]["amplitud"]),
                    "fono": np.interp(canales_final["df_dzdt"]["tiempo_ms"], canales_final["df_fono"]["tiempo_ms"], canales_final["df_fono"]["amplitud"]),
                })
            excel_final = excel_metricas_hemodinamicas_bytes(
                df_metricas=st.session_state.get("df_metricas_hemodinamicas_digitalizacion"),
                df_qbcx=tabla_resultados_qbcx(st.session_state.get("puntos_qbcx_final", {})) if "puntos_qbcx_final" in st.session_state else None,
                df_curva=st.session_state.get("df_curva_icg"),
                df_canales=df_multi_export_final,
                df_final=df_final,
                aprendizaje=st.session_state.get("aprendizaje_cursores", {}),
            )
            st.download_button(
                "Descargar Excel con todas las metricas hemodinamicas",
                data=excel_final,
                file_name=f"{nombre_paciente} metricas hemodinamicas ICG.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
    except Exception as e:
        st.error(f"PDF no disponible. Error real: {e}")
else:
    st.info("Para exportar PDF completo, primero cargar datos CGI/Z-Logic. La curva morfologica se puede analizar y descargar por separado.")
st.markdown("</div>", unsafe_allow_html=True)


# ============================================================
# AYUDA DE REQUIREMENTS
# ============================================================

with st.expander("requirements.txt recomendado"):
    st.code(
        """streamlit
pandas
numpy
matplotlib
pillow
pdfplumber
pypdf
PyMuPDF
reportlab
openpyxl
""",
        language="text",
    )
