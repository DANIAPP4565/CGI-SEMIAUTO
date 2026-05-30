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
            if bool(pm.get("validado_operador", False)):
                fuente_tiempos_vac = "desde cursores morfologicos QRS-B y B-X confirmados por operador"
            else:
                fuente_tiempos_vac = "PRELIMINAR: cursores morfologicos no confirmados visualmente"

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
        estado_ct = "desde curva dZ/dt digitalizada y cursores validados" if bool(p.get("validado_operador", False)) else "PRELIMINAR: requiere confirmacion visual de cursores"
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
    df["Valor recalculado"] = df["Valor recalculado"].apply(lambda x: round(float(x), 4) if limpiar_numero(x) is not None else x)
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
        texto = normalizar_txt(str(r.get("Variable", "")) + " " + str(r.get("Estado", "")) + " " + str(r.get("Ecuacion", "")))
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
        st.dataframe(df_der, use_container_width=True, height=560)
    with tabs[1]:
        df_t = _filtrar_derivadas_por_palabras(df_der, ["pep", "lvet", "str", "cts", "fe capan", "end(avg)", "end(est)"])
        st.dataframe(df_t, use_container_width=True, height=360)
    with tabs[2]:
        df_v = _filtrar_derivadas_por_palabras(df_der, ["ea", "ees", "acoplamiento", "chen", "vac"])
        st.dataframe(df_v, use_container_width=True, height=360)
    with tabs[3]:
        df_i = _filtrar_derivadas_por_palabras(df_der, ["bsa", "bmi", "map", "co", "ci", "si/ids", "svr", "rvs", "svri", "irv", "cft"])
        st.dataframe(df_i, use_container_width=True, height=420)
    with tabs[4]:
        df_c = _filtrar_derivadas_por_palabras(df_der, ["contractilidad", "heather", "velocidad", "aceleracion", "pendiente", "area sistolica", "simetria", "dz/dt"])
        st.dataframe(df_c, use_container_width=True, height=460)


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
            raise ValueError(f"No se detectó trazo suficiente en canal {nombre_canal}. Revisar banda/crop.")

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
        raise ValueError(f"No se pudo seguir la curva del canal {nombre_canal}. Ajustar recorte/banda.")

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
    df_dz = _digitalizar_banda_trazo(
        arr, 0.02, 0.38, rr, umbral_oscuridad=155, suavizado=suavizado,
        invertir=invertir_dzdt, max_salto_px=max(6, int(h*0.06)),
        amplitud_mm_por_px=mm_por_px, escala_y_por_mm=escala_dzdt,
        preferir_trazo_azul=True, excluir_relleno_amarillo=True, nombre_canal="dzdt"
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
    y_fono = norm(df_fono["amplitud"].to_numpy(float), 0.95 * float(amplificar_fono)) + base_fono

    fig_h = 7.2 if expandir_vertical else 5.2
    fig, ax = plt.subplots(figsize=(13.5, fig_h))

    ax.plot(t, y_dz, linewidth=2.1, color="#1f77b4", label="dZ/dt digitalizada")
    ax.plot(df_ecg["tiempo_ms"], y_ecg, linewidth=1.6, color="#2ca02c", label="ECG digitalizado")
    ax.plot(df_fono["tiempo_ms"], y_fono, linewidth=1.7, color="#ff7f0e", label="Fonocardiograma digitalizado")

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
    ax.set_title("Digitalización automática multicanal: dZ/dt + ECG verde medio + Fonocardiograma naranja inferior")
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


# ============================================================
# PDF
# ============================================================

def generar_pdf_informe(df: pd.DataFrame, morfo: Optional[Dict[str, Any]] = None) -> bytes:
    if A4 is None:
        raise RuntimeError("ReportLab no esta instalado. Agregar reportlab a requirements.txt")
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=1.4*cm, leftMargin=1.4*cm, topMargin=1.3*cm, bottomMargin=1.2*cm)
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name="TituloAzul", parent=styles["Title"], fontSize=16, leading=20, textColor=colors.HexColor("#082F49"), spaceAfter=8))
    styles.add(ParagraphStyle(name="Sub", parent=styles["Heading2"], fontSize=12, leading=15, textColor=colors.HexColor("#0B4F7A"), spaceBefore=8, spaceAfter=6))
    styles.add(ParagraphStyle(name="NormalSmall", parent=styles["BodyText"], fontSize=8.4, leading=10.5))
    elems: List[Any] = []
    elems.append(Paragraph(safe_pdf_text(TITULO_APP), styles["TituloAzul"]))
    elems.append(Paragraph(safe_pdf_text(AUTOR_APP), styles["NormalSmall"]))
    elems.append(Spacer(1, 8))

    basal, parado = seleccionar_basal_y_parado(df)
    datos = basal or {}
    paciente = datos.get("paciente", "No disponible")
    elems.append(Paragraph("Datos del paciente y condiciones del estudio", styles["Sub"]))
    tabla_datos = [
        ["Paciente", safe_pdf_text(paciente), "Fecha", safe_pdf_text(datos.get("fecha_estudio", "No disponible"))],
        ["Edad", safe_pdf_text(fmt(datos.get("edad"),0)), "Sexo", safe_pdf_text(datos.get("sexo", "No disponible"))],
        ["PAS/PAD", f"{fmt(datos.get('pas'),0)} / {fmt(datos.get('pad'),0)} mmHg", "FC", f"{fmt(datos.get('fc'),0)} lpm"],
        ["Condicion basal", safe_pdf_text(datos.get("posicion", "ESTUDIO BASAL")), "Archivo", safe_pdf_text(datos.get("archivo_origen", ""))],
    ]
    table = Table(tabla_datos, colWidths=[3*cm, 5*cm, 2.5*cm, 6*cm])
    table.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#EAF6FF")),
        ("GRID", (0,0), (-1,-1), .35, colors.HexColor("#D7E3EE")),
        ("FONTNAME", (0,0), (-1,-1), "Helvetica"),
        ("FONTSIZE", (0,0), (-1,-1), 8),
        ("VALIGN", (0,0), (-1,-1), "TOP"),
    ]))
    elems.append(table)

    elems.append(Paragraph("Informe hemodinamico integrado", styles["Sub"]))
    for par in generar_informe_texto(df).split("\n"):
        if par.strip():
            elems.append(Paragraph(safe_pdf_text(par), styles["NormalSmall"]))
            elems.append(Spacer(1, 2))

    # Tabla de variables disponibles
    elems.append(Paragraph("Variables hemodinamicas extraidas", styles["Sub"]))
    mostrar = ["ic", "gc", "irv", "rvs", "ca", "cft", "cftnr", "iv", "iac", "ih", "cts", "ea", "ees", "ava", "ds", "ids", "pep", "lvet", "rr", "dzdt_max", "z0"]
    filas = [["Variable", "Valor basal"]]
    for k in mostrar:
        filas.append([k.upper(), fmt(datos.get(k), 2 if k not in {"irv", "rvs", "fc"} else 0)])
    tvar = Table(filas, colWidths=[5*cm, 5*cm])
    tvar.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#082F49")),
        ("TEXTCOLOR", (0,0), (-1,0), colors.white),
        ("GRID", (0,0), (-1,-1), .3, colors.HexColor("#D7E3EE")),
        ("FONTSIZE", (0,0), (-1,-1), 8),
    ]))
    elems.append(tvar)

    # Tabla de variables recalculadas por ecuaciones integradas
    df_derivadas_pdf = calcular_variables_icg_derivadas(datos, morfo)
    elems.append(Paragraph("Variables recalculadas por ecuaciones ICG integradas", styles["Sub"]))
    if df_derivadas_pdf is not None and not df_derivadas_pdf.empty:
        data_der = [["Variable", "Valor", "Unidad", "Estado"]]
        for _, rr in df_derivadas_pdf.iterrows():
            data_der.append([
                safe_pdf_text(rr.get("Variable", "")),
                safe_pdf_text(rr.get("Valor recalculado", "")),
                safe_pdf_text(rr.get("Unidad", "")),
                safe_pdf_text(rr.get("Estado", "")),
            ])
        tder = Table(data_der, colWidths=[4.2*cm, 3.0*cm, 3.2*cm, 6.4*cm])
        tder.setStyle(TableStyle([
            ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#0B4F7A")),
            ("TEXTCOLOR", (0,0), (-1,0), colors.white),
            ("GRID", (0,0), (-1,-1), .3, colors.HexColor("#D7E3EE")),
            ("VALIGN", (0,0), (-1,-1), "TOP"),
            ("FONTSIZE", (0,0), (-1,-1), 7.2),
        ]))
        elems.append(tder)
        elems.append(Paragraph("Nota: las variables recalculadas son controles internos de la app. Cuando el fabricante use algoritmos propietarios, el valor importado del equipo debe mantenerse separado del valor recalculado.", styles["NormalSmall"]))
    else:
        elems.append(Paragraph("No se recalcularon variables por falta de datos suficientes.", styles["NormalSmall"]))

    if morfo is not None and morfo.get("df_curva") is not None and morfo.get("puntos") is not None:
        elems.append(PageBreak())
        elems.append(Paragraph("Analisis morfologico semiautomatico de señal ICG", styles["TituloAzul"]))
        elems.append(Paragraph(safe_pdf_text(morfo.get("interpretacion", "")), styles["NormalSmall"]))
        elems.append(Spacer(1, 8))
        png_multi = morfo.get("png_canales")
        if png_multi:
            elems.append(Paragraph("Digitalizacion automatica dZ/dt + ECG + fonocardiograma", styles["Sub"]))
            img_buf_multi = io.BytesIO(png_multi)
            elems.append(RLImage(img_buf_multi, width=17.5*cm, height=7.2*cm))
            elems.append(Spacer(1, 8))
        png = morfo.get("png")
        if png:
            elems.append(Paragraph("Cursores QRS-B-C-X sobre dZ/dt", styles["Sub"]))
            img_buf = io.BytesIO(png)
            elems.append(RLImage(img_buf, width=17.5*cm, height=7.2*cm))
            elems.append(Spacer(1, 8))
        df_tab = tabla_resultados_qbcx(morfo["puntos"])
        data = [df_tab.columns.tolist()] + df_tab.values.tolist()
        data = [[safe_pdf_text(c) for c in row] for row in data]
        tt = Table(data, colWidths=[4.2*cm, 3.4*cm, 9.4*cm])
        tt.setStyle(TableStyle([
            ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#082F49")),
            ("TEXTCOLOR", (0,0), (-1,0), colors.white),
            ("GRID", (0,0), (-1,-1), .3, colors.HexColor("#D7E3EE")),
            ("VALIGN", (0,0), (-1,-1), "TOP"),
            ("FONTSIZE", (0,0), (-1,-1), 7.4),
        ]))
        elems.append(tt)
        elems.append(Spacer(1, 8))

        df_der_morfo_pdf = calcular_variables_icg_derivadas(datos, morfo)
        if df_der_morfo_pdf is not None and not df_der_morfo_pdf.empty:
            elems.append(Paragraph("Metricas hemodinamicas calculadas desde la digitalizacion corregida", styles["Sub"]))
            data_m = [["Variable", "Valor", "Unidad", "Estado"]]
            for _, rr_m in df_der_morfo_pdf.iterrows():
                data_m.append([
                    safe_pdf_text(rr_m.get("Variable", "")),
                    safe_pdf_text(rr_m.get("Valor recalculado", "")),
                    safe_pdf_text(rr_m.get("Unidad", "")),
                    safe_pdf_text(rr_m.get("Estado", "")),
                ])
            tmet = Table(data_m, colWidths=[4.2*cm, 2.6*cm, 3.1*cm, 7.0*cm])
            tmet.setStyle(TableStyle([
                ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#0B4F7A")),
                ("TEXTCOLOR", (0,0), (-1,0), colors.white),
                ("GRID", (0,0), (-1,-1), .3, colors.HexColor("#D7E3EE")),
                ("VALIGN", (0,0), (-1,-1), "TOP"),
                ("FONTSIZE", (0,0), (-1,-1), 6.8),
            ]))
            elems.append(tmet)
            elems.append(Spacer(1, 8))

        vars_ct_pdf = calcular_contractilidad_morfologica(morfo.get("df_curva"), morfo.get("puntos"), z0=limpiar_numero(datos.get("z0")))
        interp_ct_pdf = interpretar_contractilidad_morfologica(vars_ct_pdf)
        if vars_ct_pdf.get("dzdt_max_rel") is not None:
            elems.append(Paragraph("Contractilidad morfologica relativa", styles["Sub"]))
            filas_ct = [["Variable", "Valor"]]
            for etiqueta, clave in [
                ("dZ/dt maximo relativo", "dzdt_max_rel"),
                ("Indice de Heather relativo", "hi_rel"),
                ("Indice de velocidad relativo", "vi_rel"),
                ("Indice de aceleracion relativo", "aci_rel"),
                ("Pendiente ascendente B-C", "pendiente_ascendente"),
                ("Pendiente descendente C-X", "pendiente_descendente"),
                ("Area sistolica relativa B-X", "area_sistolica_rel"),
                ("Simetria sistolica", "simetria_sistolica"),
            ]:
                val = vars_ct_pdf.get(clave)
                if val is not None:
                    filas_ct.append([etiqueta, safe_pdf_text(round(float(val), 4))])
            tct = Table(filas_ct, colWidths=[7.0*cm, 5.0*cm])
            tct.setStyle(TableStyle([
                ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#0B4F7A")),
                ("TEXTCOLOR", (0,0), (-1,0), colors.white),
                ("GRID", (0,0), (-1,-1), .3, colors.HexColor("#D7E3EE")),
                ("FONTSIZE", (0,0), (-1,-1), 7.4),
            ]))
            elems.append(tct)
            elems.append(Paragraph(safe_pdf_text(interp_ct_pdf), styles["NormalSmall"]))
            elems.append(Spacer(1, 6))
        elems.append(Paragraph("Nota: el modulo es semiautomatico; la aceptacion final de cursores corresponde al operador medico. Las variables morfologicas de contractilidad son relativas y no sustituyen los indices absolutos del equipo sin calibracion vertical real de dZ/dt.", styles["NormalSmall"]))

    doc.build(elems)
    buffer.seek(0)
    return buffer.getvalue()


# ============================================================
# INTERFAZ
# ============================================================

st.markdown("<div class='card'>", unsafe_allow_html=True)
st.subheader("1. Carga de estudios CGI / Z-Logic")
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
        st.dataframe(df_derivadas, use_container_width=True)
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

    validado_operador = st.checkbox(
        "Confirmo visualmente que QRS, B, C y X corresponden a la señal dZ/dt original",
        value=bool(st.session_state.get("qbcx_validado_operador", False)),
        help="Sin esta confirmacion, Ea/Ees y FE Capan derivados de los cursores deben considerarse preliminares."
    )

    pfinal = recalcular_puntos_manuales(df_curva, qms, bms, cms, xms)
    pfinal["validado_operador"] = bool(validado_operador)
    if bool(validado_operador) and not bool(st.session_state.get("qbcx_validado_operador", False)):
        actualizar_aprendizaje_cursores(pauto, pfinal, True)
    st.session_state["qbcx_validado_operador"] = bool(validado_operador)

    if not validado_operador:
        pfinal.setdefault("alertas", []).append("Cursores no confirmados visualmente por el operador; calculos morfologicos preliminares.")
        pfinal["confianza"] = "Preliminar - requiere confirmacion visual"
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

    cols_dl = st.columns(4)
    cols_dl[0].download_button("Descargar dZ/dt CSV", df_curva.to_csv(index=False).encode("utf-8"), "curva_dzdt_icg_digitalizada.csv", "text/csv")
    cols_dl[1].download_button("Descargar cursores CSV", tab_qbcx.to_csv(index=False).encode("utf-8"), "cursores_qbcx_icg.csv", "text/csv")
    cols_dl[2].download_button("Descargar grafico QBCX PNG", png_bytes, "grafico_qbcx_icg.png", "image/png")
    if "canales_zlogic_pdf" in st.session_state:
        canales = st.session_state["canales_zlogic_pdf"]
        df_multi = pd.DataFrame({
            "tiempo_ms": canales["df_dzdt"]["tiempo_ms"],
            "dzdt": canales["df_dzdt"]["amplitud"],
            "ecg": np.interp(canales["df_dzdt"]["tiempo_ms"], canales["df_ecg"]["tiempo_ms"], canales["df_ecg"]["amplitud"]),
            "fono": np.interp(canales["df_dzdt"]["tiempo_ms"], canales["df_fono"]["tiempo_ms"], canales["df_fono"]["amplitud"]),
        })
        cols_dl[3].download_button("Descargar dZ/dt+ECG+Fono CSV", df_multi.to_csv(index=False).encode("utf-8"), "canales_zlogic_digitalizados.csv", "text/csv")

st.markdown("</div>", unsafe_allow_html=True)


# ============================================================
# DESCARGA PDF FINAL
# ============================================================

st.markdown("<div class='card'>", unsafe_allow_html=True)
st.subheader("6. Exportacion")
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
            "Descargar PDF completo integrado",
            data=pdf,
            file_name=f"{nombre_paciente} CGI ICG MORFOLOGIA.pdf",
            mime="application/pdf",
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
