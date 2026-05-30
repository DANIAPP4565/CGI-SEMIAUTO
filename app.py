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
}

RANGOS: Dict[str, Tuple[float, float]] = {
    "edad": (0, 120), "peso": (25, 250), "talla": (90, 230), "imc": (10, 80),
    "pas": (60, 260), "pad": (30, 160), "fc": (35, 190),
    "ic": (0.8, 8.0), "gc": (1.0, 25.0), "irv": (700, 7000), "rvs": (300, 4500),
    "ca": (0.1, 10.0), "cft": (5, 120), "cftnr": (1, 220),
    "iv": (0, 200), "iac": (0, 80), "ih": (0, 80), "cts": (0.05, 1.5),
    "ea": (0.1, 10), "ees": (0.1, 25), "ava": (0.1, 5),
    "ds": (10, 250), "ids": (5, 150), "pep": (40, 180), "lvet": (150, 500), "z0": (5, 80),
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
    for clave in ["cftnr", "irv", "rvs", "iac", "ih", "cts", "ava", "ic", "gc", "ca", "cft", "iv", "ea", "ees", "ds", "ids", "pas", "pad", "fc", "pep", "lvet", "z0", "paciente", "dni", "edad", "sexo", "peso", "talla", "imc", "superficie_corporal", "fecha_estudio", "diagnostico", "posicion"]:
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
    area_sis = float(np.trapz(y[seg_bx] - yb, t[seg_bx]))
    area_asc = float(np.trapz(np.maximum(y[seg_bc] - yb, 0), t[seg_bc]))
    area_desc = float(np.trapz(np.maximum(y[seg_cx] - yx, 0), t[seg_cx]))

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
    mostrar = ["ic", "gc", "irv", "rvs", "ca", "cft", "cftnr", "iv", "iac", "ih", "cts", "ea", "ees", "ava", "ds", "ids", "pep", "lvet", "z0"]
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
        png = morfo.get("png")
        if png:
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
            dfx = leer_archivo(f)
            if not dfx.empty:
                partes.append(dfx)
            else:
                errores.append(f"{f.name}: sin datos estructurados detectables")
        except Exception as e:
            errores.append(f"{f.name}: {e}")
    if partes:
        st.session_state["df_final"] = pd.concat(partes, ignore_index=True)
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
    "Carga una captura o imagen de la señal promedio dZ/dt del cardiógrafo. "
    "La app digitaliza la curva, propone QRS-B-C-X y permite corregir manualmente los cursores."
)

img_file = st.file_uploader("Imagen/captura de curva promedio ICG", type=["png", "jpg", "jpeg"], key="img_icg")

if img_file is not None:
    imagen = Image.open(img_file).convert("RGB")
    ancho, alto = imagen.size
    st.image(imagen, caption="Imagen original", use_container_width=True)

    st.warning(
        "Para evitar curvas deformadas, recortar solo la señal dZ/dt promedio. "
        "No incluir ECG, fonocardiograma, carotidograma, texto, números, bordes ni más de una curva."
    )

    fc_det = obtener_fc_desde_df(df_final)
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
    st.session_state["qbcx_validado_operador"] = bool(validado_operador)

    pfinal = recalcular_puntos_manuales(df_curva, qms, bms, cms, xms)
    pfinal["validado_operador"] = bool(validado_operador)
    if not validado_operador:
        pfinal.setdefault("alertas", []).append("Cursores no confirmados visualmente por el operador; calculos morfologicos preliminares.")
        pfinal["confianza"] = "Preliminar - requiere confirmacion visual"
    st.session_state["puntos_qbcx_final"] = pfinal
    fig = graficar_curva_icg_con_puntos(df_curva, pfinal)
    st.pyplot(fig, use_container_width=True)
    png_bytes = fig_to_png_bytes(fig)
    st.session_state["qbcx_png"] = png_bytes
    tab_qbcx = tabla_resultados_qbcx(pfinal)
    st.dataframe(tab_qbcx, use_container_width=True)
    interpretacion = texto_interpretacion_qbcx(pfinal)
    st.info(interpretacion)
    st.session_state["interpretacion_qbcx"] = interpretacion

    if not df_final.empty:
        basal_para_morfo, _ = seleccionar_basal_y_parado(df_final)
        df_der_morfo = calcular_variables_icg_derivadas(basal_para_morfo or {}, {"puntos": pfinal, "df_curva": df_curva})
        st.markdown("### Variables recalculadas integrando cursores y contractilidad morfologica")
        if not df_der_morfo.empty:
            st.dataframe(df_der_morfo, use_container_width=True)
        else:
            st.info("No se recalcularon variables adicionales por falta de datos suficientes.")

    coldl1, coldl2, coldl3 = st.columns(3)
    coldl1.download_button("Descargar curva CSV", df_curva.to_csv(index=False).encode("utf-8"), "curva_icg_digitalizada.csv", "text/csv")
    coldl2.download_button("Descargar cursores CSV", tab_qbcx.to_csv(index=False).encode("utf-8"), "cursores_qbcx_icg.csv", "text/csv")
    coldl3.download_button("Descargar grafico PNG", png_bytes, "grafico_qbcx_icg.png", "image/png")

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
