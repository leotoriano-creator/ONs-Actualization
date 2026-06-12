# -*- coding: utf-8 -*-
"""
Descarga variables para la hoja VARIABLES de Base ONs:
    - BADLAR TNA desde API Alquimia
    - TAMAR TNA desde API Alquimia
    - COM3500 / A3500 desde BCRA

Salida:
    DataStorage/variables_mercado.xlsx

Hojas:
    - BADLAR
    - TAMAR
    - COM3500
    - VARIABLES_CONSOLIDADO

Requisitos:
    pip install pandas requests openpyxl xlrd
"""

from __future__ import annotations

import re
import sys
from pathlib import Path
from datetime import datetime

import pandas as pd
import requests


# ============================================================
# CONFIG
# ============================================================

TASAS_URL = "https://api.alquimiaconsultora.com/api/monetario/tasas-mercado"
COM3500_URL = "https://www.bcra.gob.ar/Pdfs/PublicacionesEstadisticas/com3500.xls"

SCRIPT_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = SCRIPT_DIR.parent if SCRIPT_DIR.name.lower() == "getdata" else Path.cwd()

DATA_STORAGE = PROJECT_ROOT / "DataStorage"
OUTPUT_FILE = DATA_STORAGE / "variables_mercado.xlsx"

TIMEOUT_SECONDS = 90


# ============================================================
# LOG
# ============================================================

def log(msg: str) -> None:
    ahora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"{ahora} | {msg}")


# ============================================================
# HELPERS
# ============================================================

def normalizar_texto(x) -> str:
    if x is None:
        return ""

    try:
        if pd.isna(x):
            return ""
    except Exception:
        pass

    return str(x).strip()


def normalizar_columna(col) -> str:
    s = str(col).strip().lower()
    s = (
        s.replace("á", "a")
         .replace("é", "e")
         .replace("í", "i")
         .replace("ó", "o")
         .replace("ú", "u")
         .replace("ñ", "n")
    )
    s = re.sub(r"[^a-z0-9]+", "_", s)
    s = re.sub(r"_+", "_", s).strip("_")
    return s


def normalizar_fecha(x):
    if x is None:
        return pd.NaT

    try:
        if pd.isna(x):
            return pd.NaT
    except Exception:
        pass

    # Primero formato ISO, típico del endpoint: YYYY-MM-DD.
    dt = pd.to_datetime(x, errors="coerce", dayfirst=False)

    if pd.isna(dt):
        # Fallback para fechas tipo dd/mm/yyyy del BCRA.
        dt = pd.to_datetime(x, errors="coerce", dayfirst=True)

    if pd.isna(dt):
        return pd.NaT

    return dt.normalize()


def normalizar_tasa_pct(x):
    """
    Devuelve tasa en puntos porcentuales.
    Ejemplo:
        34.3 -> 34.3
        0.343 -> 34.3
    """
    if x is None:
        return pd.NA

    try:
        if pd.isna(x):
            return pd.NA
    except Exception:
        pass

    if isinstance(x, str):
        s = x.strip().replace("%", "").replace(",", ".")
        s = re.sub(r"[^\d.\-+]", "", s)
    else:
        s = x

    try:
        val = float(s)
    except Exception:
        return pd.NA

    if abs(val) <= 1:
        val = val * 100

    return val


def normalizar_numero(x):
    if x is None:
        return pd.NA

    try:
        if pd.isna(x):
            return pd.NA
    except Exception:
        pass

    if isinstance(x, str):
        s = x.strip().replace(",", ".")
        s = re.sub(r"[^\d.\-+]", "", s)
    else:
        s = x

    try:
        return float(s)
    except Exception:
        return pd.NA


# ============================================================
# TASAS API ALQUIMIA
# ============================================================

def descargar_json_tasas() -> list[dict]:
    headers = {
        "User-Agent": "Mozilla/5.0",
        "Accept": "application/json, text/plain, */*",
    }

    log(f"Descargando tasas desde: {TASAS_URL}")
    response = requests.get(TASAS_URL, headers=headers, timeout=TIMEOUT_SECONDS)
    response.raise_for_status()

    data = response.json()

    if isinstance(data, list):
        return data

    if isinstance(data, dict):
        for key in ["rows", "data", "result", "results", "items"]:
            if key in data and isinstance(data[key], list):
                return data[key]

    raise ValueError("Formato JSON no reconocido para tasas-mercado.")


def preparar_df_tasas_largo() -> pd.DataFrame:
    data = descargar_json_tasas()
    df = pd.DataFrame(data)

    if df.empty:
        raise ValueError("La API de tasas-mercado devolvió datos vacíos.")

    required = ["Date", "Category 1", "Category 2", "Category 3", "Category 4", "Value"]
    missing = [c for c in required if c not in df.columns]

    if missing:
        raise ValueError(
            f"Faltan columnas esperadas en tasas-mercado: {missing}. "
            f"Columnas disponibles: {list(df.columns)}"
        )

    df = df[required].copy()

    df["Date"] = df["Date"].apply(normalizar_fecha)
    df["Value"] = df["Value"].apply(normalizar_tasa_pct)

    for col in ["Category 1", "Category 2", "Category 3", "Category 4"]:
        df[col] = df[col].astype(str).str.strip()

    df = df.dropna(subset=["Date", "Value"]).copy()

    log(f"Filas válidas del endpoint de tasas: {len(df):,}")
    return df


def filtrar_serie_tasa(df_largo: pd.DataFrame, nombre_serie: str) -> pd.DataFrame:
    """
    Extrae una serie TNA desde el endpoint largo.

    Para BADLAR:
        Category 1 = BADLAR
        Category 2 = Pesos
        Category 4 = TNA
        Prioridad Category 3 = Bancos privados

    Para TAMAR:
        Category 1 = TAMAR
        Category 4 = TNA
        Si hay varias variantes, prioriza Bancos privados / Total de forma robusta.
    """
    serie_upper = nombre_serie.upper()

    tmp = df_largo[
        (df_largo["Category 1"].str.upper() == serie_upper) &
        (df_largo["Category 4"].str.upper() == "TNA")
    ].copy()

    if serie_upper == "BADLAR":
        tmp = tmp[tmp["Category 2"].str.upper() == "PESOS"].copy()

    if tmp.empty:
        disponibles = (
            df_largo[["Category 1", "Category 2", "Category 3", "Category 4"]]
            .drop_duplicates()
            .sort_values(["Category 1", "Category 2", "Category 3", "Category 4"])
            .head(80)
        )
        raise ValueError(
            f"No encontré datos para {nombre_serie} TNA. "
            f"Primeras combinaciones disponibles:\n{disponibles.to_string(index=False)}"
        )

    # Priorizamos Bancos privados si existe, que es la BADLAR que se ve en la base.
    tmp["prioridad"] = 99

    cat3_upper = tmp["Category 3"].str.upper()
    tmp.loc[cat3_upper == "BANCOS PRIVADOS", "prioridad"] = 1
    tmp.loc[cat3_upper == "TOTAL", "prioridad"] = 2

    # Para TAMAR, por si viniera con otra etiqueta, dejamos lo que haya después de ordenar.
    tmp = tmp.sort_values(["Date", "prioridad"]).copy()
    tmp = tmp.drop_duplicates(subset=["Date"], keep="first")

    out = tmp[["Date", "Value", "Category 2", "Category 3"]].copy()
    out.columns = ["fecha", f"{serie_upper}_TNA", "categoria_2", "categoria_3"]

    out = out.sort_values("fecha").reset_index(drop=True)

    log(f"{serie_upper} filas válidas: {len(out):,}")
    if not out.empty:
        log(f"{serie_upper} última fecha: {out['fecha'].max().strftime('%Y-%m-%d')}")
        log(f"{serie_upper} última tasa: {out[f'{serie_upper}_TNA'].iloc[-1]}")

    return out


def descargar_badlar_tamar() -> tuple[pd.DataFrame, pd.DataFrame]:
    df_largo = preparar_df_tasas_largo()

    df_badlar = filtrar_serie_tasa(df_largo, "BADLAR")
    df_tamar = filtrar_serie_tasa(df_largo, "TAMAR")

    return df_badlar, df_tamar


# ============================================================
# COM3500 / A3500
# ============================================================

def leer_xls_bcra(url: str) -> pd.DataFrame:
    """
    Lee el XLS de BCRA.
    Requiere:
        pip install xlrd
    """
    log(f"Descargando COM3500 desde: {url}")

    headers = {"User-Agent": "Mozilla/5.0"}
    response = requests.get(url, headers=headers, timeout=TIMEOUT_SECONDS)
    response.raise_for_status()

    return pd.read_excel(response.content, header=None)


def detectar_tabla_com3500(raw: pd.DataFrame) -> pd.DataFrame:
    """
    El archivo com3500.xls puede traer títulos antes de la tabla.
    Buscamos filas con fecha válida y número plausible de TC.
    """
    registros = []

    for _, row in raw.iterrows():
        valores = list(row.values)

        fecha = pd.NaT

        # Fecha válida.
        for v in valores:
            dt = normalizar_fecha(v)
            if not pd.isna(dt) and 1990 <= dt.year <= 2100:
                fecha = dt
                break

        if pd.isna(fecha):
            continue

        nums = []
        for v in valores:
            n = normalizar_numero(v)
            if not pd.isna(n):
                nums.append(float(n))

        # TC histórico amplio.
        plausibles = [n for n in nums if 0.1 <= n <= 100000]

        if not plausibles:
            continue

        # Normalmente el TC es el último número relevante de la fila.
        tc = plausibles[-1]
        registros.append({"fecha": fecha, "COM3500": tc})

    out = pd.DataFrame(registros)

    if out.empty:
        raise ValueError("No pude parsear datos válidos del XLS de COM3500.")

    out = out.dropna(subset=["fecha", "COM3500"]).copy()
    out = out.sort_values("fecha").drop_duplicates(subset=["fecha"], keep="last").reset_index(drop=True)

    log(f"COM3500 filas válidas: {len(out):,}")
    if not out.empty:
        log(f"COM3500 última fecha: {out['fecha'].max().strftime('%Y-%m-%d')}")
        log(f"COM3500 último valor: {out['COM3500'].iloc[-1]}")

    return out


def descargar_com3500() -> pd.DataFrame:
    raw = leer_xls_bcra(COM3500_URL)
    return detectar_tabla_com3500(raw)


# ============================================================
# CONSOLIDADO
# ============================================================

def armar_consolidado(df_badlar: pd.DataFrame, df_tamar: pd.DataFrame, df_com3500: pd.DataFrame) -> pd.DataFrame:
    badlar = df_badlar[["fecha", "BADLAR_TNA"]].copy()
    tamar = df_tamar[["fecha", "TAMAR_TNA"]].copy()
    com = df_com3500[["fecha", "COM3500"]].copy()

    df = pd.merge(badlar, tamar, on="fecha", how="outer")
    df = pd.merge(df, com, on="fecha", how="outer")

    df = df.sort_values("fecha").reset_index(drop=True)

    # Forward-fill para poder cubrir feriados/fines de semana cuando haga falta.
    df["BADLAR_TNA_FFILL"] = df["BADLAR_TNA"].ffill()
    df["TAMAR_TNA_FFILL"] = df["TAMAR_TNA"].ffill()
    df["COM3500_FFILL"] = df["COM3500"].ffill()

    return df


def exportar_excel(
    df_badlar: pd.DataFrame,
    df_tamar: pd.DataFrame,
    df_com3500: pd.DataFrame,
    df_consolidado: pd.DataFrame,
) -> None:
    DATA_STORAGE.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        df_badlar.to_excel(writer, sheet_name="BADLAR", index=False)
        df_tamar.to_excel(writer, sheet_name="TAMAR", index=False)
        df_com3500.to_excel(writer, sheet_name="COM3500", index=False)
        df_consolidado.to_excel(writer, sheet_name="VARIABLES_CONSOLIDADO", index=False)

        wb = writer.book

        for ws in wb.worksheets:
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions

            for col in ["A", "B", "C", "D", "E", "F", "G"]:
                ws.column_dimensions[col].width = 20

            for cell in ws["A"][1:]:
                cell.number_format = "dd/mm/yyyy"

            for col in ["B", "C", "D", "E", "F", "G"]:
                for cell in ws[col][1:]:
                    cell.number_format = "0.0000"

    log(f"Excel exportado: {OUTPUT_FILE}")


# ============================================================
# MAIN
# ============================================================

def main() -> int:
    try:
        df_badlar, df_tamar = descargar_badlar_tamar()
        df_com3500 = descargar_com3500()

        df_consolidado = armar_consolidado(
            df_badlar=df_badlar,
            df_tamar=df_tamar,
            df_com3500=df_com3500,
        )

        exportar_excel(
            df_badlar=df_badlar,
            df_tamar=df_tamar,
            df_com3500=df_com3500,
            df_consolidado=df_consolidado,
        )

        log("Proceso finalizado OK.")
        return 0

    except Exception as e:
        log(f"ERROR: {repr(e)}")
        return 1


if __name__ == "__main__":
    sys.exit(main())
