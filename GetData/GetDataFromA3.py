# -*- coding: utf-8 -*-
"""
Descarga colocaciones A3 - Mercado primario
Filtro: corporativas / ON privadas

Uso:
    1) Ajustar FECHA_DESDE si hace falta
    2) Ejecutar:
        python bajar_a3_on_corporativas.py

Salida:
    output/a3_on_corporativas.xlsx

Requisitos:
    pip install pandas requests openpyxl
"""

from __future__ import annotations

import re
import sys
import json
from pathlib import Path
from datetime import datetime, date
from urllib.parse import quote

import pandas as pd
import requests


# ============================================================
# CONFIG
# ============================================================

# Fecha inicial de descarga.
# La fecha final se calcula automáticamente como hoy.
FECHA_DESDE = "2026-05-12"

BASE_URL = "https://api.marketdata.mae.com.ar/api/mercado/licitacionesporestado/Todos"

OUTPUT_DIR = Path("DataStorage")
OUTPUT_FILE = OUTPUT_DIR / "a3_on_corporativas.xlsx"

TIMEOUT_SECONDS = 45

# Si querés dejar afuera sólo FF pero incluir variantes ON/VCP, dejalo en True.
INCLUIR_ON_VCP = True


# ============================================================
# HELPERS
# ============================================================

def log(msg: str) -> None:
    ahora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"{ahora} | {msg}")


def normalizar_texto(x) -> str:
    if x is None:
        return ""
    return str(x).strip()


def construir_endpoint(fecha_desde: str = FECHA_DESDE, fecha_hasta: str | None = None) -> str:
    """
    Construye el endpoint de A3/MAE con fechaHasta dinámica.

    Ejemplo final:
        .../Todos?oTitulo={%22estado%22:%22H%22,%22fechaDesde%22:%222026-05-12%22,%22fechaHasta%22:%222026-06-12%22}

    fecha_hasta:
        - Si viene None, usa la fecha de hoy del sistema.
        - Formato esperado: YYYY-MM-DD.
    """
    if fecha_hasta is None:
        fecha_hasta = date.today().strftime("%Y-%m-%d")

    payload = {
        "estado": "H",
        "fechaDesde": fecha_desde,
        "fechaHasta": fecha_hasta,
    }

    # compact JSON: sin espacios, como suele esperar este tipo de endpoint.
    payload_json = json.dumps(payload, ensure_ascii=False, separators=(",", ":"))

    # quote con safe='{}' para dejar las llaves visibles y encodear comillas como %22.
    payload_encoded = quote(payload_json, safe="{}")

    return f"{BASE_URL}?oTitulo={payload_encoded}"


def descargar_json(endpoint_url: str) -> list[dict]:
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/120.0 Safari/537.36"
        ),
        "Accept": "application/json, text/plain, */*",
    }

    log(f"Descargando endpoint: {endpoint_url}")
    response = requests.get(endpoint_url, headers=headers, timeout=TIMEOUT_SECONDS)
    response.raise_for_status()

    data = response.json()

    if isinstance(data, dict):
        # Por si el endpoint devuelve {"data": [...]} o similar.
        for key in ["data", "items", "result", "results"]:
            if key in data and isinstance(data[key], list):
                return data[key]

        raise ValueError(
            "El endpoint devolvió un JSON tipo dict, pero no encontré una lista "
            "en keys típicas: data/items/result/results."
        )

    if not isinstance(data, list):
        raise ValueError(f"Formato inesperado del JSON: {type(data)}")

    return data


def es_on_corporativa(item: dict) -> bool:
    """
    Filtro principal.

    En la muestra de A3 aparecen:
        - tipo = 'Pública' para Tesoro / provincias.
        - tipo = 'Privada - ON' para ON corporativas.
        - tipo = 'Privada - ON/VCP BAJO IMPACTO' para variantes ON/VCP.
        - tipo = 'Privada - FF' para fideicomisos financieros.

    Nos quedamos con privadas que sean ON y excluimos FF.
    """
    tipo = normalizar_texto(item.get("tipo")).upper()
    titulo = normalizar_texto(item.get("titulo")).upper()
    emisor = normalizar_texto(item.get("emisor")).upper()

    es_privada = "PRIVADA" in tipo

    if INCLUIR_ON_VCP:
        es_on = ("ON" in tipo) or titulo.startswith("ON ")
    else:
        es_on = tipo == "PRIVADA - ON" or titulo.startswith("ON ")

    no_es_ff = "FF" not in tipo

    # Doble seguro para sacar sector público si viniera mal clasificado.
    no_es_sector_publico = not any(palabra in emisor for palabra in [
        "MINISTERIO DE ECONOMÍA",
        "MINISTERIO DE ECONOMIA",
        "TESORO NACIONAL",
        "GOBIERNO",
        "PROVINCIA",
        "MUNICIPALIDAD",
    ])

    return es_privada and es_on and no_es_ff and no_es_sector_publico


def parse_fecha_utc_a_fecha(x):
    """
    Convierte strings tipo '2026-05-14T13:00:00Z' a fecha naive.
    Si viene '0001-01-01T00:00:00', lo deja vacío.
    """
    s = normalizar_texto(x)
    if not s or s.startswith("0001-01-01"):
        return pd.NaT

    dt = pd.to_datetime(s, errors="coerce", utc=True)

    if pd.isna(dt):
        return pd.NaT

    # En A3 las fechas vienen con Z. Para Excel conviene dejar date/datetime naive.
    return dt.tz_convert(None)


def extraer_valor_corte_numero(valor_corte: str):
    """
    Extrae el primer número relevante de valor_Corte.
    Ejemplos:
        'TASA DE CORTE: 5.25% N.A.' -> 5.25
        'MARGEN DE CORTE: 2.99%' -> 2.99
        'PRECIO DE CORTE: 0.977275' -> 0.977275
    """
    s = normalizar_texto(valor_corte)
    if not s:
        return pd.NA

    # Busca números con punto o coma decimal.
    matches = re.findall(r"[-+]?\d+(?:[.,]\d+)?", s)
    if not matches:
        return pd.NA

    try:
        return float(matches[0].replace(",", "."))
    except Exception:
        return pd.NA


def clasificar_variable_corte(valor_corte: str, variable_licitar: str):
    s = normalizar_texto(valor_corte).upper()
    v = normalizar_texto(variable_licitar).upper()

    if "MARGEN" in s or "MARGEN" in v:
        return "Margen"
    if "TASA" in s or "TASA" in v:
        return "Tasa"
    if "PRECIO" in s or "PRECIO" in v:
        return "Precio"
    if "TEM" in s or "TEM" in v:
        return "TEM"

    return v.title() if v else ""


def limpiar_dataframe(data_filtrada: list[dict]) -> pd.DataFrame:
    df = pd.DataFrame(data_filtrada)

    if df.empty:
        return df

    # Aseguramos columnas aunque algún día falten.
    columnas_base = [
        "id",
        "fechaInicio",
        "fechaFin",
        "fechaLiquidacion",
        "fechaVencimiento",
        "fechaVencimientoEspecie",
        "titulo",
        "emisor",
        "industria",
        "moneda",
        "monedaMonto",
        "montoaLicitar",
        "monto_Adjudicado",
        "ampliableHasta",
        "variableLicitar",
        "valor_Corte",
        "duration",
        "rueda",
        "modalidad",
        "liquidador",
        "estado",
        "tipo",
        "colocador",
        "sistema_Adjudicacion",
        "observaciones",
        "comentario",
        "fechaModificacion",
        "plazoEspecie",
        "existeArchivo",
    ]

    for col in columnas_base:
        if col not in df.columns:
            df[col] = pd.NA

    # Fechas.
    for col in [
        "fechaInicio",
        "fechaFin",
        "fechaLiquidacion",
        "fechaVencimiento",
        "fechaVencimientoEspecie",
        "fechaModificacion",
    ]:
        df[col] = df[col].apply(parse_fecha_utc_a_fecha)

    # Texto básico.
    for col in [
        "titulo",
        "emisor",
        "industria",
        "moneda",
        "monedaMonto",
        "ampliableHasta",
        "variableLicitar",
        "valor_Corte",
        "duration",
        "rueda",
        "modalidad",
        "liquidador",
        "estado",
        "tipo",
        "colocador",
        "sistema_Adjudicacion",
        "observaciones",
        "comentario",
        "plazoEspecie",
    ]:
        df[col] = df[col].apply(normalizar_texto)

    # Números.
    for col in ["montoaLicitar", "monto_Adjudicado"]:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    # Campos útiles derivados.
    df["valor_corte_tipo"] = df.apply(
        lambda r: clasificar_variable_corte(r.get("valor_Corte"), r.get("variableLicitar")),
        axis=1,
    )

    df["valor_corte_num"] = df["valor_Corte"].apply(extraer_valor_corte_numero)

    df["es_on_corporativa"] = True
    df["fecha_descarga"] = datetime.now()

    # Orden final más cómodo para análisis.
    columnas_finales = [
        "fechaInicio",
        "fechaLiquidacion",
        "fechaVencimiento",
        "titulo",
        "emisor",
        "tipo",
        "moneda",
        "montoaLicitar",
        "monto_Adjudicado",
        "variableLicitar",
        "valor_Corte",
        "valor_corte_tipo",
        "valor_corte_num",
        "duration",
        "ampliableHasta",
        "observaciones",
        "fechaModificacion",
    ]

    columnas_finales = [c for c in columnas_finales if c in df.columns]

    df = df[columnas_finales].copy()

    # Orden cronológico.
    df = df.sort_values(
        by=["fechaInicio", "emisor", "titulo"],
        ascending=[True, True, True],
        na_position="last",
    ).reset_index(drop=True)

    return df


def exportar_excel(df: pd.DataFrame, output_file: Path) -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="ON_Corporativas", index=False)

        ws = writer.book["ON_Corporativas"]

        # Freeze panes y autofilter.
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

        # Anchos básicos.
        widths = {
            "A": 20,   # fechaInicio
            "B": 20,   # fechaLiquidacion
            "C": 20,   # fechaVencimiento
            "D": 55,   # titulo
            "E": 38,   # emisor
            "F": 22,   # tipo
            "G": 14,   # moneda
            "H": 18,   # montoaLicitar
            "I": 20,   # monto_Adjudicado
            "J": 18,   # variableLicitar
            "K": 35,   # valor_Corte
            "L": 16,   # valor_corte_tipo
            "M": 16,   # valor_corte_num
            "N": 16,   # duration
            "O": 28,   # ampliableHasta
            "P": 55,   # observaciones
            "Q": 20,   # fechaModificacion
        }

        for col, width in widths.items():
            ws.column_dimensions[col].width = width

        # Formatos numéricos/fecha simples.
        fecha_cols = ["A", "B", "C", "Q"]
        for col in fecha_cols:
            for cell in ws[col][1:]:
                cell.number_format = "dd/mm/yyyy hh:mm"

        for col in ["H", "I"]:
            for cell in ws[col][1:]:
                cell.number_format = '#,##0'

        for col in ["M"]:
            for cell in ws[col][1:]:
                cell.number_format = '0.00'

    log(f"Excel exportado: {output_file}")


def main() -> int:
    try:
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

        endpoint_url = construir_endpoint()
        log(f"Fecha desde: {FECHA_DESDE}")
        log(f"Fecha hasta: {date.today().strftime('%Y-%m-%d')}")

        data = descargar_json(endpoint_url)

        log(f"Registros totales descargados: {len(data):,}")

        data_filtrada = [item for item in data if es_on_corporativa(item)]

        log(f"Registros filtrados como ON corporativas: {len(data_filtrada):,}")

        df = limpiar_dataframe(data_filtrada)

        if df.empty:
            log("ALERTA: el filtro no dejó registros. Revisión necesaria.")
        else:
            resumen_tipo = df["tipo"].value_counts(dropna=False)
            log("Resumen por tipo filtrado:")
            for tipo, cantidad in resumen_tipo.items():
                log(f"  - {tipo}: {cantidad}")

        exportar_excel(df, OUTPUT_FILE)

        log("Proceso finalizado OK.")
        return 0

    except Exception as e:
        log(f"ERROR: {e}")
        return 1


if __name__ == "__main__":
    sys.exit(main())
