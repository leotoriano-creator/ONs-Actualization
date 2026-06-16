# -*- coding: utf-8 -*-
"""
RunAll - Automatización A3 / Base ONs

Corre en orden:
    1) GetData/GetDataFromA3.py
    2) GetData/GetA3500&TAMAR.py
    3) ExcelModifier/ModifyExcelUSD&TAMAR.py
    4) ExcelModifier/ModifyExcel.py

Railway:
    Start Command:
        python RunAll/RunAll.py

    Variable obligatoria:
        GOOGLE_SERVICE_ACCOUNT_JSON
"""

from __future__ import annotations

import os
import sys
import subprocess
from pathlib import Path
from datetime import datetime


SCRIPT_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = SCRIPT_DIR.parent

GETDATA_DIR = PROJECT_ROOT / "GetData"
EXCEL_MODIFIER_DIR = PROJECT_ROOT / "ExcelModifier"
DATA_STORAGE_DIR = PROJECT_ROOT / "DataStorage"
LOGS_DIR = PROJECT_ROOT / "Logs"

SCRIPTS = [
    {
        "name": "Bajar colocaciones A3",
        "path": GETDATA_DIR / "GetDataFromA3.py",
        "required_outputs": [DATA_STORAGE_DIR / "a3_on_corporativas.xlsx"],
    },
    {
        "name": "Bajar BADLAR / TAMAR / COM3500",
        "path": GETDATA_DIR / "GetA3500&TAMAR.py",
        "required_outputs": [DATA_STORAGE_DIR / "variables_mercado.xlsx"],
    },
    {
        "name": "Actualizar solapa VARIABLES",
        "path": EXCEL_MODIFIER_DIR / "ModifyExcelUSD&TAMAR.py",
        "required_outputs": [],
    },
    {
        "name": "Actualizar solapa BASE",
        "path": EXCEL_MODIFIER_DIR / "ModifyExcel.py",
        "required_outputs": [],
    },
]

LOG_LINES: list[str] = []


def now_str() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def log(msg: str) -> None:
    line = f"{now_str()} | {msg}"
    LOG_LINES.append(line)
    print(line, flush=True)


def guardar_log() -> Path:
    LOGS_DIR.mkdir(parents=True, exist_ok=True)
    log_file = LOGS_DIR / f"runall_a3_ons_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
    log_file.write_text("\n".join(LOG_LINES), encoding="utf-8")
    return log_file


def validar_estructura() -> None:
    errores: list[str] = []

    if not GETDATA_DIR.exists():
        errores.append(f"No existe carpeta GetData: {GETDATA_DIR}")

    if not EXCEL_MODIFIER_DIR.exists():
        errores.append(f"No existe carpeta ExcelModifier: {EXCEL_MODIFIER_DIR}")

    DATA_STORAGE_DIR.mkdir(parents=True, exist_ok=True)
    LOGS_DIR.mkdir(parents=True, exist_ok=True)

    for item in SCRIPTS:
        if not item["path"].exists():
            errores.append(f"No existe script: {item['path']}")

    google_credentials = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON", "").strip()
    if not google_credentials:
        errores.append("No existe la variable de entorno GOOGLE_SERVICE_ACCOUNT_JSON.")

    if errores:
        raise FileNotFoundError("\n".join(errores))

    log("Variable GOOGLE_SERVICE_ACCOUNT_JSON detectada correctamente.")


def validar_outputs(outputs: list[Path], etapa: str) -> None:
    for output in outputs:
        if not output.exists():
            raise FileNotFoundError(
                f"La etapa '{etapa}' terminó con código 0, "
                f"pero no encontré el archivo esperado: {output}"
            )


def correr_script(script_path: Path, name: str) -> int:
    log("=" * 80)
    log(f"INICIO: {name}")
    log(f"Script: {script_path}")
    log("=" * 80)

    result = subprocess.run(
        [sys.executable, str(script_path)],
        cwd=str(PROJECT_ROOT),
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        env=os.environ.copy(),
    )

    if result.stdout:
        print(result.stdout, end="", flush=True)
        LOG_LINES.append("")
        LOG_LINES.append(f"STDOUT - {name}")
        LOG_LINES.append("-" * 80)
        LOG_LINES.append(result.stdout)

    if result.stderr:
        print(result.stderr, end="", flush=True)
        LOG_LINES.append("")
        LOG_LINES.append(f"STDERR - {name}")
        LOG_LINES.append("-" * 80)
        LOG_LINES.append(result.stderr)

    if result.returncode == 0:
        log(f"OK: {name}")
    else:
        log(f"ERROR: {name} terminó con código {result.returncode}")

    return result.returncode


def main() -> int:
    start = datetime.now()

    try:
        log("=" * 80)
        log("INICIO RUNALL A3 / BASE ONS")
        log(f"Proyecto: {PROJECT_ROOT}")
        log(f"Python: {sys.executable}")
        log("=" * 80)

        validar_estructura()

        for item in SCRIPTS:
            returncode = correr_script(item["path"], item["name"])

            if returncode != 0:
                raise RuntimeError(f"Falló la etapa: {item['name']}")

            validar_outputs(item.get("required_outputs", []), item["name"])

        elapsed = datetime.now() - start
        log("=" * 80)
        log("RUNALL FINALIZADO OK")
        log(f"Duración: {elapsed}")
        log("=" * 80)

        log_file = guardar_log()
        print(f"{now_str()} | Log guardado en: {log_file}", flush=True)
        return 0

    except Exception as e:
        elapsed = datetime.now() - start
        log("=" * 80)
        log("RUNALL FINALIZADO CON ERROR")
        log(f"Error: {repr(e)}")
        log(f"Duración: {elapsed}")
        log("=" * 80)

        log_file = guardar_log()
        print(f"{now_str()} | Log guardado en: {log_file}", flush=True)
        return 1


if __name__ == "__main__":
    sys.exit(main())
