# -*- coding: utf-8 -*-
"""
RunAll - Automatización A3 / Base ONs

Corre en orden:
    1) GetData/GetDataFromA3.py
       Baja colocaciones A3 y genera DataStorage/a3_on_corporativas.xlsx

    2) GetData/GetA3500&TAMAR.py
       Baja BADLAR, TAMAR y COM3500 y genera DataStorage/variables_mercado.xlsx

    3) ExcelModifier/ModifyExcelUSD&TAMAR.py
       Actualiza la solapa VARIABLES del Excel Base ONs - Nueva.xlsx

    4) ExcelModifier/ModifyExcel.py
       Actualiza la solapa BASE del Excel Base ONs - Nueva.xlsx

Estructura esperada:
A3 Scrapper/
│
├── DataStorage/
├── ExcelModifier/
│   ├── ModifyExcel.py
│   └── ModifyExcelUSD&TAMAR.py
├── GetData/
│   ├── GetDataFromA3.py
│   └── GetA3500&TAMAR.py
├── Logs/
├── RunAll/
│   └── RunAll.py
└── ServiceAccount/
    └── service_account.json
"""

from __future__ import annotations

import sys
import subprocess
from pathlib import Path
from datetime import datetime


# ============================================================
# PATHS
# ============================================================

SCRIPT_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = SCRIPT_DIR.parent

GETDATA_DIR = PROJECT_ROOT / "GetData"
EXCEL_MODIFIER_DIR = PROJECT_ROOT / "ExcelModifier"
DATA_STORAGE_DIR = PROJECT_ROOT / "DataStorage"
LOGS_DIR = PROJECT_ROOT / "Logs"
SERVICE_ACCOUNT_FILE = PROJECT_ROOT / "ServiceAccount" / "service_account.json"


SCRIPTS = [
    {
        "name": "Bajar colocaciones A3",
        "path": GETDATA_DIR / "GetDataFromA3.py",
        "required_outputs": [
            DATA_STORAGE_DIR / "a3_on_corporativas.xlsx",
        ],
    },
    {
        "name": "Bajar BADLAR / TAMAR / COM3500",
        "path": GETDATA_DIR / "GetA3500&TAMAR.py",
        "required_outputs": [
            DATA_STORAGE_DIR / "variables_mercado.xlsx",
        ],
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


# ============================================================
# LOG
# ============================================================

LOG_LINES: list[str] = []


def now_str() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def log(msg: str) -> None:
    line = f"{now_str()} | {msg}"
    LOG_LINES.append(line)
    print(line)


def guardar_log() -> Path:
    LOGS_DIR.mkdir(parents=True, exist_ok=True)

    log_file = LOGS_DIR / f"runall_a3_ons_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
    log_file.write_text("\n".join(LOG_LINES), encoding="utf-8")

    return log_file


# ============================================================
# VALIDACIONES
# ============================================================

def validar_estructura() -> None:
    errores = []

    if not GETDATA_DIR.exists():
        errores.append(f"No existe carpeta GetData: {GETDATA_DIR}")

    if not EXCEL_MODIFIER_DIR.exists():
        errores.append(f"No existe carpeta ExcelModifier: {EXCEL_MODIFIER_DIR}")

    if not DATA_STORAGE_DIR.exists():
        DATA_STORAGE_DIR.mkdir(parents=True, exist_ok=True)

    if not LOGS_DIR.exists():
        LOGS_DIR.mkdir(parents=True, exist_ok=True)

    if not SERVICE_ACCOUNT_FILE.exists():
        errores.append(f"No existe service account: {SERVICE_ACCOUNT_FILE}")

    for item in SCRIPTS:
        script_path = item["path"]
        if not script_path.exists():
            errores.append(f"No existe script: {script_path}")

    if errores:
        raise FileNotFoundError("\n".join(errores))


def validar_outputs(outputs: list[Path], etapa: str) -> None:
    for output in outputs:
        if not output.exists():
            raise FileNotFoundError(
                f"La etapa '{etapa}' terminó OK, pero no encontré el archivo esperado: {output}"
            )


# ============================================================
# EJECUCIÓN
# ============================================================

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
    )

    if result.stdout:
        print(result.stdout)
        LOG_LINES.append("")
        LOG_LINES.append(f"STDOUT - {name}")
        LOG_LINES.append("-" * 80)
        LOG_LINES.append(result.stdout)

    if result.stderr:
        print(result.stderr)
        LOG_LINES.append("")
        LOG_LINES.append(f"STDERR - {name}")
        LOG_LINES.append("-" * 80)
        LOG_LINES.append(result.stderr)

    if result.returncode == 0:
        log(f"OK: {name}")
    else:
        log(f"ERROR: {name} terminó con código {result.returncode}")

    return result.returncode


# ============================================================
# MAIN
# ============================================================

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
            name = item["name"]
            path = item["path"]

            returncode = correr_script(path, name)

            if returncode != 0:
                raise RuntimeError(f"Falló la etapa: {name}")

            validar_outputs(item.get("required_outputs", []), name)

        elapsed = datetime.now() - start

        log("=" * 80)
        log("RUNALL FINALIZADO OK")
        log(f"Duración: {elapsed}")
        log("=" * 80)

        log_file = guardar_log()
        print(f"{now_str()} | Log guardado en: {log_file}")

        return 0

    except Exception as e:
        elapsed = datetime.now() - start

        log("=" * 80)
        log("RUNALL FINALIZADO CON ERROR")
        log(f"Error: {repr(e)}")
        log(f"Duración: {elapsed}")
        log("=" * 80)

        log_file = guardar_log()
        print(f"{now_str()} | Log guardado en: {log_file}")

        return 1


if __name__ == "__main__":
    sys.exit(main())