import argparse
import os
import re
import shutil
import subprocess
import sys
from pathlib import Path


PROJECT_DIR = Path(__file__).resolve().parent
VENV_PYINSTALLER = PROJECT_DIR / ".venv" / "Scripts" / "pyinstaller.exe"
MAIN_PY = PROJECT_DIR / "main.py"
DIST_DIR = PROJECT_DIR / "dist"
BUILD_DIR = PROJECT_DIR / "build"


def next_version():
    pattern = re.compile(r"notificaciones_sitfa_v(\d+)\.exe$", re.IGNORECASE)
    max_version = 0
    if DIST_DIR.exists():
        for item in DIST_DIR.iterdir():
            match = pattern.match(item.name)
            if match:
                max_version = max(max_version, int(match.group(1)))
    return max_version + 1


def build(version):
    if not VENV_PYINSTALLER.exists():
        raise FileNotFoundError(f"No se encontró PyInstaller en: {VENV_PYINSTALLER}")
    if not MAIN_PY.exists():
        raise FileNotFoundError(f"No se encontró main.py en: {MAIN_PY}")

    DIST_DIR.mkdir(exist_ok=True)

    cmd = [
        str(VENV_PYINSTALLER),
        "--noconfirm",
        "--onefile",
        "--name",
        "notificaciones_sitfa",
        "--add-data",
        "Notifica liq CON RD.doc;.",
        "--add-data",
        "Notifica liq sin rd.doc;.",
        str(MAIN_PY),
    ]

    subprocess.run(cmd, cwd=str(PROJECT_DIR), check=True)

    base_exe = DIST_DIR / "notificaciones_sitfa.exe"
    versioned_exe = DIST_DIR / f"notificaciones_sitfa_v{version}.exe"
    if not base_exe.exists():
        raise FileNotFoundError(f"No se encontró el ejecutable generado: {base_exe}")

    if versioned_exe.exists():
        versioned_exe.unlink()

    shutil.copy2(base_exe, versioned_exe)
    print(f"Ejecutable generado: {versioned_exe}")
    return versioned_exe


def main():
    parser = argparse.ArgumentParser(description="Empaqueta el proyecto en un EXE versionado.")
    parser.add_argument("--version", type=int, help="Versión a usar. Si no se indica, se incrementa automáticamente.")
    args = parser.parse_args()

    version = args.version if args.version else next_version()
    print(f"Usando versión: v{version}")
    build(version)


if __name__ == "__main__":
    try:
        main()
    except subprocess.CalledProcessError as e:
        print(f"Error ejecutando PyInstaller: {e}")
        sys.exit(e.returncode)
    except Exception as e:
        print(f"Error: {e}")
        sys.exit(1)
