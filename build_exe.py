import argparse
import re
import shutil
import subprocess
import sys
from pathlib import Path


PROJECT_DIR = Path(__file__).resolve().parent
VENV_PYINSTALLER = PROJECT_DIR / ".venv" / "Scripts" / "pyinstaller.exe"
ENTRY_PY = PROJECT_DIR / "web_ui.py"
DIST_DIR = PROJECT_DIR / "dist"
VERSION_PY = PROJECT_DIR / "version.py"


def pyinstaller_command():
    """Usa el PyInstaller del venv, del PATH o del Python actual."""
    if VENV_PYINSTALLER.exists():
        return [str(VENV_PYINSTALLER)]
    executable = shutil.which("pyinstaller")
    if executable:
        return [executable]
    return [sys.executable, "-m", "PyInstaller"]


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
    if not ENTRY_PY.exists():
        raise FileNotFoundError(f"No se encontró web_ui.py en: {ENTRY_PY}")
    for required in ("Notifica liq CON RD.doc", "Notifica liq sin rd.doc", "IEDriverServer.exe"):
        if not (PROJECT_DIR / required).exists():
            raise FileNotFoundError(f"No se encontró el archivo requerido: {PROJECT_DIR / required}")

    DIST_DIR.mkdir(exist_ok=True)
    VERSION_PY.write_text(f'"""Version generada durante el build."""\nAPP_VERSION = {version}\n', encoding="utf-8")

    helper_cmd = pyinstaller_command() + [
        "--noconfirm",
        "--onefile",
        "--windowed",
        "--name",
        "updater_helper",
        "--specpath",
        str(PROJECT_DIR / "build"),
        str(PROJECT_DIR / "updater_helper.py"),
    ]
    subprocess.run(helper_cmd, cwd=str(PROJECT_DIR), check=True)

    cmd = pyinstaller_command() + [
        "--noconfirm",
        "--onefile",
        "--name",
        "notificaciones_sitfa",
        "--specpath",
        str(PROJECT_DIR / "build"),
        "--add-data",
        f"{PROJECT_DIR / 'Notifica liq CON RD.doc'};.",
        "--add-data",
        f"{PROJECT_DIR / 'Notifica liq sin rd.doc'};.",
        "--add-data",
        f"{PROJECT_DIR / 'IEDriverServer.exe'};.",
        "--hidden-import",
        "main",
        "--hidden-import",
        "pythoncom",
        "--hidden-import",
        "win32com.client",
        "--hidden-import",
        "win32gui",
        "--hidden-import",
        "win32con",
        str(ENTRY_PY),
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
    print(f"Ejecutable del actualizador: {DIST_DIR / 'updater_helper.exe'}")
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
    except subprocess.CalledProcessError as exc:
        print(f"Error ejecutando PyInstaller: {exc}")
        sys.exit(exc.returncode)
    except Exception as exc:
        print(f"Error: {exc}")
        sys.exit(1)
