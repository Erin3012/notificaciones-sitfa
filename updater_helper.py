import argparse
import ctypes
import os
import shutil
import subprocess
import sys
import time


def _proceso_termino(pid):
    if os.name != "nt":
        return True
    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    proceso = kernel32.OpenProcess(0x00100000, False, pid)
    if not proceso:
        return True
    try:
        return kernel32.WaitForSingleObject(proceso, 0) != 0x00000102
    finally:
        kernel32.CloseHandle(proceso)


def reemplazar_y_reiniciar(pid, source, target):
    for _ in range(120):
        if _proceso_termino(pid):
            break
        time.sleep(0.5)
    else:
        raise RuntimeError("El programa anterior no termino a tiempo.")

    ultimo_error = None
    for _ in range(20):
        try:
            staged = f"{target}.update"
            shutil.copyfile(source, staged)
            os.replace(staged, target)
            try:
                os.remove(source)
                os.rmdir(os.path.dirname(source))
            except OSError:
                pass
            ultimo_error = None
            break
        except OSError as error:
            ultimo_error = error
            time.sleep(0.5)
    if ultimo_error:
        raise ultimo_error
    subprocess.Popen([target], cwd=os.path.dirname(target), close_fds=True)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--pid", type=int, required=True)
    parser.add_argument("--source", required=True)
    parser.add_argument("--target", required=True)
    args = parser.parse_args()
    reemplazar_y_reiniciar(args.pid, args.source, args.target)


if __name__ == "__main__":
    try:
        main()
    except Exception as error:
        with open(os.path.join(os.path.dirname(sys.executable), "actualizacion_error.log"), "w", encoding="utf-8") as archivo:
            archivo.write(str(error))
        raise

