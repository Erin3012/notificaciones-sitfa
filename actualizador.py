import json
import os
import re
import subprocess
import sys
import tempfile
import time
import urllib.error
import urllib.request
from tkinter import Canvas, Label, Tk, messagebox

from version import APP_VERSION


REPOSITORIO = "Erin3012/notificaciones-sitfa"
API_RELEASE_LATEST = f"https://api.github.com/repos/{REPOSITORIO}/releases/latest"
MINIMO_EXE_BYTES = 100 * 1024
TIEMPO_MINIMO_ACTUALIZADOR = 10.0
INTERVALO_MENSAJE = 1.5
MENSAJES_LANZAMIENTO = (
    "Inicializando nucleo de actualizacion",
    "Contactando con el servidor de misiones",
    "Verificando version del sistema",
    "Cargando complementos",
    "Sincronizando modulos del robot",
    "Renderizando motor de ejecucion",
    "Optimizando recursos de combate",
    "Preparando salto de version",
    "Calibrando protocolos SITFA",
)


class ErrorActualizacion(Exception):
    pass


class VentanaActualizacion:
    def __init__(self):
        self.root = Tk()
        self.inicio = time.monotonic()
        self.mensaje_index = 0
        self.proximo_mensaje = self.inicio
        self.progreso_actual = 0
        self.root.title("Robot SITFA")
        self.root.geometry("640x310")
        self.root.resizable(False, False)
        self.root.configure(bg="#f3f6fb")
        self.root.protocol("WM_DELETE_WINDOW", self._bloquear_cierre)

        self.root.update_idletasks()
        ancho = self.root.winfo_screenwidth()
        alto = self.root.winfo_screenheight()

        x = (ancho - 640) // 2
        y = (alto - 310) // 2
        self.root.geometry(f"640x310+{x}+{y}")

        encabezado = Canvas(self.root, width=640, height=106, bg="#ffffff", highlightthickness=0)
        encabezado.pack(fill="x")
        encabezado.create_polygon(32, 30, 50, 20, 68, 30, 68, 51, 50, 61, 32, 51, fill="#0b74c9", outline="")
        encabezado.create_text(50, 41, text="S", fill="#ffffff", font=("Segoe UI", 19, "bold"))
        encabezado.create_text(86, 35, text="SITFA", anchor="w", fill="#17365d", font=("Segoe UI", 24, "bold"))
        encabezado.create_text(595, 35, text=f"v{APP_VERSION}", anchor="e", fill="#17365d", font=("Segoe UI", 10, "bold"))
        encabezado.create_text(595, 64, text="ACTUALIZACION", anchor="e", fill="#0b74c9", font=("Segoe UI", 8, "bold"))

        tarjeta = Canvas(self.root, width=560, height=143, bg="#ffffff", highlightthickness=1, highlightbackground="#dce6f2")
        tarjeta.pack(pady=(18, 0))
        tarjeta.create_oval(25, 26, 73, 74, fill="#dceefe", outline="")
        tarjeta.create_text(49, 50, text="↻", fill="#0b74c9", font=("Segoe UI", 25, "bold"))
        tarjeta.create_text(92, 35, text="Actualizacion del sistema", anchor="w", fill="#17365d", font=("Segoe UI", 14, "bold"))
        self.estado = Label(tarjeta, text="Inicializando nucleo de actualizacion", bg="#ffffff", fg="#60738e", font=("Segoe UI", 10))
        self.estado.place(x=92, y=59, anchor="w")

        self.lienzo = Canvas(tarjeta, width=430, height=20, bg="#ffffff", highlightthickness=0)
        self.lienzo.place(x=92, y=92)
        self.lienzo.create_rectangle(0, 2, 430, 18, fill="#e4ebf4", outline="")
        self.barra = self.lienzo.create_rectangle(0, 2, 0, 18, fill="#1677d2", outline="")
        self.porcentaje = Label(tarjeta, text="0%", bg="#ffffff", fg="#1677d2", font=("Segoe UI", 11, "bold"))
        self.porcentaje.place(x=535, y=101, anchor="e")

        self.detalle = Label(self.root, text=f"Version instalada: v{APP_VERSION}   |   No cierres la aplicacion durante el proceso", bg="#f3f6fb", fg="#60738e", font=("Segoe UI", 9))
        self.detalle.pack(pady=(10, 0))

        pie = Canvas(self.root, width=560, height=58, bg="#f3f6fb", highlightthickness=0)
        pie.pack_forget()
        for x, titulo, texto in ((0, "Seguro", "Datos protegidos"), (190, "Confiable", "Actualizaciones oficiales"), (380, "Siempre contigo", "Servicio actualizado")):
            pie.create_rectangle(x, 0, x + 175, 54, fill="#ffffff", outline="#dce6f2")
            pie.create_text(x + 16, 18, text="●", fill="#1677d2", font=("Segoe UI", 12, "bold"))
            pie.create_text(x + 37, 17, text=titulo, anchor="w", fill="#17365d", font=("Segoe UI", 9, "bold"))
            pie.create_text(x + 37, 36, text=texto, anchor="w", fill="#71819a", font=("Segoe UI", 8))
        self.root.update_idletasks()

    def _bloquear_cierre(self):
        pass

    def mensaje(self, texto):
        self.estado.configure(text=texto)
        self.root.update_idletasks()

    def mensaje_rotativo(self):
        ahora = time.monotonic()
        if ahora < self.proximo_mensaje:
            return
        self.mensaje(MENSAJES_LANZAMIENTO[self.mensaje_index % len(MENSAJES_LANZAMIENTO)])
        self.mensaje_index += 1
        self.proximo_mensaje = ahora + INTERVALO_MENSAJE

    def avance(self, actual, total):
        if total > 0:
            porcentaje_descarga = min(100, int(actual * 100 / total))
            self.set_progress(20 + int(porcentaje_descarga * 0.70))
            self.mensaje_rotativo()
        else:
            self.set_progress(min(20, self.progreso_actual + 1))
            self.mensaje_rotativo()
        self.root.update_idletasks()

    def set_progress(self, porcentaje):
        self.progreso_actual = max(0, min(100, int(porcentaje)))
        ancho = int(430 * self.progreso_actual / 100)
        self.lienzo.coords(self.barra, 0, 2, ancho, 18)
        self.porcentaje.configure(text=f"{self.progreso_actual}%")
        self.root.update_idletasks()

    def completar_tiempo_minimo(self):
        while time.monotonic() - self.inicio < TIEMPO_MINIMO_ACTUALIZADOR:
            transcurrido = time.monotonic() - self.inicio
            progreso_minimo = int(min(100, (transcurrido / TIEMPO_MINIMO_ACTUALIZADOR) * 100))
            self.set_progress(max(self.progreso_actual, progreso_minimo))
            self.mensaje_rotativo()
            time.sleep(0.25)
        self.set_progress(100)

    def cerrar(self):
        try:
            self.root.destroy()
        except Exception:
            pass


def _solicitar_json(url):
    solicitud = urllib.request.Request(url, headers={"Accept": "application/vnd.github+json", "User-Agent": "Robot-SITFA-Updater"})
    try:
        with urllib.request.urlopen(solicitud, timeout=20) as respuesta:
            return json.loads(respuesta.read().decode("utf-8"))
    except (urllib.error.URLError, urllib.error.HTTPError, TimeoutError, json.JSONDecodeError) as error:
        raise ErrorActualizacion(f"No se pudo consultar GitHub: {error}") from error


def _version_desde_tag(tag):
    coincidencia = re.fullmatch(r"v(\d+)", str(tag or "").strip())
    if not coincidencia:
        raise ErrorActualizacion(f"El release de GitHub tiene un tag invalido: {tag!r}")
    return int(coincidencia.group(1))


def _descargar_exe(url, destino, ventana):
    solicitud = urllib.request.Request(url, headers={"Accept": "application/octet-stream", "User-Agent": "Robot-SITFA-Updater"})
    try:
        with urllib.request.urlopen(solicitud, timeout=60) as respuesta:
            total = int(respuesta.headers.get("Content-Length") or 0)
            descargado = 0
            with open(destino, "wb") as archivo:
                while True:
                    bloque = respuesta.read(1024 * 1024)
                    if not bloque:
                        break
                    archivo.write(bloque)
                    descargado += len(bloque)
                    ventana.avance(descargado, total)
    except (urllib.error.URLError, urllib.error.HTTPError, OSError, TimeoutError) as error:
        raise ErrorActualizacion(f"No se pudo descargar la actualizacion: {error}") from error
    if not os.path.isfile(destino) or os.path.getsize(destino) < MINIMO_EXE_BYTES:
        raise ErrorActualizacion("La descarga de la actualizacion esta incompleta o es invalida.")


def _asset_exe(release):
    candidatos = [asset for asset in release.get("assets") or [] if str(asset.get("name", "")).lower().endswith(".exe") and "updater_helper" not in str(asset.get("name", "")).lower()]
    if not candidatos:
        raise ErrorActualizacion("El release no contiene el EXE principal.")
    return candidatos[0]


def _iniciar_helper(helper, origen, destino):
    if not os.path.isfile(helper):
        raise ErrorActualizacion("Falta updater_helper.exe junto al programa. Instala la distribucion completa.")
    argumentos = [helper, "--pid", str(os.getpid()), "--source", origen, "--target", destino]
    flags = getattr(subprocess, "DETACHED_PROCESS", 0) | getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0)
    subprocess.Popen(argumentos, creationflags=flags, close_fds=True)


def verificar_actualizacion_obligatoria():
    """Devuelve True si la aplicacion puede continuar; False si debe terminar."""
    if not getattr(sys, "frozen", False):
        return True

    ventana = VentanaActualizacion()
    try:
        release = _solicitar_json(API_RELEASE_LATEST)
        version_remota = _version_desde_tag(release.get("tag_name"))
        if version_remota <= APP_VERSION:
            ventana.mensaje("Sistema actualizado. Verificando modulos instalados")
            ventana.completar_tiempo_minimo()
            ventana.cerrar()
            return True

        asset = _asset_exe(release)
        ventana.mensaje(f"Descargando la version v{version_remota}...")
        carpeta_temp = tempfile.mkdtemp(prefix="robot_sitfa_update_")
        descarga = os.path.join(carpeta_temp, asset["name"])
        _descargar_exe(asset["browser_download_url"], descarga, ventana)
        destino = os.path.abspath(sys.executable)
        helper = os.path.join(os.path.dirname(destino), "updater_helper.exe")
        ventana.set_progress(94)
        ventana.mensaje("Aplicando parche de estabilidad del nucleo")
        ventana.mensaje("Preparando la instalacion...")
        _iniciar_helper(helper, descarga, destino)
        ventana.set_progress(100)
        ventana.completar_tiempo_minimo()
        ventana.mensaje("Reiniciando con la nueva version...")
        ventana.cerrar()
        return False
    except (ErrorActualizacion, OSError) as error:
        ventana.completar_tiempo_minimo()
        ventana.cerrar()
        messagebox.showerror("Actualizacion obligatoria", f"No se puede iniciar Robot SITFA hasta actualizarlo.\n\n{error}")
        return False
    except Exception as error:
        ventana.completar_tiempo_minimo()
        ventana.cerrar()
        messagebox.showerror("Actualizacion obligatoria", f"No se pudo verificar la version de Robot SITFA.\n\n{error}")
        return False

