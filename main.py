import tkinter as tk
from tkinter import simpledialog, messagebox, filedialog
import win32com.client
import getpass
import time
import threading
import math
import os
import sys
import unicodedata
import tempfile
import re
import socket
import uuid
import traceback

try:
    from selenium import webdriver
    from selenium.common.exceptions import (
        NoAlertPresentException,
        NoSuchWindowException,
        TimeoutException,
        WebDriverException,
    )
    from selenium.webdriver.common.by import By
    from selenium.webdriver.ie.options import Options as IeOptions
    from selenium.webdriver.ie.service import Service as IeService
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.webdriver.support.select import Select
    from selenium.webdriver.support.ui import WebDriverWait
    SELENIUM_IMPORT_ERROR = None
except Exception as exc:
    webdriver = None
    NoAlertPresentException = NoSuchWindowException = TimeoutException = WebDriverException = Exception
    By = IeOptions = IeService = Select = WebDriverWait = EC = None
    SELENIUM_IMPORT_ERROR = exc

# Variables globales para las rutas de los documentos de plantilla seleccionados
ruta_plantilla_con_rd = None
ruta_plantilla_sin_rd = None

# Puente opcional para interfaces externas (por ejemplo, la interfaz web).
# Cuando no se configura, la aplicación conserva sus diálogos de Tkinter.
UI_CONFIRM_CALLBACK = None
UI_LIQUIDACION_CALLBACK = None
UI_STOP_EVENT = None


def automation_should_stop():
    return bool(UI_STOP_EVENT is not None and UI_STOP_EVENT.is_set())


def solicitar_datos_liquidacion(parent, **kwargs):
    if UI_LIQUIDACION_CALLBACK:
        return UI_LIQUIDACION_CALLBACK(**kwargs)
    return LiquidacionInfoDialog(parent, **kwargs).result

LOGIN_URL = "http://www.familia.pjud/SITFAWEB/jsp/Login/Login.jsp"
POST_LOGIN_URL = "http://www.familia.pjud/SITFAWEB/MenuLinkAction.do?opMenu=ConsultaRolTramitar&opRol=59"


def cargar_dotenv_simple(ruta_archivo=".env"):
    if not os.path.exists(ruta_archivo):
        return

    try:
        with open(ruta_archivo, "r", encoding="utf-8") as fh:
            for raw_line in fh:
                line = raw_line.strip()
                if not line or line.startswith("#") or "=" not in line:
                    continue
                clave, valor = line.split("=", 1)
                clave = clave.strip()
                valor = valor.strip().strip('"').strip("'")
                if clave and clave not in os.environ:
                    os.environ[clave] = valor
    except Exception as e:
        print(f"[WARN] No se pudo cargar .env: {e}")


cargar_dotenv_simple()


FILL_LOGIN_SCRIPT = """
var form = document.forms["InicioAplicacionForm"];
if (!form) {
    throw new Error("No se encontro el formulario InicioAplicacionForm");
}
var data = arguments[0] || {};
function setValue(name, value) {
    var el = form.elements[name];
    if (el) {
        el.value = value || "";
    }
}
setValue("username", data.username);
setValue("password", data.password);
setValue("nomequipo", data.machineName);
setValue("usuequipo", data.windowsUser);
setValue("ipequipo", data.localIp);
setValue("MAC_equipo", data.macAddress);
return true;
"""

SUBMIT_LOGIN_SCRIPT = """
var form = document.forms["InicioAplicacionForm"];
if (!form) {
    throw new Error("No se encontro el formulario InicioAplicacionForm");
}
form.submit();
"""

def resolver_rutas_plantillas():
    """
    Prioridad:
    1) Carpeta del .exe / script externo.
    2) Carpeta embebida por PyInstaller.
    3) None para pedirla manualmente.
    """
    candidatos_base = []

    exe_dir = os.path.dirname(os.path.abspath(sys.argv[0])) if sys.argv and sys.argv[0] else None
    if exe_dir:
        candidatos_base.append(exe_dir)

    # Cuando corre empaquetado, PyInstaller expone la ruta temporal en _MEIPASS.
    meipass_dir = getattr(sys, "_MEIPASS", None)
    if meipass_dir:
        candidatos_base.append(meipass_dir)

    # Ãšltimo fallback: carpeta del script.
    script_dir = os.path.dirname(os.path.abspath(__file__))
    candidatos_base.append(script_dir)

    ruta_con_rd = None
    ruta_sin_rd = None

    for base_dir in candidatos_base:
        if not ruta_con_rd:
            ruta = os.path.join(base_dir, "Notifica liq CON RD.doc")
            if os.path.exists(ruta):
                ruta_con_rd = ruta
        if not ruta_sin_rd:
            ruta = os.path.join(base_dir, "Notifica liq sin rd.doc")
            if os.path.exists(ruta):
                ruta_sin_rd = ruta
        if ruta_con_rd and ruta_sin_rd:
            break

    return ruta_con_rd, ruta_sin_rd


def get_local_ip():
    try:
        with socket.socket(socket.AF_INET, socket.SOCK_DGRAM) as sock:
            sock.connect(("8.8.8.8", 80))
            return sock.getsockname()[0]
    except OSError:
        return "0"


def get_mac_address():
    mac = uuid.getnode()
    return "-".join(f"{(mac >> shift) & 0xFF:02X}" for shift in range(40, -1, -8))


def safe_get(driver, url, timeout=20):
    debug_login(f"safe_get inicio url={url!r} timeout={timeout}")
    try:
        before_url = driver.current_url
        debug_login(f"safe_get before_url={before_url!r}")
    except WebDriverException:
        before_url = ""
        debug_login("safe_get no pudo leer current_url inicial")

    try:
        debug_login(f"safe_get intentando driver.get({url!r})")
        driver.get(url)
    except WebDriverException as exc:
        debug_login(f"safe_get driver.get fallo: {exc}")
        try:
            debug_login("safe_get intentando execute_script fallback")
            driver.execute_script(
                "var targetUrl = arguments[0]; "
                "window.setTimeout(function() { window.location.href = targetUrl; }, 0); "
                "return true;",
                url,
            )
        except WebDriverException as exc2:
            debug_login(f"safe_get fallo execute_script: {exc2}")
            raise RuntimeError(f"No se pudo iniciar navegacion por JavaScript: {exc2}") from exc2

    deadline = time.time() + timeout
    last_url = ""
    while time.time() < deadline:
        time.sleep(0.5)
        try:
            current_url = driver.current_url
            debug_login(f"safe_get current_url={current_url!r}")
            if current_url != last_url:
                last_url = current_url
            if current_url and current_url != "about:blank" and current_url != before_url:
                debug_login("safe_get navegacion completada")
                return
        except WebDriverException:
            debug_login("safe_get current_url lanzo WebDriverException")
            pass

    debug_login("safe_get termino por timeout")


def wait_for_ready(driver, timeout=20):
    debug_login(f"wait_for_ready inicio timeout={timeout}")
    try:
        WebDriverWait(driver, timeout).until(
            lambda current: current.execute_script("return document.readyState") in {"interactive", "complete"}
        )
        debug_login("wait_for_ready completo")
    except TimeoutException:
        debug_login("wait_for_ready timeout")
        pass
    except WebDriverException as exc:
        debug_login(f"wait_for_ready WebDriverException: {exc}")
        raise


def wait_for_form(driver, form_name, timeout=20):
    debug_login(f"wait_for_form inicio form={form_name!r} timeout={timeout}")
    WebDriverWait(driver, timeout).until(
        lambda current: current.execute_script(
            "return !!document.forms[arguments[0]];",
            form_name,
        )
    )
    debug_login(f"wait_for_form encontrado form={form_name!r}")


def switch_to_frame_with_form(driver, form_name, timeout=20):
    debug_login(f"switch_to_frame_with_form inicio form={form_name!r} timeout={timeout}")
    deadline = time.time() + timeout

    while time.time() < deadline:
        driver.switch_to.default_content()
        try:
            if driver.execute_script("return !!document.forms[arguments[0]];", form_name):
                debug_login(f"switch_to_frame_with_form encontrado en documento principal form={form_name!r}")
                return
        except WebDriverException:
            pass

        frames = driver.find_elements(By.TAG_NAME, "frame") + driver.find_elements(By.TAG_NAME, "iframe")
        for frame in frames:
            driver.switch_to.default_content()
            try:
                driver.switch_to.frame(frame)
                if driver.execute_script("return !!document.forms[arguments[0]];", form_name):
                    debug_login(f"switch_to_frame_with_form encontrado en frame={getattr(frame, 'get_attribute', lambda *_: '')('name')!r}")
                    return
            except WebDriverException:
                continue

        time.sleep(0.5)

    driver.switch_to.default_content()
    debug_login(f"switch_to_frame_with_form no encontro form={form_name!r}")
    raise TimeoutException(f"No encontre el formulario {form_name}.")


def wait_for_query_form_inputs(driver, timeout=20):
    WebDriverWait(driver, timeout).until(
        lambda current: current.execute_script(
            """
            var form = document.forms["TramitarPpalForm"];
            if (!form) {
                return false;
            }
            var tip = form.elements["TIP_Causa"];
            var rol = form.elements["ROL_Causa"];
            var era = form.elements["ERA_Causa"];
            return !!(tip && rol && era && tip.options && tip.options.length && era.options && era.options.length);
            """
        )
    )


def ensure_query_form_ready(driver, timeout=15):
    debug_login(f"ensure_query_form_ready inicio timeout={timeout}")
    try:
        switch_to_frame_with_form(driver, "TramitarPpalForm", timeout=2)
        wait_for_query_form_inputs(driver, timeout=min(timeout, 5))
        debug_login("ensure_query_form_ready reutilizo formulario existente")
        return "reused"
    except TimeoutException:
        debug_login("ensure_query_form_ready no encontro formulario reutilizable, navegando")
        driver.switch_to.default_content()

    safe_get(driver, POST_LOGIN_URL, timeout=timeout)
    switch_to_frame_with_form(driver, "TramitarPpalForm", timeout=timeout)
    wait_for_query_form_inputs(driver, timeout=timeout)
    debug_login("ensure_query_form_ready completo tras navegacion")
    return "navigated"


def submit_login(driver, username, password):
    debug_login("submit_login inicio")
    wait_for_form(driver, "InicioAplicacionForm", timeout=20)
    debug_login("submit_login formulario listo")
    data = {
        "username": username,
        "password": password,
        "machineName": socket.gethostname() or "0",
        "windowsUser": getpass.getuser() or "0",
        "localIp": get_local_ip(),
        "macAddress": get_mac_address(),
    }
    debug_login("submit_login llenando formulario")
    driver.execute_script(FILL_LOGIN_SCRIPT, data)
    debug_login("submit_login enviando formulario")
    driver.execute_script(SUBMIT_LOGIN_SCRIPT)
    debug_login("submit_login submit ejecutado")

# Herramientas de control de ventanas
try:
    import win32gui
    import win32con
except ImportError:
    win32gui = None
    win32con = None

class LiquidacionInfoDialog(simpledialog.Dialog):
    """Cuadro de diÃ¡logo para elegir tipo de liquidaciÃ³n y calcular cuotas."""
    def body(self, master):
        tk.Label(master, text=f"Procesando: {self.title_text}", font=("Arial", 10, "bold")).grid(row=0, column=0, columnspan=2, pady=10)
        
        tk.Label(master, text="Tipo de LiquidaciÃ³n:").grid(row=1, column=0, padx=10, pady=5, sticky="w")
        self.var_tipo = tk.StringVar(master)
        self.var_tipo.set("Peso Rojo")
        
        tk.Radiobutton(master, text="Peso Rojo", variable=self.var_tipo, value="Peso Rojo").grid(row=2, column=0, padx=20, sticky="w")
        tk.Radiobutton(master, text="Manual Excel", variable=self.var_tipo, value="Manual Excel").grid(row=3, column=0, padx=20, sticky="w")

        tk.Label(master, text="Total liquidaciÃ³n:").grid(row=4, column=0, padx=10, pady=5, sticky="w")
        self.entry_total_liq = tk.Entry(master, width=15)
        self.entry_total_liq.grid(row=4, column=1, padx=10, pady=5)
        if getattr(self, "initial_total_liq", ""):
            self.entry_total_liq.insert(0, limpiar_separadores_numericos(self.initial_total_liq))
        
        tk.Label(master, text="Monto pensiÃ³n mensual:").grid(row=5, column=0, padx=10, pady=5, sticky="w")
        self.entry_monto_pen = tk.Entry(master, width=15)
        self.entry_monto_pen.grid(row=5, column=1, padx=10, pady=5)
        if getattr(self, "initial_monto_pen", ""):
            self.entry_monto_pen.insert(0, limpiar_separadores_numericos(self.initial_monto_pen))

        tk.Label(master, text="NÂ° Cuenta LAV:").grid(row=6, column=0, padx=10, pady=10, sticky="w")
        self.entry_cuenta_lav = tk.Entry(master, width=25)
        self.entry_cuenta_lav.grid(row=6, column=1, padx=10, pady=10)
        if getattr(self, "initial_cuenta_lav", ""):
            self.entry_cuenta_lav.insert(0, self.initial_cuenta_lav)
        self.entry_cuenta_lav.focus_set()

        return self.entry_total_liq

    def __init__(self, parent, title=None, title_text="", initial_total_liq="", initial_monto_pen="", initial_cuenta_lav=""):
        self.result = None
        self.entry_total_liq = None
        self.entry_monto_pen = None
        self.entry_cuenta_lav = None
        self.var_tipo = None
        self.title_text = title_text
        self.initial_total_liq = initial_total_liq
        self.initial_monto_pen = initial_monto_pen
        self.initial_cuenta_lav = initial_cuenta_lav
        super().__init__(parent, title)

    def apply(self):
        try:
            total_liq = float(self.entry_total_liq.get().replace(",", "."))
            monto_pen = float(self.entry_monto_pen.get().replace(",", "."))
            if monto_pen > 0:
                cuotas = math.ceil(total_liq / monto_pen)
            else:
                cuotas = 0
        except (ValueError, TypeError):
            cuotas = 0
        
        cuenta_lav = self.entry_cuenta_lav.get()
        self.result = {
            "tipo": self.var_tipo.get(),
            "cuotas": cuotas,
            "cuenta_lav": cuenta_lav,
            "total_liq": self.entry_total_liq.get().strip()
        }

def extraer_datos_liquidacion_desde_pdf(pdf_path):
    """Extrae total de liquidacion y el ultimo valor de Pension reajustada desde un PDF."""
    import re
    datos = {
        "total_liq": "",
        "monto_pen": "",
        "cuenta_lav": "",
    }

    if not pdf_path or not os.path.exists(pdf_path):
        return datos

    def limpiar_texto(valor):
        texto = str(valor or "").strip()
        texto = " ".join(texto.split())
        return texto

    def normalizar_comparacion(valor):
        texto = limpiar_texto(valor).lower()
        texto = unicodedata.normalize("NFKD", texto)
        texto = "".join(ch for ch in texto if not unicodedata.combining(ch))
        return texto

    def normalizar_monto(valor):
        texto = limpiar_texto(valor)
        if not texto:
            return ""
        texto = texto.replace("$", "").replace(" ", "")
        try:
            numero = float(texto.replace(".", "").replace(",", "."))
            return "${:,.0f}".format(int(round(numero))).replace(",", ".")
        except Exception:
            return valor.strip() if isinstance(valor, str) else texto

    try:
        import fitz
        doc = fitz.open(pdf_path)
    except Exception as e:
        print(f"    [PDF] No se pudo abrir el PDF para extraer datos: {e}")
        return datos

    paginas_texto = []
    try:
        for pagina in doc:
            try:
                pagina_texto = pagina.get_text("text") or ""
            except Exception:
                pagina_texto = ""
            paginas_texto.append(pagina_texto)

            try:
                tablas = pagina.find_tables()
            except Exception as e:
                print(f"    [PDF] Error buscando tablas: {e}")
                continue

            for tabla in tablas.tables:
                try:
                    filas = tabla.extract()
                except Exception as e:
                    print(f"    [PDF] Error extrayendo tabla: {e}")
                    continue

                if not filas:
                    continue

                encabezado = [normalizar_comparacion(c) for c in filas[0] if c]
                idx_pension = None
                idx_total = None
                if any("pension reajustada" in c for c in encabezado):
                    for idx, celda in enumerate(encabezado):
                        if "pension reajustada" in celda:
                            idx_pension = idx
                        if celda == "total":
                            idx_total = idx

                    ultimo_monto_pension = ""
                    for fila in filas[1:]:
                        celdas = [limpiar_texto(c) for c in fila]
                        fila_norm = " | ".join(normalizar_comparacion(c) for c in celdas if c)

                        if not datos["cuenta_lav"] and "cuenta lav" in fila_norm:
                            for celda in celdas:
                                m_lav = re.search(r"\b([0-9]{6,})\b", celda)
                                if m_lav:
                                    datos["cuenta_lav"] = m_lav.group(1)
                                    break

                        if "total liquidacion de deuda" in fila_norm and celdas:
                            datos["total_liq"] = normalizar_monto(celdas[-1])

                        if idx_pension is not None and len(celdas) > idx_pension:
                            valor_pension = limpiar_texto(celdas[idx_pension])
                            if valor_pension and valor_pension != "-":
                                ultimo_monto_pension = valor_pension

                    if ultimo_monto_pension:
                        datos["monto_pen"] = normalizar_monto(ultimo_monto_pension)

                # Captura de LAV desde tablas de resumen que no son la de detalle.
                if not datos["cuenta_lav"]:
                    fila_completa = " | ".join(normalizar_comparacion(c) for c in filas[0] if c)
                    if "cuenta lav" in fila_completa:
                        for fila in filas:
                            for celda in fila:
                                m_lav = re.search(r"\b([0-9]{6,})\b", limpiar_texto(celda))
                                if m_lav:
                                    datos["cuenta_lav"] = m_lav.group(1)
                                    break
                            if datos["cuenta_lav"]:
                                break

        # Fallback por texto plano si la tabla no entrego los datos esperados.
        if not datos["total_liq"] or not datos["monto_pen"] or not datos["cuenta_lav"]:
            texto_completo = "\n".join(paginas_texto)
            lineas = [linea.strip() for linea in texto_completo.splitlines() if linea.strip()]
            for linea in lineas:
                linea_norm = linea.lower()
                if not datos["total_liq"] and "total liquidacion de deuda" in linea_norm:
                    m = re.search(r"total liquidaci[oó]n de deuda.*?(\$[\d\.\,]+)", linea, flags=re.I)
                    if m:
                        datos["total_liq"] = normalizar_monto(m.group(1))
                if not datos["monto_pen"] and "pension reajustada" in normalizar_comparacion(linea) and "$" in linea:
                    m = re.search(r"pensi[oó]n reajustada\s*(\$[\d\.\,]+)", linea, flags=re.I)
                    if m:
                        datos["monto_pen"] = normalizar_monto(m.group(1))
                if not datos["cuenta_lav"]:
                    m = re.search(
                        r"(?:cuenta\s*n\s*°?\s*lav|cuenta\s*lav|lav)\s*[:#°\-\s]*([0-9]{6,})",
                        normalizar_comparacion(linea),
                        flags=re.I,
                    )
                    if m:
                        datos["cuenta_lav"] = m.group(1)

        # Si el PDF viene como tabla, aseguramos que el ultimo monto de pension reajustada
        # sea el ultimo valor visible de esa columna.
        if not datos["monto_pen"]:
            try:
                ultima_pension = ""
                for pagina in fitz.open(pdf_path):
                    try:
                        tablas = pagina.find_tables()
                    except Exception:
                        continue
                    for tabla in tablas.tables:
                        try:
                            filas = tabla.extract()
                        except Exception:
                            continue
                        if not filas:
                            continue
                        encabezado = [limpiar_texto(c).lower() for c in filas[0]]
                        if "pension reajustada" not in " | ".join(encabezado):
                            continue
                        idx_pension = None
                        for idx, celda in enumerate(encabezado):
                            if "pension reajustada" in celda:
                                idx_pension = idx
                                break
                        if idx_pension is None:
                            continue
                        for fila in filas[1:]:
                            if len(fila) > idx_pension:
                                valor = limpiar_texto(fila[idx_pension])
                                if valor and valor != "-":
                                    ultima_pension = valor
                if ultima_pension:
                    datos["monto_pen"] = normalizar_monto(ultima_pension)
            except Exception as e:
                print(f"    [PDF] Error en fallback de pension reajustada: {e}")

        # Total de liquidacion: priorizar la fila total de deuda
        if not datos["total_liq"]:
            try:
                texto_completo = "\n".join(paginas_texto)
                m = re.search(r"TOTAL LIQUIDACI[oó]N DE DEUDA.*?(\$[\d\.\,]+)", texto_completo, flags=re.I | re.S)
                if m:
                    datos["total_liq"] = normalizar_monto(m.group(1))
            except Exception:
                pass

        if not datos["cuenta_lav"]:
            try:
                texto_completo = "\n".join(paginas_texto)
                m = re.search(
                    r"(?:cuenta\s*n\s*°?\s*lav|cuenta\s*lav|lav)\s*[:#°\-\s]*([0-9]{6,})",
                    normalizar_comparacion(texto_completo),
                    flags=re.I,
                )
                if m:
                    datos["cuenta_lav"] = m.group(1)
            except Exception:
                pass
    finally:
        try:
            doc.close()
        except Exception:
            pass

    print(f"    [PDF] Extraido -> total_liq={datos['total_liq']!r}, monto_pen={datos['monto_pen']!r}, cuenta_lav={datos['cuenta_lav']!r}")
    return datos

def cerrar_popups_ie_background():
    """Hilo secundario que busca y acepta alertas de IE."""
    if not win32gui: return
    for _ in range(12): # Monitorear durante 6 segundos
        def callback(hwnd, _):
            if win32gui.IsWindowVisible(hwnd):
                title = win32gui.GetWindowText(hwnd)
                if "Mensaje de la pÃ¡gina web" in title or "Internet Explorer" in title:
                    win32gui.PostMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
                    win32gui.PostMessage(hwnd, win32con.WM_KEYUP, win32con.VK_RETURN, 0)
                    #print(f"DEBUG: Pop-up '{title}' cerrado.")
        win32gui.EnumWindows(callback, None)
        time.sleep(0.5)

def obtener_doc_seguro(window_or_frame, reintentos=3, pausa=0.5):
    if not window_or_frame:
        print("    [DEBUG] window_or_frame es None en obtener_doc_seguro")
        return None

    ultimo_doc = None
    ultimo_estado = None
    for intento in range(1, reintentos + 1):
        try:
            doc = window_or_frame.Document
            ultimo_doc = doc
            ultimo_estado = getattr(doc, "readyState", None)
            if ultimo_estado in ("complete", "interactive"):
                return doc
            if intento < reintentos:
                print(f"    [DEBUG] doc.readyState es '{ultimo_estado}', esperando...")
                time.sleep(pausa)
        except Exception as e:
            print(f"    [DEBUG] Error en obtener_doc_seguro: {e}")
            if intento < reintentos:
                time.sleep(pausa)

    if ultimo_doc is not None:
        print(f"    [DEBUG] Retornando doc con readyState '{ultimo_estado}' tras espera.")
        return ultimo_doc
    return None

def obtener_body_frame(ie):
    try:
        doc = ie.Document
        
        # 1. Â¿EstÃ¡n los elementos directamente en la pÃ¡gina principal?
        try:
            if doc.getElementsByName("TIP_Causa").length > 0:
                return ie  # Devolver la ventana principal
        except: pass

        # 2. Buscar etiquetas <frame> o <iframe> en el HTML
        for tag in ["frame", "iframe"]:
            try:
                elementos = doc.getElementsByTagName(tag)
                if elementos.length > 0:
                    for i in range(elementos.length):
                        nodo = elementos[i]
                        nombre = nodo.getAttribute("name")
                        if nombre and str(nombre).lower() == "body":
                            # Intentar obtener el 'window' del frame
                            try:
                                if getattr(nodo, "contentWindow", None):
                                    return nodo.contentWindow
                            except: pass
                            
                            # Intentar obtenerlo mediante la colecciÃ³n COM
                            try:
                                return doc.frames(i)
                            except: pass
                            try:
                                return doc.frames[i]
                            except: pass
            except Exception as e1:
                print(f"    [DEBUG] Error buscando <{tag}>: {e1}")

        # 3. IteraciÃ³n clÃ¡sica de la colecciÃ³n frames
        try:
            frames = doc.frames
            length = getattr(frames, 'length', 0)
            
            try:
                res = frames("body")
                if res: return res
            except: pass

            for i in range(length):
                try:
                    f = frames(i)
                    if getattr(f, 'name', '').lower() == "body":
                        return f
                except: pass
                try:
                    f = frames[i]
                    if getattr(f, 'name', '').lower() == "body":
                        return f
                except: pass
        except Exception as e2:
            print(f"    [DEBUG] Error en doc.frames: {e2}")

    except Exception as e:
        print(f"    [DEBUG] Error general en obtener_body_frame: {e}")
    
    print("    [DEBUG] No se encontrÃ³ el frame 'body' y tampoco estaban los elementos en la pÃ¡gina principal.")
    return None

def agregar_nomenclatura(doc, valor_nomenclatura):
    try:
        selects = doc.getElementsByName("cmbNomenclaturas")
        if selects.length > 0:
            selects[0].value = str(valor_nomenclatura)
            time.sleep(0.5)
            btns = doc.getElementsByName("BtnAgregarNomenclatura")
            if btns.length > 0:
                btns[0].click()
                time.sleep(1.5)
                return True
    except: pass
    return False

def normalizar_texto(valor):
    """Deja el texto en minÃºsculas, sin tildes y con espacios limpios."""
    if valor is None:
        return ""
    texto = str(valor).strip().lower()
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(ch for ch in texto if not unicodedata.combining(ch))
    return " ".join(texto.split())

def extraer_sujeto_desde_celdas(celdas):
    """Intenta obtener el tipo de sujeto desde distintas columnas de la fila."""
    candidatos = []
    if len(celdas) > 2:
        candidatos.append(celdas[2])
    if len(celdas) > 4:
        candidatos.append(celdas[4])
    if len(celdas) > 5:
        candidatos.append(celdas[5])

    for candidato in candidatos:
        texto_norm = normalizar_texto(candidato)
        if texto_norm:
            return texto_norm
    return ""

def iterar_documentos_com(doc):
    """Devuelve el documento principal y los documentos accesibles desde frames/iframes."""
    docs = []
    if not doc:
        return docs
    docs.append(doc)
    try:
        frames = getattr(doc, "frames", None)
        if frames:
            for i in range(getattr(frames, "length", 0)):
                try:
                    frame = frames(i)
                except Exception:
                    try:
                        frame = frames[i]
                    except Exception:
                        frame = None
                if not frame:
                    continue
                try:
                    doc_frame = frame.document
                except Exception:
                    try:
                        doc_frame = frame.Document
                    except Exception:
                        doc_frame = None
                if doc_frame:
                    docs.append(doc_frame)
    except Exception:
        pass
    return docs

def buscar_elemento_por_texto(doc, textos_objetivo, tags=("input", "button", "a")):
    """Busca un control cuyo texto visible o atributo principal coincida con alguno de los textos objetivo."""
    objetivos = [normalizar_texto(t) for t in textos_objetivo if normalizar_texto(t)]
    if not objetivos or not doc:
        return None

    for doc_actual in iterar_documentos_com(doc):
        for tag in tags:
            try:
                elementos = doc_actual.getElementsByTagName(tag)
            except Exception:
                continue

            length = getattr(elementos, "length", 0)
            for i in range(length):
                try:
                    el = elementos[i]
                except Exception:
                    continue

                candidatos = []
                for attr in ("value", "innerText", "title", "alt", "name"):
                    try:
                        valor = getattr(el, attr, "")
                    except Exception:
                        valor = ""
                    if valor:
                        candidatos.append(str(valor))
                    try:
                        attr_val = el.getAttribute(attr)
                    except Exception:
                        attr_val = None
                    if attr_val:
                        candidatos.append(str(attr_val))

                texto = normalizar_texto(" ".join(candidatos))
                if not texto:
                    continue

                for objetivo in objetivos:
                    if objetivo in texto:
                        return el
    return None

def esperar_elemento_por_texto(ie, textos_objetivo, timeout=25, pausa=0.5):
    """Espera hasta encontrar un elemento por texto en la pagina o en sus frames."""
    fin = time.time() + timeout
    ultimo_doc = None
    while time.time() < fin:
        try:
            body = obtener_body_frame(ie)
            doc = obtener_doc_seguro(body)
            if doc:
                ultimo_doc = doc
                elemento = buscar_elemento_por_texto(doc, textos_objetivo)
                if elemento:
                    return elemento
        except Exception as e:
            print(f"    [DEBUG] Error esperando elemento por texto: {e}")
        time.sleep(pausa)

    if ultimo_doc:
        return buscar_elemento_por_texto(ultimo_doc, textos_objetivo)
    return None

def establecer_valor_campo(doc, nombre_campo, valor):
    """Asigna valor a un campo y dispara eventos para que IE lo reconozca."""
    if not doc:
        return False

    try:
        import json
        script = (
            "(function(){"
            f"var el=document.getElementsByName({json.dumps(nombre_campo)})[0];"
            "if(!el){return;}"
            "try{el.focus();}catch(e){}"
            f"el.value={json.dumps(str(valor))};"
            "try{if(el.fireEvent){el.fireEvent('onchange');el.fireEvent('onblur');}}catch(e){}"
            "try{if(el.blur){el.blur();}}catch(e){}"
            "})();"
        )
        try:
            doc.parentWindow.execScript(script, "JavaScript")
            return True
        except Exception:
            pass
    except Exception:
        pass

    try:
        campos = doc.getElementsByName(nombre_campo)
    except Exception:
        return False

    if not campos or getattr(campos, "length", 0) == 0:
        return False

    try:
        campo = campos[0]
        try:
            campo.focus()
        except Exception:
            pass
        campo.value = str(valor)
        try:
            campo.fireEvent("onchange")
        except Exception:
            pass
        try:
            campo.fireEvent("onblur")
        except Exception:
            pass
        try:
            campo.blur()
        except Exception:
            pass
        return True
    except Exception as e:
        print(f"    [DEBUG] No se pudo asignar {nombre_campo}: {e}")
        return False

def leer_valor_campo(doc, nombre_campo):
    """Lee el valor actual de un campo por nombre para debug."""
    if not doc:
        return None
    try:
        campos = doc.getElementsByName(nombre_campo)
        if not campos or getattr(campos, "length", 0) == 0:
            return None
        return getattr(campos[0], "value", None)
    except Exception as e:
        print(f"    [DEBUG] No se pudo leer {nombre_campo}: {e}")
        return None

def debug_formulario_rit(doc, etiqueta=""):
    """Imprime el estado del formulario de RIT antes de consultar."""
    try:
        tip = leer_valor_campo(doc, "TIP_Causa")
        rol = leer_valor_campo(doc, "ROL_Causa")
        era = leer_valor_campo(doc, "ERA_Causa")
        print(f"    [DEBUG] {etiqueta} TIP_Causa={tip!r} ROL_Causa={rol!r} ERA_Causa={era!r}")
    except Exception as e:
        print(f"    [DEBUG] Error en debug_formulario_rit: {e}")

def extraer_ruc_desde_documento(doc):
    """Busca el RUC en el documento principal o en un popup."""
    import re
    if not doc:
        return ""

    patron = re.compile(r'\b\d{1,2}\s*-\s*\d\s*-\s*\d{6,8}-[0-9Kk]\b|\b\d{1,2}\.\d{3}\.\d{3}-[0-9Kk]\b|\b\d{8,9}-[0-9Kk]\b')

    try:
        body_text = ""
        try:
            body_text = str(getattr(doc.body, "innerText", "") or "")
        except Exception:
            pass
        if body_text:
            m = patron.search(body_text)
            if m:
                return m.group(0).replace(" ", "").upper()
    except Exception:
        pass

    try:
        all_inputs = doc.getElementsByTagName("input")
        for inp in all_inputs:
            candidato = " ".join([
                str(getattr(inp, "value", "")),
                str(getattr(inp, "id", "")),
                str(getattr(inp, "name", "")),
            ])
            m = patron.search(candidato)
            if m:
                return m.group(0).replace(" ", "").upper()
    except Exception:
        pass

    try:
        all_tds = doc.getElementsByTagName("td")
        for td in all_tds:
            candidato = str(getattr(td, "innerText", "")).strip()
            m = patron.search(candidato)
            if m:
                return m.group(0).replace(" ", "").upper()
    except Exception:
        pass

    return ""

def formatear_total_liquidacion(valor):
    """Formatea un monto como moneda chilena sin decimales, por ejemplo $1.234.567."""
    if valor is None:
        return ""
    texto = str(valor).strip().replace("$", "").replace(".", "").replace(" ", "")
    if not texto:
        return ""
    try:
        numero = float(texto.replace(",", "."))
        entero = int(round(numero))
        return "${:,.0f}".format(entero).replace(",", ".")
    except Exception:
        return str(valor).strip()

def limpiar_separadores_numericos(valor):
    """Deja un valor numÃ©rico sin separadores de miles ni sÃ­mbolos de moneda."""
    if valor is None:
        return ""
    texto = str(valor).strip()
    if not texto:
        return ""
    return texto.replace("$", "").replace(".", "").replace(",", "").replace(" ", "")

def eliminar_filas_con_variables_pendientes(doc_word):
    """Elimina filas de tablas que todavÃ­a contienen variables sin reemplazar."""
    try:
        tablas = getattr(doc_word, "Tables", None)
        if not tablas:
            return 0

        filas_eliminadas = 0
        total_tablas = getattr(tablas, "Count", 0)
        for idx_tabla in range(total_tablas, 0, -1):
            try:
                tabla = tablas(idx_tabla)
                filas = tabla.Rows
                total_filas = getattr(filas, "Count", 0)
            except Exception:
                continue

            for idx_fila in range(total_filas, 0, -1):
                try:
                    fila = filas(idx_fila)
                    texto_fila = str(getattr(fila.Range, "Text", "") or "")
                except Exception:
                    continue

                if "<<" in texto_fila and ">>" in texto_fila:
                    try:
                        fila.Delete()
                        filas_eliminadas += 1
                        print(f"    [Word] Fila eliminada por variable pendiente en tabla {idx_tabla}, fila {idx_fila}.")
                    except Exception as e_del:
                        print(f"    [Word] No se pudo eliminar la fila {idx_fila} de la tabla {idx_tabla}: {e_del}")
        return filas_eliminadas
    except Exception as e:
        print(f"    [Word] Error eliminando filas con variables pendientes: {e}")
        return 0

def limpiar_pdfs_en_carpeta(ruta_carpeta):
    """Elimina los PDFs existentes para dejar la carpeta lista para una nueva iteraciÃ³n."""
    eliminados = 0
    try:
        if not ruta_carpeta or not os.path.isdir(ruta_carpeta):
            return 0
        for nombre in os.listdir(ruta_carpeta):
            if nombre.lower().endswith(".pdf"):
                ruta_pdf = os.path.join(ruta_carpeta, nombre)
                try:
                    os.remove(ruta_pdf)
                    eliminados += 1
                except Exception as e:
                    print(f"    [PDF] No se pudo eliminar '{ruta_pdf}': {e}")
    except Exception as e:
        print(f"    [PDF] Error limpiando carpeta '{ruta_carpeta}': {e}")
    return eliminados

def esperar_pdf_unico_en_carpeta(ruta_carpeta, timeout_segundos=60):
    """Espera hasta encontrar un Ãºnico PDF en la carpeta."""
    if not ruta_carpeta or not os.path.isdir(ruta_carpeta):
      return ""

def pedir_confirmacion_pdf_descargado():
    """Pregunta al usuario si el PDF ya se descargÃ³."""
    return messagebox.askyesno(
        "PDF descargado",
        "Â¿Ya se descargÃ³ el PDF de la liquidaciÃ³n?\n\n"
        "SÃ­: buscarÃ© un PDF en la carpeta seleccionada.\n"
        "No: continuarÃ© sin autollenar los datos."
    )

    fin = time.time() + timeout_segundos
    while time.time() < fin:
        try:
            pdfs = [
                os.path.join(ruta_carpeta, nombre)
                for nombre in os.listdir(ruta_carpeta)
                if nombre.lower().endswith(".pdf")
            ]
            if len(pdfs) == 1 and os.path.exists(pdfs[0]):
                return pdfs[0]
            if len(pdfs) > 1:
                pdfs.sort(key=lambda p: os.path.getmtime(p), reverse=True)
                return pdfs[0]
        except Exception:
            pass
        time.sleep(1)
    return ""

def clicar_pdf_liquidacion(ie, doc, id_escrito, reintentos=10, pausa=0.5):
    """Busca y pulsa el botÃ³n de generar PDF asociado al id de la liquidaciÃ³n."""
    if not doc or not id_escrito:
        return False

    for intento in range(1, reintentos + 1):
        try:
            doc = obtener_doc_seguro(obtener_body_frame(ie))
            if not doc:
                time.sleep(pausa)
                continue

            imagenes = doc.getElementsByTagName("img")
            print(f"    [PDF] Intento {intento}: revisando {getattr(imagenes, 'length', 0)} imÃ¡genes para id '{id_escrito}'.")
            for img in imagenes:
                try:
                    src = str(getattr(img, "src", "")).lower()
                    onclick = str(img.getAttribute("onclick") or "")
                    if "generarpdf.gif" in src and id_escrito in onclick:
                        img.click()
                        print(f"    [PDF] Click realizado sobre generarpdf.gif para id '{id_escrito}'.")
                        return True
                except Exception:
                    continue
        except Exception as e:
            print(f"    [PDF] Error buscando botÃ³n de PDF: {e}")
        time.sleep(pausa)

    print(f"    [PDF] No se encontrÃ³ el botÃ³n de PDF para id '{id_escrito}'.")
    return False

def insertar_pdf_como_imagenes_en_word(doc_word, pdf_path):
    """Convierte cada pÃ¡gina del PDF en imagen e inserta al final del documento de Word."""
    try:
        import fitz
    except Exception as e:
        print(f"    [Word] No se pudo importar fitz para PDF: {e}")
        return False

    if not pdf_path or not os.path.exists(pdf_path):
        print(f"    [Word] PDF no vÃ¡lido: {pdf_path}")
        return False

    try:
        with tempfile.TemporaryDirectory() as temp_dir:
            pdf = fitz.open(pdf_path)
            print(f"    [Word] PDF pÃ¡ginas detectadas: {len(pdf)}")
            selection = doc_word.Application.Selection
            try:
                doc_word.Activate()
                selection.Collapse(Direction=0)  # wdCollapseEnd
            except Exception as e_act:
                print(f"    [Word] No se pudo activar el documento antes de insertar: {e_act}")
            for i, pagina in enumerate(pdf):
                pixmap = pagina.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
                nombre_imagen = os.path.join(temp_dir, f"temp_page_{i}.png")
                pixmap.save(nombre_imagen)

                try:
                    selection.EndKey(Unit=6)  # wdStory
                except Exception:
                    pass
                try:
                    selection.TypeParagraph()
                except Exception:
                    pass
                try:
                    selection.InlineShapes.AddPicture(
                        FileName=nombre_imagen,
                        LinkToFile=False,
                        SaveWithDocument=True
                    )
                except Exception as e_pic:
                    print(f"    [Word] AddPicture fallÃ³ en pÃ¡gina {i}: {e_pic}")
                    try:
                        rng = doc_word.Range(doc_word.Content.End - 1, doc_word.Content.End - 1)
                        rng.InlineShapes.AddPicture(
                            FileName=nombre_imagen,
                            LinkToFile=False,
                            SaveWithDocument=True
                        )
                    except Exception as e_rng:
                        print(f"    [Word] Fallback Range fallÃ³ en pÃ¡gina {i}: {e_rng}")
                        return False
                if i < len(pdf) - 1:
                    try:
                        selection.TypeParagraph()
                    except Exception:
                        pass
        print(f"    [Word] PDF insertado como imÃ¡genes al final del documento: {pdf_path}")
        return True
    except Exception as e:
        print(f"    [Word] Error insertando PDF como imÃ¡genes: {e}")
        return False

# Variables globales para manejar el documento temporal
doc_abierto_actual = None
ruta_temp_actual = None

def cerrar_y_eliminar_word_temporal(cerrar_documento=False):
    global doc_abierto_actual, ruta_temp_actual
    try:
        if doc_abierto_actual:
            if cerrar_documento:
                try:
                    # 0 = wdDoNotSaveChanges (no guardar cambios)
                    doc_abierto_actual.Close(0)
                except Exception as e:
                    print(f"No se pudo cerrar el documento en Word: {e}")
            doc_abierto_actual = None
            ruta_temp_actual = None # Limpiar tambiÃ©n la ruta
    except Exception as e:
        print(f"Error general limpiando documento: {e}")

def sanitizar_nombre_archivo(nombre):
    """Elimina caracteres no válidos para nombres de archivo en Windows."""
    if not nombre:
        return "documento"
    invalidos = '<>:"/\\|?*'
    limpio = "".join("_" if ch in invalidos else ch for ch in str(nombre))
    limpio = limpio.strip().rstrip(".")
    return limpio or "documento"


def resolver_carpeta_aplicacion():
    """Devuelve la carpeta del proyecto o la carpeta junto al EXE empaquetado."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))

def guardar_word_sin_cerrar(nombre_base_salida):
    """Guarda el documento Word abierto en la carpeta salidas sin cerrarlo."""
    global doc_abierto_actual, ruta_temp_actual
    if not doc_abierto_actual or not ruta_temp_actual:
        print("    [Word] No hay documento abierto para guardar.")
        return ""

    try:
        carpeta_proyecto = resolver_carpeta_aplicacion()
        carpeta_salida = os.path.join(carpeta_proyecto, "salidas")
        os.makedirs(carpeta_salida, exist_ok=True)
        extension = os.path.splitext(ruta_temp_actual)[1] or ".doc"
        nombre_archivo = f"{sanitizar_nombre_archivo(nombre_base_salida)}{extension}"
        ruta_salida = os.path.join(carpeta_salida, nombre_archivo)

        if os.path.exists(ruta_salida):
            try:
                os.remove(ruta_salida)
            except Exception as e_rm:
                print(f"    [Word] No se pudo reemplazar '{ruta_salida}': {e_rm}")

        try:
            doc_abierto_actual.SaveAs2(ruta_salida)
        except Exception:
            doc_abierto_actual.SaveAs(ruta_salida)

        print(f"    [Word] Documento guardado como: {ruta_salida}")
        return ruta_salida
    except Exception as e:
        print(f"    [Word] Error guardando el documento: {e}")
        return ""


def guardar_y_cerrar_word_actual(nombre_base_salida):
    """Guarda el documento Word abierto con un nombre basado en el RIT y lo cierra."""
    global doc_abierto_actual, ruta_temp_actual
    if not doc_abierto_actual or not ruta_temp_actual:
        print("    [Word] No hay documento abierto para guardar.")
        return ""

    try:
        carpeta_base = resolver_carpeta_aplicacion()
        carpeta_salida = os.path.join(carpeta_base, "salidas")
        os.makedirs(carpeta_salida, exist_ok=True)
        extension = os.path.splitext(ruta_temp_actual)[1] or ".doc"
        nombre_archivo = f"{sanitizar_nombre_archivo(nombre_base_salida)}{extension}"
        ruta_salida = os.path.join(carpeta_salida, nombre_archivo)

        if os.path.exists(ruta_salida):
            try:
                os.remove(ruta_salida)
            except Exception as e_rm:
                print(f"    [Word] No se pudo reemplazar '{ruta_salida}': {e_rm}")

        try:
            doc_abierto_actual.SaveAs2(ruta_salida)
        except Exception:
            doc_abierto_actual.SaveAs(ruta_salida)

        print(f"    [Word] Documento guardado como: {ruta_salida}")

        try:
            doc_abierto_actual.Close(0)
        except Exception as e_close:
            print(f"    [Word] No se pudo cerrar el documento guardado: {e_close}")

        doc_abierto_actual = None
        ruta_temp_actual = None
        return ruta_salida
    except Exception as e:
        print(f"    [Word] Error guardando el documento: {e}")
        return ""

def modificar_word_abierto(texto, cuotas, datos_litigantes=None, datos_resolucion=None, cuenta_lav=None, ruc_completo=None, total_liquidacion=None):
    global doc_abierto_actual, ruta_temp_actual, ruta_plantilla_con_rd, ruta_plantilla_sin_rd
    import os
    
    try:
        
        # Determinar quÃ© plantilla usar
        if cuotas >= 3:
            ruta_doc_a_usar = ruta_plantilla_con_rd
        else:
            ruta_doc_a_usar = ruta_plantilla_sin_rd

        if not ruta_doc_a_usar or not os.path.exists(ruta_doc_a_usar):
            print(f"No se encontrÃ³ el documento plantilla: {ruta_doc_a_usar}")
            messagebox.showerror("Error", f"No se encontrÃ³ el documento plantilla: {ruta_doc_a_usar}")
            return False

        ruta_doc_a_usar = os.path.abspath(ruta_doc_a_usar)
        print(f"    [Word] Ruta confirmada: {ruta_doc_a_usar}")

        print(f"    [Word] Abriendo plantilla: {ruta_doc_a_usar}")
        os.startfile(ruta_doc_a_usar)

        try:
            word = win32com.client.GetActiveObject("Word.Application")
        except Exception:
            word = win32com.client.Dispatch("Word.Application")

        word.Visible = True
        try:
            word.WindowState = 0  # wdWindowStateNormal
        except Exception:
            pass

        # Esperar a que Word termine de abrir la plantilla antes de editarla.
        doc_word = None
        for _ in range(30):
            try:
                doc_word = word.ActiveDocument
                if doc_word and os.path.basename(getattr(doc_word, "FullName", "")).lower() == os.path.basename(ruta_doc_a_usar).lower():
                    break
            except Exception:
                doc_word = None
            time.sleep(0.5)

        if not doc_word:
            print("    [Word] No se pudo obtener el documento activo despuÃ©s de esperar la apertura.")
            return False

        doc_word.Activate()
        try:
            doc_word.Windows(1).Visible = True
            doc_word.Windows(1).Activate()
        except Exception:
            pass
        time.sleep(0.5)
        doc_abierto_actual = doc_word
        ruta_temp_actual = ruta_doc_a_usar
        
        # Debug: mostrar contenido del documento para ver qué marcadores hay
        try:
            contenido = doc_word.Content.Text
            print(f"    [DEBUG] Longitud del documento: {len(contenido)} caracteres")
            # Buscar patrones << >> en el documento
            import re
            marcadores = re.findall(r'<<[^>]+>>', contenido)
            if marcadores:
                print(f"    [DEBUG] Marcadores encontrados en la plantilla: {marcadores[:20]}")
            else:
                print(f"    [DEBUG] No se encontraron marcadores << >> en la plantilla")
                print(f"    [DEBUG] Primeros 500 caracteres del documento: {contenido[:500]}")
        except Exception as e:
            print(f"    [DEBUG] Error leyendo contenido del documento: {e}")
                
        # --- REEMPLAZO DE VARIABLES ---
        print("\n--- INICIANDO REEMPLAZO EN WORD ---")
        reemplazos = {}
        
        if datos_litigantes:
            for lit in datos_litigantes:
                s = normalizar_texto(lit["sujeto"])
                r = lit["rut"]
                n = lit["nombre"]
                
                if s == "ddo":
                    reemplazos["<<NOMBRE DEMANDADO>>"] = n
                    reemplazos["<<rut ddo>>"] = r
                    reemplazos["<<RUT DDO>>"] = r
                elif s == "dte":
                    reemplazos["<<NOMBRE DEMANDANTE>>"] = n
                    reemplazos["<<rut dte>>"] = r
                    reemplazos["<<RUT DTE>>"] = r
                elif s == "nino":
                    reemplazos["<<NOMBRE NIÃ‘O>>"] = n
                    reemplazos["<<rut niÃ±o>>"] = r
                    reemplazos["<<RUT NIÃ‘O>>"] = r
                elif s == "nina":
                    reemplazos["<<NOMBRE NIÃ‘A>>"] = n
                    reemplazos["<<rut niÃ±a>>"] = r
                    reemplazos["<<RUT NIÃ‘A>>"] = r
                elif s == "adolescente" or s == "nna":
                    reemplazos["<<NOMBRE NNA>>"] = n
                    reemplazos["<<rut nna>>"] = r
                    reemplazos["<<RUT NNA>>"] = r
                elif s == "solicitante":
                    reemplazos["<<NOMBRE SOLICITANTE>>"] = n
                    reemplazos["<<rut solicitante>>"] = r
                    reemplazos["<<RUT SOLICITANTE>>"] = r

        # Compatibilidad limpia para plantillas sin tildes.
        if "<<NOMBRE NIÃƒâ€˜O>>" in reemplazos and "<<NOMBRE NINO>>" not in reemplazos:
            reemplazos["<<NOMBRE NINO>>"] = reemplazos["<<NOMBRE NIÃƒâ€˜O>>"]
            reemplazos["<<rut nino>>"] = reemplazos.get("<<rut niÃƒÂ±o>>", reemplazos.get("<<RUT NIÃƒâ€˜O>>", ""))
            reemplazos["<<RUT NINO>>"] = reemplazos.get("<<RUT NIÃƒâ€˜O>>", reemplazos.get("<<rut niÃƒÂ±o>>", ""))
        if "<<NOMBRE NIÃƒâ€˜A>>" in reemplazos and "<<NOMBRE NINA>>" not in reemplazos:
            reemplazos["<<NOMBRE NINA>>"] = reemplazos["<<NOMBRE NIÃƒâ€˜A>>"]
            reemplazos["<<rut nina>>"] = reemplazos.get("<<rut niÃƒÂ±a>>", reemplazos.get("<<RUT NIÃƒâ€˜A>>", ""))
            reemplazos["<<RUT NINA>>"] = reemplazos.get("<<RUT NIÃƒâ€˜A>>", reemplazos.get("<<rut niÃƒÂ±a>>", ""))
        if "<<NOMBRE NNA>>" in reemplazos and "<<NOMBRE NNA>>" not in reemplazos:
            pass

        if datos_resolucion:
            folio_val = datos_resolucion.get("folio")
            fecha_val = datos_resolucion.get("fecha")
            rit_val = datos_resolucion.get("rit")
            
            # Criterio: Si hay folio (nÃºmero), usarlo. Si no, usar la fecha en su lugar.
            valor_para_folio = folio_val if folio_val else fecha_val
            
            if valor_para_folio:
                reemplazos["<<FOLIO>>"] = valor_para_folio
                reemplazos["<<folio>>"] = valor_para_folio
                
            if fecha_val:
                reemplazos["<<FECHA ING>>"] = fecha_val
                reemplazos["<<fecha ing>>"] = fecha_val
                reemplazos["<<FECHA>>"] = fecha_val
                reemplazos["<<fecha>>"] = fecha_val
                
            if rit_val:
                reemplazos["<<RIT>>"] = rit_val
                reemplazos["<<rit>>"] = rit_val
                
        # Reemplazo de cuotas
        reemplazos["<<cuota>>"] = str(cuotas)
        reemplazos["<<CUOTA>>"] = str(cuotas)
        reemplazos["<<cuotas>>"] = str(cuotas)
        reemplazos["<<CUOTAS>>"] = str(cuotas)

        if cuenta_lav:
            reemplazos["<<CUENTA_LAV>>"] = cuenta_lav

        if ruc_completo:
            reemplazos["<<RUC>>"] = ruc_completo

        if total_liquidacion:
            total_formateado = formatear_total_liquidacion(total_liquidacion)
            reemplazos["<<TOTAL LIQUIDACION>>"] = total_formateado
            reemplazos["<<total liquidacion>>"] = total_formateado



        if reemplazos:
            print(f"    Diccionario de reemplazos a aplicar:")
            for k, v in reemplazos.items():
                print(f"      {k} -> {v}")
                
            # Ejecutar reemplazo en todo el Word (Find and Replace)
            for buscar, reemplazar_por in reemplazos.items():
                if reemplazar_por:
                    try:
                        # wdReplaceAll = 2, wdFindContinue = 1
                        doc_word.Content.Find.ClearFormatting()
                        doc_word.Content.Find.Replacement.ClearFormatting()
                        
                        resultado = doc_word.Content.Find.Execute(
                            FindText=buscar,
                            MatchCase=False,
                            MatchWholeWord=False,
                            MatchWildcards=False,
                            MatchSoundsLike=False,
                            MatchAllWordForms=False,
                            Forward=True,
                            Wrap=1, # Continuar buscando desde el principio si llega al final
                            Format=False,
                            ReplaceWith=reemplazar_por,
                            Replace=2
                        )
                        print(f"    [Word] Buscando '{buscar}' -> Resultado Find: {resultado}")
                    except Exception as e_rep:
                        print(f"    [Word] Error reemplazando '{buscar}': {e_rep}")
        else:
            print("    No se recibieron datos para reemplazar en el Word.")

        filas_eliminadas = eliminar_filas_con_variables_pendientes(doc_word)
        print(f"    [Word] Filas eliminadas por variables pendientes: {filas_eliminadas}")
        print("--- FIN REEMPLAZO EN WORD ---\n")
        print("    [Word] Documento listo para revisiÃ³n manual.")

        print("    [Word] Cambios aplicados en memoria. No se guardÃ³ el archivo.")
        return True
    except Exception as e:
        print(f"Error modificando word: {e}")
    return False

def leer_excel_causas(ruta_archivo):
    import os
    datos = []
    try:
        ruta_absoluta = os.path.abspath(ruta_archivo)
        excel = win32com.client.Dispatch("Excel.Application")
        wb = excel.Workbooks.Open(ruta_absoluta)
        sheet = wb.ActiveSheet
        row = 2
        while True:
            rit = sheet.Cells(row, 1).Value
            if not rit: break
            datos.append({"rit": str(rit).strip()})
            row += 1
        wb.Close(False)
        excel.Quit()
    except Exception as e:
        print(f"Error al leer Excel: {e}")
    return datos


def preparar_causas_con_jueces(causas, jueces=None):
    """Asocia cada RIT al juez cuyo índice coincide con el último dígito del rol."""
    if jueces is None:
        return causas or []

    nombres = [str(juez or "").strip() for juez in jueces]
    if len(nombres) != 10:
        raise ValueError("Se requieren exactamente 10 jueces, indexados del 0 al 9.")

    resultado = []
    for causa in causas or []:
        rit = str(causa.get("rit", "") if isinstance(causa, dict) else causa).strip()
        partes = rit.split("-")
        if len(partes) < 2 or not partes[1].strip() or not partes[1].strip()[-1].isdigit():
            raise ValueError(f"RIT inválido o sin rol numérico: {rit}")
        indice_juez = int(partes[1].strip()[-1])
        resultado.append({"rit": rit, "juez": nombres[indice_juez]})
    return resultado


BROWSER_BACKEND = os.environ.get("SITFA_BROWSER_BACKEND", "selenium_ie_mode").strip().lower()
SITFA_URL = os.environ.get("SITFA_URL", "").strip()
SITFA_USER = os.environ.get("SITFA_USER", "").strip()
SITFA_PASSWORD = os.environ.get("SITFA_PASSWORD", "").strip()
SITFA_IE_MODE = os.environ.get("SITFA_IE_MODE", "edge").strip().lower()
SITFA_DEBUG_LOGIN = os.environ.get("SITFA_DEBUG_LOGIN", "1").strip() not in {"0", "false", "no"}


def debug_login(message):
    if SITFA_DEBUG_LOGIN:
        print(f"[LOGIN DEBUG] {message}")


def resolver_msedge_path():
    candidatos = [
        os.environ.get("EDGE_PATH", "").strip(),
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    ]
    for ruta in candidatos:
        if ruta and os.path.exists(ruta):
            return ruta
    return ""


def resolver_iedriver_path():
    candidatos_base = []
    exe_dir = os.path.dirname(os.path.abspath(sys.argv[0])) if sys.argv and sys.argv[0] else ""
    if exe_dir:
        candidatos_base.append(exe_dir)

    meipass_dir = getattr(sys, "_MEIPASS", None)
    if meipass_dir:
        candidatos_base.append(meipass_dir)

    candidatos_base.append(os.path.dirname(os.path.abspath(__file__)))

    candidatos = []
    ruta_env = os.environ.get("IEDRIVER_PATH", "").strip()
    if ruta_env:
        candidatos.append(ruta_env)

    for base_dir in candidatos_base:
        candidatos.append(os.path.join(base_dir, "IEDriverServer.exe"))
        candidatos.append(os.path.join(base_dir, "drivers", "IEDriverServer.exe"))

    for ruta in candidatos:
        if ruta and os.path.exists(ruta):
            return ruta
    return ""


def iniciar_driver_selenium_edge_ie(root):
    if not webdriver:
        raise RuntimeError(f"Selenium no esta disponible: {SELENIUM_IMPORT_ERROR}")

    debug_login("iniciar_driver_selenium_edge_ie inicio")
    debug_login(f"edge_path={resolver_msedge_path()!r}")
    debug_login(f"driver_path={resolver_iedriver_path()!r}")
    debug_login(f"SITFA_IE_MODE={SITFA_IE_MODE!r}")

    edge_path = resolver_msedge_path()
    if not edge_path:
        raise RuntimeError("No se encontro msedge.exe. Configura EDGE_PATH o instala Microsoft Edge.")

    driver_path = resolver_iedriver_path()
    if not driver_path:
        raise RuntimeError(
            "No se encontro IEDriverServer.exe. Colocalo junto al ejecutable/script o configura IEDRIVER_PATH."
        )

    options = IeOptions()
    options.page_load_strategy = os.environ.get("SITFA_PAGE_LOAD_STRATEGY", "eager")
    options.initial_browser_url = SITFA_URL or LOGIN_URL
    options.attach_to_edge_chrome = True
    if SITFA_IE_MODE != "ie":
        options.edge_executable_path = edge_path
    options.ignore_protected_mode_settings = True
    options.ignore_zoom_level = True
    options.ensure_clean_session = False
    options.require_window_focus = True
    options.native_events = False

    service = IeService(executable_path=driver_path)
    debug_login("creando webdriver.Ie")
    driver = webdriver.Ie(service=service, options=options)
    debug_login("webdriver.Ie creado")
    driver.set_page_load_timeout(10)

    return driver


def selenium_find_first(driver, selectors, timeout=5):
    fin = time.time() + timeout
    while time.time() < fin:
        for by, value in selectors:
            try:
                elementos = driver.find_elements(by, value)
                visibles = [el for el in elementos if el.is_displayed()]
                if visibles:
                    return visibles[0]
                if elementos:
                    return elementos[0]
            except Exception:
                continue
        time.sleep(0.2)
    return None


def selenium_set_input(driver, element, value):
    try:
        element.clear()
    except Exception:
        pass
    try:
        element.send_keys(value)
        return True
    except Exception:
        try:
            driver.execute_script(
                """
                arguments[0].focus();
                arguments[0].value = arguments[1];
                try {
                    if (arguments[0].fireEvent) {
                        arguments[0].fireEvent('onchange');
                        arguments[0].fireEvent('onblur');
                    } else if (document.createEvent) {
                        var evt = document.createEvent('HTMLEvents');
                        evt.initEvent('change', true, false);
                        arguments[0].dispatchEvent(evt);
                        var evt2 = document.createEvent('HTMLEvents');
                        evt2.initEvent('input', true, false);
                        arguments[0].dispatchEvent(evt2);
                    } else if (document.createEventObject) {
                        var evt3 = document.createEventObject();
                        arguments[0].fireEvent('onchange', evt3);
                        arguments[0].fireEvent('onblur', evt3);
                    }
                } catch(e) {}
                arguments[0].blur();
                """,
                element,
                value,
            )
            return True
        except Exception:
            return False


def intentar_login_sitfa(driver):
    if not (SITFA_USER and SITFA_PASSWORD):
        debug_login("intentar_login_sitfa sin credenciales")
        return False

    debug_login("intentar_login_sitfa inicio")
    try:
        debug_login(f"estado inicial url={driver.current_url!r} title={driver.title!r}")
        debug_login(f"window_handles={getattr(driver, 'window_handles', [])!r}")
    except Exception as exc:
        debug_login(f"no se pudo leer estado inicial: {exc}")

    for step in range(20):
        try:
            debug_login(f"esperando login form paso={step + 1}/20")
            wait_for_ready(driver, timeout=1)
            if driver.execute_script("return !!document.forms['InicioAplicacionForm'];"):
                debug_login("formulario de login detectado")
                break
        except Exception as exc:
            debug_login(f"espera login form fallo en paso {step + 1}: {exc}")
        time.sleep(0.5)
    else:
        debug_login("no se detecto el formulario de login luego de esperar")
        return False

    try:
        debug_login("enviando submit_login")
        submit_login(driver, SITFA_USER, SITFA_PASSWORD)
    except Exception as exc:
        print(f"[WARN] No se pudo enviar el login automatico: {exc}")
        debug_login(f"error en submit_login: {exc}\n{traceback.format_exc()}")
        return False

    wait_for_ready(driver, timeout=10)
    try:
        debug_login(f"abriendo POST_LOGIN_URL={POST_LOGIN_URL!r}")
        safe_get(driver, POST_LOGIN_URL, timeout=10)
    except Exception as exc:
        debug_login(f"fallo safe_get POST_LOGIN_URL: {exc}")
        pass
    wait_for_ready(driver, timeout=8)
    try:
        debug_login(f"login final url={driver.current_url!r} title={driver.title!r}")
        debug_login(f"window_handles={getattr(driver, 'window_handles', [])!r}")
    except Exception as exc:
        debug_login(f"no se pudo leer estado final: {exc}")
    debug_login("intentar_login_sitfa completo")
    return True


def selenium_wait_document_ready(driver, timeout=10):
    fin = time.time() + timeout
    while time.time() < fin:
        try:
            estado = driver.execute_script("return document.readyState")
            if estado in ("interactive", "complete"):
                return True
        except Exception:
            pass
        time.sleep(0.2)
    return False


def selenium_list_frame_candidates(driver):
    driver.switch_to.default_content()
    frames = []
    for tag in ("frame", "iframe"):
        try:
            frames.extend(driver.find_elements(By.TAG_NAME, tag))
        except Exception:
            continue

    def prioridad(frame):
        try:
            nombre = (frame.get_attribute("name") or "").strip().lower()
        except Exception:
            nombre = ""
        return 0 if nombre == "body" else 1

    try:
        frames.sort(key=prioridad)
    except Exception:
        pass
    return [None] + frames


def selenium_run_in_contexts(driver, callback):
    ultimo_error = None
    for frame in selenium_list_frame_candidates(driver):
        try:
            driver.switch_to.default_content()
            if frame is not None:
                driver.switch_to.frame(frame)
            resultado = callback()
            if resultado:
                return resultado
        except Exception as exc:
            ultimo_error = exc
            continue
    if ultimo_error:
        print(f"    [DEBUG] Selenium contexto sin resultado: {ultimo_error}")
    return None


def selenium_find_elements_by_name(driver, nombre):
    def _buscar():
        elementos = driver.find_elements(By.NAME, nombre)
        return elementos if elementos else None

    return selenium_run_in_contexts(driver, _buscar) or []


def selenium_form_field_value(driver, nombre_campo, valor=None):
    def _operar():
        script = """
        var form = document.forms['TramitarPpalForm'] || document.querySelector("form[name='TramitarPpalForm']");
        if (!form) return null;
        var el = form.elements ? form.elements[arguments[0]] : null;
        if (!el && form[arguments[0]]) el = form[arguments[0]];
        if (!el) {
            var candidatos = form.getElementsByTagName ? form.getElementsByTagName('*') : [];
            for (var i = 0; i < candidatos.length; i++) {
                var cand = candidatos[i];
                var candName = cand.name || cand.id || '';
                if (candName === arguments[0]) {
                    el = cand;
                    break;
                }
            }
        }
        if (!el) return null;
        if (arguments.length > 1) {
            try { el.value = arguments[1]; } catch(e) {}
            try { el.setAttribute('value', arguments[1]); } catch(e) {}
            try {
                if (el.fireEvent) {
                    el.fireEvent('onchange');
                    el.fireEvent('onblur');
                }
            } catch(e) {}
        }
        try {
            return el.value != null ? String(el.value) : '';
        } catch(e) {
            return '';
        }
        """
        if valor is None:
            return driver.execute_script(script, nombre_campo)
        return driver.execute_script(script, nombre_campo, str(valor))

    return selenium_run_in_contexts(driver, _operar)


def selenium_set_field_value(driver, nombre_campo, valor):
    try:
        return str(selenium_form_field_value(driver, nombre_campo, valor) or "") == str(valor)
    except Exception:
        return False


def selenium_read_field_value(driver, nombre_campo):
    try:
        return selenium_form_field_value(driver, nombre_campo)
    except Exception:
        return None


def debug_formulario_rit_selenium(driver, etiqueta=""):
    try:
        tip = selenium_read_field_value(driver, "TIP_Causa")
        rol = selenium_read_field_value(driver, "ROL_Causa")
        era = selenium_read_field_value(driver, "ERA_Causa")
        print(f"    [DEBUG] {etiqueta} TIP_Causa={tip!r} ROL_Causa={rol!r} ERA_Causa={era!r}")
    except Exception as e:
        print(f"    [DEBUG] Error en debug_formulario_rit_selenium: {e}")


def selenium_find_element_by_text(driver, textos_objetivo, tags=("input", "button", "a"), timeout=25, pausa=0.5):
    import json
    objetivos = [normalizar_texto(t) for t in textos_objetivo if normalizar_texto(t)]
    if not objetivos:
        return None
    fin = time.time() + timeout

    SCRIPT_FIND_BY_TEXT = """
    var textos = arguments[0];
    var tags = arguments[1];
    function normalizar(valor) {
        return (valor || '').toLowerCase();
    }
    for (var t = 0; t < tags.length; t++) {
        var elements = document.getElementsByTagName(tags[t]);
        for (var i = 0; i < elements.length; i++) {
            var el = elements[i];
            var attrs = [
                el.value || '',
                el.textContent || '',
                el.title || '',
                el.alt || '',
                el.name || '',
                el.innerText || ''
            ].join(' ');
            var texto = normalizar(attrs);
            for (var j = 0; j < textos.length; j++) {
                if (texto.indexOf(textos[j]) >= 0) {
                    return i + '|' + tags[t];
                }
            }
        }
    }
    return '';
    """

    def _buscar_en_contexto():
        try:
            selenium_wait_document_ready(driver, timeout=2)
        except Exception:
            pass
        resultado = driver.execute_script(SCRIPT_FIND_BY_TEXT, objetivos, list(tags))
        if resultado:
            idx_str, tag = resultado.split("|", 1)
            idx = int(idx_str)
            elementos = driver.find_elements(By.TAG_NAME, tag)
            if idx < len(elementos):
                return elementos[idx]
        return None

    while time.time() < fin:
        try:
            elemento = selenium_run_in_contexts(driver, _buscar_en_contexto)
            if elemento:
                return elemento
        except Exception as e:
            print(f"    [DEBUG] Error en selenium_find_element_by_text: {e}")
        time.sleep(pausa)
    return None


def selenium_click_element(driver, element):
    try:
        driver.execute_script("arguments[0].click();", element)
        return True
    except Exception:
        pass
    try:
        element.click()
        return True
    except Exception:
        pass

    def _click_en_contexto():
        try:
            driver.execute_script("arguments[0].click();", element)
            return True
        except Exception:
            pass
        try:
            element.click()
            return True
        except Exception:
            return None

    resultado = selenium_run_in_contexts(driver, _click_en_contexto)
    return bool(resultado)


def selenium_accept_alerts(driver, timeout=6):
    fin = time.time() + timeout
    acepto = False
    while time.time() < fin:
        try:
            alerta = driver.switch_to.alert
            texto = alerta.text
            print(f"    [DEBUG] Alerta detectada: {texto}")
            alerta.accept()
            acepto = True
            time.sleep(0.3)
        except NoAlertPresentException:
            time.sleep(0.2)
        except Exception as e:
            print(f"    [DEBUG] Error aceptando alerta Selenium: {e}")
            time.sleep(0.2)
    return acepto


def selenium_extraer_ruc(driver):
    patron = re.compile(r"\b\d{1,2}\s*-\s*\d\s*-\s*\d{6,8}-[0-9Kk]\b|\b\d{1,2}\.\d{3}\.\d{3}-[0-9Kk]\b|\b\d{8,9}-[0-9Kk]\b")

    def _buscar():
        textos = []
        try:
            body_text = driver.execute_script("return document.body ? document.body.innerText || '' : '';")
            if body_text:
                textos.append(str(body_text))
        except Exception:
            pass
        try:
            valores = driver.execute_script(
                """
                return Array.from(document.querySelectorAll('input')).map(function(el) {
                    return [el.value || '', el.id || '', el.name || ''].join(' ');
                });
                """
            )
            textos.extend(valores or [])
        except Exception:
            pass
        try:
            celdas = driver.execute_script(
                "return Array.from(document.querySelectorAll('td')).map(function(el) { return el.innerText || ''; });"
            )
            textos.extend(celdas or [])
        except Exception:
            pass

        for texto in textos:
            m = patron.search(str(texto))
            if m:
                return m.group(0).replace(" ", "").upper()
        return None

    return selenium_run_in_contexts(driver, _buscar) or ""


def selenium_obtener_info_liquidacion(driver):
    def _buscar():
        return driver.execute_script(
            """
            function normalizar(valor) {
                return (valor || '').toLowerCase()
                    .normalize('NFD')
                    .replace(/[\\u0300-\\u036f]/g, '');
            }
            const tds = Array.from(document.querySelectorAll('td.texto, td'));
            for (const td of tds) {
                const texto = normalizar(td.innerText || '');
                if (!texto.includes('liquidacion de deuda')) {
                    continue;
                }
                const tr = td.closest('tr');
                if (!tr) {
                    continue;
                }
                const celdas = Array.from(tr.querySelectorAll('td')).map(td => (td.innerText || '').trim());
                const checkbox = tr.querySelector('input[type="checkbox"]');
                if (checkbox) {
                    checkbox.click();
                }
                return {
                    id_escrito: checkbox ? (checkbox.value || '') : '',
                    folio: celdas[0] || '',
                    fecha: celdas[4] || '',
                    celdas: celdas
                };
            }
            return null;
            """
        )

    return selenium_run_in_contexts(driver, _buscar) or {}


def selenium_clicar_pdf_liquidacion(driver, id_escrito, reintentos=12, pausa=0.5):
    if not id_escrito:
        return False

    script = """
    var id_escrito = arguments[0];
    var imagenes = document.getElementsByTagName("img");
    for (var i = 0; i < imagenes.length; i++) {
        var img = imagenes[i];
        var src = (img.src || '').toLowerCase();
        var onclick = img.getAttribute("onclick") || '';
        if (src.indexOf('generarpdf.gif') >= 0 && onclick.indexOf(id_escrito) >= 0) {
            try { img.click(); } catch(e) {}
            return true;
        }
    }
    return false;
    """

    for intento in range(1, reintentos + 1):
        print(f"    [PDF] Intento {intento}: buscando imagenes para id '{id_escrito}'.")
        try:
            if driver.execute_script(script, id_escrito):
                print(f"    [PDF] Click realizado sobre generarpdf.gif para id '{id_escrito}'.")
                return True
        except Exception as e:
            print(f"    [PDF] Error en intento {intento}: {e}")
        time.sleep(pausa)

    print(f"    [PDF] No se encontro el boton de PDF para id '{id_escrito}'.")
    return False


def selenium_click_radio_grales(driver):
    try:
        return bool(driver.execute_script("""
            var radios = document.querySelectorAll("input[name='seleccion']");
            for (var i = 0; i < radios.length; i++) {
                if (radios[i].value === 'Grales') {
                    radios[i].click();
                    return true;
                }
            }
            return false;
        """))
    except Exception:
        return False


def selenium_agregar_nomenclatura(driver, valor_nomenclatura):
    script = """
    var select = document.querySelector("select[name='cmbNomenclaturas'], [name='cmbNomenclaturas']");
    if (!select) return false;
    select.value = arguments[0];
    var event = document.createEvent('HTMLEvents');
    event.initEvent('change', true, false);
    select.dispatchEvent(event);
    var btn = document.querySelector("[name='BtnAgregarNomenclatura']");
    if (btn) {
        try { btn.click(); } catch(e) {}
        return true;
    }
    return false;
    """
    try:
        return bool(driver.execute_script(script, str(valor_nomenclatura)))
    except Exception:
        return False


def selenium_seleccionar_juez(driver, apellido_juez):
    if not apellido_juez:
        return False
    script = """
    var select = document.querySelector("select[name='CRR_IdFuncionarioTrmPend'], [name='CRR_IdFuncionarioTrmPend']");
    if (!select) return false;
    var apellido = arguments[0].toLowerCase();
    for (var i = 0; i < select.options.length; i++) {
        if ((select.options[i].text || '').toLowerCase().indexOf(apellido) >= 0) {
            select.value = select.options[i].value;
            try {
                var event = document.createEvent('HTMLEvents');
                event.initEvent('change', true, false);
                select.dispatchEvent(event);
            } catch(e) {}
            return true;
        }
    }
    return false;
    """
    try:
        return bool(driver.execute_script(script, apellido_juez))
    except Exception:
        return False


def selenium_set_glosa(driver, texto_final):
    try:
        return bool(driver.execute_script("""
            var campo = document.getElementById('GLS_Ref');
            if (!campo) {
                var refs = document.getElementsByName('referencia');
                if (refs.length > 0) campo = refs[0];
            }
            if (!campo) return false;
            campo.focus();
            campo.value = arguments[0];
            try {
                if (campo.fireEvent) {
                    campo.fireEvent('onchange');
                    campo.fireEvent('onblur');
                } else if (document.createEvent) {
                    var evt = document.createEvent('HTMLEvents');
                    evt.initEvent('change', true, false);
                    campo.dispatchEvent(evt);
                    var evt2 = document.createEvent('HTMLEvents');
                    evt2.initEvent('input', true, false);
                    campo.dispatchEvent(evt2);
                } else if (document.createEventObject) {
                    var evt3 = document.createEventObject();
                    campo.fireEvent('onchange', evt3);
                    campo.fireEvent('onblur', evt3);
                }
            } catch(e) {}
            campo.blur();
            return true;
        """, texto_final))
    except Exception:
        return False


def selenium_click_actuacion(driver):
    script = """
    var elements = document.querySelectorAll("input[name='irAccionTramitarT'], button[name='irAccionTramitarT']");
    if (!elements.length) return false;
    for (var i = 0; i < elements.length; i++) {
        var el = elements[i];
        var text = (el.value || '') + ' ' + (el.innerText || '') + ' ' + (el.textContent || '') + ' ' + (el.name || '');
        if (text.toLowerCase().indexOf('actu') >= 0) {
            var form = el.form || document.forms['TramitarPpalForm'];
            if (form && form.elements['FLG_Actuacion']) {
                try { form.elements['FLG_Actuacion'].value = '1'; } catch(e) {}
            }
            try { el.removeAttribute('disabled'); } catch(e) {}
            try { el.disabled = false; } catch(e) {}
            try { el.scrollIntoView(true); } catch(e) {}
            try { el.click(); } catch(e) {}
            return true;
        }
    }
    return false;
    """
    try:
        return bool(driver.execute_script(script))
    except Exception as exc:
        debug_login(f"selenium_click_actuacion fallo: {exc}")
        return False


def selenium_override_show_modal_dialog(driver):
    try:
        driver.execute_script(
            "window.showModalDialog = function(url, arg, opt) { window.open(url, '_blank'); return 'mock'; };"
        )
        return True
    except Exception:
        return False


def selenium_wait_new_window(driver, handles_previos, timeout=10):
    fin = time.time() + timeout
    previos = set(handles_previos)
    while time.time() < fin:
        try:
            actuales = driver.window_handles
            nuevos = [h for h in actuales if h not in previos]
            if nuevos:
                return nuevos[-1]
        except Exception:
            pass
        time.sleep(0.2)
    return ""


def selenium_extraer_litigantes_popup(driver):
    datos_litigantes = []

    def _extraer():
        tabla = driver.execute_script(
            """
            let tabla = document.getElementById('TablaLitigantes');
            if (!tabla) {
                tabla = Array.from(document.querySelectorAll('table')).find(function(tbl) {
                    return (tbl.id || '').toLowerCase() === 'tablalitigantes';
                }) || null;
            }
            if (!tabla) {
                return null;
            }
            return Array.from(tabla.querySelectorAll('tr')).map(function(tr) {
                return Array.from(tr.querySelectorAll('td')).map(function(td) {
                    return (td.innerText || '').trim();
                });
            });
            """
        )

        if not tabla:
            return None

        encontrados = []
        for celdas in tabla:
            if len(celdas) < 6:
                continue
            sujeto = (celdas[2] or "").upper()
            rut = celdas[3] or ""
            nombre = (celdas[5] or "").replace("\r\n", " ").replace("\n", " ").strip()
            sujeto_limpio = sujeto.replace(".", "").strip()
            if any(s in sujeto for s in ["DTE", "DDO", "ADOLESCENTE", "NIÑA", "NIÑO", "NNA", "SOLICITANTE"]):
                if rut and not any(x["rut"] == rut for x in encontrados):
                    encontrados.append({"sujeto": sujeto_limpio, "rut": rut, "nombre": nombre})

        if encontrados:
            return encontrados

        for celdas in tabla:
            if len(celdas) < 4:
                continue
            candidato_sujeto = " ".join([celdas[i] for i in (2, 4, 5) if len(celdas) > i and celdas[i]])
            sujeto_norm = normalizar_texto(candidato_sujeto)
            if not any(s in sujeto_norm for s in ["dte", "ddo", "adolescente", "nina", "nino", "nna", "solicitante"]):
                continue
            rut = celdas[3] if len(celdas) > 3 else ""
            nombre = (celdas[5] if len(celdas) > 5 else "").replace("\r\n", " ").replace("\n", " ").strip()
            if rut and not any(x["rut"] == rut for x in encontrados):
                encontrados.append({
                    "sujeto": candidato_sujeto.replace(".", "").strip().upper(),
                    "rut": rut,
                    "nombre": nombre,
                })

        return encontrados if encontrados else None

    resultado = selenium_run_in_contexts(driver, _extraer)
    if resultado:
        datos_litigantes.extend(resultado)
    return datos_litigantes


def selenium_ir_al_buscador(driver):
    script = """
    var links = document.querySelectorAll('a');
    for (var i = 0; i < links.length; i++) {
        var href = links[i].href || '';
        var text = (links[i].innerText || '').toLowerCase();
        if (text.indexOf('volver al buscador') >= 0 || text.indexOf('consulta rol tramitar') >= 0 || href.indexOf('opMenu=ConsultaRolTramitar') >= 0) {
            try { links[i].click(); } catch(e) {}
            return true;
        }
    }
    return false;
    """
    try:
        return bool(driver.execute_script(script))
    except Exception:
        return False


def automatizar_selenium_edge_ie_con_rit(archivo_excel=None, carpeta_pdfs=None, causas=None, jueces=None):
    root = tk.Tk()
    root.withdraw()

    if not archivo_excel:
        archivo_excel = filedialog.askopenfilename(title="Selecciona el archivo Excel", filetypes=[("Excel files", "*.xlsx *.xls")])
    if not archivo_excel:
        return

    if not carpeta_pdfs:
        carpeta_pdfs = filedialog.askdirectory(title="Selecciona la carpeta donde se descargara el PDF")
    if not carpeta_pdfs:
        messagebox.showinfo("Cancelado", "No se selecciono una carpeta para los PDF. Proceso terminado.")
        return

    global ruta_plantilla_con_rd, ruta_plantilla_sin_rd

    ruta_plantilla_con_rd, ruta_plantilla_sin_rd = resolver_rutas_plantillas()

    if not ruta_plantilla_con_rd:
        ruta_plantilla_con_rd = filedialog.askopenfilename(
            title="Selecciona la plantilla Word 'Notifica liq CON RD.doc'",
            filetypes=[("Word Documents", "*.doc *.docx")]
        )
        if not ruta_plantilla_con_rd:
            messagebox.showinfo("Cancelado", "Seleccion de plantilla 'con RD' cancelada. Proceso terminado.")
            return

    if not ruta_plantilla_sin_rd:
        ruta_plantilla_sin_rd = filedialog.askopenfilename(
            title="Selecciona la plantilla Word 'Notifica liq sin rd.doc'",
            filetypes=[("Word Documents", "*.doc *.docx")]
        )
        if not ruta_plantilla_sin_rd:
            messagebox.showinfo("Cancelado", "Seleccion de plantilla 'sin RD' cancelada. Proceso terminado.")
            return

    if causas is None:
        causas = leer_excel_causas(archivo_excel)
    causas = preparar_causas_con_jueces(causas, jueces)
    if not causas:
        messagebox.showerror("Error", "No se encontraron datos en el Excel.")
        return

    driver = None
    try:
        driver = iniciar_driver_selenium_edge_ie(root)
        login_intentado = intentar_login_sitfa(driver)
        if login_intentado:
            print("    [DEBUG] Se intento login automatico con credenciales de .env.")
        else:
            print("    [DEBUG] No se detecto formulario de login automatico; se mantiene flujo asistido.")

        try:
            estado_form = ensure_query_form_ready(driver, timeout=15)
            print(f"    [DEBUG] Formulario de consulta preparado con modo: {estado_form}")
        except Exception as exc:
            print(f"    [WARN] No se pudo preparar el formulario de consulta automaticamente: {exc}")

        for index, item in enumerate(causas):
            if automation_should_stop():
                print("[SITFA] Detención solicitada desde la interfaz.")
                break
            rit_completo = item["rit"]
            apellido_juez = item["juez"].lower()

            limpiar_pdfs_en_carpeta(carpeta_pdfs)

            try:
                p = rit_completo.split("-")
                tipo_causa, rit_num, anio_causa = p[0].strip().upper(), p[1].strip(), p[2].strip()
            except Exception as e:
                print(f"[{rit_completo}] Error al separar RIT: {e}")
                continue

            try:
                if not selenium_set_field_value(driver, "TIP_Causa", tipo_causa):
                    print(f"[{rit_completo}] No se pudo cargar TIP_Causa.")
                    continue
                if not selenium_set_field_value(driver, "ERA_Causa", anio_causa):
                    print(f"[{rit_completo}] No se pudo cargar ERA_Causa.")
                    continue
                if not selenium_set_field_value(driver, "ROL_Causa", rit_num):
                    print(f"[{rit_completo}] No se pudo cargar ROL_Causa.")
                    continue

                time.sleep(1.5)

                debug_formulario_rit_selenium(driver, etiqueta=f"[{rit_completo}] antes de Consultar")

                btn_consultar = selenium_find_element_by_text(driver, ["Consultar"], tags=("input", "button"), timeout=10)
                if not btn_consultar:
                    print(f"[{rit_completo}] No se encontro el boton 'Consultar'.")
                    continue

                if not selenium_click_element(driver, btn_consultar):
                    print(f"[{rit_completo}] No se pudo pulsar 'Consultar'.")
                    continue

                time.sleep(1.0)
                selenium_accept_alerts(driver, timeout=6)
                time.sleep(1.0)

                driver.switch_to.default_content()
                try:
                    selenium_wait_document_ready(driver, timeout=10)
                except Exception:
                    pass

                print(f"    [DEBUG] [{rit_completo}] Buscando pestaña Liquidacion...")
                tab_liquidacion = selenium_find_element_by_text(
                    driver,
                    ["Liquidación", "Liquidacion"],
                    tags=("a", "td", "div", "span"),
                    timeout=10,
                    pausa=0.5,
                )
                if tab_liquidacion:
                    print(f"    [DEBUG] [{rit_completo}] Click en pestaña Liquidacion")
                    selenium_click_element(driver, tab_liquidacion)
                    time.sleep(2.0)
                    driver.switch_to.default_content()
                    try:
                        selenium_wait_document_ready(driver, timeout=10)
                    except Exception:
                        pass

                print(f"    [DEBUG] [{rit_completo}] Buscando boton Resolucion...")
                resolucion_btn = selenium_find_element_by_text(
                    driver,
                    ["Resolución", "Resolucion"],
                    timeout=25,
                    pausa=0.5,
                )
                if not resolucion_btn:
                    print(f"[{rit_completo}] No se encontro el boton 'Resolucion'.")
                    try:
                        page_text = driver.execute_script("return document.body ? document.body.innerText.substring(0, 500) : '';")
                        print(f"    [DEBUG] Contenido parcial de la pagina: {page_text[:200]}")
                    except Exception as e:
                        print(f"    [DEBUG] No se pudo obtener contenido de la pagina: {e}")
                    continue

                print(f"    [DEBUG] Boton Resolucion encontrado")
                if not selenium_click_element(driver, resolucion_btn):
                    print(f"[{rit_completo}] No se pudo pulsar 'Resolucion'.")
                    continue
                time.sleep(2.0)

                info_liq = selenium_obtener_info_liquidacion(driver)
                id_escrito = info_liq.get("id_escrito", "")
                folio_liq = info_liq.get("folio", "") if str(info_liq.get("folio", "")).isdigit() else ""
                fecha_liq = info_liq.get("fecha", "") if "/" in str(info_liq.get("fecha", "")) else ""

                if info_liq.get("celdas"):
                    print(f"\n    [DEBUG] Fila encontrada: {info_liq['celdas']}")
                    print(f"    [DEBUG] Folio extraido: '{folio_liq}', Fecha extraida: '{fecha_liq}'")

                if id_escrito:
                    time.sleep(1.0)
                    if not selenium_clicar_pdf_liquidacion(driver, id_escrito, reintentos=12, pausa=0.5):
                        print(f"[{rit_completo}] No se pudo disparar el PDF de la liquidacion.")

                datos_pdf = {"total_liq": "", "monto_pen": "", "cuenta_lav": ""}
                pdf_path = ""
                if pedir_confirmacion_pdf_descargado():
                    pdfs = [
                        os.path.join(carpeta_pdfs, nombre)
                        for nombre in os.listdir(carpeta_pdfs)
                        if nombre.lower().endswith(".pdf")
                    ]
                    if pdfs:
                        pdfs.sort(key=lambda p: os.path.getmtime(p), reverse=True)
                        pdf_path = pdfs[0]
                        print(f"[{rit_completo}] PDF detectado: {pdf_path}")
                        datos_pdf = extraer_datos_liquidacion_desde_pdf(pdf_path)
                        print(
                            f"    [PDF] Datos extraidos -> total: {datos_pdf['total_liq']}, "
                            f"monto pension: {datos_pdf['monto_pen']}, cuenta LAV: {datos_pdf['cuenta_lav']}"
                        )
                    else:
                        print(f"[{rit_completo}] El usuario confirmo PDF, pero no hay ninguno en la carpeta.")
                else:
                    print(f"[{rit_completo}] El usuario indico que el PDF no esta listo. Se continua sin autollenado.")

                info = solicitar_datos_liquidacion(
                    root,
                    title="Liquidacion",
                    title_text=rit_completo,
                    initial_total_liq=datos_pdf.get("total_liq", ""),
                    initial_monto_pen=datos_pdf.get("monto_pen", ""),
                    initial_cuenta_lav=datos_pdf.get("cuenta_lav", "")
                )
                if not info:
                    print(f"[{rit_completo}] Se cancelo el dialogo de cuotas.")
                    continue
                if automation_should_stop():
                    print("[SITFA] Detención solicitada desde la interfaz.")
                    break
                tipo, cuotas = info["tipo"], info["cuotas"]
                total_liquidacion = info.get("total_liq", "")

                selenium_click_radio_grales(driver)
                time.sleep(0.8)

                noms = ["385034", "1015"] if tipo == "Peso Rojo" and cuotas >= 3 else []
                if tipo == "Peso Rojo" and cuotas < 3:
                    noms = ["385034"]
                elif tipo != "Peso Rojo" and cuotas >= 3:
                    noms = ["1015"]
                elif tipo != "Peso Rojo" and cuotas < 3:
                    noms = ["1025"]

                for nomenclatura in noms:
                    selenium_agregar_nomenclatura(driver, nomenclatura)
                    time.sleep(1.0)

                selenium_seleccionar_juez(driver, apellido_juez)

                texto_final = "Notif Liq RD" if cuotas >= 3 else "Notif Liq sin RD"
                selenium_set_glosa(driver, texto_final)

                debug_login("intentando click en Actuacion")
                try:
                    actuacion_clicked = selenium_click_actuacion(driver)
                    debug_login(f"click Actuacion resultado={actuacion_clicked!r}")
                    selenium_accept_alerts(driver, timeout=3)
                    wait_for_ready(driver, timeout=5)
                except Exception as exc:
                    print(f"    [DEBUG] No se pudo hacer click en Actuacion: {exc}")
                    debug_login(f"error click Actuacion: {exc}\n{traceback.format_exc()}")

                ruc_completo = selenium_extraer_ruc(driver)
                if ruc_completo:
                    print(f"    [DEBUG] RUC extraido desde pagina principal: {ruc_completo}")
                else:
                    print("    [DEBUG] No se pudo extraer RUC desde la pagina principal.")

                print(f"\n--- EXTRAYENDO DATOS PARA {rit_completo} ---")
                datos_litigantes = []
                popup_handle = ""
                main_handle = driver.current_window_handle
                handles_previos = list(driver.window_handles)

                selenium_override_show_modal_dialog(driver)
                btn_mod = selenium_find_element_by_text(driver, ["Modificación", "Modificacion"], tags=("button", "input"), timeout=10)
                if btn_mod and selenium_click_element(driver, btn_mod):
                    popup_handle = selenium_wait_new_window(driver, handles_previos, timeout=10)
                else:
                    print("No se encontro el boton de Modificacion.")

                if popup_handle:
                    try:
                        driver.switch_to.window(popup_handle)
                        selenium_wait_document_ready(driver, timeout=8)
                        time.sleep(2.0)
                        if not ruc_completo:
                            ruc_completo = selenium_extraer_ruc(driver)
                            if ruc_completo:
                                print(f"    [DEBUG] RUC encontrado como respaldo en popup: {ruc_completo}")
                        datos_litigantes = selenium_extraer_litigantes_popup(driver)
                    except Exception as e_popup:
                        print(f"Error extrayendo datos del litigante: {e_popup}")
                    finally:
                        try:
                            driver.close()
                        except Exception:
                            pass
                        driver.switch_to.window(main_handle)
                else:
                    print("    [DEBUG] No se detecto ventana popup de litigantes.")

                for litigante in datos_litigantes:
                    print(f" -> Encontrado: [{litigante['sujeto']}] RUT: {litigante['rut']} | Nombre: {litigante['nombre']}")
                print("------------------------------------------\n")

                datos_res = {"folio": folio_liq, "fecha": fecha_liq, "rit": rit_completo}
                modificar_word_abierto(
                    texto=f"{rit_completo}: {texto_final}",
                    cuotas=cuotas,
                    datos_litigantes=datos_litigantes,
                    datos_resolucion=datos_res,
                    cuenta_lav=info["cuenta_lav"],
                    ruc_completo=ruc_completo,
                    total_liquidacion=total_liquidacion
                )

                if pdf_path and doc_abierto_actual:
                    insertar_pdf_como_imagenes_en_word(doc_abierto_actual, pdf_path)

                ruta_guardada = guardar_y_cerrar_word_actual(rit_completo)
                if ruta_guardada:
                    print(f"    [Word] Archivo final listo: {ruta_guardada}")

                print(f"    [DEBUG] Intentando volver al buscador para {rit_completo}...")
                if not selenium_ir_al_buscador(driver):
                    print(f"    [DEBUG] No se encontro el enlace 'Volver al buscador' para {rit_completo}.")

                if index < len(causas) - 1:
                    if not messagebox.askokcancel(
                        "Causa Finalizada",
                        f"Se termino de procesar: {rit_completo}\n\n"
                        "Presiona OK para continuar con la siguiente causa o Cancelar para detener el proceso."
                    ):
                        cerrar_y_eliminar_word_temporal(cerrar_documento=False)
                        break
                else:
                    cerrar_y_eliminar_word_temporal(cerrar_documento=False)

            except Exception as e:
                print(f"Error procesando {rit_completo}: {e}")
                cerrar_y_eliminar_word_temporal(cerrar_documento=False)

        messagebox.showinfo("Proceso Terminado", "El procesamiento por lotes ha finalizado.")
        cerrar_y_eliminar_word_temporal(cerrar_documento=False)
    finally:
        if driver:
            try:
                driver.quit()
            except Exception:
                pass

def automatizar_ie_con_rit(archivo_excel=None, carpeta_pdfs=None, causas=None, jueces=None):
    root = tk.Tk()
    root.withdraw()

    if not archivo_excel:
        archivo_excel = filedialog.askopenfilename(title="Selecciona el archivo Excel", filetypes=[("Excel files", "*.xlsx *.xls")])
    if not archivo_excel: 
        return

    if not carpeta_pdfs:
        carpeta_pdfs = filedialog.askdirectory(title="Selecciona la carpeta donde se descargarÃ¡ el PDF")
    if not carpeta_pdfs:
        messagebox.showinfo("Cancelado", "No se seleccionÃ³ una carpeta para los PDF. Proceso terminado.")
        return

    # --- Plantillas Word ---
    global ruta_plantilla_con_rd, ruta_plantilla_sin_rd, doc_abierto_actual, ruta_temp_actual

    ruta_plantilla_con_rd, ruta_plantilla_sin_rd = resolver_rutas_plantillas()

    if not ruta_plantilla_con_rd:
        ruta_plantilla_con_rd = filedialog.askopenfilename(
            title="Selecciona la plantilla Word 'Notifica liq CON RD.doc'",
            filetypes=[("Word Documents", "*.doc *.docx")]
        )
        if not ruta_plantilla_con_rd:
            messagebox.showinfo("Cancelado", "SelecciÃ³n de plantilla 'con RD' cancelada. Proceso terminado.")
            return

    if not ruta_plantilla_sin_rd:
        ruta_plantilla_sin_rd = filedialog.askopenfilename(
            title="Selecciona la plantilla Word 'Notifica liq sin rd.doc'",
            filetypes=[("Word Documents", "*.doc *.docx")]
        )
        if not ruta_plantilla_sin_rd:
            messagebox.showinfo("Cancelado", "SelecciÃ³n de plantilla 'sin RD' cancelada. Proceso terminado.")
            return
    # -----------------------------------

    if causas is None:
        causas = leer_excel_causas(archivo_excel)
    causas = preparar_causas_con_jueces(causas, jueces)
    if not causas:
        messagebox.showerror("Error", "No se encontraron datos en el Excel.")
        return

    ie = None
    try:
        shell = win32com.client.Dispatch("Shell.Application")
        for window in shell.Windows():
            if "iexplore.exe" in getattr(window, 'FullName', '').lower() and "SITFAWEB" in window.LocationURL:
                ie = window
                break
        if not ie:
            messagebox.showerror("Error", "No se encontrÃ³ SITFA en IE.")
            return
        ie.Visible = 1
    except: return

    for index, item in enumerate(causas):
        if automation_should_stop():
            print("[SITFA] Detención solicitada desde la interfaz.")
            break
        rit_completo = item["rit"]
        apellido_juez = item["juez"].lower()

        limpiar_pdfs_en_carpeta(carpeta_pdfs)
        
        try:
            p = rit_completo.split("-")
            tipo_causa, rit_num, anio_causa = p[0].strip().upper(), p[1].strip(), p[2].strip()
        except Exception as e:
            print(f"[{rit_completo}] Error al separar RIT: {e}")
            continue

        try:
            # A. BÃºsqueda inicial
            body = obtener_body_frame(ie)
            doc = obtener_doc_seguro(body)
            if not doc:
                print(f"[{rit_completo}] No se pudo obtener el documento del frame 'body'.")
                continue

            form_rit = None
            try:
                form_rit = doc.forms("TramitarPpalForm")
            except Exception:
                try:
                    form_rit = doc.TramitarPpalForm
                except Exception:
                    form_rit = None

            if not establecer_valor_campo(doc, "TIP_Causa", tipo_causa):
                print(f"[{rit_completo}] No se pudo cargar TIP_Causa.")
                continue
            if not establecer_valor_campo(doc, "ERA_Causa", anio_causa):
                print(f"[{rit_completo}] No se pudo cargar ERA_Causa.")
                continue
            if not establecer_valor_campo(doc, "ROL_Causa", rit_num):
                print(f"[{rit_completo}] No se pudo cargar ROL_Causa.")
                continue

            if form_rit:
                try:
                    form_rit.TIP_Causa.value = tipo_causa
                    form_rit.ERA_Causa.value = anio_causa
                    form_rit.ROL_Causa.value = rit_num
                    print(f"    [DEBUG] Valores cargados en TramitarPpalForm directamente.")
                except Exception as e:
                    print(f"    [DEBUG] No se pudo escribir en TramitarPpalForm: {e}")

            time.sleep(0.5)
            debug_formulario_rit(doc, etiqueta=f"[{rit_completo}] antes de Consultar")

            # El aviso "Existe 1 Trámite pendiente..." aparece justo después de
            # pulsar Consultar, así que hay que tener activo el monitor antes de
            # ese click para que la confirmación no bloquee el flujo.
            threading.Thread(target=cerrar_popups_ie_background, daemon=True).start()

            for inp in doc.getElementsByTagName("input"):
                if inp.getAttribute("value") == "Consultar":
                    try:
                        print(f"    [DEBUG] Boton Consultar detectado: value={getattr(inp, 'value', None)!r} name={getattr(inp, 'name', None)!r} type={getattr(inp, 'type', None)!r}")
                    except Exception:
                        pass
                    inp.click()
                    time.sleep(1.0)
                    break
            
            # B. Esperar boton Resolucion
            resolucion_btn = esperar_elemento_por_texto(
                ie,
                ["Resolución", "Resolucion"],
                timeout=25,
                pausa=0.5,
            )
            
            if not resolucion_btn:
                print(f"[{rit_completo}] No se encontro el boton 'Resolucion'.")
                continue

            print(f"    [DEBUG] [{rit_completo}] Haciendo click en Resolucion...")
            resolucion_btn.click()
            time.sleep(0.5)
            
            threading.Thread(target=cerrar_popups_ie_background, daemon=True).start()
            time.sleep(2.0)

            print(f"    [DEBUG] [{rit_completo}] Obteniendo documento del frame body...")
            # C. Checkbox y PDF por ID
            doc = obtener_doc_seguro(obtener_body_frame(ie))
            if not doc:
                print(f"[{rit_completo}] No se pudo obtener el documento del frame body despues de Resolucion.")
                continue
            
            print(f"    [DEBUG] [{rit_completo}] Documento obtenido. Buscando pestaña Liquidacion...")
            
            # Buscar y hacer click en la pestaña "Liquidación"
            tab_liquidacion = None
            try:
                links = doc.getElementsByTagName("a")
                print(f"    [DEBUG] [{rit_completo}] Encontrados {links.length} enlaces en el documento")
                for i in range(links.length):
                    link = links[i]
                    texto = str(link.innerText or link.textContent or "").strip()
                    href = str(link.getAttribute("href") or "")
                    print(f"    [DEBUG] Enlace {i}: texto='{texto}', href='{href}'")
                    if "liquidación" in texto.lower() or "liquidacion" in texto.lower():
                        tab_liquidacion = link
                        print(f"    [DEBUG] [{rit_completo}] Pestaña Liquidacion encontrada: '{texto}'")
                        break
            except Exception as e:
                print(f"    [DEBUG] Error buscando pestaña Liquidacion: {e}")
            
            if tab_liquidacion:
                print(f"    [DEBUG] [{rit_completo}] Haciendo click en pestaña Liquidacion...")
                tab_liquidacion.click()
                time.sleep(2.0)
                doc = obtener_doc_seguro(obtener_body_frame(ie))
                if not doc:
                    print(f"[{rit_completo}] No se pudo obtener el documento despues de click en Liquidacion.")
                    continue
                print(f"    [DEBUG] [{rit_completo}] Documento recargado despues de click en Liquidacion")
            else:
                print(f"    [DEBUG] [{rit_completo}] No se encontro pestaña Liquidacion, continuando...")
            
            print(f"    [DEBUG] [{rit_completo}] Buscando fila de liquidacion de deuda...")
            id_escrito = None
            folio_liq = ""
            fecha_liq = ""
            
            tds = doc.getElementsByTagName("td")
            print(f"    [DEBUG] [{rit_completo}] Encontrados {tds.length} elementos TD en el documento")
            for i in range(tds.length):
                texto_td = str(tds[i].innerText or "").lower()
                clase_td = str(tds[i].className or "")
                if "liquidaci" in texto_td and "deuda" in texto_td and clase_td == "texto":
                    print(f"    [DEBUG] [{rit_completo}] TD encontrado en posicion {i}: '{texto_td[:50]}...'")
                    tr = tds[i]
                    while tr and tr.tagName.upper() != "TR": tr = tr.parentNode
                    if tr:
                        # Extraer celdas de esta fila específica
                        try:
                            celdas_tr = tr.getElementsByTagName("td")
                            textos_celdas = [str(celdas_tr[c].innerText).strip() for c in range(celdas_tr.length)]
                            print(f"\n    [DEBUG] Fila encontrada: {textos_celdas}")
                            
                            if len(textos_celdas) >= 6:
                                # Buscar folio (primer número en la fila)
                                folio_candidato = ""
                                fecha_candidata = ""
                                for idx, celda in enumerate(textos_celdas):
                                    if celda.isdigit() and not folio_candidato:
                                        folio_candidato = celda
                                        print(f"    [DEBUG] Folio encontrado en posicion {idx}: '{celda}'")
                                    elif "/" in celda and len(celda) >= 8 and not fecha_candidata:
                                        fecha_candidata = celda
                                        print(f"    [DEBUG] Fecha encontrada en posicion {idx}: '{celda}'")
                                
                                if folio_candidato.isdigit():
                                    folio_liq = folio_candidato
                                if "/" in fecha_candidata:
                                    fecha_liq = fecha_candidata
                                    
                            print(f"    [DEBUG] Folio extraído: '{folio_liq}', Fecha extraída: '{fecha_liq}'")
                        except Exception as e_celdas:
                            print(f"    [DEBUG] Error extrayendo folio/fecha de la fila: {e_celdas}")

                        for inp in tr.getElementsByTagName("input"):
                            if inp.getAttribute("type") == "checkbox":
                                id_escrito = inp.value
                                print(f"    [DEBUG] Checkbox encontrado con id_escrito: '{id_escrito}'")
                                inp.click()
                                break
                    if id_escrito: 
                        print(f"    [DEBUG] Usando primera liquidación encontrada (más reciente)")
                        break

            if id_escrito:
                time.sleep(1.0)
                clic_pdf = clicar_pdf_liquidacion(ie, doc, id_escrito, reintentos=12, pausa=0.5)
                if not clic_pdf:
                    print(f"[{rit_completo}] No se pudo disparar el PDF de la liquidaciÃ³n.")
            
            # D. DiÃ¡logo de cuotas
            datos_pdf = {"total_liq": "", "monto_pen": "", "cuenta_lav": ""}
            pdf_path = ""
            if pedir_confirmacion_pdf_descargado():
                pdfs = [
                    os.path.join(carpeta_pdfs, nombre)
                    for nombre in os.listdir(carpeta_pdfs)
                    if nombre.lower().endswith(".pdf")
                ]
                if pdfs:
                    pdfs.sort(key=lambda p: os.path.getmtime(p), reverse=True)
                    pdf_path = pdfs[0]
                    print(f"[{rit_completo}] PDF detectado: {pdf_path}")
                    datos_pdf = extraer_datos_liquidacion_desde_pdf(pdf_path)
                    print(
                        f"    [PDF] Datos extraÃ­dos -> total: {datos_pdf['total_liq']}, "
                        f"monto pensiÃ³n: {datos_pdf['monto_pen']}, cuenta LAV: {datos_pdf['cuenta_lav']}"
                    )
                else:
                    print(f"[{rit_completo}] El usuario confirmÃ³ PDF, pero no hay ninguno en la carpeta. Se mostrarÃ¡ el diÃ¡logo vacÃ­o.")
            else:
                print(f"[{rit_completo}] El usuario indicÃ³ que el PDF no estÃ¡ listo. Se continÃºa sin autollenado.")

            info = solicitar_datos_liquidacion(
                root,
                title="LiquidaciÃ³n",
                title_text=rit_completo,
                initial_total_liq=datos_pdf.get("total_liq", ""),
                initial_monto_pen=datos_pdf.get("monto_pen", ""),
                initial_cuenta_lav=datos_pdf.get("cuenta_lav", "")
            )
            if not info:
                print(f"[{rit_completo}] Se cancelÃ³ el diÃ¡logo de cuotas.")
                continue
            tipo, cuotas = info['tipo'], info['cuotas']
            total_liquidacion = info.get('total_liq', "")

            # E. Nomenclaturas y Juez
            for r in doc.getElementsByName("seleccion"):
                if r.getAttribute("value") == "Grales": r.click(); break
            time.sleep(0.8)

            noms = []
            if tipo == "Peso Rojo":
                noms = ["385034", "1015"] if cuotas >= 3 else ["385034"]
            else:
                noms = ["1015"] if cuotas >= 3 else ["1025"]

            for n in noms:
                doc = obtener_doc_seguro(obtener_body_frame(ie))
                agregar_nomenclatura(doc, n)

            if apellido_juez:
                doc = obtener_doc_seguro(obtener_body_frame(ie))
                select_juez = doc.getElementsByName("CRR_IdFuncionarioTrmPend")
                if select_juez.length > 0:
                    opciones = select_juez[0].getElementsByTagName("option")
                    for opt in opciones:
                        if apellido_juez in str(opt.innerText).lower():
                            select_juez[0].value = opt.value
                            break
            
            # F. Glosa y Word
            texto_final = "Notif Liq RD" if cuotas >= 3 else "Notif Liq sin RD"
            doc = obtener_doc_seguro(obtener_body_frame(ie))
            campo_ref = doc.getElementById("GLS_Ref")
            if not campo_ref:
                refs = doc.getElementsByName("referencia")
                if refs.length > 0: campo_ref = refs[0]
            if campo_ref: campo_ref.value = texto_final

            # RUC: intentar capturarlo desde la pÃ¡gina principal antes de abrir el popup de litigantes.
            ruc_completo = extraer_ruc_desde_documento(doc)
            if ruc_completo:
                print(f"    [DEBUG] RUC extraÃ­do desde pÃ¡gina principal: {ruc_completo}")
            else:
                print("    [DEBUG] No se pudo extraer RUC desde la pÃ¡gina principal.")

            # --- OBTENER DATOS DE CABECERA (SUJETO, RUT, NOMBRE) ---
            print(f"\n--- EXTRAYENDO DATOS PARA {rit_completo} ---")
            datos_litigantes = []
            try:
                # Evitar bloqueo de showModalDialog
                try:
                    doc.parentWindow.execScript("window.showModalDialog = function(url, arg, opt) { window.open(url, '_blank'); return 'mock'; };", "JavaScript")
                except Exception: pass

                btn_mod = None
                btns = doc.getElementsByTagName("button")
                for i in range(btns.length):
                    btn_text = str(btns[i].innerText or "").lower()
                    btn_onclick = str(btns[i].getAttribute("onclick") or "").lower()
                    if "modificación" in btn_text or "modificacion" in btn_text or "showpopupcabecera" in btn_onclick:
                        btn_mod = btns[i]
                        print(f"    [DEBUG] Botón Modificación encontrado: texto='{btn_text}', onclick='{btn_onclick}'")
                        break
                
                if btn_mod:
                    btn_mod.click()
                    time.sleep(5.0) # Dar tiempo a que cargue la ventana
                    
                    ventana_popup = None
                    shell_windows = win32com.client.Dispatch("Shell.Application").Windows()
                    for window in shell_windows:
                        try:
                            if "iexplore.exe" in getattr(window, 'FullName', '').lower():
                                url = getattr(window, 'LocationURL', '').lower()
                                title = getattr(window, 'LocationName', '').lower()
                                if "despliegue de pop up" in title:
                                    ventana_popup = window
                                    break
                                if "popup" in url and ventana_popup is None:
                                    ventana_popup = window
                                if ("cabecera" in title or "modifica" in title) and ventana_popup is None:
                                    ventana_popup = window
                        except:
                            pass
                    if not ventana_popup:
                        # Ãšltimo intento: encontrar cualquier pestaÃ±a cuyo tÃ­tulo mencione pop up.
                        for window in shell_windows:
                            try:
                                if "iexplore.exe" in getattr(window, 'FullName', '').lower():
                                    title = getattr(window, 'LocationName', '').lower()
                                    if "pop up" in title or "despliegue" in title:
                                        ventana_popup = window
                                        break
                            except:
                                pass
                        
                    if ventana_popup:
                        d = ventana_popup.Document
                        print(f"    [DEBUG] Ventana popup detectada: title='{getattr(ventana_popup, 'LocationName', '')}', url='{getattr(ventana_popup, 'LocationURL', '')}'")
                        try:
                            print(f"    [DEBUG] readyState inicial popup: {d.readyState}")
                        except Exception:
                            pass
                        # Ya no buscamos el RUC acÃ¡ salvo como respaldo si no se obtuvo antes.
                        if not ruc_completo:
                            try:
                                ruc_popup = extraer_ruc_desde_documento(d)
                                if ruc_popup:
                                    ruc_completo = ruc_popup
                                    print(f"    [DEBUG] RUC encontrado como respaldo en popup: {ruc_completo}")
                            except Exception as e:
                                print(f"Error al buscar RUC en ventana emergente: {e}")

                        # Dump resumido del DOM para depurar quÃ© estÃ¡ leyendo realmente.
                        try:
                            all_tables_dbg = d.getElementsByTagName("table")
                            print(f"    [DEBUG] dump tables count: {getattr(all_tables_dbg, 'length', 0)}")
                            for i in range(min(getattr(all_tables_dbg, "length", 0), 5)):
                                try:
                                    tabla_dbg = all_tables_dbg[i]
                                    filas_dbg = tabla_dbg.getElementsByTagName("tr")
                                    print(f"    [DEBUG] table[{i}] id='{str(getattr(tabla_dbg, 'id', ''))}' rows={getattr(filas_dbg, 'length', 0)}")
                                    for j in range(min(getattr(filas_dbg, "length", 0), 5)):
                                        try:
                                            tds_dbg = filas_dbg[j].getElementsByTagName("td")
                                            textos_dbg = [str(tds_dbg[k].innerText).strip() for k in range(min(getattr(tds_dbg, "length", 0), 8))]
                                            print(f"    [DEBUG] table[{i}] row[{j}] -> {textos_dbg}")
                                        except Exception as e_row:
                                            print(f"    [DEBUG] error row table[{i}][{j}]: {e_row}")
                                except Exception as e_table:
                                    print(f"    [DEBUG] error table[{i}]: {e_table}")
                        except Exception as e_dump:
                            print(f"    [DEBUG] error dumping popup DOM: {e_dump}")

                        for _ in range(10):
                            if d.readyState == "complete": break
                            time.sleep(0.5)
                        print(f"    [DEBUG] readyState popup tras espera: {getattr(d, 'readyState', 'desconocido')}")
                            
                        try:
                            tabla = None
                            filas = []
                            docs_para_buscar = iterar_documentos_com(d)
                            print(f"    [DEBUG] documentos accesibles para buscar: {len(docs_para_buscar)}")

                            for idx_doc, doc_busqueda in enumerate(docs_para_buscar):
                                try:
                                    tablas = doc_busqueda.getElementsByTagName("table")
                                    print(f"    [DEBUG] doc[{idx_doc}] tables={getattr(tablas, 'length', 0)}")
                                    for i in range(getattr(tablas, "length", 0)):
                                        try:
                                            tabla_id = str(getattr(tablas[i], "id", ""))
                                        except Exception:
                                            tabla_id = ""
                                        if tabla_id.lower() == "tablalitigantes":
                                            tabla = tablas[i]
                                            filas = tabla.getElementsByTagName("tr")
                                            print(f"    [DEBUG] TablaLitigantes encontrada en doc[{idx_doc}] con {getattr(filas, 'length', 0)} filas")
                                            break
                                    if tabla:
                                        break
                                except Exception as e_doc:
                                    print(f"    [DEBUG] Error buscando en doc[{idx_doc}]: {e_doc}")

                            if not tabla:
                                print("    [DEBUG] No se encontrÃ³ 'TablaLitigantes' en ningÃºn documento accesible.")
                                continue

                            print(f"    [DEBUG] total <tr> a revisar: {getattr(filas, 'length', 0)}")

                            for tr_idx in range(filas.length):
                                tds = filas[tr_idx].getElementsByTagName("td")
                                if tds.length < 6:
                                    continue

                                celdas = [str(tds[c].innerText).strip() for c in range(tds.length)]

                                # Estructura tablaLitigantes:
                                # 0 Est, 1 Sujeto, 2 Rut, 3 Pers, 4 Nombre, 5 Caj
                                sujeto = celdas[1].upper() if len(celdas) > 1 else ""
                                rut = celdas[2] if len(celdas) > 2 else ""
                                nombre = celdas[4] if len(celdas) > 4 else ""

                                nombre = nombre.replace('\r\n', ' ').replace('\n', ' ').replace('&nbsp;', ' ').strip()
                                sujeto_limpio = sujeto.replace('.', '').replace('&nbsp;', '').strip()

                                # Detectar los tipos que realmente aparecen en esta tabla
                                if any(s in sujeto_limpio for s in ["DTE", "DDO", "ADOLESCENTE", "NIÑO", "NIÑA", "NNA", "SOLICITANTE"]):
                                    if rut and not any(l['rut'] == rut for l in datos_litigantes):
                                        datos_litigantes.append({
                                            "sujeto": sujeto_limpio,
                                            "rut": rut,
                                            "nombre": nombre
                                        })
                                        print(f"    [DEBUG] Litigante detectado: sujeto='{sujeto_limpio}', rut='{rut}', nombre='{nombre}'")
                        except Exception as e:
                            print(f"Error extrayendo datos del litigante: {e}")

                        # Fallback: si la detección exacta falló o cambió el orden de columnas,
                        # recorrer nuevamente la tabla con comparaciones normalizadas.
                        try:
                            if not datos_litigantes:
                                filas_fallback = tabla.getElementsByTagName("tr") if tabla else d.getElementsByTagName("tr")
                                for tr_idx in range(filas_fallback.length):
                                    tds = filas_fallback[tr_idx].getElementsByTagName("td")
                                    if tds.length < 4:
                                        continue
                                    celdas = [str(tds[c].innerText).strip() for c in range(tds.length)]
                                    
                                    # Estructura tablaLitigantes: 1 Sujeto, 2 Rut, 4 Nombre
                                    candidato_sujeto = celdas[1] if len(celdas) > 1 else ""
                                    rut = celdas[2] if len(celdas) > 2 else ""
                                    nombre = celdas[4] if len(celdas) > 4 else ""
                                    
                                    sujeto_norm = normalizar_texto(candidato_sujeto)
                                    if not any(s in sujeto_norm for s in ["dte", "ddo", "adolescente", "niño", "niña", "nina", "nino", "nna", "solicitante"]):
                                        continue
                                    
                                    nombre = nombre.replace('\r\n', ' ').replace('\n', ' ').replace('&nbsp;', ' ').strip()
                                    if rut and not any(l['rut'] == rut for l in datos_litigantes):
                                        datos_litigantes.append({
                                            "sujeto": candidato_sujeto.replace('.', '').replace('&nbsp;', '').strip().upper(),
                                            "rut": rut,
                                            "nombre": nombre
                                        })
                                        print(f"    [DEBUG] Litigante detectado por fallback: sujeto='{candidato_sujeto}', rut='{rut}', nombre='{nombre}'")
                        except Exception as e_fb:
                            print(f"Error en fallback de litigantes: {e_fb}")
                            
                        # No cerrar la pestaÃ±a si es un tab compartido; solo liberar si realmente era una ventana separada.
                        try:
                            if "popup" in str(getattr(ventana_popup, "LocationURL", "")).lower():
                                ventana_popup.Quit()
                        except:
                            pass
                else:
                    print("No se encontrÃ³ el botÃ³n de ModificaciÃ³n.")
                    
            except Exception as e_dbg:
                print(f"Error en extracciÃ³n de cabecera: {e_dbg}")
                
            # Mostrar los datos limpios extraÃ­dos
            for litigante in datos_litigantes:
                print(f" -> Encontrado: [{litigante['sujeto']}] RUT: {litigante['rut']} | Nombre: {litigante['nombre']}")
            print("------------------------------------------\n")

            datos_res = {"folio": folio_liq, "fecha": fecha_liq, "rit": rit_completo}
            modificar_word_abierto(
                texto=f"{rit_completo}: {texto_final}",
                cuotas=cuotas,
                datos_litigantes=datos_litigantes,
                datos_resolucion=datos_res,
                cuenta_lav=info['cuenta_lav'],
                ruc_completo=ruc_completo,
                total_liquidacion=total_liquidacion
            )

            if pdf_path and doc_abierto_actual is not None:
                insertar_pdf_como_imagenes_en_word(doc_abierto_actual, pdf_path)

            ruta_guardada = guardar_word_sin_cerrar(rit_completo)
            if ruta_guardada:
                print(f"    [Word] Archivo guardado: {ruta_guardada}")
            
            # G. Volver al buscador
            print(f"    [DEBUG] Intentando volver al buscador para {rit_completo}...")
            links = ie.Document.getElementsByTagName("a")
            for link in links:
                if "opMenu=ConsultaRolTramitar" in str(link.getAttribute("href")):
                    print(f"    [DEBUG] Encontrado enlace 'Volver al buscador'. Clickeando...")
                    link.click()
                    time.sleep(3.0) # Aumentar el tiempo de espera
                    # Esperar a que la página de IE esté lista después de la navegación
                    for _ in range(10): # Esperar hasta 5 segundos
                        if ie.ReadyState == 4: # READYSTATE_COMPLETE
                            print(f"    [DEBUG] IE ReadyState: COMPLETE después de volver al buscador.")
                            break
                        time.sleep(0.5)
                    else:
                        print(f"    [DEBUG] IE ReadyState no COMPLETE después de volver al buscador en {rit_completo}.")
                    break
            else:
                print(f"    [DEBUG] No se encontró el enlace 'Volver al buscador' para {rit_completo}.")
            
            # Pausa de confirmación para seguir con la siguiente causa
            if index < len(causas) - 1:
                if not messagebox.askokcancel("Causa Finalizada", f"Se terminó de procesar: {rit_completo}\n\nPresiona OK para continuar con la siguiente causa o Cancelar para detener el proceso."):
                    cerrar_y_eliminar_word_temporal(cerrar_documento=True)
                    break
                else:
                    # Usuario presionó OK, cerrar el Word actual antes de continuar
                    try:
                        if doc_abierto_actual is not None:
                            messagebox.showinfo("Cerrando documento", f"Se cerrará el documento Word de {rit_completo} al aceptar este mensaje.")
                            doc_abierto_actual.Close(0)
                            print(f"    [Word] Documento cerrado después de confirmación")
                            doc_abierto_actual = None
                            ruta_temp_actual = None
                    except Exception as e_close:
                        print(f"    [Word] Error cerrando documento: {e_close}")
            else:
                # Última causa, cerrar el Word
                try:
                    if doc_abierto_actual is not None:
                        messagebox.showinfo("Cerrando documento", f"Se cerrará el documento Word de {rit_completo} al aceptar este mensaje.")
                        doc_abierto_actual.Close(0)
                        print(f"    [Word] Documento cerrado (última causa)")
                        doc_abierto_actual = None
                        ruta_temp_actual = None
                except Exception as e_close:
                    print(f"    [Word] Error cerrando documento: {e_close}")

        except Exception as e:
            print(f"Error procesando {rit_completo}: {e}")
            cerrar_y_eliminar_word_temporal(cerrar_documento=False)

    messagebox.showinfo("Proceso Terminado", "El procesamiento por lotes ha finalizado.")
    cerrar_y_eliminar_word_temporal(cerrar_documento=False)

def automatizar_con_backend(archivo_excel=None, carpeta_pdfs=None, causas=None, jueces=None):
    automatizar_ie_con_rit(
        archivo_excel=archivo_excel,
        carpeta_pdfs=carpeta_pdfs,
        causas=causas,
        jueces=jueces,
    )

if __name__ == "__main__":
    automatizar_con_backend()



