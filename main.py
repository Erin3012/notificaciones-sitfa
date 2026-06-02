import tkinter as tk
from tkinter import simpledialog, messagebox, filedialog
import win32com.client
import time
import threading
import math
import os
import sys
import unicodedata
import tempfile

# Variables globales para las rutas de los documentos de plantilla seleccionados
ruta_plantilla_con_rd = None
ruta_plantilla_sin_rd = None

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

def guardar_y_cerrar_word_actual(nombre_base_salida):
    """Guarda el documento Word abierto con un nombre basado en el RIT y lo cierra."""
    global doc_abierto_actual, ruta_temp_actual
    if not doc_abierto_actual or not ruta_temp_actual:
        print("    [Word] No hay documento abierto para guardar.")
        return ""

    try:
        carpeta_base = os.path.dirname(os.path.abspath(__file__))
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
    global doc_abierto_actual, ruta_temp_actual
    import os
    
    try:
        global ruta_plantilla_con_rd, ruta_plantilla_sin_rd, doc_abierto_actual, ruta_temp_actual
        
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
            juez = sheet.Cells(row, 2).Value
            if not rit: break
            datos.append({"rit": str(rit).strip(), "juez": str(juez).strip() if juez else ""})
            row += 1
        wb.Close(False)
        excel.Quit()
    except Exception as e:
        print(f"Error al leer Excel: {e}")
    return datos

def automatizar_ie_con_rit():
    root = tk.Tk()
    root.withdraw()

    archivo_excel = filedialog.askopenfilename(title="Selecciona el archivo Excel", filetypes=[("Excel files", "*.xlsx *.xls")])
    if not archivo_excel: 
        return

    carpeta_pdfs = filedialog.askdirectory(title="Selecciona la carpeta donde se descargarÃ¡ el PDF")
    if not carpeta_pdfs:
        messagebox.showinfo("Cancelado", "No se seleccionÃ³ una carpeta para los PDF. Proceso terminado.")
        return

    # --- Plantillas Word ---
    global ruta_plantilla_con_rd, ruta_plantilla_sin_rd

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

    causas = leer_excel_causas(archivo_excel)
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

            resolucion_btn.click()
            time.sleep(2.0)

            # C. Checkbox y PDF por ID
            doc = obtener_doc_seguro(obtener_body_frame(ie))
            id_escrito = None
            folio_liq = ""
            fecha_liq = ""
            
            tds = doc.getElementsByTagName("td")
            for i in range(tds.length):
                if "liquidaciÃ³n de deuda" in str(tds[i].innerText).lower() and str(tds[i].className) == "texto":
                    tr = tds[i]
                    while tr and tr.tagName.upper() != "TR": tr = tr.parentNode
                    if tr:
                        # Extraer celdas de esta fila especÃ­fica
                        try:
                            celdas_tr = tr.getElementsByTagName("td")
                            textos_celdas = [str(celdas_tr[c].innerText).strip() for c in range(celdas_tr.length)]
                            print(f"\n    [DEBUG] Fila encontrada: {textos_celdas}")
                            
                            if len(textos_celdas) >= 6:
                                # SegÃºn el HTML aportado, Folio estÃ¡ en td 0 y Fecha en td 4
                                folio_candidato = textos_celdas[0]
                                fecha_candidata = textos_celdas[4]
                                
                                if folio_candidato.isdigit():
                                    folio_liq = folio_candidato
                                if "/" in fecha_candidata:
                                    fecha_liq = fecha_candidata
                                    
                            print(f"    [DEBUG] Folio extraÃ­do: '{folio_liq}', Fecha extraÃ­da: '{fecha_liq}'")
                        except Exception as e_celdas:
                            print(f"    [DEBUG] Error extrayendo folio/fecha de la fila: {e_celdas}")

                        for inp in tr.getElementsByTagName("input"):
                            if inp.getAttribute("type") == "checkbox":
                                id_escrito = inp.value
                                inp.click()
                                break
                    if id_escrito: break

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

            info = LiquidacionInfoDialog(
                root,
                title="LiquidaciÃ³n",
                title_text=rit_completo,
                initial_total_liq=datos_pdf.get("total_liq", ""),
                initial_monto_pen=datos_pdf.get("monto_pen", ""),
                initial_cuenta_lav=datos_pdf.get("cuenta_lav", "")
            ).result
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
                    btn_text = str(btns[i].innerText).lower()
                    if "modificaciÃ³n" in btn_text:
                        btn_mod = btns[i]
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

                                # Estructura real del HTML:
                                # 0 SRCeI, 1 Fecha, 2 Tipo Lit., 3 RUT, 4 Persona, 5 Nombre o RazÃ³n Social
                                sujeto = celdas[2].upper() if len(celdas) > 2 else ""
                                rut = celdas[3] if len(celdas) > 3 else ""
                                nombre = celdas[5] if len(celdas) > 5 else ""

                                nombre = nombre.replace('\r\n', ' ').replace('\n', ' ').strip()
                                sujeto_limpio = sujeto.replace('.', '').strip()

                                # Detectar los tipos que realmente aparecen en esta tabla
                                if any(s in sujeto for s in ["DTE", "DDO", "ADOLESCENTE", "NIÃ‘A", "NIÃ‘O", "NNA", "SOLICITANTE"]):
                                    if rut and not any(l['rut'] == rut for l in datos_litigantes):
                                        datos_litigantes.append({
                                            "sujeto": sujeto_limpio,
                                            "rut": rut,
                                            "nombre": nombre
                                        })
                                        print(f"    [DEBUG] Litigante detectado: sujeto='{sujeto_limpio}', rut='{rut}', nombre='{nombre}'")
                        except Exception as e:
                            print(f"Error extrayendo datos del litigante: {e}")

                        # Fallback: si la detecciÃ³n exacta fallÃ³ o cambiÃ³ el orden de columnas,
                        # recorrer nuevamente la tabla con comparaciones normalizadas.
                        try:
                            if not datos_litigantes:
                                filas_fallback = tabla.getElementsByTagName("tr") if tabla else d.getElementsByTagName("tr")
                                for tr_idx in range(filas_fallback.length):
                                    tds = filas_fallback[tr_idx].getElementsByTagName("td")
                                    if tds.length < 4:
                                        continue
                                    celdas = [str(tds[c].innerText).strip() for c in range(tds.length)]
                                    candidato_sujeto = " ".join(
                                        [celdas[i] for i in (2, 4, 5) if len(celdas) > i and celdas[i]]
                                    )
                                    sujeto_norm = normalizar_texto(candidato_sujeto)
                                    if not any(s in sujeto_norm for s in ["dte", "ddo", "adolescente", "nina", "nino", "nna", "solicitante"]):
                                        continue
                                    rut = celdas[3] if len(celdas) > 3 else ""
                                    nombre = celdas[5] if len(celdas) > 5 else ""
                                    nombre = nombre.replace('\r\n', ' ').replace('\n', ' ').strip()
                                    if rut and not any(l['rut'] == rut for l in datos_litigantes):
                                        datos_litigantes.append({
                                            "sujeto": candidato_sujeto.replace('.', '').strip().upper(),
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

            if pdf_path and doc_abierto_actual:
                insertar_pdf_como_imagenes_en_word(doc_abierto_actual, pdf_path)

            ruta_guardada = guardar_y_cerrar_word_actual(rit_completo)
            if ruta_guardada:
                print(f"    [Word] Archivo final listo: {ruta_guardada}")
            
            # G. Volver al buscador
            print(f"    [DEBUG] Intentando volver al buscador para {rit_completo}...")
            links = ie.Document.getElementsByTagName("a")
            for link in links:
                if "opMenu=ConsultaRolTramitar" in str(link.getAttribute("href")):
                    print(f"    [DEBUG] Encontrado enlace 'Volver al buscador'. Clickeando...")
                    link.click()
                    time.sleep(3.0) # Aumentar el tiempo de espera
                    # Esperar a que la pÃ¡gina de IE estÃ© lista despuÃ©s de la navegaciÃ³n
                    for _ in range(10): # Esperar hasta 5 segundos
                        if ie.ReadyState == 4: # READYSTATE_COMPLETE
                            print(f"    [DEBUG] IE ReadyState: COMPLETE despuÃ©s de volver al buscador.")
                            break
                        time.sleep(0.5)
                    else:
                        print(f"    [DEBUG] IE ReadyState no COMPLETE despuÃ©s de volver al buscador en {rit_completo}.")
                    break
            else:
                print(f"    [DEBUG] No se encontrÃ³ el enlace 'Volver al buscador' para {rit_completo}.")
            
            # NUEVO: Pausa de confirmaciÃ³n para seguir con la siguiente causa
            if index < len(causas) - 1:
                if not messagebox.askokcancel("Causa Finalizada", f"Se terminÃ³ de procesar: {rit_completo}\n\nPresiona OK para continuar con la siguiente causa o Cancelar para detener el proceso."):
                    cerrar_y_eliminar_word_temporal(cerrar_documento=False)
                    break
            else:
                cerrar_y_eliminar_word_temporal(cerrar_documento=False)

        except Exception as e:
            print(f"Error procesando {rit_completo}: {e}")
            cerrar_y_eliminar_word_temporal(cerrar_documento=False)

    messagebox.showinfo("Proceso Terminado", "El procesamiento por lotes ha finalizado.")
    cerrar_y_eliminar_word_temporal(cerrar_documento=False)

if __name__ == "__main__":
    automatizar_ie_con_rit()



