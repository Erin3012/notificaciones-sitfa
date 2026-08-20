"""Interfaz web local para ejecutar la automatización SITFA.

Se ejecuta únicamente en el equipo local y conserva la lógica de main.py.
Abrir con: python web_ui.py
"""

from __future__ import annotations

import io
import json
import os
import shutil
import subprocess
import threading
import time
import traceback
import webbrowser
import math
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path


HOST = "127.0.0.1"
PORT = 8765
PROJECT_DIR = Path(__file__).resolve().parent


class LogBuffer(io.TextIOBase):
    def __init__(self, original):
        self.original = original
        self._lock = threading.Lock()
        self._lines: list[str] = []

    def write(self, text):
        if text:
            with self._lock:
                self._lines.extend(text.splitlines(True))
                self._lines = self._lines[-2000:]
            self.original.write(text)
            self.original.flush()
        return len(text)

    def flush(self):
        self.original.flush()

    def snapshot(self):
        with self._lock:
            return "".join(self._lines)

    def clear(self):
        with self._lock:
            self._lines.clear()


class RunState:
    def __init__(self):
        self.lock = threading.Lock()
        self.running = False
        self.started_at = None
        self.error = None
        self.pending = None

    def as_dict(self):
        with self.lock:
            return {
                "running": self.running,
                "started_at": self.started_at,
                "error": self.error,
                "pending": self.pending and self.pending["kind"],
            }


LOG = LogBuffer(__import__("sys").stdout)
STATE = RunState()
STOP_EVENT = threading.Event()


class PendingInteraction:
    def __init__(self, kind, payload):
        self.kind = kind
        self.payload = payload
        self.event = threading.Event()
        self.response = None


def wait_for_interaction(kind, payload):
    interaction = PendingInteraction(kind, payload)
    with STATE.lock:
        STATE.pending = {"kind": kind, "payload": payload, "interaction": interaction}
    while not interaction.event.wait(0.2):
        if STOP_EVENT.is_set():
            return None
    with STATE.lock:
        if STATE.pending and STATE.pending.get("interaction") is interaction:
            STATE.pending = None
    return interaction.response


def request_confirmation(title, message):
    response = wait_for_interaction("confirm", {"title": title, "message": message})
    return bool(response) and not STOP_EVENT.is_set()


def request_liquidacion(**kwargs):
    payload = {
        "title": kwargs.get("title") or "Liquidación",
        "title_text": kwargs.get("title_text", ""),
        "initial_total_liq": kwargs.get("initial_total_liq", ""),
        "initial_monto_pen": kwargs.get("initial_monto_pen", ""),
        "initial_cuenta_lav": kwargs.get("initial_cuenta_lav", ""),
    }
    response = wait_for_interaction("liquidacion", payload)
    if not isinstance(response, dict):
        return response

    def amount(value):
        text = str(value or "").strip().replace("$", "").replace(".", "").replace(" ", "").replace(",", "")
        try:
            return float(text) if text else 0
        except ValueError:
            return 0

    total = amount(response.get("total_liq"))
    pension = amount(response.get("monto_pen"))
    response["cuotas"] = math.ceil(total / pension) if pension > 0 else 0
    return response


def notify_web(level, title, message):
    print(f"[{level.upper()}] {title}: {message}")


class WebMessageBox:
    def askyesno(self, title, message):
        return request_confirmation(title, message)

    def askokcancel(self, title, message):
        return request_confirmation(title, message)

    def showinfo(self, title, message):
        notify_web("info", title, message)

    def showerror(self, title, message):
        notify_web("error", title, message)


def choose_path(kind: str) -> str:
    import tkinter as tk
    from tkinter import filedialog

    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    try:
        if kind == "excel":
            return filedialog.askopenfilename(
                title="Selecciona el archivo Excel",
                filetypes=[("Excel", "*.xlsx *.xls"), ("Todos los archivos", "*.*")],
            )
        return filedialog.askdirectory(title="Selecciona la carpeta de descarga de PDF")
    finally:
        root.destroy()


def read_rits_from_excel(archivo_excel: str):
    import main
    import pythoncom

    pythoncom.CoInitialize()
    try:
        return [item["rit"] for item in main.leer_excel_causas(archivo_excel)]
    finally:
        pythoncom.CoUninitialize()


def run_automation(archivo_excel: str, carpeta_pdfs: str, causas, jueces):
    import main
    import pythoncom

    main.UI_CONFIRM_CALLBACK = request_confirmation
    main.UI_LIQUIDACION_CALLBACK = request_liquidacion
    main.UI_STOP_EVENT = STOP_EVENT
    main.messagebox = WebMessageBox()
    pythoncom.CoInitialize()

    with STATE.lock:
        STATE.running = True
        STATE.started_at = time.strftime("%Y-%m-%d %H:%M:%S")
        STATE.error = None

    print("\n=== SITFA: inicio de automatización ===")
    print(f"Excel: {archivo_excel}")
    print(f"Carpeta PDF: {carpeta_pdfs}")
    try:
        main.automatizar_con_backend(
            archivo_excel=archivo_excel,
            carpeta_pdfs=carpeta_pdfs,
            causas=causas,
            jueces=jueces,
        )
        print("=== SITFA: proceso finalizado ===")
    except Exception as exc:
        with STATE.lock:
            STATE.error = str(exc)
        print(f"[ERROR] {exc}")
        traceback.print_exc()
    finally:
        pythoncom.CoUninitialize()
        main.UI_CONFIRM_CALLBACK = None
        main.UI_LIQUIDACION_CALLBACK = None
        main.UI_STOP_EVENT = None
        with STATE.lock:
            STATE.pending = None
        with STATE.lock:
            STATE.running = False


def start_run(payload):
    archivo_excel = str(payload.get("archivo_excel", "")).strip()
    carpeta_pdfs = str(payload.get("carpeta_pdfs", "")).strip()
    if not archivo_excel or not carpeta_pdfs:
        raise ValueError("Debes indicar el archivo Excel y la carpeta de PDF.")
    if not Path(archivo_excel).is_file():
        raise ValueError("El archivo Excel no existe.")
    if not Path(carpeta_pdfs).is_dir():
        raise ValueError("La carpeta de PDF no existe.")
    causas = payload.get("causas") or []
    jueces = payload.get("jueces") or []
    if not causas:
        raise ValueError("Debes cargar o agregar al menos un RIT.")
    if len(jueces) != 10 or any(not str(juez).strip() for juez in jueces):
        raise ValueError("Debes ingresar los 10 jueces, desde el índice 0 al 9.")
    clean_causas = []
    for causa in causas:
        rit = str(causa.get("rit", "") if isinstance(causa, dict) else causa).strip().upper()
        partes = rit.split("-")
        if len(partes) != 3 or not partes[0] or not partes[1].isdigit() or not partes[2].isdigit():
            raise ValueError(f"RIT inválido: {rit}")
        clean_causas.append({"rit": rit})

    with STATE.lock:
        if STATE.running:
            raise ValueError("Ya hay una automatización en ejecución.")
    STOP_EVENT.clear()
    threading.Thread(
        target=run_automation,
        args=(archivo_excel, carpeta_pdfs, clean_causas, [str(juez).strip() for juez in jueces]),
        daemon=True,
        name="sitfa-automation",
    ).start()


HTML = r'''<!doctype html>
<html lang="es">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>SITFA · Automatización</title>
<style>
:root { --navy:#102a43; --navy-2:#163b5c; --teal:#008b8b; --teal-dark:#006d6d; --bg:#f4f7fa; --card:#fff; --text:#203040; --muted:#66788a; --line:#dce5ec; --ok:#16855b; --warn:#a66a00; }
* { box-sizing:border-box; }
body { margin:0; min-height:100vh; background:var(--bg); color:var(--text); font:15px/1.45 "Segoe UI",Arial,sans-serif; }
header { background:var(--navy); color:white; padding:24px clamp(20px,5vw,72px); }
header small { color:#b9d2e5; letter-spacing:.08em; text-transform:uppercase; }
header h1 { margin:5px 0 0; font-size:clamp(24px,3vw,34px); font-weight:650; }
main { max-width:1180px; margin:28px auto; padding:0 20px 40px; display:grid; grid-template-columns:minmax(320px, .9fr) minmax(420px, 1.4fr); gap:20px; }
.card { background:var(--card); border:1px solid var(--line); border-radius:14px; box-shadow:0 5px 20px #19324d0d; padding:22px; }
.card h2 { margin:0 0 5px; font-size:20px; color:var(--navy); }
.hint { color:var(--muted); margin:0 0 20px; }
label { display:block; margin:16px 0 7px; font-weight:600; }
.path-row { display:flex; gap:8px; }
input[type=text] { width:100%; min-width:0; border:1px solid #b9c9d6; border-radius:8px; padding:11px 12px; color:var(--text); background:#fbfdff; }
input:focus { outline:3px solid #008b8b33; border-color:var(--teal); }
button { border:0; border-radius:8px; padding:11px 15px; font:inherit; font-weight:650; cursor:pointer; transition:.15s ease; }
button:focus-visible { outline:3px solid #f3c969; outline-offset:2px; }
.secondary { background:#e9f0f5; color:var(--navy); white-space:nowrap; }
.secondary:hover { background:#dce8ef; }
.primary { background:var(--teal); color:white; width:100%; margin-top:24px; font-size:16px; }
.primary:hover { background:var(--teal-dark); }
button:disabled { opacity:.55; cursor:not-allowed; }
.status { display:flex; align-items:center; gap:10px; margin:0 0 18px; padding:12px 14px; border-radius:9px; background:#edf3f7; font-weight:650; }
.dot { width:10px; height:10px; border-radius:50%; background:#8495a5; }
.dot.running { background:#d18a00; box-shadow:0 0 0 5px #d18a0020; }
.dot.done { background:var(--ok); }
.log-head { display:flex; justify-content:space-between; align-items:center; gap:12px; margin-bottom:10px; }
.log-head h2 { margin:0; }
#logs { height:500px; overflow:auto; margin:0; padding:15px; background:#0d1b2a; border-radius:9px; color:#d8e8f3; font:12px/1.55 Consolas,"Courier New",monospace; white-space:pre-wrap; word-break:break-word; }
.notice { margin-top:14px; color:var(--muted); font-size:13px; }
.danger { background:#b53b3b; color:white; margin-top:10px; width:100%; }
.danger:hover { background:#922d2d; }
.modal-backdrop { position:fixed; inset:0; display:none; align-items:center; justify-content:center; padding:20px; background:#102a4380; z-index:10; }
.modal-backdrop.open { display:flex; }
.modal { width:min(520px,100%); background:white; border-radius:14px; padding:24px; box-shadow:0 20px 60px #102a4366; }
.modal h2 { margin:0 0 8px; color:var(--navy); }
.modal p { white-space:pre-wrap; color:var(--text); }
.modal-actions { display:flex; justify-content:flex-end; gap:10px; margin-top:20px; }
.modal-actions button { width:auto; }
.form-grid { display:grid; grid-template-columns:1fr 1fr; gap:12px; }
.form-grid label { margin-top:8px; }
.section-label { display:flex; align-items:center; justify-content:space-between; gap:10px; margin-top:22px; }
.section-label h3 { margin:0; color:var(--navy); font-size:16px; }
.rit-list { display:flex; flex-direction:column; gap:8px; max-height:260px; overflow:auto; padding:2px; }
.rit-row { display:flex; gap:8px; align-items:center; }
.rit-row input { flex:1; }
.rit-assignment { min-width:150px; color:var(--muted); font-size:12px; }
.icon-button { padding:9px 12px; }
.judge-grid { display:grid; grid-template-columns:1fr 1fr; gap:8px 12px; }
.judge-row { display:flex; align-items:center; gap:8px; }
.judge-index { width:25px; color:var(--muted); font-weight:700; text-align:center; }
@media (max-width:520px) { .form-grid { grid-template-columns:1fr; } }
@media (max-width:800px) { main { grid-template-columns:1fr; } #logs { height:360px; } }
</style>
</head>
<body>
<header><small>Notificaciones SITFA</small><h1>Centro de automatización</h1></header>
<main>
  <section class="card">
    <h2>Preparar ejecución</h2>
    <p class="hint">Selecciona los datos de entrada. La automatización seguirá usando SITFA, Word y los archivos de salida configurados.</p>
    <label for="excel">Archivo Excel de RIT</label>
    <div class="path-row"><input id="excel" type="text" placeholder="C:\\...\\causas.xlsx" autocomplete="off"><button class="secondary" onclick="choose('excel')">Elegir</button></div>
    <div class="section-label"><h3>RIT a procesar <span id="rit-count" class="notice">(0)</span></h3><button class="secondary icon-button" onclick="addRit()">+ Agregar RIT</button></div>
    <div id="rit-list" class="rit-list"><p class="notice">Selecciona un Excel o agrega un RIT manualmente.</p></div>
    <div class="section-label"><h3>Jueces asociados</h3><span class="notice">Índice 0 al 9</span></div>
    <div id="judge-list" class="judge-grid"></div>
    <label for="pdfs">Carpeta para descargar PDF</label>
    <div class="path-row"><input id="pdfs" type="text" placeholder="C:\\...\\descargas" autocomplete="off"><button class="secondary" onclick="choose('folder')">Elegir</button></div>
    <button id="start" class="primary" onclick="startRun()">Iniciar automatización</button>
    <button id="stop" class="danger" onclick="stopRun()" disabled>Detener automatización</button>
    <p class="notice">La ventana permanecerá disponible mientras el proceso trabaja. Las confirmaciones y datos de liquidación se solicitan dentro de esta página.</p>
  </section>
  <section class="card">
    <div class="log-head"><h2>Registro de actividad</h2><button class="secondary" onclick="clearLogs()">Limpiar</button></div>
    <div id="status" class="status"><span id="dot" class="dot"></span><span id="status-text">Listo para iniciar</span></div>
    <pre id="logs" aria-live="polite">Aún no hay actividad registrada.</pre>
  </section>
</main>
<div id="modal-backdrop" class="modal-backdrop" role="dialog" aria-modal="true" aria-labelledby="modal-title">
  <div class="modal">
    <h2 id="modal-title"></h2>
    <p id="modal-message"></p>
    <div id="liquidacion-fields" class="form-grid" style="display:none">
      <div><label for="tipo-liq">Tipo de liquidación</label><select id="tipo-liq" style="width:100%;padding:11px;border:1px solid #b9c9d6;border-radius:8px"><option>Peso Rojo</option><option>Manual Excel</option></select></div>
      <div><label for="total-liq">Total liquidación</label><input id="total-liq" type="text"></div>
      <div><label for="monto-pen">Monto pensión mensual</label><input id="monto-pen" type="text"></div>
      <div><label for="cuenta-lav">N° Cuenta LAV</label><input id="cuenta-lav" type="text"></div>
    </div>
    <div class="modal-actions"><button id="modal-cancel" class="secondary" onclick="respond(false)">Cancelar</button><button id="modal-ok" class="primary" style="width:auto;margin-top:0" onclick="respond(true)">Continuar</button></div>
  </div>
</div>
<script>
const $ = id => document.getElementById(id);
let lastLogs = '';
let currentPending = null;
async function choose(kind) {
  const res = await fetch('/api/choose', {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({kind})});
  const data = await res.json();
  if (!data.ok) return showLocalMessage(data.error);
  $(kind === 'excel' ? 'excel' : 'pdfs').value = data.path;
  if (kind === 'excel') await loadExcel();
}
let rits = [];
function renderRits() {
  $('rit-count').textContent = '(' + rits.length + ')';
  const list = $('rit-list');
  if (!rits.length) { list.innerHTML = '<p class="notice">No hay RIT cargados.</p>'; return; }
  list.innerHTML = rits.map((rit, i) => { const idx = judgeIndex(rit); const judge = idx >= 0 ? (getJudges()[idx] || 'sin asignar') : 'RIT inválido'; return '<div class="rit-row"><input type="text" value="' + escapeHtml(rit) + '" aria-label="RIT ' + (i+1) + '" onchange="updateRit(' + i + ', this.value)"><span class="rit-assignment">Índice ' + (idx >= 0 ? idx : '-') + ': ' + escapeHtml(judge) + '</span><button class="secondary icon-button" onclick="removeRit(' + i + ')" aria-label="Eliminar RIT">Eliminar</button></div>'; }).join('');
}
function escapeHtml(value) { return String(value).replace(/[&<>'"]/g, char => ({'&':'&amp;','<':'&lt;','>':'&gt;',"'":'&#39;','"':'&quot;'}[char])); }
function updateRit(index, value) { rits[index] = value.trim().toUpperCase(); }
function judgeIndex(rit) { const parts = String(rit || '').split('-'); return parts.length === 3 && /\d/.test(parts[1]) ? Number(parts[1].slice(-1)) : -1; }
function addRit() { rits.push(''); renderRits(); const inputs = $('rit-list').querySelectorAll('input'); if (inputs.length) inputs[inputs.length-1].focus(); }
function removeRit(index) { rits.splice(index, 1); renderRits(); }
async function loadExcel() {
  const res = await fetch('/api/load_excel', {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({archivo_excel:$('excel').value})});
  const data = await res.json();
  if (!data.ok) return showLocalMessage(data.error);
  rits = data.rits || []; renderRits();
}
function renderJudges() {
  $('judge-list').innerHTML = Array.from({length:10}, (_, i) => '<div class="judge-row"><span class="judge-index">' + i + '</span><input id="judge-' + i + '" type="text" placeholder="Nombre del juez" aria-label="Juez índice ' + i + '"></div>').join('');
}
function getJudges() { return Array.from({length:10}, (_, i) => $('judge-' + i).value.trim()); }
async function startRun() {
  const res = await fetch('/api/start', {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({archivo_excel:$('excel').value, carpeta_pdfs:$('pdfs').value, causas:rits.map(rit => ({rit})), jueces:getJudges()})});
  const data = await res.json();
  if (!data.ok) showLocalMessage(data.error);
  refresh();
}
async function stopRun() { await fetch('/api/stop', {method:'POST'}); refresh(); }
async function clearLogs() {
  const res = await fetch('/api/logs/clear', {method:'POST', cache:'no-store'});
  const data = await res.json();
  if (!data.ok) return showLocalMessage(data.error || 'No se pudieron limpiar los logs.');
  lastLogs = '';
  $('logs').textContent = 'Registro limpiado.';
  setTimeout(refresh, 150);
}
function showLocalMessage(message) { $('modal-title').textContent='Aviso'; $('modal-message').textContent=message; $('liquidacion-fields').style.display='none'; $('modal-cancel').style.display='none'; $('modal-ok').textContent='Cerrar'; currentPending={local:true}; $('modal-backdrop').classList.add('open'); }
async function respond(ok) {
  if (currentPending && currentPending.local) { $('modal-backdrop').classList.remove('open'); currentPending=null; return; }
  let response = ok;
  if (ok && currentPending && currentPending.kind === 'liquidacion') response = {tipo:$('tipo-liq').value, total_liq:$('total-liq').value, monto_pen:$('monto-pen').value, cuenta_lav:$('cuenta-lav').value};
  await fetch('/api/respond', {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({response})});
  $('modal-backdrop').classList.remove('open'); currentPending=null; refresh();
}
async function refreshPending() {
  const res = await fetch('/api/pending'), data = await res.json();
  if (!data.pending || (currentPending && currentPending.id === data.pending.id)) return;
  currentPending = data.pending; $('modal-title').textContent=data.pending.payload.title || 'Confirmación'; $('modal-message').textContent=data.pending.payload.message || ('Procesando: ' + (data.pending.payload.title_text || ''));
  $('modal-cancel').style.display='inline-block'; $('modal-ok').textContent=data.pending.kind === 'confirm' ? 'Continuar' : 'Guardar y continuar';
  if (data.pending.kind === 'liquidacion') { const p=data.pending.payload; $('liquidacion-fields').style.display='grid'; $('total-liq').value=p.initial_total_liq || ''; $('monto-pen').value=p.initial_monto_pen || ''; $('cuenta-lav').value=p.initial_cuenta_lav || ''; } else $('liquidacion-fields').style.display='none';
  $('modal-backdrop').classList.add('open');
}
async function refresh() {
  const [statusRes, logsRes] = await Promise.all([fetch('/api/status'), fetch('/api/logs')]);
  const status = await statusRes.json(), logs = await logsRes.text();
  const running = status.running;
  $('start').disabled = running;
  $('stop').disabled = !running;
  $('status-text').textContent = running ? 'Automatización en ejecución…' : (status.error ? 'Terminó con error' : (status.started_at ? 'Proceso finalizado' : 'Listo para iniciar'));
  $('dot').className = 'dot ' + (running ? 'running' : (status.started_at && !status.error ? 'done' : ''));
  if (logs !== lastLogs) { $('logs').textContent = logs || 'Aún no hay actividad registrada.'; $('logs').scrollTop = $('logs').scrollHeight; lastLogs = logs; }
}
renderJudges(); renderRits(); setInterval(refresh, 1000); setInterval(refreshPending, 400); refresh();
</script>
</body>
</html>'''


class Handler(BaseHTTPRequestHandler):
    def log_message(self, *_args):
        return

    def send_json(self, data, status=HTTPStatus.OK):
        raw = json.dumps(data, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(raw)))
        self.end_headers()
        self.wfile.write(raw)

    def read_json(self):
        length = int(self.headers.get("Content-Length", "0"))
        return json.loads(self.rfile.read(length) or b"{}")

    def do_GET(self):
        if self.path == "/":
            raw = HTML.encode("utf-8")
            self.send_response(HTTPStatus.OK)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Cache-Control", "no-store")
            self.send_header("Content-Length", str(len(raw)))
            self.end_headers()
            self.wfile.write(raw)
        elif self.path == "/api/status":
            self.send_json(STATE.as_dict())
        elif self.path == "/api/logs":
            raw = LOG.snapshot().encode("utf-8", errors="replace")
            self.send_response(HTTPStatus.OK)
            self.send_header("Content-Type", "text/plain; charset=utf-8")
            self.send_header("Content-Length", str(len(raw)))
            self.end_headers()
            self.wfile.write(raw)
        elif self.path == "/api/pending":
            with STATE.lock:
                pending = STATE.pending
                result = None
                if pending:
                    result = {
                        "id": id(pending["interaction"]),
                        "kind": pending["kind"],
                        "payload": pending["payload"],
                    }
            self.send_json({"pending": result})
        else:
            self.send_error(HTTPStatus.NOT_FOUND)

    def do_DELETE(self):
        if self.path == "/api/logs":
            LOG.clear()
            self.send_json({"ok": True})
        else:
            self.send_error(HTTPStatus.NOT_FOUND)

    def do_POST(self):
        try:
            payload = self.read_json()
            if self.path == "/api/load_excel":
                archivo_excel = str(payload.get("archivo_excel", "")).strip()
                if not Path(archivo_excel).is_file():
                    raise ValueError("El archivo Excel no existe.")
                self.send_json({"ok": True, "rits": read_rits_from_excel(archivo_excel)})
            elif self.path == "/api/start":
                start_run(payload)
                self.send_json({"ok": True})
            elif self.path == "/api/stop":
                with STATE.lock:
                    running = STATE.running
                    pending = STATE.pending
                if running:
                    STOP_EVENT.set()
                    if pending:
                        pending["interaction"].response = False
                        pending["interaction"].event.set()
                    print("[SITFA] Solicitud de detención recibida desde la interfaz.")
                self.send_json({"ok": True, "running": running})
            elif self.path == "/api/respond":
                with STATE.lock:
                    pending = STATE.pending
                if not pending:
                    raise ValueError("No hay ninguna solicitud pendiente.")
                pending["interaction"].response = payload.get("response")
                pending["interaction"].event.set()
                self.send_json({"ok": True})
            elif self.path == "/api/logs/clear":
                LOG.clear()
                self.send_json({"ok": True})
            elif self.path == "/api/choose":
                kind = payload.get("kind")
                if kind not in {"excel", "folder"}:
                    raise ValueError("Tipo de selector no válido.")
                self.send_json({"ok": True, "path": choose_path(kind)})
            else:
                self.send_error(HTTPStatus.NOT_FOUND)
        except ValueError as exc:
            self.send_json({"ok": False, "error": str(exc)}, HTTPStatus.BAD_REQUEST)
        except Exception as exc:
            self.send_json({"ok": False, "error": str(exc)}, HTTPStatus.INTERNAL_SERVER_ERROR)


def main():
    server = ThreadingHTTPServer((HOST, PORT), Handler)
    url = f"http://{HOST}:{PORT}"
    print(f"SITFA UI disponible en {url}")
    print("Presiona Ctrl+C en esta ventana para cerrar la interfaz.")
    chrome_candidates = [
        shutil.which("chrome"),
        os.path.expandvars(r"%ProgramFiles%\Google\Chrome\Application\chrome.exe"),
        os.path.expandvars(r"%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe"),
        os.path.expandvars(r"%LocalAppData%\Google\Chrome\Application\chrome.exe"),
    ]
    chrome_path = next((path for path in chrome_candidates if path and Path(path).is_file()), None)
    if chrome_path:
        subprocess.Popen([chrome_path, url], close_fds=True)
    else:
        webbrowser.open(url)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nInterfaz cerrada.")
    finally:
        server.server_close()


if __name__ == "__main__":
    import sys
    sys.stdout = LOG
    main()
