# -*- coding: utf-8 -*-
"""
TAREO MÓVIL – GRUPO DE COSECHA | PRIZE PRO
Flask + SQLite/PostgreSQL + PWA. Listo para GitHub + Render.

Mejoras incluidas:
- Login con formato móvil verde/blanco como referencia.
- Pantalla principal: Soporte, Configuraciones, Sincronización y Hojas de Tareo.
- Creación de hoja: fecha, grupo, subgrupo, labor y responsable.
- Detalle de hoja con tabs: Labores, Trabajadores, Rend./Avance por Labor.
- Registro por labor-consumidor, detalle de trabajador por labor y lecturas por balde.
- Adaptado a desktop y celular con diseño responsive tipo app.
"""
import os, re, sqlite3, base64
from datetime import datetime, date
from functools import wraps
from io import BytesIO

from flask import Flask, request, redirect, url_for, session, flash, jsonify, send_file, render_template_string, Response
from openpyxl import Workbook, load_workbook
from werkzeug.security import generate_password_hash, check_password_hash

try:
    import psycopg2
    import psycopg2.extras
except Exception:
    psycopg2 = None

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PERSIST_DIR = os.getenv("PERSIST_DIR", "/data" if os.path.isdir("/data") else BASE_DIR)
FOTO_DIR = os.path.join(PERSIST_DIR, "fotos_marcacion")
os.makedirs(FOTO_DIR, exist_ok=True)
DB_PATH = os.path.join(PERSIST_DIR, "asistencia_tareo.db")
DATABASE_URL = os.getenv("DATABASE_URL", "").strip()
USE_POSTGRES = bool(DATABASE_URL)

app = Flask(__name__)
app.secret_key = os.getenv("SECRET_KEY", "cambiar-clave-en-render")
app.config["MAX_CONTENT_LENGTH"] = 30 * 1024 * 1024

# ========================= DB =========================
def is_pg(): return USE_POSTGRES and psycopg2 is not None

def get_conn():
    if is_pg():
        return psycopg2.connect(DATABASE_URL, cursor_factory=psycopg2.extras.RealDictCursor)
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def qmark(sql): return sql.replace("?", "%s") if is_pg() else sql

def row_to_dict(r): return dict(r) if r else None

def rows_to_dict(rows): return [row_to_dict(r) for r in (rows or [])]

def execute(sql, params=(), fetchone=False, fetchall=False, commit=False):
    conn = get_conn(); cur = conn.cursor()
    cur.execute(qmark(sql), params)
    data = None
    if fetchone: data = cur.fetchone()
    if fetchall: data = cur.fetchall()
    if commit: conn.commit()
    cur.close(); conn.close()
    return data

def scalar(sql, params=()):
    r = row_to_dict(execute(sql, params, fetchone=True))
    return list(r.values())[0] if r else 0

def now_str(): return datetime.now().strftime("%Y-%m-%d %H:%M:%S")
def today_str(): return date.today().strftime("%Y-%m-%d")

def init_db():
    conn = get_conn(); cur = conn.cursor()
    idtype = "SERIAL PRIMARY KEY" if is_pg() else "INTEGER PRIMARY KEY AUTOINCREMENT"
    cur.execute(qmark(f"""CREATE TABLE IF NOT EXISTS usuarios(
        id {idtype}, usuario TEXT UNIQUE NOT NULL, password_hash TEXT NOT NULL,
        nombres TEXT, rol TEXT DEFAULT 'operador', estado TEXT DEFAULT 'ACTIVO', creado_en TEXT)"""))
    cur.execute(qmark(f"""CREATE TABLE IF NOT EXISTS trabajadores(
        id {idtype}, dni TEXT UNIQUE NOT NULL, trabajador TEXT, empresa TEXT,
        area TEXT, cargo TEXT, actividad TEXT, planilla TEXT, estado TEXT DEFAULT 'ACTIVO', fecha_carga TEXT)"""))
    cur.execute(qmark(f"""CREATE TABLE IF NOT EXISTS asistencia(
        id {idtype}, dni TEXT NOT NULL, trabajador TEXT, empresa TEXT, area TEXT, cargo TEXT,
        tipo TEXT NOT NULL, fecha TEXT NOT NULL, hora TEXT NOT NULL, fecha_hora TEXT NOT NULL,
        metodo TEXT, foto_path TEXT, latitud TEXT, longitud TEXT, registrado_por TEXT, observacion TEXT)"""))
    cur.execute(qmark(f"""CREATE TABLE IF NOT EXISTS tareos(
        id {idtype}, hoja_id INTEGER, dni TEXT NOT NULL, trabajador TEXT, empresa TEXT, area TEXT, cargo TEXT,
        fecha TEXT NOT NULL, labor TEXT, lote TEXT, fundo TEXT, horas REAL DEFAULT 0,
        cantidad REAL DEFAULT 0, unidad TEXT, observacion TEXT, registrado_por TEXT, creado_en TEXT)"""))
    cur.execute(qmark(f"""CREATE TABLE IF NOT EXISTS hojas_tareo(
        id {idtype}, fecha TEXT NOT NULL, grupo TEXT, subgrupo TEXT, labor TEXT, responsable TEXT,
        estado TEXT DEFAULT 'ABIERTA', registros INTEGER DEFAULT 0, horas_total REAL DEFAULT 0, rendimiento_total REAL DEFAULT 0,
        creado_por TEXT, creado_en TEXT)"""))
    cur.execute(qmark(f"""CREATE TABLE IF NOT EXISTS lecturas_balde(
        id {idtype}, hoja_id INTEGER, dni TEXT, trabajador TEXT, fecha_hora TEXT, a_diurno REAL DEFAULT 0, a_noct REAL DEFAULT 0, registrado_por TEXT)"""))

    cur.execute(qmark("SELECT id FROM usuarios WHERE usuario=?"), ("admin",))
    if not cur.fetchone():
        cur.execute(qmark("INSERT INTO usuarios(usuario,password_hash,nombres,rol,estado,creado_en) VALUES(?,?,?,?,?,?)"),
                    ("admin", generate_password_hash("admin123"), "ADMINISTRADOR", "admin", "ACTIVO", now_str()))
    conn.commit(); cur.close(); conn.close()

# ========================= UTIL =========================
def normalizar_columna(c):
    c = str(c or "").strip().upper()
    for a,b in {"Á":"A","É":"E","Í":"I","Ó":"O","Ú":"U","Ñ":"N"}.items(): c = c.replace(a,b)
    return re.sub(r"\s+", " ", c)

def limpiar_dni(v):
    solo = re.sub(r"\D", "", str(v or ""))
    return solo[-8:] if len(solo) >= 8 else solo

def limpiar_texto(v, upper=True):
    s = "" if v is None else str(v).strip()
    return s.upper() if upper else s

def login_required(f):
    @wraps(f)
    def w(*args, **kwargs):
        if not session.get("usuario"): return redirect(url_for("login"))
        return f(*args, **kwargs)
    return w

def admin_required(f):
    @wraps(f)
    def w(*args, **kwargs):
        if not session.get("usuario"): return redirect(url_for("login"))
        if session.get("rol") != "admin":
            flash("Solo administrador puede ingresar a esta opción.", "danger")
            return redirect(url_for("home"))
        return f(*args, **kwargs)
    return w

def excel_response(headers, rows, filename, sheet="DATOS"):
    wb = Workbook(); ws = wb.active; ws.title = sheet[:31]
    ws.append(headers)
    for r in rows:
        ws.append([r.get(h.lower(), r.get(h, "")) for h in headers])
    for col in ws.columns:
        max_len = max(len(str(c.value or "")) for c in col)
        ws.column_dimensions[col[0].column_letter].width = min(max(max_len + 2, 12), 45)
    out = BytesIO(); wb.save(out); out.seek(0)
    return send_file(out, as_attachment=True, download_name=filename, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ========================= UI =========================
BASE_HTML = r"""
<!doctype html><html lang="es"><head>
<meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<meta name="theme-color" content="#2f773b"><link rel="manifest" href="{{ url_for('manifest') }}">
<title>{{ title or 'Tareo Móvil' }}</title>
<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet">
<link href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.css" rel="stylesheet">
<script src="https://unpkg.com/html5-qrcode" type="text/javascript"></script>
<style>
:root{--verde:#2f773b;--verde2:#3f8748;--verde3:#276a33;--verdeClaro:#eaf5eb;--line:#e8ece8;--txt:#2d5b35;--gris:#6b7d6d;--amarillo:#ffc20e}
*{box-sizing:border-box}body{margin:0;background:#fff;font-family:Inter,Segoe UI,Arial,sans-serif;color:#21472a}.app-bg{min-height:100vh;background:linear-gradient(180deg,#fff 0%,#fff 62%,#fbfbfb 100%)}.shell{width:min(1120px,100%);margin:0 auto;padding:18px}.phone-wrap{max-width:430px;margin:0 auto}.desktop-grid{display:grid;grid-template-columns:390px 1fr;gap:18px;align-items:start}.header-title{text-align:center;color:#166534;font-family:Georgia,serif;font-weight:900;letter-spacing:.5px;font-size:23px;line-height:1.13;margin:4px 0 22px;text-transform:uppercase}.green-hero{background:var(--verde);border-radius:0 0 18px 18px;min-height:145px;padding:12px 16px 22px;color:white;text-align:center;position:relative}.green-top{display:flex;justify-content:space-between;align-items:center;font-size:11px;font-weight:800}.avatar{width:78px;height:78px;border-radius:999px;background:white;color:var(--verde);display:grid;place-items:center;margin:10px auto 2px;font-size:43px;box-shadow:0 8px 20px rgba(0,0,0,.13)}.login-name{font-size:11px;font-weight:800}.white-input{height:36px;background:white;border-radius:10px;box-shadow:0 5px 13px rgba(0,0,0,.18);border:0}.floating-card{background:white;border-radius:10px;box-shadow:0 8px 18px rgba(0,0,0,.15);padding:12px}.tile{width:74px;height:70px;border-radius:8px;background:white;box-shadow:0 7px 17px rgba(0,0,0,.14);display:flex;flex-direction:column;align-items:center;justify-content:center;color:var(--verde);font-weight:900;font-size:10px;text-align:center}.tile i{font-size:26px;margin-bottom:5px}.bottom-sync{position:fixed;left:10px;bottom:8px;color:#317a3e;font-size:10px;font-weight:700}.bottom-out{position:fixed;right:14px;bottom:8px;color:#c84c4c;font-size:20px}.tab-main{display:flex;gap:0;background:#fafafa;padding-left:0;border-top:4px solid #d7d7d7}.tab-main a{flex:1;text-align:center;text-decoration:none;color:#508557;font-weight:900;font-size:13px;padding:14px 8px;border-radius:7px 7px 0 0;background:#fff}.tab-main a.active{background:var(--verde);color:white;box-shadow:0 3px 7px rgba(0,0,0,.18)}.subtabs{display:flex;background:#fff}.subtabs a{flex:1;text-align:center;padding:13px 5px;text-decoration:none;color:#4b8a54;font-weight:900;font-size:12px}.subtabs a.active{background:var(--verde);color:white}.panel-green{background:var(--verde);color:white;text-align:center;padding:21px 12px 42px}.panel-green i{font-size:38px}.panel-green h4{font-size:11px;font-weight:900;margin:5px 0 0}.toolstrip{background:white;margin:-25px 9px 5px;border-radius:9px;min-height:49px;box-shadow:0 5px 13px rgba(0,0,0,.22);display:flex;align-items:center;gap:20px;padding:7px 14px;color:var(--verde);font-size:24px}.toolstrip button,.toolstrip a{border:0;background:transparent;color:var(--verde);font-size:24px;text-decoration:none}.info-bar{margin:0 9px;background:var(--verde);color:white;border-radius:2px;display:grid;grid-template-columns:1fr 1fr 1fr 1fr 22px;align-items:center;font-size:10px;font-weight:900;height:23px}.info-bar div{text-align:center;border-right:1px solid rgba(255,255,255,.28)}.worker-card{background:white;margin:10px 12px;border-radius:10px;box-shadow:0 3px 12px rgba(0,0,0,.20);padding:11px 13px;color:#397443;position:relative}.worker-title{display:grid;grid-template-columns:1fr 1fr;gap:8px;font-size:9px;font-weight:900;text-transform:uppercase}.worker-title b{font-size:10px;color:var(--verde)}.worker-grid{display:grid;grid-template-columns:1fr 1fr 1fr;gap:7px;margin-top:7px}.worker-grid label{font-size:8px;font-weight:900;color:#6c8a6f;margin-bottom:1px}.mini-input{height:25px;border:1px solid #9ebaa0;border-radius:3px;font-size:10px;padding:3px 5px;width:100%;font-weight:800;color:#315f39}.mini-badge{border-radius:3px;color:white;font-size:8px;font-weight:900;text-align:center;padding:4px 3px;text-transform:uppercase}.bg-y{background:var(--amarillo)!important;color:white}.bg-g{background:#42b852!important}.person-dot{width:42px;height:42px;border-radius:999px;background:#378145;color:white;display:grid;place-items:center;font-size:26px;float:left;margin-right:9px}.small-label{font-size:8px;color:#79937b;font-weight:900}.small-value{font-size:9px;color:#466a49;font-weight:900}.leaf{width:120px;height:120px;border-radius:70% 30% 70% 30%;background:linear-gradient(135deg,#ffd8bd,#eef4c7,#cbd9b6);opacity:.72;margin:20px auto 0;transform:rotate(-20deg)}.card-pro{background:white;border:1px solid var(--line);border-radius:18px;box-shadow:0 8px 20px rgba(0,0,0,.10)}.btn-green{background:var(--verde);border-color:var(--verde);color:white;font-weight:900;border-radius:9px}.btn-green:hover{background:var(--verde3);color:white}.form-control,.form-select{border-radius:9px;border:1px solid #dfe7df;font-weight:700;font-size:13px}.form-label{font-size:12px;font-weight:900;color:#3e7545}.page-card{border-radius:13px;overflow:hidden;border:1px solid #e5e7e5;background:white}.list-table th{font-size:11px;color:#497550}.list-table td{font-size:12px;vertical-align:middle}.status-pill{display:inline-block;background:#39b54a;color:white;border-radius:4px;padding:4px 8px;font-size:9px;font-weight:900}.top-actions{display:flex;gap:10px;flex-wrap:wrap;justify-content:center;margin-top:-16px}.top-actions .tile{width:82px;height:76px}.login-page .shell{padding:0}.login-form{margin:-7px auto 0;width:92%;max-width:360px}.login-form .floating-card{padding:13px 14px 18px}.alert{border-radius:12px;font-size:13px}.desk-panel{display:block}.mobile-only{display:none}
@media(max-width:860px){.shell{padding:0}.desktop-grid{display:block}.desk-panel{display:none}.mobile-only{display:block}.header-title{font-size:17px;margin:16px 7px 20px}.page-card{border-radius:0;border-left:0;border-right:0}.phone-wrap{max-width:100%}.green-hero{border-radius:0}.worker-card{margin-left:9px;margin-right:9px}.toolstrip{gap:15px}.info-bar{font-size:8.5px}.bottom-sync,.bottom-out{position:fixed}.desktop-pad{padding:0 0 28px}.tab-main a,.subtabs a{font-size:11px}.worker-grid{gap:5px}.floating-card{border-radius:9px}.top-actions .tile{width:72px;height:70px}}
</style></head><body class="{{ 'login-page' if not session.get('usuario') else '' }}"><div class="app-bg"><main class="shell">
{% with messages=get_flashed_messages(with_categories=true) %}{% if messages %}<div class="phone-wrap mt-2">{% for cat,msg in messages %}<div class="alert alert-{{cat}} shadow-sm">{{msg}}</div>{% endfor %}</div>{% endif %}{% endwith %}
{{ body|safe }}</main></div>
<audio id="sndOk"><source src="data:audio/wav;base64,UklGRjQAAABXQVZFZm10IBAAAAABAAEAESsAACJWAAACABAAZGF0YRAAAAAAAP//AAD//wAA//8AAP//AAD//wAA//8=" type="audio/wav"></audio>
<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/js/bootstrap.bundle.min.js"></script><script>function beep(){try{let a=document.getElementById('sndOk');a.currentTime=0;a.play().catch(()=>{});}catch(e){}}if('serviceWorker' in navigator){navigator.serviceWorker.register('/sw.js').catch(()=>{});}</script>
</body></html>
"""

def render_page(body, title="Tareo Móvil", **ctx):
    return render_template_string(BASE_HTML, body=render_template_string(body, **ctx), title=title)

# ========================= AUTH + HOME =========================
@app.route('/login', methods=['GET','POST'])
def login():
    if request.method == 'POST':
        usuario = request.form.get('usuario','').strip()
        password = request.form.get('password','')
        row = row_to_dict(execute('SELECT * FROM usuarios WHERE usuario=?', (usuario,), fetchone=True))
        if row and row.get('estado','ACTIVO') == 'ACTIVO' and check_password_hash(row['password_hash'], password):
            session['usuario']=usuario; session['rol']=row.get('rol','operador'); session['nombres']=row.get('nombres') or usuario
            return redirect(url_for('home'))
        flash('Usuario/clave incorrecta o usuario inactivo.', 'danger')
    body = """
    <div class="phone-wrap">
      <div class="green-hero" style="min-height:225px">
        <div class="green-top"><span><i class="bi bi-headset"></i> Soporte</span><span><i class="bi bi-gear"></i> Config.</span></div>
        <div class="avatar"><i class="bi bi-person-circle"></i></div><div class="login-name">INICIAR SESIÓN</div>
      </div>
      <form method="post" class="login-form">
        <div class="floating-card">
          <input class="form-control white-input mb-2" name="usuario" required autofocus placeholder="Usuario">
          <input class="form-control white-input mb-3" name="password" type="password" required placeholder="Clave">
          <button class="btn btn-green w-100"><i class="bi bi-box-arrow-in-right me-1"></i> INGRESAR</button>
          <div class="text-center small mt-2 text-muted">Demo: admin / admin123</div>
        </div>
      </form>
      <div class="d-flex justify-content-center gap-3 mt-4">
        <div class="tile"><i class="bi bi-list-check"></i>TAREO</div><div class="tile"><i class="bi bi-file-earmark-bar-graph"></i>REPORTES<br>TAREO</div>
      </div>
      <div class="leaf"></div>
      <div class="bottom-sync"><i class="bi bi-arrow-repeat"></i> Sincronizar Tablas Maestras<br>Actualizado hasta: {{ now }}</div><a class="bottom-out"><i class="bi bi-box-arrow-right"></i></a>
    </div>"""
    return render_page(body, title='Login Tareo Móvil', now=now_str())

@app.route('/logout')
def logout():
    session.clear(); return redirect(url_for('login'))

@app.route('/')
@login_required
def home():
    hojas = rows_to_dict(execute('SELECT * FROM hojas_tareo ORDER BY fecha DESC, id DESC LIMIT 8', fetchall=True))
    body = """
    <div class="desktop-grid">
      <div class="phone-wrap">
        <div class="green-hero" style="min-height:220px">
          <div class="green-top"><a class="text-white text-decoration-none" href="{{url_for('soporte')}}"><i class="bi bi-headset"></i> Soporte</a><a class="text-white text-decoration-none" href="{{url_for('configuraciones')}}"><i class="bi bi-gear"></i> Config.</a></div>
          <div class="avatar"><i class="bi bi-person-circle"></i></div><div class="login-name">{{ session.get('nombres','USUARIO') }}</div>
          <div class="white-input mt-3"></div>
        </div>
        <div class="top-actions">
          <a class="tile text-decoration-none" href="{{url_for('hojas_tareo')}}"><i class="bi bi-list-check"></i>TAREO</a>
          <a class="tile text-decoration-none" href="{{url_for('reportes')}}"><i class="bi bi-file-earmark-bar-graph"></i>REPORTES<br>TAREO</a>
          <a class="tile text-decoration-none" href="{{url_for('sincronizacion')}}"><i class="bi bi-arrow-repeat"></i>SINC.</a>
        </div>
        <div class="leaf"></div>
        <div class="bottom-sync"><i class="bi bi-arrow-repeat"></i> Sincronizar Tablas Maestras<br>Actualizado hasta: {{ now }}</div><a href="{{url_for('logout')}}" class="bottom-out"><i class="bi bi-box-arrow-right"></i></a>
      </div>
      <div class="desk-panel">
        <h1 class="header-title">TAREO MÓVIL – GRUPO DE COSECHA</h1>
        <div class="card-pro p-4 mb-3"><div class="d-flex justify-content-between align-items-center"><div><h4 class="fw-bold text-success mb-1">Hojas recientes</h4><div class="text-muted small">Crea una hoja y registra labores, trabajadores y avances.</div></div><a class="btn btn-green" href="{{url_for('crear_hoja')}}"><i class="bi bi-plus-lg"></i> Crear hoja</a></div></div>
        <div class="card-pro p-3"><div class="table-responsive"><table class="table list-table"><thead><tr><th>Fecha</th><th>Grupo</th><th>Subgrupo</th><th>Labor</th><th>Responsable</th><th>Estado</th><th></th></tr></thead><tbody>{% for h in hojas %}<tr><td>{{h.fecha}}</td><td>{{h.grupo}}</td><td>{{h.subgrupo}}</td><td>{{h.labor}}</td><td>{{h.responsable}}</td><td><span class="status-pill">{{h.estado}}</span></td><td><a class="btn btn-sm btn-green" href="{{url_for('detalle_hoja', hoja_id=h.id)}}">Abrir</a></td></tr>{% else %}<tr><td colspan="7" class="text-center text-muted py-4">Sin hojas creadas.</td></tr>{% endfor %}</tbody></table></div></div>
      </div>
    </div>"""
    return render_page(body, hojas=hojas, now=now_str())

# ========================= HOJAS TAREO =========================
@app.route('/hojas')
@login_required
def hojas_tareo():
    hojas = rows_to_dict(execute('SELECT * FROM hojas_tareo ORDER BY fecha DESC, id DESC LIMIT 50', fetchall=True))
    body = """
    <div class="phone-wrap desktop-pad">
      <h2 class="header-title">HOJAS DE TAREO PARA CREAR</h2>
      <div class="floating-card mb-2 d-flex gap-2"><a class="btn btn-green flex-fill" href="{{url_for('crear_hoja')}}"><i class="bi bi-plus-lg"></i> NUEVA HOJA</a><a class="btn btn-outline-success" href="{{url_for('home')}}"><i class="bi bi-house"></i></a></div>
      {% for h in hojas %}<a class="text-decoration-none" href="{{url_for('detalle_hoja', hoja_id=h.id)}}"><div class="worker-card"><div class="worker-title"><div><b>{{h.labor}}</b><br>{{h.fecha}}</div><div class="text-end"><b>{{h.grupo}}</b><br>{{h.subgrupo}}</div></div><div class="mt-2 small"><b>RESPONSABLE:</b> {{h.responsable}} &nbsp; <span class="status-pill">{{h.estado}}</span></div><div class="info-bar mt-2 mx-0"><div><i class="bi bi-calendar"></i> {{h.fecha}}</div><div><i class="bi bi-list"></i> {{h.registros}} Reg.</div><div><i class="bi bi-clock"></i> {{'%.2f'|format(h.horas_total or 0)}} H.</div><div>A.Rend {{'%.2f'|format(h.rendimiento_total or 0)}}</div><span>⌄</span></div></div></a>{% else %}<div class="alert alert-light border text-center">Aún no tienes hojas. Crea la primera.</div>{% endfor %}
    </div>"""
    return render_page(body, hojas=hojas)

@app.route('/hojas/crear', methods=['GET','POST'])
@login_required
def crear_hoja():
    if request.method == 'POST':
        fecha = request.form.get('fecha') or today_str()
        grupo = limpiar_texto(request.form.get('grupo'))
        subgrupo = limpiar_texto(request.form.get('subgrupo'))
        labor = limpiar_texto(request.form.get('labor'))
        responsable = limpiar_texto(request.form.get('responsable'))
        if not grupo or not labor or not responsable:
            flash('Completa grupo, labor y responsable.', 'danger')
            return redirect(url_for('crear_hoja'))
        execute('INSERT INTO hojas_tareo(fecha,grupo,subgrupo,labor,responsable,estado,creado_por,creado_en) VALUES(?,?,?,?,?,?,?,?)',
                (fecha,grupo,subgrupo,labor,responsable,'ABIERTA',session.get('usuario'),now_str()), commit=True)
        hid = scalar('SELECT MAX(id) AS id FROM hojas_tareo')
        return redirect(url_for('detalle_hoja', hoja_id=hid))
    body = """
    <div class="phone-wrap desktop-pad"><h2 class="header-title">CREAR HOJA DE TAREO</h2><div class="page-card">
      <div class="panel-green"><i class="bi bi-clipboard2-plus"></i><h4>NUEVA HOJA – FECHA, GRUPO, SUBGRUPO, LABOR Y RESPONSABLE</h4></div>
      <form method="post" class="floating-card" style="margin:-24px 10px 12px">
        <label class="form-label">FECHA</label><input type="date" name="fecha" class="form-control mb-2" value="{{today}}" required>
        <label class="form-label">GRUPO</label><input name="grupo" class="form-control mb-2" list="grupos" placeholder="GRUPO COSECHA" required><datalist id="grupos"><option>GRUPO COSECHA</option><option>GRUPO CAMPO</option><option>GRUPO EMPAQUE</option></datalist>
        <label class="form-label">SUBGRUPO</label><input name="subgrupo" class="form-control mb-2" placeholder="SUBGRUPO / CUADRILLA">
        <label class="form-label">LABOR</label><input name="labor" class="form-control mb-2" list="labores" placeholder="COSECHA" required><datalist id="labores"><option>COSECHA</option><option>RALEO</option><option>PODA</option><option>LIMPIEZA</option></datalist>
        <label class="form-label">RESPONSABLE</label><input name="responsable" class="form-control mb-3" placeholder="APELLIDOS Y NOMBRES" required>
        <button class="btn btn-green w-100"><i class="bi bi-check-circle"></i> CREAR Y ENTRAR</button><a class="btn btn-outline-secondary w-100 mt-2" href="{{url_for('hojas_tareo')}}">VOLVER</a>
      </form></div></div>"""
    return render_page(body, today=today_str())

@app.route('/hoja/<int:hoja_id>')
@login_required
def detalle_hoja(hoja_id):
    tab = request.args.get('tab','labores')
    h = row_to_dict(execute('SELECT * FROM hojas_tareo WHERE id=?', (hoja_id,), fetchone=True))
    if not h:
        flash('Hoja no encontrada.', 'danger'); return redirect(url_for('hojas_tareo'))
    tareos = rows_to_dict(execute('SELECT * FROM tareos WHERE hoja_id=? ORDER BY creado_en DESC LIMIT 80', (hoja_id,), fetchall=True))
    lecturas = rows_to_dict(execute('SELECT * FROM lecturas_balde WHERE hoja_id=? ORDER BY fecha_hora DESC LIMIT 80', (hoja_id,), fetchall=True))
    trabajadores = rows_to_dict(execute('SELECT dni, trabajador FROM trabajadores WHERE estado="ACTIVO" ORDER BY trabajador LIMIT 200', fetchall=True)) if not is_pg() else rows_to_dict(execute("SELECT dni, trabajador FROM trabajadores WHERE estado='ACTIVO' ORDER BY trabajador LIMIT 200", fetchall=True))
    registros = len(tareos); horas_total = sum(float(x.get('horas') or 0) for x in tareos); rend_total = sum(float(x.get('cantidad') or 0) for x in tareos)
    execute('UPDATE hojas_tareo SET registros=?, horas_total=?, rendimiento_total=? WHERE id=?', (registros, horas_total, rend_total, hoja_id), commit=True)
    body = """
    <div class="phone-wrap desktop-pad"><h2 class="header-title">TAREO MÓVIL – {{ 'DETALLE DE TRABAJADOR POR LABOR' if tab=='trabajadores' else ('DETALLE NÚMERO DE LECTURAS POR BALDE' if tab=='rendimiento' else 'GRUPO DE COSECHA') }}</h2>
      <div class="page-card">
        <div class="tab-main"><a class="{{'active' if tab=='labores' else ''}}" href="{{url_for('detalle_hoja',hoja_id=h.id,tab='labores')}}">LABORES</a><a class="{{'active' if tab=='trabajadores' else ''}}" href="{{url_for('detalle_hoja',hoja_id=h.id,tab='trabajadores')}}">TRABAJADORES</a><a class="{{'active' if tab=='rendimiento' else ''}}" href="{{url_for('detalle_hoja',hoja_id=h.id,tab='rendimiento')}}">REND./AVANCE</a></div>
        <div class="subtabs"><a class="{{'active' if tab=='labores' else ''}}" href="{{url_for('detalle_hoja',hoja_id=h.id,tab='labores')}}">Labores</a><a class="{{'active' if tab=='trabajadores' else ''}}" href="{{url_for('detalle_hoja',hoja_id=h.id,tab='trabajadores')}}">Trab.por Labor</a><a class="{{'active' if tab=='rendimiento' else ''}}" href="{{url_for('detalle_hoja',hoja_id=h.id,tab='rendimiento')}}">Rend/Avance por Labor</a></div>
        <div class="panel-green"><i class="bi {{ 'bi-people-fill' if tab=='trabajadores' else ('bi-person-badge' if tab=='rendimiento' else 'bi-people') }}"></i><h4>{{ 'TRABAJADORES' if tab=='trabajadores' else ('PRODUCTIVIDAD POR TRABAJADOR' if tab=='rendimiento' else 'REGISTRO DE ACT-LAB-CONSUMIDOR') }}</h4></div>
        <div class="toolstrip"><button data-bs-toggle="modal" data-bs-target="#modalRegistro"><i class="bi bi-list-check"></i></button><i class="bi bi-search"></i>{% if tab=='trabajadores' %}<i class="bi bi-clock"></i><i class="bi bi-box-arrow-in-right"></i><i class="bi bi-person-plus"></i>{% elif tab=='rendimiento' %}<i class="bi bi-search"></i>{% else %}<i class="bi bi-files"></i><i class="bi bi-arrow-clockwise"></i>{% endif %}<a href="{{url_for('hojas_tareo')}}"><i class="bi bi-chevron-left"></i></a></div>
        <div class="info-bar"><div><i class="bi bi-calendar"></i> {{h.fecha}}</div><div><i class="bi bi-list"></i> {{registros}} Reg.</div><div><i class="bi bi-clock"></i> {{'%.2f'|format(horas_total)}} H.</div><div>A.Rend {{'%.2f'|format(rend_total)}}</div><span>⌄</span></div>
        {% if tab=='labores' %}
          {% for r in tareos %}<div class="worker-card"><div class="worker-title"><div>ACTIVIDAD<br><b>{{h.labor}}</b></div><div>A PORTAR / H.TRANSC / AVANCE<br><b>{{'%.2f'|format(r.horas or 0)}} &nbsp;&nbsp; {{'%.2f'|format(r.cantidad or 0)}}</b></div></div><div class="mt-2"><span class="small-label">ITEM</span> <b>001</b> &nbsp; <span class="small-label">CONSUMIDOR</span> <b>{{r.lote or 'FUNDO HEFEI'}}</b></div><div class="mt-1"><span class="person-dot"><i class="bi bi-people-fill"></i></span><div class="small-label">RESPONSABLE</div><div class="small-value">{{h.responsable}}</div><div class="small-label mt-1">GRUPO</div><div class="small-value">{{h.grupo}}</div></div><div class="worker-grid mt-2"><div><div class="mini-badge bg-y">{{r.unidad or 'DIURNO-NOCT'}}</div></div><div><div class="mini-badge bg-y">JORNAL</div></div><div><div class="mini-badge bg-g">F.TOTAL</div></div></div></div>{% else %}<div class="worker-card text-center text-muted">Presiona <b>+</b> para registrar la primera labor.</div>{% endfor %}
        {% elif tab=='trabajadores' %}
          {% for r in tareos %}<div class="worker-card"><div class="worker-title"><div>TRABAJADOR<br><b>{{r.trabajador}}</b></div><div>NRO.DOCUMENTO<br><b>{{r.dni}}</b></div></div><div class="worker-grid"><div><label>HORA INICIO</label><input class="mini-input" value="06:30"></div><div><label>HORA FIN</label><input class="mini-input" value="16:30"></div><div><label>H.NORMAL</label><input class="mini-input" value="{{'%.2f'|format(r.horas or 0)}}"></div><div><label>A.DIURNO</label><input class="mini-input" value="{{'%.2f'|format(r.cantidad or 0)}}"></div><div><label>A.NOCT</label><input class="mini-input" value="0.00"></div><div><label>ESTADO</label><div class="mini-badge bg-g">FIN TOTAL</div></div></div></div>{% else %}<div class="worker-card text-center text-muted">Sin trabajadores registrados en esta hoja.</div>{% endfor %}
        {% else %}
          {% for l in lecturas %}<div class="worker-card"><span class="person-dot"><i class="bi bi-person-circle"></i></span><div class="worker-title"><div>TRABAJADOR<br><b>{{l.trabajador}}</b></div><div>NRO.DOC.<br><b>{{l.dni}}</b></div></div><div class="small-label mt-1">HORA TOMA REGISTRO</div><div class="small-value">{{l.fecha_hora}}</div><div class="worker-grid"><div><label>A.DIURNO</label><div class="mini-badge bg-y">{{'%.2f'|format(l.a_diurno or 0)}}</div></div><div><label>A.NOCT.</label><div class="mini-badge bg-y">{{'%.2f'|format(l.a_noct or 0)}}</div></div><div class="text-end"><i class="bi bi-chevron-left text-success"></i></div></div></div>{% else %}<div class="worker-card text-center text-muted">Sin lecturas de balde registradas.</div>{% endfor %}
        {% endif %}<div class="leaf"></div>
      </div>
    </div>
    <div class="modal fade" id="modalRegistro" tabindex="-1"><div class="modal-dialog modal-dialog-centered"><div class="modal-content"><form method="post" action="{{url_for('guardar_registro_hoja', hoja_id=h.id, tab=tab)}}"><div class="modal-header"><h5 class="modal-title fw-bold text-success">Registrar en hoja</h5><button class="btn-close" data-bs-dismiss="modal"></button></div><div class="modal-body"><label class="form-label">TRABAJADOR / DNI</label><select name="dni" class="form-select mb-2" required>{% for t in trabajadores %}<option value="{{t.dni}}">{{t.dni}} - {{t.trabajador}}</option>{% endfor %}</select><div class="row g-2"><div class="col-6"><label class="form-label">HORAS</label><input name="horas" type="number" step="0.01" class="form-control" value="9.75"></div><div class="col-6"><label class="form-label">AVANCE</label><input name="cantidad" type="number" step="0.01" class="form-control" value="7.00"></div><div class="col-6"><label class="form-label">CONSUMIDOR / LOTE</label><input name="lote" class="form-control" value="FUNDO HEFEI"></div><div class="col-6"><label class="form-label">UNIDAD</label><select name="unidad" class="form-select"><option>JORNAL</option><option>KG</option><option>BALDE</option><option>JABA</option></select></div></div></div><div class="modal-footer"><button class="btn btn-green w-100">GUARDAR</button></div></form></div></div></div>
    """
    return render_page(body, h=h, tab=tab, tareos=tareos, lecturas=lecturas, trabajadores=trabajadores, registros=registros, horas_total=horas_total, rend_total=rend_total)

@app.route('/hoja/<int:hoja_id>/registro/<tab>', methods=['POST'])
@login_required
def guardar_registro_hoja(hoja_id, tab):
    h = row_to_dict(execute('SELECT * FROM hojas_tareo WHERE id=?', (hoja_id,), fetchone=True))
    if not h: flash('Hoja no encontrada.', 'danger'); return redirect(url_for('hojas_tareo'))
    dni = limpiar_dni(request.form.get('dni'))
    t = row_to_dict(execute('SELECT * FROM trabajadores WHERE dni=?', (dni,), fetchone=True))
    if not t: flash('Trabajador no encontrado en base.', 'danger'); return redirect(url_for('detalle_hoja', hoja_id=hoja_id, tab=tab))
    horas = float(request.form.get('horas') or 0); cantidad = float(request.form.get('cantidad') or 0)
    if tab == 'rendimiento':
        execute('INSERT INTO lecturas_balde(hoja_id,dni,trabajador,fecha_hora,a_diurno,a_noct,registrado_por) VALUES(?,?,?,?,?,?,?)',
                (hoja_id,dni,t.get('trabajador',''),now_str(),cantidad,0,session.get('usuario')), commit=True)
    execute('''INSERT INTO tareos(hoja_id,dni,trabajador,empresa,area,cargo,fecha,labor,lote,fundo,horas,cantidad,unidad,observacion,registrado_por,creado_en)
               VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',
            (hoja_id,dni,t.get('trabajador',''),t.get('empresa',''),t.get('area',''),t.get('cargo',''),h.get('fecha'),h.get('labor'),limpiar_texto(request.form.get('lote')),h.get('grupo'),horas,cantidad,limpiar_texto(request.form.get('unidad')),'',session.get('usuario'),now_str()), commit=True)
    flash('Registro guardado correctamente.', 'success')
    return redirect(url_for('detalle_hoja', hoja_id=hoja_id, tab=tab))

# ========================= SOPORTE / CONFIG / SINC =========================
@app.route('/soporte')
@login_required
def soporte():
    return render_page('<div class="phone-wrap"><h2 class="header-title">SOPORTE</h2><div class="floating-card"><b>Canal de soporte</b><p class="small text-muted mb-2">Registra incidencias de sincronización, acceso o lectura de fotocheck.</p><textarea class="form-control" rows="4" placeholder="Describe el problema..."></textarea><a class="btn btn-green w-100 mt-3" href="{{url_for(\'home\')}}">ENVIAR / VOLVER</a></div></div>')

@app.route('/configuraciones')
@login_required
def configuraciones():
    body = '<div class="phone-wrap"><h2 class="header-title">CONFIGURACIONES</h2><div class="floating-card"><a class="btn btn-green w-100 mb-2" href="{{url_for(\'cargar_base\')}}"><i class="bi bi-file-earmark-excel"></i> Base trabajadores</a><a class="btn btn-outline-success w-100 mb-2" href="{{url_for(\'usuarios\')}}"><i class="bi bi-people"></i> Usuarios</a><a class="btn btn-outline-secondary w-100" href="{{url_for(\'home\')}}">Volver</a></div></div>'
    return render_page(body)

@app.route('/sincronizacion')
@login_required
def sincronizacion():
    total = scalar('SELECT COUNT(*) AS c FROM trabajadores')
    body = '<div class="phone-wrap"><h2 class="header-title">SINCRONIZACIÓN</h2><div class="panel-green"><i class="bi bi-arrow-repeat"></i><h4>TABLAS MAESTRAS</h4></div><div class="floating-card" style="margin:-20px 10px 0"><p><b>Trabajadores:</b> {{total}}</p><p><b>Última actualización:</b> {{now}}</p><a class="btn btn-green w-100" href="{{url_for(\'home\')}}">SINCRONIZADO</a></div></div>'
    return render_page(body, total=total, now=now_str())

# ========================= REPORTES / CARGA / USUARIOS =========================
@app.route('/reportes')
@login_required
def reportes():
    desde = request.args.get('desde') or today_str(); hasta = request.args.get('hasta') or today_str(); q=request.args.get('q','').strip()
    params=[desde,hasta]; where='WHERE fecha>=? AND fecha<=?'
    if q:
        like=f"%{q.upper()}%"; where += ' AND (dni LIKE ? OR UPPER(trabajador) LIKE ? OR UPPER(labor) LIKE ?)'; params += [like,like,like]
    tareos = rows_to_dict(execute(f'SELECT * FROM tareos {where} ORDER BY creado_en DESC LIMIT 500', params, fetchall=True))
    body = """
    <div class="phone-wrap desktop-pad"><h2 class="header-title">REPORTES TAREO</h2><form class="floating-card mb-2"><div class="row g-2"><div class="col-6"><input class="form-control" type="date" name="desde" value="{{desde}}"></div><div class="col-6"><input class="form-control" type="date" name="hasta" value="{{hasta}}"></div><div class="col-9"><input class="form-control" name="q" value="{{q}}" placeholder="DNI / trabajador / labor"></div><div class="col-3 d-grid"><button class="btn btn-green"><i class="bi bi-search"></i></button></div></div><a class="btn btn-outline-success w-100 mt-2" href="{{url_for('exportar_tareos',desde=desde,hasta=hasta,q=q)}}">EXPORTAR EXCEL</a></form>{% for r in tareos %}<div class="worker-card"><div class="worker-title"><div>{{r.fecha}}<br><b>{{r.trabajador}}</b></div><div class="text-end">{{r.dni}}<br><b>{{r.labor}}</b></div></div><div class="worker-grid"><div><label>HORAS</label><div class="mini-input">{{r.horas}}</div></div><div><label>CANT.</label><div class="mini-input">{{r.cantidad}}</div></div><div><label>UNIDAD</label><div class="mini-input">{{r.unidad}}</div></div></div></div>{% else %}<div class="alert alert-light border text-center">Sin datos.</div>{% endfor %}</div>"""
    return render_page(body, desde=desde, hasta=hasta, q=q, tareos=tareos)

@app.route('/exportar/tareos')
@login_required
def exportar_tareos():
    desde=request.args.get('desde') or today_str(); hasta=request.args.get('hasta') or today_str(); q=request.args.get('q','').strip()
    params=[desde,hasta]; where='WHERE fecha>=? AND fecha<=?'
    if q:
        like=f"%{q.upper()}%"; where += ' AND (dni LIKE ? OR UPPER(trabajador) LIKE ? OR UPPER(labor) LIKE ?)'; params += [like,like,like]
    rows=rows_to_dict(execute(f'SELECT fecha,dni,trabajador,empresa,area,cargo,labor,fundo,lote,horas,cantidad,unidad,observacion,registrado_por,creado_en FROM tareos {where} ORDER BY creado_en DESC', params, fetchall=True))
    headers=['FECHA','DNI','TRABAJADOR','EMPRESA','AREA','CARGO','LABOR','FUNDO','LOTE','HORAS','CANTIDAD','UNIDAD','OBSERVACION','REGISTRADO_POR','CREADO_EN']
    return excel_response(headers, rows, f'tareos_{desde}_a_{hasta}.xlsx', 'TAREOS')

@app.route('/cargar-base', methods=['GET','POST'])
@admin_required
def cargar_base():
    if request.method == 'POST':
        f = request.files.get('archivo')
        if not f or not f.filename.lower().endswith('.xlsx'):
            flash('Suba un Excel .xlsx válido.', 'danger'); return redirect(url_for('cargar_base'))
        wb = load_workbook(f, data_only=True, read_only=True); ws = wb.active
        rows = list(ws.iter_rows(values_only=True))
        if not rows: flash('Excel vacío.', 'danger'); return redirect(url_for('cargar_base'))
        headers = [normalizar_columna(c) for c in rows[0]]
        if 'DNI' not in headers:
            flash('La plantilla debe tener la columna DNI.', 'danger'); return redirect(url_for('cargar_base'))
        colmap = {'TRABAJADOR':'trabajador','NOMBRE':'trabajador','APELLIDOS Y NOMBRES':'trabajador','EMPRESA':'empresa','AREA':'area','ÁREA':'area','CARGO':'cargo','ACTIVIDAD':'actividad','PLANILLA':'planilla','ESTADO':'estado'}
        ins=upd=omi=0; conn=get_conn(); cur=conn.cursor(); ahora=now_str()
        for row in rows[1:]:
            item={headers[i]: row[i] if i < len(row) else '' for i in range(len(headers))}
            dni=limpiar_dni(item.get('DNI'))
            if len(dni)!=8: omi+=1; continue
            data={'trabajador':'','empresa':'','area':'','cargo':'','actividad':'','planilla':'','estado':'ACTIVO'}
            for col,key in colmap.items():
                if col in headers and item.get(col) not in (None,''): data[key]=limpiar_texto(item.get(col))
            cur.execute(qmark('SELECT id FROM trabajadores WHERE dni=?'), (dni,))
            if cur.fetchone():
                cur.execute(qmark('UPDATE trabajadores SET trabajador=?,empresa=?,area=?,cargo=?,actividad=?,planilla=?,estado=?,fecha_carga=? WHERE dni=?'), (data['trabajador'],data['empresa'],data['area'],data['cargo'],data['actividad'],data['planilla'],data['estado'],ahora,dni)); upd+=1
            else:
                cur.execute(qmark('INSERT INTO trabajadores(dni,trabajador,empresa,area,cargo,actividad,planilla,estado,fecha_carga) VALUES(?,?,?,?,?,?,?,?,?)'), (dni,data['trabajador'],data['empresa'],data['area'],data['cargo'],data['actividad'],data['planilla'],data['estado'],ahora)); ins+=1
        conn.commit(); cur.close(); conn.close(); flash(f'Carga completa. Insertados: {ins} | Actualizados: {upd} | Omitidos: {omi}', 'success')
        return redirect(url_for('cargar_base'))
    trabajadores = rows_to_dict(execute('SELECT * FROM trabajadores ORDER BY fecha_carga DESC, trabajador LIMIT 100', fetchall=True))
    body = """
    <div class="phone-wrap desktop-pad"><h2 class="header-title">CONFIGURACIÓN – TRABAJADORES</h2><form method="post" enctype="multipart/form-data" class="floating-card mb-2"><label class="form-label">Archivo Excel .xlsx</label><input class="form-control mb-2" type="file" name="archivo" accept=".xlsx" required><button class="btn btn-green w-100">CARGAR BASE</button><a class="btn btn-outline-success w-100 mt-2" href="{{url_for('plantilla_trabajadores')}}">PLANTILLA EXCEL</a></form>{% for r in trabajadores %}<div class="worker-card"><div class="worker-title"><div>{{r.dni}}<br><b>{{r.trabajador}}</b></div><div class="text-end">{{r.empresa}}<br><b>{{r.cargo}}</b></div></div></div>{% endfor %}</div>"""
    return render_page(body, trabajadores=trabajadores)

@app.route('/plantilla-trabajadores')
@admin_required
def plantilla_trabajadores():
    headers=['DNI','TRABAJADOR','EMPRESA','AREA','CARGO','ACTIVIDAD','PLANILLA','ESTADO']
    rows=[{'DNI':'12345678','TRABAJADOR':'APELLIDOS Y NOMBRES','EMPRESA':'AQUANQA I','AREA':'CAMPO','CARGO':'OPERARIO','ACTIVIDAD':'COSECHA','PLANILLA':'AGRARIO','ESTADO':'ACTIVO'}]
    return excel_response(headers, rows, 'plantilla_trabajadores_tareo_movil.xlsx', 'TRABAJADORES')

@app.route('/usuarios', methods=['GET','POST'])
@admin_required
def usuarios():
    if request.method == 'POST':
        usuario=request.form.get('usuario','').strip(); nombres=limpiar_texto(request.form.get('nombres')); clave=request.form.get('clave',''); rol=request.form.get('rol','operador'); estado=request.form.get('estado','ACTIVO')
        if not usuario or not clave: flash('Usuario y clave son obligatorios.', 'danger')
        else:
            try:
                execute('INSERT INTO usuarios(usuario,password_hash,nombres,rol,estado,creado_en) VALUES(?,?,?,?,?,?)', (usuario,generate_password_hash(clave),nombres,rol,estado,now_str()), commit=True); flash('Usuario creado.', 'success')
            except Exception as e: flash(f'No se pudo crear usuario: {e}', 'danger')
        return redirect(url_for('usuarios'))
    users=rows_to_dict(execute('SELECT usuario,nombres,rol,estado,creado_en FROM usuarios ORDER BY usuario', fetchall=True))
    body='<div class="phone-wrap"><h2 class="header-title">USUARIOS</h2><form method="post" class="floating-card mb-2"><input class="form-control mb-2" name="usuario" placeholder="Usuario" required><input class="form-control mb-2" name="nombres" placeholder="Nombres"><input class="form-control mb-2" type="password" name="clave" placeholder="Clave" required><select class="form-select mb-2" name="rol"><option value="operador">operador</option><option value="admin">admin</option></select><select class="form-select mb-2" name="estado"><option>ACTIVO</option><option>INACTIVO</option></select><button class="btn btn-green w-100">Guardar</button></form>{% for u in users %}<div class="worker-card"><b>{{u.usuario}}</b> · {{u.rol}}<br><span class="small text-muted">{{u.nombres}} · {{u.estado}}</span></div>{% endfor %}</div>'
    return render_page(body, users=users)

# ========================= COMPAT API MARCACIÓN BÁSICA =========================
@app.route('/api/trabajador/<dni>')
@login_required
def api_trabajador(dni):
    dni = limpiar_dni(dni)
    if len(dni) != 8: return jsonify(ok=False, msg='DNI inválido.')
    t = row_to_dict(execute('SELECT * FROM trabajadores WHERE dni=?', (dni,), fetchone=True))
    if not t: return jsonify(ok=False, msg='DNI no encontrado en la base de trabajadores.')
    return jsonify(ok=True, trabajador=t, sugerido='ENTRADA')

# ========================= PWA =========================
@app.route('/manifest.webmanifest')
@app.route('/manifest.json')
def manifest():
    return jsonify({"name":"Tareo Móvil PRIZE","short_name":"Tareo","start_url":"/","display":"standalone","background_color":"#ffffff","theme_color":"#2f773b","icons":[]})

@app.route('/sw.js')
def sw():
    return Response("self.addEventListener('install',e=>self.skipWaiting()); self.addEventListener('fetch',e=>{});", mimetype='application/javascript')

try:
    init_db()
except Exception as e:
    print('ERROR inicializando base de datos:', e)

if __name__ == '__main__':
    port = int(os.getenv('PORT', '5000'))
    app.run(host='0.0.0.0', port=port, debug=False)
