# -*- coding: utf-8 -*-
"""
TAREO MÓVIL – GRUPO DE COSECHA | PRIZE PRO
Flask + SQLite/PostgreSQL + PWA. Listo para GitHub + Render.

Mejoras incluidas:
- Login con formato móvil verde/blanco como referencia.
- Pantalla principal: Soporte, Configuraciones, Sincronización y Hojas de Tareo.
- Creación de hoja: fecha, grupo, subgrupo, labor, responsable, turno y tipo.
- Detalle de hoja con tabs: Labores, Trabajadores, Rend./Avance por Labor, con iconos funcionales.
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

def _add_column_if_missing(cur, table, column, ddl):
    """Migración segura para SQLite/PostgreSQL."""
    try:
        if is_pg():
            cur.execute(qmark(f"ALTER TABLE {table} ADD COLUMN IF NOT EXISTS {column} {ddl}"))
        else:
            cur.execute(f"PRAGMA table_info({table})")
            cols = [r[1] for r in cur.fetchall()]
            if column not in cols:
                cur.execute(f"ALTER TABLE {table} ADD COLUMN {column} {ddl}")
    except Exception as e:
        print(f"No se pudo agregar columna {table}.{column}:", e)

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
        id {idtype}, hoja_id INTEGER, labor_id INTEGER, dni TEXT NOT NULL, trabajador TEXT, empresa TEXT, area TEXT, cargo TEXT,
        fecha TEXT NOT NULL, labor TEXT, lote TEXT, fundo TEXT, horas REAL DEFAULT 0,
        cantidad REAL DEFAULT 0, unidad TEXT, observacion TEXT, registrado_por TEXT, creado_en TEXT,
        hora_inicio TEXT, hora_fin TEXT, ref_inicio TEXT, ref_fin TEXT, turno TEXT, tipo_tareo TEXT)"""))
    cur.execute(qmark(f"""CREATE TABLE IF NOT EXISTS hojas_tareo(
        id {idtype}, fecha TEXT NOT NULL, grupo TEXT, subgrupo TEXT, labor TEXT, responsable TEXT,
        turno TEXT DEFAULT 'DIA', tipo_tareo TEXT DEFAULT 'JORNAL',
        estado TEXT DEFAULT 'ABIERTA', registros INTEGER DEFAULT 0, horas_total REAL DEFAULT 0, rendimiento_total REAL DEFAULT 0,
        creado_por TEXT, creado_en TEXT)"""))
    cur.execute(qmark(f"""CREATE TABLE IF NOT EXISTS hoja_labores(
        id {idtype}, hoja_id INTEGER NOT NULL, grupo TEXT, subgrupo TEXT, labor TEXT,
        turno TEXT DEFAULT 'DIA', tipo_tareo TEXT DEFAULT 'JORNAL', responsable TEXT, creado_en TEXT, creado_por TEXT)"""))
    cur.execute(qmark(f"""CREATE TABLE IF NOT EXISTS actividades_maestras(
        id {idtype}, cod_actividad TEXT, desc_actividad TEXT, cod_labor TEXT, desc_labor TEXT,
        cod_consumidor TEXT, desc_consumidor TEXT, estado TEXT DEFAULT 'ACTIVO', fecha_carga TEXT)"""))
    cur.execute(qmark(f"""CREATE TABLE IF NOT EXISTS lecturas_balde(
        id {idtype}, hoja_id INTEGER, labor_id INTEGER, dni TEXT, trabajador TEXT, fecha_hora TEXT,
        a_diurno REAL DEFAULT 0, a_noct REAL DEFAULT 0, metodo TEXT, registrado_por TEXT)"""))

    # Migraciones sobre bases ya existentes en Render
    for col, ddl in [('turno', "TEXT DEFAULT 'DIA'"), ('tipo_tareo', "TEXT DEFAULT 'JORNAL'")]:
        _add_column_if_missing(cur, 'hojas_tareo', col, ddl)
    for col, ddl in [('labor_id','INTEGER'),('hora_inicio','TEXT'),('hora_fin','TEXT'),('ref_inicio','TEXT'),('ref_fin','TEXT'),('turno','TEXT'),('tipo_tareo','TEXT')]:
        _add_column_if_missing(cur, 'tareos', col, ddl)
    for col, ddl in [('labor_id','INTEGER'),('metodo','TEXT')]:
        _add_column_if_missing(cur, 'lecturas_balde', col, ddl)

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
*{box-sizing:border-box}body{margin:0;background:#fff;font-family:Inter,Segoe UI,Arial,sans-serif;color:#21472a}.app-bg{min-height:100vh;background:linear-gradient(180deg,#fff 0%,#fff 62%,#fbfbfb 100%)}.shell{width:min(1120px,100%);margin:0 auto;padding:18px}.phone-wrap{max-width:430px;margin:0 auto}.desktop-grid{display:grid;grid-template-columns:360px 1fr;gap:28px;align-items:start}.home-desktop-list .worker-card{margin:10px 0}.home-desktop-list{max-width:640px;margin:0 auto}.header-title{text-align:center;color:#166534;font-family:Georgia,serif;font-weight:900;letter-spacing:.5px;font-size:23px;line-height:1.13;margin:4px 0 22px;text-transform:uppercase}.green-hero{background:var(--verde);border-radius:0 0 18px 18px;min-height:145px;padding:12px 16px 22px;color:white;text-align:center;position:relative;overflow:visible}.tareo-hero{min-height:124px!important;padding-bottom:42px!important}.tareo-toolbar{margin:12px 12px 8px!important;position:relative;z-index:4}.tareo-list-page .worker-card:first-of-type{margin-top:12px}.back-mini{display:inline-grid;place-items:center;width:36px;height:36px;border-radius:999px;color:var(--verde);text-decoration:none;font-size:24px}.green-top{display:flex;justify-content:space-between;align-items:center;font-size:11px;font-weight:800}.avatar{width:78px;height:78px;border-radius:999px;background:white;color:var(--verde);display:grid;place-items:center;margin:10px auto 2px;font-size:43px;box-shadow:0 8px 20px rgba(0,0,0,.13)}.login-name{font-size:11px;font-weight:800}.white-input{height:36px;background:white;border-radius:10px;box-shadow:0 5px 13px rgba(0,0,0,.18);border:0}.floating-card{background:white;border-radius:10px;box-shadow:0 8px 18px rgba(0,0,0,.15);padding:12px}.tile{width:74px;height:70px;border-radius:8px;background:white;box-shadow:0 7px 17px rgba(0,0,0,.14);display:flex;flex-direction:column;align-items:center;justify-content:center;color:var(--verde);font-weight:900;font-size:10px;text-align:center}.tile i{font-size:26px;margin-bottom:5px}.bottom-sync{position:fixed;left:10px;bottom:8px;color:#317a3e;font-size:10px;font-weight:700}.bottom-out{position:fixed;right:14px;bottom:8px;color:#c84c4c;font-size:20px}.tab-main{display:flex;gap:0;background:#fafafa;padding-left:0;border-top:4px solid #d7d7d7}.tab-main a{flex:1;text-align:center;text-decoration:none;color:#508557;font-weight:900;font-size:13px;padding:14px 8px;border-radius:7px 7px 0 0;background:#fff}.tab-main a.active{background:var(--verde);color:white;box-shadow:0 3px 7px rgba(0,0,0,.18)}.subtabs{display:flex;background:#fff}.subtabs a{flex:1;text-align:center;padding:13px 5px;text-decoration:none;color:#4b8a54;font-weight:900;font-size:12px}.subtabs a.active{background:var(--verde);color:white}.panel-green{background:var(--verde);color:white;text-align:center;padding:21px 12px 42px}.panel-green i{font-size:38px}.panel-green h4{font-size:11px;font-weight:900;margin:5px 0 0}.toolstrip{background:white;margin:-25px 9px 5px;border-radius:9px;min-height:49px;box-shadow:0 5px 13px rgba(0,0,0,.22);display:flex;align-items:center;gap:20px;padding:7px 14px;color:var(--verde);font-size:24px}.toolstrip button,.toolstrip a{border:0;background:transparent;color:var(--verde);font-size:24px;text-decoration:none}.info-bar{margin:0 9px;background:var(--verde);color:white;border-radius:2px;display:grid;grid-template-columns:1fr 1fr 1fr 1fr 22px;align-items:center;font-size:10px;font-weight:900;height:23px}.info-bar div{text-align:center;border-right:1px solid rgba(255,255,255,.28)}.worker-card{background:white;margin:10px 12px;border-radius:10px;box-shadow:0 3px 12px rgba(0,0,0,.20);padding:11px 13px;color:#397443;position:relative}.worker-title{display:grid;grid-template-columns:1fr 1fr;gap:8px;font-size:9px;font-weight:900;text-transform:uppercase}.worker-title b{font-size:10px;color:var(--verde)}.worker-grid{display:grid;grid-template-columns:1fr 1fr 1fr;gap:7px;margin-top:7px}.worker-grid label{font-size:8px;font-weight:900;color:#6c8a6f;margin-bottom:1px}.mini-input{height:25px;border:1px solid #9ebaa0;border-radius:3px;font-size:10px;padding:3px 5px;width:100%;font-weight:800;color:#315f39}.mini-badge{border-radius:3px;color:white;font-size:8px;font-weight:900;text-align:center;padding:4px 3px;text-transform:uppercase}.bg-y{background:var(--amarillo)!important;color:white}.bg-g{background:#42b852!important}.person-dot{width:42px;height:42px;border-radius:999px;background:#378145;color:white;display:grid;place-items:center;font-size:26px;float:left;margin-right:9px}.small-label{font-size:8px;color:#79937b;font-weight:900}.small-value{font-size:9px;color:#466a49;font-weight:900}.leaf{width:120px;height:120px;border-radius:70% 30% 70% 30%;background:linear-gradient(135deg,#ffd8bd,#eef4c7,#cbd9b6);opacity:.72;margin:20px auto 0;transform:rotate(-20deg)}.card-pro{background:white;border:1px solid var(--line);border-radius:18px;box-shadow:0 8px 20px rgba(0,0,0,.10)}.btn-green{background:var(--verde);border-color:var(--verde);color:white;font-weight:900;border-radius:9px}.btn-green:hover{background:var(--verde3);color:white}.form-control,.form-select{border-radius:9px;border:1px solid #dfe7df;font-weight:700;font-size:13px}.form-label{font-size:12px;font-weight:900;color:#3e7545}.page-card{border-radius:13px;overflow:hidden;border:1px solid #e5e7e5;background:white}.list-table th{font-size:11px;color:#497550}.list-table td{font-size:12px;vertical-align:middle}.status-pill{display:inline-block;background:#39b54a;color:white;border-radius:4px;padding:4px 8px;font-size:9px;font-weight:900}.top-actions{display:flex;gap:12px;flex-wrap:wrap;justify-content:center;margin-top:14px;position:relative;z-index:5}.top-actions .tile{width:82px;height:76px}.login-page .shell{padding:0}.login-form{margin:-7px auto 0;width:92%;max-width:360px}.login-form .floating-card{padding:13px 14px 18px}.alert{border-radius:12px;font-size:13px}.desk-panel{display:block}.mobile-only{display:none}.clock-box{width:116px;height:116px;border:5px solid var(--verde);border-radius:999px;margin:8px auto;display:grid;place-items:center;color:var(--verde);font-weight:900;background:#fff;box-shadow:0 4px 14px rgba(47,119,59,.18)}.clock-box i{font-size:38px}.scan-box{border:2px dashed #8dbf93;border-radius:12px;padding:10px;background:#f8fff9}.toolstrip .hint{font-size:9px;font-weight:900;color:#2f773b;margin-left:-18px;margin-right:0}.splash-card{height:94vh;max-height:760px;background:#23773f;border-radius:10px;box-shadow:0 4px 12px rgba(0,0,0,.25);display:flex;flex-direction:column;align-items:center;justify-content:center;color:white;position:relative}.splash-logo{width:145px;height:145px;border-radius:999px;background:#fff;border:6px solid #92bd33;display:grid;place-items:center;color:#23773f;font-size:66px;box-shadow:0 3px 10px rgba(0,0,0,.22)}.splash-title{font-weight:900;margin-top:18px;letter-spacing:.5px}.splash-foot{position:absolute;bottom:26px;text-align:center;font-size:11px;color:#d9f2df;font-weight:700}.role-toggle{display:grid;grid-template-columns:1fr 1fr;gap:8px}.role-toggle label{border:1px solid #dce7dc;border-radius:9px;padding:9px;text-align:center;font-size:12px;font-weight:900;color:#2f773b}.role-toggle input{display:none}.role-toggle input:checked+span{background:#2f773b;color:white;border-radius:7px;padding:7px 9px;display:block}.bottom-nav{position:sticky;bottom:0;background:white;border-top:1px solid #e8ece8;display:flex;justify-content:space-around;padding:7px 0;color:#477b4d;font-size:10px;font-weight:800}.bottom-nav a{text-decoration:none;color:#477b4d;text-align:center}.bottom-nav i{display:block;font-size:17px}.copy-list{max-height:260px;overflow:auto;border:1px solid #e5ede5;border-radius:9px;padding:8px;background:#fbfffb}.clock-face{width:180px;height:180px;border-radius:999px;background:#e7e0ef;margin:8px auto;position:relative;display:grid;place-items:center;color:#5d44aa;touch-action:none;cursor:pointer;user-select:none}.clock-hand{width:65px;height:3px;background:#6b4eb8;position:absolute;transform-origin:left center;transform:rotate(-35deg);left:90px;top:90px}.clock-dot{width:12px;height:12px;border-radius:999px;background:#6b4eb8;position:absolute;left:84px;top:84px}.clock-num{position:absolute;font-size:12px;color:#37303c}.clock-bubble{position:absolute;right:20px;top:50px;background:#6b4eb8;color:white;border-radius:999px;padding:9px;font-weight:900}.field-required{box-shadow:inset 4px 0 0 var(--verde)}.big-plus{font-size:34px!important;line-height:1}.big-plus .bi-plus{font-size:22px!important;margin-left:-12px;font-weight:900}.labor-card-compact{padding:13px 16px}.labor-card-compact .worker-title{font-size:8px}.labor-card-compact .worker-title b{font-size:9px;line-height:1.15}.labor-card-compact .labor-main{font-size:15px!important;line-height:1.1;color:#146c35}.labor-card-compact .resp-main{font-size:13px!important;line-height:1.1;color:#146c35}.worker-queue{border:1px dashed #8cc79b;border-radius:12px;background:#f7fff8;padding:9px;margin-top:10px;max-height:155px;overflow:auto}.queue-item{display:flex;justify-content:space-between;gap:8px;align-items:center;border-bottom:1px solid #e2f3e5;padding:6px 0;font-size:12px}.queue-item:last-child{border-bottom:0}.scan-ok{background:#d1fae5;border:1px solid #86efac;color:#166534;border-radius:10px;padding:8px;font-size:12px;font-weight:800}.scan-bad{background:#fee2e2;border:1px solid #fecaca;color:#991b1b;border-radius:10px;padding:8px;font-size:12px;font-weight:800}.time-click{cursor:pointer;background:#fbfffb}.time-click:focus{outline:2px solid #2f773b}.report-wrap{max-width:540px;margin:0 auto}.config-header{display:flex;align-items:center;gap:8px;justify-content:center;position:relative}.config-header .back-mini{position:absolute;left:0}.btn-plus-fab{display:inline-flex!important;align-items:center;gap:2px}.btn-plus-fab i:first-child{font-size:30px!important}.btn-plus-fab i:last-child{font-size:20px!important;margin-left:-12px;margin-top:12px}.field-help{font-size:10px;color:#5f7d65;font-weight:800} 
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


def get_actividades_maestras(limit=5000):
    try:
        rows = rows_to_dict(execute("SELECT * FROM actividades_maestras WHERE COALESCE(estado,'ACTIVO')='ACTIVO' ORDER BY desc_actividad, desc_labor, desc_consumidor LIMIT ?", (limit,), fetchall=True))
    except Exception:
        rows = []
    return rows

def js_master_options(rows):
    import json
    return json.dumps(rows, ensure_ascii=False)

# ========================= AUTH + HOME =========================

@app.route('/inicio')
def inicio():
    body = """
    <div class="phone-wrap"><div class="splash-card">
      <div class="splash-logo"><i class="bi bi-clipboard2-data"></i></div>
      <div class="splash-title">TAREO MOVIL</div>
      <a class="btn btn-light btn-sm mt-4 fw-bold text-success" href="{{url_for('login')}}">ENTRAR</a>
      <div class="splash-foot">Nisira Systems S.A.C.<br>v.1.0</div>
    </div></div>"""
    return render_page(body, title='Tareo Móvil')

@app.route('/login', methods=['GET','POST'])
def login():
    if request.method == 'POST':
        usuario = request.form.get('usuario','').strip()
        password = request.form.get('password','')
        row = row_to_dict(execute('SELECT * FROM usuarios WHERE usuario=?', (usuario,), fetchone=True))
        if row and row.get('estado','ACTIVO') == 'ACTIVO' and check_password_hash(row['password_hash'], password):
            if request.form.get('modo') == 'admin' and row.get('rol') != 'admin':
                flash('Este acceso es solo para administradores.', 'danger'); return redirect(url_for('login'))
            session['usuario']=usuario; session['rol']=row.get('rol','operador'); session['nombres']=row.get('nombres') or usuario
            return redirect(url_for('home'))
        flash('Usuario/clave incorrecta o usuario inactivo.', 'danger')
    body = """
    <div class="phone-wrap">
      <div class="green-hero" style="min-height:245px;border-radius:0 0 22px 22px">
        <div class="green-top"><span><i class="bi bi-headset"></i> Soporte</span><span><i class="bi bi-gear"></i> Config.</span></div>
        <div class="splash-logo" style="width:96px;height:96px;font-size:42px;margin:14px auto 6px"><i class="bi bi-clipboard2-data"></i></div>
        <div class="splash-title">TAREO MOVIL</div><div class="login-name">INICIAR SESIÓN</div>
      </div>
      <form method="post" class="login-form">
        <div class="floating-card">
          <div class="role-toggle mb-2"><label><input type="radio" name="modo" value="usuario" checked><span>USUARIO</span></label><label><input type="radio" name="modo" value="admin"><span>ADMINISTRADOR</span></label></div>
          <input class="form-control white-input mb-2" name="usuario" required autofocus placeholder="Usuario">
          <input class="form-control white-input mb-3" name="password" type="password" required placeholder="Clave">
          <button class="btn btn-green w-100"><i class="bi bi-box-arrow-in-right me-1"></i> INGRESAR</button>
          <div class="text-center small mt-2 text-muted">Admin demo: admin / admin123</div>
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
def home():
    if not session.get('usuario'):
        return redirect(url_for('inicio'))
    hojas = rows_to_dict(execute('SELECT * FROM hojas_tareo ORDER BY fecha DESC, id DESC LIMIT 8', fetchall=True))
    body = """
    <div class="desktop-grid">
      <div class="phone-wrap">
        <div class="green-hero" style="min-height:220px">
          <div class="green-top"><a class="text-white text-decoration-none" href="{{url_for('soporte')}}"><i class="bi bi-headset"></i> Soporte</a>{% if session.get('rol')=='admin' %}<a class="text-white text-decoration-none" href="{{url_for('configuraciones')}}"><i class="bi bi-gear"></i> Config.</a>{% else %}<span></span>{% endif %}</div>
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
    <div class="phone-wrap desktop-pad tareo-list-page">
      <div class="page-card">
        <div class="green-hero tareo-hero" style="border-radius:0 0 12px 12px">
          <div class="green-top"><span>v.1.0</span><a class="text-white text-decoration-none" href="{{url_for('home')}}"><i class="bi bi-house"></i></a></div>
          <i class="bi bi-list-check" style="font-size:34px;margin-top:14px"></i><div class="login-name mt-1">TAREOS</div>
        </div>
        <div class="toolstrip tareo-toolbar">
          <a class="btn-plus-fab" title="Crear hoja" href="{{url_for('crear_hoja')}}"><i class="bi bi-list-task"></i><i class="bi bi-plus-circle-fill"></i></a>
          <a title="Plantilla Excel" href="{{url_for('plantilla_trabajadores')}}"><i class="bi bi-file-earmark-excel"></i></a>
          <a title="Sincronizar" href="{{url_for('sincronizacion')}}"><i class="bi bi-arrow-clockwise"></i></a>
        </div>
        {% for h in hojas %}
          <a class="text-decoration-none" href="{{url_for('detalle_hoja', hoja_id=h.id)}}"><div class="worker-card">
            <span class="person-dot" style="border-radius:8px"><i class="bi bi-clipboard2-check"></i></span>
            <div class="worker-title"><div>RESPONSABLE<br><b>{{h.responsable}}</b></div><div class="text-end">PRESUPUESTO<br><b>{{h.tipo_tareo or 'JORNAL'}}</b></div></div>
            <div class="worker-grid"><div><label>SUCURSAL</label><div class="small-value">{{h.grupo}}</div></div><div><label>PLANILLA</label><div class="small-value">AGR. PACKING</div></div><div><label>DOCUMENTO</label><div class="small-value">{{h.id}}</div></div></div>
            <div class="small-label mt-2">ZONA CONSUMIDOR</div><div class="small-value">{{h.subgrupo}}</div>
            <div class="worker-grid"><div><label>FECHA</label><div class="small-value">{{h.fecha}}</div></div><div><label>ESTADO</label><div class="mini-badge bg-y">{{h.estado}}</div></div><div class="text-end"><i class="bi bi-chevron-left text-success"></i></div></div>
          </div></a>
        {% else %}<div class="worker-card text-center text-muted">No hay hojas. Presiona <b>+</b> para crear.</div>{% endfor %}
        <div class="leaf"></div>
        <div class="bottom-nav"><a href="{{url_for('hojas_tareo')}}"><i class="bi bi-list-check"></i>Listado de Tareos</a><a href="{{url_for('home')}}"><i class="bi bi-file-text"></i>Detalle</a></div>
      </div>
    </div>"""
    return render_page(body, hojas=hojas)

@app.route('/hojas/crear', methods=['GET','POST'])
@login_required
def crear_hoja():
    if request.method == 'POST':
        fecha = request.form.get('fecha') or today_str()
        grupo = limpiar_texto(request.form.get('actividad') or request.form.get('grupo'))
        subgrupo = limpiar_texto(request.form.get('labor') or request.form.get('subgrupo'))
        labor = limpiar_texto(request.form.get('consumidor') or request.form.get('consumidor_desc') or request.form.get('labor_consumidor') or '')
        responsable = limpiar_texto(request.form.get('responsable'))
        turno = limpiar_texto(request.form.get('turno') or 'DIA')
        tipo_tareo = limpiar_texto(request.form.get('tipo_tareo') or 'JORNAL')
        if turno not in ('DIA','NOCHE'): turno = 'DIA'
        if tipo_tareo not in ('JORNAL','RENDIMIENTO'): tipo_tareo = 'JORNAL'
        if not grupo or not subgrupo or not responsable:
            flash('Debe seleccionar Actividad y Labor, además de responsable.', 'danger')
            return redirect(url_for('crear_hoja'))
        execute('INSERT INTO hojas_tareo(fecha,grupo,subgrupo,labor,responsable,turno,tipo_tareo,estado,creado_por,creado_en) VALUES(?,?,?,?,?,?,?,?,?,?)',
                (fecha,grupo,subgrupo,labor,responsable,turno,tipo_tareo,'ABIERTA',session.get('usuario'),now_str()), commit=True)
        hid = scalar('SELECT MAX(id) AS id FROM hojas_tareo')
        execute('INSERT INTO hoja_labores(hoja_id,grupo,subgrupo,labor,turno,tipo_tareo,responsable,creado_en,creado_por) VALUES(?,?,?,?,?,?,?,?,?)',
                (hid,grupo,subgrupo,labor,turno,tipo_tareo,responsable,now_str(),session.get('usuario')), commit=True)
        return redirect(url_for('hojas_tareo'))
    maestros = get_actividades_maestras()
    body = """
    <div class="phone-wrap desktop-pad"><h2 class="header-title">CREAR HOJA DE TAREO</h2><div class="page-card">
      <div class="panel-green"><i class="bi bi-clipboard2-plus"></i><h4>NUEVA HOJA – FECHA, ACTIVIDAD, LABOR, CONSUMIDOR, RESPONSABLE, TURNO Y TIPO</h4></div>
      <form method="post" class="floating-card" style="margin:-24px 10px 12px">
        <label class="form-label">FECHA</label><input type="date" name="fecha" class="form-control mb-2 field-required" value="{{today}}" required>
        <label class="form-label">ACTIVIDAD</label><input id="actividadInput" name="actividad" class="form-control mb-2 field-required" list="actividad_list" placeholder="Digite primeras letras de la actividad" required autocomplete="off"><datalist id="actividad_list"></datalist>
        <label class="form-label">LABOR</label><input id="laborInput" name="labor" class="form-control mb-2 field-required" list="labor_list" placeholder="Seleccione labor según actividad" required autocomplete="off"><datalist id="labor_list"></datalist>
        <label class="form-label">CONSUMIDOR <span class="text-muted">(opcional)</span></label><input id="consumidorInput" name="consumidor" class="form-control mb-2" list="consumidor_list" placeholder="Consumidor / zona / campo"><datalist id="consumidor_list"></datalist>
        <label class="form-label">RESPONSABLE</label><input name="responsable" class="form-control mb-2 field-required" placeholder="APELLIDOS Y NOMBRES" required>
        <div class="row g-2 mb-3"><div class="col-6"><label class="form-label">TURNO</label><select name="turno" class="form-select field-required"><option>DIA</option><option>NOCHE</option></select></div><div class="col-6"><label class="form-label">TIPO</label><select name="tipo_tareo" class="form-select field-required"><option>JORNAL</option><option>RENDIMIENTO</option></select></div></div>
        <button class="btn btn-green w-100"><i class="bi bi-check-circle"></i> GUARDAR HOJA</button><a class="btn btn-outline-secondary w-100 mt-2" href="{{url_for('hojas_tareo')}}">VOLVER</a>
      </form></div></div>
      <script>
        const MAESTROS={{ maestros_json|safe }};
        const uniq=a=>[...new Set(a.filter(Boolean))].sort();
        function fillList(id, arr){const dl=document.getElementById(id); dl.innerHTML=''; arr.slice(0,300).forEach(v=>{const o=document.createElement('option'); o.value=v; dl.appendChild(o);});}
        function refreshActividad(){fillList('actividad_list', uniq(MAESTROS.map(x=>x.desc_actividad || x.cod_actividad)));}
        function refreshLabor(){const a=(actividadInput.value||'').toUpperCase(); const rows=MAESTROS.filter(x=>!a || String(x.desc_actividad||'').toUpperCase()===a || String(x.cod_actividad||'').toUpperCase()===a || String(x.desc_actividad||'').toUpperCase().includes(a)); fillList('labor_list', uniq(rows.map(x=>x.desc_labor || x.cod_labor))); refreshConsumidor();}
        function refreshConsumidor(){const a=(actividadInput.value||'').toUpperCase(), l=(laborInput.value||'').toUpperCase(); const rows=MAESTROS.filter(x=>(!a || String(x.desc_actividad||'').toUpperCase().includes(a)||String(x.cod_actividad||'').toUpperCase().includes(a)) && (!l || String(x.desc_labor||'').toUpperCase().includes(l)||String(x.cod_labor||'').toUpperCase().includes(l))); fillList('consumidor_list', uniq(rows.map(x=>x.desc_consumidor || x.cod_consumidor)));}
        actividadInput.addEventListener('input', refreshLabor); laborInput.addEventListener('input', refreshConsumidor); document.addEventListener('DOMContentLoaded',()=>{refreshActividad();refreshLabor();});
      </script>
    """
    return render_page(body, today=today_str(), maestros_json=js_master_options(maestros))

@app.route('/hoja/<int:hoja_id>')
@login_required
def detalle_hoja(hoja_id):
    tab = request.args.get('tab','labores')
    h = row_to_dict(execute('SELECT * FROM hojas_tareo WHERE id=?', (hoja_id,), fetchone=True))
    if not h:
        flash('Hoja no encontrada.', 'danger'); return redirect(url_for('hojas_tareo'))
    labores = rows_to_dict(execute('SELECT * FROM hoja_labores WHERE hoja_id=? ORDER BY id DESC', (hoja_id,), fetchall=True))
    tareos = rows_to_dict(execute('SELECT * FROM tareos WHERE hoja_id=? ORDER BY creado_en DESC LIMIT 100', (hoja_id,), fetchall=True))
    lecturas = rows_to_dict(execute('SELECT * FROM lecturas_balde WHERE hoja_id=? ORDER BY fecha_hora DESC LIMIT 100', (hoja_id,), fetchall=True))
    registros = len(tareos); horas_total = sum(float(x.get('horas') or 0) for x in tareos); rend_total = sum(float(x.get('cantidad') or 0) for x in tareos)
    execute('UPDATE hojas_tareo SET registros=?, horas_total=?, rendimiento_total=? WHERE id=?', (registros, horas_total, rend_total, hoja_id), commit=True)
    body = """
    <div class="phone-wrap desktop-pad"><h2 class="header-title">TAREO MÓVIL – {{ 'DETALLE DE TRABAJADOR POR LABOR' if tab=='trabajadores' else ('DETALLE NÚMERO DE LECTURAS POR BALDE' if tab=='rendimiento' else 'GRUPO DE COSECHA') }}</h2>
      <div class="page-card">
        <div class="tab-main"><a class="{{'active' if tab=='labores' else ''}}" href="{{url_for('detalle_hoja',hoja_id=h.id,tab='labores')}}">LABORES</a><a class="{{'active' if tab=='trabajadores' else ''}}" href="{{url_for('detalle_hoja',hoja_id=h.id,tab='trabajadores')}}">TRABAJADORES</a><a class="{{'active' if tab=='rendimiento' else ''}}" href="{{url_for('detalle_hoja',hoja_id=h.id,tab='rendimiento')}}">REND./AVANCE</a></div>
        <div class="subtabs"><a class="{{'active' if tab=='labores' else ''}}" href="{{url_for('detalle_hoja',hoja_id=h.id,tab='labores')}}">Labores</a><a class="{{'active' if tab=='trabajadores' else ''}}" href="{{url_for('detalle_hoja',hoja_id=h.id,tab='trabajadores')}}">Trab.por Labor</a><a class="{{'active' if tab=='rendimiento' else ''}}" href="{{url_for('detalle_hoja',hoja_id=h.id,tab='rendimiento')}}">Rend/Avance por Labor</a></div>
        <div class="panel-green"><i class="bi {{ 'bi-people-fill' if tab=='trabajadores' else ('bi-person-badge' if tab=='rendimiento' else 'bi-people') }}"></i><h4>{{ 'TRABAJADORES – QR / CÓDIGO BARRAS / DIGITACIÓN' if tab=='trabajadores' else ('PRODUCTIVIDAD POR TRABAJADOR – BALDE / QR / DIGITACIÓN' if tab=='rendimiento' else 'REGISTRO DE ACTIVIDAD, LABOR, CONSUMIDOR, TURNO Y TIPO') }}</h4></div>
        <div class="toolstrip">
          <button title="Crear labor/grupo/subgrupo" data-bs-toggle="modal" data-bs-target="#modalLabor"><i class="bi bi-list-check"></i></button>
          <button title="Buscar trabajador" data-bs-toggle="modal" data-bs-target="#modalBuscar"><i class="bi bi-search"></i></button>
          {% if tab=='trabajadores' %}
            <button title="Elegir horarios" data-bs-toggle="modal" data-bs-target="#modalHora"><i class="bi bi-clock"></i></button>
            <button title="Entrada / salida" data-bs-toggle="modal" data-bs-target="#modalHora"><i class="bi bi-box-arrow-in-right"></i></button>
            <button title="Registrar trabajador" data-bs-toggle="modal" data-bs-target="#modalRegistro"><i class="bi bi-person-plus"></i></button>
          {% elif tab=='rendimiento' %}
            <button title="Registrar avance" data-bs-toggle="modal" data-bs-target="#modalAvance"><i class="bi bi-upc-scan"></i></button>
            <button title="Refrescar" onclick="location.reload()"><i class="bi bi-arrow-clockwise"></i></button>
          {% else %}
            <button title="Copiar labor" data-bs-toggle="modal" data-bs-target="#modalCopiar"><i class="bi bi-files"></i></button>
            <button title="Refrescar" onclick="location.reload()"><i class="bi bi-arrow-clockwise"></i></button>
          {% endif %}
          <a href="{{url_for('hojas_tareo')}}" title="Volver"><i class="bi bi-chevron-left"></i></a>
        </div>
        <div class="info-bar"><div><i class="bi bi-calendar"></i> {{h.fecha}}</div><div><i class="bi bi-list"></i> {{registros}} Reg.</div><div><i class="bi bi-clock"></i> {{'%.2f'|format(horas_total)}} H.</div><div>A.Rend {{'%.2f'|format(rend_total)}}</div><span>⌄</span></div>
        {% if tab=='labores' %}
          {% for l in labores %}<a class="text-decoration-none" href="{{url_for('detalle_hoja',hoja_id=h.id,tab='trabajadores')}}"><div class="worker-card labor-card-compact"><div class="worker-title"><div>ACTIVIDAD<br><b>{{l.grupo}}</b></div><div class="text-end">LABOR<br><b>{{l.subgrupo or 'SIN LABOR'}}</b></div></div><div class="mt-2"><span class="small-label">CONSUMIDOR</span> <b class="labor-main">{{l.labor or 'SIN CONSUMIDOR'}}</b><br><span class="small-label">RESPONSABLE</span> <b class="resp-main">{{l.responsable or h.responsable}}</b></div><div class="worker-grid mt-2"><div><div class="mini-badge {{'bg-y' if l.turno=='NOCHE' else 'bg-g'}}">{{l.turno}}</div></div><div><div class="mini-badge bg-y">{{l.tipo_tareo}}</div></div><div><div class="mini-badge bg-g">ACTIVA</div></div></div></div></a>{% else %}<div class="worker-card text-center text-muted">Presiona <b>+</b> para crear actividad, labor y consumidor.</div>{% endfor %}
        {% elif tab=='trabajadores' %}
          {% for r in tareos %}<div class="worker-card"><div class="worker-title"><div>TRABAJADOR<br><b>{{r.trabajador}}</b></div><div>NRO.DOCUMENTO<br><b>{{r.dni}}</b></div></div><div class="worker-grid"><div><label>HORA INICIO</label><input class="mini-input" value="{{r.hora_inicio or ('22:00' if r.turno=='NOCHE' else '06:30')}}"></div><div><label>HORA FIN</label><input class="mini-input" value="{{r.hora_fin or ('06:00' if r.turno=='NOCHE' else '16:30')}}"></div><div><label>H.NORMAL</label><input class="mini-input" value="{{'%.2f'|format(r.horas or 0)}}"></div><div><label>REF. INI</label><input class="mini-input" value="{{r.ref_inicio or '12:00'}}"></div><div><label>REF. FIN</label><input class="mini-input" value="{{r.ref_fin or '13:00'}}"></div><div><label>ESTADO</label><div class="mini-badge bg-g">FIN TOTAL</div></div></div></div>{% else %}<div class="worker-card text-center text-muted">Presiona el <b>hombresito +</b> para registrar trabajador por QR/código/digitación.</div>{% endfor %}
        {% else %}
          {% for l in lecturas %}<div class="worker-card"><span class="person-dot"><i class="bi bi-person-circle"></i></span><div class="worker-title"><div>TRABAJADOR<br><b>{{l.trabajador}}</b></div><div>NRO.DOC.<br><b>{{l.dni}}</b></div></div><div class="small-label mt-1">HORA TOMA REGISTRO</div><div class="small-value">{{l.fecha_hora}} · {{l.metodo or 'DIGITACIÓN'}}</div><div class="worker-grid"><div><label>A.DIURNO</label><div class="mini-badge bg-y">{{'%.2f'|format(l.a_diurno or 0)}}</div></div><div><label>A.NOCT.</label><div class="mini-badge bg-y">{{'%.2f'|format(l.a_noct or 0)}}</div></div><div class="text-end"><i class="bi bi-chevron-left text-success"></i></div></div></div>{% else %}<div class="worker-card text-center text-muted">Presiona el icono de escaneo para registrar avance por QR/código/digitación.</div>{% endfor %}
        {% endif %}<div class="leaf"></div>
      </div>
    </div>

    <div class="modal fade" id="modalLabor" tabindex="-1"><div class="modal-dialog modal-dialog-centered"><div class="modal-content"><form method="post" action="{{url_for('guardar_labor_hoja', hoja_id=h.id, tab=tab)}}"><div class="modal-header"><h5 class="modal-title fw-bold text-success"><i class="bi bi-plus-square"></i> Crear nueva labor</h5><button class="btn-close" data-bs-dismiss="modal"></button></div><div class="modal-body"><div class="alert alert-light border small mb-2">Complete los datos y presione <b>CREAR LABOR</b>. Al cerrar con X no se guarda nada.</div><label class="form-label">ACTIVIDAD</label><input name="grupo" class="form-control mb-2" placeholder="Digite actividad" required autocomplete="off"><label class="form-label">LABOR</label><input name="subgrupo" class="form-control mb-2" placeholder="Digite labor" required autocomplete="off"><label class="form-label">CONSUMIDOR (opcional)</label><input name="labor" class="form-control mb-2" placeholder="Consumidor / zona / campo"><label class="form-label">RESPONSABLE</label><input name="responsable" class="form-control mb-2" placeholder="APELLIDOS Y NOMBRES"><div class="row g-2"><div class="col-6"><label class="form-label">TURNO</label><select name="turno" class="form-select"><option>DIA</option><option>NOCHE</option></select></div><div class="col-6"><label class="form-label">TIPO</label><select name="tipo_tareo" class="form-select"><option>JORNAL</option><option>RENDIMIENTO</option></select></div></div></div><div class="modal-footer"><button class="btn btn-green w-100" type="submit">CREAR LABOR</button></div></form></div></div></div>
    <div class="modal fade" id="modalCopiar" tabindex="-1"><div class="modal-dialog modal-dialog-centered"><div class="modal-content"><form method="post" action="{{url_for('copiar_labor_hoja', hoja_id=h.id, tab=tab)}}"><div class="modal-header"><h5 class="modal-title fw-bold text-success"><i class="bi bi-files"></i> Copiar labor existente</h5><button class="btn-close" data-bs-dismiss="modal"></button></div><div class="modal-body"><div class="alert alert-light border small">Selecciona el documento/labor que deseas copiar. No se copiará nada hasta presionar <b>COPIAR SELECCIONADO</b>.</div><div class="copy-list">{% for l in labores %}<label class="d-block mb-2"><input type="radio" name="labor_id_origen" value="{{l.id}}" required> <b>{{l.labor}}</b><br><span class="small text-muted">{{l.grupo}} / {{l.subgrupo}} / {{l.turno}} / {{l.tipo_tareo}}</span></label>{% endfor %}</div><label class="form-label mt-2">Nuevo nombre de labor (opcional)</label><input name="labor_nueva" class="form-control" placeholder="Dejar vacío para copiar igual"></div><div class="modal-footer"><button class="btn btn-green w-100">COPIAR SELECCIONADO</button></div></form></div></div></div>
    <div class="modal fade" id="modalBuscar" tabindex="-1"><div class="modal-dialog modal-dialog-centered"><div class="modal-content"><div class="modal-header"><h5 class="modal-title fw-bold text-success"><i class="bi bi-search"></i> Buscar trabajador</h5><button class="btn-close" data-bs-dismiss="modal"></button></div><div class="modal-body"><input id="buscarDni" class="form-control mb-2" placeholder="DNI / QR / código barras"><button class="btn btn-green w-100" onclick="buscarTrabajadorLibre()">BUSCAR</button><div id="buscarResultado" class="alert alert-light border mt-2">Esperando búsqueda.</div></div></div></div></div>
    <div class="modal fade" id="modalHora" tabindex="-1"><div class="modal-dialog modal-dialog-centered"><div class="modal-content"><div class="modal-header"><h5 class="modal-title fw-bold text-success"><i class="bi bi-clock"></i> Horarios de trabajo y refrigerio</h5><button class="btn-close" data-bs-dismiss="modal"></button></div><div class="modal-body"><div class="clock-face"><span class="clock-num" style="top:8px;left:86px">0</span><span class="clock-num" style="top:25px;right:45px">5</span><span class="clock-bubble">10</span><span class="clock-num" style="top:88px;right:20px">15</span><span class="clock-num" style="bottom:45px;right:38px">20</span><span class="clock-num" style="bottom:18px;left:86px">30</span><span class="clock-num" style="bottom:45px;left:38px">35</span><span class="clock-num" style="top:88px;left:20px">40</span><span class="clock-num" style="top:55px;left:28px">50</span><span class="clock-num" style="top:28px;left:55px">55</span><span class="clock-hand"></span><span class="clock-dot"></span></div><div class="alert alert-success small"><b>Turno NOCHE:</b> recomendado 22:00 a 06:00. Puedes modificarlo según tu operación.</div><div class="row g-2"><div class="col-6"><label class="form-label">Inicio trabajo</label><input id="horaInicioDefault" type="time" class="form-control" value="06:30"></div><div class="col-6"><label class="form-label">Fin trabajo</label><input id="horaFinDefault" type="time" class="form-control" value="16:30"></div><div class="col-6"><label class="form-label">Inicio refrigerio</label><input id="refInicioDefault" type="time" class="form-control" value="12:00"></div><div class="col-6"><label class="form-label">Fin refrigerio</label><input id="refFinDefault" type="time" class="form-control" value="13:00"></div></div><button class="btn btn-green w-100 mt-3" type="button" onclick="aplicarHorarioRegistro()" data-bs-dismiss="modal">APLICAR AL REGISTRO</button></div></div></div></div>
    <div class="modal fade" id="modalRegistro" tabindex="-1"><div class="modal-dialog modal-dialog-centered"><div class="modal-content"><form method="post" action="{{url_for('guardar_registro_hoja', hoja_id=h.id, tab='trabajadores')}}" id="frmTrab"><div class="modal-header"><h5 class="modal-title fw-bold text-success"><i class="bi bi-person-plus"></i> Registrar trabajador</h5><button class="btn-close" data-bs-dismiss="modal"></button></div><div class="modal-body"><div class="scan-box mb-2"><label class="form-label">DNI / QR / CÓDIGO BARRAS</label><div class="input-group"><input name="dni" id="dniTrab" class="form-control" placeholder="Escanee o digite DNI" autocomplete="off"><button type="button" class="btn btn-green" onclick="abrirScanner('readerTrab','dniTrab')"><i class="bi bi-upc-scan"></i></button></div><div id="readerTrab" style="display:none;margin-top:8px"></div><div id="dniStatus" class="mt-2 field-help">Digite/escanee DNI. Se reconocerá automáticamente.</div><input type="hidden" name="dnis_masivos" id="dnisMasivos"><div id="workerQueue" class="worker-queue"><div class="text-muted small text-center">Trabajadores detectados aparecerán aquí.</div></div></div><label class="form-label">LABOR</label><select name="labor_id" class="form-select mb-2">{% for l in labores %}<option value="{{l.id}}">{{l.grupo}} / {{l.subgrupo}} / {{l.labor}} / {{l.turno}} / {{l.tipo_tareo}}</option>{% endfor %}</select><div class="row g-2"><div class="col-6"><label class="form-label">TURNO</label><select name="turno" id="turnoTrab" class="form-select" onchange="setTurnoHorario()"><option>DIA</option><option>NOCHE</option></select></div><div class="col-6"><label class="form-label">TIPO</label><select name="tipo_tareo" class="form-select"><option>JORNAL</option><option>RENDIMIENTO</option></select></div><div class="col-6"><label class="form-label">H. INICIO</label><input name="hora_inicio" id="horaInicioTrab" type="time" class="form-control time-click" value="06:30" onclick="abrirRelojPara(this)"></div><div class="col-6"><label class="form-label">H. FIN</label><input name="hora_fin" id="horaFinTrab" type="time" class="form-control time-click" value="16:30" onclick="abrirRelojPara(this)"></div><div class="col-6"><label class="form-label">REF. INI</label><input name="ref_inicio" id="refInicioTrab" type="time" class="form-control time-click" value="12:00" onclick="abrirRelojPara(this)"></div><div class="col-6"><label class="form-label">REF. FIN</label><input name="ref_fin" id="refFinTrab" type="time" class="form-control time-click" value="13:00" onclick="abrirRelojPara(this)"></div></div><input name="horas" type="hidden" value="9.75"><input name="cantidad" type="hidden" value="0.00"><div class="scan-ok mt-3">Pre-registro activo: los DNI detectados quedan en la lista superior hasta presionar guardar.</div></div><div class="modal-footer"><button class="btn btn-green w-100">GUARDAR PRE-REGISTRO / TODOS</button></div></form></div></div></div>
    <div class="modal fade" id="modalAvance" tabindex="-1"><div class="modal-dialog modal-dialog-centered"><div class="modal-content"><form method="post" action="{{url_for('guardar_registro_hoja', hoja_id=h.id, tab='rendimiento')}}"><div class="modal-header"><h5 class="modal-title fw-bold text-success"><i class="bi bi-upc-scan"></i> Registrar avance / lectura</h5><button class="btn-close" data-bs-dismiss="modal"></button></div><div class="modal-body"><div class="scan-box mb-2"><label class="form-label">DNI / QR / CÓDIGO BARRAS</label><div class="input-group"><input name="dni" id="dniAvance" class="form-control" placeholder="Escanee o digite DNI" required><button type="button" class="btn btn-green" onclick="abrirScanner('readerAvance','dniAvance')"><i class="bi bi-upc-scan"></i></button></div><div id="readerAvance" style="display:none;margin-top:8px"></div></div><label class="form-label">LABOR</label><select name="labor_id" class="form-select mb-2">{% for l in labores %}<option value="{{l.id}}">{{l.labor}} / {{l.turno}} / {{l.tipo_tareo}}</option>{% endfor %}</select><div class="row g-2"><div class="col-6"><label class="form-label">A. DIURNO</label><input name="cantidad" type="number" step="0.01" class="form-control" value="1.00"></div><div class="col-6"><label class="form-label">A. NOCT.</label><input name="a_noct" type="number" step="0.01" class="form-control" value="0.00"></div><div class="col-6"><label class="form-label">UNIDAD</label><select name="unidad" class="form-select"><option>BALDE</option><option>KG</option><option>JABA</option><option>UNIDAD</option></select></div><div class="col-6"><label class="form-label">MÉTODO</label><select name="metodo" class="form-select"><option>QR/CÓDIGO</option><option>DIGITACIÓN</option><option>LECTOR USB</option></select></div></div></div><div class="modal-footer"><button class="btn btn-green w-100">GUARDAR AVANCE</button></div></form></div></div></div>
    <script>
      function limpiarDni(v){let raw=(v||'').toString();let m=raw.match(/(?:^|\D)(\d{8})(?:\D|$)/);let d=m?m[1]:raw.replace(/\D/g,'');return d.length>=8?d.slice(-8):d;}
      async function buscarTrabajadorLibre(){let dni=limpiarDni(document.getElementById('buscarDni').value);document.getElementById('buscarDni').value=dni;let box=document.getElementById('buscarResultado');if(dni.length!==8){box.className='alert alert-warning mt-2';box.innerText='Ingrese DNI válido de 8 dígitos.';return;}let r=await fetch('/api/trabajador/'+dni);let j=await r.json();if(!j.ok){box.className='alert alert-danger mt-2';box.innerText=j.msg;return;}box.className='alert alert-success mt-2';box.innerHTML='<b>'+j.trabajador.trabajador+'</b><br>'+j.trabajador.dni+' · '+(j.trabajador.cargo||'');beep();}
      let scanner=null;function abrirScanner(readerId,inputId){let el=document.getElementById(readerId);el.style.display='block';if(scanner){scanner.stop().catch(()=>{});scanner=null;}scanner=new Html5Qrcode(readerId);scanner.start({facingMode:'environment'},{fps:10,qrbox:220},decoded=>{document.getElementById(inputId).value=limpiarDni(decoded);document.getElementById(inputId).dispatchEvent(new Event('input'));beep();scanner.stop().catch(()=>{});el.style.display='none';}).catch(()=>alert('No se pudo activar cámara. Revise permisos.'));}
      function setTurnoHorario(){let t=document.getElementById('turnoTrab')?.value;if(t==='NOCHE'){horaInicioTrab.value='22:00';horaFinTrab.value='06:00';}else{horaInicioTrab.value='06:30';horaFinTrab.value='16:30';}}
      let activeTimeInput=null;
      function abrirRelojPara(inp){activeTimeInput=inp;['horaInicioDefault','horaFinDefault','refInicioDefault','refFinDefault'].forEach((id,i)=>{let el=document.getElementById(id); if(!el)return; const src=[horaInicioTrab,horaFinTrab,refInicioTrab,refFinTrab][i]; el.value=src.value;}); const m=new bootstrap.Modal(document.getElementById('modalHora')); m.show();}
      function aplicarHorarioRegistro(){
        if(document.getElementById('horaInicioTrab')){horaInicioTrab.value=horaInicioDefault.value;horaFinTrab.value=horaFinDefault.value;refInicioTrab.value=refInicioDefault.value;refFinTrab.value=refFinDefault.value;}
        beep();
      }
      function bindClock(){
        document.querySelectorAll('#modalHora input[type=time]').forEach(inp=>{inp.addEventListener('focus',()=>activeTimeInput=inp);inp.addEventListener('click',()=>activeTimeInput=inp);});
        activeTimeInput=document.getElementById('horaInicioDefault');
        const face=document.querySelector('#modalHora .clock-face'); if(!face)return;
        const hand=face.querySelector('.clock-hand'), bubble=face.querySelector('.clock-bubble');
        const setFromEvent=(ev)=>{
          const r=face.getBoundingClientRect(); const touch=ev.touches?ev.touches[0]:ev;
          const cx=r.left+r.width/2, cy=r.top+r.height/2; const x=touch.clientX-cx, y=touch.clientY-cy;
          let deg=Math.atan2(y,x)*180/Math.PI + 90; if(deg<0)deg+=360;
          let minute=Math.round(deg/6/5)*5; if(minute>=60)minute=0;
          const inp=activeTimeInput||document.getElementById('horaInicioDefault');
          let [hh]=String(inp.value||'00:00').split(':'); inp.value=String(hh).padStart(2,'0')+':'+String(minute).padStart(2,'0');
          if(hand) hand.style.transform='rotate('+(deg-90)+'deg)';
          if(bubble) bubble.textContent=String(minute).padStart(2,'0');
          const map={horaInicioDefault:'horaInicioTrab',horaFinDefault:'horaFinTrab',refInicioDefault:'refInicioTrab',refFinDefault:'refFinTrab'};
          if(map[inp.id] && document.getElementById(map[inp.id])) document.getElementById(map[inp.id]).value=inp.value;
          ev.preventDefault();
        };
        let drag=false; face.addEventListener('mousedown',e=>{drag=true;setFromEvent(e)}); window.addEventListener('mousemove',e=>{if(drag)setFromEvent(e)}); window.addEventListener('mouseup',()=>drag=false);
        face.addEventListener('touchstart',e=>{drag=true;setFromEvent(e)},{passive:false}); face.addEventListener('touchmove',e=>{if(drag)setFromEvent(e)},{passive:false}); face.addEventListener('touchend',()=>drag=false);
      }
      const workerMap=new Map();
      function renderQueue(){const q=document.getElementById('workerQueue'), h=document.getElementById('dnisMasivos'); if(!q||!h)return; h.value=[...workerMap.keys()].join(','); if(workerMap.size===0){q.innerHTML='<div class="text-muted small text-center">Trabajadores detectados aparecerán aquí.</div>';return;} q.innerHTML=[...workerMap.entries()].map(([dni,n])=>'<div class="queue-item"><div><b>'+dni+'</b><br><span>'+n+'</span></div><button type="button" class="btn btn-sm btn-outline-danger" onclick="workerMap.delete(\''+dni+'\');renderQueue()">×</button></div>').join('');}
      async function detectarDniTrab(){const inp=document.getElementById('dniTrab'), st=document.getElementById('dniStatus'); if(!inp||!st)return; const dni=limpiarDni(inp.value); if(dni.length!==8)return; inp.value=dni; let r=await fetch('/api/trabajador/'+dni); let j=await r.json(); if(!j.ok){st.className='scan-bad mt-2';st.textContent=j.msg;return;} workerMap.set(dni,j.trabajador.trabajador||'TRABAJADOR'); renderQueue(); st.className='scan-ok mt-2'; st.innerHTML='✓ Reconocido: <b>'+j.trabajador.trabajador+'</b>'; beep(); setTimeout(()=>{inp.value='';inp.focus();},250);}
      function bindDniAuto(){const inp=document.getElementById('dniTrab'); if(!inp)return; let t=null; inp.addEventListener('input',()=>{clearTimeout(t); if(limpiarDni(inp.value).length>=8)t=setTimeout(detectarDniTrab,120);}); inp.addEventListener('keydown',e=>{if(e.key==='Enter'){e.preventDefault();detectarDniTrab();}});}
      
      // Mejora reloj v7: arrastrar/click mueve la manija y actualiza el campo activo en tiempo real.
      function bindClockV7(){
        document.querySelectorAll('#modalHora input[type=time]').forEach(inp=>{
          inp.addEventListener('focus',()=>activeTimeInput=inp);
          inp.addEventListener('click',()=>activeTimeInput=inp);
        });
        const face=document.querySelector('#modalHora .clock-face'); if(!face || face.dataset.v7)return; face.dataset.v7='1';
        const hand=face.querySelector('.clock-hand'), bubble=face.querySelector('.clock-bubble');
        function updateFromPointer(ev){
          const r=face.getBoundingClientRect(); const e=ev.touches?ev.touches[0]:ev;
          const cx=r.left+r.width/2, cy=r.top+r.height/2;
          const dx=e.clientX-cx, dy=e.clientY-cy;
          let deg=Math.atan2(dy,dx)*180/Math.PI + 90; if(deg<0)deg+=360;
          let minute=Math.round((deg/360)*60/5)*5; if(minute>=60)minute=0;
          const inp=activeTimeInput || document.getElementById('horaInicioDefault');
          let [hh]=String(inp.value||'00:00').split(':');
          inp.value=String(hh).padStart(2,'0')+':'+String(minute).padStart(2,'0');
          if(hand) hand.style.transform='rotate('+(deg-90)+'deg)';
          if(bubble) bubble.textContent=String(minute).padStart(2,'0');
          const map={horaInicioDefault:'horaInicioTrab',horaFinDefault:'horaFinTrab',refInicioDefault:'refInicioTrab',refFinDefault:'refFinTrab'};
          if(map[inp.id] && document.getElementById(map[inp.id])) document.getElementById(map[inp.id]).value=inp.value;
          ev.preventDefault();
        }
        let dragging=false;
        face.addEventListener('pointerdown',e=>{dragging=true; face.setPointerCapture && face.setPointerCapture(e.pointerId); updateFromPointer(e);});
        face.addEventListener('pointermove',e=>{if(dragging)updateFromPointer(e);});
        face.addEventListener('pointerup',()=>dragging=false);
        face.addEventListener('pointercancel',()=>dragging=false);
      }
      const _oldAbrirRelojPara = abrirRelojPara;
      abrirRelojPara = function(inp){ activeTimeInput=inp; _oldAbrirRelojPara(inp); setTimeout(bindClockV7,150); };

      document.addEventListener('DOMContentLoaded',()=>{bindClock();bindClockV7();bindDniAuto();});
    </script>
    """
    return render_page(body, h=h, tab=tab, tareos=tareos, lecturas=lecturas, labores=labores, registros=registros, horas_total=horas_total, rend_total=rend_total)


@app.route('/hoja/<int:hoja_id>/copiar-labor/<tab>', methods=['POST'])
@login_required
def copiar_labor_hoja(hoja_id, tab):
    origen_id = request.form.get('labor_id_origen')
    lab = row_to_dict(execute('SELECT * FROM hoja_labores WHERE id=? AND hoja_id=?', (origen_id, hoja_id), fetchone=True))
    if not lab:
        flash('Selecciona una labor válida para copiar.', 'danger')
        return redirect(url_for('detalle_hoja', hoja_id=hoja_id, tab='labores'))
    labor_nueva = limpiar_texto(request.form.get('labor_nueva') or lab.get('labor'))
    execute('INSERT INTO hoja_labores(hoja_id,grupo,subgrupo,labor,turno,tipo_tareo,responsable,creado_en,creado_por) VALUES(?,?,?,?,?,?,?,?,?)',
            (hoja_id, lab.get('grupo'), lab.get('subgrupo'), labor_nueva, lab.get('turno'), lab.get('tipo_tareo'), lab.get('responsable'), now_str(), session.get('usuario')), commit=True)
    flash('Labor copiada correctamente.', 'success')
    return redirect(url_for('detalle_hoja', hoja_id=hoja_id, tab='labores'))

@app.route('/hoja/<int:hoja_id>/labor/<tab>', methods=['POST'])
@login_required
def guardar_labor_hoja(hoja_id, tab):
    h = row_to_dict(execute('SELECT * FROM hojas_tareo WHERE id=?', (hoja_id,), fetchone=True))
    if not h: flash('Hoja no encontrada.', 'danger'); return redirect(url_for('hojas_tareo'))
    grupo = limpiar_texto(request.form.get('grupo') or h.get('grupo'))
    subgrupo = limpiar_texto(request.form.get('subgrupo') or h.get('subgrupo'))
    labor = limpiar_texto(request.form.get('labor') or h.get('labor'))
    responsable = limpiar_texto(request.form.get('responsable') or h.get('responsable'))
    turno = limpiar_texto(request.form.get('turno') or 'DIA')
    tipo_tareo = limpiar_texto(request.form.get('tipo_tareo') or 'JORNAL')
    if not grupo or not subgrupo:
        flash('Debe seleccionar Actividad y Labor.', 'danger')
        return redirect(url_for('detalle_hoja', hoja_id=hoja_id, tab='labores'))
    execute('INSERT INTO hoja_labores(hoja_id,grupo,subgrupo,labor,turno,tipo_tareo,responsable,creado_en,creado_por) VALUES(?,?,?,?,?,?,?,?,?)',
            (hoja_id,grupo,subgrupo,labor,turno,tipo_tareo,responsable,now_str(),session.get('usuario')), commit=True)
    flash('Nueva labor creada dentro de la hoja.', 'success')
    return redirect(url_for('detalle_hoja', hoja_id=hoja_id, tab='labores'))

@app.route('/hoja/<int:hoja_id>/registro/<tab>', methods=['POST'])
@login_required
def guardar_registro_hoja(hoja_id, tab):
    h = row_to_dict(execute('SELECT * FROM hojas_tareo WHERE id=?', (hoja_id,), fetchone=True))
    if not h:
        flash('Hoja no encontrada.', 'danger')
        return redirect(url_for('hojas_tareo'))

    dnis_raw = request.form.get('dnis_masivos') or request.form.get('dni') or ''
    dnis = []
    for part in re.split(r'[,;\s]+', dnis_raw):
        d = limpiar_dni(part)
        if len(d) == 8 and d not in dnis:
            dnis.append(d)
    if not dnis:
        flash('Debe digitar o escanear al menos un DNI válido.', 'danger')
        return redirect(url_for('detalle_hoja', hoja_id=hoja_id, tab=tab))

    labor_id = request.form.get('labor_id') or None
    lab = row_to_dict(execute('SELECT * FROM hoja_labores WHERE id=? AND hoja_id=?', (labor_id, hoja_id), fetchone=True)) if labor_id else None
    labor = (lab or h).get('labor')
    grupo = (lab or h).get('grupo')
    turno = limpiar_texto(request.form.get('turno') or (lab or h).get('turno') or 'DIA')
    tipo_tareo = limpiar_texto(request.form.get('tipo_tareo') or (lab or h).get('tipo_tareo') or 'JORNAL')
    try:
        horas = float(request.form.get('horas') or 0)
        cantidad = float(request.form.get('cantidad') or 0)
        a_noct = float(request.form.get('a_noct') or 0)
    except Exception:
        flash('Horas / avance inválido.', 'danger')
        return redirect(url_for('detalle_hoja', hoja_id=hoja_id, tab=tab))
    hora_inicio = request.form.get('hora_inicio') or ('22:00' if turno == 'NOCHE' else '06:30')
    hora_fin = request.form.get('hora_fin') or ('06:00' if turno == 'NOCHE' else '16:30')
    ref_inicio = request.form.get('ref_inicio') or '12:00'
    ref_fin = request.form.get('ref_fin') or '13:00'
    unidad = limpiar_texto(request.form.get('unidad') or ('BALDE' if tab == 'rendimiento' else tipo_tareo))
    metodo = limpiar_texto(request.form.get('metodo') or 'DIGITACIÓN')

    ok = 0; no_encontrados = []
    for dni in dnis:
        t = row_to_dict(execute('SELECT * FROM trabajadores WHERE dni=?', (dni,), fetchone=True))
        if not t:
            no_encontrados.append(dni)
            continue
        h_reg = horas
        if tab == 'rendimiento':
            execute('INSERT INTO lecturas_balde(hoja_id,labor_id,dni,trabajador,fecha_hora,a_diurno,a_noct,metodo,registrado_por) VALUES(?,?,?,?,?,?,?,?,?)',
                    (hoja_id,labor_id,dni,t.get('trabajador',''),now_str(),cantidad,a_noct,metodo,session.get('usuario')), commit=True)
            h_reg = 0
        execute('''INSERT INTO tareos(hoja_id,labor_id,dni,trabajador,empresa,area,cargo,fecha,labor,lote,fundo,horas,cantidad,unidad,observacion,registrado_por,creado_en,hora_inicio,hora_fin,ref_inicio,ref_fin,turno,tipo_tareo)
                   VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',
                (hoja_id,labor_id,dni,t.get('trabajador',''),t.get('empresa',''),t.get('area',''),t.get('cargo',''),h.get('fecha'),labor,limpiar_texto(request.form.get('lote') or grupo),grupo,h_reg,cantidad,unidad,'',session.get('usuario'),now_str(),hora_inicio,hora_fin,ref_inicio,ref_fin,turno,tipo_tareo), commit=True)
        ok += 1
    msg = f'Registro guardado correctamente. Trabajadores registrados: {ok}.'
    if no_encontrados:
        msg += ' No encontrados: ' + ', '.join(no_encontrados)
    flash(msg, 'success' if ok else 'danger')
    return redirect(url_for('detalle_hoja', hoja_id=hoja_id, tab=tab))

# ========================= SOPORTE / CONFIG / SINC =========================
@app.route('/soporte')
@login_required
def soporte():
    return render_page('<div class="phone-wrap"><h2 class="header-title">SOPORTE</h2><div class="floating-card"><b>Canal de soporte</b><p class="small text-muted mb-2">Registra incidencias de sincronización, acceso o lectura de fotocheck.</p><textarea class="form-control" rows="4" placeholder="Describe el problema..."></textarea><a class="btn btn-green w-100 mt-3" href="{{url_for(\'home\')}}">ENVIAR / VOLVER</a></div></div>')

@app.route('/configuraciones')
@admin_required
def configuraciones():
    body = '<div class="phone-wrap"><a class="back-mini" href="{{url_for(\'home\')}}"><i class="bi bi-chevron-left"></i></a><h2 class="header-title">CONFIGURACIÓN</h2><div class="floating-card"><a class="btn btn-green w-100 mb-2" href="{{url_for(\'cargar_base\')}}"><i class="bi bi-people"></i> Cargar trabajadores</a><a class="btn btn-green w-100 mb-2" href="{{url_for(\'cargar_actividades\')}}"><i class="bi bi-diagram-3"></i> Actividades / Labores / Consumidores</a><a class="btn btn-outline-success w-100 mb-2" href="{{url_for(\'plantilla_trabajadores\')}}">Plantilla trabajadores</a><a class="btn btn-outline-success w-100 mb-2" href="{{url_for(\'plantilla_actividades\')}}">Plantilla actividades</a><a class="btn btn-outline-secondary w-100 mb-2" href="{{url_for(\'usuarios\')}}"><i class="bi bi-people"></i> Usuarios</a><a class="btn btn-outline-secondary w-100" href="{{url_for(\'home\')}}">Volver</a></div></div>'
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
    <div class="phone-wrap desktop-pad report-wrap"><div class="config-header"><a class="back-mini" href="{{url_for('home')}}"><i class="bi bi-chevron-left"></i></a><h2 class="header-title mb-2">REPORTES TAREO</h2></div><form class="floating-card mb-2"><div class="row g-2"><div class="col-6"><input class="form-control" type="date" name="desde" value="{{desde}}"></div><div class="col-6"><input class="form-control" type="date" name="hasta" value="{{hasta}}"></div><div class="col-9"><input class="form-control" name="q" value="{{q}}" placeholder="DNI / trabajador / labor"></div><div class="col-3 d-grid"><button class="btn btn-green"><i class="bi bi-search"></i></button></div></div><a class="btn btn-outline-success w-100 mt-2" href="{{url_for('exportar_tareos',desde=desde,hasta=hasta,q=q)}}">EXPORTAR EXCEL</a></form>{% for r in tareos %}<div class="worker-card"><div class="worker-title"><div>{{r.fecha}}<br><b>{{r.trabajador}}</b></div><div class="text-end">{{r.dni}}<br><b>{{r.labor}}</b></div></div><div class="worker-grid"><div><label>HORAS</label><div class="mini-input">{{r.horas}}</div></div><div><label>CANT.</label><div class="mini-input">{{r.cantidad}}</div></div><div><label>UNIDAD</label><div class="mini-input">{{r.unidad}}</div></div></div></div>{% else %}<div class="alert alert-light border text-center">Sin datos.</div>{% endfor %}</div>"""
    return render_page(body, desde=desde, hasta=hasta, q=q, tareos=tareos)

@app.route('/exportar/tareos')
@login_required
def exportar_tareos():
    desde=request.args.get('desde') or today_str(); hasta=request.args.get('hasta') or today_str(); q=request.args.get('q','').strip()
    params=[desde,hasta]; where='WHERE fecha>=? AND fecha<=?'
    if q:
        like=f"%{q.upper()}%"; where += ' AND (dni LIKE ? OR UPPER(trabajador) LIKE ? OR UPPER(labor) LIKE ?)'; params += [like,like,like]
    rows=rows_to_dict(execute(f'SELECT fecha,dni,trabajador,empresa,area,cargo,labor,fundo,lote,horas,cantidad,unidad,turno,tipo_tareo,hora_inicio,hora_fin,ref_inicio,ref_fin,observacion,registrado_por,creado_en FROM tareos {where} ORDER BY creado_en DESC', params, fetchall=True))
    headers=['FECHA','DNI','TRABAJADOR','EMPRESA','AREA','CARGO','LABOR','FUNDO','LOTE','HORAS','CANTIDAD','UNIDAD','TURNO','TIPO_TAREO','HORA_INICIO','HORA_FIN','REF_INICIO','REF_FIN','OBSERVACION','REGISTRADO_POR','CREADO_EN']
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
    <div class="phone-wrap desktop-pad"><div class="config-header"><a class="back-mini" href="{{url_for('configuraciones')}}"><i class="bi bi-chevron-left"></i></a><h2 class="header-title">CONFIGURACIÓN – TRABAJADORES</h2></div><form method="post" enctype="multipart/form-data" class="floating-card mb-2"><label class="form-label">Archivo Excel .xlsx</label><input class="form-control mb-2" type="file" name="archivo" accept=".xlsx" required><button class="btn btn-green w-100">CARGAR BASE</button><a class="btn btn-outline-success w-100 mt-2" href="{{url_for('plantilla_trabajadores')}}">PLANTILLA EXCEL</a></form>{% for r in trabajadores %}<div class="worker-card"><div class="worker-title"><div>{{r.dni}}<br><b>{{r.trabajador}}</b></div><div class="text-end">{{r.empresa}}<br><b>{{r.cargo}}</b></div></div></div>{% endfor %}</div>"""
    return render_page(body, trabajadores=trabajadores)

@app.route('/plantilla-trabajadores')
@admin_required
def plantilla_trabajadores():
    headers=['DNI','TRABAJADOR','EMPRESA','CARGO','ESTADO','FECHA']
    rows=[{'DNI':'12345678','TRABAJADOR':'APELLIDOS Y NOMBRES','EMPRESA':'AQUANQA I','CARGO':'OPERARIO','ESTADO':'ACTIVO','FECHA':today_str()}]
    return excel_response(headers, rows, 'plantilla_trabajadores.xlsx', 'TRABAJADORES')


@app.route('/cargar-actividades', methods=['GET','POST'])
@admin_required
def cargar_actividades():
    if request.method == 'POST':
        f = request.files.get('archivo')
        if not f or not f.filename.lower().endswith('.xlsx'):
            flash('Suba un Excel .xlsx válido.', 'danger'); return redirect(url_for('cargar_actividades'))
        wb = load_workbook(f, data_only=True, read_only=True); ws = wb.active
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            flash('Excel vacío.', 'danger'); return redirect(url_for('cargar_actividades'))
        headers = [normalizar_columna(c) for c in rows[0]]
        aliases = {
            'COD. ACTIVIDAD':'cod_actividad','COD ACTIVIDAD':'cod_actividad','COD_ACTIVIDAD':'cod_actividad',
            'DESCRIPCION ACTIVIDAD':'desc_actividad','DESCRIPCIÓN ACTIVIDAD':'desc_actividad','DESC ACTIVIDAD':'desc_actividad',
            'COD. LABOR':'cod_labor','COD LABOR':'cod_labor','COD_LABOR':'cod_labor',
            'DESCRIPCION LABOR':'desc_labor','DESCRIPCIÓN LABOR':'desc_labor','DESC LABOR':'desc_labor',
            'COD. CONSUMIDOR':'cod_consumidor','COD CONSUMIDOR':'cod_consumidor','COD_CONSUMIDOR':'cod_consumidor',
            'DESCRIPCION CONSUMIDOR':'desc_consumidor','DESCRIPCIÓN CONSUMIDOR':'desc_consumidor','DESC CONSUMIDOR':'desc_consumidor'
        }
        required = ['desc_actividad','desc_labor']
        mapped = {i: aliases.get(h) for i,h in enumerate(headers)}
        if not any(v=='desc_actividad' for v in mapped.values()) or not any(v=='desc_labor' for v in mapped.values()):
            flash('La plantilla debe tener Descripción Actividad y Descripción Labor.', 'danger'); return redirect(url_for('cargar_actividades'))
        conn=get_conn(); cur=conn.cursor(); ins=0; omi=0; ahora=now_str()
        for row in rows[1:]:
            data={'cod_actividad':'','desc_actividad':'','cod_labor':'','desc_labor':'','cod_consumidor':'','desc_consumidor':''}
            for i,val in enumerate(row):
                k=mapped.get(i)
                if k: data[k]=limpiar_texto(val)
            if not data['desc_actividad'] or not data['desc_labor']:
                omi += 1; continue
            cur.execute(qmark('INSERT INTO actividades_maestras(cod_actividad,desc_actividad,cod_labor,desc_labor,cod_consumidor,desc_consumidor,estado,fecha_carga) VALUES(?,?,?,?,?,?,?,?)'),
                        (data['cod_actividad'],data['desc_actividad'],data['cod_labor'],data['desc_labor'],data['cod_consumidor'],data['desc_consumidor'],'ACTIVO',ahora)); ins += 1
        conn.commit(); cur.close(); conn.close()
        flash(f'Actividades cargadas. Insertados: {ins} | Omitidos: {omi}', 'success')
        return redirect(url_for('cargar_actividades'))
    datos = rows_to_dict(execute('SELECT * FROM actividades_maestras ORDER BY fecha_carga DESC, desc_actividad, desc_labor LIMIT 150', fetchall=True))
    body = """
    <div class="phone-wrap desktop-pad"><a class="back-mini" href="{{url_for('configuraciones')}}"><i class="bi bi-chevron-left"></i></a><h2 class="header-title">ACTIVIDADES / LABORES / CONSUMIDORES</h2>
    <form method="post" enctype="multipart/form-data" class="floating-card mb-2"><label class="form-label">Archivo Excel .xlsx</label><input class="form-control mb-2" type="file" name="archivo" accept=".xlsx" required><button class="btn btn-green w-100">CARGAR ACTIVIDADES</button><a class="btn btn-outline-success w-100 mt-2" href="{{url_for('plantilla_actividades')}}">PLANTILLA ACTIVIDADES</a></form>
    {% for r in datos %}<div class="worker-card"><div class="worker-title"><div>ACTIVIDAD<br><b>{{r.desc_actividad}}</b></div><div class="text-end">LABOR<br><b>{{r.desc_labor}}</b></div></div><div class="small-label mt-2">CONSUMIDOR</div><div class="small-value">{{r.desc_consumidor or 'NO OBLIGATORIO'}}</div></div>{% else %}<div class="alert alert-light border text-center">Sin actividades cargadas.</div>{% endfor %}</div>"""
    return render_page(body, datos=datos)

@app.route('/plantilla-actividades')
@admin_required
def plantilla_actividades():
    headers=['Cod. Actividad','Descripción Actividad','Cod. Labor','Descripción Labor','Cod. Consumidor','Descripción Consumidor']
    rows=[{'Cod. Actividad':'ACT001','Descripción Actividad':'COSECHA','Cod. Labor':'LAB001','Descripción Labor':'COSECHA MANUAL','Cod. Consumidor':'CON001','Descripción Consumidor':'CAMPO 01'},
          {'Cod. Actividad':'ACT001','Descripción Actividad':'COSECHA','Cod. Labor':'LAB002','Descripción Labor':'COSECHA SELECTIVA','Cod. Consumidor':'','Descripción Consumidor':''}]
    return excel_response(headers, rows, 'plantilla_actividades_labores_consumidores.xlsx', 'ACTIVIDADES')

@app.route('/api/actividades-maestras')
@login_required
def api_actividades_maestras():
    return jsonify(ok=True, data=get_actividades_maestras())

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
    body='<div class="phone-wrap"><a class="back-mini" href="{{url_for(\'configuraciones\')}}"><i class="bi bi-chevron-left"></i></a><h2 class="header-title">USUARIOS</h2><form method="post" class="floating-card mb-2"><input class="form-control mb-2" name="usuario" placeholder="Usuario" required><input class="form-control mb-2" name="nombres" placeholder="Nombres"><input class="form-control mb-2" type="password" name="clave" placeholder="Clave" required><select class="form-select mb-2" name="rol"><option value="operador">operador</option><option value="admin">admin</option></select><select class="form-select mb-2" name="estado"><option>ACTIVO</option><option>INACTIVO</option></select><button class="btn btn-green w-100">Guardar</button></form>{% for u in users %}<div class="worker-card"><b>{{u.usuario}}</b> · {{u.rol}}<br><span class="small text-muted">{{u.nombres}} · {{u.estado}}</span></div>{% endfor %}</div>'
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
