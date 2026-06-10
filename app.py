# -*- coding: utf-8 -*-
"""
CONTROL DE ASISTENCIA Y TAREO - PRIZE PRO
Listo para GitHub + Render. Flask + SQLite/PostgreSQL + PWA.

Incluye:
- Login admin/operador.
- Carga de trabajadores por Excel.
- Captura de fotocheck QR/codigo de barras por camara o lector USB.
- Marcacion de entrada/salida con foto opcional desde celular.
- Registro de tareos por trabajador, labor, lote, horas, cantidad y observacion.
- Dashboard, filtros y exportacion Excel.
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
def is_pg():
    return USE_POSTGRES and psycopg2 is not None

def get_conn():
    if is_pg():
        return psycopg2.connect(DATABASE_URL, cursor_factory=psycopg2.extras.RealDictCursor)
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def qmark(sql):
    return sql.replace("?", "%s") if is_pg() else sql

def row_to_dict(r):
    return dict(r) if r else None

def rows_to_dict(rows):
    return [row_to_dict(r) for r in (rows or [])]

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
    cur.execute(qmark(f"""
        CREATE TABLE IF NOT EXISTS usuarios(
            id {idtype}, usuario TEXT UNIQUE NOT NULL, password_hash TEXT NOT NULL,
            nombres TEXT, rol TEXT DEFAULT 'operador', estado TEXT DEFAULT 'ACTIVO', creado_en TEXT
        )
    """))
    cur.execute(qmark(f"""
        CREATE TABLE IF NOT EXISTS trabajadores(
            id {idtype}, dni TEXT UNIQUE NOT NULL, trabajador TEXT, empresa TEXT,
            area TEXT, cargo TEXT, actividad TEXT, planilla TEXT, estado TEXT DEFAULT 'ACTIVO', fecha_carga TEXT
        )
    """))
    cur.execute(qmark(f"""
        CREATE TABLE IF NOT EXISTS asistencia(
            id {idtype}, dni TEXT NOT NULL, trabajador TEXT, empresa TEXT, area TEXT, cargo TEXT,
            tipo TEXT NOT NULL, fecha TEXT NOT NULL, hora TEXT NOT NULL, fecha_hora TEXT NOT NULL,
            metodo TEXT, foto_path TEXT, latitud TEXT, longitud TEXT, registrado_por TEXT, observacion TEXT
        )
    """))
    cur.execute(qmark(f"""
        CREATE TABLE IF NOT EXISTS tareos(
            id {idtype}, dni TEXT NOT NULL, trabajador TEXT, empresa TEXT, area TEXT, cargo TEXT,
            fecha TEXT NOT NULL, labor TEXT, lote TEXT, fundo TEXT, horas REAL DEFAULT 0,
            cantidad REAL DEFAULT 0, unidad TEXT, observacion TEXT, registrado_por TEXT, creado_en TEXT
        )
    """))
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
            return redirect(url_for("dashboard"))
        return f(*args, **kwargs)
    return w

def guardar_foto_base64(data_url, dni, tipo):
    if not data_url or not data_url.startswith("data:image"):
        return ""
    try:
        header, data = data_url.split(",", 1)
        raw = base64.b64decode(data)
        name = f"{dni}_{tipo}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.jpg"
        path = os.path.join(FOTO_DIR, name)
        with open(path, "wb") as f: f.write(raw)
        return path
    except Exception:
        return ""

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

# ========================= HTML =========================
BASE_HTML = r"""
<!doctype html><html lang="es"><head>
<meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="theme-color" content="#16a34a"><link rel="manifest" href="{{ url_for('manifest') }}">
<title>{{ title or 'Asistencia y Tareo PRIZE' }}</title>
<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet">
<link href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.css" rel="stylesheet">
<script src="https://unpkg.com/html5-qrcode" type="text/javascript"></script>
<script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
<style>
:root{--green:#16a34a;--green2:#22c55e;--dark:#07111f;--panel:#0f172a;--muted:#64748b;--line:#e5e7eb}
*{box-sizing:border-box} body{margin:0;background:linear-gradient(135deg,#f8fafc,#eef7f1);font-family:Inter,Segoe UI,Arial,sans-serif;color:#0f172a}.sidebar{position:fixed;left:0;top:0;bottom:0;width:290px;background:radial-gradient(circle at top,#14532d 0,#0f172a 44%,#07111f 100%);color:white;padding:17px 13px;z-index:20;box-shadow:12px 0 34px rgba(2,6,23,.2)}.content{margin-left:290px;padding:22px}.logoBox{display:flex;align-items:center;gap:12px;padding:10px 8px 18px;border-bottom:1px solid rgba(255,255,255,.12);margin-bottom:14px}.logoIcon{width:52px;height:52px;border-radius:16px;background:#fff;color:#15803d;display:grid;place-items:center;font-size:28px}.brandTitle{font-size:19px;font-weight:900;line-height:1}.brandTitle small{display:block;color:#bbf7d0;font-size:12px;font-weight:700;margin-top:4px}.navLabel{font-size:11px;text-transform:uppercase;letter-spacing:.09em;color:#86efac;margin:16px 10px 6px}.nav-link{display:flex;align-items:center;gap:12px;color:#cbd5e1;border-radius:16px;margin:6px 2px;padding:13px 14px;font-weight:750}.nav-link i{font-size:19px;min-width:24px;text-align:center}.nav-link:hover,.nav-link.active{background:linear-gradient(135deg,var(--green),#15803d);color:white;box-shadow:0 8px 18px rgba(22,163,74,.25)}.userBox{position:absolute;bottom:14px;left:14px;right:14px;background:rgba(255,255,255,.09);border-radius:18px;padding:12px;color:#d1fae5;font-size:13px}.topbar{background:rgba(255,255,255,.88);backdrop-filter:blur(14px);border:1px solid rgba(226,232,240,.9);border-radius:24px;padding:15px 18px;margin-bottom:18px;box-shadow:0 12px 30px rgba(15,23,42,.06)}.card-pro{border:1px solid rgba(226,232,240,.95);border-radius:24px;box-shadow:0 14px 34px rgba(15,23,42,.07);background:rgba(255,255,255,.97)}.kpi{border-radius:24px;padding:20px;background:white;border:1px solid #e2e8f0;box-shadow:0 12px 26px rgba(15,23,42,.06);position:relative;overflow:hidden}.kpi:after{content:"";position:absolute;right:-20px;top:-20px;width:86px;height:86px;border-radius:50%;background:#dcfce7}.kpiIcon{width:46px;height:46px;border-radius:16px;display:grid;place-items:center;background:#dcfce7;color:#166534;font-size:23px}.btn-pro{border-radius:15px;font-weight:800}.form-control,.form-select{border-radius:15px;padding:11px 13px;border:1px solid #dbe3ef}.form-control:focus,.form-select:focus{box-shadow:0 0 0 .25rem rgba(22,163,74,.16);border-color:#22c55e}.table{font-size:14px}.table thead th{color:#475569;background:#f8fafc}.badge-soft{background:#dcfce7;color:#166534;border-radius:999px;padding:8px 12px;font-weight:800}.scanPulse{animation:pulse 1s ease-in-out 1}@keyframes pulse{0%{box-shadow:0 0 0 0 rgba(34,197,94,.6)}100%{box-shadow:0 0 0 18px rgba(34,197,94,0)}}.login-page{min-height:100vh;background:linear-gradient(180deg,#f8fbfc 0%,#ecf7f1 48%,#0f7f60 100%)}.login-page .content{margin-left:0;min-height:100vh;display:flex;align-items:center;justify-content:center}.login-card{width:min(480px,92vw);padding:44px!important;border-radius:32px!important}.login-badge{width:92px;height:92px;border-radius:999px;background:#ecfdf5;color:#15803d;display:grid;place-items:center;margin:-10px auto 16px;font-size:44px}
@media(max-width:920px){.sidebar{position:sticky;top:0;width:100%;bottom:auto;border-radius:0 0 22px 22px;padding:8px;z-index:50}.content{margin-left:0;padding:10px}.logoBox{padding:4px 4px 8px;margin-bottom:7px}.logoIcon{width:34px;height:34px;border-radius:10px;font-size:19px}.brandTitle{font-size:14px}.brandTitle small{font-size:9px}.navMenu{display:flex!important;flex-direction:row!important;overflow-x:auto;gap:8px;padding:8px;border-radius:18px;background:rgba(2,6,23,.35);white-space:nowrap}.nav-link{flex:0 0 auto;margin:0;padding:9px 12px;border-radius:13px;font-size:12px;gap:6px;background:rgba(255,255,255,.08)}.nav-link i{font-size:15px;min-width:16px}.navLabel,.userBox{display:none}.topbar{border-radius:18px;padding:12px}.topbar h2{font-size:22px}.card-pro{border-radius:18px}.row.g-4{--bs-gutter-y:12px}}
</style></head><body class="{{ 'login-page' if not session.get('usuario') else '' }}">
{% if session.get('usuario') %}<aside class="sidebar"><div class="logoBox"><div class="logoIcon"><i class="bi bi-qr-code-scan"></i></div><div class="brandTitle">PRIZE PRO<small>Asistencia y Tareo</small></div></div><div class="navMenu nav flex-column"><a class="nav-link {% if active=='dashboard' %}active{% endif %}" href="{{ url_for('dashboard') }}"><i class="bi bi-speedometer2"></i><span>Dashboard</span></a><a class="nav-link {% if active=='marcar' %}active{% endif %}" href="{{ url_for('marcar') }}"><i class="bi bi-camera"></i><span>Marcar asistencia</span></a><a class="nav-link {% if active=='tareo' %}active{% endif %}" href="{{ url_for('tareo') }}"><i class="bi bi-clipboard2-check"></i><span>Registrar tareo</span></a><a class="nav-link {% if active=='reportes' %}active{% endif %}" href="{{ url_for('reportes') }}"><i class="bi bi-table"></i><span>Reportes</span></a>{% if session.get('rol')=='admin' %}<div class="navLabel">Admin</div><a class="nav-link {% if active=='carga' %}active{% endif %}" href="{{ url_for('cargar_base') }}"><i class="bi bi-file-earmark-excel"></i><span>Cargar base</span></a><a class="nav-link {% if active=='usuarios' %}active{% endif %}" href="{{ url_for('usuarios') }}"><i class="bi bi-people"></i><span>Usuarios</span></a>{% endif %}<a class="nav-link" href="{{ url_for('logout') }}"><i class="bi bi-box-arrow-right"></i><span>Salir</span></a></div><div class="userBox"><i class="bi bi-person-circle me-1"></i>{{ session.get('usuario') }} · {{ session.get('rol')|upper }}</div></aside>{% endif %}
<main class="content">{% with messages=get_flashed_messages(with_categories=true) %}{% if messages %}{% for cat,msg in messages %}<div class="alert alert-{{cat}} alert-dismissible fade show card-pro">{{msg}}<button class="btn-close" data-bs-dismiss="alert"></button></div>{% endfor %}{% endif %}{% endwith %}{{ body|safe }}</main>
<audio id="sndOk"><source src="data:audio/wav;base64,UklGRjQAAABXQVZFZm10IBAAAAABAAEAESsAACJWAAACABAAZGF0YRAAAAAAAP//AAD//wAA//8AAP//AAD//wAA//8=" type="audio/wav"></audio>
<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/js/bootstrap.bundle.min.js"></script><script>function beep(){try{const a=document.getElementById('sndOk');a.currentTime=0;a.play().catch(()=>{});}catch(e){}} if('serviceWorker' in navigator){navigator.serviceWorker.register('/sw.js').catch(()=>{});}</script>
</body></html>
"""

def render_page(body, title="Asistencia y Tareo PRIZE", active="dashboard", **ctx):
    return render_template_string(BASE_HTML, body=render_template_string(body, **ctx), title=title, active=active)

# ========================= AUTH =========================
@app.route('/login', methods=['GET','POST'])
def login():
    if request.method == 'POST':
        usuario = request.form.get('usuario','').strip()
        password = request.form.get('password','')
        row = row_to_dict(execute('SELECT * FROM usuarios WHERE usuario=?', (usuario,), fetchone=True))
        if row and row.get('estado','ACTIVO') == 'ACTIVO' and check_password_hash(row['password_hash'], password):
            session['usuario']=usuario; session['rol']=row.get('rol','operador'); session['nombres']=row.get('nombres') or usuario
            return redirect(url_for('dashboard'))
        flash('Usuario/clave incorrecta o usuario inactivo.', 'danger')
    body = """
    <div class="card card-pro login-card text-center"><div class="login-badge"><i class="bi bi-qr-code-scan"></i></div><h2 class="fw-bold mb-1">PRIZE PRO</h2><div class="text-muted mb-4">Control de asistencia y tareo</div><form method="post" class="text-start"><label class="fw-bold">Usuario</label><input class="form-control mb-3" name="usuario" required autofocus placeholder="admin"><label class="fw-bold">Clave</label><input class="form-control mb-3" name="password" type="password" required placeholder="admin123"><button class="btn btn-success btn-pro w-100"><i class="bi bi-shield-check me-2"></i>Ingresar</button></form></div>
    """
    return render_page(body, title='Login', active='')

@app.route('/logout')
def logout():
    session.clear(); return redirect(url_for('login'))

# ========================= DASHBOARD =========================
@app.route('/')
@login_required
def dashboard():
    f = today_str()
    total_trab = scalar('SELECT COUNT(*) AS c FROM trabajadores')
    entradas = scalar("SELECT COUNT(*) AS c FROM asistencia WHERE fecha=? AND tipo='ENTRADA'", (f,))
    salidas = scalar("SELECT COUNT(*) AS c FROM asistencia WHERE fecha=? AND tipo='SALIDA'", (f,))
    tareos_hoy = scalar('SELECT COUNT(*) AS c FROM tareos WHERE fecha=?', (f,))
    ult = rows_to_dict(execute('SELECT dni,trabajador,tipo,fecha_hora,metodo,registrado_por FROM asistencia ORDER BY fecha_hora DESC LIMIT 8', fetchall=True))
    labor = rows_to_dict(execute('SELECT COALESCE(labor,\'SIN LABOR\') AS labor, COUNT(*) AS total FROM tareos GROUP BY labor ORDER BY total DESC LIMIT 8', fetchall=True))
    import json
    body = """
    <div class="topbar d-flex justify-content-between align-items-center flex-wrap gap-2"><div><h2 class="fw-bold mb-0">Dashboard de control</h2><div class="text-muted">Asistencia por fotocheck QR/código de barras y tareo diario.</div></div><span class="badge-soft">{{ fecha }}</span></div>
    <div class="row g-3 mb-3"><div class="col-md-3"><div class="kpi"><div class="d-flex justify-content-between"><div><div class="text-muted">Trabajadores</div><div class="fs-2 fw-bold">{{total_trab}}</div></div><div class="kpiIcon"><i class="bi bi-people"></i></div></div></div></div><div class="col-md-3"><div class="kpi"><div class="d-flex justify-content-between"><div><div class="text-muted">Entradas hoy</div><div class="fs-2 fw-bold text-success">{{entradas}}</div></div><div class="kpiIcon"><i class="bi bi-box-arrow-in-right"></i></div></div></div></div><div class="col-md-3"><div class="kpi"><div class="d-flex justify-content-between"><div><div class="text-muted">Salidas hoy</div><div class="fs-2 fw-bold">{{salidas}}</div></div><div class="kpiIcon"><i class="bi bi-box-arrow-right"></i></div></div></div></div><div class="col-md-3"><div class="kpi"><div class="d-flex justify-content-between"><div><div class="text-muted">Tareos hoy</div><div class="fs-2 fw-bold">{{tareos_hoy}}</div></div><div class="kpiIcon"><i class="bi bi-clipboard2-check"></i></div></div></div></div></div>
    <div class="row g-3"><div class="col-lg-7"><div class="card card-pro p-3"><div class="d-flex justify-content-between"><h5 class="fw-bold mb-0">Últimas marcaciones</h5><a class="btn btn-success btn-pro" href="{{url_for('marcar')}}">Marcar</a></div><div class="table-responsive mt-3"><table class="table table-hover align-middle"><thead><tr><th>DNI</th><th>Trabajador</th><th>Tipo</th><th>Fecha/hora</th><th>Método</th><th>Usuario</th></tr></thead><tbody>{% for r in ult %}<tr><td class="fw-bold">{{r.dni}}</td><td>{{r.trabajador}}</td><td><span class="badge bg-{{ 'success' if r.tipo=='ENTRADA' else 'secondary' }}">{{r.tipo}}</span></td><td>{{r.fecha_hora}}</td><td>{{r.metodo}}</td><td>{{r.registrado_por}}</td></tr>{% else %}<tr><td colspan="6" class="text-center text-muted py-4">Sin marcaciones.</td></tr>{% endfor %}</tbody></table></div></div></div><div class="col-lg-5"><div class="card card-pro p-4"><h5 class="fw-bold">Tareos por labor</h5><canvas id="chartLabor" height="170"></canvas></div></div></div>
    <script>new Chart(document.getElementById('chartLabor'),{type:'bar',data:{labels:{{ labels|safe }},datasets:[{label:'Registros',data:{{ values|safe }} }]},options:{plugins:{legend:{display:false}},scales:{y:{beginAtZero:true,ticks:{precision:0}}}}});</script>
    """
    return render_page(body, active='dashboard', fecha=f, total_trab=total_trab, entradas=entradas, salidas=salidas, tareos_hoy=tareos_hoy, ult=ult, labels=json.dumps([x['labor'] for x in labor]), values=json.dumps([x['total'] for x in labor]))

# ========================= MARCACION =========================
@app.route('/marcar')
@login_required
def marcar():
    body = """
    <div class="topbar d-flex justify-content-between align-items-center flex-wrap gap-2"><div><h2 class="fw-bold mb-0">Marcar asistencia</h2><div class="text-muted">Escanee fotocheck QR/código de barras con cámara, lector USB o digitación.</div></div><span class="badge-soft"><i class="bi bi-camera me-1"></i>Celular / PC</span></div>
    <div class="row g-4"><div class="col-lg-5"><div class="card card-pro p-4" id="scanCard"><h5 class="fw-bold"><i class="bi bi-upc-scan me-2"></i>Captura de fotocheck</h5><label class="form-label fw-bold mt-2">DNI / QR / Código de barras</label><div class="input-group mb-3"><input id="dni" class="form-control" placeholder="Escanee o digite DNI" maxlength="30" autofocus autocomplete="off"><button class="btn btn-success btn-pro" onclick="buscarDni('DIGITACION')">Buscar</button></div><div class="d-grid gap-2"><button class="btn btn-outline-success btn-pro" onclick="iniciarCamara()"><i class="bi bi-camera-video me-1"></i>Escanear con cámara</button><button class="btn btn-outline-secondary btn-pro" onclick="detenerCamara()">Detener cámara</button></div><div id="reader" class="mt-3" style="width:100%;display:none"></div><div class="alert alert-light border mt-3 small mb-0"><b>Recomendación:</b> el QR o código de barras del fotocheck debe contener el DNI de 8 dígitos.</div></div></div>
    <div class="col-lg-7"><div class="card card-pro p-4"><h5 class="fw-bold"><i class="bi bi-person-check me-2"></i>Confirmación</h5><div id="msg" class="alert alert-info">Esperando captura.</div><form id="frm" style="display:none" onsubmit="guardarMarcacion(event)"><input type="hidden" id="metodo" value="DIGITACION"><div class="row g-3"><div class="col-md-4"><label class="fw-bold">DNI</label><input class="form-control" id="f_dni" readonly></div><div class="col-md-8"><label class="fw-bold">Trabajador</label><input class="form-control" id="trabajador" readonly></div><div class="col-md-6"><label class="fw-bold">Empresa</label><input class="form-control" id="empresa" readonly></div><div class="col-md-6"><label class="fw-bold">Área / Cargo</label><input class="form-control" id="area_cargo" readonly></div><div class="col-md-6"><label class="fw-bold">Tipo de marca</label><select id="tipo" class="form-select"><option>ENTRADA</option><option>SALIDA</option></select></div><div class="col-md-6"><label class="fw-bold">Fecha/hora</label><input class="form-control" id="fecha_hora" readonly></div><div class="col-12"><label class="fw-bold">Foto evidencia opcional</label><video id="video" autoplay playsinline style="width:100%;max-height:230px;border-radius:18px;background:#0f172a"></video><canvas id="canvas" style="display:none"></canvas><div class="d-flex gap-2 mt-2 flex-wrap"><button type="button" class="btn btn-outline-success btn-pro" onclick="activarFoto()">Activar foto</button><button type="button" class="btn btn-outline-secondary btn-pro" onclick="capturarFoto()">Capturar foto</button></div><input type="hidden" id="foto_data"></div><div class="col-12"><label class="fw-bold">Observación</label><textarea id="observacion" class="form-control" rows="2" placeholder="Opcional"></textarea></div></div><button class="btn btn-success btn-pro mt-3 px-4"><i class="bi bi-check-circle me-1"></i>Guardar marcación</button><button type="button" class="btn btn-outline-secondary btn-pro mt-3" onclick="limpiar()">Nueva captura</button></form></div></div></div>
    <script>
    let html5QrCode=null,buscando=false,timer=null,streamFoto=null; const dniInput=document.getElementById('dni'); dniInput.focus();
    dniInput.addEventListener('input',()=>{clearTimeout(timer); const d=limpiarDni(dniInput.value); if(d.length>=8){dniInput.value=d; timer=setTimeout(()=>buscarDni('AUTO/LECTOR'),180);}}); dniInput.addEventListener('keydown',e=>{if(e.key==='Enter'){e.preventDefault();buscarDni('DIGITACION');}});
    function limpiarDni(v){let raw=(v||'').toString(); let m=raw.match(/(?:^|\D)(\d{8})(?:\D|$)/); let d=m?m[1]:raw.replace(/\D/g,''); return d.length>=8?d.slice(-8):d;} function mostrar(txt,tipo){const m=document.getElementById('msg');m.className='alert alert-'+tipo;m.innerText=txt;}
    async function buscarDni(metodo='DIGITACION'){if(buscando)return; const dni=limpiarDni(dniInput.value); dniInput.value=dni; document.getElementById('metodo').value=metodo; if(dni.length!==8){mostrar('Ingrese DNI válido de 8 dígitos.','warning');return;} buscando=true; const r=await fetch('/api/trabajador/'+dni); const data=await r.json(); buscando=false; if(!data.ok){frm.style.display='none';mostrar(data.msg,'danger');return;} beep(); scanCard.classList.add('scanPulse'); setTimeout(()=>scanCard.classList.remove('scanPulse'),900); const t=data.trabajador; frm.style.display='block'; f_dni.value=t.dni||''; trabajador.value=t.trabajador||''; empresa.value=t.empresa||''; area_cargo.value=((t.area||'')+' / '+(t.cargo||'')).replace(/^ \/ | \/ $/g,''); tipo.value=data.sugerido||'ENTRADA'; fecha_hora.value=new Date().toLocaleString(); mostrar('Trabajador encontrado. Confirme la marcación.','success');}
    async function guardarMarcacion(e){e.preventDefault(); const payload={dni:f_dni.value,tipo:tipo.value,metodo:metodo.value,foto_data:foto_data.value,observacion:observacion.value,latitud:'',longitud:''}; if(navigator.geolocation){navigator.geolocation.getCurrentPosition(async p=>{payload.latitud=p.coords.latitude;payload.longitud=p.coords.longitude; await enviar(payload);},async()=>await enviar(payload),{timeout:1600});}else{await enviar(payload);} }
    async function enviar(payload){const r=await fetch('/api/marcacion',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(payload)}); const data=await r.json(); mostrar(data.msg,data.ok?'success':'danger'); if(data.ok){beep(); setTimeout(()=>limpiar(),700);}}
    function limpiar(){frm.style.display='none'; dniInput.value=''; ['f_dni','trabajador','empresa','area_cargo','foto_data','observacion'].forEach(id=>{const el=document.getElementById(id); if(el) el.value='';}); mostrar('Esperando captura.','info'); dniInput.focus();}
    function iniciarCamara(){reader.style.display='block'; html5QrCode=new Html5Qrcode('reader'); html5QrCode.start({facingMode:'environment'},{fps:10,qrbox:250},decoded=>{dniInput.value=limpiarDni(decoded); buscarDni('QR/CAMARA'); detenerCamara();}).catch(()=>mostrar('No se pudo activar cámara. Revise permisos del navegador.','warning'));}
    function detenerCamara(){if(html5QrCode){html5QrCode.stop().catch(()=>{});html5QrCode=null;} reader.style.display='none';}
    async function activarFoto(){try{streamFoto=await navigator.mediaDevices.getUserMedia({video:{facingMode:'user'},audio:false}); video.srcObject=streamFoto;}catch(e){mostrar('No se pudo activar foto evidencia.','warning');}}
    function capturarFoto(){if(!video.srcObject){mostrar('Primero active la foto.','warning');return;} canvas.width=video.videoWidth||640; canvas.height=video.videoHeight||480; canvas.getContext('2d').drawImage(video,0,0,canvas.width,canvas.height); foto_data.value=canvas.toDataURL('image/jpeg',0.72); mostrar('Foto capturada. Ya puede guardar.','success');}
    </script>
    """
    return render_page(body, active='marcar')

# ========================= TAREO =========================
@app.route('/tareo')
@login_required
def tareo():
    body = """
    <div class="topbar d-flex justify-content-between align-items-center flex-wrap gap-2"><div><h2 class="fw-bold mb-0">Registro de tareo</h2><div class="text-muted">Registre labor, lote, horas y cantidad por trabajador.</div></div><span class="badge-soft"><i class="bi bi-clipboard2-check me-1"></i>Tareo diario</span></div>
    <div class="row g-4"><div class="col-lg-5"><div class="card card-pro p-4" id="scanCard"><h5 class="fw-bold"><i class="bi bi-upc-scan me-2"></i>Buscar trabajador</h5><label class="form-label fw-bold mt-2">DNI / QR / Código de barras</label><div class="input-group mb-3"><input id="dni" class="form-control" placeholder="Escanee o digite DNI" maxlength="30" autofocus autocomplete="off"><button class="btn btn-success btn-pro" onclick="buscarDni('DIGITACION')">Buscar</button></div><div class="d-grid gap-2"><button class="btn btn-outline-success btn-pro" onclick="iniciarCamara()">Escanear con cámara</button><button class="btn btn-outline-secondary btn-pro" onclick="detenerCamara()">Detener cámara</button></div><div id="reader" class="mt-3" style="width:100%;display:none"></div></div></div>
    <div class="col-lg-7"><div class="card card-pro p-4"><h5 class="fw-bold"><i class="bi bi-list-check me-2"></i>Detalle del tareo</h5><div id="msg" class="alert alert-info">Esperando trabajador.</div><form id="frm" style="display:none" onsubmit="guardarTareo(event)"><div class="row g-3"><div class="col-md-4"><label class="fw-bold">DNI</label><input class="form-control" id="f_dni" readonly></div><div class="col-md-8"><label class="fw-bold">Trabajador</label><input class="form-control" id="trabajador" readonly></div><div class="col-md-6"><label class="fw-bold">Fecha</label><input type="date" class="form-control" id="fecha" required></div><div class="col-md-6"><label class="fw-bold">Labor</label><input class="form-control" id="labor" list="labores" placeholder="COSECHA / SELECCIÓN / CAMPO" required><datalist id="labores"><option>COSECHA</option><option>SELECCIÓN</option><option>RALEO</option><option>PODA</option><option>EMPAQUE</option><option>LIMPIEZA</option></datalist></div><div class="col-md-4"><label class="fw-bold">Fundo/Sede</label><input class="form-control" id="fundo" placeholder="Opcional"></div><div class="col-md-4"><label class="fw-bold">Lote</label><input class="form-control" id="lote" placeholder="Ej. LOTE 01"></div><div class="col-md-4"><label class="fw-bold">Horas</label><input type="number" step="0.01" min="0" class="form-control" id="horas" value="8"></div><div class="col-md-6"><label class="fw-bold">Cantidad</label><input type="number" step="0.01" min="0" class="form-control" id="cantidad" value="0"></div><div class="col-md-6"><label class="fw-bold">Unidad</label><select id="unidad" class="form-select"><option>HORAS</option><option>JORNAL</option><option>KG</option><option>JABAS</option><option>UNIDADES</option></select></div><div class="col-12"><label class="fw-bold">Observación</label><textarea id="observacion" class="form-control" rows="2"></textarea></div></div><button class="btn btn-success btn-pro mt-3 px-4">Guardar tareo</button><button type="button" class="btn btn-outline-secondary btn-pro mt-3" onclick="limpiar()">Nuevo</button></form></div></div></div>
    <script>
    let html5QrCode=null,buscando=false,timer=null; const dniInput=document.getElementById('dni'); fecha.value=new Date().toISOString().slice(0,10); dniInput.focus(); dniInput.addEventListener('input',()=>{clearTimeout(timer); const d=limpiarDni(dniInput.value); if(d.length>=8){dniInput.value=d; timer=setTimeout(()=>buscarDni('AUTO/LECTOR'),180);}}); dniInput.addEventListener('keydown',e=>{if(e.key==='Enter'){e.preventDefault();buscarDni('DIGITACION');}}); function limpiarDni(v){let raw=(v||'').toString(); let m=raw.match(/(?:^|\D)(\d{8})(?:\D|$)/); let d=m?m[1]:raw.replace(/\D/g,''); return d.length>=8?d.slice(-8):d;} function mostrar(txt,tipo){msg.className='alert alert-'+tipo;msg.innerText=txt;}
    async function buscarDni(metodo='DIGITACION'){if(buscando)return; const dni=limpiarDni(dniInput.value); dniInput.value=dni; if(dni.length!==8){mostrar('Ingrese DNI válido.','warning');return;} buscando=true; const r=await fetch('/api/trabajador/'+dni); const data=await r.json(); buscando=false; if(!data.ok){frm.style.display='none';mostrar(data.msg,'danger');return;} beep(); scanCard.classList.add('scanPulse'); setTimeout(()=>scanCard.classList.remove('scanPulse'),900); const t=data.trabajador; frm.style.display='block'; f_dni.value=t.dni||''; trabajador.value=t.trabajador||''; mostrar('Trabajador encontrado. Complete el tareo.','success'); labor.focus();}
    async function guardarTareo(e){e.preventDefault(); const payload={dni:f_dni.value,fecha:fecha.value,labor:labor.value.toUpperCase(),fundo:fundo.value.toUpperCase(),lote:lote.value.toUpperCase(),horas:horas.value,cantidad:cantidad.value,unidad:unidad.value,observacion:observacion.value.toUpperCase()}; const r=await fetch('/api/tareo',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(payload)}); const data=await r.json(); mostrar(data.msg,data.ok?'success':'danger'); if(data.ok){beep(); setTimeout(()=>limpiar(),700);}}
    function limpiar(){frm.style.display='none'; dniInput.value=''; ['f_dni','trabajador','labor','fundo','lote','observacion'].forEach(id=>{const el=document.getElementById(id); if(el) el.value='';}); horas.value='8'; cantidad.value='0'; fecha.value=new Date().toISOString().slice(0,10); mostrar('Esperando trabajador.','info'); dniInput.focus();}
    function iniciarCamara(){reader.style.display='block'; html5QrCode=new Html5Qrcode('reader'); html5QrCode.start({facingMode:'environment'},{fps:10,qrbox:250},decoded=>{dniInput.value=limpiarDni(decoded); buscarDni('QR/CAMARA'); detenerCamara();}).catch(()=>mostrar('No se pudo activar cámara.','warning'));}
    function detenerCamara(){if(html5QrCode){html5QrCode.stop().catch(()=>{});html5QrCode=null;} reader.style.display='none';}
    </script>
    """
    return render_page(body, active='tareo')

# ========================= API =========================
@app.route('/api/trabajador/<dni>')
@login_required
def api_trabajador(dni):
    dni = limpiar_dni(dni)
    if len(dni) != 8: return jsonify(ok=False, msg='DNI inválido.')
    t = row_to_dict(execute('SELECT * FROM trabajadores WHERE dni=?', (dni,), fetchone=True))
    if not t: return jsonify(ok=False, msg='DNI no encontrado en la base de trabajadores.')
    ultima = row_to_dict(execute('SELECT tipo FROM asistencia WHERE dni=? AND fecha=? ORDER BY fecha_hora DESC LIMIT 1', (dni, today_str()), fetchone=True))
    sugerido = 'SALIDA' if ultima and ultima.get('tipo') == 'ENTRADA' else 'ENTRADA'
    return jsonify(ok=True, trabajador=t, sugerido=sugerido)

@app.route('/api/marcacion', methods=['POST'])
@login_required
def api_marcacion():
    p = request.get_json(force=True)
    dni = limpiar_dni(p.get('dni'))
    tipo = limpiar_texto(p.get('tipo') or 'ENTRADA')
    if tipo not in ('ENTRADA','SALIDA'): return jsonify(ok=False, msg='Tipo de marca inválido.')
    t = row_to_dict(execute('SELECT * FROM trabajadores WHERE dni=?', (dni,), fetchone=True))
    if not t: return jsonify(ok=False, msg='Trabajador no encontrado.')
    fecha_hora = now_str(); fecha, hora = fecha_hora[:10], fecha_hora[11:]
    foto_path = guardar_foto_base64(p.get('foto_data'), dni, tipo)
    execute('''INSERT INTO asistencia(dni,trabajador,empresa,area,cargo,tipo,fecha,hora,fecha_hora,metodo,foto_path,latitud,longitud,registrado_por,observacion) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',
            (dni,t.get('trabajador',''),t.get('empresa',''),t.get('area',''),t.get('cargo',''),tipo,fecha,hora,fecha_hora,limpiar_texto(p.get('metodo')),foto_path,str(p.get('latitud','')),str(p.get('longitud','')),session.get('usuario'),limpiar_texto(p.get('observacion'))), commit=True)
    return jsonify(ok=True, msg=f'Marcación {tipo} guardada correctamente para {t.get("trabajador", dni)}.')

@app.route('/api/tareo', methods=['POST'])
@login_required
def api_tareo():
    p = request.get_json(force=True)
    dni = limpiar_dni(p.get('dni'))
    t = row_to_dict(execute('SELECT * FROM trabajadores WHERE dni=?', (dni,), fetchone=True))
    if not t: return jsonify(ok=False, msg='Trabajador no encontrado.')
    fecha = p.get('fecha') or today_str()
    try: horas = float(p.get('horas') or 0); cantidad = float(p.get('cantidad') or 0)
    except Exception: return jsonify(ok=False, msg='Horas/cantidad inválidas.')
    if horas < 0 or horas > 24: return jsonify(ok=False, msg='Horas debe estar entre 0 y 24.')
    execute('''INSERT INTO tareos(dni,trabajador,empresa,area,cargo,fecha,labor,lote,fundo,horas,cantidad,unidad,observacion,registrado_por,creado_en) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',
            (dni,t.get('trabajador',''),t.get('empresa',''),t.get('area',''),t.get('cargo',''),fecha,limpiar_texto(p.get('labor')),limpiar_texto(p.get('lote')),limpiar_texto(p.get('fundo')),horas,cantidad,limpiar_texto(p.get('unidad')),limpiar_texto(p.get('observacion')),session.get('usuario'),now_str()), commit=True)
    return jsonify(ok=True, msg='Tareo guardado correctamente.')

# ========================= CARGA Y REPORTES =========================
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
                if col in headers and item.get(col) not in (None,''):
                    data[key]=limpiar_texto(item.get(col))
            cur.execute(qmark('SELECT id FROM trabajadores WHERE dni=?'), (dni,))
            if cur.fetchone():
                cur.execute(qmark('UPDATE trabajadores SET trabajador=?,empresa=?,area=?,cargo=?,actividad=?,planilla=?,estado=?,fecha_carga=? WHERE dni=?'), (data['trabajador'],data['empresa'],data['area'],data['cargo'],data['actividad'],data['planilla'],data['estado'],ahora,dni)); upd+=1
            else:
                cur.execute(qmark('INSERT INTO trabajadores(dni,trabajador,empresa,area,cargo,actividad,planilla,estado,fecha_carga) VALUES(?,?,?,?,?,?,?,?,?)'), (dni,data['trabajador'],data['empresa'],data['area'],data['cargo'],data['actividad'],data['planilla'],data['estado'],ahora)); ins+=1
        conn.commit(); cur.close(); conn.close()
        flash(f'Carga completa. Insertados: {ins} | Actualizados: {upd} | Omitidos: {omi}', 'success')
        return redirect(url_for('cargar_base'))
    q = request.args.get('q','').strip()
    where=''; params=[]
    if q:
        like=f"%{q.upper()}%"; where='WHERE dni LIKE ? OR UPPER(trabajador) LIKE ? OR UPPER(empresa) LIKE ?'; params=[like,like,like]
    trabajadores = rows_to_dict(execute(f'SELECT * FROM trabajadores {where} ORDER BY fecha_carga DESC, trabajador LIMIT 300', params, fetchall=True))
    body = """
    <div class="topbar d-flex justify-content-between align-items-center flex-wrap gap-2"><div><h2 class="fw-bold mb-0">Cargar base de trabajadores</h2><div class="text-muted">Carga incremental por DNI. Si existe, actualiza; si no existe, inserta.</div></div><a class="btn btn-outline-success btn-pro" href="{{url_for('plantilla_trabajadores')}}">Plantilla Excel</a></div>
    <div class="card card-pro p-4 mb-3"><form method="post" enctype="multipart/form-data"><label class="fw-bold">Archivo Excel .xlsx</label><input class="form-control mb-3" type="file" name="archivo" accept=".xlsx" required><button class="btn btn-success btn-pro">Cargar / actualizar</button></form></div>
    <div class="card card-pro p-3"><div class="d-flex justify-content-between flex-wrap gap-2"><h5 class="fw-bold mb-0">Trabajadores cargados</h5><form class="d-flex gap-2"><input class="form-control" name="q" value="{{q}}" placeholder="Buscar DNI, trabajador o empresa"><button class="btn btn-dark btn-pro"><i class="bi bi-search"></i></button></form></div><div class="table-responsive mt-3"><table class="table table-hover"><thead><tr><th>DNI</th><th>Trabajador</th><th>Empresa</th><th>Área</th><th>Cargo</th><th>Actividad</th><th>Estado</th></tr></thead><tbody>{% for r in trabajadores %}<tr><td class="fw-bold">{{r.dni}}</td><td>{{r.trabajador}}</td><td>{{r.empresa}}</td><td>{{r.area}}</td><td>{{r.cargo}}</td><td>{{r.actividad}}</td><td><span class="badge bg-success">{{r.estado}}</span></td></tr>{% else %}<tr><td colspan="7" class="text-center text-muted py-4">Sin base cargada.</td></tr>{% endfor %}</tbody></table></div></div>
    """
    return render_page(body, active='carga', trabajadores=trabajadores, q=q)

@app.route('/reportes')
@login_required
def reportes():
    desde = request.args.get('desde') or today_str(); hasta = request.args.get('hasta') or today_str(); q=request.args.get('q','').strip()
    params=[desde,hasta]; where='WHERE fecha>=? AND fecha<=?'
    if q:
        like=f"%{q.upper()}%"; where += ' AND (dni LIKE ? OR UPPER(trabajador) LIKE ? OR UPPER(empresa) LIKE ?)'; params += [like,like,like]
    asistencia = rows_to_dict(execute(f'SELECT * FROM asistencia {where} ORDER BY fecha_hora DESC LIMIT 500', params, fetchall=True))
    tareos = rows_to_dict(execute(f'SELECT * FROM tareos {where} ORDER BY creado_en DESC LIMIT 500', params, fetchall=True))
    body = """
    <div class="topbar d-flex justify-content-between align-items-center flex-wrap gap-2"><div><h2 class="fw-bold mb-0">Reportes</h2><div class="text-muted">Asistencia y tareos con filtros por fecha.</div></div><div class="d-flex gap-2 flex-wrap"><a class="btn btn-outline-success btn-pro" href="{{url_for('exportar_asistencia',desde=desde,hasta=hasta,q=q)}}">Exportar asistencia</a><a class="btn btn-success btn-pro" href="{{url_for('exportar_tareos',desde=desde,hasta=hasta,q=q)}}">Exportar tareos</a></div></div>
    <form class="card card-pro p-3 mb-3"><div class="row g-2"><div class="col-md-3"><input class="form-control" type="date" name="desde" value="{{desde}}"></div><div class="col-md-3"><input class="form-control" type="date" name="hasta" value="{{hasta}}"></div><div class="col-md-5"><input class="form-control" name="q" value="{{q}}" placeholder="DNI, trabajador o empresa"></div><div class="col-md-1 d-grid"><button class="btn btn-dark btn-pro"><i class="bi bi-search"></i></button></div></div></form>
    <div class="card card-pro p-3 mb-3"><h5 class="fw-bold">Asistencia</h5><div class="table-responsive"><table class="table table-hover"><thead><tr><th>Fecha</th><th>Hora</th><th>DNI</th><th>Trabajador</th><th>Tipo</th><th>Empresa</th><th>Área</th><th>Método</th><th>Usuario</th></tr></thead><tbody>{% for r in asistencia %}<tr><td>{{r.fecha}}</td><td>{{r.hora}}</td><td class="fw-bold">{{r.dni}}</td><td>{{r.trabajador}}</td><td>{{r.tipo}}</td><td>{{r.empresa}}</td><td>{{r.area}}</td><td>{{r.metodo}}</td><td>{{r.registrado_por}}</td></tr>{% else %}<tr><td colspan="9" class="text-center text-muted py-4">Sin asistencia.</td></tr>{% endfor %}</tbody></table></div></div>
    <div class="card card-pro p-3"><h5 class="fw-bold">Tareos</h5><div class="table-responsive"><table class="table table-hover"><thead><tr><th>Fecha</th><th>DNI</th><th>Trabajador</th><th>Labor</th><th>Fundo</th><th>Lote</th><th>Horas</th><th>Cantidad</th><th>Unidad</th><th>Usuario</th></tr></thead><tbody>{% for r in tareos %}<tr><td>{{r.fecha}}</td><td class="fw-bold">{{r.dni}}</td><td>{{r.trabajador}}</td><td>{{r.labor}}</td><td>{{r.fundo}}</td><td>{{r.lote}}</td><td>{{r.horas}}</td><td>{{r.cantidad}}</td><td>{{r.unidad}}</td><td>{{r.registrado_por}}</td></tr>{% else %}<tr><td colspan="10" class="text-center text-muted py-4">Sin tareos.</td></tr>{% endfor %}</tbody></table></div></div>
    """
    return render_page(body, active='reportes', desde=desde, hasta=hasta, q=q, asistencia=asistencia, tareos=tareos)

@app.route('/exportar/asistencia')
@login_required
def exportar_asistencia():
    desde=request.args.get('desde') or today_str(); hasta=request.args.get('hasta') or today_str(); q=request.args.get('q','').strip()
    params=[desde,hasta]; where='WHERE fecha>=? AND fecha<=?'
    if q:
        like=f"%{q.upper()}%"; where += ' AND (dni LIKE ? OR UPPER(trabajador) LIKE ? OR UPPER(empresa) LIKE ?)'; params += [like,like,like]
    rows=rows_to_dict(execute(f'SELECT fecha,hora,dni,trabajador,empresa,area,cargo,tipo,metodo,latitud,longitud,registrado_por,observacion FROM asistencia {where} ORDER BY fecha_hora DESC', params, fetchall=True))
    headers=['FECHA','HORA','DNI','TRABAJADOR','EMPRESA','AREA','CARGO','TIPO','METODO','LATITUD','LONGITUD','REGISTRADO_POR','OBSERVACION']
    return excel_response(headers, rows, f'asistencia_{desde}_a_{hasta}.xlsx', 'ASISTENCIA')

@app.route('/exportar/tareos')
@login_required
def exportar_tareos():
    desde=request.args.get('desde') or today_str(); hasta=request.args.get('hasta') or today_str(); q=request.args.get('q','').strip()
    params=[desde,hasta]; where='WHERE fecha>=? AND fecha<=?'
    if q:
        like=f"%{q.upper()}%"; where += ' AND (dni LIKE ? OR UPPER(trabajador) LIKE ? OR UPPER(empresa) LIKE ?)'; params += [like,like,like]
    rows=rows_to_dict(execute(f'SELECT fecha,dni,trabajador,empresa,area,cargo,labor,fundo,lote,horas,cantidad,unidad,observacion,registrado_por,creado_en FROM tareos {where} ORDER BY creado_en DESC', params, fetchall=True))
    headers=['FECHA','DNI','TRABAJADOR','EMPRESA','AREA','CARGO','LABOR','FUNDO','LOTE','HORAS','CANTIDAD','UNIDAD','OBSERVACION','REGISTRADO_POR','CREADO_EN']
    return excel_response(headers, rows, f'tareos_{desde}_a_{hasta}.xlsx', 'TAREOS')

@app.route('/plantilla-trabajadores')
@admin_required
def plantilla_trabajadores():
    headers=['DNI','TRABAJADOR','EMPRESA','AREA','CARGO','ACTIVIDAD','PLANILLA','ESTADO']
    rows=[{'DNI':'12345678','TRABAJADOR':'APELLIDOS Y NOMBRES','EMPRESA':'AQUANQA I','AREA':'CAMPO','CARGO':'OPERARIO','ACTIVIDAD':'COSECHA','PLANILLA':'AGRARIO','ESTADO':'ACTIVO'}]
    return excel_response(headers, rows, 'plantilla_trabajadores_asistencia_tareo.xlsx', 'TRABAJADORES')

# ========================= USUARIOS =========================
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
    body="""
    <div class="topbar"><h2 class="fw-bold mb-0">Usuarios</h2><div class="text-muted">Administradores y operadores.</div></div><div class="row g-3"><div class="col-lg-5"><div class="card card-pro p-4"><h5 class="fw-bold">Crear usuario</h5><form method="post"><label class="fw-bold">Usuario</label><input class="form-control mb-2" name="usuario" required><label class="fw-bold">Nombres</label><input class="form-control mb-2" name="nombres"><label class="fw-bold">Clave</label><input class="form-control mb-2" type="password" name="clave" required><label class="fw-bold">Rol</label><select class="form-select mb-2" name="rol"><option value="operador">operador</option><option value="admin">admin</option></select><label class="fw-bold">Estado</label><select class="form-select mb-3" name="estado"><option>ACTIVO</option><option>INACTIVO</option></select><button class="btn btn-success btn-pro">Guardar</button></form></div></div><div class="col-lg-7"><div class="card card-pro p-3"><h5 class="fw-bold">Lista</h5><div class="table-responsive"><table class="table table-hover"><thead><tr><th>Usuario</th><th>Nombres</th><th>Rol</th><th>Estado</th><th>Creado</th></tr></thead><tbody>{% for u in users %}<tr><td class="fw-bold">{{u.usuario}}</td><td>{{u.nombres}}</td><td>{{u.rol}}</td><td>{{u.estado}}</td><td>{{u.creado_en}}</td></tr>{% endfor %}</tbody></table></div></div></div></div>
    """
    return render_page(body, active='usuarios', users=users)

# ========================= PWA =========================
@app.route('/manifest.webmanifest')
@app.route('/manifest.json')
def manifest():
    return jsonify({"name":"Asistencia y Tareo PRIZE","short_name":"Asistencia","start_url":"/","display":"standalone","background_color":"#eef7f1","theme_color":"#16a34a","icons":[]})

@app.route('/sw.js')
def sw():
    return Response("self.addEventListener('install',e=>self.skipWaiting()); self.addEventListener('fetch',e=>{});", mimetype='application/javascript')

if __name__ == '__main__':
    init_db()
    port = int(os.getenv('PORT', '5000'))
    app.run(host='0.0.0.0', port=port, debug=False)
