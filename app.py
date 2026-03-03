import os, io, threading
from datetime import datetime
from zoneinfo import ZoneInfo
from flask import Flask, request, jsonify, session, send_from_directory
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.files.file import File
from openpyxl import load_workbook
from functools import wraps

write_lock = threading.Lock()
app = Flask(__name__, static_folder='public', static_url_path='')

# ═══════════════════════════════════════════
#  CREDENCIALES — solo estas van en Render
# ═══════════════════════════════════════════
APP_USER    = os.environ.get('APP_USER',    'Usuario123')
APP_PASS    = os.environ.get('APP_PASS',    'Contraseña2026')
SECRET_KEY  = os.environ.get('SESSION_SECRET', 'crg-sodexo-2026')
SP_USUARIO  = os.environ.get('SP_USUARIO')
SP_PASSWORD = os.environ.get('SP_PASSWORD')
app.secret_key = SECRET_KEY

def build_users():
    users = {APP_USER: APP_PASS}
    for par in os.environ.get('USERS_EXTRA', '').split(','):
        if ':' in par:
            u, p = par.strip().split(':', 1)
            users[u.strip()] = p.strip()
    return users

# ═══════════════════════════════════════════
#  SHAREPOINT — rutas fijas
# ═══════════════════════════════════════════
SITE_CMD     = "https://sodexo.sharepoint.com/sites/Basededatos-CommandCenter"
BASE_NOMEN   = "/sites/Basededatos-CommandCenter/Documents partages/General/Banca/Automatizaciones/Codigos/Hub Tickets/Bd_nomenclaturas"
RUTA_CECO    = f"{BASE_NOMEN}/Bd Ceco.xlsx"
RUTA_ESTADO  = f"{BASE_NOMEN}/Bd Estado cliente.xlsx"
RUTA_FAMILIA = f"{BASE_NOMEN}/Bd Familia y Sub familia.xlsx"
RUTA_RESP    = f"{BASE_NOMEN}/Bd Responsable.xlsx"
SITE_TKT     = "https://sodexo.sharepoint.com/sites/Basededatos-CommandCenter"
RUTA_TKT     = "/sites/Basededatos-CommandCenter/Documents partages/General/Banca/Base de Datos/BD Temis/Registro web de tickets.xlsx"
SHEET_TKT    = "Hoja1"

# ═══════════════════════════════════════════
#  HELPERS
# ═══════════════════════════════════════════
def get_ctx(site):
    return ClientContext(site).with_credentials(UserCredential(SP_USUARIO, SP_PASSWORD))

def read_sp(site, path):
    import openpyxl
    buf = io.BytesIO(File.open_binary(get_ctx(site), path).content)
    ws  = openpyxl.load_workbook(buf, data_only=True).worksheets[0]
    return [list(r) for r in ws.iter_rows(values_only=True)]

def write_sp(site, path, sheet, row_data):
    buf = io.BytesIO(File.open_binary(get_ctx(site), path).content)
    wb  = load_workbook(buf)
    ws  = wb[sheet] if sheet in wb.sheetnames else wb.worksheets[0]
    ws.append(row_data)
    nr  = ws.max_row
    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    folder = get_ctx(site).web.get_folder_by_server_relative_url('/'.join(path.split('/')[:-1]))
    folder.upload_file(path.split('/')[-1], out.getvalue()).execute_query()
    return nr

def fmt_date(s):
    try: return datetime.strptime(s, '%Y-%m-%d').strftime('%d/%m/%Y 00:00:00')
    except: return s or ''

def fmt_dt(d): return d.strftime('%d/%m/%Y %H:%M:%S')

def s(v): return str(v or '').strip()

def auth_required(f):
    @wraps(f)
    def dec(*a, **k):
        if not session.get('authed'):
            return jsonify({'error': 'No autorizado'}), 401
        return f(*a, **k)
    return dec

# ═══════════════════════════════════════════
#  RUTAS
# ═══════════════════════════════════════════
@app.route('/')
def index(): return send_from_directory('public', 'index.html')

@app.route('/api/login', methods=['POST'])
def login():
    d = request.get_json()
    users = build_users()
    u, p = d.get('usuario', ''), d.get('password', '')
    if u in users and users[u] == p:
        session['authed'] = True
        session['user'] = u
        return jsonify({'ok': True})
    return jsonify({'ok': False, 'error': 'Usuario o contraseña incorrectos'}), 401

@app.route('/api/logout', methods=['POST'])
def logout():
    session.clear()
    return jsonify({'ok': True})

@app.route('/api/check')
def check():
    return jsonify({'authed': session.get('authed', False)})

@app.route('/api/data')
@auth_required
def get_data():
    try:
        ceco = {}
        for r in read_sp(SITE_CMD, RUTA_CECO)[1:]:
            c, v = s(r[0]), s(r[1])
            if c and v: ceco.setdefault(c, []).append(v)

        estado = {}
        for r in read_sp(SITE_CMD, RUTA_ESTADO)[1:]:
            c, v = s(r[0]), s(r[1])
            if c and v: estado.setdefault(c, []).append(v)

        familia = {}
        for r in read_sp(SITE_CMD, RUTA_FAMILIA)[1:]:
            c, f = s(r[0]), s(r[1])
            sf = s(r[2]) if len(r) > 2 else ''
            if c and f:
                familia.setdefault(c, {}).setdefault(f, [])
                if sf: familia[c][f].append(sf)

        resp = []
        for r in read_sp(SITE_CMD, RUTA_RESP)[1:]:
            banco  = s(r[0])
            correo = s(r[1]) if len(r) > 1 else ''
            nombre = s(r[2]) if len(r) > 2 else ''
            if nombre: resp.append({'nombre': nombre, 'correo': correo, 'banco': banco})

        return jsonify({'ceco': ceco, 'estado': estado, 'familia': familia, 'responsable': resp})
    except Exception as e:
        print(f'Error data: {e}')
        return jsonify({'error': str(e)}), 500

@app.route('/api/ticket', methods=['POST'])
@auth_required
def save_ticket():
    try:
        d = request.get_json()
        wo_new = s(d.get('wo', '')).upper()

        with write_lock:
            # ── Verificar W.O duplicado ──
            for r in read_sp(SITE_TKT, RUTA_TKT)[1:]:
                if s(r[0]).upper() == wo_new:
                    return jsonify({'ok': False, 'duplicado': True, 'wo': wo_new}), 409

            now = datetime.now(ZoneInfo('America/Lima'))
            row = [
                d.get('wo', ''),           # 0  A - W.O
                d.get('ceco', ''),         # 1  B - Ceco
                d.get('familia', ''),      # 2  C - Familia
                d.get('sub_familia', ''),  # 3  D - Sub Familia
                d.get('descripcion_ot', ''),# 4 E - Descripción OT
                d.get('usuario', ''),      # 5  F - Usuario que creo la OT
                d.get('correo', ''),       # 6  G - Correo Usuario
                d.get('detalle', ''),      # 7  H - Detalle
                'RM',                      # 8  I - Tipo de OT
                d.get('tipo_ot', ''),      # 9  J - Tipo de Ticket (CBM/CBP)
                'ACK',                     # 10 K - Estado
                fmt_date(d.get('fecha_apertura')),  # 11 L - Fecha modif. estado
                d.get('usuario', ''),      # 12 M - Modificado por
                fmt_date(d.get('fecha_apertura')),  # 13 N - Fecha de Apertura Cliente
                fmt_dt(now),               # 14 O - Fecha de Inicio Real (hora Lima)
                '',                        # 15 P - Fecha de cierre Real (vacío)
                d.get('prioridad', ''),    # 16 Q - Prioridad
                d.get('cliente', '')       # 17 R - Cliente
            ]
            nr = write_sp(SITE_TKT, RUTA_TKT, SHEET_TKT, row)

        return jsonify({'ok': True, 'fila': nr})
    except Exception as e:
        print(f'Error ticket: {e}')
        return jsonify({'error': str(e)}), 500

@app.route('/api/control')
@auth_required
def get_control():
    try:
        tickets = []
        for r in read_sp(SITE_TKT, RUTA_TKT)[1:]:
            wo = s(r[0])
            if not wo: continue
            tickets.append({
                'wo':          wo,
                'ceco':        s(r[1]),
                'familia':     s(r[2]),
                'descripcion': s(r[4]),
                'usuario':     s(r[5]),
                'correo':      s(r[6]),
                'fecha':       s(r[13]),
                'prioridad':   s(r[16]),
                'cliente':     s(r[17]),
                'estado_reg':  s(r[27]) if len(r) > 27 else ''
            })
        tickets.reverse()  # más recientes primero
        return jsonify({'tickets': tickets})
    except Exception as e:
        print(f'Error control: {e}')
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 3000)))
