const express = require('express');
const path    = require('path');
const session = require('express-session');
const fetch   = (...a) => import('node-fetch').then(({default:f}) => f(...a));

const app  = express();
const PORT = process.env.PORT || 3000;

// ── Middlewares ──
app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(session({
  secret: process.env.SESSION_SECRET || 'crg-sodexo-secret-2026',
  resave: false,
  saveUninitialized: false,
  cookie: { secure: false, maxAge: 8 * 60 * 60 * 1000 } // 8 horas
}));

// ── Credenciales desde variables de entorno ──
const APP_USER = process.env.APP_USER     || 'Usuario123';
const APP_PASS = process.env.APP_PASS     || 'Contraseña2026';

// Credenciales Microsoft (Graph API — Service Account de Sodexo)
const MS_TENANT   = process.env.MS_TENANT_ID;
const MS_CLIENT   = process.env.MS_CLIENT_ID;
const MS_SECRET   = process.env.MS_CLIENT_SECRET;

// IDs de SharePoint
const SP_SITE         = process.env.SP_SITE_ID;
const SP_BD_CECO      = process.env.SP_BD_CECO_ID;
const SP_BD_ESTADO    = process.env.SP_BD_ESTADO_ID;
const SP_BD_FAMILIA   = process.env.SP_BD_FAMILIA_ID;
const SP_BD_RESP      = process.env.SP_BD_RESP_ID;
const SP_BD_TICKETS   = process.env.SP_BD_TICKETS_ID;
const SP_SHEET        = process.env.SP_TICKETS_SHEET || 'Hoja1';

// ══════════════════════════════════════════════
//  MICROSOFT GRAPH — Token con Client Credentials
// ══════════════════════════════════════════════
let _token = null;
let _tokenExpiry = 0;

async function getGraphToken() {
  if (_token && Date.now() < _tokenExpiry) return _token;

  const url = `https://login.microsoftonline.com/${MS_TENANT}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    grant_type:    'client_credentials',
    client_id:     MS_CLIENT,
    client_secret: MS_SECRET,
    scope:         'https://graph.microsoft.com/.default'
  });

  const r = await fetch(url, { method: 'POST', body });
  const d = await r.json();
  if (!d.access_token) throw new Error('No se pudo obtener token de Microsoft: ' + JSON.stringify(d));

  _token = d.access_token;
  _tokenExpiry = Date.now() + (d.expires_in - 60) * 1000;
  return _token;
}

async function gGet(path) {
  const t = await getGraphToken();
  const r = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
    headers: { Authorization: `Bearer ${t}` }
  });
  if (!r.ok) throw new Error(`Graph GET ${path} → ${r.status}: ${await r.text()}`);
  return r.json();
}

async function gPatch(path, body) {
  const t = await getGraphToken();
  const r = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
    method: 'PATCH',
    headers: { Authorization: `Bearer ${t}`, 'Content-Type': 'application/json' },
    body: JSON.stringify(body)
  });
  if (!r.ok) throw new Error(`Graph PATCH ${path} → ${r.status}: ${await r.text()}`);
  return r.json();
}

async function readSheet(itemId, sheet = 'Hoja1') {
  const d = await gGet(`/sites/${SP_SITE}/drive/items/${itemId}/workbook/worksheets('${sheet}')/usedRange`);
  return d.values || [];
}

// ── Helper: letra de columna ──
function colLetter(n) {
  let s = '';
  while (n > 0) { s = String.fromCharCode(65 + (n - 1) % 26) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

// ── Helper: formato fecha ──
function fmtDate(dateStr) {
  if (!dateStr) return '';
  const [y, m, d] = dateStr.split('-');
  return `${d}/${m}/${y} 00:00:00`;
}
function fmtDateTime(d) {
  return `${String(d.getDate()).padStart(2,'0')}/${String(d.getMonth()+1).padStart(2,'0')}/${d.getFullYear()} ${String(d.getHours()).padStart(2,'0')}:${String(d.getMinutes()).padStart(2,'0')}:${String(d.getSeconds()).padStart(2,'0')}`;
}

// ══════════════════════════════════════════════
//  MIDDLEWARE DE AUTH
// ══════════════════════════════════════════════
function requireAuth(req, res, next) {
  if (req.session && req.session.authed) return next();
  res.status(401).json({ error: 'No autorizado' });
}

// ══════════════════════════════════════════════
//  RUTAS
// ══════════════════════════════════════════════

// Servir el frontend
app.get('/', (req, res) => res.sendFile(path.join(__dirname, 'public', 'index.html')));

// ── Login ──
app.post('/api/login', (req, res) => {
  const { usuario, password } = req.body;
  if (usuario === APP_USER && password === APP_PASS) {
    req.session.authed = true;
    res.json({ ok: true });
  } else {
    res.status(401).json({ ok: false, error: 'Usuario o contraseña incorrectos' });
  }
});

// ── Logout ──
app.post('/api/logout', (req, res) => {
  req.session.destroy();
  res.json({ ok: true });
});

// ── Check sesión ──
app.get('/api/check', (req, res) => {
  res.json({ authed: !!(req.session && req.session.authed) });
});

// ── Cargar datos de los desplegables ──
app.get('/api/data', requireAuth, async (req, res) => {
  try {
    const [rowsCeco, rowsEstado, rowsFam, rowsResp] = await Promise.all([
      readSheet(SP_BD_CECO),
      readSheet(SP_BD_ESTADO),
      readSheet(SP_BD_FAMILIA),
      readSheet(SP_BD_RESP)
    ]);

    // Ceco: A=Cliente, B=Ceco
    const ceco = {};
    rowsCeco.slice(1).forEach(r => {
      const c = (r[0]||'').toString().trim(), v = (r[1]||'').toString().trim();
      if (c && v) { if (!ceco[c]) ceco[c] = []; ceco[c].push(v); }
    });

    // Estado: A=Cliente, B=Estado
    const estado = {};
    rowsEstado.slice(1).forEach(r => {
      const c = (r[0]||'').toString().trim(), v = (r[1]||'').toString().trim();
      if (c && v) { if (!estado[c]) estado[c] = []; estado[c].push(v); }
    });

    // Familia: A=Cliente, B=Familia, C=Sub Familia
    const familia = {};
    rowsFam.slice(1).forEach(r => {
      const c = (r[0]||'').toString().trim(), f = (r[1]||'').toString().trim(), sf = (r[2]||'').toString().trim();
      if (c && f) {
        if (!familia[c]) familia[c] = {};
        if (!familia[c][f]) familia[c][f] = [];
        if (sf) familia[c][f].push(sf);
      }
    });

    // Responsable: A=Nombre, B=Correo
    const responsable = [];
    rowsResp.slice(1).forEach(r => {
      const n = (r[0]||'').toString().trim(), e = (r[1]||'').toString().trim();
      if (n) responsable.push({ nombre: n, correo: e });
    });

    res.json({ ceco, estado, familia, responsable });
  } catch (err) {
    console.error('Error cargando datos:', err.message);
    res.status(500).json({ error: err.message });
  }
});

// ── Guardar ticket ──
app.post('/api/ticket', requireAuth, async (req, res) => {
  try {
    const {
      wo, ceco, familia, sub_familia, descripcion_ot,
      usuario, correo, detalle, tipo_ot, prioridad,
      cliente, fecha_apertura
    } = req.body;

    const now = new Date();

    // Orden exacto de columnas según tabla entregada
    const row = [
      wo,
      ceco,
      familia,
      sub_familia || '',
      descripcion_ot,
      usuario,
      correo,
      detalle,
      tipo_ot,
      'RM',                    // Tipo de Ticket (predeterminado)
      'ACK',                   // Estado (predeterminado)
      fmtDate(fecha_apertura), // Fecha modificación estado = Fecha apertura
      usuario,                 // Modificado por = Usuario que creó OT
      fmtDate(fecha_apertura), // Fecha de Apertura Cliente
      fmtDateTime(now),        // Fecha de Inicio Real = ahora
      prioridad,
      cliente
    ];

    // Obtener fila siguiente
    const used = await gGet(`/sites/${SP_SITE}/drive/items/${SP_BD_TICKETS}/workbook/worksheets('${SP_SHEET}')/usedRange`);
    const nextRow = (used.rowCount || 1) + 1;
    const addr = `A${nextRow}:${colLetter(row.length)}${nextRow}`;

    await gPatch(
      `/sites/${SP_SITE}/drive/items/${SP_BD_TICKETS}/workbook/worksheets('${SP_SHEET}')/range(address='${addr}')`,
      { values: [row] }
    );

    res.json({ ok: true, fila: nextRow });
  } catch (err) {
    console.error('Error guardando ticket:', err.message);
    res.status(500).json({ error: err.message });
  }
});

app.listen(PORT, () => console.log(`CRG corriendo en puerto ${PORT}`));
