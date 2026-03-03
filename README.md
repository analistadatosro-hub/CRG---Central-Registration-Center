# CRG – Central Registration Center v2.0

Sistema de registro centralizado de tickets TEMIS para Sodexo.

---

## Archivos del proyecto

```
crg-form/
├── server.js          ← Servidor Node.js (lógica + conexión SharePoint)
├── package.json       ← Dependencias
├── .gitignore         ← Excluye node_modules y .env
└── public/
    └── index.html     ← Formulario web (login + form)
```

---

## Variables de entorno (configurar en Render)

| Variable            | Descripción |
|---------------------|-------------|
| `APP_USER`          | Usuario de acceso al formulario (`Usuario123`) |
| `APP_PASS`          | Contraseña de acceso (`Contraseña2026`) |
| `SESSION_SECRET`    | Texto secreto para las sesiones (cualquier texto largo) |
| `MS_TENANT_ID`      | Tenant ID de Sodexo en Azure AD |
| `MS_CLIENT_ID`      | Client ID de la app registrada en Azure |
| `MS_CLIENT_SECRET`  | Client Secret de la app en Azure |
| `SP_SITE_ID`        | ID del sitio SharePoint |
| `SP_BD_CECO_ID`     | ID del archivo Bd Ceco.xlsx en SharePoint |
| `SP_BD_ESTADO_ID`   | ID del archivo Bd Estado cliente.xlsx |
| `SP_BD_FAMILIA_ID`  | ID del archivo Bd Familia y Sub familia.xlsx |
| `SP_BD_RESP_ID`     | ID del archivo Bd Responsable.xlsx |
| `SP_BD_TICKETS_ID`  | ID del archivo Registro web de tickets.xlsx |
| `SP_TICKETS_SHEET`  | Nombre de la hoja (por defecto: `Hoja1`) |

---

## PASO 1 — Registrar app en Azure AD (con permisos de aplicación)

> ⚠️ Esta app usa **Client Credentials** (credenciales fijas tuyas), NO login de usuarios.

1. Ir a **portal.azure.com** → **Azure Active Directory** → **App registrations** → **New registration**
2. Nombre: `CRG Temis`
3. Supported account types: **Accounts in this organizational directory only**
4. Redirect URI: dejar vacío (no se necesita)
5. Clic en **Register**
6. Copiar **Application (client) ID** → `MS_CLIENT_ID`
7. Copiar **Directory (tenant) ID** → `MS_TENANT_ID`

### Crear Client Secret:
1. En la app → **Certificates & secrets** → **New client secret**
2. Descripción: `crg-temis`, Expira: 24 meses
3. Copiar el **Value** inmediatamente → `MS_CLIENT_SECRET`

### Permisos de aplicación (no delegados):
1. **API permissions** → **Add a permission** → **Microsoft Graph** → **Application permissions**
2. Agregar: `Files.ReadWrite.All`, `Sites.ReadWrite.All`
3. Clic en **Grant admin consent** ← requiere ser admin del tenant

---

## PASO 2 — Obtener IDs de SharePoint

### Site ID:
En Graph Explorer (https://developer.microsoft.com/graph/graph-explorer):
```
GET https://graph.microsoft.com/v1.0/sites/sodexo.sharepoint.com:/sites/Basededatos-CommandCenter
```
Copiar el campo `id` → `SP_SITE_ID`

### ID de cada archivo Excel:
```
GET https://graph.microsoft.com/v1.0/sites/{SP_SITE_ID}/drive/root:/Documents partages/General/Banca/Automatizaciones/Codigos/Hub Tickets/Bd_nomenclaturas/Bd Ceco.xlsx
```
Copiar `id` → `SP_BD_CECO_ID`

Repetir para cada archivo cambiando el nombre al final de la URL.

---

## PASO 3 — Subir a GitHub y Render

### GitHub:
1. Subir estos archivos al repo: `server.js`, `package.json`, `.gitignore`, `public/index.html`
2. **NO** subir `node_modules/` ni `.env`

### Render:
1. **New +** → **Web Service** (no Static Site)
2. Conectar repo de GitHub
3. Configurar:
   - **Build Command:** `npm install`
   - **Start Command:** `npm start`
   - **Environment:** Node
4. En **Environment Variables** agregar todas las variables de la tabla de arriba
5. Deploy

---

## Estructura esperada de los Excel en SharePoint

### Bd Ceco.xlsx
| A: Cliente | B: Ceco |
|---|---|
| BCP | LIMA1-REGION1-SANTA ANITA |

### Bd Estado cliente.xlsx
| A: Cliente | B: Estado |
|---|---|
| BCP | Cerrado |

### Bd Familia y Sub familia.xlsx
| A: Cliente | B: Familia | C: Sub Familia |
|---|---|---|
| BCP | EQUIPO | ATM |

### Bd Responsable.xlsx
| A: Nombre | B: Correo |
|---|---|
| Juan Perez | juan.perez@sodexo.com |

### Registro web de tickets.xlsx — Columnas en orden:
`W.O | Ceco | Familia | Sub Familia | Descripción OT | Usuario creó OT | Correo | Detalle | Tipo OT | Tipo Ticket | Estado | Fecha Modif Estado | Modificado por | Fecha Apertura | Fecha Inicio Real | Prioridad | Cliente`
