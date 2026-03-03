# CRG – Central Registration Center
**Sistema de registro centralizado de tickets TEMIS para Sodexo**

---

## ¿Qué hace esta app?
- Login con cuenta corporativa Microsoft (@sodexo.com)
- Lee los desplegables (Ceco, Familia, Sub Familia, Descripción OT, Responsables) directamente desde los Excel en SharePoint
- Al guardar un ticket, escribe la nueva fila en el Excel de registro en SharePoint

---

## Pasos para ponerlo en funcionamiento

### PASO 1 — Registrar la app en Azure AD

1. Ir a **portal.azure.com** → **Azure Active Directory** → **App registrations** → **New registration**
2. Nombre: `CRG Temis`
3. Supported account types: **Accounts in this organizational directory only (Sodexo only)**
4. Redirect URI: tipo **SPA** → URL: `https://TU-APP.onrender.com` (la URL que te dará Render)
5. Hacer clic en **Register**
6. Copiar el **Application (client) ID** → este es tu `clientId`
7. Copiar el **Directory (tenant) ID** → este es tu `tenantId`

#### Permisos necesarios:
1. En la app → **API permissions** → **Add a permission** → **Microsoft Graph** → **Delegated permissions**
2. Agregar: `Files.ReadWrite.All`, `Sites.ReadWrite.All`, `User.Read`
3. Hacer clic en **Grant admin consent**

---

### PASO 2 — Obtener los IDs de SharePoint (Graph Explorer)

1. Ir a **https://developer.microsoft.com/en-us/graph/graph-explorer**
2. Iniciar sesión con tu cuenta de Sodexo

#### Obtener el Site ID:
```
GET https://graph.microsoft.com/v1.0/sites/sodexo.sharepoint.com:/sites/Basededatos-CommandCenter
```
Copiar el valor `id` del resultado → este es tu `siteId`

#### Obtener el ID de cada archivo Excel:
```
GET https://graph.microsoft.com/v1.0/sites/{siteId}/drive/root:/Documents partages/General/Banca/Automatizaciones/Codigos/Hub Tickets/Bd_nomenclaturas/Bd Ceco.xlsx
```
Copiar el valor `id` → este es tu `bdCecoId`

Repetir para cada archivo:
- `Bd Estado cliente.xlsx` → `bdEstadoId`
- `Bd Familia y Sub familia.xlsx` → `bdFamiliaId`
- `Bd Responsable.xlsx` → `bdRespId`
- `Registro web de tickets.xlsx` → `bdTicketsId`

---

### PASO 3 — Configurar el archivo index.html

Abrir `index.html` y completar el objeto `CFG` al inicio del script:

```javascript
const CFG = {
  clientId:   "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",  // App ID de Azure
  tenantId:   "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",  // Tenant ID de Sodexo
  redirectUri: window.location.origin,

  siteId:         "sodexo.sharepoint.com,xxxxxx,xxxxxx",
  bdCecoId:       "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx",
  bdEstadoId:     "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx",
  bdFamiliaId:    "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx",
  bdRespId:       "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx",
  bdTicketsId:    "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx",
  ticketsSheet:   "Hoja1"   // Nombre exacto de la hoja en el Excel de tickets
};
```

---

### PASO 4 — Estructura esperada de los Excel en SharePoint

#### Bd Ceco.xlsx
| A (Cliente) | B (Ceco) |
|---|---|
| BCP | LIMA1-REGION1-SANTA ANITA |
| BBVA | Sede Central BBVA |

#### Bd Estado cliente.xlsx
| A (Cliente) | B (Estado) |
|---|---|
| BCP | Cerrado |
| BBVA | VISA |

#### Bd Familia y Sub familia.xlsx
| A (Cliente) | B (Familia) | C (Sub Familia) |
|---|---|---|
| BCP | EQUIPO | ATM |
| BCP | EQUIPO | AIRE ACONDICIONADO |
| BCP | MOBILIARIO | Mueble |

#### Bd Responsable.xlsx
| A (Nombre) | B (Correo) |
|---|---|
| Juan Perez | juan.perez@sodexo.com |

#### Registro web de tickets.xlsx (columnas en orden exacto)
| W.O | Ceco | Familia | Sub Familia | Descripción OT | Usuario creó OT | Correo | Detalle | Tipo OT | Tipo Ticket | Estado | Fecha Modif Estado | Modificado por | Fecha Apertura | Fecha Inicio Real | Prioridad | Cliente |

---

### PASO 5 — Subir a GitHub y Render

1. Crear repositorio en GitHub y subir `index.html`
2. En **render.com** → New → **Static Site**
3. Conectar el repo de GitHub
4. Build command: *(vacío)*
5. Publish directory: `.`
6. Hacer deploy

**Importante:** Después de obtener la URL de Render (ej: `https://crg-temis.onrender.com`), volver a Azure AD → App registrations → tu app → Authentication → agregar esa URL exacta como Redirect URI (tipo SPA).

---

## Notas
- Si SharePoint no responde, la app usa datos embebidos de BCP como fallback
- Los campos predeterminados (Tipo Ticket=RM, Estado=ACK) se envían automáticamente
- La fecha de apertura la elige el usuario; la fecha de inicio real se toma automáticamente del momento del envío
