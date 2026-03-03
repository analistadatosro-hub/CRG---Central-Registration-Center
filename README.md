# CRG – Central Registration Center v2.0 (Python/Flask)

---

## Archivos del proyecto

```
crg-form/
├── app.py             ← Servidor Python/Flask
├── requirements.txt   ← Librerías Python
├── render.yaml        ← Configuración de Render
├── .gitignore
└── public/
    ├── index.html     ← Formulario web
    └── sodexo_logo.png
```

---

## Variables de entorno (configurar en Render)

| Variable           | Descripción | Ejemplo |
|--------------------|-------------|---------|
| `APP_USER`         | Usuario para entrar al formulario | `Usuario123` |
| `APP_PASS`         | Contraseña para entrar al formulario | `Contraseña2026` |
| `SESSION_SECRET`   | Texto secreto para sesiones | `cualquier-texto-largo` |
| `SP_USUARIO`       | Tu email corporativo de Sodexo | `rodrigo.obando@sodexo.com` |
| `SP_PASSWORD`      | Tu contraseña de Sodexo | `tucontraseña` |
| `SITE_COMMAND`     | URL del site SharePoint de las BDs | `https://sodexo.sharepoint.com/sites/Basededatos-CommandCenter` |
| `RUTA_BD_CECO`     | Ruta del archivo Bd Ceco.xlsx | `/sites/Basededatos-CommandCenter/Documents partages/General/Banca/Automatizaciones/Codigos/Hub Tickets/Bd_nomenclaturas/Bd Ceco.xlsx` |
| `RUTA_BD_ESTADO`   | Ruta del archivo Bd Estado cliente.xlsx | igual pero con el nombre del archivo |
| `RUTA_BD_FAMILIA`  | Ruta del archivo Bd Familia y Sub familia.xlsx | igual |
| `RUTA_BD_RESP`     | Ruta del archivo Bd Responsable.xlsx | igual |
| `SITE_TICKETS`     | URL del site donde está el archivo de tickets | `https://sodexo.sharepoint.com/sites/Basededatos-CommandCenter` |
| `RUTA_BD_TICKETS`  | Ruta del archivo Registro web de tickets.xlsx | `/sites/Basededatos-CommandCenter/Documents partages/General/Banca/Automatizaciones/Codigos/Hub Tickets/Registro web de tickets.xlsx` |
| `SHEET_TICKETS`    | Nombre de la hoja en el Excel de tickets | `Hoja1` |

---

## Pasos para subir a Render

1. Subir todos los archivos a GitHub (reemplazar los anteriores)
2. En Render → tu servicio actual → **Settings**
3. Cambiar **Language** de Node a **Python**
4. **Build Command:** `pip install -r requirements.txt`
5. **Start Command:** `gunicorn app:app`
6. En **Environment Variables** agregar todas las variables de la tabla
7. **Manual Deploy** → Deploy latest commit

---

## Estructura esperada de los Excel en SharePoint

### Bd Ceco.xlsx
| A: Cliente | B: Ceco |
|---|---|
| BCP | LIMA1-REGION1-SANTA ANITA |
| BBVA | Sede Central BBVA |

### Bd Estado cliente.xlsx
| A: Cliente | B: Estado |
|---|---|
| BCP | Cerrado |
| BBVA | VISA |

### Bd Familia y Sub familia.xlsx
| A: Cliente | B: Familia | C: Sub Familia |
|---|---|---|
| BCP | EQUIPO | ATM |
| BCP | EQUIPO | AIRE ACONDICIONADO |

### Bd Responsable.xlsx
| A: Nombre | B: Correo |
|---|---|
| Juan Perez | juan.perez@sodexo.com |

### Registro web de tickets.xlsx — columnas en orden exacto:
`W.O | Ceco | Familia | Sub Familia | Descripción OT | Usuario creó OT | Correo | Detalle | Tipo OT | Tipo Ticket | Estado | Fecha Modif Estado | Modificado por | Fecha Apertura | Fecha Inicio Real | Prioridad | Cliente`
