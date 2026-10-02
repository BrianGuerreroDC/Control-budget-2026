# Control Budget Compliance 2026

Dashboard estático (`index.html`) + una función serverless (`api/data.js`) que lee
**Cuadro Control 2026.xlsx** directo desde Microsoft Graph con credenciales de
aplicación. Ninguna persona necesita conectar su propia cuenta de Microsoft 365:
el backend se autentica una sola vez por el equipo, y todos los que abran el
enlace ven los mismos datos, actualizados automáticamente cada minuto.

## 0. Opción rápida, sin credenciales de Azure (enlace compartido)

1. En SharePoint/OneDrive: **Compartir → Configuración del vínculo → "Cualquier persona con el vínculo"**, permiso **Solo lectura**.
2. En Vercel → Settings → Environment Variables agrega `EXCEL_SHARE_URL` con ese enlace.
3. Listo: `api/data.js` descarga el Excel con ese enlace (no necesita las secciones 1–3 de abajo).
   Prueba: abre el enlace en una ventana de incógnito; si descarga/abre el archivo sin pedir login, funciona.
   Quien tenga el enlace puede ver el archivo, por eso va solo como variable de entorno.

## 1. (Alternativa) Registrar la app en Azure AD (lo hace tu equipo de TI)

1. **Azure Portal → Microsoft Entra ID → App registrations → New registration.**
   Nombre sugerido: `Monitor Budget Compliance 2026`. Tipo de cuenta: solo tu
   organización.
2. **Certificates & secrets → New client secret.** Copia el *value* apenas se
   genera (no se vuelve a mostrar). Anota también el **Directory (tenant) ID**
   y el **Application (client) ID** de la pantalla "Overview".
3. **API permissions → Add a permission → Microsoft Graph → Application
   permissions** (no "Delegated"). Agrega `Files.Read.All` (o `Sites.Read.All`
   si el archivo vive en un sitio de SharePoint). Luego **Grant admin consent**
   — este paso solo lo puede hacer un administrador del tenant.
4. **(Recomendado) Restringir el acceso solo a ese archivo/sitio**, en vez de
   dejar que la app lea todo OneDrive de la organización, con una
   *Application Access Policy* (`New-ApplicationAccessPolicy` en PowerShell de
   Exchange Online/SharePoint) o limitando el permiso a `Sites.Selected` y
   dando acceso explícito solo al sitio del archivo.

## 2. Ubicar el archivo (`GRAPH_DRIVE_ID` y `GRAPH_ITEM_ID`)

La forma más simple es con **Graph Explorer** (https://developer.microsoft.com/graph/graph-explorer),
con la cuenta dueña del archivo:

- `GET /me/drive/root:/RUTA/AL/ARCHIVO/Cuadro Control 2026.xlsx` (ajusta la
  ruta) → la respuesta trae `"parentReference": {"driveId": "..."}` y
  `"id": "..."` — esos son `GRAPH_DRIVE_ID` y `GRAPH_ITEM_ID`.
- Si el archivo vive en un sitio de SharePoint (no en un OneDrive personal),
  primero `GET /sites/{hostname}:/sites/{sitio}` para obtener el `site id`, y
  luego `GET /sites/{site-id}/drive/root:/RUTA/ARCHIVO.xlsx` para obtener el
  `driveId`/`id` igual que arriba.

## 3. Variables de entorno en Vercel

En el proyecto de Vercel → **Settings → Environment Variables**, agrega
(nunca las escribas en el código ni las compartas por chat):

| Variable              | Valor                                            |
|------------------------|---------------------------------------------------|
| `AZURE_TENANT_ID`      | Directory (tenant) ID                             |
| `AZURE_CLIENT_ID`      | Application (client) ID                           |
| `AZURE_CLIENT_SECRET`  | El *secret value* del paso 1.2                    |
| `GRAPH_DRIVE_ID`       | Del paso 2                                        |
| `GRAPH_ITEM_ID`        | Del paso 2                                        |
| `BUDGET_YEAR`          | `2026` (opcional, ya es el valor por defecto)     |
| `DATA_CACHE_TTL_MS`    | `45000` (opcional; baja este número para refrescar más seguido) |

## 4. Subir a GitHub y desplegar en Vercel

```bash
cd monitor-budget-2026
git init
git add .
git commit -m "Monitor Control Budget Compliance 2026"
git branch -M main
git remote add origin https://github.com/TU-USUARIO/TU-REPO.git
git push -u origin main
```

Luego, en Vercel: **Add New → Project → Import Git Repository** y elige ese
repositorio (usa la integración propia de Vercel con GitHub — no hace falta
ningún token adicional). No necesita *Build Command* ni *Output Directory*:
es un proyecto "Other", Vercel sirve `index.html` y detecta `api/data.js`
como función automáticamente. Una vez agregadas las variables de entorno del
paso 3, cada `git push` a `main` vuelve a desplegar solo.

## 5. Verificar

- Abre la URL que te da Vercel: el encabezado debe decir "En vivo · Cuadro
  Control 2026.xlsx" en un par de segundos.
- Si ves un aviso en vez de datos, el mensaje indica exactamente qué falta
  (credenciales, permisos, o el ID del archivo) — revisa la variable de
  entorno correspondiente en Vercel y vuelve a desplegar.
- Mientras completas la configuración, el botón **"Cargar Excel"** del
  dashboard sigue funcionando para cualquiera, cargando el archivo
  manualmente desde el navegador.

## Notas

- Los datos se refrescan solos cada 60 segundos en cada navegador abierto, y
  también al volver a la pestaña. El botón "Actualizar" fuerza una lectura
  inmediata del Excel (ignora la caché de 45s del backend).
- Toda la lógica de negocio (semáforo, presupuesto compartido FOOD/HBC,
  filtros, tendencias) vive en `index.html` y es idéntica a la versión que
  corría dentro de Claude; solo cambió de dónde se obtienen los datos.
- `lib/parse.js` contiene el parseo del Excel (compartido conceptualmente con
  la lógica del dashboard) y `api/data.js` es la única pieza nueva que habla
  con Microsoft Graph.
