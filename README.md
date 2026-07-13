# Panel de Pedidos

App autocontenida (backend + frontend en un mismo servidor Node). No usa Azure Functions,
ni Microsoft Graph, ni Entra ID. Lee el Excel directamente del enlace compartido de OneDrive.

## 1. Preparar el enlace de OneDrive

1. En OneDrive, clic derecho sobre el Excel → **Compartir**.
2. Tipo de vínculo: **"Personas de tu organización"** o **"Cualquier persona con el vínculo"**
   (usa la que permita vuestra política; si eliges la de organización, tendrás que abrir la URL
   una vez autenticado para comprobar que el servidor puede acceder — si no, habrá que valorar
   la opción de permisos de aplicación que descartamos por complejidad).
3. Permisos: **solo lectura**, pero con **descarga permitida** (no bloquear descarga) — si no,
   la API de `/content` no devuelve el fichero.
4. Copia el enlace y pégalo en `.env` como `ONEDRIVE_SHARE_LINK`.

## 2. Instalar y arrancar

```bash
cd webapp
npm install
cp .env.example .env
# edita .env con tu enlace real de OneDrive
npm start
```

Por defecto escucha en el puerto 3000 (configurable con `PORT` en `.env`).

## 3. Uso

Abre `http://tu-servidor:3000` en el navegador. El usuario introduce su email, pulsa
"Ver pedidos" y la app:
1. Descarga el Excel de OneDrive (con caché en memoria de `CACHE_SECONDS` segundos para no
   descargarlo en cada petición).
2. Filtra las filas por ese email.
3. Agrupa por número de pedido, calcula la fecha disponible de cabecera (MAX de las líneas,
   o vacía si alguna línea no tiene fecha) y clasifica en las 3 categorías.

## 4. Notas de seguridad

- **No hay autenticación** — cualquiera que sepa un email puede ver esos pedidos. Es el mismo
  nivel de seguridad que tenía la versión inicial de subida manual. Si más adelante queréis
  verificar quién es cada usuario, se puede añadir un login simple sin tocar Graph ni Entra ID
  (por ejemplo, un enlace único por usuario, o una contraseña compartida).
- El enlace de OneDrive vive solo en el `.env` del servidor — nunca lo pongas en el HTML/JS
  del frontend, porque ahí sí sería visible para cualquiera que abra el código fuente de la página.

## 5. Desplegar

Funciona igual en cualquier sitio que ejecute Node 18+:
- **Azure App Service** (Web App normal, no Functions): despliegue por Git o zip deploy.
- **VM propia / IIS**: usa `pm2 start server.js` + nginx o iisnode como reverse proxy.
- **Docker**: dockerfile mínimo — Node 18-alpine, `COPY . .`, `npm ci --omit=dev`, `CMD ["node","server.js"]`.

## Columnas esperadas en el Excel

La app busca el nombre entre corchetes en la fila de cabecera, así que aunque cambie el
prefijo de la tabla (`jokin info_pedidos_WEB[...]`) seguirá funcionando:

`xempresa_id`, `xemail`, `xnumdoc_id`, `xfecha_pedido`, `cliente`, `xalmacen_id`,
`xrepresentante_id`, `xarticulo_id`, `Articulo`, `yfecha_disponible`,
`Sumxcantidad_prin`, `Sumxexistencia`, `Sumxdisponible`.

Si falta alguna, el endpoint `/api/pedidos` devuelve un error indicando cuál.
