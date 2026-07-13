require('dotenv').config();
const express = require('express');
const path = require('path');
const XLSX = require('xlsx');
const jwt = require('jsonwebtoken');
const jwksClient = require('jwks-rsa');

const {
  PORT = 3000,
  ONEDRIVE_SHARE_LINK,
  CACHE_SECONDS = 60,
  TENANT_ID,
  CLIENT_ID
} = process.env;

if (!ONEDRIVE_SHARE_LINK) {
  console.error('Falta ONEDRIVE_SHARE_LINK en el .env');
  process.exit(1);
}
if (!TENANT_ID || !CLIENT_ID) {
  console.error('Falta TENANT_ID o CLIENT_ID en el .env (necesarios para validar el login de Microsoft)');
  process.exit(1);
}

const app = express();
app.use(express.static(path.join(__dirname, 'public')));

// --- 0. Validación del token de login de Microsoft (solo identifica al usuario, no toca Graph) ---

const jwks = jwksClient({
  jwksUri: `https://login.microsoftonline.com/${TENANT_ID}/discovery/v2.0/keys`
});

function getSigningKey(header, callback) {
  jwks.getSigningKey(header.kid, (err, key) => {
    if (err) return callback(err);
    callback(null, key.getPublicKey());
  });
}

function verifyUserToken(token) {
  return new Promise((resolve, reject) => {
    jwt.verify(
      token,
      getSigningKey,
      {
        audience: CLIENT_ID,
        issuer: `https://login.microsoftonline.com/${TENANT_ID}/v2.0`
      },
      (err, decoded) => {
        if (err) return reject(err);
        resolve(decoded);
      }
    );
  });
}

// --- 1. Convertir el enlace de "Compartir" de SharePoint/OneDrive en una URL de descarga directa ---
// Para SharePoint Online / OneDrive for Business basta con añadir download=1 a la query string
// del propio enlace de compartir (no hace falta pasar por ninguna API de Graph).

function getDownloadUrl() {
  const url = new URL(ONEDRIVE_SHARE_LINK.trim());
  url.searchParams.set('download', '1');
  return url.toString();
}

// --- 2. Descarga del Excel con caché en memoria (para no descargarlo en cada petición) ---

let cache = { buffer: null, fetchedAt: 0 };

async function downloadExcel() {
  const now = Date.now();
  if (cache.buffer && now - cache.fetchedAt < CACHE_SECONDS * 1000) {
    return cache.buffer;
  }

  const url = getDownloadUrl();
  const res = await fetch(url, {
    redirect: 'follow',
    headers: {
      // Algunos servidores de SharePoint devuelven una página de aviso/interstitial
      // a clientes sin cabeceras de navegador; simulamos una para evitarlo.
      'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36'
    }
  });

  if (!res.ok) {
    const body = await res.text().catch(() => '');
    throw new Error(`No se ha podido descargar el Excel (${res.status}). ${body}`.slice(0, 500));
  }

  const contentType = res.headers.get('content-type') || '';
  const arrayBuffer = await res.arrayBuffer();

  // Si en vez del Excel nos devuelven una página HTML, es que el enlace requiere login
  // (o algo ha ido mal) y no hemos recibido el fichero real.
  if (contentType.includes('text/html')) {
    throw new Error('El servidor ha devuelto una página web en vez del Excel. Revisa que el enlace de OneDrive permita descarga y acceso sin iniciar sesión adicional.');
  }

  cache = { buffer: Buffer.from(arrayBuffer), fetchedAt: now };
  return cache.buffer;
}

// --- 3. Parseo y clasificación ---

const FIELD_KEYS = {
  empresa: 'xempresa_id',
  email: 'xemail',
  numdoc: 'xnumdoc_id',
  fechaPedido: 'xfecha_pedido',
  cliente: 'cliente',
  almacen: 'xalmacen_id',
  representante: 'xrepresentante_id',
  articuloId: 'xarticulo_id',
  articulo: 'articulo',
  fechaDisponible: 'yfecha_disponible',
  cantidad: 'sumxcantidad_prin',
  existencia: 'sumxexistencia',
  disponible: 'sumxdisponible'
};

function extractBracketKey(header) {
  const s = String(header || '');
  const m = s.match(/\[(.+?)\]/);
  return (m ? m[1] : s).trim().toLowerCase();
}

function toDateSafe(v) {
  if (v === null || v === undefined || v === '') return null;
  if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
  if (typeof v === 'number') {
    const d = new Date(Math.round((v - 25569) * 86400 * 1000));
    return isNaN(d.getTime()) ? null : d;
  }
  if (typeof v === 'string') {
    const trimmed = v.trim();
    if (trimmed === '') return null;
    const d = new Date(trimmed);
    return isNaN(d.getTime()) ? null : d;
  }
  return null;
}

function toNumberSafe(v) {
  if (v === null || v === undefined || v === '') return 0;
  const n = typeof v === 'number' ? v : parseFloat(String(v).replace(',', '.'));
  return isNaN(n) ? 0 : n;
}

function parseWorkbook(buffer) {
  const wb = XLSX.read(buffer, { type: 'buffer', cellDates: true });
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: '' });

  if (!rows.length) throw new Error('El Excel no contiene filas.');

  const headerRow = rows[0];
  const colIndex = {};
  headerRow.forEach((h, idx) => {
    const key = extractBracketKey(h);
    for (const field in FIELD_KEYS) {
      if (FIELD_KEYS[field] === key) colIndex[field] = idx;
    }
  });

  const missing = Object.keys(FIELD_KEYS).filter((f) => !(f in colIndex));
  if (missing.length) {
    throw new Error('Columnas no encontradas en la cabecera: ' + missing.map((f) => FIELD_KEYS[f]).join(', '));
  }

  const parsed = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r || r.every((c) => c === '' || c === null || c === undefined)) continue;
    parsed.push({
      email: String(r[colIndex.email] || '').trim(),
      numdoc: String(r[colIndex.numdoc] || '').trim(),
      fechaPedido: toDateSafe(r[colIndex.fechaPedido]),
      cliente: r[colIndex.cliente],
      articulo: r[colIndex.articulo],
      articuloId: r[colIndex.articuloId],
      fechaDisponible: toDateSafe(r[colIndex.fechaDisponible]),
      cantidad: toNumberSafe(r[colIndex.cantidad]),
      existencia: toNumberSafe(r[colIndex.existencia]),
      disponible: toNumberSafe(r[colIndex.disponible])
    });
  }
  return parsed;
}

function buildOrders(rows) {
  const map = new Map();
  rows.forEach((row) => {
    if (!map.has(row.numdoc)) {
      map.set(row.numdoc, { numdoc: row.numdoc, cliente: row.cliente, fechaPedido: row.fechaPedido, lines: [] });
    }
    const order = map.get(row.numdoc);
    if (!order.cliente && row.cliente) order.cliente = row.cliente;
    if (!order.fechaPedido && row.fechaPedido) order.fechaPedido = row.fechaPedido;
    order.lines.push({
      articulo: row.articulo,
      articuloId: row.articuloId,
      cantidad: row.cantidad,
      existencia: row.existencia,
      disponible: row.disponible,
      fechaDisponible: row.fechaDisponible
    });
  });

  const orders = Array.from(map.values());
  orders.forEach((o) => {
    const allStock = o.lines.every((l) => l.disponible > 0);
    const allHaveFecha = o.lines.every((l) => l.fechaDisponible instanceof Date);
    o.fechaDisponibleHeader = allHaveFecha
      ? new Date(Math.max(...o.lines.map((l) => l.fechaDisponible.getTime())))
      : null;
    o.clasificacion = allStock ? 'con_stock' : o.fechaDisponibleHeader ? 'sin_stock_con_fecha' : 'sin_stock_sin_fecha';
  });

  const priority = { sin_stock_sin_fecha: 0, sin_stock_con_fecha: 1, con_stock: 2 };
  orders.sort((a, b) => {
    if (priority[a.clasificacion] !== priority[b.clasificacion]) return priority[a.clasificacion] - priority[b.clasificacion];
    const ta = a.fechaDisponibleHeader ? a.fechaDisponibleHeader.getTime() : Infinity;
    const tb = b.fechaDisponibleHeader ? b.fechaDisponibleHeader.getTime() : Infinity;
    if (ta !== tb) return ta - tb;
    return String(a.numdoc).localeCompare(String(b.numdoc));
  });

  return orders;
}

// --- 4. Endpoint HTTP ---

app.get('/api/pedidos', async (req, res) => {
  try {
    const authHeader = req.headers.authorization || '';
    const token = authHeader.replace(/^Bearer\s+/i, '');
    if (!token) {
      return res.status(401).json({ error: 'Falta iniciar sesión con Microsoft.' });
    }

    let decoded;
    try {
      decoded = await verifyUserToken(token);
    } catch (err) {
      return res.status(401).json({ error: 'Sesión no válida o caducada, inicia sesión de nuevo.' });
    }

    const email = (decoded.preferred_username || decoded.email || decoded.upn || '').toLowerCase().trim();
    if (!email) {
      return res.status(400).json({ error: 'Tu cuenta de Microsoft no tiene un email reconocible.' });
    }

    const fileBuffer = await downloadExcel();
    const allRows = parseWorkbook(fileBuffer);

    const userRows = allRows.filter((r) => r.email.toLowerCase() === email);
    const orders = buildOrders(userRows);

    res.json({ email, orders });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

app.get('/health', (req, res) => res.send('ok'));

app.listen(PORT, () => {
  console.log(`Panel de Pedidos escuchando en el puerto ${PORT}`);
});
