require('dotenv').config();
const express = require('express');
const path = require('path');
const XLSX = require('xlsx');

const {
  PORT = 3000,
  ONEDRIVE_SHARE_LINK,
  CACHE_SECONDS = 60
} = process.env;

if (!ONEDRIVE_SHARE_LINK) {
  console.error('Falta ONEDRIVE_SHARE_LINK en el .env');
  process.exit(1);
}

const app = express();
app.use(express.static(path.join(__dirname, 'public')));

// --- 1. Convertir el enlace de "Compartir" de OneDrive en una URL de descarga directa ---

function encodeSharingUrl(url) {
  const base64 = Buffer.from(url.trim(), 'utf-8').toString('base64');
  const base64url = base64.replace(/=+$/, '').replace(/\//g, '_').replace(/\+/g, '-');
  return `u!${base64url}`;
}

function getDownloadUrl() {
  const encoded = encodeSharingUrl(ONEDRIVE_SHARE_LINK);
  return `https://api.onedrive.com/v1.0/shares/${encoded}/root/content`;
}

// --- 2. Descarga del Excel con caché en memoria (para no descargarlo en cada petición) ---

let cache = { buffer: null, fetchedAt: 0 };

async function downloadExcel() {
  const now = Date.now();
  if (cache.buffer && now - cache.fetchedAt < CACHE_SECONDS * 1000) {
    return cache.buffer;
  }

  const url = getDownloadUrl();
  const res = await fetch(url, { redirect: 'follow' });

  if (!res.ok) {
    const body = await res.text().catch(() => '');
    throw new Error(`No se ha podido descargar el Excel (${res.status}). ${body}`.trim());
  }

  const arrayBuffer = await res.arrayBuffer();
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
    const email = String(req.query.email || '').trim().toLowerCase();
    if (!email) {
      return res.status(400).json({ error: 'Falta el parámetro email.' });
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
