// Esse Jeffe Otomasyon — tüm uygulama tek dosyada.
// Bölümler: Ayarlar · Yardımcılar · Shopify · Trendyol · Hepsiburada · Excel ·
//           E-posta · Sipariş akışı · Panel sayfası · Sunucu
const crypto = require('crypto');
const express = require('express');
const ExcelJS = require('exceljs');

// ======================================================================
// AYARLAR (değerler Railway > Variables üzerinden gelir)
// ======================================================================
const config = (() => {
  const env = (k, d = '') => (process.env[k] ?? d).toString().trim();

  return {
    port: Number(env('PORT', '3000')),
    // Hangi kanallar taransın (virgülle): shopify,trendyol,hepsiburada
    sources: env('KANALLAR', 'shopify').toLowerCase().split(',').map((x) => x.trim()).filter(Boolean),
    panel: { user: env('PANEL_USER'), password: env('PANEL_PASSWORD') },
    shopify: {
      store: env('SHOPIFY_STORE').replace('.myshopify.com', ''),
      clientId: env('SHOPIFY_CLIENT_ID'),
      clientSecret: env('SHOPIFY_CLIENT_SECRET'),
      accessToken: env('SHOPIFY_ACCESS_TOKEN'),
      apiVersion: env('SHOPIFY_API_VERSION', '2026-07'),
      // Listeye giren siparişlere eklenen etiket; bu etiketi taşıyan siparişler bir daha listelenmez.
      tag: env('SHOPIFY_ETIKET', 'etiketi çıkarıldı - otomatik'),
    },
    trendyol: {
      sellerId: env('TRENDYOL_SELLER_ID'),
      apiKey: env('TRENDYOL_API_KEY'),
      apiSecret: env('TRENDYOL_API_SECRET'),
      days: Number(env('TRENDYOL_GUN', '14')),
    },
    hepsiburada: {
      merchantId: env('HB_MERCHANT_ID'),
      username: env('HB_USERNAME') || env('HB_MERCHANT_ID'),
      password: env('HB_PASSWORD'),
      userAgent: env('HB_USER_AGENT'),
    },
    google: {
      sheetId: env('GOOGLE_SHEET_ID'),
      serviceAccount: env('GOOGLE_SERVICE_ACCOUNT_JSON'),
    },
    mail: {
      resendKey: env('RESEND_API_KEY'),
      from: env('MAIL_FROM'),
      to: env('MAIL_TO'),
    },
  };
})();

// ======================================================================
// YARDIMCILAR
// ======================================================================
const util = (() => {
  // Pazaryeri API'lerinde alan adları zaman zaman değişebildiği için
  // birkaç olası adı sırayla deneyen küçük yardımcılar.
  function pick(obj, ...paths) {
    for (const p of paths) {
      const v = p.split('.').reduce((o, k) => (o == null ? undefined : o[k]), obj);
      if (v !== undefined && v !== null && v !== '') return v;
    }
    return '';
  }

  const toNumber = (v) => {
    if (v && typeof v === 'object') v = v.amount ?? v.value;
    const n = Number(v);
    return Number.isFinite(n) ? n : null;
  };

  const toDate = (v) => {
    if (!v) return null;
    const d = new Date(typeof v === 'string' && /^\d+$/.test(v) ? Number(v) : v);
    return isNaN(d) ? null : d;
  };

  async function httpJson(url, options = {}, label = 'API') {
    const res = await fetch(url, options);
    const text = await res.text();
    if (!res.ok) throw new Error(`${label} ${res.status}: ${text.slice(0, 300)}`);
    return text ? JSON.parse(text) : {};
  }

  return { pick, toNumber, toDate, httpJson };
})();

// ======================================================================
// SHOPIFY
// ======================================================================
const shopify = (() => {
  // Shopify: gönderilmemiş (unfulfilled / kısmi), iptal edilmemiş ve henüz
  // otomatik etiketi almamış siparişler. Liste oluşunca bu siparişlere etiket eklenir.
  // Kargo anahtarı = siparişin uzun sistem ID'si (legacyResourceId).
  const cfg = config.shopify;
  const { httpJson, toNumber, toDate } = util;

  let cachedToken = null; // { value, expiresAt }

  async function getToken() {
    if (cfg.accessToken) return cfg.accessToken; // eski tip özel uygulama
    if (cachedToken && cachedToken.expiresAt > Date.now() + 60_000) return cachedToken.value;
    // Dev Dashboard uygulaması: client credentials ile 24 saatlik token
    const data = await httpJson(
      `https://${cfg.store}.myshopify.com/admin/oauth/access_token`,
      {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: new URLSearchParams({
          grant_type: 'client_credentials',
          client_id: cfg.clientId,
          client_secret: cfg.clientSecret,
        }),
      },
      'Shopify token'
    );
    cachedToken = {
      value: data.access_token,
      expiresAt: Date.now() + (Number(data.expires_in) || 86_399) * 1000,
    };
    return cachedToken.value;
  }

  async function gql(query, variables) {
    const token = await getToken();
    const data = await httpJson(
      `https://${cfg.store}.myshopify.com/admin/api/${cfg.apiVersion}/graphql.json`,
      {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', 'X-Shopify-Access-Token': token },
        body: JSON.stringify({ query, variables }),
      },
      'Shopify'
    );
    if (data.errors) throw new Error('Shopify: ' + JSON.stringify(data.errors).slice(0, 300));
    return data.data;
  }

  const ORDERS = `
  query Orders($cursor: String, $q: String!) {
    orders(first: 50, after: $cursor, sortKey: CREATED_AT, reverse: true, query: $q) {
      pageInfo { hasNextPage endCursor }
      nodes {
        id
        legacyResourceId
        name
        createdAt
        cancelledAt
        tags
        displayFulfillmentStatus
        paymentGatewayNames
        phone
        customer { firstName lastName }
        shippingAddress { name phone }
        lineItems(first: 50) {
          nodes {
            title
            variantTitle
            sku
            quantity
            currentQuantity
            unfulfilledQuantity
            originalUnitPriceSet { shopMoney { amount } }
          }
        }
      }
    }
  }`;

  const TAGS_ADD = `
  mutation TagsAdd($id: ID!, $tags: [String!]!) {
    tagsAdd(id: $id, tags: $tags) { userErrors { field message } }
  }`;

  function checkConfig() {
    if (!cfg.store || !(cfg.accessToken || (cfg.clientId && cfg.clientSecret))) {
      throw new Error('Shopify bilgileri eksik (SHOPIFY_STORE + CLIENT_ID/SECRET)');
    }
  }

  // mode 'yeni'  : gönderilmemiş, iptal edilmemiş, etiketsiz siparişler
  // mode 'gecmis': etiketlenmiş siparişler, from/to (YYYY-MM-DD) tarih aralığında
  function buildQuery({ mode = 'yeni', from, to } = {}) {
    const tag = `tag:"${cfg.tag.replace(/"/g, '')}"`;
    if (mode === 'gecmis') {
      const parts = [tag];
      if (from) parts.push(`created_at:>=${from}`);
      if (to) {
        const next = new Date(to + 'T00:00:00Z');
        next.setUTCDate(next.getUTCDate() + 1);
        parts.push(`created_at:<${next.toISOString().slice(0, 10)}`);
      }
      return parts.join(' AND ');
    }
    return `status:open AND (fulfillment_status:unfulfilled OR fulfillment_status:partial) AND NOT ${tag}`;
  }

  const DURUM = {
    UNFULFILLED: 'Gönderilmedi', PARTIALLY_FULFILLED: 'Kısmen gönderildi', FULFILLED: 'Gönderildi',
    ON_HOLD: 'Beklemede', SCHEDULED: 'Planlandı', IN_PROGRESS: 'Hazırlanıyor', OPEN: 'Açık',
  };

  async function fetchRows(opts = {}) {
    checkConfig();
    const history = opts.mode === 'gecmis';
    const q = buildQuery(opts);
    const rows = [];
    let cursor = null;

    do {
      const page = (await gql(ORDERS, { cursor, q })).orders;
      for (const o of page.nodes) {
        if (!history && o.cancelledAt) continue;
        if (!history && (o.tags || []).includes(cfg.tag)) continue; // arama filtresine ek güvence
        const customer =
          o.shippingAddress?.name ||
          [o.customer?.firstName, o.customer?.lastName].filter(Boolean).join(' ');
        for (const li of o.lineItems.nodes) {
          // currentQuantity: siparişten çıkarılan / değiştirilen ürünlerde 0 olur.
          const current = li.currentQuantity ?? li.quantity;
          const qty = history ? current : Math.min(li.unfulfilledQuantity, current);
          if (!qty) continue; // çıkarılmış, değiştirilmiş ya da (yeni listede) gönderilmiş kalem
          const unit = toNumber(li.originalUnitPriceSet?.shopMoney?.amount);
          rows.push({
            _key: o.id,
            _gid: o.id,
            kanal: 'Shopify',
            siparisNo: o.name,
            tarih: toDate(o.createdAt),
            musteri: customer,
            telefon: o.shippingAddress?.phone || o.phone || '',
            kargoFirmasi: 'DHL eCommerce',
            kargoAnahtari: String(o.legacyResourceId),
            urun: li.title,
            varyant: li.variantTitle || '',
            sku: li.sku || '',
            adet: qty,
            tutar: unit != null ? unit * qty : null,
            odemeTipi: (o.paymentGatewayNames || []).join(', '),
            durum: o.cancelledAt ? 'İptal' : (DURUM[o.displayFulfillmentStatus] || o.displayFulfillmentStatus),
            etiketler: (o.tags || []).filter((t) => t !== cfg.tag).join(', '),
          });
        }
      }
      cursor = page.pageInfo.hasNextPage ? page.pageInfo.endCursor : null;
    } while (cursor);

    return rows;
  }

  // Listeye giren siparişleri etiketler. Başarısız olanların sipariş numaralarını döner.
  async function tagOrders(rows) {
    const orders = new Map(rows.filter((r) => r._gid).map((r) => [r._gid, r.siparisNo]));
    const failed = [];
    let tagged = 0;
    for (const [id, name] of orders) {
      try {
        const res = await gql(TAGS_ADD, { id, tags: [cfg.tag] });
        const errs = res.tagsAdd.userErrors;
        if (errs.length) throw new Error(errs.map((e) => e.message).join(', '));
        tagged++;
      } catch (e) {
        failed.push(`${name} (${e.message})`);
      }
    }
    return { tagged, failed };
  }

  return { fetchRows, tagOrders, buildQuery };
})();

// ======================================================================
// TRENDYOL
// ======================================================================
const trendyol = (() => {
  // Trendyol: henüz kargoya verilmemiş paketler (Created, Picking, Invoiced).
  // Kargo kodu Trendyol'un atadığı cargoTrackingNumber.
  const cfg = config.trendyol;
  const { httpJson, pick, toNumber, toDate } = util;

  const BASE = 'https://apigw.trendyol.com/integration/order/sellers';
  const STATUSES = ['Created', 'Picking', 'Invoiced'];
  const DAY = 24 * 60 * 60 * 1000;

  async function fetchStatus(status, headers) {
    const packages = [];
    const endDate = Date.now();
    const startDate = endDate - cfg.days * DAY;
    let page = 0;
    let totalPages = 1;
    while (page < totalPages) {
      const url =
        `${BASE}/${cfg.sellerId}/orders?status=${status}` +
        `&startDate=${startDate}&endDate=${endDate}` +
        `&orderByField=PackageLastModifiedDate&orderByDirection=DESC&size=200&page=${page}`;
      const data = await httpJson(url, { headers }, `Trendyol (${status})`);
      packages.push(...(data.content || []));
      totalPages = data.totalPages || 1;
      page++;
    }
    return packages;
  }

  async function fetchRows() {
    if (!cfg.sellerId || !cfg.apiKey || !cfg.apiSecret) {
      throw new Error('Trendyol bilgileri eksik (SELLER_ID / API_KEY / API_SECRET)');
    }
    const headers = {
      Authorization: 'Basic ' + Buffer.from(`${cfg.apiKey}:${cfg.apiSecret}`).toString('base64'),
      'User-Agent': `${cfg.sellerId} - SelfIntegration`,
    };

    const rows = [];
    for (const status of STATUSES) {
      for (const p of await fetchStatus(status, headers)) {
        const customer =
          pick(p, 'shipmentAddress.fullName') ||
          [p.customerFirstName, p.customerLastName].filter(Boolean).join(' ');
        for (const line of p.lines || []) {
          rows.push({
            kanal: 'Trendyol',
            siparisNo: String(pick(p, 'orderNumber')),
            tarih: toDate(p.orderDate),
            musteri: customer,
            telefon: pick(p, 'shipmentAddress.phone'),
            kargoFirmasi: pick(p, 'cargoProviderName'),
            kargoAnahtari: String(pick(p, 'cargoTrackingNumber')),
            urun: pick(line, 'productName'),
            varyant: [pick(line, 'productColor'), pick(line, 'productSize')].filter(Boolean).join(' / '),
            sku: pick(line, 'merchantSku', 'barcode'),
            adet: toNumber(line.quantity),
            tutar: toNumber(pick(line, 'amount', 'price')),
            odemeTipi: 'Trendyol',
            durum: status,
          });
        }
      }
    }
    return rows;
  }

  return { fetchRows };
})();

// ======================================================================
// HEPSIBURADA
// ======================================================================
const hepsiburada = (() => {
  // Hepsiburada: iki grup sipariş
  //  1) Paketlenmemiş kalemler (sipariş alındı)  -> kargo kodu henüz yok
  //  2) Paketlenmiş ama kargoya verilmemiş paketler -> kargo kodu/barkod var
  // Not: Alan adları ilk gerçek çalıştırmada kontrol edilecek; pick() birden
  // fazla olası adı denediği için küçük farklar sorun çıkarmaz.
  const cfg = config.hepsiburada;
  const { httpJson, pick, toNumber, toDate } = util;

  const BASE = 'https://oms-external.hepsiburada.com';

  async function paged(path, headers, label) {
    const all = [];
    const limit = 100;
    for (let offset = 0; ; offset += limit) {
      const data = await httpJson(
        `${BASE}${path}?offset=${offset}&limit=${limit}`,
        { headers },
        label
      );
      const items = Array.isArray(data) ? data : data.items || data.data || [];
      all.push(...items);
      if (items.length < limit) break;
    }
    return all;
  }

  function lineToRow(src, line, extra) {
    return {
      kanal: 'Hepsiburada',
      siparisNo: String(pick(line, 'orderNumber') || pick(src, 'orderNumber')),
      tarih: toDate(pick(line, 'orderDate') || pick(src, 'orderDate')),
      musteri: pick(src, 'recipientName', 'customerName', 'shippingAddress.name'),
      telefon: pick(src, 'phoneNumber', 'shippingAddress.phoneNumber'),
      kargoFirmasi: pick(src, 'cargoCompany', 'cargoCompanyModel.name'),
      urun: pick(line, 'productName', 'name'),
      varyant: pick(line, 'properties', 'variantName'),
      sku: pick(line, 'merchantSku', 'sku', 'hepsiburadaSku'),
      adet: toNumber(pick(line, 'quantity')),
      tutar: toNumber(pick(line, 'totalPrice', 'price')),
      odemeTipi: 'Hepsiburada',
      ...extra,
    };
  }

  async function fetchRows() {
    if (!cfg.merchantId || !cfg.password) {
      throw new Error('Hepsiburada bilgileri eksik (MERCHANT_ID / PASSWORD)');
    }
    const headers = {
      Authorization: 'Basic ' + Buffer.from(`${cfg.username}:${cfg.password}`).toString('base64'),
      'User-Agent': cfg.userAgent || 'esse-jeffe-otomasyon',
      Accept: 'application/json',
    };

    const rows = [];

    const openItems = await paged(`/orders/merchantid/${cfg.merchantId}`, headers, 'Hepsiburada (paketlenecek)');
    for (const item of openItems) {
      rows.push(lineToRow(item, item, { kargoAnahtari: '', durum: 'Paketlenecek' }));
    }

    const packages = await paged(`/packages/merchantid/${cfg.merchantId}`, headers, 'Hepsiburada (paketler)');
    for (const p of packages) {
      const code = String(pick(p, 'barcode', 'trackingNumber', 'cargoTrackingNumber', 'packageNumber'));
      for (const line of p.items || p.lines || [p]) {
        rows.push(lineToRow(p, line, { kargoAnahtari: code, durum: 'Paketlendi' }));
      }
    }
    return rows;
  }

  return { fetchRows };
})();

// ======================================================================
// EXCEL
// ======================================================================
const excel = (() => {
  // Her satır bir sipariş. Excel ve Google Sheet aynı kolon düzenini kullanır.
  const COLUMNS = [
    { header: 'Kanal', key: 'kanal', width: 12 },
    { header: 'Sipariş No', key: 'siparisNo', width: 14 },
    { header: 'Sipariş Tarihi', key: 'tarih', width: 17 },
    { header: 'Müşteri', key: 'musteri', width: 22 },
    { header: 'Telefon', key: 'telefon', width: 16 },
    { header: 'Kargo Firması', key: 'kargoFirmasi', width: 15 },
    { header: 'Kargo Anahtarı', key: 'kargoAnahtari', width: 18 },
    { header: 'SKU', key: 'sku', width: 24 },
    { header: 'Adet', key: 'adet', width: 7 },
    { header: 'Tutar (₺)', key: 'tutar', width: 12 },
    { header: 'Ödeme', key: 'odemeTipi', width: 16 },
    { header: 'Durum', key: 'durum', width: 13 },
    { header: 'Etiketler', key: 'etiketler', width: 30 },
  ];

  const skuText = (o) => o.skus.map((x) => (x.adet > 1 ? `${x.sku} ×${x.adet}` : x.sku)).join('\n');
  const dateText = (d) => d ? new Date(d).toLocaleString('tr-TR', {
    timeZone: 'Europe/Istanbul', day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit',
  }) : '';

  // Sipariş -> düz değerler (Excel ve Sheet için ortak)
  function toRecord(o) {
    return {
      kanal: o.kanal, siparisNo: o.siparisNo, tarih: dateText(o.tarih), musteri: o.musteri,
      telefon: o.telefon, kargoFirmasi: o.kargoFirmasi, kargoAnahtari: o.kargoAnahtari,
      sku: skuText(o), adet: o.adet, tutar: o.tutar ? Math.round(o.tutar * 100) / 100 : '',
      odemeTipi: o.odemeTipi, durum: o.durum, etiketler: o.etiketler || '',
    };
  }

  const FONT = { name: 'Arial', size: 10 };

  function addSheet(wb, name, orders) {
    const ws = wb.addWorksheet(name, { views: [{ state: 'frozen', ySplit: 1 }] });
    ws.columns = COLUMNS;
    orders.forEach((o) => ws.addRow(toRecord(o)));
    ws.getColumn('tutar').numFmt = '#,##0.00';
    ['siparisNo', 'kargoAnahtari', 'telefon'].forEach((k) => (ws.getColumn(k).numFmt = '@'));
    ws.eachRow((row, i) => {
      row.font = i === 1 ? { ...FONT, bold: true, color: { argb: 'FFFFFFFF' } } : FONT;
      row.alignment = { vertical: 'top', wrapText: true };
      if (i === 1) {
        row.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF3B2A2F' } };
        row.height = 20;
      }
    });
    if (orders.length) ws.autoFilter = { from: 'A1', to: { row: 1, column: COLUMNS.length } };
  }

  async function buildWorkbook(orders, warnings = [], channels = ['Shopify', 'Trendyol', 'Hepsiburada']) {
    const wb = new ExcelJS.Workbook();
    wb.creator = 'Esse Jeffe Otomasyon';
    // Tek kanal açıksa tek sayfa; birden fazlaysa "Tümü" + kanal sayfaları.
    if (channels.length > 1) addSheet(wb, 'Tümü', orders);
    for (const kanal of channels) addSheet(wb, kanal, orders.filter((o) => o.kanal === kanal));
    if (warnings.length) {
      const ws = wb.addWorksheet('Uyarılar');
      ws.columns = [{ header: 'Uyarı', key: 'w', width: 120 }];
      warnings.forEach((w) => ws.addRow({ w }));
      ws.eachRow((row, i) => (row.font = { ...FONT, bold: i === 1, color: i === 1 ? undefined : { argb: 'FFB00020' } }));
    }
    return wb.xlsx.writeBuffer();
  }

  return { buildWorkbook, COLUMNS, toRecord };
})();

// ======================================================================
// E-POSTA
// ======================================================================
const mailer = (() => {
  // E-posta Resend üzerinden gönderilir (HTTP API; Railway'de SMTP portlarına ihtiyaç duymaz).
  const cfg = config.mail;

  async function sendMail({ subject, html, attachments = [] }) {
    if (!cfg.resendKey || !cfg.to) return { skipped: true, reason: 'E-posta ayarları eksik' };
    const res = await fetch('https://api.resend.com/emails', {
      method: 'POST',
      headers: { Authorization: `Bearer ${cfg.resendKey}`, 'Content-Type': 'application/json' },
      body: JSON.stringify({
        from: cfg.from,
        to: cfg.to.split(',').map((s) => s.trim()),
        subject,
        html,
        attachments: attachments.map((a) => ({
          filename: a.filename,
          content: Buffer.from(a.content).toString('base64'),
        })),
      }),
    });
    if (!res.ok) throw new Error(`E-posta gönderilemedi: ${res.status} ${await res.text()}`);
    return { sent: true };
  }

  return { sendMail };
})();

// ======================================================================
// GOOGLE SHEETS (her liste tablonun başına yeni bir sekme olarak eklenir)
// ======================================================================
const sheets = (() => {
  const cfg = config.google;
  const API = 'https://sheets.googleapis.com/v4/spreadsheets';
  let cached = null; // { token, expiresAt }

  const b64url = (v) => Buffer.from(typeof v === 'string' ? v : JSON.stringify(v)).toString('base64url');

  async function getToken() {
    if (cached && cached.expiresAt > Date.now() + 60_000) return cached.token;
    let sa;
    try { sa = JSON.parse(cfg.serviceAccount); } catch { throw new Error('GOOGLE_SERVICE_ACCOUNT_JSON okunamadı (JSON dosyasının tamamını yapıştırın)'); }
    const now = Math.floor(Date.now() / 1000);
    const unsigned = b64url({ alg: 'RS256', typ: 'JWT' }) + '.' + b64url({
      iss: sa.client_email, scope: 'https://www.googleapis.com/auth/spreadsheets',
      aud: 'https://oauth2.googleapis.com/token', iat: now, exp: now + 3600,
    });
    const signature = crypto.createSign('RSA-SHA256').update(unsigned).sign(sa.private_key, 'base64url');
    const data = await util.httpJson('https://oauth2.googleapis.com/token', {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({ grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer', assertion: `${unsigned}.${signature}` }),
    }, 'Google token');
    cached = { token: data.access_token, expiresAt: Date.now() + (data.expires_in || 3600) * 1000 };
    return cached.token;
  }

  async function api(path, method, body) {
    const token = await getToken();
    return util.httpJson(`${API}/${cfg.sheetId}${path}`, {
      method, headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
      body: body ? JSON.stringify(body) : undefined,
    }, 'Google Sheets');
  }

  function tabTitle(prefix) {
    const t = new Date().toLocaleString('tr-TR', {
      timeZone: 'Europe/Istanbul', day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit', second: '2-digit',
    });
    return `${prefix} ${t}`.replace(/[\[\]:*?\/\\]/g, '.');
  }

  async function writeOrders(orders, prefix) {
    if (!cfg.sheetId || !cfg.serviceAccount) return { skipped: true };
    const title = tabTitle(prefix);
    const cols = excel.COLUMNS;

    const added = await api(':batchUpdate', 'POST', {
      requests: [{ addSheet: { properties: { title, index: 0, gridProperties: { frozenRowCount: 1 } } } }],
    });
    const gid = added.replies[0].addSheet.properties.sheetId;

    const values = [cols.map((c) => c.header), ...orders.map((o) => {
      const r = excel.toRecord(o);
      return cols.map((c) => r[c.key] ?? '');
    })];
    const range = encodeURIComponent(`'${title.replace(/'/g, "''")}'!A1`);
    // RAW: uzun ID'ler ve telefonlar sayıya dönüşmeden metin olarak kalır.
    await api(`/values/${range}?valueInputOption=RAW`, 'PUT', { values });

    const n = cols.length;
    await api(':batchUpdate', 'POST', { requests: [
      { repeatCell: {
          range: { sheetId: gid, startRowIndex: 0, endRowIndex: 1, startColumnIndex: 0, endColumnIndex: n },
          cell: { userEnteredFormat: { textFormat: { bold: true, foregroundColor: { red: 1, green: 1, blue: 1 } },
                  backgroundColor: { red: 0.231, green: 0.165, blue: 0.184 } } },
          fields: 'userEnteredFormat(textFormat,backgroundColor)' } },
      { repeatCell: {
          range: { sheetId: gid, startRowIndex: 1, startColumnIndex: 0, endColumnIndex: n },
          cell: { userEnteredFormat: { wrapStrategy: 'WRAP', verticalAlignment: 'TOP' } },
          fields: 'userEnteredFormat(wrapStrategy,verticalAlignment)' } },
      ...cols.map((c, i) => ({ updateDimensionProperties: {
          range: { sheetId: gid, dimension: 'COLUMNS', startIndex: i, endIndex: i + 1 },
          properties: { pixelSize: Math.round(c.width * 8) }, fields: 'pixelSize' } })),
    ] });

    return { url: `https://docs.google.com/spreadsheets/d/${cfg.sheetId}/edit#gid=${gid}`, title };
  }

  return { writeOrders };
})();

// ======================================================================
// SİPARİŞ AKIŞI
// ======================================================================
const orders = (() => {
  // Sipariş akışı iki adımlı:
  //  1) fetchOrders: siparişleri çeker, panelde tablo olarak gösterilmek üzere döner
  //     ve satırları kısa süreliğine sunucuda saklar.
  //  2) exportSelected: panelde seçilen siparişlerle Excel'i oluşturur, e-postalar;
  //     "yeni" listede Shopify siparişlerini etiketler.
  const { buildWorkbook } = excel;
  const { sendMail } = mailer;

  const ALL = { shopify: ['Shopify', shopify], trendyol: ['Trendyol', trendyol], hepsiburada: ['Hepsiburada', hepsiburada] };

  const sessions = new Map(); // id -> { mode, rows, warnings, channels }
  const SESSION_MS = 60 * 60 * 1000;

  function remember(data) {
    const id = crypto.randomUUID();
    sessions.set(id, data);
    setTimeout(() => sessions.delete(id), SESSION_MS).unref();
    return id;
  }

  // Her sipariş tek satır; SKU'lar alt alta listelenir.
  function groupByOrder(rows) {
    const map = new Map();
    for (const r of rows) {
      if (!map.has(r._key)) {
        map.set(r._key, {
          key: r._key, kanal: r.kanal, siparisNo: r.siparisNo, tarih: r.tarih, musteri: r.musteri,
          telefon: r.telefon, kargoFirmasi: r.kargoFirmasi, kargoAnahtari: r.kargoAnahtari,
          odemeTipi: r.odemeTipi, durum: r.durum, etiketler: r.etiketler || '',
          tutar: 0, adet: 0, skus: [],
        });
      }
      const o = map.get(r._key);
      o.skus.push({ sku: r.sku || r.urun || '(SKU yok)', adet: r.adet });
      o.adet += r.adet || 0;
      o.tutar += r.tutar || 0;
    }
    return [...map.values()].sort((a, b) => (b.tarih || 0) - (a.tarih || 0));
  }

  async function fetchOrders({ mode = 'yeni', from, to } = {}) {
    // Geçmiş (etiketli) siparişler yalnızca Shopify'da var.
    const keys = mode === 'gecmis' ? ['shopify'] : config.sources;
    const sources = keys.filter((k) => ALL[k]).map((k) => ALL[k]);

    const results = await Promise.allSettled(sources.map(([, src]) => src.fetchRows({ mode, from, to })));
    const rows = [];
    const warnings = [];
    results.forEach((r, i) => {
      const name = sources[i][0];
      if (r.status === 'fulfilled') {
        for (const row of r.value) rows.push({ ...row, _key: row._key || `${name}:${row.siparisNo}` });
      } else {
        warnings.push(`${name}: ${r.reason.message}`);
        console.error(`[${name}]`, r.reason);
      }
    });

    const channels = sources.map(([name]) => name);
    const sessionId = remember({ mode, rows, warnings, channels, from, to });
    return { sessionId, mode, orders: groupByOrder(rows), warnings };
  }

  function fileName(prefix, d = new Date()) {
    const p = (n) => String(n).padStart(2, '0');
    const tr = new Date(d.toLocaleString('en-US', { timeZone: 'Europe/Istanbul' }));
    return `${prefix}_${tr.getFullYear()}-${p(tr.getMonth() + 1)}-${p(tr.getDate())}_${p(tr.getHours())}${p(tr.getMinutes())}.xlsx`;
  }

  async function exportSelected({ sessionId, keys, email = true }) {
    const s = sessions.get(sessionId);
    if (!s) throw new Error('Liste süresi doldu, siparişleri yeniden getirin.');
    const wanted = new Set(keys || []);
    const rows = s.rows.filter((r) => wanted.has(r._key));
    if (!rows.length) throw new Error("Seçilen siparişler bu listeden zaten Excel'e alınmış; siparişleri yeniden getirin.");

    const warnings = [...s.warnings];
    const selected = groupByOrder(rows);
    const history = s.mode === 'gecmis';
    let tagging = null;
    let buffer = await buildWorkbook(selected, warnings, s.channels);

    // Etiketleme yalnızca yeni listede ve Excel sorunsuz oluştuktan sonra yapılır.
    if (!history && rows.some((r) => r._gid)) {
      tagging = await shopify.tagOrders(rows);
      if (tagging.failed.length) {
        warnings.push(`Shopify etiketlenemeyen siparişler (bir sonraki listede tekrar çıkar): ${tagging.failed.join('; ')}`);
      }
      // Aynı listeden ikinci kez etiketleme yapılmasın diye seçilenleri oturumdan çıkar.
      s.rows = s.rows.filter((r) => !wanted.has(r._key));
    }

    // Google Sheet: tablonun başına yeni sekme
    let sheet = { skipped: true };
    try {
      sheet = await sheets.writeOrders(selected, history ? 'Geçmiş' : 'Liste');
    } catch (e) {
      console.error(e);
      sheet = { error: e.message };
      warnings.push(`Google Sheet'e yazılamadı: ${e.message}`);
    }

    if (warnings.length !== s.warnings.length) buffer = await buildWorkbook(selected, warnings, s.channels);

    const filename = fileName(history ? 'gecmis_siparisler' : 'siparisler');
    const orderCount = selected.length;

    let mail = { skipped: true };
    if (email) {
      const title = history
        ? `Etiketlenmiş siparişler (${s.from || '…'} – ${s.to || '…'})`
        : 'Kargoya verilecek siparişler';
      const sheetLine = sheet.url ? `<p><a href="${sheet.url}">Google Sheet'te aç</a> (sekme: ${sheet.title})</p>` : '';
      const tagLine = tagging ? `<p>${tagging.tagged} Shopify siparişine "${config.shopify.tag}" etiketi eklendi.</p>` : '';
      const warn = warnings.length ? `<p style="color:#b00020"><b>Uyarı:</b><br>${warnings.join('<br>')}</p>` : '';
      try {
        mail = await sendMail({
          subject: `${title} – ${orderCount} sipariş`,
          html: `<p>${title}: <b>${orderCount}</b> sipariş. Excel ekte.</p>${sheetLine}${tagLine}${warn}`,
          attachments: [{ filename, content: buffer }],
        });
      } catch (e) {
        mail = { error: e.message };
        console.error(e);
      }
    }
    return { buffer, filename, orderCount, warnings, mail, sheet, tagged: tagging ? tagging.tagged : 0 };
  }

  return { fetchOrders, exportSelected };
})();

// ======================================================================
// PANEL SAYFASI
// ======================================================================
const PAGE = `<!doctype html>
<html lang="tr"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Esse Jeffe Otomasyon</title>
<style>
  :root{--ink:#2b2023;--muted:#7a6a6e;--brand:#3b2a2f;--bg:#f6f2ee;--line:#e6dcd6;--zebra:#fcfaf8;--warn:#b00020;--ok:#1d6b3a}
  *{box-sizing:border-box}
  body{font-family:system-ui,-apple-system,Segoe UI,Roboto,sans-serif;background:var(--bg);color:var(--ink);margin:0;padding:20px 24px;font-size:15px}
  .wrap{max-width:1800px;margin:0 auto}
  h1{font-size:22px;margin:0 0 16px}
  .tabs{display:flex;gap:6px;margin-bottom:-1px}
  .tab{padding:11px 18px;border:1px solid var(--line);border-bottom:0;border-radius:10px 10px 0 0;background:#efe8e3;cursor:pointer;font-weight:600;color:var(--muted)}
  .tab.on{background:#fff;color:var(--ink)}
  .card{background:#fff;border:1px solid var(--line);border-radius:0 12px 12px 12px;padding:20px}
  .bar{display:flex;flex-wrap:wrap;gap:10px;align-items:center;margin-bottom:12px}
  button{padding:12px 18px;font-size:15px;font-weight:600;border:0;border-radius:10px;background:var(--brand);color:#fff;cursor:pointer}
  button:disabled{opacity:.5;cursor:not-allowed}
  input[type=date]{padding:10px;border:1px solid var(--line);border-radius:8px;font:inherit}
  label.chk{display:flex;align-items:center;gap:6px;color:var(--muted)}
  .scroll{overflow:auto;max-height:calc(100vh - 230px);border:1px solid var(--line);border-radius:10px}
  table{border-collapse:collapse;width:100%;font-size:15px;min-width:1100px}
  th,td{padding:12px 14px;border-bottom:1px solid var(--line);text-align:left;vertical-align:top;line-height:1.45}
  th{background:#f3ece7;font-weight:700;position:sticky;top:0;z-index:1;white-space:nowrap}
  tbody tr:nth-child(even) td{background:var(--zebra)}
  tbody tr:hover td{background:#f5ede8}
  tr.off td{opacity:.4}
  td.num{text-align:right;white-space:nowrap}
  td.nowrap{white-space:nowrap}
  .sku{display:block;font-family:ui-monospace,Menlo,monospace;font-size:14px;white-space:nowrap}
  .mono{font-family:ui-monospace,Menlo,monospace;font-size:14px}
  .tags{color:#5b4a4f;min-width:220px}
  input[type=checkbox]{width:18px;height:18px}
  .empty{padding:40px;text-align:center;color:var(--muted)}
  #hint{font-size:14px;color:var(--muted);margin-bottom:10px}
  #msg{margin-top:12px;line-height:1.7}
  .warn{color:var(--warn)} .ok{color:var(--ok)}
  a.dl{font-weight:700;color:var(--brand);margin-right:16px}
  .hide{display:none}
</style></head>
<body><div class="wrap">
  <h1>Esse Jeffe Otomasyon</h1>
  <div class="tabs">
    <div class="tab on" data-mode="yeni">Yeni siparişler</div>
    <div class="tab" data-mode="gecmis">Geçmiş (etiketlenmiş)</div>
  </div>
  <div class="card">
    <div class="bar">
      <span id="range" class="hide">
        <input type="date" id="bas"> – <input type="date" id="bit">
      </span>
      <button id="fetch">Siparişleri getir</button>
      <span style="flex:1"></span>
      <label class="chk"><input type="checkbox" id="mail" checked> E-posta da gönder</label>
      <button id="export" disabled>Excel oluştur</button>
    </div>
    <div id="hint"></div>
    <div class="scroll"><table>
      <thead><tr>
        <th><input type="checkbox" id="all"></th><th>Kanal</th><th>Sipariş</th><th>Tarih</th><th>Müşteri</th>
        <th>Telefon</th><th>SKU</th><th>Tutar</th><th>Ödeme</th><th>Kargo anahtarı</th><th>Durum</th><th>Etiketler</th>
      </tr></thead>
      <tbody id="rows"><tr><td colspan="12" class="empty">"Siparişleri getir"e basın.</td></tr></tbody>
    </table></div>
    <div id="msg"></div>
  </div>
</div>
<script>
let mode = 'yeni', session = null, orders = [];
const COLS = 12;
const $ = (id) => document.getElementById(id);
const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const money = (n) => n ? n.toLocaleString('tr-TR', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + ' ₺' : '';
const when = (d) => d ? new Date(d).toLocaleString('tr-TR', { day:'2-digit', month:'2-digit', hour:'2-digit', minute:'2-digit' }) : '';
const iso = (d) => d.toISOString().slice(0, 10);
const empty = (t) => '<tr><td colspan="' + COLS + '" class="empty">' + t + '</td></tr>';

$('bas').value = iso(new Date(Date.now() - 6 * 864e5)); $('bit').value = iso(new Date());

function setHint() {
  $('hint').textContent = mode === 'yeni'
    ? 'Excel oluşturulunca seçilen Shopify siparişleri "etiketi çıkarıldı - otomatik" olarak etiketlenir ve bir daha bu listede çıkmaz.'
    : 'Daha önce etiketlenmiş Shopify siparişleri. Bu sekmeden Excel almak etiketleri değiştirmez.';
}
setHint();

document.querySelectorAll('.tab').forEach((t) => t.onclick = () => {
  document.querySelectorAll('.tab').forEach((x) => x.classList.toggle('on', x === t));
  mode = t.dataset.mode; session = null; orders = [];
  $('range').classList.toggle('hide', mode !== 'gecmis');
  $('rows').innerHTML = empty('"Siparişleri getir"e basın.');
  $('msg').innerHTML = ''; setHint(); update();
});

function selected() { return [...document.querySelectorAll('.pick:checked')].map((c) => c.value); }
function update() {
  const n = selected().length;
  $('export').disabled = !n;
  $('export').textContent = n ? 'Excel oluştur (' + n + ')' : 'Excel oluştur';
  $('all').checked = n && n === orders.length;
  document.querySelectorAll('.pick').forEach((c) => c.closest('tr').classList.toggle('off', !c.checked));
}
$('all').onchange = (e) => { document.querySelectorAll('.pick').forEach((c) => c.checked = e.target.checked); update(); };

function render() {
  if (!orders.length) {
    $('rows').innerHTML = empty(mode === 'yeni' ? 'Kargoya verilecek yeni sipariş yok.' : 'Bu tarih aralığında etiketlenmiş sipariş yok.');
    return update();
  }
  $('rows').innerHTML = orders.map((o) =>
    '<tr><td><input type="checkbox" class="pick" value="' + esc(o.key) + '" checked></td>' +
    '<td>' + esc(o.kanal) + '</td><td class="nowrap"><b>' + esc(o.siparisNo) + '</b></td><td class="nowrap">' + when(o.tarih) + '</td>' +
    '<td>' + esc(o.musteri) + '</td><td class="nowrap">' + esc(o.telefon) + '</td>' +
    '<td>' + o.skus.map((x) => '<span class="sku">' + esc(x.sku) + (x.adet > 1 ? ' ×' + esc(x.adet) : '') + '</span>').join('') + '</td>' +
    '<td class="num">' + money(o.tutar) + '</td><td>' + esc(o.odemeTipi) + '</td>' +
    '<td class="mono">' + esc(o.kargoAnahtari) + '</td><td>' + esc(o.durum) + '</td>' +
    '<td class="tags">' + esc(o.etiketler) + '</td></tr>'
  ).join('');
  document.querySelectorAll('.pick').forEach((c) => c.onchange = update);
  update();
}

async function call(url, opt) {
  const r = await fetch(url, opt);
  const d = await r.json().catch(() => ({}));
  if (!r.ok) throw new Error(d.error || 'Bir hata oluştu');
  return d;
}

$('fetch').onclick = async () => {
  $('fetch').disabled = true; $('fetch').textContent = 'Getiriliyor…'; $('msg').innerHTML = '';
  try {
    const q = mode === 'gecmis' ? '&bas=' + $('bas').value + '&bit=' + $('bit').value : '';
    const d = await call('/api/siparisler?mod=' + mode + q);
    session = d.sessionId; orders = d.orders; render();
    $('msg').innerHTML = '<span>' + orders.length + ' sipariş bulundu.</span>' +
      (d.warnings.length ? '<p class="warn">' + d.warnings.map(esc).join('<br>') + '</p>' : '');
  } catch (e) { $('msg').innerHTML = '<p class="warn">' + esc(e.message) + '</p>'; }
  finally { $('fetch').disabled = false; $('fetch').textContent = 'Siparişleri getir'; }
};

$('export').onclick = async () => {
  const keys = selected();
  if (mode === 'yeni' && !confirm(keys.length + ' sipariş Excel\\'e alınacak ve Shopify\\'da etiketlenecek. Devam edilsin mi?')) return;
  $('export').disabled = true; $('export').textContent = 'Hazırlanıyor…';
  try {
    const d = await call('/api/excel', { method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ sessionId: session, keys, email: $('mail').checked }) });
    const mail = d.mail.sent ? 'E-postanıza gönderildi.' :
      d.mail.error ? '<span class="warn">E-posta gönderilemedi: ' + esc(d.mail.error) + '</span>' :
      $('mail').checked ? 'E-posta ayarlı değil.' : '';
    $('msg').innerHTML = '<span class="ok">' + d.orderCount + ' sipariş listeye alındı.</span>' +
      (d.tagged ? ' ' + d.tagged + ' sipariş etiketlendi.' : '') + ' ' + mail +
      (d.warnings.length ? '<p class="warn">' + d.warnings.map(esc).join('<br>') + '</p>' : '') +
      '<br><a class="dl" href="' + d.download + '">⬇ Excel\\'i indir</a>' +
      (d.sheetUrl ? '<a class="dl" href="' + esc(d.sheetUrl) + '" target="_blank" rel="noopener">↗ Google Sheet\\'te aç</a>' : '');
    if (mode === 'yeni') { orders = orders.filter((o) => !keys.includes(o.key)); render(); }
  } catch (e) { $('msg').innerHTML = '<p class="warn">' + esc(e.message) + '</p>'; update(); }
  finally { if ($('export').textContent === 'Hazırlanıyor…') update(); }
};
</script></body></html>`;

// ======================================================================
// SUNUCU
// ======================================================================
const { fetchOrders, exportSelected } = orders;

const app = express();

// --- Basit şifre koruması (tarayıcının kendi giriş penceresi) ---
function auth(req, res, next) {
  if (!config.panel.user || !config.panel.password) {
    return res.status(500).send('PANEL_USER ve PANEL_PASSWORD tanımlanmamış.');
  }
  const [type, value] = (req.headers.authorization || '').split(' ');
  if (type === 'Basic' && value) {
    const [u, ...rest] = Buffer.from(value, 'base64').toString().split(':');
    const p = rest.join(':');
    const same = (a, b) =>
      a.length === b.length && crypto.timingSafeEqual(Buffer.from(a), Buffer.from(b));
    if (same(u, config.panel.user) && same(p, config.panel.password)) return next();
  }
  res.set('WWW-Authenticate', 'Basic realm="Esse Jeffe Otomasyon"').status(401).send('Giriş gerekli');
}

app.get('/health', (_req, res) => res.send('ok')); // Railway sağlık kontrolü
app.use(auth);

// Oluşturulan dosyalar 1 saat boyunca panelden indirilebilir.
const files = new Map();
const keep = (buffer, filename) => {
  const id = crypto.randomUUID();
  files.set(id, { buffer, filename });
  setTimeout(() => files.delete(id), 60 * 60 * 1000).unref();
  return id;
};

let busy = false;
const once = (fn) => async (req, res) => {
  if (busy) return res.status(409).json({ error: 'Başka bir işlem sürüyor, birkaç saniye bekleyin.' });
  busy = true;
  try {
    await fn(req, res);
  } catch (e) {
    console.error(e);
    res.status(500).json({ error: e.message });
  } finally {
    busy = false;
  }
};

const isDate = (v) => /^\d{4}-\d{2}-\d{2}$/.test(v || '');

app.get('/api/siparisler', once(async (req, res) => {
  const mode = req.query.mod === 'gecmis' ? 'gecmis' : 'yeni';
  const from = isDate(req.query.bas) ? req.query.bas : undefined;
  const to = isDate(req.query.bit) ? req.query.bit : undefined;
  if (mode === 'gecmis' && (!from || !to)) throw new Error('Başlangıç ve bitiş tarihi seçin.');
  res.json(await fetchOrders({ mode, from, to }));
}));

app.post('/api/excel', express.json(), once(async (req, res) => {
  const { sessionId, keys, email } = req.body || {};
  const out = await exportSelected({ sessionId, keys, email: email !== false });
  res.json({
    orderCount: out.orderCount,
    tagged: out.tagged,
    warnings: out.warnings,
    mail: out.mail,
    sheetUrl: out.sheet && out.sheet.url,
    download: `/indir/${keep(out.buffer, out.filename)}`,
  });
}));

app.get('/indir/:id', (req, res) => {
  const f = files.get(req.params.id);
  if (!f) return res.status(404).send('Dosyanın süresi dolmuş, listeyi yeniden oluşturun.');
  res.set({
    'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    'Content-Disposition': `attachment; filename="${f.filename}"`,
  });
  res.send(Buffer.from(f.buffer));
});

app.get('/', (_req, res) => res.type('html').send(PAGE));



app.listen(config.port, () => console.log(`Panel hazır: port ${config.port}`));
