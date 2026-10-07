// Esse Jeffe Otomasyon — tüm uygulama tek dosyada.
// Bölümler: Ayarlar · Yardımcılar · Shopify · Trendyol · Hepsiburada · Kayıt · Excel ·
//           Google Sheets · E-posta · Sipariş akışı · Panel sayfası · Sunucu
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
    // Trendyol/Hepsiburada'da Drive'a aktarılanların kaydı burada tutulur (Railway volume yolu: /data)
    dataDir: env('DATA_DIR', './data'),
    panel: { user: env('PANEL_USER'), password: env('PANEL_PASSWORD') },
    shopify: {
      store: env('SHOPIFY_STORE').replace('.myshopify.com', ''),
      clientId: env('SHOPIFY_CLIENT_ID'),
      clientSecret: env('SHOPIFY_CLIENT_SECRET'),
      accessToken: env('SHOPIFY_ACCESS_TOKEN'),
      apiVersion: env('SHOPIFY_API_VERSION', '2026-07'),
      tags: {
        panel: env('SHOPIFY_ETIKET_PANEL', 'etiket oluşturuldu - otomatik'),   // 2. aşama
        drive: env('SHOPIFY_ETIKET_DRIVE', "drive'a aktarıldı - otomatik"),   // 3. aşama
        eski: env('SHOPIFY_ETIKET', 'etiketi çıkarıldı - otomatik'),          // önceki sürüm; 3. aşama sayılır
      },
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
  // Aşamalar etiketlerle izlenir:
  //  yeni  : açık, gönderilmemiş, bizim etiketlerimizden hiçbiri yok
  //  panel : "etiket oluşturuldu - otomatik" var, Drive etiketi yok
  //  drive : "drive'a aktarıldı - otomatik" (veya önceki sürümün etiketi) var
  // Kargo anahtarı = siparişin uzun sistem ID'si (legacyResourceId).
  const cfg = config.shopify;
  const T = cfg.tags;
  const OURS = [T.panel, T.drive, T.eski];
  const { httpJson, toNumber, toDate } = util;

  let cachedToken = null;

  async function getToken() {
    if (cfg.accessToken) return cfg.accessToken;
    if (cachedToken && cachedToken.expiresAt > Date.now() + 60_000) return cachedToken.value;
    const data = await httpJson(
      `https://${cfg.store}.myshopify.com/admin/oauth/access_token`,
      {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: new URLSearchParams({ grant_type: 'client_credentials', client_id: cfg.clientId, client_secret: cfg.clientSecret }),
      },
      'Shopify token'
    );
    cachedToken = { value: data.access_token, expiresAt: Date.now() + (Number(data.expires_in) || 86_399) * 1000 };
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
        id legacyResourceId name createdAt cancelledAt tags
        displayFulfillmentStatus paymentGatewayNames phone
        customer { firstName lastName }
        shippingAddress { name phone address1 address2 city province zip }
        lineItems(first: 50) {
          nodes {
            title sku quantity currentQuantity unfulfilledQuantity
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

  const DURUM = {
    UNFULFILLED: 'Gönderilmedi', PARTIALLY_FULFILLED: 'Kısmen gönderildi', FULFILLED: 'Gönderildi',
    ON_HOLD: 'Beklemede', SCHEDULED: 'Planlandı', IN_PROGRESS: 'Hazırlanıyor', OPEN: 'Açık',
  };

  function checkConfig() {
    if (!cfg.store || !(cfg.accessToken || (cfg.clientId && cfg.clientSecret))) {
      throw new Error('Shopify bilgileri eksik (SHOPIFY_STORE + CLIENT_ID/SECRET)');
    }
  }

  const q = (t) => `tag:"${t.replace(/"/g, '')}"`;

  function dateRange(from, to) {
    const parts = [];
    if (from) parts.push(`created_at:>=${from}`);
    if (to) {
      const next = new Date(to + 'T00:00:00Z');
      next.setUTCDate(next.getUTCDate() + 1);
      parts.push(`created_at:<${next.toISOString().slice(0, 10)}`);
    }
    return parts;
  }

  // Shopify araması kaba filtre; kesin ayrım aşağıda etiketlere bakılarak yapılır.
  function buildQuery({ stage, from, to }) {
    if (stage === 'panel') return q(T.panel);
    // Drive etiketi kesme işareti içerdiği için aramada tarihe göre çekilip etikete kodda bakılır.
    if (stage === 'drive') return dateRange(from, to).join(' AND ');
    return `status:open AND (fulfillment_status:unfulfilled OR fulfillment_status:partial)` +
      OURS.map((t) => ` AND NOT ${q(t)}`).join('');
  }

  function inStage(stage, tags) {
    const has = (t) => tags.includes(t);
    if (stage === 'yeni') return !OURS.some(has);
    if (stage === 'panel') return has(T.panel) && !has(T.drive) && !has(T.eski);
    return has(T.drive) || has(T.eski);
  }

  function address(a) {
    if (!a) return '';
    return [a.address1, a.address2, a.city, a.province, a.zip].filter(Boolean).join(', ');
  }

  async function fetchRows({ stage = 'yeni', from, to } = {}) {
    checkConfig();
    const query = buildQuery({ stage, from, to });
    const rows = [];
    let cursor = null;
    do {
      const page = (await gql(ORDERS, { cursor, q: query })).orders;
      for (const o of page.nodes) {
        const tags = o.tags || [];
        if (!inStage(stage, tags)) continue;
        if (stage === 'yeni' && o.cancelledAt) continue;
        const customer = o.shippingAddress?.name || [o.customer?.firstName, o.customer?.lastName].filter(Boolean).join(' ');
        for (const li of o.lineItems.nodes) {
          // currentQuantity: siparişten çıkarılan / değiştirilen ürünlerde 0 olur.
          const current = li.currentQuantity ?? li.quantity;
          const qty = stage === 'drive' ? current : Math.min(li.unfulfilledQuantity, current);
          if (!qty) continue;
          const unit = toNumber(li.originalUnitPriceSet?.shopMoney?.amount);
          rows.push({
            _key: o.id,
            _ref: { gid: o.id },
            kanal: 'Shopify',
            siparisNo: o.name,
            tarih: toDate(o.createdAt),
            musteri: customer,
            telefon: o.shippingAddress?.phone || o.phone || '',
            adres: address(o.shippingAddress),
            kargoFirmasi: 'DHL eCommerce',
            kargoAnahtari: String(o.legacyResourceId),
            urun: li.title,
            sku: li.sku || '',
            adet: qty,
            tutar: unit != null ? unit * qty : null,
            odemeTipi: (o.paymentGatewayNames || []).join(', '),
            durum: o.cancelledAt ? 'İptal' : (DURUM[o.displayFulfillmentStatus] || o.displayFulfillmentStatus),
            etiketler: tags.filter((t) => !OURS.includes(t)).join(', '),
          });
        }
      }
      cursor = page.pageInfo.hasNextPage ? page.pageInfo.endCursor : null;
    } while (cursor);
    return rows;
  }

  // Siparişlere etiket ekler. orders: [{ key, siparisNo, rows }]
  async function addTag(orders, tag) {
    const done = [];
    const failed = [];
    for (const o of orders) {
      try {
        const res = await gql(TAGS_ADD, { id: o.rows[0]._ref.gid, tags: [tag] });
        const errs = res.tagsAdd.userErrors;
        if (errs.length) throw new Error(errs.map((e) => e.message).join(', '));
        done.push(o.key);
      } catch (e) {
        failed.push(`${o.siparisNo} (${e.message})`);
      }
    }
    return { done, failed };
  }

  // 1 → 2: panele çek
  const advance = (orders) => addTag(orders, T.panel);
  // 2 → 3: Drive'a aktarıldı
  const markExported = (orders) => addTag(orders, T.drive);

  return { fetchRows, advance, markExported, usesStore: false };
})();

// ======================================================================
// TRENDYOL
// ======================================================================
const trendyol = (() => {
  // yeni : Created paketler
  // panel: Picking / Invoiced paketler (Drive'a aktarılanlar sistem kaydından düşülür)
  // Panele çek = paketi "Picking" (Trendyol panelinde "İşleme Alındı") yapar.
  const cfg = config.trendyol;
  const { httpJson, pick, toNumber, toDate } = util;

  const BASE = 'https://apigw.trendyol.com/integration/order/sellers';
  const STATUS = { yeni: ['Created'], panel: ['Picking', 'Invoiced'] };
  const DURUM = { Created: 'Yeni', Picking: 'İşleme Alındı', Invoiced: 'Faturalandı' };
  const DAY = 24 * 60 * 60 * 1000;

  function headers() {
    if (!cfg.sellerId || !cfg.apiKey || !cfg.apiSecret) {
      throw new Error('Trendyol bilgileri eksik (SELLER_ID / API_KEY / API_SECRET)');
    }
    return {
      Authorization: 'Basic ' + Buffer.from(`${cfg.apiKey}:${cfg.apiSecret}`).toString('base64'),
      'User-Agent': `${cfg.sellerId} - SelfIntegration`,
      'Content-Type': 'application/json',
    };
  }

  async function fetchStatus(status, h) {
    const packages = [];
    const endDate = Date.now();
    const startDate = endDate - cfg.days * DAY;
    for (let page = 0, total = 1; page < total; page++) {
      const url = `${BASE}/${cfg.sellerId}/orders?status=${status}&startDate=${startDate}&endDate=${endDate}` +
        `&orderByField=PackageLastModifiedDate&orderByDirection=DESC&size=200&page=${page}`;
      const data = await httpJson(url, { headers: h }, `Trendyol (${status})`);
      packages.push(...(data.content || []));
      total = data.totalPages || 1;
    }
    return packages;
  }

  function address(a = {}) {
    return a.fullAddress || [a.address1, a.address2, a.district, a.city].filter(Boolean).join(', ');
  }

  async function fetchRows({ stage = 'yeni' } = {}) {
    const h = headers();
    const rows = [];
    for (const status of STATUS[stage] || []) {
      for (const p of await fetchStatus(status, h)) {
        const customer = pick(p, 'shipmentAddress.fullName') || [p.customerFirstName, p.customerLastName].filter(Boolean).join(' ');
        for (const line of p.lines || []) {
          rows.push({
            _key: `Trendyol:${p.id}`,
            _ref: { packageId: p.id, lineId: pick(line, 'id', 'lineId'), qty: toNumber(line.quantity) },
            kanal: 'Trendyol',
            siparisNo: String(pick(p, 'orderNumber')),
            tarih: toDate(p.orderDate),
            musteri: customer,
            telefon: pick(p, 'shipmentAddress.phone'),
            adres: address(p.shipmentAddress),
            kargoFirmasi: pick(p, 'cargoProviderName'),
            kargoAnahtari: String(pick(p, 'cargoTrackingNumber')),
            urun: pick(line, 'productName'),
            sku: pick(line, 'merchantSku', 'barcode'),
            adet: toNumber(line.quantity),
            tutar: toNumber(pick(line, 'amount', 'price')),
            odemeTipi: 'Trendyol',
            durum: DURUM[status] || status,
            etiketler: '',
          });
        }
      }
    }
    return rows;
  }

  // 1 → 2: paketleri "Picking" yapar
  async function advance(orders) {
    const h = headers();
    const done = [];
    const failed = [];
    for (const o of orders) {
      const packageId = o.rows[0]._ref.packageId;
      try {
        await httpJson(`${BASE}/${cfg.sellerId}/shipment-packages/${packageId}`, {
          method: 'PUT',
          headers: h,
          body: JSON.stringify({
            lines: o.rows.map((r) => ({ lineId: Number(r._ref.lineId), quantity: r._ref.qty })),
            params: {},
            status: 'Picking',
          }),
        }, 'Trendyol durum güncelleme');
        done.push(o.key);
      } catch (e) {
        failed.push(`${o.siparisNo} (${e.message})`);
      }
    }
    return { done, failed };
  }

  return { fetchRows, advance, usesStore: true };
})();

// ======================================================================
// HEPSIBURADA
// ======================================================================
const hepsiburada = (() => {
  // yeni : paketlenecek kalemler (sipariş numarasına göre gruplanır)
  // panel: paketlenmiş, kargoya verilmemiş paketler (Drive'a aktarılanlar sistem kaydından düşülür)
  // Panele çek = siparişin tüm kalemlerini tek pakette paketler ("Gönderime Hazır").
  // Not: Alan adları ilk gerçek çalıştırmada kontrol edilecek; pick() birden fazla olası adı dener.
  const cfg = config.hepsiburada;
  const { httpJson, pick, toNumber, toDate } = util;
  const BASE = 'https://oms-external.hepsiburada.com';

  function headers() {
    if (!cfg.merchantId || !cfg.password) throw new Error('Hepsiburada bilgileri eksik (MERCHANT_ID / PASSWORD)');
    return {
      Authorization: 'Basic ' + Buffer.from(`${cfg.username}:${cfg.password}`).toString('base64'),
      'User-Agent': cfg.userAgent || 'esse-jeffe-otomasyon',
      Accept: 'application/json',
      'Content-Type': 'application/json',
    };
  }

  async function paged(path, h, label) {
    const all = [];
    const limit = 100;
    for (let offset = 0; ; offset += limit) {
      const data = await httpJson(`${BASE}${path}?offset=${offset}&limit=${limit}`, { headers: h }, label);
      const items = Array.isArray(data) ? data : data.items || data.data || [];
      all.push(...items);
      if (items.length < limit) break;
    }
    return all;
  }

  function address(src) {
    const a = src.shippingAddress || src.deliveryAddress || {};
    return pick(a, 'address', 'fullAddress') ||
      [pick(a, 'address'), pick(a, 'district', 'town'), pick(a, 'city')].filter(Boolean).join(', ') ||
      pick(src, 'shippingAddressDetail', 'address');
  }

  function lineToRow(src, line, extra) {
    return {
      kanal: 'Hepsiburada',
      siparisNo: String(pick(line, 'orderNumber') || pick(src, 'orderNumber')),
      tarih: toDate(pick(line, 'orderDate') || pick(src, 'orderDate')),
      musteri: pick(src, 'recipientName', 'customerName', 'shippingAddress.name'),
      telefon: pick(src, 'phoneNumber', 'shippingAddress.phoneNumber'),
      adres: address(src),
      kargoFirmasi: pick(src, 'cargoCompany', 'cargoCompanyModel.name'),
      urun: pick(line, 'productName', 'name'),
      sku: pick(line, 'merchantSku', 'sku', 'hepsiburadaSku'),
      adet: toNumber(pick(line, 'quantity')),
      tutar: toNumber(pick(line, 'totalPrice', 'price')),
      odemeTipi: 'Hepsiburada',
      etiketler: '',
      ...extra,
    };
  }

  async function fetchRows({ stage = 'yeni' } = {}) {
    const h = headers();
    const rows = [];
    if (stage === 'yeni') {
      for (const item of await paged(`/orders/merchantid/${cfg.merchantId}`, h, 'Hepsiburada (paketlenecek)')) {
        const no = String(pick(item, 'orderNumber'));
        rows.push(lineToRow(item, item, {
          _key: `Hepsiburada:${no}`,
          _ref: { lineId: pick(item, 'id', 'lineItemId'), qty: toNumber(pick(item, 'quantity')) },
          kargoAnahtari: '',
          durum: 'Paketlenecek',
        }));
      }
    } else if (stage === 'panel') {
      for (const p of await paged(`/packages/merchantid/${cfg.merchantId}`, h, 'Hepsiburada (paketler)')) {
        const no = String(pick(p, 'packageNumber', 'id', 'barcode'));
        const code = String(pick(p, 'barcode', 'trackingNumber', 'cargoTrackingNumber', 'packageNumber'));
        for (const line of p.items || p.lines || [p]) {
          rows.push(lineToRow(p, line, { _key: `Hepsiburada:paket:${no}`, _ref: {}, kargoAnahtari: code, durum: 'Gönderime Hazır' }));
        }
      }
    }
    return rows;
  }

  // 1 → 2: siparişin tüm kalemlerini paketler
  async function advance(orders) {
    const h = headers();
    const done = [];
    const failed = [];
    for (const o of orders) {
      try {
        await httpJson(`${BASE}/packages/merchantid/${cfg.merchantId}`, {
          method: 'POST',
          headers: h,
          body: JSON.stringify({ lineItemRequests: o.rows.map((r) => ({ id: r._ref.lineId, quantity: r._ref.qty })) }),
        }, 'Hepsiburada paketleme');
        done.push(o.key);
      } catch (e) {
        failed.push(`${o.siparisNo} (${e.message})`);
      }
    }
    return { done, failed };
  }

  return { fetchRows, advance, usesStore: true };
})();

// ======================================================================
// KAYIT (Trendyol/Hepsiburada aktarım kaydı)
// ======================================================================
const store = (() => {
  // Trendyol / Hepsiburada'da Drive'a aktarılan siparişlerin kaydı (Shopify bunu etiketle tutar).
  // Railway'de kalıcı olması için DATA_DIR bir volume'a bağlanmalı (/data).
  const fs = require('fs');
  const path = require('path');
  const file = path.join(config.dataDir, 'aktarilanlar.json');
  let cache = null;

  function load() {
    if (cache) return cache;
    try { cache = JSON.parse(fs.readFileSync(file, 'utf8')); } catch { cache = {}; }
    return cache;
  }

  function save() {
    fs.mkdirSync(config.dataDir, { recursive: true });
    const tmp = file + '.tmp';
    fs.writeFileSync(tmp, JSON.stringify(cache));
    fs.renameSync(tmp, file);
  }

  const has = (key) => !!load()[key];

  function add(orders) {
    const db = load();
    const now = new Date().toISOString();
    for (const o of orders) {
      const { rows, ...rest } = o;
      db[o.key] = { ...rest, aktarimTarihi: now };
    }
    save();
  }

  // Sipariş tarihine göre (yoksa aktarım tarihine göre) aralık; YYYY-MM-DD, İstanbul saati
  function list({ from, to, kanallar }) {
    const day = (d) => new Date(d).toLocaleDateString('en-CA', { timeZone: 'Europe/Istanbul' });
    return Object.values(load()).filter((o) =>
      kanallar.includes(o.kanal) &&
      (!from || day(o.tarih || o.aktarimTarihi) >= from) && (!to || day(o.tarih || o.aktarimTarihi) <= to));
  }

  return { has, add, list };
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
    { header: 'Adres', key: 'adres', width: 40 },
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
      telefon: o.telefon, adres: o.adres || '', kargoFirmasi: o.kargoFirmasi, kargoAnahtari: o.kargoAnahtari,
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
    ['siparisNo', 'kargoAnahtari', 'telefon', 'adres'].forEach((k) => (ws.getColumn(k).numFmt = '@'));
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
// GOOGLE SHEETS (her aktarım tablonun başına yeni sekme olarak eklenir)
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

  return { writeOrders, configured: () => !!(cfg.sheetId && cfg.serviceAccount) };
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
// SİPARİŞ AKIŞI (3 aşama)
// ======================================================================
const flow = (() => {
  // Üç aşama:
  //  yeni  → [Panele çek]   → panel  (Shopify etiket, Trendyol İşleme Alındı, Hepsiburada paketle)
  //  panel → [Drive'a aktar] → drive (Google Sheet'e yaz, Shopify etiket / pazaryeri kaydı)
  //  drive : tarih aralığıyla görüntüleme ve Excel indirme
  const { buildWorkbook } = excel;
  const { sendMail } = mailer;
  const ALL = { shopify: ['Shopify', shopify], trendyol: ['Trendyol', trendyol], hepsiburada: ['Hepsiburada', hepsiburada] };
  const active = () => config.sources.filter((k) => ALL[k]).map((k) => ALL[k]);
  const byName = Object.fromEntries(Object.values(ALL));

  const sessions = new Map();
  function remember(data) {
    const id = crypto.randomUUID();
    sessions.set(id, data);
    setTimeout(() => sessions.delete(id), 60 * 60 * 1000).unref();
    return id;
  }

  // Ürün satırlarını siparişlere toplar; SKU'lar alt alta listelenir.
  function group(rows) {
    const map = new Map();
    for (const r of rows) {
      if (!map.has(r._key)) {
        map.set(r._key, {
          key: r._key, kanal: r.kanal, siparisNo: r.siparisNo, tarih: r.tarih, musteri: r.musteri,
          telefon: r.telefon, adres: r.adres || '', kargoFirmasi: r.kargoFirmasi, kargoAnahtari: r.kargoAnahtari,
          odemeTipi: r.odemeTipi, durum: r.durum, etiketler: r.etiketler || '',
          tutar: 0, adet: 0, skus: [], rows: [],
        });
      }
      const o = map.get(r._key);
      o.rows.push(r);
      o.skus.push({ sku: r.sku || r.urun || '(SKU yok)', adet: r.adet });
      o.adet += r.adet || 0;
      o.tutar += r.tutar || 0;
    }
    return [...map.values()];
  }

  const publicView = ({ rows, ...rest }) => rest;
  const sortDesc = (list) => list.sort((a, b) => new Date(b.tarih || 0) - new Date(a.tarih || 0));

  async function list({ stage = 'yeni', from, to } = {}) {
    const sources = active();
    const warnings = [];
    let orders = [];

    const results = await Promise.allSettled(sources.map(async ([name, src]) => {
      if (stage === 'drive' && src.usesStore) {
        return store.list({ from, to, kanallar: [name] }).map((o) => ({ ...o, rows: [] }));
      }
      let grouped = group(await src.fetchRows({ stage, from, to }));
      if (stage === 'panel' && src.usesStore) grouped = grouped.filter((o) => !store.has(o.key));
      return grouped;
    }));
    results.forEach((r, i) => {
      if (r.status === 'fulfilled') orders.push(...r.value);
      else {
        warnings.push(`${sources[i][0]}: ${r.reason.message}`);
        console.error(`[${sources[i][0]}]`, r.reason);
      }
    });

    orders = sortDesc(orders);
    const sessionId = remember({ stage, from, to, orders: new Map(orders.map((o) => [o.key, o])) });
    return { sessionId, stage, orders: orders.map(publicView), warnings };
  }

  function take(sessionId, keys, stage) {
    const s = sessions.get(sessionId);
    if (!s) throw new Error('Liste süresi doldu, siparişleri yeniden getirin.');
    if (stage && s.stage !== stage) throw new Error('Bu işlem bu sekmede yapılamaz.');
    const selected = (keys || []).map((k) => s.orders.get(k)).filter(Boolean);
    if (!selected.length) throw new Error('Seçilen siparişler bu listeden zaten işlendi; listeyi yeniden getirin.');
    return { s, selected };
  }

  const byChannel = (orders) => orders.reduce((m, o) => ((m[o.kanal] ||= []).push(o), m), {});

  // 1 → 2
  async function advance({ sessionId, keys }) {
    const { s, selected } = take(sessionId, keys, 'yeni');
    const summary = {};
    const failed = [];
    for (const [kanal, orders] of Object.entries(byChannel(selected))) {
      const res = await byName[kanal].advance(orders);
      summary[kanal] = res.done.length;
      failed.push(...res.failed.map((f) => `${kanal} ${f}`));
      res.done.forEach((k) => s.orders.delete(k));
    }
    return { summary, failed, done: Object.values(summary).reduce((a, b) => a + b, 0) };
  }

  function fileName(prefix) {
    const p = (n) => String(n).padStart(2, '0');
    const tr = new Date(new Date().toLocaleString('en-US', { timeZone: 'Europe/Istanbul' }));
    return `${prefix}_${tr.getFullYear()}-${p(tr.getMonth() + 1)}-${p(tr.getDate())}_${p(tr.getHours())}${p(tr.getMinutes())}.xlsx`;
  }

  async function mailList({ title, orders, buffer, filename, sheet, warnings }) {
    const sheetLine = sheet && sheet.url ? `<p><a href="${sheet.url}">Google Sheet'te aç</a> (sekme: ${sheet.title})</p>` : '';
    const warn = warnings.length ? `<p style="color:#b00020"><b>Uyarı:</b><br>${warnings.join('<br>')}</p>` : '';
    try {
      return await sendMail({
        subject: `${title} – ${orders.length} sipariş`,
        html: `<p>${title}: <b>${orders.length}</b> sipariş. Excel ekte.</p>${sheetLine}${warn}`,
        attachments: [{ filename, content: buffer }],
      });
    } catch (e) {
      console.error(e);
      return { error: e.message };
    }
  }

  // 2 → 3
  async function exportDrive({ sessionId, keys, email = true }) {
    if (!sheets.configured()) throw new Error('Google Sheet ayarları eksik (GOOGLE_SHEET_ID / GOOGLE_SERVICE_ACCOUNT_JSON).');
    const { s, selected } = take(sessionId, keys, 'panel');
    const channels = active().map(([n]) => n);
    const warnings = [];

    // Önce Sheet'e yazılır; yazılamazsa hiçbir sipariş 3. aşamaya geçmez.
    const sheet = await sheets.writeOrders(selected, 'Liste');

    const groups = byChannel(selected);
    for (const [kanal, orders] of Object.entries(groups)) {
      if (byName[kanal].usesStore) {
        store.add(orders);
        orders.forEach((o) => s.orders.delete(o.key));
      } else {
        const res = await byName[kanal].markExported(orders);
        res.done.forEach((k) => s.orders.delete(k));
        if (res.failed.length) {
          warnings.push(`${kanal} etiketlenemeyen siparişler (Sheet'e yazıldı, panelde 2. aşamada kalır): ${res.failed.join('; ')}`);
        }
      }
    }

    const buffer = await buildWorkbook(selected, warnings, channels);
    const filename = fileName('siparisler');
    const mail = email ? await mailList({ title: "Drive'a aktarılan siparişler", orders: selected, buffer, filename, sheet, warnings }) : { skipped: true };
    return { orderCount: selected.length, buffer, filename, sheet, mail, warnings };
  }

  // Herhangi bir sekmeden seçilenleri sadece Excel olarak almak (durum değiştirmez)
  async function download({ sessionId, keys, email = false }) {
    const { s, selected } = take(sessionId, keys);
    const channels = active().map(([n]) => n);
    const buffer = await buildWorkbook(selected, [], channels);
    const filename = fileName(s.stage === 'drive' ? 'aktarilanlar' : 'liste');
    const mail = email ? await mailList({ title: 'Sipariş listesi', orders: selected, buffer, filename, warnings: [] }) : { skipped: true };
    return { orderCount: selected.length, buffer, filename, mail, warnings: [] };
  }

  return { list, advance, exportDrive, download };
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
  .wrap{max-width:1900px;margin:0 auto}
  h1{font-size:22px;margin:0 0 16px}
  .tabs{display:flex;gap:6px;margin-bottom:-1px;flex-wrap:wrap}
  .tab{padding:11px 18px;border:1px solid var(--line);border-bottom:0;border-radius:10px 10px 0 0;background:#efe8e3;cursor:pointer;font-weight:600;color:var(--muted)}
  .tab.on{background:#fff;color:var(--ink)}
  .tab .n{display:inline-block;min-width:22px;margin-left:6px;padding:1px 7px;border-radius:999px;background:#e2d6cf;font-size:12px;text-align:center}
  .card{background:#fff;border:1px solid var(--line);border-radius:0 12px 12px 12px;padding:20px}
  .bar{display:flex;flex-wrap:wrap;gap:10px;align-items:center;margin-bottom:10px}
  button{padding:12px 18px;font-size:15px;font-weight:600;border:0;border-radius:10px;background:var(--brand);color:#fff;cursor:pointer}
  button.ghost{background:#efe8e3;color:var(--ink)}
  button:disabled{opacity:.5;cursor:not-allowed}
  input[type=date]{padding:10px;border:1px solid var(--line);border-radius:8px;font:inherit}
  label.chk{display:flex;align-items:center;gap:6px;color:var(--muted)}
  .scroll{overflow:auto;max-height:calc(100vh - 250px);border:1px solid var(--line);border-radius:10px}
  table{border-collapse:collapse;width:100%;font-size:15px;min-width:1400px}
  th,td{padding:12px 14px;border-bottom:1px solid var(--line);text-align:left;vertical-align:top;line-height:1.45}
  th{background:#f3ece7;font-weight:700;position:sticky;top:0;z-index:1;white-space:nowrap}
  tbody tr:nth-child(even) td{background:var(--zebra)}
  tbody tr:hover td{background:#f5ede8}
  tr.off td{opacity:.4}
  td.num{text-align:right;white-space:nowrap}
  td.nowrap{white-space:nowrap}
  td.addr{min-width:240px;max-width:340px}
  .sku{display:block;font-family:ui-monospace,Menlo,monospace;font-size:14px;white-space:nowrap}
  .mono{font-family:ui-monospace,Menlo,monospace;font-size:14px}
  .tags{color:#5b4a4f;min-width:200px}
  input[type=checkbox]{width:18px;height:18px}
  .empty{padding:40px;text-align:center;color:var(--muted)}
  #hint{font-size:14px;color:var(--muted);margin-bottom:12px}
  #msg{margin-top:12px;line-height:1.7}
  .warn{color:var(--warn)} .ok{color:var(--ok)}
  a.dl{font-weight:700;color:var(--brand);margin-right:16px}
  .hide{display:none}
</style></head>
<body><div class="wrap">
  <h1>Esse Jeffe Otomasyon</h1>
  <div class="tabs">
    <div class="tab on" data-stage="yeni">1 · Yeni gelen siparişler<span class="n" id="n-yeni">–</span></div>
    <div class="tab" data-stage="panel">2 · Panele çekilenler<span class="n" id="n-panel">–</span></div>
    <div class="tab" data-stage="drive">3 · Etiket oluşturulanlar</div>
  </div>
  <div class="card">
    <div class="bar">
      <span id="range" class="hide"><input type="date" id="bas"> – <input type="date" id="bit"></span>
      <button id="fetch" class="ghost">Yenile</button>
      <span style="flex:1"></span>
      <label class="chk" id="mailBox"><input type="checkbox" id="mail" checked> E-posta da gönder</label>
      <button id="excel" class="ghost" disabled>Excel indir</button>
      <button id="action" disabled>Panele çek</button>
    </div>
    <div id="hint"></div>
    <div class="scroll"><table>
      <thead><tr>
        <th><input type="checkbox" id="all"></th><th>Kanal</th><th>Sipariş</th><th>Tarih</th><th>Müşteri</th><th>Telefon</th>
        <th>Adres</th><th>SKU</th><th>Tutar</th><th>Ödeme</th><th>Kargo anahtarı</th><th>Durum</th><th>Etiketler</th>
      </tr></thead>
      <tbody id="rows"></tbody>
    </table></div>
    <div id="msg"></div>
  </div>
</div>
<script>
const COLS = 13;
const STAGES = {
  yeni:  { action: 'Panele çek', hint: 'Hiç dokunulmamış siparişler. "Panele çek": Shopify\\'da "etiket oluşturuldu - otomatik" etiketi eklenir, Trendyol\\'da "İşleme Alındı" yapılır, Hepsiburada\\'da paketlenip "Gönderime Hazır"a geçer.' },
  panel: { action: "Drive'a aktar", hint: 'Panele çekilmiş siparişler. "Drive\\'a aktar": seçilenler Google Sheet\\'e yeni bir sekme olarak yazılır ve 3. aşamaya geçer (Shopify\\'da "drive\\'a aktarıldı - otomatik" etiketi eklenir).' },
  drive: { action: null, hint: 'Drive\\'a aktarılmış siparişler (sipariş tarihine göre). Bu sekmede hiçbir siparişin durumu değişmez.' },
};
let stage = 'yeni', session = null, orders = [];
const $ = (id) => document.getElementById(id);
const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const money = (n) => n ? n.toLocaleString('tr-TR', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + ' ₺' : '';
const when = (d) => d ? new Date(d).toLocaleString('tr-TR', { day:'2-digit', month:'2-digit', hour:'2-digit', minute:'2-digit' }) : '';
const iso = (d) => d.toISOString().slice(0, 10);
const empty = (t) => '<tr><td colspan="' + COLS + '" class="empty">' + t + '</td></tr>';
$('bas').value = iso(new Date(Date.now() - 6 * 864e5)); $('bit').value = iso(new Date());

function setStage(s) {
  stage = s; session = null; orders = [];
  document.querySelectorAll('.tab').forEach((x) => x.classList.toggle('on', x.dataset.stage === s));
  $('range').classList.toggle('hide', s !== 'drive');
  $('action').classList.toggle('hide', !STAGES[s].action);
  $('action').dataset.label = STAGES[s].action || '';
  $('mailBox').classList.toggle('hide', s === 'yeni');
  $('mail').checked = s === 'panel';
  $('fetch').textContent = s === 'drive' ? 'Getir' : 'Yenile';
  $('hint').textContent = STAGES[s].hint;
  $('msg').innerHTML = '';
  $('rows').innerHTML = empty(s === 'drive' ? 'Tarih aralığı seçip "Getir"e basın.' : 'Yükleniyor…');
  update();
  if (s !== 'drive') load();
}
document.querySelectorAll('.tab').forEach((t) => t.onclick = () => setStage(t.dataset.stage));

function selected() { return [...document.querySelectorAll('.pick:checked')].map((c) => c.value); }
function update() {
  const n = selected().length;
  const label = $('action').dataset.label;
  $('action').disabled = !n; $('excel').disabled = !n;
  $('action').textContent = n ? label + ' (' + n + ')' : label;
  $('excel').textContent = n ? 'Excel indir (' + n + ')' : 'Excel indir';
  $('all').checked = n && n === orders.length;
  document.querySelectorAll('.pick').forEach((c) => c.closest('tr').classList.toggle('off', !c.checked));
}
$('all').onchange = (e) => { document.querySelectorAll('.pick').forEach((c) => c.checked = e.target.checked); update(); };

function render() {
  if (stage !== 'drive') $('n-' + stage).textContent = orders.length;
  if (!orders.length) {
    $('rows').innerHTML = empty({ yeni: 'Yeni sipariş yok.', panel: 'Panele çekilmiş sipariş yok.', drive: 'Bu tarih aralığında aktarılmış sipariş yok.' }[stage]);
    return update();
  }
  $('rows').innerHTML = orders.map((o) =>
    '<tr><td><input type="checkbox" class="pick" value="' + esc(o.key) + '" checked></td>' +
    '<td>' + esc(o.kanal) + '</td><td class="nowrap"><b>' + esc(o.siparisNo) + '</b></td><td class="nowrap">' + when(o.tarih) + '</td>' +
    '<td>' + esc(o.musteri) + '</td><td class="nowrap">' + esc(o.telefon) + '</td><td class="addr">' + esc(o.adres) + '</td>' +
    '<td>' + o.skus.map((x) => '<span class="sku">' + esc(x.sku) + (x.adet > 1 ? ' ×' + esc(x.adet) : '') + '</span>').join('') + '</td>' +
    '<td class="num">' + money(o.tutar) + '</td><td>' + esc(o.odemeTipi) + '</td>' +
    '<td class="mono">' + esc(o.kargoAnahtari) + '</td><td>' + esc(o.durum) + '</td><td class="tags">' + esc(o.etiketler) + '</td></tr>'
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
const post = (url, body) => call(url, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body) });
const warnHtml = (w) => w && w.length ? '<p class="warn">' + w.map(esc).join('<br>') + '</p>' : '';

async function load() {
  $('fetch').disabled = true;
  const my = stage;
  try {
    const q = my === 'drive' ? '&bas=' + $('bas').value + '&bit=' + $('bit').value : '';
    const d = await call('/api/liste?asama=' + my + q);
    if (my !== stage) return;
    session = d.sessionId; orders = d.orders; render();
    $('msg').innerHTML = warnHtml(d.warnings);
  } catch (e) { if (my === stage) { $('rows').innerHTML = empty(''); $('msg').innerHTML = warnHtml([e.message]); } }
  finally { $('fetch').disabled = false; }
}
$('fetch').onclick = load;

function countBy(keys) {
  const c = {};
  orders.filter((o) => keys.includes(o.key)).forEach((o) => c[o.kanal] = (c[o.kanal] || 0) + 1);
  return c;
}
function busy(btn, on) { btn.disabled = on; if (on) btn.textContent = 'İşleniyor…'; else update(); }
function mailText(m) {
  if (!m) return '';
  return m.sent ? 'E-postanıza gönderildi.' : m.error ? '<span class="warn">E-posta gönderilemedi: ' + esc(m.error) + '</span>' : '';
}

$('action').onclick = async () => {
  const keys = selected();
  const c = countBy(keys);
  const what = { Shopify: 'etiketlenecek', Trendyol: '"İşleme Alındı" yapılacak', Hepsiburada: 'paketlenip "Gönderime Hazır"a geçecek' };
  const text = stage === 'yeni'
    ? 'Panele çekilecek:\\n' + Object.entries(c).map(([k, v]) => '• ' + k + ': ' + v + ' sipariş ' + what[k]).join('\\n') + '\\n\\nTrendyol ve Hepsiburada işlemleri geri alınamaz. Devam edilsin mi?'
    : keys.length + " sipariş Google Sheet'e yazılacak ve 3. aşamaya geçecek. Devam edilsin mi?";
  if (!confirm(text)) return;
  busy($('action'), true);
  try {
    if (stage === 'yeni') {
      const d = await post('/api/panele-cek', { sessionId: session, keys });
      const parts = Object.entries(d.summary).map(([k, v]) => k + ': ' + v);
      $('msg').innerHTML = '<span class="ok">' + d.done + ' sipariş panele çekildi.</span> ' + esc(parts.join(' · ')) + warnHtml(d.failed);
    } else {
      const d = await post('/api/drive', { sessionId: session, keys, email: $('mail').checked });
      $('msg').innerHTML = '<span class="ok">' + d.orderCount + " sipariş Drive'a aktarıldı.</span> " + mailText(d.mail) + warnHtml(d.warnings) +
        '<br>' + (d.sheetUrl ? '<a class="dl" href="' + esc(d.sheetUrl) + '" target="_blank" rel="noopener">↗ Google Sheet\\'te aç</a>' : '') +
        '<a class="dl" href="' + d.download + '">⬇ Excel\\'i indir</a>';
    }
    const done = new Set(keys);
    const msg = $('msg').innerHTML;
    await load();
    $('msg').innerHTML = msg + $('msg').innerHTML;
  } catch (e) { $('msg').innerHTML = warnHtml([e.message]); }
  finally { busy($('action'), false); }
};

$('excel').onclick = async () => {
  busy($('excel'), true);
  try {
    const d = await post('/api/excel', { sessionId: session, keys: selected(), email: stage !== 'yeni' && $('mail').checked });
    $('msg').innerHTML = '<span class="ok">' + d.orderCount + ' sipariş Excel\\'e alındı.</span> ' + mailText(d.mail) + '<br><a class="dl" href="' + d.download + '">⬇ Excel\\'i indir</a>';
    window.location.href = d.download;
  } catch (e) { $('msg').innerHTML = warnHtml([e.message]); }
  finally { busy($('excel'), false); }
};

setStage('yeni');
call('/api/liste?asama=panel').then((d) => $('n-panel').textContent = d.orders.length).catch(() => {});
</script></body></html>`;

// ======================================================================
// SUNUCU
// ======================================================================
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
    const same = (a, b) => a.length === b.length && crypto.timingSafeEqual(Buffer.from(a), Buffer.from(b));
    if (same(u, config.panel.user) && same(p, config.panel.password)) return next();
  }
  res.set('WWW-Authenticate', 'Basic realm="Esse Jeffe Otomasyon"').status(401).send('Giriş gerekli');
}

app.get('/health', (_req, res) => res.send('ok')); // Railway sağlık kontrolü
app.use(auth);
app.use(express.json());

// Oluşturulan Excel dosyaları 1 saat boyunca panelden indirilebilir.
const files = new Map();
const keep = (buffer, filename) => {
  const id = crypto.randomUUID();
  files.set(id, { buffer, filename });
  setTimeout(() => files.delete(id), 60 * 60 * 1000).unref();
  return `/indir/${id}`;
};

const handle = (fn) => async (req, res) => {
  try {
    res.json(await fn(req));
  } catch (e) {
    console.error(e);
    res.status(e.status || 500).json({ error: e.message });
  }
};

// Durum değiştiren işlemler aynı anda iki kez çalışmasın.
let busy = false;
const exclusive = (fn) => async (req) => {
  if (busy) throw Object.assign(new Error('Başka bir işlem sürüyor, birkaç saniye bekleyin.'), { status: 409 });
  busy = true;
  try { return await fn(req); } finally { busy = false; }
};

const isDate = (v) => /^\d{4}-\d{2}-\d{2}$/.test(v || '');

app.get('/api/liste', handle(async (req) => {
  const stage = ['yeni', 'panel', 'drive'].includes(req.query.asama) ? req.query.asama : 'yeni';
  const from = isDate(req.query.bas) ? req.query.bas : undefined;
  const to = isDate(req.query.bit) ? req.query.bit : undefined;
  if (stage === 'drive' && (!from || !to)) throw new Error('Başlangıç ve bitiş tarihi seçin.');
  return flow.list({ stage, from, to });
}));

app.post('/api/panele-cek', handle(exclusive(async (req) => {
  const { sessionId, keys } = req.body || {};
  return flow.advance({ sessionId, keys });
})));

app.post('/api/drive', handle(exclusive(async (req) => {
  const { sessionId, keys, email } = req.body || {};
  const out = await flow.exportDrive({ sessionId, keys, email: email !== false });
  return {
    orderCount: out.orderCount, warnings: out.warnings, mail: out.mail,
    sheetUrl: out.sheet && out.sheet.url, download: keep(out.buffer, out.filename),
  };
})));

app.post('/api/excel', handle(async (req) => {
  const { sessionId, keys, email } = req.body || {};
  const out = await flow.download({ sessionId, keys, email: !!email });
  return { orderCount: out.orderCount, mail: out.mail, download: keep(out.buffer, out.filename) };
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

