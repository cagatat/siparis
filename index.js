// Esse Jeffe Otomasyon — tüm uygulama tek dosyada.
// Bölümler: Ayarlar · Yardımcılar · Shopify · Trendyol · Hepsiburada · Kayıt · Adres kontrolü ·
//           Excel · Google Sheets · E-posta · Sipariş akışı · Otomatik adres kontrolü · Panel sayfası · Sunucu
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
    // Trendyol/Hepsiburada'da Excel'e aktarılanların kaydı burada tutulur (Railway volume yolu: /data)
    dataDir: env('DATA_DIR', './data'),
    panel: { user: env('PANEL_USER'), password: env('PANEL_PASSWORD') },
    shopify: {
      // Site (essejeffe.com) bağlıysa siparişler Shopify yerine siteden okunur. Site, Shopify'ın GraphQL biçiminde cevap verir
      // (dilara_kopru.py, POST /api/dilara/gql); etiketler, aktarım listesi ve kargo anahtarı aynı mantıkla çalışır.
      siteUrl: env('MAGAZA_SITE_URL').replace(/\/+$/, ''),
      siteAnahtar: env('SITE_ANAHTAR'),
      store: env('SHOPIFY_STORE').replace('.myshopify.com', ''),
      clientId: env('SHOPIFY_CLIENT_ID'),
      clientSecret: env('SHOPIFY_CLIENT_SECRET'),
      accessToken: env('SHOPIFY_ACCESS_TOKEN'),
      apiVersion: env('SHOPIFY_API_VERSION', '2026-07'),
      tags: {
        excel: env('SHOPIFY_ETIKET_EXCEL', 'excel'),                           // Excel'e aktarılan siparişler
        // Önceki sürümlerin etiketleri: "drive" ve "eski" aktarılmış sayılır, "panel" bekleyen sayılır.
        panel: 'etiket oluşturuldu - otomatik',
        drive: "drive'a aktarıldı - otomatik",
        eski: 'etiketi çıkarıldı - otomatik',
        adres: env('SHOPIFY_ETIKET_ADRES', 'adres düzeltildi - otomatik'),    // adres kontrolü değişiklik yaptıysa
      },
    },
    // Adres kontrolü: Google Maps anahtarı varsa o, yoksa OpenStreetMap kullanılır.
    adres: {
      // Ana anahtar: "evet" olmadıkça adreslere hiçbir yerde dokunulmaz, kontrol de yapılmaz
      duzeltme: env('ADRES_DUZELTME', 'hayir').toLowerCase() === 'evet',
      otomatik: env('ADRES_OTOMATIK', 'evet').toLowerCase() !== 'hayir', // yeni siparişlerde otomatik düzeltme
      aralikDk: Math.max(1, Number(env('ADRES_ARALIK_DK', '5')) || 5),
      googleKey: env('GOOGLE_MAPS_API_KEY'),
      email: env('ADRES_EMAIL') || env('MAIL_TO').split(',')[0].trim(),
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
      tab: env('GOOGLE_SHEET_SAYFA', 'Siparişler'),   // tüm aktarımların yazıldığı tek sayfa
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

  // Okunur ödeme yöntemi (yalnız ekranda; Excel/Sheet'teki "Ödeme" kolonu değişmez).
  // Dilara'nın taslaktan girdiği kapıda siparişlerde altyapı "manual", ödeme bekliyor → Kapıda.
  function odemeYontemi(gateways = [], finansal = '', tags = []) {
    const g = gateways.join(' ').toLowerCase();               // altyapı adları İngilizce ("CASH ON DELIVERY")
    const t = tags.join(' ').toLocaleLowerCase('tr');         // etiketler Türkçe ("İade", "Ödeme bekleniyor")
    if (/de[ğg]i[şs]im/.test(t)) return 'Değişim';
    if (/\biade\b/.test(t)) return 'İade';
    if (/cash on delivery|\bcod\b|kap[ıi]da|teslimatta/.test(g)) return 'Kapıda';
    if (/havale|eft|bank|transfer/.test(g) || /[öo]deme bekleniyor|dekont|havale/.test(t)) return finansal === 'PAID' ? 'Havale/EFT (ödendi)' : 'Havale/EFT';
    if (/paytr|kredi|card|kart|shopify_payments|iyzico/.test(g)) return 'Kart';
    if (finansal === 'PAID') return 'Ödendi';
    if (!g || /manual/.test(g)) return ['PENDING', 'PARTIALLY_PAID', 'AUTHORIZED'].includes(finansal) ? 'Kapıda' : (gateways.join(', ') || 'Belirsiz');
    return gateways.join(', ');
  }

  return { pick, toNumber, toDate, httpJson, odemeYontemi };
})();

// ======================================================================
// SHOPIFY
// ======================================================================
const shopify = (() => {
  // Durum tek etiketle izlenir:
  //  bekleyen  : açık, gönderilmemiş, "excel" etiketi yok
  //  aktarilan : "excel" etiketi (veya önceki sürümlerin aktarım etiketi) var
  // Kargo anahtarı = siparişin uzun sistem ID'si (legacyResourceId).
  const cfg = config.shopify;
  const T = cfg.tags;
  const DONE = [T.excel, T.drive, T.eski];                 // aktarılmış sayılan etiketler
  const HIDDEN = [T.excel, T.panel, T.drive, T.eski, T.adres]; // Etiketler kolonunda gösterilmeyenler
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
    if (cfg.siteUrl) {
      const data = await httpJson(
        `${cfg.siteUrl}/api/dilara/gql`,
        {
          method: 'POST',
          headers: { 'Content-Type': 'application/json', 'x-site-anahtar': cfg.siteAnahtar },
          body: JSON.stringify({ query, variables }),
        },
        'Site'
      );
      if (data.errors) throw new Error('Site: ' + JSON.stringify(data.errors).slice(0, 300));
      return data.data;
    }
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

  const ORDER_FIELDS = `
        id legacyResourceId name createdAt cancelledAt tags
        displayFulfillmentStatus displayFinancialStatus paymentGatewayNames phone
        customer { firstName lastName }
        shippingAddress { name firstName lastName company phone address1 address2 city province provinceCode zip countryCodeV2 }
        lineItems(first: 50) {
          nodes {
            title sku quantity currentQuantity unfulfilledQuantity
            originalUnitPriceSet { shopMoney { amount } }
          }
        }`;
  const ORDERS = `
  query Orders($cursor: String, $q: String!) {
    orders(first: 50, after: $cursor, sortKey: CREATED_AT, reverse: true, query: $q) {
      pageInfo { hasNextPage endCursor }
      nodes { ${ORDER_FIELDS} }
    }
  }`;

  const ORDER_UPDATE = `
  mutation OrderUpdate($input: OrderInput!) {
    orderUpdate(input: $input) { userErrors { field message } }
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
    if (cfg.siteUrl) {
      if (!cfg.siteAnahtar) throw new Error('Site bilgileri eksik (MAGAZA_SITE_URL + SITE_ANAHTAR)');
      return;
    }
    if (!cfg.store || !(cfg.accessToken || (cfg.clientId && cfg.clientSecret))) {
      throw new Error('Shopify bilgileri eksik (SHOPIFY_STORE + CLIENT_ID/SECRET)');
    }
  }

  const q = (t) => `tag:"${t.replace(/"/g, '')}"`;

  // Shopify araması kaba filtre; kesin ayrım aşağıda etiketlere bakılarak yapılır.
  // (Eski Drive etiketi kesme işareti içerdiği için aramaya konmaz, kodda elenir.)
  function buildQuery() {
    return `status:open AND (fulfillment_status:unfulfilled OR fulfillment_status:partial)` +
      [T.excel, T.eski].map((t) => ` AND NOT ${q(t)}`).join('');
  }

  function inStage(stage, tags) {
    const done = DONE.some((t) => tags.includes(t));
    return stage === 'aktarilan' ? done : !done;
  }

  function address(a) {
    if (!a) return '';
    return [a.address1, a.address2, a.city, a.province, a.zip].filter(Boolean).join(', ');
  }

  // Bir siparişin ürün satırları (stage: 'bekleyen' → gönderilmemiş adet, 'aktarilan' → güncel adet)
  async function orderRows(o, stage, rows) {
        const tags = o.tags || [];
        const customer = o.shippingAddress?.name || [o.customer?.firstName, o.customer?.lastName].filter(Boolean).join(' ');
        // Adres kontrolü (sadece okuma; düzeltmeyi otomatik adres kontrolü yapar)
        const ak = stage === 'aktarilan' || !config.adres.duzeltme ? null : await adres.check(o.shippingAddress);
        for (const li of o.lineItems.nodes) {
          // currentQuantity: siparişten çıkarılan / değiştirilen ürünlerde 0 olur.
          const current = li.currentQuantity ?? li.quantity;
          const qty = stage === 'aktarilan' ? current : Math.min(li.unfulfilledQuantity, current);
          if (!qty) continue;
          const unit = toNumber(li.originalUnitPriceSet?.shopMoney?.amount);
          rows.push({
            _key: o.id,
            _ref: { gid: o.id, address: o.shippingAddress },
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
            odeme: util.odemeYontemi(o.paymentGatewayNames || [], o.displayFinancialStatus, tags),
            durum: o.cancelledAt ? 'İptal' : (DURUM[o.displayFulfillmentStatus] || o.displayFulfillmentStatus),
            etiketler: tags.filter((t) => !HIDDEN.includes(t)).join(', '),
            adresDurum: ak ? ak.durum : '',
            adresNot: ak ? ak.notlar.join(' · ') : '',
            adresDuzeltildi: tags.includes(T.adres),
          });
        }
  }

  async function fetchRows({ stage = 'bekleyen' } = {}) {
    checkConfig();
    const query = buildQuery();
    const rows = [];
    let cursor = null;
    do {
      const page = (await gql(ORDERS, { cursor, q: query })).orders;
      for (const o of page.nodes) {
        if (!inStage(stage, o.tags || [])) continue;
        if (stage === 'bekleyen' && o.cancelledAt) continue;
        await orderRows(o, stage, rows);
      }
      cursor = page.pageInfo.hasNextPage ? page.pageInfo.endCursor : null;
    } while (cursor);
    return rows;
  }

  // Bir aktarımdaki siparişler (kimlikleriyle; durum ne olursa olsun)
  async function fetchByIds(ids) {
    checkConfig();
    const rows = [];
    for (let i = 0; i < ids.length; i += 100) {
      const d = await gql(`query($ids: [ID!]!) { nodes(ids: $ids) { ... on Order { ${ORDER_FIELDS} } } }`, { ids: ids.slice(i, i + 100) });
      for (const o of d.nodes || []) if (o && o.id) await orderRows(o, 'aktarilan', rows);
    }
    return rows;
  }

  // Aktarım listesi (numara, zaman, siparişler) uygulamanın gizli alanında tutulur: Shopify yönetiminde görünmez,
  // sunucu yeniden kurulsa da kaybolmaz.
  const AKTARIM_NS = 'esse_siparis', AKTARIM_KEY = 'aktarimlar';
  async function aktarimlariOku() {
    checkConfig();
    const d = await gql(`{ currentAppInstallation { id metafield(namespace: "${AKTARIM_NS}", key: "${AKTARIM_KEY}") { value } } }`);
    const ai = d.currentAppInstallation;
    let liste = [];
    try { liste = JSON.parse(ai.metafield?.value || '[]'); } catch { liste = []; }
    return { ownerId: ai.id, liste: Array.isArray(liste) ? liste : [] };
  }
  async function aktarimlariYaz(ownerId, liste) {
    const res = await gql(`mutation($m: [MetafieldsSetInput!]!) { metafieldsSet(metafields: $m) { userErrors { field message } } }`,
      { m: [{ ownerId, namespace: AKTARIM_NS, key: AKTARIM_KEY, type: 'json', value: JSON.stringify(liste) }] });
    const errs = res.metafieldsSet.userErrors;
    if (errs.length) throw new Error('aktarım kaydı yazılamadı: ' + errs.map((e) => e.message).join(', '));
  }

  // Uygulamanın gizli alanında başka bir kayıt (ör. etiket ayarları): okunamazsa hata fırlatır
  async function gizliOku(key) {
    checkConfig();
    const d = await gql(`{ currentAppInstallation { id metafield(namespace: "${AKTARIM_NS}", key: "${key}") { value } } }`);
    const ai = d.currentAppInstallation;
    let deger = null;
    try { deger = ai.metafield?.value ? JSON.parse(ai.metafield.value) : null; } catch { deger = null; }
    return { ownerId: ai.id, deger };
  }
  async function gizliYaz(key, deger) {
    const { ownerId } = await gizliOku(key);
    const res = await gql(`mutation($m: [MetafieldsSetInput!]!) { metafieldsSet(metafields: $m) { userErrors { field message } } }`,
      { m: [{ ownerId, namespace: AKTARIM_NS, key, type: 'json', value: JSON.stringify(deger) }] });
    const errs = res.metafieldsSet.userErrors;
    if (errs.length) throw new Error(errs.map((e) => e.message).join(', '));
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

  // Adres kontrolü + düzeltme. Döner: 'duzeltildi' | 'tamam' | 'sorunlu'
  async function fixAddress(o) {
    const a = o.rows[0]._ref.address;
    const r = await adres.check(a, { full: true });
    if (r.degisiklik) {
      // Shopify adres güncellerken soyad ister. Müşteri ad-soyadı tek kutuya yazdıysa
      // aynı tam adı ad + soyad olarak ikiye ayırırız; görünen isim değişmez.
      let firstName = String(a.firstName || '').trim();
      let lastName = String(a.lastName || '').trim();
      if (!lastName) {
        const parts = String(a.name || firstName).trim().split(/\s+/).filter(Boolean);
        lastName = parts.pop() || '';
        firstName = parts.join(' ');
      }
      if (!lastName) throw new Error('siparişte alıcı adı yok, adres Shopify\'da elle düzeltilmeli');
      const input = {
        id: o.rows[0]._ref.gid,
        shippingAddress: {
          firstName, lastName, company: a.company, phone: a.phone,
          address1: r.degisiklik.address1, address2: r.degisiklik.address2 || null,
          city: r.degisiklik.city, provinceCode: a.provinceCode, zip: a.zip, countryCode: a.countryCodeV2 || 'TR',
        },
      };
      const res = await gql(ORDER_UPDATE, { input });
      const errs = res.orderUpdate.userErrors;
      if (errs.length) throw new Error('adres güncellenemedi: ' + errs.map((e) => e.message).join(', '));
      await gql(TAGS_ADD, { id: input.id, tags: [T.adres] });
    }
    return { sonuc: r.durum === 'sorunlu' ? 'sorunlu' : r.degisiklik ? 'duzeltildi' : 'tamam', notlar: r.notlar };
  }

  async function fixAddresses(orders) {
    const adresSonuc = { duzeltildi: [], sorunlu: [], hata: [] };
    for (const o of orders) {
      try {
        const r = await fixAddress(o);
        if (r.sonuc === 'duzeltildi') adresSonuc.duzeltildi.push(o.siparisNo);
        if (r.sonuc === 'sorunlu') adresSonuc.sorunlu.push(`${o.siparisNo}: ${r.notlar.join(' · ')}`);
      } catch (e) {
        console.error(`[adres] ${o.siparisNo}:`, e.message);
        adresSonuc.hata.push(`${o.siparisNo} (${e.message})`);
      }
    }
    return adresSonuc;
  }

  // Excel'e aktarıldı: tek etiket
  const markExported = (orders) => addTag(orders, T.excel);

  return { fetchRows, fetchByIds, markExported, fixAddresses, aktarimlariOku, aktarimlariYaz, gizliOku, gizliYaz, usesStore: false };
})();

// ======================================================================
// TRENDYOL
// ======================================================================
const trendyol = (() => {
  // bekleyen: Created / Picking / Invoiced paketler (Excel'e aktarılanlar sistem kaydından düşülür)
  // Excel'e aktar = Created paketi "Picking" (Trendyol panelinde "İşleme Alındı") yapar.
  const cfg = config.trendyol;
  const { httpJson, pick, toNumber, toDate } = util;

  const BASE = 'https://apigw.trendyol.com/integration/order/sellers';
  const STATUS = ['Created', 'Picking', 'Invoiced'];
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

  async function fetchRows() {
    const h = headers();
    const rows = [];
    for (const status of STATUS) {
      for (const p of await fetchStatus(status, h)) {
        const customer = pick(p, 'shipmentAddress.fullName') || [p.customerFirstName, p.customerLastName].filter(Boolean).join(' ');
        for (const line of p.lines || []) {
          rows.push({
            _key: `Trendyol:${p.id}`,
            _ref: { packageId: p.id, lineId: pick(line, 'id', 'lineId'), qty: toNumber(line.quantity), status },
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
            odemeTipi: 'Trendyol', odeme: 'Pazaryeri',
            durum: DURUM[status] || status,
            etiketler: '',
          });
        }
      }
    }
    return rows;
  }

  // Excel'e aktarırken: henüz işleme alınmamış (Created) paketleri "Picking" yapar
  async function advance(orders) {
    const h = headers();
    const done = [];
    const failed = [];
    for (const o of orders) {
      if (o.rows[0]._ref.status !== 'Created') { done.push(o.key); continue; }
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
  // bekleyen: paketlenecek kalemler + paketlenmiş, kargoya verilmemiş paketler
  //           (Excel'e aktarılanlar sistem kaydından düşülür)
  // Excel'e aktar = paketlenmemiş siparişin tüm kalemlerini tek pakette paketler ("Gönderime Hazır").
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
      odemeTipi: 'Hepsiburada', odeme: 'Pazaryeri',
      etiketler: '',
      ...extra,
    };
  }

  async function fetchRows() {
    const h = headers();
    const rows = [];
    {
      for (const item of await paged(`/orders/merchantid/${cfg.merchantId}`, h, 'Hepsiburada (paketlenecek)')) {
        const no = String(pick(item, 'orderNumber'));
        rows.push(lineToRow(item, item, {
          _key: `Hepsiburada:${no}`,
          _ref: { lineId: pick(item, 'id', 'lineItemId'), qty: toNumber(pick(item, 'quantity')), paketle: true },
          kargoAnahtari: '',
          durum: 'Paketlenecek',
        }));
      }
    }
    {
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

  // Excel'e aktarırken: paketlenmemiş siparişin tüm kalemlerini paketler
  async function advance(orders) {
    const h = headers();
    const done = [];
    const failed = [];
    for (const o of orders) {
      if (!o.rows[0]._ref.paketle) { done.push(o.key); continue; }
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
  // Trendyol / Hepsiburada'da Excel'e aktarılan siparişlerin kaydı (Shopify bunu etiketle tutar).
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

  // Kanal + sipariş numarasıyla bakılır (Hepsiburada'da paketlenince paket anahtarı değişir).
  const noKey = (o) => `${o.kanal}#${o.siparisNo}`;
  const has = (o) => !!load()[noKey(o)];

  function add(orders) {
    const db = load();
    const now = new Date().toISOString();
    for (const o of orders) {
      const { rows, ...rest } = o;
      db[noKey(o)] = { ...rest, aktarimTarihi: now };
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

  const get = (k) => load()[k] || null;

  return { has, add, list, get, noKey };
})();

// ======================================================================
// ADRES KONTROLÜ (mahalle / ilçe / il)
// ======================================================================
const adres = (() => {
  // Shopify adres kontrolü. Mağazanın ödeme formunda:
  //    Adres (address1)                → mahalle, sokak, kapı no
  //    Apartman, daire vb. (address2)  → ilçe
  //    Şehir (city)                    → il
  //  (Alanlar farklı kullanılmışsa il için province/city, ilçe için address2/city/adres metni de denenir.)
  //  • il ve ilçe geçerli mi, ilçe o ile mi ait?
  //  • adres satırında mahalle var mı, varsa o ilçede gerçekten var mı?
  //  • mahalle yoksa sokak + ilçe + il ile haritada aranır, bulunursa adresin başına eklenir
  //  • adres satırının başında/sonunda tekrar yazılmış il ve ilçe adları temizlenir
  // Mahalle listesi: turkey-neighbourhoods paketi (il → ilçe → mahalle).
  const tn = require('turkey-neighbourhoods');
  const cfg = config.adres;

  // Türkçe karakterleri düzleyip küçük harfe çevirir; karakter sayısı korunur (1:1).
  const MAP = { 'ç': 'c', 'Ç': 'c', 'ğ': 'g', 'Ğ': 'g', 'ı': 'i', 'I': 'i', 'İ': 'i', 'ö': 'o', 'Ö': 'o', 'ş': 's', 'Ş': 's', 'ü': 'u', 'Ü': 'u', 'â': 'a', 'Â': 'a', 'î': 'i', 'Î': 'i', 'û': 'u', 'Û': 'u' };
  const fold = (s) => String(s || '').split('').map((c) => MAP[c] ?? c.toLowerCase()).join('');
  const key = (s) => fold(s).replace(/[^a-z0-9]/g, '');
  const SUFFIX = /\b(mahallesi|mahalle|mah|mh)\b\.?/g;
  const baseKey = (s) => key(fold(s).replace(SUFFIX, ' '));

  function lev(a, b) {
    if (Math.abs(a.length - b.length) > 2) return 99;
    const d = Array.from({ length: a.length + 1 }, (_, i) => [i]);
    for (let j = 1; j <= b.length; j++) d[0][j] = j;
    for (let i = 1; i <= a.length; i++) for (let j = 1; j <= b.length; j++) {
      d[i][j] = Math.min(d[i - 1][j] + 1, d[i][j - 1] + 1, d[i - 1][j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
    }
    return d[a.length][b.length];
  }

  // Tam eşleşme; yoksa tek bir yakın eşleşme (yazım hatası) kabul edilir.
  function bestMatch(k, list, getKey) {
    if (!k) return null;
    const exact = list.filter((x) => getKey(x) === k);
    if (exact.length === 1) return exact[0];
    if (k.length < 5) return null;
    const near = list.filter((x) => lev(getKey(x), k) <= (k.length >= 9 ? 2 : 1));
    return near.length === 1 ? near[0] : null;
  }

  const cityCodes = Object.keys(tn.cityNamesByCode);

  function findIl(a) {
    const m = /^TR-?(\d{2})$/i.exec(a.provinceCode || '');
    if (m && tn.cityNamesByCode[m[1]]) return m[1];
    for (const v of [a.province, a.city]) {
      const hit = bestMatch(key(v), cityCodes, (c) => key(tn.cityNamesByCode[c]));
      if (hit) return hit;
    }
    // Şehir alanına il ile birlikte başka şey de yazılmış olabilir ("Adana Yüreğir")
    const words = fold(a.city).replace(/[^a-z0-9\s]/g, ' ').split(/\s+/).filter(Boolean);
    return cityCodes.find((c) => words.includes(key(tn.cityNamesByCode[c]))) || null;
  }

  // Metindeki kelimeler arasında o ilin bir ilçesi geçiyor mu?
  function ilceInText(ilKod, text) {
    const words = fold(text).replace(/[^a-z0-9\s]/g, ' ').split(/\s+/).filter(Boolean);
    return tn.getDistrictsByCityCode(ilKod).find((d) => {
      const parts = fold(d).replace(/[^a-z0-9\s]/g, ' ').split(/\s+/).filter(Boolean);
      return words.some((_, i) => parts.every((p, j) => words[i + j] === p));
    }) || null;
  }

  const findIlce = (ilKod, text) => bestMatch(key(text), tn.getDistrictsByCityCode(ilKod), key);

  const mahalleler = (ilKod, ilce) => tn.getNeighbourhoodsByCityCodeAndDistrict(ilKod, ilce) || [];
  const mahAdi = (m) => m.replace(/\s+Mah\.?$/i, '');

  // Adres metninde mahalle arar: önce "X Mah./Mahallesi/Mh." kalıbı, sonra son ekiz geçen mahalle adı.
  function findMahalle(text, list) {
    const f = fold(text);
    // 1) son ekli kalıp: ekten önceki 1-4 kelimeyi uzundan kısaya dene
    const re = /\b(mahallesi|mahalle|mah|mh)\b\.?/g;
    let m;
    const written = [];
    while ((m = re.exec(f))) {
      const words = f.slice(0, m.index).replace(/[^a-z0-9\s]/g, ' ').trim().split(/\s+/).filter(Boolean);
      for (let n = Math.min(4, words.length); n >= 1; n--) {
        const cand = words.slice(-n).join('');
        const hit = bestMatch(cand, list, baseKey);
        if (hit) return { mahalle: hit, yazili: true };
      }
      written.push(String(text).slice(0, m.index).replace(/[^\p{L}\p{N}\s]/gu, ' ').trim().split(/\s+/).slice(-2).join(' '));
    }
    if (written.length) return { mahalle: null, yazili: true, yanlis: written[0] };
    // 2) ekiz yazılmış mahalle adı (cadde/sokak adı değilse)
    const words = f.replace(/[^a-z0-9\s]/g, ' ').split(/\s+/).filter(Boolean);
    let best = null;
    for (const mh of list) {
      const parts = fold(mahAdi(mh)).replace(/[^a-z0-9\s]/g, ' ').split(/\s+/).filter(Boolean);
      for (let i = 0; i + parts.length <= words.length; i++) {
        if (parts.every((p, j) => words[i + j] === p)) {
          const next = words[i + parts.length] || '';
          if (/^(cad|cd|caddesi|sok|sk|sokak|sokagi|blv|bulv|bulvari|bulvar|yolu|meydani|sitesi|site|apt|ap|apartman|evleri|konutlari|bey|hanim|efendi|pasa)/.test(next)) continue;
          if (parts.join('').length < 5) continue; // "Bey", "Yeni" gibi kısa adlar ekiz yazıldıysa mahalle sayılmaz
          if (!best || parts.length > best.n) best = { mh, n: parts.length };
        }
      }
    }
    return best ? { mahalle: best.mh, yazili: false } : { mahalle: null, yazili: false };
  }

  // Satırın başında/sonunda tekrar yazılmış il, ilçe, Türkiye adlarını temizler.
  function cleanLine(line, names) {
    let s = String(line || '').trim();
    const drop = new Set(names.map(key).concat(['turkiye', 'turkey']));
    const SEP = '[\\s,/\\\\\\-–.()]';
    for (let i = 0; i < 6; i++) {
      const before = s;
      const end = new RegExp(`${SEP}*([^\\s,/\\\\\\-–()]+)${SEP}*$`).exec(s);
      if (end && drop.has(key(end[1])) && s.slice(0, end.index).trim()) s = s.slice(0, end.index).trim();
      const start = new RegExp(`^${SEP}*([^\\s,/\\\\\\-–()]+)${SEP}+`).exec(s);
      if (start && drop.has(key(start[1]))) {
        const rest = s.slice(start[0].length);
        if (rest.trim() && !/^(cad|cd|caddesi|sok|sk|sokak|blv|bulv|bulvar|yolu)/.test(fold(rest))) s = rest.trim();
      }
      if (s === before) break;
    }
    return s.replace(/\s{2,}/g, ' ').replace(/[\s,/-]+$/, '').trim();
  }

  // Sokak bilgisini sadeleştirir (kapı/daire no gibi aramayı bozan kısımlar atılır).
  const streetOnly = (s) => s.replace(/\b(no|kapi|daire|d|kat|k|blok|apt|apartmani|site|sitesi)\b\s*[:.]?\s*[\w/-]*/gi, ' ').replace(/\s{2,}/g, ' ').trim();

  let lastNominatim = 0;
  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

  // ---- Adres satırını sadeleştirme ---------------------------------------------
  // Sadece emin olunan dört parça standart yazılır ve başa şu sırayla alınır:
  //   Xxx Mah. → Xxx Cad. → Xxx Sok. → Xxx Apt.
  // Geri kalan her şey (No, Kat, Daire, site adı, blok, notlar...) yazıldığı gibi,
  // aynı sırayla arkaya eklenir; kısaltılmaz, büyük/küçük harfi değiştirilmez.
  const KURUM = /\b(okul|okulu|ilkokulu|ortaokulu|lise|lisesi|kolej|koleji|anaokulu|kres|kresi|hastane|hastanesi|saglik|polikligi|poliklinik|universite|universitesi|fakulte|fakultesi|kampus|kampusu|yurt|yurdu|cami|camii|belediye|belediyesi|karakol|karakolu|adliye|adliyesi|havalimani|avm|otel|hotel|fabrika|fabrikasi|osb|sanayi|kisla|kislasi|postane|muhtarlik|kaymakamlik|valilik|plaza|is merkezi|ishani)\b/;
  const isKurum = (s) => KURUM.test(fold(s));

  const KW = {
    mah: ['mahallesi', 'mahalle', 'mah', 'mh'],
    cad: ['caddesi', 'cadde', 'cad', 'cd'],
    sok: ['sokagi', 'sokak', 'sok', 'sk'],
    apt: ['apartmani', 'apartman', 'apt', 'ap'],
    // aşağıdakiler sadece sınır olarak tanınır, dokunulmaz
    diger: ['bulvari', 'bulvar', 'bulv', 'blv', 'yolu', 'sitesi', 'site', 'evleri', 'konutlari', 'rezidans', 'residence',
      'no', 'numara', 'kat', 'daire', 'blok', 'blk'],
  };
  const kind = (t) => Object.keys(KW).find((k) => KW[k].includes(t)) || null;
  const STREET = ['cad', 'sok'];
  const isNum = (f) => /^[0-9]/.test(f);

  const up = (c) => (c === 'i' ? 'İ' : c === 'ı' ? 'I' : c.toLocaleUpperCase('tr-TR'));
  const title = (w) => {
    if (/^[0-9]/.test(w)) return w;
    const lower = w.toLocaleLowerCase('tr-TR');
    return up(lower[0]) + lower.slice(1);
  };

  // Kelimelere böler; her kelimenin orijinal metindeki yeri ve önündeki boşluk/ayraç saklanır.
  function tokenize(line) {
    const s = String(line || '');
    const toks = [];
    const re = /[^\s,]+/g;
    let m;
    while ((m = re.exec(s))) {
      // "mah.159.sok.ev" gibi bitişik yazımları noktadan sonra böl
      const parts = m[0].split(/(?<=[.:])(?=[^\s.:])/);
      let pos = m.index;
      parts.forEach((p) => {
        const clean = p.replace(/^[.:;()]+|[.:;()]+$/g, '');
        toks.push({ orig: p, clean, f: fold(clean), start: pos, end: pos + p.length });
        pos += p.length;
      });
    }
    toks.forEach((t, i) => { t.gap = s.slice(i ? toks[i - 1].end : 0, t.start); });
    return toks.filter((t) => t.clean || t.orig.trim());
  }

  function streetKeys(text) {
    const toks = tokenize(text);
    const out = [];
    toks.forEach((t, i) => {
      if (!STREET.includes(kind(t.f)) && !['bulvari', 'bulvar', 'bulv', 'blv', 'yolu'].includes(t.f)) return;
      const name = [];
      for (let j = i - 1; j >= 0 && name.length < 3; j--) {
        if (kind(toks[j].f)) break;
        name.unshift(toks[j].clean);
        if (isNum(toks[j].f) || /,/.test(toks[j].gap)) break;
      }
      if (name.length) out.push(key(name.join('')));
    });
    return out;
  }

  // Satırı yeniden düzenler. Emin olunamazsa null döner (satıra dokunulmaz).
  //  mahalleResmi: resmi mahalle adı (haritadan bulunduysa ya da yazılanın doğrulanmış hali)
  function restructure(line, mahalleResmi, list) {
    const toks = tokenize(line);
    const used = new Set();
    const seg = { mah: null, cad: [], sok: [], apt: null };

    for (let i = 0; i < toks.length; i++) {
      const k = kind(toks[i].f);
      if (!['mah', 'cad', 'sok', 'apt'].includes(k)) continue;
      // anahtar kelimeden geriye doğru isim kelimelerini topla
      const name = [];
      for (let j = i - 1; j >= 0 && name.length < 4; j--) {
        if (used.has(j) || kind(toks[j].f)) break;
        if (name.length && isNum(toks[j].f)) break;      // önceki parçanın numarası
        name.unshift(j);
        if (isNum(toks[j].f)) break;                       // "159 Sok." → isim sadece numara
        if (/,/.test(toks[j].gap)) break;                  // virgül sınırı
      }
      if (!name.length) return null;
      if (k === 'mah') {
        if (seg.mah) return null;
        // resmi listede eşleşen en kısa sondan başlayan kelime grubu mahalle adıdır
        let hitAt = -1;
        for (let n = 1; n <= name.length; n++) {
          const cand = name.slice(-n).map((j) => toks[j].f).join('');
          if (list.some((m) => baseKey(m) === key(cand)) || (n === name.length && bestMatch(key(cand), list, baseKey))) { hitAt = n; break; }
        }
        if (hitAt < 0) return null;
        const idx = name.slice(-hitAt);
        idx.forEach((j) => used.add(j));
        used.add(i);
        seg.mah = true;
        continue;
      }
      if (name.length > 3) return null;                    // isim sınırından emin değiliz
      name.forEach((j) => used.add(j));
      used.add(i);
      const text = name.map((j) => title(toks[j].clean) + (isNum(toks[j].f) && /\.$/.test(toks[j].orig) ? '.' : '')).join(' ');
      if (k === 'apt') { if (seg.apt) return null; seg.apt = text; } else seg[k].push(text);
    }

    // Mahalle eksiz yazılmışsa ("Cihangir 1234 sk") o kelimeler de mahalle parçası sayılır
    if (mahalleResmi && !seg.mah) {
      const want = fold(mahAdi(mahalleResmi)).replace(/[^a-z0-9\s]/g, ' ').split(/\s+/).filter(Boolean);
      for (let i = 0; i + want.length <= toks.length; i++) {
        if (want.every((w, j) => !used.has(i + j) && toks[i + j].f === w)) {
          want.forEach((_, j) => used.add(i + j));
          break;
        }
      }
    }

    const parts = [];
    if (mahalleResmi) parts.push(`${mahAdi(mahalleResmi)} Mah.`);
    else if (seg.mah) return null;
    seg.cad.forEach((t) => parts.push(`${t} Cad.`));
    seg.sok.forEach((t) => parts.push(`${t} Sok.`));
    if (seg.apt) parts.push(`${seg.apt} Apt.`);

    // geri kalanlar: yazıldığı gibi, aynı sırayla
    let rest = '';
    let prevUsed = true;
    toks.forEach((t, i) => {
      if (used.has(i)) { prevUsed = true; return; }
      rest += (rest && !prevUsed ? t.gap : rest ? ' ' : '') + t.orig;
      prevUsed = false;
    });
    rest = rest.replace(/^[\s,/-]+/, '').trim();
    return [...parts, rest].filter(Boolean).join(' ');
  }

  // Sokak/cadde adının karşılaştırma anahtarı: "159. Sokak" → "159", "Petunya Sk." → "petunya"
  const STREET_WORDS = /\b(sokagi|sokak|sok|sk|caddesi|cadde|cad|cd|bulvari|bulvar|bulv|blv|yolu)\b/g;
  const streetKey = (s) => key(fold(s).replace(STREET_WORDS, ' '));

  async function googleSearch(query, ilce, ilAd) {
    const url = 'https://maps.googleapis.com/maps/api/geocode/json?language=tr&region=tr' +
      `&address=${encodeURIComponent(`${query}, ${ilce}, ${ilAd}, Türkiye`)}` +
      `&components=${encodeURIComponent(`country:TR|administrative_area:${ilAd}`)}&key=${cfg.googleKey}`;
    const d = await util.httpJson(url, {}, 'Google Maps');
    if (d.status && d.status !== 'OK' && d.status !== 'ZERO_RESULTS') {
      throw new Error(`${d.status}${d.error_message ? ': ' + d.error_message : ''}`);
    }
    return (d.results || []).slice(0, 3).map((r) => {
      const comps = r.address_components || [];
      const road = comps.find((c) => (c.types || []).includes('route'));
      return { road: road ? road.long_name : '', comps: comps.map((c) => c.long_name) };
    });
  }

  async function osmSearch(query, ilce, ilAd) {
    const wait = 1100 - (Date.now() - lastNominatim);
    if (wait > 0) await sleep(wait); // OpenStreetMap kuralı: saniyede en fazla 1 istek
    lastNominatim = Date.now();
    const url = 'https://nominatim.openstreetmap.org/search?format=jsonv2&addressdetails=1&countrycodes=tr&limit=3' +
      `&q=${encodeURIComponent(`${query}, ${ilce}, ${ilAd}`)}` + (cfg.email ? `&email=${encodeURIComponent(cfg.email)}` : '');
    const d = await util.httpJson(url, { headers: { 'User-Agent': 'esse-jeffe-otomasyon/1.0', 'Accept-Language': 'tr' } }, 'OpenStreetMap');
    return (d || []).map((r) => ({ road: (r.address || {}).road || '', comps: Object.values(r.address || {}).map(String) }));
  }

  // Haritada arar. Mahalle ancak şu üçü birden tutarsa kabul edilir:
  //  sonuç bizim sokak/caddemize düşmüş, bizim ilçemizde, ve mahalle o ilçenin resmi listesinde var.
  async function geocode(sKeys, query, ilce, ilAd, list) {
    let results = null; // [{ road, comps }]
    if (cfg.googleKey) {
      try {
        results = await googleSearch(query, ilce, ilAd);
      } catch (e) {
        // Google reddederse (faturalandırma, kota vb.) ücretsiz haritaya geçilir
        console.error('[adres] harita araması (Google, OpenStreetMap\'e geçiliyor):', e.message);
      }
    }
    if (!results) {
      try {
        results = await osmSearch(query, ilce, ilAd);
      } catch (e) {
        console.error('[adres] harita araması (OpenStreetMap):', e.message);
        return null;
      }
    }
    const found = new Set();
    for (const r of results) {
      if (!r.road || !sKeys.includes(streetKey(r.road))) continue; // başka sokağa düşen sonuç
      if (!r.comps.some((c) => key(c) === key(ilce))) continue;    // başka ilçeye düşen sonuç
      for (const c of r.comps) {
        const hit = list.find((m) => baseKey(m) === baseKey(c));
        if (hit) found.add(hit);
      }
    }
    return found.size === 1 ? [...found][0] : null;
  }

  // a: Shopify shippingAddress. full=false: sadece kontrol (haritaya gitmez, değiştirmez).
  async function check(a, { full = false } = {}) {
    if (!a) return { durum: 'sorunlu', notlar: ['Teslimat adresi yok'] };
    const notlar = [];
    const ilKod = findIl(a);
    if (!ilKod) return { durum: 'sorunlu', notlar: [`İl tanınamadı (Şehir: ${a.city || 'boş'})`] };
    const ilAd = tn.cityNamesByCode[ilKod];

    // İlçe: önce "Apartman, daire vb." alanı, sonra Şehir alanı, en son adres metni
    let ilce = null;
    let ilceAlani = null;
    if ((ilce = findIlce(ilKod, a.address2) || ilceInText(ilKod, a.address2))) ilceAlani = 'address2';
    else if ((ilce = findIlce(ilKod, a.city))) ilceAlani = 'city';
    else if (!String(a.address2 || '').trim() && (ilce = ilceInText(ilKod, a.city))) ilceAlani = 'tasinacak';
    else if ((ilce = ilceInText(ilKod, a.address1))) {
      // İlçe sadece adres satırında yazıyor: "Apartman, daire vb." boşsa oraya taşınır
      ilceAlani = String(a.address2 || '').trim() ? 'address1' : 'tasinacak';
      if (ilceAlani === 'address1') notlar.push(`İlçe "Apartman, daire vb." alanında yok (adreste ${ilce} geçiyor)`);
    } else {
      notlar.push(`İlçe bulunamadı ("Apartman, daire vb." alanı: ${a.address2 || 'boş'}) – ${ilAd} ilinin ilçelerinden biri olmalı`);
      return { durum: 'sorunlu', notlar, ilAd };
    }
    // Mahalle/sokak metni: ilçe alanı olarak kullanılan satır hariç
    const metin = [a.address1, ilceAlani === 'address2' ? '' : a.address2].filter(Boolean).join(' ');
    const list = mahalleler(ilKod, ilce);
    const bulunan = findMahalle(metin, list);

    // Okul, hastane, üniversite vb. kurum adreslerine dokunulmaz
    if (isKurum(metin)) {
      return {
        durum: bulunan.mahalle ? 'tamam' : 'kurum',
        notlar: bulunan.mahalle ? [] : ['Kurum adresi (okul, hastane vb.) – otomatik düzenleme yapılmadı'],
        ilAd, ilce, mahalle: bulunan.mahalle, eklendi: false, degisiklik: null,
      };
    }

    // il/ilçe tekrarları temizlenmiş hali (sadece mahalle-ilçe-il uyumluysa kullanılır)
    const orijinal = String(a.address1 || '').trim();
    let address1 = cleanLine(a.address1, [ilAd, ilce]);
    // İlçe alanı olarak kullanılan satıra dokunulmaz; boşsa ve ilçe adreste yazıyorsa oraya yazılır
    const address2 = ilceAlani === 'address2' ? String(a.address2 || '').trim()
      : ilceAlani === 'tasinacak' ? ilce
      : cleanLine(a.address2, [ilAd, ilce]);
    let mahalle = bulunan.mahalle;
    let eklendi = false;

    if (!mahalle && bulunan.yanlis) {
      notlar.push(`"${bulunan.yanlis}" mahallesi ${ilce} / ${ilAd} içinde bulunamadı`);
    } else if (!mahalle && full) {
      const sKeys = streetKeys(address1);
      if (sKeys.length) {
        mahalle = await geocode(sKeys, streetOnly(address1), ilce, ilAd, list);
        eklendi = !!mahalle;
      }
    }
    // Mahalle–ilçe–il uyumlu değilse adres satırından hiçbir şey silinmez (il/ilçe tekrarı dahil)
    const tutarli = !!mahalle && notlar.length === 0;
    if (!tutarli) address1 = orijinal;
    if (!mahalle && !bulunan.yanlis) notlar.push('Adreste mahalle yok' + (full ? ', haritada da kesin olarak bulunamadı' : ''));

    // Düzenleme: Mah./Cad./Sok./Apt. emin olunan yerlerde standart yazılır ve başa alınır,
    // geri kalan her şey yazıldığı gibi kalır. Emin olunamazsa satıra dokunulmaz.
    let bicimlendi = false;
    if (!bulunan.yanlis) {
      // Mahalle adı satıra sadece yazılı halini düzeltmek ya da haritadan bulunanı eklemek için yazılır
      const yaz = mahalle;
      const yeni = restructure(address1, yaz, list);
      if (yeni) { address1 = yeni; bicimlendi = true; }
      else if (eklendi) address1 = `${mahAdi(mahalle)} Mah. ${address1}`.trim();
    }

    // Şehir alanı ilçe olarak kullanılmıyorsa sadece il adı kalmalı (yine sadece uyumluysa)
    let city = String(a.city || '').trim();
    if (tutarli && ilceAlani !== 'city' && key(city) !== key(ilAd)) city = ilAd;

    const degisti = address1 !== String(a.address1 || '').trim() || address2 !== String(a.address2 || '').trim() || city !== String(a.city || '').trim();
    return {
      durum: notlar.length ? 'sorunlu' : 'tamam',
      notlar, ilAd, ilce, mahalle, eklendi, bicimlendi,
      degisiklik: degisti ? { address1, address2, city } : null,
    };
  }

  return { check, fold, key };
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

  // SKU sıralama: sipariş içindeki SKU'lar alfabetik (sayılar doğal sırada), siparişler de
  // ilk SKU'larına göre sıralanır — aynı ürünler Excel'de alt alta gelir.
  const cmp = new Intl.Collator('tr', { numeric: true, sensitivity: 'base' }).compare;
  const skuSirala = (o) => ({ ...o, skus: [...o.skus].sort((a, b) => cmp(a.sku, b.sku)) });
  function sortBySku(orders) {
    return orders.map(skuSirala)
      .sort((a, b) => cmp(a.skus[0]?.sku || '', b.skus[0]?.sku || '') || cmp(a.siparisNo, b.siparisNo));
  }
  // Paneldeki sıralama: siparişler gönderilen sırada kalır (yalnız sipariş içindeki SKU'lar sıralanır)
  const ekranSirasi = (orders) => orders.map(skuSirala);

  const skuText = (o) => o.skus.map((x) => (x.adet > 1 ? `${x.sku} ×${x.adet}` : x.sku)).join('\n');
  const dateText = (d) => d ? new Date(d).toLocaleString('tr-TR', {
    timeZone: 'Europe/Istanbul', day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit',
  }) : '';

  // Excel ve Google Sheet'te görünen kanal adları
  const fold = (s) => adres.fold(s);
  const KANAL_ADI = { Shopify: 'website', Trendyol: 'trendyol', Hepsiburada: 'hepsiburada' };
  const kanalAdi = (k) => KANAL_ADI[k] || k;

  // Kargo firması adları: "Marketplace" eki atılır, HepsiJET'in tüm yazımları tek ada indirilir
  function kargoAdi(name) {
    const n = String(name || '').trim();
    if (fold(n).replace(/[^a-z]/g, '').includes('hepsijet')) return 'HepsiJet';
    return n.replace(/\s*marketplace\s*/gi, ' ').replace(/\s{2,}/g, ' ').trim();
  }

  // Sipariş -> düz değerler (Excel ve Sheet için ortak)
  function toRecord(o) {
    return {
      kanal: kanalAdi(o.kanal), siparisNo: o.siparisNo, tarih: dateText(o.tarih), musteri: o.musteri,
      telefon: o.telefon, adres: o.adres || '', kargoFirmasi: kargoAdi(o.kargoFirmasi), kargoAnahtari: o.kargoAnahtari,
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
    for (const kanal of channels) addSheet(wb, kanalAdi(kanal), orders.filter((o) => o.kanal === kanal));
    if (warnings.length) {
      const ws = wb.addWorksheet('Uyarılar');
      ws.columns = [{ header: 'Uyarı', key: 'w', width: 120 }];
      warnings.forEach((w) => ws.addRow({ w }));
      ws.eachRow((row, i) => (row.font = { ...FONT, bold: i === 1, color: i === 1 ? undefined : { argb: 'FFB00020' } }));
    }
    return wb.xlsx.writeBuffer();
  }

  return { buildWorkbook, COLUMNS, toRecord, sortBySku, ekranSirasi };
})();

// ======================================================================
// GOOGLE SHEETS (tek sayfa; her aktarımda sayfa yenilenir)
// ======================================================================
const sheets = (() => {
  // Tek sayfa: her aktarımda başlık dışındaki tüm kayıtlar silinir, seçilen siparişler 2. satırdan itibaren yazılır.
  const cfg = config.google;
  const API = 'https://sheets.googleapis.com/v4/spreadsheets';
  const TAB = config.google.tab;
  let cached = null; // { token, expiresAt }
  let tabId = null;

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

  async function api(path, method = 'GET', body) {
    const token = await getToken();
    return util.httpJson(`${API}/${cfg.sheetId}${path}`, {
      method, headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
      body: body ? JSON.stringify(body) : undefined,
    }, 'Google Sheets');
  }

  // Sheet kolonları: Excel kolonlarının önüne "Aktarım" (aktarım tarihi-saati) eklenir.
  const columns = () => [{ header: 'Aktarım', key: 'aktarim', width: 16 }, ...excel.COLUMNS];
  const range = (a1) => encodeURIComponent(`'${TAB.replace(/'/g, "''")}'!${a1}`);

  const fmt = (gid, start, end, header) => ({
    repeatCell: {
      range: { sheetId: gid, startRowIndex: start, endRowIndex: end, startColumnIndex: 0, endColumnIndex: columns().length },
      cell: { userEnteredFormat: header
        ? { textFormat: { bold: true, foregroundColor: { red: 1, green: 1, blue: 1 } }, backgroundColor: { red: 0.231, green: 0.165, blue: 0.184 }, wrapStrategy: 'CLIP', verticalAlignment: 'MIDDLE' }
        : { textFormat: { bold: false, foregroundColor: { red: 0, green: 0, blue: 0 } }, backgroundColor: { red: 1, green: 1, blue: 1 }, wrapStrategy: 'WRAP', verticalAlignment: 'TOP' } },
      fields: 'userEnteredFormat(textFormat,backgroundColor,wrapStrategy,verticalAlignment)',
    },
  });

  // Sayfa yoksa oluşturur, başlık satırını yazar. Sayfanın sheetId'sini döner.
  async function ensureTab() {
    if (tabId != null) return tabId;
    const meta = await api('?fields=sheets.properties(sheetId,title)');
    const found = (meta.sheets || []).find((s) => s.properties.title === TAB);
    if (found) {
      tabId = found.properties.sheetId;
      const head = await api(`/values/${range('A1:A1')}`);
      if (head.values && head.values.length) return tabId;
    } else {
      const added = await api(':batchUpdate', 'POST', {
        requests: [{ addSheet: { properties: { title: TAB, index: 0, gridProperties: { frozenRowCount: 1 } } } }],
      });
      tabId = added.replies[0].addSheet.properties.sheetId;
    }
    const cols = columns();
    await api(`/values/${range('A1')}?valueInputOption=RAW`, 'PUT', { values: [cols.map((c) => c.header)] });
    await api(':batchUpdate', 'POST', { requests: [
      fmt(tabId, 0, 1, true),
      { updateSheetProperties: { properties: { sheetId: tabId, gridProperties: { frozenRowCount: 1 } }, fields: 'gridProperties.frozenRowCount' } },
      ...cols.map((c, i) => ({ updateDimensionProperties: {
        range: { sheetId: tabId, dimension: 'COLUMNS', startIndex: i, endIndex: i + 1 },
        properties: { pixelSize: Math.round(c.width * 8) }, fields: 'pixelSize' } })),
    ] });
    return tabId;
  }

  async function writeOrders(orders, etiket) {
    if (!cfg.sheetId || !cfg.serviceAccount) return { skipped: true };
    let gid;
    try {
      gid = await ensureTab();
    } catch (e) {
      tabId = null; // sayfa silinmiş olabilir; bir sonraki denemede yeniden bakılır
      throw e;
    }
    const cols = columns();
    const stamp = new Date().toLocaleString('tr-TR', {
      timeZone: 'Europe/Istanbul', day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit',
    });
    const values = orders.map((o) => {
      const r = { aktarim: etiket || stamp, ...excel.toRecord(o) };
      return cols.map((c) => r[c.key] ?? '');
    });
    const n = values.length;

    // Başlık hariç tüm eski kayıtları sil, seçilenleri 2. satırdan itibaren yaz.
    try {
      await api(`/values/${range('A2:ZZ')}:clear`, 'POST', {});
    } catch (e) {
      tabId = null; // sayfa silinmiş olabilir; bir sonraki denemede yeniden bakılır
      throw e;
    }
    // RAW: uzun ID'ler ve telefonlar sayıya dönüşmeden metin olarak kalır.
    await api(`/values/${range('A2')}?valueInputOption=RAW`, 'PUT', { values });
    await api(':batchUpdate', 'POST', { requests: [fmt(gid, 1, 1 + n, false)] });

    return { url: `https://docs.google.com/spreadsheets/d/${cfg.sheetId}/edit#gid=${gid}`, title: TAB };
  }

  return { writeOrders, configured: () => !!(cfg.sheetId && cfg.serviceAccount) };
})();

// ======================================================================
// SİPARİŞ AKIŞI (tek sayfa: siparişleri çek → Excel'e aktar)
// ======================================================================
const flow = (() => {
  //  Siparişleri çek : henüz Excel'e aktarılmamış siparişleri listeler (hiçbir şeyi değiştirmez)
  //  Excel'e aktar   : seçilenleri Excel'e (ayarlıysa Google Sheet'e) yazar; e-posta gönderilmez (kullanıcı isteği);
  //                    Shopify → "excel" etiketi; Trendyol/Hepsiburada → işleme alınır + kayıt
  //  Aktarılanlar    : tarih aralığıyla görüntüleme + Excel'i yeniden indirme (durum değiştirmez)
  const { buildWorkbook } = excel;
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
          odemeTipi: r.odemeTipi, odeme: r.odeme || r.odemeTipi || '', durum: r.durum, etiketler: r.etiketler || '',
          adresDurum: r.adresDurum || '', adresNot: r.adresNot || '', adresDuzeltildi: !!r.adresDuzeltildi,
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

  // ---- Aktarımlar: her "Excel'e aktar" numaralı bir aktarımdır (no, zaman, siparişler) ----
  // Liste Shopify açıksa uygulamanın gizli alanında, değilse DATA_DIR/aktarimlar.json'da tutulur.
  const fs = require('fs');
  const path = require('path');
  const aktarimDosyasi = path.join(config.dataDir, 'aktarimlar.json');
  async function aktarimlar() {
    if (config.sources.includes('shopify')) return shopify.aktarimlariOku();
    try { return { liste: JSON.parse(fs.readFileSync(aktarimDosyasi, 'utf8')) }; } catch { return { liste: [] }; }
  }
  async function aktarimlariKaydet(kayit, liste) {
    // Kayıt boyutu sınırlı: en eski aktarımlar düşülür (≈ son birkaç yüz aktarım kalır)
    while (liste.length > 1 && JSON.stringify(liste).length > 90000) liste.shift();
    if (config.sources.includes('shopify')) return shopify.aktarimlariYaz(kayit.ownerId, liste);
    fs.mkdirSync(config.dataDir, { recursive: true });
    fs.writeFileSync(aktarimDosyasi, JSON.stringify(liste));
  }
  const trZaman = (iso) => new Date(iso).toLocaleString('tr-TR', {
    timeZone: 'Europe/Istanbul', day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' });
  const aktarimEtiketi = (a) => `#${a.no} · ${trZaman(a.zaman)}`;
  const trGun = (iso) => new Date(iso).toLocaleDateString('en-CA', { timeZone: 'Europe/Istanbul' });

  // Seçilen günün aktarımları (yeniden eskiye)
  async function aktarimListesi({ tarih } = {}) {
    const { liste } = await aktarimlar();
    return liste.filter((a) => !tarih || trGun(a.zaman) === tarih).reverse()
      .map((a) => ({ no: a.no, zaman: a.zaman, adet: (a.siparisler || []).length }));
  }

  // stage: 'bekleyen' (aktarılmamış) | 'aktarim' (bir aktarımın siparişleri, no ile)
  async function list({ stage = 'bekleyen', no } = {}) {
    const sources = active();
    const warnings = [];
    let orders = [];
    let aktarim = null;
    if (stage === 'aktarim') {
      aktarim = (await aktarimlar()).liste.find((a) => a.no === Number(no));
      if (!aktarim) throw new Error(`#${no} numaralı aktarım bulunamadı.`);
    }

    const results = await Promise.allSettled(sources.map(async ([name, src]) => {
      if (aktarim) {
        const benim = (aktarim.siparisler || []).filter((x) => (x.kanal || 'Shopify') === name);
        if (!benim.length) return [];
        if (src.usesStore) return benim.map((x) => store.get(x.k)).filter(Boolean).map((o) => ({ ...o, rows: [] }));
        return group(await src.fetchByIds(benim.map((x) => x.id)));
      }
      const grouped = group(await src.fetchRows({ stage }));
      return src.usesStore ? grouped.filter((o) => !store.has(o)) : grouped;
    }));
    results.forEach((r, i) => {
      if (r.status === 'fulfilled') orders.push(...r.value);
      else {
        warnings.push(`${sources[i][0]}: ${r.reason.message}`);
        console.error(`[${sources[i][0]}]`, r.reason);
      }
    });

    orders = aktarim ? excel.sortBySku(orders) : sortDesc(orders);
    const sessionId = remember({ stage, aktarim, orders: new Map(orders.map((o) => [o.key, o])) });
    return { sessionId, stage, aktarim: aktarim && { no: aktarim.no, zaman: aktarim.zaman, etiket: aktarimEtiketi(aktarim) },
      orders: orders.map(publicView), warnings };
  }

  function take(sessionId, keys, stage) {
    const s = sessions.get(sessionId);
    if (!s) throw new Error('Liste süresi doldu, "Siparişleri çek"e yeniden basın.');
    if (stage && s.stage !== stage) throw new Error('Bu işlem bu listede yapılamaz.');
    const selected = (keys || []).map((k) => s.orders.get(k)).filter(Boolean);
    if (!selected.length) throw new Error('Seçilen siparişler zaten aktarıldı; "Siparişleri çek"e yeniden basın.');
    return { s, selected };
  }

  const byChannel = (orders) => orders.reduce((m, o) => ((m[o.kanal] ||= []).push(o), m), {});

  function fileName(prefix) {
    const p = (n) => String(n).padStart(2, '0');
    const tr = new Date(new Date().toLocaleString('en-US', { timeZone: 'Europe/Istanbul' }));
    return `${prefix}_${tr.getFullYear()}-${p(tr.getMonth() + 1)}-${p(tr.getDate())}_${p(tr.getHours())}${p(tr.getMinutes())}.xlsx`;
  }


  // Excel'e aktar: Excel dosyası her zaman oluşur; Google Sheet ayarlıysa o da yazılır.
  // sirali: siparişler paneldeki sırayla (keys sırası) yazılır; değilse eskisi gibi SKU sırasıyla.
  async function exportExcel({ sessionId, keys, sirali = false }) {
    const { s, selected } = take(sessionId, keys, 'bekleyen');
    const channels = active().map(([n]) => n);
    const warnings = [];
    const exported = [];
    // Aktarım numarası: listeyi okuyamazsak hiçbir siparişe dokunmadan dururuz
    const kayit = await aktarimlar();
    const aktarim = { no: kayit.liste.reduce((m, a) => Math.max(m, Number(a.no) || 0), 0) + 1, zaman: new Date().toISOString(), siparisler: [] };

    for (const [kanal, orders] of Object.entries(byChannel(selected))) {
      const src = byName[kanal];
      if (src.usesStore) {
        // Trendyol: Created → Picking · Hepsiburada: paketlenmemişse paketle. Başarılı olanlar kaydedilir.
        const res = await src.advance(orders);
        const ok = orders.filter((o) => res.done.includes(o.key));
        if (ok.length) store.add(ok);
        exported.push(...ok);
        if (res.failed.length) warnings.push(`${kanal} işleme alınamayan siparişler (Excel'e yazılmadı, listede kalır): ${res.failed.join('; ')}`);
      } else {
        const res = await src.markExported(orders);
        exported.push(...orders.filter((o) => res.done.includes(o.key)));
        if (res.failed.length) warnings.push(`${kanal} etiketlenemeyen siparişler (Excel'e yazılmadı, listede kalır): ${res.failed.join('; ')}`);
      }
    }
    exported.forEach((o) => s.orders.delete(o.key));
    if (!exported.length) throw new Error('Hiçbir sipariş aktarılamadı. ' + warnings.join(' '));

    aktarim.siparisler = exported.map((o) => (byName[o.kanal].usesStore
      ? { kanal: o.kanal, k: store.noKey(o), ad: o.siparisNo } : { id: o.rows[0]._ref.gid, ad: o.siparisNo }));
    try { await aktarimlariKaydet(kayit, [...kayit.liste, aktarim]); }
    catch (e) { console.error(e); warnings.push(`Aktarım #${aktarim.no} geçmişe kaydedilemedi (${e.message}); Excel'i şimdi indirin.`); }
    const etiket = aktarimEtiketi(aktarim);

    let sheet = null;
    const sira = new Map((keys || []).map((k, i) => [k, i]));
    const ordered = sirali ? excel.ekranSirasi([...exported].sort((a, b) => sira.get(a.key) - sira.get(b.key))) : excel.sortBySku(exported);
    if (sheets.configured()) {
      try { sheet = await sheets.writeOrders(ordered, etiket); }
      catch (e) { console.error(e); warnings.push(`Google Sheet'e yazılamadı: ${e.message}`); }
    }

    const buffer = await buildWorkbook(ordered, warnings, channels);
    const filename = fileName(`aktarim-${aktarim.no}`);
    return { orderCount: exported.length, aktarim: { no: aktarim.no, zaman: aktarim.zaman, etiket }, buffer, filename, sheet, warnings,
      etiketler: etiketVerisi(ordered, etiket) };
  }

  // Aktarılanlar listesinden seçilenleri yeniden Excel olarak almak (durum değiştirmez)
  async function download({ sessionId, keys, sirali = false }) {
    const { s, selected } = take(sessionId, keys);
    const buffer = await buildWorkbook(sirali ? excel.ekranSirasi(selected) : excel.sortBySku(selected), [], active().map(([n]) => n));
    return { orderCount: selected.length, buffer, filename: fileName(s.aktarim ? `aktarim-${s.aktarim.no}` : 'liste') };
  }

  // Etiket değişkenleri: şablondaki {{AD}} yerlerine yazılır; değerler Excel/Sheet satırıyla aynı (excel.toRecord)
  function etiketVerisi(orders, etiket) {
    return orders.map((o) => {
      const r = excel.toRecord(o);
      const t = (v) => (v == null ? '' : String(v));
      return {
        AKTARIM: t(etiket), KANAL: t(r.kanal), SIPARIS_NO: t(r.siparisNo), TARIH: t(r.tarih), MUSTERI: t(r.musteri),
        TELEFON: t(r.telefon), ADRES: t(r.adres), KARGO_FIRMASI: t(r.kargoFirmasi), KARGO_ANAHTARI: t(r.kargoAnahtari), SKU: t(r.sku),
        ADET: t(r.adet), TUTAR: t(r.tutar), ODEME: t(r.odemeTipi), DURUM: t(r.durum), ETIKETLER: t(r.etiketler),
      };
    });
  }

  // ---- Etiket ayarları: ölçü (mm), sipariş etiketi şablonu (HTML), serbest etiket (HTML) ----
  // Uygulamanın gizli alanında (esse_siparis.etiket) tutulur; okunamazsa/yazılamazsa DATA_DIR/etiket-ayar.json.
  // Şablonlarda "cqw" birimi etiket genişliğinin yüzdesidir: ölçü değişince düzen orantılı kalır.
  const ETIKET_VARSAYILAN = {
    genislik: 100, yukseklik: 100,
    // bicim: PPLA resim yükleme biçimi (B/b = BMP, P/p = PCX; "Yazıcı testi" ile seçilir) · koyuluk 0 = yazıcının kendi ayarı
    yazici: { yol: 'dogrudan', bicim: 'B', cevir: false, ters: false, koyuluk: 0 },
    sablon: `<style>
.e{height:100%;box-sizing:border-box;padding:5cqw 4cqw;display:flex;flex-direction:column;align-items:center;justify-content:center;
  text-align:center;font-family:"Open Sans","Segoe UI",Arial,sans-serif;color:#000}
.site{font-weight:700;font-size:4.4cqw}
.sku{font-weight:700;font-size:5.8cqw;line-height:1.15;margin-top:5cqw}
.barkod{width:76cqw;height:26cqw;margin-top:6cqw}
.no{font-family:Menlo,"DejaVu Sans Mono",monospace;font-size:4.8cqw;letter-spacing:.9cqw;line-height:1.1}
.kargo{font-weight:700;font-size:6.2cqw;margin-top:2.5cqw;line-height:1.1}
.kanal{font-weight:700;font-size:4.4cqw;line-height:1.1}
.alt{font-size:3.8cqw;margin-top:3cqw;line-height:1.35}
</style>
<div class="e">
  <div class="site">www.essejeffe.com</div>
  <div class="sku">{{SKU}}</div>
  <div class="barkod">{{BARKOD}}</div>
  <div class="no">{{KARGO_ANAHTARI}}</div>
  <div class="kargo">{{KARGO_FIRMASI}}</div>
  <div class="kanal">{{KANAL}}</div>
  <div class="alt">{{MUSTERI}}<br>{{SIPARIS_NO}}</div>
</div>`,
    serbest: `<style>
.s{height:100%;box-sizing:border-box;padding:6cqw;display:flex;flex-direction:column;align-items:center;justify-content:center;
  text-align:center;font-family:"Open Sans","Segoe UI",Arial,sans-serif;color:#000;gap:3cqw}
</style>
<div class="s">
  <div style="font-size:9cqw;font-weight:700">Yazınızı buraya yazın</div>
  <div style="font-size:5cqw">ikinci satır</div>
</div>`,
  };
  const etiketDosyasi = path.join(config.dataDir, 'etiket-ayar.json');
  const ayarTemizle = (a) => {
    const mm = (v, d) => { const n = Number(v); return Number.isFinite(n) && n >= 10 && n <= 300 ? Math.round(n * 10) / 10 : d; };
    const html = (v, d) => (typeof v === 'string' && v.length <= 50000 ? v : d);
    const D = ETIKET_VARSAYILAN;
    const y = a?.yazici || {}, DY = D.yazici;
    const koyuluk = Number(y.koyuluk);
    return { genislik: mm(a?.genislik, D.genislik), yukseklik: mm(a?.yukseklik, D.yukseklik),
      sablon: html(a?.sablon, D.sablon), serbest: html(a?.serbest, D.serbest),
      // Yazıcıya gönderme: dogrudan = PPLA (Mac yardımcısı), pencere = tarayıcının yazdırma penceresi
      yazici: { yol: ['dogrudan', 'pencere'].includes(y.yol) ? y.yol : DY.yol,
        bicim: ['B', 'b', 'P', 'p'].includes(y.bicim) ? y.bicim : DY.bicim,
        cevir: typeof y.cevir === 'boolean' ? y.cevir : DY.cevir, ters: typeof y.ters === 'boolean' ? y.ters : DY.ters,
        koyuluk: Number.isInteger(koyuluk) && koyuluk >= 0 && koyuluk <= 20 ? koyuluk : DY.koyuluk } };
  };
  async function etiketAyarOku() {
    if (config.sources.includes('shopify')) {
      try { const { deger } = await shopify.gizliOku('etiket'); if (deger) return { ayar: ayarTemizle(deger), varsayilan: ETIKET_VARSAYILAN }; }
      catch (e) { console.error('[etiket-ayar] gizli alan okunamadı:', e.message); }
    }
    let dosya = null;
    try { dosya = JSON.parse(fs.readFileSync(etiketDosyasi, 'utf8')); } catch {}
    return { ayar: ayarTemizle(dosya || {}), varsayilan: ETIKET_VARSAYILAN };
  }
  async function etiketAyarYaz(a) {
    const ayar = ayarTemizle(a);
    if (config.sources.includes('shopify')) {
      try { await shopify.gizliYaz('etiket', ayar); return { ayar }; }
      catch (e) { console.error('[etiket-ayar] gizli alana yazılamadı:', e.message); }
    }
    fs.mkdirSync(config.dataDir, { recursive: true });
    fs.writeFileSync(etiketDosyasi, JSON.stringify(ayar));
    return { ayar, uyari: 'Ayarlar sunucuda geçici dosyaya kaydedildi; sunucu yeniden kurulunca sıfırlanabilir.' };
  }
  // Seçilen siparişlerin etiket verisi (ekrandaki sırayla; durum değiştirmez)
  function etiketler({ sessionId, keys }) {
    const { s, selected } = take(sessionId, keys);
    return { etiketler: etiketVerisi(excel.ekranSirasi(selected), s.aktarim ? aktarimEtiketi(s.aktarim) : '') };
  }

  return { list, exportExcel, download, aktarimListesi, etiketler, etiketVerisi, etiketAyarOku, etiketAyarYaz };
})();

// ======================================================================
// OTOMATİK ADRES KONTROLÜ (yeni Shopify siparişleri, birkaç dakikada bir)
// ======================================================================
const oto = (() => {
  // Bekleyen (Excel'e aktarılmamış) Shopify siparişlerinin adresleri düzenli aralıklarla kontrol edilir.
  // Aynı adres bir kez kontrol edildikten sonra, adres değişmedikçe tekrar haritaya sorulmaz.
  const cfg = config.adres;
  const checked = new Map(); // sipariş → kontrol edilen adresin özeti
  const durum = { duzeltme: cfg.duzeltme, acik: cfg.duzeltme && cfg.otomatik && config.sources.includes('shopify'), aralikDk: cfg.aralikDk, son: null, duzeltilen: 0, hata: null };
  let calisiyor = false;

  const ozet = (a) => JSON.stringify([a && a.address1, a && a.address2, a && a.city]);

  async function tur() {
    if (calisiyor) return;
    calisiyor = true;
    try {
      const rows = await shopify.fetchRows({ stage: 'bekleyen' });
      const map = new Map();
      for (const r of rows) {
        if (!map.has(r._key)) map.set(r._key, { key: r._key, siparisNo: r.siparisNo, rows: [] });
        map.get(r._key).rows.push(r);
      }
      const bekleyen = [...map.values()].filter((o) => checked.get(o.key) !== ozet(o.rows[0]._ref.address));
      if (bekleyen.length) {
        const sonuc = await shopify.fixAddresses(bekleyen);
        bekleyen.forEach((o) => checked.set(o.key, ozet(o.rows[0]._ref.address)));
        durum.duzeltilen += sonuc.duzeltildi.length;
        if (sonuc.duzeltildi.length || sonuc.hata.length) {
          console.log(`[oto-adres] ${bekleyen.length} sipariş kontrol edildi, ${sonuc.duzeltildi.length} düzeltildi, ${sonuc.sorunlu.length} sorunlu, ${sonuc.hata.length} hata`);
        }
      }
      durum.son = new Date().toISOString();
      durum.hata = null;
    } catch (e) {
      durum.hata = e.message;
      console.error('[oto-adres]', e.message);
    } finally {
      calisiyor = false;
    }
  }

  function baslat() {
    if (!durum.acik) return;
    setTimeout(tur, 30 * 1000);
    setInterval(tur, cfg.aralikDk * 60 * 1000);
    console.log(`[oto-adres] açık: her ${cfg.aralikDk} dakikada bir yeni siparişler kontrol edilecek`);
  }

  return { baslat, durum };
})();

// ======================================================================
// ETİKET YAZICISI YARDIMCISI (Mac) — panel etiketleri PPLA olarak 127.0.0.1:9123'e gönderir;
// bu küçük zsh betiğini launchd her bağlantıda çalıştırır, veriyi Argox kuyruğuna ham (raw) basar.
// Kurulum: panelde "Etiket ayarları › Mac'e yardımcıyı kur" komutu Terminal'e yapıştırılır. Ek program gerekmez.
// ======================================================================
const yazici = (() => {
  const PORT = 9123, ETIKET = 'com.essejeffe.etiket';
  const YARDIMCI = `#!/bin/zsh
# Esse Jeffe etiket yardımcısı: launchd her bağlantı için bunu çalıştırır (girdi/çıktı = bağlantı).
# Tarayıcıdan gelen PPLA verisini (base64) Argox kuyruğuna ham (raw) olarak gönderir.
export LC_ALL=C PATH=/usr/bin:/bin:/usr/sbin:/sbin
D=\${0:A:h}
CR=$'\\r'
yanit() {
  printf 'HTTP/1.1 %s\\r\\nAccess-Control-Allow-Origin: *\\r\\nAccess-Control-Allow-Methods: GET, POST, OPTIONS\\r\\nAccess-Control-Allow-Headers: Content-Type\\r\\nAccess-Control-Allow-Private-Network: true\\r\\nContent-Type: application/json\\r\\nCache-Control: no-store\\r\\nContent-Length: %d\\r\\nConnection: close\\r\\n\\r\\n%s' "$1" \${#2} "$2"
  exit 0
}
jstr() { local s=\${1//\\\\/\\\\\\\\}; s=\${s//\\"/\\\\\\"}; s=\${s//$'\\n'/ }; s=\${s//$CR/}; s=\${s//$'\\t'/ }; print -rn -- "\\"$s\\""; }
kuyruk() {
  local k=""
  [[ -s $D/kuyruk ]] && k=$(<$D/kuyruk)
  [[ -z $k ]] && k=$(lpstat -e 2>/dev/null | grep -i argox | head -1)
  [[ -z $k ]] && k=$(lpstat -p 2>/dev/null | awk 'tolower($0) ~ /argox/ {print $2; exit}')
  print -rn -- $k
}
IFS= read -r ilk || exit 0
ilk=\${ilk%$CR}
yontem=\${ilk%% *}; yol=\${\${ilk#* }%% *}
uzunluk=0
while IFS= read -r h; do
  h=\${h%$CR}
  [[ -z $h ]] && break
  case \${(L)h} in content-length:*) uzunluk=\${\${h#*:}//[^0-9]/} ;; esac
done
[[ $yontem == OPTIONS ]] && yanit '204 No Content' ''
case "$yontem $yol" in
  'GET /durum')
    k=$(kuyruk)
    yanit '200 OK' "{\\"ok\\":true,\\"surum\\":1,\\"kuyruk\\":$(jstr $k)}" ;;
  'POST /yazdir')
    (( uzunluk > 0 && uzunluk < 60000000 )) || yanit '400 Bad Request' '{"ok":false,"hata":"bos istek"}'
    k=$(kuyruk)
    [[ -z $k ]] && yanit '500 Internal Server Error' '{"ok":false,"hata":"argox kuyrugu bulunamadi"}'
    t=$(mktemp -d /tmp/esse-etiket.XXXXXX) || yanit '500 Internal Server Error' '{"ok":false,"hata":"gecici dosya"}'
    head -c $uzunluk > $t/govde
    if ! base64 -D < $t/govde > $t/etiket.prn 2>/dev/null; then base64 -d < $t/govde > $t/etiket.prn 2>/dev/null; fi
    if [[ ! -s $t/etiket.prn ]]; then rm -rf $t; yanit '400 Bad Request' '{"ok":false,"hata":"veri cozulemedi"}'; fi
    cikti=$(lp -d "$k" -o raw $t/etiket.prn 2>&1); s=$?
    rm -rf $t
    (( s == 0 )) && yanit '200 OK' "{\\"ok\\":true,\\"kuyruk\\":$(jstr $k),\\"is\\":$(jstr $cikti)}"
    yanit '500 Internal Server Error' "{\\"ok\\":false,\\"kuyruk\\":$(jstr $k),\\"hata\\":$(jstr $cikti)}" ;;
esac
yanit '404 Not Found' '{"ok":false,"hata":"yok"}'
`;
  const plist = `<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0"><dict>
  <key>Label</key><string>${ETIKET}</string>
  <key>ProgramArguments</key><array><string>/bin/zsh</string><string>$D/yardimci.zsh</string></array>
  <key>inetdCompatibility</key><dict><key>Wait</key><false/></dict>
  <key>Sockets</key><dict><key>Dinle</key><dict>
    <key>SockNodeName</key><string>127.0.0.1</string><key>SockServiceName</key><string>${PORT}</string>
    <key>SockType</key><string>stream</string><key>SockFamily</key><string>IPv4</string>
  </dict></dict>
</dict></plist>`;
  const kurulum = `#!/bin/zsh
set -e
D="$HOME/.esse-etiket"; P="$HOME/Library/LaunchAgents/${ETIKET}.plist"
mkdir -p "$D" "$HOME/Library/LaunchAgents"
cat > "$D/yardimci.zsh" <<'YARDIMCI_SONU'
${YARDIMCI}YARDIMCI_SONU
chmod 755 "$D/yardimci.zsh"
K=$(lpstat -e 2>/dev/null | grep -i argox | head -1 || true)
print -r -- "$K" > "$D/kuyruk"
cat > "$P" <<PLIST_SONU
${plist}
PLIST_SONU
launchctl bootout gui/$(id -u) "$P" 2>/dev/null || true
launchctl bootstrap gui/$(id -u) "$P"
sleep 1
if curl -s -m 3 http://127.0.0.1:${PORT}/durum >/dev/null; then
  if [[ -n $K ]]; then echo "✅ Etiket yardımcısı kuruldu. Yazıcı: $K"; else echo "⚠️ Yardımcı kuruldu ama Argox yazıcısı bulunamadı (lpstat -p ile bakın)."; fi
else echo "❌ Yardımcı başlatılamadı."; fi
`;
  const kaldir = `launchctl bootout gui/$(id -u)/${ETIKET} 2>/dev/null; rm -rf ~/.esse-etiket ~/Library/LaunchAgents/${ETIKET}.plist; echo "Etiket yardımcısı kaldırıldı."`;
  // Terminal'e tek satır olarak yapıştırılır (base64: yapıştırırken bozulmaz)
  const komutlar = () => ({ port: PORT, kaldir,
    kur: "echo '" + Buffer.from(kurulum).toString('base64') + "' | base64 -D | zsh" });
  return { PORT, YARDIMCI, kurulum, komutlar };
})();

// ======================================================================
// PANEL SAYFASI
// ======================================================================
const PAGE = `<!doctype html>
<html lang="tr"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Esse Jeffe Siparişler</title>
<style>
  :root{--ink:#1f1719;--muted:#5e4f53;--brand:#3b2a2f;--bg:#f6f2ee;--line:#e4d9d3;--zebra:#fbf8f6;--warn:#b00020;--ok:#1d6b3a}
  *{box-sizing:border-box}
  body{font-family:"Segoe UI",system-ui,-apple-system,Roboto,"Helvetica Neue",Arial,sans-serif;background:var(--bg);color:var(--ink);
    margin:0;padding:12px 16px;font-size:14.5px;line-height:1.35;-webkit-font-smoothing:antialiased;font-variant-numeric:tabular-nums}
  .wrap{max-width:1900px;margin:0 auto}
  h1{font-size:18px;margin:0 0 10px;display:flex;align-items:baseline;gap:10px;flex-wrap:wrap}
  h1 small{font-size:12px;font-weight:400;color:var(--muted)}
  .card{background:#fff;border:1px solid var(--line);border-radius:10px;padding:12px}
  .bar{display:flex;flex-wrap:wrap;gap:8px;align-items:center;margin-bottom:8px}
  button{padding:8px 16px;font:inherit;font-weight:600;border:0;border-radius:8px;background:var(--brand);color:#fff;cursor:pointer}
  button.ghost{background:#efe8e3;color:var(--ink)}
  button.link{background:none;color:var(--muted);padding:6px 4px;text-decoration:underline;font-weight:500}
  button:disabled{opacity:.5;cursor:not-allowed}
  input[type=date]{padding:6px 8px;border:1px solid var(--line);border-radius:6px;font:inherit}
  .count{color:var(--muted)}
  .scroll{overflow:auto;max-height:calc(100vh - 150px);border:1px solid var(--line);border-radius:8px}
  table{border-collapse:collapse;width:100%;min-width:980px}
  th,td{padding:6px 9px;border-bottom:1px solid var(--line);text-align:left;vertical-align:top}
  th{background:#f3ece7;font-size:13px;font-weight:700;position:sticky;top:0;z-index:1;white-space:nowrap;color:var(--muted)}
  th:first-child,td:first-child{width:32px;padding-right:0}
  tbody tr:nth-child(even) td{background:var(--zebra)}
  tbody tr:hover td{background:#f3e9e3}
  tr.off td{opacity:.4}
  td.num{text-align:right;white-space:nowrap}
  td.nowrap{white-space:nowrap}
  td.addr{min-width:200px;max-width:280px}
  th.sr{cursor:pointer;user-select:none}
  th.sr:hover{color:var(--ink)}
  th.sr .ok-i{font-size:11px;margin-left:3px;color:var(--brand)}
  select{padding:7px 8px;border:1px solid var(--line);border-radius:8px;font:inherit;background:#fff}
  .tag{display:block;white-space:nowrap}
  .odeme{white-space:nowrap;font-weight:600}
  td.addr.bad{background:#fde8e8 !important;color:#8a1020;box-shadow:inset 3px 0 0 #c62828}
  .anote{display:block;margin-top:2px;font-size:12.5px;font-weight:600}
  .afix{display:block;margin-top:2px;font-size:12.5px;color:var(--ok)}
  .sku{display:block;font-family:ui-monospace,"Cascadia Mono",Consolas,Menlo,monospace;font-size:13.5px;font-weight:500;white-space:nowrap}
  .mono{font-family:ui-monospace,"Cascadia Mono",Consolas,Menlo,monospace;font-size:13px;color:var(--muted)}
  .tags{color:var(--muted);font-size:13px}
  input[type=checkbox]{width:16px;height:16px;margin:2px 0 0}
  .empty{padding:28px;text-align:center;color:var(--muted)}
  #msg{margin-bottom:8px;line-height:1.6}
  #msg:empty{display:none}
  .warn{color:var(--warn)} .ok{color:var(--ok)}
  a.dl{font-weight:700;color:var(--brand);margin-right:16px}
  .hide{display:none !important}
  .batches{display:flex;flex-wrap:wrap;gap:6px}
  button.chip{padding:6px 12px;border-radius:999px;background:#efe8e3;color:var(--ink);font-weight:600}
  button.chip.on{background:var(--brand);color:#fff}
  .muted{color:var(--muted)}
  body.gomulu{padding:8px}
  body.gomulu h1{font-size:0;margin:0 0 6px}
  body.gomulu h1 small{font-size:11px}
  .duzen{display:flex;gap:16px;flex-wrap:wrap;margin:8px 0}
  .duzen-sol{flex:1 1 420px;display:flex;flex-direction:column;gap:6px}
  .duzen-sag{flex:0 0 auto;display:flex;flex-direction:column;gap:6px}
  textarea#sablon,textarea#serbest{min-height:340px;font-family:ui-monospace,Menlo,Consolas,monospace;font-size:13px;line-height:1.45;
    padding:10px;border:1px solid var(--line);border-radius:8px;resize:vertical;tab-size:2}
  .yazici-kutu{border:1px solid var(--line);border-radius:10px;padding:10px 12px;margin-top:12px;display:flex;flex-direction:column;gap:6px}
  #yKurKomut{width:100%;box-sizing:border-box;font-family:ui-monospace,Menlo,monospace;font-size:11px;word-break:break-all}
  .onizleme{width:340px;height:340px;display:flex;align-items:center;justify-content:center;background:#e9e2dc;border-radius:10px;overflow:hidden}
  .onizleme iframe{flex:none;border:0;background:#fff;box-shadow:0 1px 6px rgba(0,0,0,.25);border-radius:6px;transform-origin:center}
</style></head>
<body><div class="wrap">
  <h1>Esse Jeffe Siparişler <small>{{SURUM}}</small></h1>
  <div class="card">
    <div class="bar" id="barMain">
      <button id="fetch">Siparişleri çek</button>
      <button id="export" disabled>Excel'e aktar</button>
      <button id="etiketBekleyen" class="ghost hide">🏷️ Etiketleri yazdır</button>
      <span class="count" id="count"></span>
      <span style="flex:1"></span>
      <label class="muted" for="sirala">Sırala</label>
      <select id="sirala" aria-label="Sıralama">
        <option value="kanal">Kanala göre</option>
        <option value="sku">SKU'ya göre</option>
        <option value="odeme">Ödeme yöntemine göre (önce kapıda)</option>
        <option value="tarih-yeni">Tarihe göre (yeni → eski)</option>
        <option value="tarih-eski">Tarihe göre (eski → yeni)</option>
        <option value="etiket">Etikete göre (değişim/iade önce)</option>
        <option value="ozel" hidden>Sütuna göre</option>
      </select>
      <button class="link" id="toPast">Önceki aktarımlar</button>
    </div>
    <div class="bar hide" id="barPast">
      <b>Önceki aktarımlar</b>
      <input type="date" id="gun" aria-label="Aktarım günü">
      <span id="batches" class="batches"></span>
      <span style="flex:1"></span>
      <button id="etiket" class="ghost" disabled>🏷️ Etiketleri yazdır</button>
      <button id="excel" class="ghost" disabled>Excel indir</button>
      <button class="link" id="toMain">← Bekleyen siparişler</button>
    </div>
    <div class="bar hide" id="barEtiket">
      <b id="etiketBaslik">Etiket ayarları</b>
      <span style="flex:1"></span>
      <button class="link" id="etiketGeri">← Siparişler</button>
    </div>
    <div class="muted" id="etiketBilgi" style="font-size:12px;margin:4px 0">🏷️ Etiketler bu panelden doğrudan yazıcıya gider · Ölçü: <span id="olcuYazi"></span> ·
      <button class="link" id="toAyar">Etiket ayarları</button> · <button class="link" id="toSerbest">Serbest etiket</button></div>
    <div id="llDurum"></div>
    <div id="msg"></div>
    <div id="ayarAlan" class="hide">
      <div class="bar">
        <label>Genişlik <input type="number" id="olcuG" min="10" max="300" step="0.5" style="width:80px"> mm</label>
        <label>Yükseklik <input type="number" id="olcuY" min="10" max="300" step="0.5" style="width:80px"> mm</label>
        <span class="muted" style="font-size:12px">Ölçü sipariş etiketi ve serbest etiket için ortaktır.</span>
      </div>
      <div class="duzen">
        <div class="duzen-sol">
          <label class="muted" for="sablon">Sipariş etiketi şablonu (HTML)</label>
          <textarea id="sablon" spellcheck="false"></textarea>
          <div class="muted" style="font-size:12px;line-height:1.6">Değişkenler: {{SKU}} {{BARKOD}} {{KARGO_ANAHTARI}} {{KARGO_FIRMASI}} {{KANAL}} {{MUSTERI}} {{SIPARIS_NO}}
            {{TELEFON}} {{ADRES}} {{TARIH}} {{ADET}} {{TUTAR}} {{ODEME}} {{DURUM}} {{ETIKETLER}} {{AKTARIM}} ·
            {{BARKOD:metin}} istediğiniz metnin barkodu · <b>cqw</b> = etiket genişliğinin %1'i (ölçü değişince düzen orantılı kalır)</div>
        </div>
        <div class="duzen-sag"><div class="muted" style="font-size:12px">Önizleme (ilk sipariş ya da örnek)</div><div class="onizleme" id="ayarOnizleme"></div></div>
      </div>
      <div class="bar">
        <button id="ayarKaydet">Kaydet</button>
        <button class="ghost" id="ayarDeneme">Deneme etiketi yazdır</button>
        <button class="ghost" id="ayarSifirla">Varsayılan şablona dön</button>
      </div>
      <div class="yazici-kutu">
        <b>Yazıcıya gönderme</b>
        <div class="bar">
          <label><input type="radio" name="yYol" value="dogrudan"> Doğrudan Argox'a (yazdırma penceresi açılmaz)</label>
          <label><input type="radio" name="yYol" value="pencere"> Yazdırma penceresiyle (eski yol)</label>
        </div>
        <div style="font-size:13px">Mac yardımcısı: <span id="yDurum"></span> <button class="link" id="yKontrol">Yeniden kontrol et</button></div>
        <div class="bar">
          <label>Resim biçimi <select id="yBicim"><option value="B">1</option><option value="b">2</option><option value="P">3</option><option value="p">4</option></select></label>
          <label><input type="checkbox" id="yCevir"> Baş aşağı çıkıyorsa çevir</label>
          <label><input type="checkbox" id="yTers"> Siyah/beyaz ters çıkıyorsa düzelt</label>
          <label>Koyuluk <select id="yKoyu"><option value="0">Yazıcının ayarı</option></select></label>
        </div>
        <div class="bar">
          <button class="ghost" id="yYaziTesti">Yazı testi</button>
          <button class="ghost" id="yResimTesti">Resim testi (seçili biçim)</button>
          <span class="muted" style="font-size:12px">Resim testinde numara düz ve okunur çıkıyorsa o biçimi seçip Kaydet'e basın; çıkmıyorsa sıradaki numarayı deneyin.</span>
        </div>
        <details id="yKurulum"><summary>Mac'e yardımcıyı kur / kaldır</summary>
          <ol style="margin:6px 0;padding-left:20px;font-size:13px;line-height:1.6">
            <li>Mac'te <b>Terminal</b>'i açın (⌘ + boşluk → "Terminal").</li>
            <li>Aşağıdaki komutu kopyalayıp Terminal'e yapıştırın, Enter'a basın. "✅ Etiket yardımcısı kuruldu" yazmalı.</li>
            <li>Bu sayfada "Yeniden kontrol et"e basın; tarayıcı izin sorarsa "İzin ver" deyin. Mac yeniden başlasa da yardımcı kendiliğinden çalışır.</li>
          </ol>
          <textarea id="yKurKomut" readonly rows="3"></textarea>
          <div class="bar"><button class="ghost" id="yKopyala">Komutu kopyala</button></div>
          <div class="muted" style="font-size:12px">Kaldırmak için Terminal'e: <code id="yKaldir"></code></div>
        </details>
      </div>
    </div>
    <div id="serbestAlan" class="hide">
      <div class="duzen">
        <div class="duzen-sol">
          <label class="muted" for="serbest">Serbest etiket (HTML) — ne yazarsanız o basılır</label>
          <textarea id="serbest" spellcheck="false"></textarea>
          <div class="muted" style="font-size:12px">{{BARKOD:metin}} ile barkod ekleyebilirsiniz · Ölçü "Etiket ayarları"ndan.</div>
        </div>
        <div class="duzen-sag"><div class="muted" style="font-size:12px">Önizleme</div><div class="onizleme" id="serbestOnizleme"></div></div>
      </div>
      <div class="bar">
        <label>Kopya <input type="number" id="kopya" min="1" max="200" value="1" style="width:70px"></label>
        <button id="serbestYazdir">🖨️ Yazdır</button>
        <button class="ghost" id="serbestKaydet">Şablonu kaydet</button>
        <button class="ghost" id="serbestSifirla">Boş şablona dön</button>
      </div>
    </div>
    <div class="scroll" id="tabloAlan"><table id="tablo">
      <thead><tr>
        <th><input type="checkbox" id="all"></th><th class="sr" data-k="kanal">Kanal</th><th class="sr" data-k="siparisNo">Sipariş</th><th class="sr" data-k="tarih">Tarih</th><th class="sr" data-k="musteri">Müşteri</th><th>Telefon</th>
        <th>Adres</th><th class="sr" data-k="sku">SKU</th><th class="sr" data-k="tutar">Tutar</th><th class="sr" data-k="odeme">Ödeme</th><th>Kargo anahtarı</th><th class="sr" data-k="durum">Durum</th><th class="sr" data-k="etiket">Etiketler</th>
      </tr></thead>
      <tbody id="rows"></tbody>
    </table></div>
  </div>
</div>
<script>
const COLS = 13;
if (window.self !== window.top) document.body.classList.add('gomulu');   // Dilara panelinin içinde: başlık küçülür
let mode = 'bekleyen', session = null, orders = [];
const $ = (id) => document.getElementById(id);
const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const money = (n) => n ? n.toLocaleString('tr-TR', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + ' ₺' : '';
const when = (d) => d ? new Date(d).toLocaleString('tr-TR', { day:'2-digit', month:'2-digit', hour:'2-digit', minute:'2-digit' }) : '';
const bugun = () => new Date().toLocaleDateString('en-CA', { timeZone: 'Europe/Istanbul' });
const saat = (d) => new Date(d).toLocaleTimeString('tr-TR', { timeZone: 'Europe/Istanbul', hour: '2-digit', minute: '2-digit' });
const empty = (t) => '<tr><td colspan="' + COLS + '" class="empty">' + t + '</td></tr>';
let aktarimNo = null;

// ---- Sıralama (seçim hatırlanır; Excel ve Google Sheet de ekrandaki sırayla yazılır) ----
const HAZIR = { kanal: { k: 'kanal', yon: 1 }, sku: { k: 'sku', yon: 1 }, odeme: { k: 'odeme', yon: 1 }, 'tarih-yeni': { k: 'tarih', yon: -1 }, 'tarih-eski': { k: 'tarih', yon: 1 }, etiket: { k: 'etiket', yon: 1 } };
// Satış kanalı: Shopify → web, sonra pazaryerleri
const KANAL_AD = { Shopify: 'web', Trendyol: 'trendyol', Hepsiburada: 'hepsiburada' };
const KANAL_SIRA = ['Shopify', 'Trendyol', 'Hepsiburada'];
const ODEME_SIRA = ['Kapıda', 'Havale/EFT', 'Havale/EFT (ödendi)', 'Kart', 'Ödendi', 'Değişim', 'İade', 'Pazaryeri'];
const coll = new Intl.Collator('tr', { numeric: true, sensitivity: 'base' });
const ilkSku = (o) => (o.skus || []).map((x) => x.sku).sort(coll.compare)[0] || '';
// v14: varsayılan kanala göre; anahtar değişti, v13'te hatırlanan seçim bir kez sıfırlanır
let sira = (() => { try { const x = JSON.parse(localStorage.getItem('siralama2')); if (x && x.k === 'kaynak') x.k = 'kanal'; return x || { ad: 'kanal', ...HAZIR.kanal }; } catch { return { ad: 'kanal', ...HAZIR.kanal }; } })();
function siraKaydet() { try { localStorage.setItem('siralama2', JSON.stringify(sira)); } catch {} }
function anahtar(o, k) {
  if (k === 'sku') return ilkSku(o);
  if (k === 'tarih') return new Date(o.tarih || 0).getTime();
  if (k === 'tutar') return Number(o.tutar) || 0;
  if (k === 'odeme') { const i = ODEME_SIRA.indexOf(o.odeme); return i < 0 ? 99 : i; }
  if (k === 'kanal') { const i = KANAL_SIRA.indexOf(o.kanal); return i < 0 ? 99 : i; }
  if (k === 'etiket') { const t = (o.etiketler || '').toLocaleLowerCase('tr'); return (/değişim|degisim/.test(t) ? '0' : /iade/.test(t) ? '1' : t ? '2' : '3') + t; }
  return String(o[k] ?? '');
}
function sirala(liste) {
  const { k, yon } = sira;
  const fark = (a, b) => { const x = anahtar(a, k), y = anahtar(b, k); return typeof x === 'number' ? x - y : coll.compare(x, y); };
  return [...liste].sort((a, b) => yon * fark(a, b) || coll.compare(ilkSku(a), ilkSku(b)) || coll.compare(a.siparisNo, b.siparisNo));
}
function siraGoster() {
  $('sirala').value = HAZIR[sira.ad] ? sira.ad : 'ozel';
  document.querySelectorAll('th.sr').forEach((th) => {
    th.querySelector('.ok-i')?.remove();
    if (th.dataset.k === sira.k) th.insertAdjacentHTML('beforeend', '<span class="ok-i">' + (sira.yon > 0 ? '▲' : '▼') + '</span>');
  });
}
$('sirala').onchange = (e) => { if (HAZIR[e.target.value]) { sira = { ad: e.target.value, ...HAZIR[e.target.value] }; siraKaydet(); render(); } };
document.querySelectorAll('th.sr').forEach((th) => th.onclick = () => {
  const k = th.dataset.k;
  sira = sira.k === k ? { ad: 'ozel', k, yon: -sira.yon } : { ad: 'ozel', k, yon: k === 'tarih' || k === 'tutar' ? -1 : 1 };
  siraKaydet(); render();
});

function setMode(m) {
  const etiketMi = m === 'ayar' || m === 'serbest';
  mode = m; session = null; orders = [];
  $('barMain').classList.toggle('hide', m !== 'bekleyen');
  $('barPast').classList.toggle('hide', m !== 'aktarim');
  $('barEtiket').classList.toggle('hide', !etiketMi);
  $('etiketBilgi').classList.toggle('hide', etiketMi);
  $('tabloAlan').classList.toggle('hide', etiketMi);
  $('ayarAlan').classList.toggle('hide', m !== 'ayar');
  $('serbestAlan').classList.toggle('hide', m !== 'serbest');
  $('msg').innerHTML = ''; $('llDurum').innerHTML = '';
  $('rows').innerHTML = empty('');
  update();
  if (m === 'ayar') {
    $('etiketBaslik').textContent = 'Etiket ayarları';
    $('olcuG').value = etiketAyar.genislik; $('olcuY').value = etiketAyar.yukseklik; $('sablon').value = etiketAyar.sablon; ayarOnizle(); yaziciFormDoldur();
  } else if (m === 'serbest') {
    $('etiketBaslik').textContent = 'Serbest etiket';
    $('serbest').value = etiketAyar.serbest; serbestOnizle();
  } else if (m === 'bekleyen') load();
  else { if (!$('gun').value) $('gun').value = bugun(); gunuYukle(); }
}
$('toPast').onclick = () => setMode('aktarim');

// Seçilen günün aktarımları: her biri bir düğme (#no · saat · adet); en sonuncusu açılır
async function gunuYukle() {
  session = null; orders = []; aktarimNo = null; update();
  $('batches').innerHTML = '<span class="muted">Yükleniyor…</span>';
  $('rows').innerHTML = empty('');
  try {
    const d = await call('api/aktarimlar?tarih=' + $('gun').value);
    if (!d.aktarimlar.length) {
      $('batches').innerHTML = '<span class="muted">Bu gün aktarım yok.</span>';
      $('rows').innerHTML = empty('Başka bir gün seçin.');
      return;
    }
    $('batches').innerHTML = d.aktarimlar.map((a) =>
      '<button class="chip" data-no="' + a.no + '">#' + a.no + ' · ' + saat(a.zaman) + ' · ' + a.adet + ' sipariş</button>').join('');
    document.querySelectorAll('button.chip').forEach((b) => b.onclick = () => aktarimiAc(Number(b.dataset.no)));
    aktarimiAc(d.aktarimlar[0].no);
  } catch (e) { $('batches').innerHTML = ''; $('msg').innerHTML = warnHtml([e.message]); }
}
$('gun').onchange = gunuYukle;

async function aktarimiAc(no) {
  aktarimNo = no;
  document.querySelectorAll('button.chip').forEach((b) => b.classList.toggle('on', Number(b.dataset.no) === no));
  $('msg').innerHTML = ''; $('rows').innerHTML = empty('Yükleniyor…');
  try {
    const d = await call('api/liste?liste=aktarim&no=' + no);
    if (mode !== 'aktarim' || aktarimNo !== no) return;
    session = d.sessionId; orders = d.orders; $('rows').innerHTML = ''; render();
    $('msg').innerHTML = '<b>Aktarım ' + esc(d.aktarim.etiket) + '</b> · ' + orders.length + ' sipariş' + warnHtml(d.warnings);
  } catch (e) { $('rows').innerHTML = empty(''); $('msg').innerHTML = warnHtml([e.message]); }
}
$('toMain').onclick = () => setMode('bekleyen');

// ---- Etiketler: panel HTML şablondan çizer, tek yazdırma penceresiyle yazıcıya gönderir (Label LIVE yok) ----
// Code128 (B) barkod: SVG
const C128 = ['212222','222122','222221','121223','121322','131222','122213','122312','132212','221213','221312','231212','112232','122132',
  '122231','113222','123122','123221','223211','221132','221231','213212','223112','312131','311222','321122','321221','312212','322112',
  '322211','212123','212321','232121','111323','131123','131321','112313','132113','132311','211313','231113','231311','112133','112331',
  '132131','113123','113321','133121','313121','211331','231131','213113','213311','213131','311123','311321','331121','312113','312311',
  '332111','314111','221411','431111','111224','111422','121124','121421','141122','141221','112214','112412','122114','122411','142112',
  '142211','241211','221114','413111','241112','134111','111242','121142','121241','114212','124112','124211','411212','421112','421211',
  '212141','214121','412121','111143','111341','131141','114113','114311','411113','411311','113141','114131','311141','411131','211412',
  '211214','211232','2331112'];
function barkodSvg(metin) {
  const t = String(metin || '').replace(/[^\\x20-\\x7e]/g, '?');
  if (!t) return '';
  // Rakam dizileri C kümesiyle (iki rakam bir kod) yazılır: barkod kısalır, çubuklar kalınlaşır
  const rakam = (i) => { let n = 0; while (i + n < t.length && t[i + n] >= '0' && t[i + n] <= '9') n++; return n; };
  let kume = rakam(0) >= 4 || (rakam(0) === t.length && t.length >= 2 && t.length % 2 === 0) ? 'C' : 'B';
  const kod = [kume === 'C' ? 105 : 104];
  for (let i = 0; i < t.length;) {
    if (kume === 'C') {
      if (rakam(i) >= 2) { kod.push(Number(t.substr(i, 2))); i += 2; continue; }
      kod.push(100); kume = 'B';
    }
    const n = rakam(i);
    if (n >= 4 && (n >= 6 || i + n === t.length)) {
      if (n % 2) { kod.push(t.charCodeAt(i) - 32); i++; }
      kod.push(99); kume = 'C'; continue;
    }
    kod.push(t.charCodeAt(i) - 32); i++;
  }
  let top = kod[0];
  for (let i = 1; i < kod.length; i++) top += kod[i] * i;
  kod.push(top % 103, 106);
  let x = 10, rect = '', moduller = '0'.repeat(10);   // 10 modül sessiz alan
  for (const k of kod) {
    const desen = C128[k];
    for (let i = 0; i < desen.length; i++) {
      const w = Number(desen[i]);
      if (i % 2 === 0) rect += '<rect x="' + x + '" y="0" width="' + w + '" height="100"/>';
      moduller += (i % 2 === 0 ? '1' : '0').repeat(w);
      x += w;
    }
  }
  moduller += '0'.repeat(10);
  return '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ' + (x + 10) + ' 100" preserveAspectRatio="none" shape-rendering="crispEdges"' +
    ' data-desen="' + moduller + '" style="display:block;width:100%;height:100%"><g fill="#000">' + rect + '</g></svg>';
}
const unesc = (s) => String(s).replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&#39;/g, "'").replace(/&amp;/g, '&');
// Şablonu bir etiketin değerleriyle doldurur
function sablonDoldur(sablon, v) {
  v = v || {};
  let h = String(sablon || '').replace(/\\{\\{\\s*([A-Z_]+)\\s*\\}\\}/g, (m, ad) => {
    if (ad === 'BARKOD') return barkodSvg(v.KARGO_ANAHTARI);
    if (!(ad in v)) return m;
    return esc(v[ad]).replace(/\\n/g, '<br>');
  });
  return h.replace(/\\{\\{\\s*BARKOD:([^}]*)\\}\\}/g, (m, metin) => barkodSvg(unesc(metin.trim())));
}
// Her etiket bir sayfa; sayfa boyutu = etiket ölçüsü
function etiketBelgesi(sayfalar) {
  const g = etiketAyar.genislik, y = etiketAyar.yukseklik;
  return '<!doctype html><html><head><meta charset="utf-8"><style>@page{size:' + g + 'mm ' + y + 'mm;margin:0}' +
    'html,body{margin:0;padding:0;background:#fff;-webkit-print-color-adjust:exact;print-color-adjust:exact}' +
    '.sayfa{width:' + g + 'mm;height:' + y + 'mm;overflow:hidden;position:relative;box-sizing:border-box;container-type:size;break-after:page;page-break-after:always}' +
    '.sayfa:last-child{break-after:auto;page-break-after:auto}</style></head><body>' +
    sayfalar.map((x) => '<div class="sayfa">' + x + '</div>').join('') + '</body></html>';
}
// Gizli çerçevede yazdırma penceresini açar (Mac'te yazıcı olarak Argox seçilir; ayar bir kez hatırlanır)
function belgeYazdir(html) {
  return new Promise((bitti) => {
    const f = document.createElement('iframe');
    f.setAttribute('aria-hidden', 'true');
    f.style.cssText = 'position:fixed;right:0;bottom:0;width:0;height:0;border:0;visibility:hidden';
    f.onload = async () => {
      try { await f.contentDocument.fonts.ready; } catch {}
      await new Promise((r) => setTimeout(r, 150));
      try { f.contentWindow.focus(); f.contentWindow.print(); } finally { setTimeout(() => { f.remove(); bitti(); }, 1000); }
    };
    f.srcdoc = html;
    document.body.appendChild(f);
  });
}
// Önizleme: tek etiket, kutuya sığacak şekilde küçültülmüş
function onizle(kutu, icerik) {
  const g = etiketAyar.genislik, y = etiketAyar.yukseklik, mmPx = 96 / 25.4;
  const olcek = Math.min(320 / (g * mmPx), 320 / (y * mmPx), 3);
  let f = kutu.querySelector('iframe');
  if (!f) { f = document.createElement('iframe'); f.title = 'Etiket önizleme'; kutu.appendChild(f); }
  f.style.width = g + 'mm'; f.style.height = y + 'mm'; f.style.transform = 'scale(' + olcek + ')';
  f.srcdoc = etiketBelgesi([icerik]);
}

let etiketAyar = { genislik: 100, yukseklik: 100, sablon: '', serbest: '' }, etiketVarsayilan = null;
function olcuGoster() { $('olcuYazi').textContent = etiketAyar.genislik + ' × ' + etiketAyar.yukseklik + ' mm'; }
async function etiketAyarYukle() {
  try { const d = await call('api/etiket-ayar'); etiketAyar = d.ayar; etiketVarsayilan = d.varsayilan; olcuGoster(); }
  catch (e) { $('llDurum').innerHTML = warnHtml(['Etiket ayarları okunamadı: ' + e.message]); }
}
const ORNEK = { SKU: 'K-060-YESIL-S\\nK-092-MAVI-S', KARGO_ANAHTARI: 'EJ1001599626', KARGO_FIRMASI: 'DHL eCommerce', KANAL: 'website',
  MUSTERI: 'çiğdem bay', SIPARIS_NO: '#100159', TELEFON: '05551112233', ADRES: 'Gül Sok. No 5, Kadıköy, İstanbul', TARIH: '26.11.2026 10:00',
  ADET: '2', TUTAR: '2798', ODEME: 'Kapıda', DURUM: 'Gönderilmedi', ETIKETLER: '', AKTARIM: '#1 · 26.11.2026 10:00' };
let ornekVeri = ORNEK;

let sonEtiketler = null, yaziliyor = false;
async function etiketYazdir(liste) {
  if (!liste || !liste.length) return;
  if (!etiketAyar.sablon) await etiketAyarYukle();
  await etiketGonder(liste.map((v) => sablonDoldur(etiketAyar.sablon, v)));
}

// ---- Doğrudan yazdırma (v20): etiket 203 dpi siyah-beyaz resme çevrilir, PPLA komutlarıyla Mac'teki yardımcıya
// (127.0.0.1:9123) gönderilir; yardımcı Argox kuyruğuna ham basar. Yazdırma penceresi/sürücü aradan çıkar.
const YARDIMCI = 'http://127.0.0.1:9123/';
const NOKTA_MM = 8;   // 203 dpi ≈ 8 nokta/mm
const yaziciAyar = () => etiketAyar.yazici || { yol: 'dogrudan', bicim: 'B', cevir: false, ters: false, koyuluk: 0 };
async function yardimciDurum() {
  try {
    const r = await fetch(YARDIMCI + 'durum', { cache: 'no-store', signal: AbortSignal.timeout(3000) });
    return await r.json();
  } catch { return null; }
}
async function yardimciyaGonder(bayt) {
  let r;
  try {
    r = await fetch(YARDIMCI + 'yazdir', { method: 'POST', headers: { 'Content-Type': 'text/plain' }, body: b64(bayt), signal: AbortSignal.timeout(120000) });
  } catch { throw Object.assign(new Error("Mac'teki etiket yardımcısına ulaşılamadı."), { yardimciYok: true }); }
  const d = await r.json().catch(() => ({}));
  if (!r.ok || !d.ok) throw new Error('Yazıcıya gönderilemedi: ' + (d.hata || ('hata ' + r.status)));
  return d;
}
const metinBayt = (t) => Uint8Array.from(t, (c) => c.charCodeAt(0) & 255);
function birlestir(parcalar) {
  let n = 0; for (const p of parcalar) n += p.length;
  const u = new Uint8Array(n); let i = 0;
  for (const p of parcalar) { u.set(p, i); i += p.length; }
  return u;
}
function b64(u) {
  let s = '';
  for (let i = 0; i < u.length; i += 0x8000) s += String.fromCharCode.apply(null, u.subarray(i, i + 0x8000));
  return btoa(s);
}

// Bir etiketi (HTML) nokta matrisine çevirir (1 = siyah). Barkodlar tam sayı nokta genişliğinde yeniden çizilir: okunur kalır.
async function etiketResmi(icerik) {
  const W = Math.round(etiketAyar.genislik * NOKTA_MM), H = Math.round(etiketAyar.yukseklik * NOKTA_MM);
  const stil = 'html,body{margin:0;padding:0;background:#fff}.sayfa{width:' + W + 'px;height:' + H +
    'px;overflow:hidden;position:relative;box-sizing:border-box;container-type:size;background:#fff;color:#000}';
  // 1) Gerçek boyutta (1 piksel = 1 nokta) görünmez çerçevede yerleşim; barkodların yeri ölçülür
  const f = document.createElement('iframe');
  f.setAttribute('aria-hidden', 'true');
  f.style.cssText = 'position:fixed;left:-20000px;top:0;width:' + W + 'px;height:' + H + 'px;border:0;visibility:hidden';
  const yuklendi = new Promise((r) => { f.onload = r; });
  f.srcdoc = '<!doctype html><html><head><meta charset="utf-8"><style>' + stil + '</style></head><body><div class="sayfa">' + icerik + '</div></body></html>';
  document.body.appendChild(f);
  try {
    await yuklendi;
    const d = f.contentDocument;
    try { await d.fonts.ready; } catch {}
    const sayfa = d.querySelector('.sayfa'), k = sayfa.getBoundingClientRect();
    const barkodlar = [...d.querySelectorAll('svg[data-desen]')].map((s) => {
      const r = s.getBoundingClientRect();
      return { x: r.left - k.left, y: r.top - k.top, w: r.width, h: r.height, desen: s.getAttribute('data-desen') };
    });
    // 2) Aynı yerleşim SVG (foreignObject) olarak tuvale çizilir
    const svg = '<svg xmlns="http://www.w3.org/2000/svg" width="' + W + '" height="' + H + '"><foreignObject x="0" y="0" width="' + W + '" height="' + H + '">' +
      '<div xmlns="http://www.w3.org/1999/xhtml"><style>' + stil + '</style>' + new XMLSerializer().serializeToString(sayfa) + '</div></foreignObject></svg>';
    const img = new Image();
    await new Promise((ok, hata) => {
      img.onload = ok; img.onerror = () => hata(new Error('Etiket resme çevrilemedi (şablonda hatalı HTML olabilir).'));
      img.src = 'data:image/svg+xml;charset=utf-8,' + encodeURIComponent(svg);
    });
    const c = document.createElement('canvas'); c.width = W; c.height = H;
    const g = c.getContext('2d', { willReadFrequently: true });
    g.fillStyle = '#fff'; g.fillRect(0, 0, W, H); g.drawImage(img, 0, 0, W, H);
    let p;
    try { p = g.getImageData(0, 0, W, H).data; }
    catch { throw new Error('Bu tarayıcı etiketi resme çeviremiyor; paneli Chrome ile açın ya da yazdırma penceresini kullanın.'); }
    const nokta = new Uint8Array(W * H);
    for (let i = 0, j = 0; i < W * H; i++, j += 4) if (p[j] * 0.299 + p[j + 1] * 0.587 + p[j + 2] * 0.114 < 150) nokta[i] = 1;
    for (const b of barkodlar) barkodCiz(nokta, W, H, b);
    return { W, H, nokta };
  } finally { f.remove(); }
}
function barkodCiz(nokta, W, H, b) {
  const n = (b.desen || '').length;
  if (!n || b.w < n) return;   // yer yetmiyorsa çizildiği gibi kalır
  const m = Math.floor(b.w / n), x0 = Math.round(b.x + (b.w - m * n) / 2);
  const y0 = Math.max(0, Math.round(b.y)), y1 = Math.min(H, Math.round(b.y + b.h));
  const xa = Math.max(0, Math.floor(b.x)), xb = Math.min(W, Math.ceil(b.x + b.w));
  for (let y = y0; y < y1; y++) {
    const s = y * W;
    for (let x = xa; x < xb; x++) nokta[s + x] = 0;
    for (let i = 0; i < n; i++) if (b.desen[i] === '1') {
      for (let q = 0; q < m; q++) { const x = x0 + i * m + q; if (x >= 0 && x < W) nokta[s + x] = 1; }
    }
  }
}
// Tek renkli BMP / PCX (bit 0 = siyah; "ters" seçiliyse tersi)
function bmp1(r, ters) {
  const { W, H, nokta } = r, satir = Math.ceil(W / 32) * 4, veri = satir * H, bas = 62;
  const u = new Uint8Array(bas + veri), v = new DataView(u.buffer);
  u[0] = 0x42; u[1] = 0x4d; v.setUint32(2, bas + veri, true); v.setUint32(10, bas, true);
  v.setUint32(14, 40, true); v.setInt32(18, W, true); v.setInt32(22, H, true); v.setUint16(26, 1, true); v.setUint16(28, 1, true);
  v.setUint32(34, veri, true); v.setUint32(38, 7992, true); v.setUint32(42, 7992, true); v.setUint32(46, 2, true);
  u[58] = 255; u[59] = 255; u[60] = 255;   // renk 0 = siyah, 1 = beyaz
  for (let y = 0; y < H; y++) {
    const o = bas + (H - 1 - y) * satir;    // BMP satırları alttan üste
    for (let x = 0; x < satir * 8; x++) {
      const siyah = x < W ? nokta[y * W + x] === 1 : false;
      if (siyah === ters) u[o + (x >> 3)] |= 0x80 >> (x & 7);
    }
  }
  return u;
}
function pcx1(r, ters) {
  const { W, H, nokta } = r, satir = Math.ceil(W / 16) * 2;
  const bas = new Uint8Array(128), v = new DataView(bas.buffer);
  bas[0] = 10; bas[1] = 5; bas[2] = 1; bas[3] = 1;
  v.setUint16(8, W - 1, true); v.setUint16(10, H - 1, true); v.setUint16(12, 203, true); v.setUint16(14, 203, true);
  bas[19] = 255; bas[20] = 255; bas[21] = 255;   // renk 0 = siyah, 1 = beyaz
  bas[65] = 1; v.setUint16(66, satir, true); v.setUint16(68, 1, true);
  const cikti = [bas], s = new Uint8Array(satir);
  for (let y = 0; y < H; y++) {
    s.fill(0);
    for (let x = 0; x < satir * 8; x++) {
      const siyah = x < W ? nokta[y * W + x] === 1 : false;
      if (siyah === ters) s[x >> 3] |= 0x80 >> (x & 7);
    }
    const rle = [];
    for (let i = 0; i < satir;) {
      let n = 1;
      while (i + n < satir && n < 63 && s[i + n] === s[i]) n++;
      if (n > 1 || s[i] >= 0xc0) rle.push(0xc0 | n);
      rle.push(s[i]); i += n;
    }
    cikti.push(Uint8Array.from(rle));
  }
  return birlestir(cikti);
}
// Bir etiketin PPLA komutları: resmi yazıcı belleğine yükle → etikete yerleştir → bas → resmi sil
function pplaEtiket(r, y) {
  const ad = 'EJETIKET';
  if (y.cevir) r = { ...r, nokta: r.nokta.slice().reverse() };
  const resim = y.bicim === 'P' || y.bicim === 'p' ? pcx1(r, y.ters) : bmp1(r, y.ters);
  return [metinBayt('\\x02IA' + y.bicim + ad + '\\r'), resim,
    metinBayt('\\x02L\\r' + 'D11\\r' + (y.koyuluk ? 'H' + String(y.koyuluk).padStart(2, '0') + '\\r' : '') +
      '1Y1100000000000' + ad + '\\r' + 'E\\r' + '\\x02xAG' + ad + '\\r')];
}
async function dogrudanBas(sayfalar, ilerleme, y = yaziciAyar()) {
  const durum = await yardimciDurum();
  if (!durum) throw Object.assign(new Error("Mac'teki etiket yardımcısına ulaşılamadı."), { yardimciYok: true });
  if (!durum.kuyruk) throw new Error('Yardımcı çalışıyor ama Argox yazıcısı bulunamadı (yazıcı Mac’e bağlı ve açık mı?).');
  const onbellek = new Map();   // aynı içerik (serbest etiket kopyaları) bir kez çizilir
  for (let i = 0; i < sayfalar.length; i += 20) {   // 20'şer etiketlik parçalar halinde
    const parca = [];
    for (let j = i; j < Math.min(i + 20, sayfalar.length); j++) {
      if (ilerleme) ilerleme(j + 1);
      if (!onbellek.has(sayfalar[j])) onbellek.set(sayfalar[j], await etiketResmi(sayfalar[j]));
      parca.push(...pplaEtiket(onbellek.get(sayfalar[j]), y));
    }
    await yardimciyaGonder(birlestir(parca));
  }
}
// Tüm "Yazdır" düğmeleri buradan geçer: ayara göre doğrudan (PPLA) ya da yazdırma penceresi
async function etiketGonder(sayfalar) {
  if (yaziliyor || !sayfalar || !sayfalar.length) return;
  yaziliyor = true;
  const n = sayfalar.length, yaz = (h) => { $('llDurum').innerHTML = h; };
  try {
    if (yaziciAyar().yol === 'pencere') {
      yaz('<span class="muted">🏷️ ' + n + ' etiket hazırlanıyor… Yazdırma penceresinde yazıcı olarak Argox&#39;u seçin.</span>');
      await belgeYazdir(etiketBelgesi(sayfalar));
      yaz('<span class="ok">🏷️ ' + n + ' etiket yazdırma penceresine gönderildi.</span>');
      return;
    }
    await dogrudanBas(sayfalar, (i) => yaz('<span class="muted">🏷️ Etiket ' + i + ' / ' + n + ' hazırlanıyor…</span>'));
    yaz('<span class="ok">🏷️ ' + n + ' etiket yazıcıya gönderildi.</span>');
  } catch (e) {
    yaz(warnHtml([e.message + (e.yardimciYok ? " Yardımcı bu Mac'e kurulu mu? (Etiket ayarları › Mac'e yardımcıyı kur) Tarayıcı izin sorduysa “İzin ver” deyin." : '')]) +
      ' <button class="link" id="pencereyleBas">Yazdırma penceresiyle yazdır</button>');
    $('pencereyleBas').onclick = () => belgeYazdir(etiketBelgesi(sayfalar));
  } finally { yaziliyor = false; }
}

// ---- Etiket ayarları › Yazıcıya gönderme ----
for (let k = 4; k <= 20; k += 2) $('yKoyu').insertAdjacentHTML('beforeend', '<option value="' + k + '">' + k + '</option>');
function yaziciFormDoldur() {
  const y = yaziciAyar();
  document.querySelectorAll('input[name=yYol]').forEach((r) => { r.checked = r.value === y.yol; });
  $('yBicim').value = y.bicim; $('yCevir').checked = y.cevir; $('yTers').checked = y.ters; $('yKoyu').value = String(y.koyuluk);
  yaziciKontrol();
}
function yaziciFormOku() {
  const yol = document.querySelector('input[name=yYol]:checked');
  return { yol: yol ? yol.value : 'dogrudan', bicim: $('yBicim').value, cevir: $('yCevir').checked, ters: $('yTers').checked,
    koyuluk: Number($('yKoyu').value) || 0 };
}
async function yaziciKontrol() {
  $('yDurum').innerHTML = '<span class="muted">kontrol ediliyor…</span>';
  const d = await yardimciDurum();
  $('yDurum').innerHTML = !d ? '<span class="warn">bulunamadı — bu Mac’e kurulu değil ya da tarayıcı izin vermedi</span>'
    : d.kuyruk ? '<span class="ok">✅ çalışıyor · yazıcı: ' + esc(d.kuyruk) + '</span>'
    : '<span class="warn">çalışıyor ama Argox yazıcısı bulunamadı (yazıcı bağlı ve açık mı?)</span>';
}
$('yKontrol').onclick = yaziciKontrol;
$('yKurulum').addEventListener('toggle', async () => {
  if (!$('yKurulum').open || $('yKurKomut').value) return;
  try { const d = await call('api/yazici-kurulum'); $('yKurKomut').value = d.kur; $('yKaldir').textContent = d.kaldir; }
  catch (e) { $('yKurKomut').value = 'Komut alınamadı: ' + e.message; }
});
$('yKopyala').onclick = async () => {
  try { await navigator.clipboard.writeText($('yKurKomut').value); $('yKopyala').textContent = 'Kopyalandı ✓'; }
  catch { $('yKurKomut').select(); document.execCommand('copy'); $('yKopyala').textContent = 'Kopyalandı ✓'; }
  setTimeout(() => { $('yKopyala').textContent = 'Komutu kopyala'; }, 2000);
};
// Yazı testi: resim olmadan, yazıcının kendi yazı tipiyle tek satır (yardımcı ve yazıcı bağlantısını dener)
$('yYaziTesti').onclick = async () => {
  try {
    const durum = await yardimciDurum();
    if (!durum) throw new Error("Mac'teki etiket yardımcısına ulaşılamadı.");
    await yardimciyaGonder(metinBayt('\\x02L\\r' + 'D11\\r' + '121100000500050ESSE JEFFE TEST\\r' + 'E\\r'));
    $('llDurum').innerHTML = '<span class="ok">Yazı testi gönderildi: etikette “ESSE JEFFE TEST” yazmalı.</span>';
  } catch (e) { $('llDurum').innerHTML = warnHtml([e.message]); }
};
// Resim testi: seçili biçimle numaralı bir deneme etiketi (doğru çıkan biçim numarası seçilip kaydedilir)
$('yResimTesti').onclick = () => {
  const y = yaziciFormOku(), no = $('yBicim').selectedOptions[0].textContent;
  etiketAyar.genislik = Number($('olcuG').value) || etiketAyar.genislik;
  etiketAyar.yukseklik = Number($('olcuY').value) || etiketAyar.yukseklik;
  const icerik = '<div style="height:100%;box-sizing:border-box;border:1.2cqw solid #000;display:flex;flex-direction:column;align-items:center;' +
    'justify-content:center;gap:2cqw;font-family:Arial,Helvetica,sans-serif;color:#000;text-align:center">' +
    '<div style="font-size:7cqw;font-weight:700">↑ ÜST ↑</div><div style="font-size:34cqw;font-weight:700;line-height:1">' + esc(no) + '</div>' +
    '<div style="font-size:6cqw">çğışöü ÇĞİŞÖÜ</div><div style="width:72cqw;height:16cqw">' + barkodSvg('EJ1001599626') + '</div>' +
    '<div style="font-size:4.5cqw">EJ1001599626</div></div>';
  etiketGonderAyarla(y, [icerik]);
};
async function etiketGonderAyarla(y, sayfalar) {
  const once = etiketAyar.yazici;
  etiketAyar.yazici = { ...y, yol: 'dogrudan' };
  try { await etiketGonder(sayfalar); } finally { etiketAyar.yazici = once; }
}

// ---- Etiket ayarları / Serbest etiket ekranları ----
function ayarOnizle() {
  etiketAyar.genislik = Number($('olcuG').value) || etiketAyar.genislik;
  etiketAyar.yukseklik = Number($('olcuY').value) || etiketAyar.yukseklik;
  onizle($('ayarOnizleme'), sablonDoldur($('sablon').value, ornekVeri));
}
function serbestOnizle() { onizle($('serbestOnizleme'), sablonDoldur($('serbest').value, {})); }
let onizZaman = null;
const gecikmeli = (fn) => () => { clearTimeout(onizZaman); onizZaman = setTimeout(fn, 250); };
$('sablon').oninput = gecikmeli(ayarOnizle);
$('olcuG').oninput = gecikmeli(ayarOnizle); $('olcuY').oninput = gecikmeli(ayarOnizle);
$('serbest').oninput = gecikmeli(serbestOnizle);
// Tab tuşu metin kutusunda girinti yapsın
for (const id of ['sablon', 'serbest']) $(id).addEventListener('keydown', (e) => {
  if (e.key !== 'Tab') return; e.preventDefault();
  const t = e.target, b = t.selectionStart; t.setRangeText('  ', b, t.selectionEnd, 'end'); t.dispatchEvent(new Event('input'));
});
async function etiketKaydet(degisen, mesajYeri) {
  try {
    const d = await post('api/etiket-ayar', { ...etiketAyar, ...degisen });
    etiketAyar = d.ayar; olcuGoster();
    $('msg').innerHTML = '<span class="ok">Kaydedildi.</span>' + warnHtml(d.uyari ? [d.uyari] : []);
  } catch (e) { $('msg').innerHTML = warnHtml([e.message]); }
}
$('ayarKaydet').onclick = () => etiketKaydet({ genislik: Number($('olcuG').value), yukseklik: Number($('olcuY').value), sablon: $('sablon').value,
  yazici: yaziciFormOku() });
$('ayarSifirla').onclick = () => { if (etiketVarsayilan) { $('sablon').value = etiketVarsayilan.sablon; ayarOnizle(); } };
$('ayarDeneme').onclick = async () => {
  ayarOnizle();
  const once = etiketAyar.yazici; etiketAyar.yazici = yaziciFormOku();   // kaydetmeden ekrandaki yazıcı seçimiyle dener
  try { await etiketGonder([sablonDoldur($('sablon').value, ornekVeri)]); } finally { etiketAyar.yazici = once; }
};
$('serbestKaydet').onclick = () => etiketKaydet({ serbest: $('serbest').value });
$('serbestSifirla').onclick = () => { if (etiketVarsayilan) { $('serbest').value = etiketVarsayilan.serbest; serbestOnizle(); } };
$('serbestYazdir').onclick = () => {
  const n = Math.min(200, Math.max(1, Number($('kopya').value) || 1));
  const tek = sablonDoldur($('serbest').value, {});
  etiketGonder(Array.from({ length: n }, () => tek));
};
async function etiketEkrani(tur) {
  await etiketAyarYukle();
  // Önizlemede ekrandaki ilk sipariş kullanılır (varsa)
  if (session && selected().length) {
    try { const d = await post('api/etiket', { sessionId: session, keys: selected().slice(0, 1) }); ornekVeri = d.etiketler[0] || ORNEK; } catch { ornekVeri = ORNEK; }
  } else ornekVeri = ORNEK;
  setMode(tur);
}
$('toAyar').onclick = () => etiketEkrani('ayar');
$('toSerbest').onclick = () => etiketEkrani('serbest');
$('etiketGeri').onclick = () => setMode('bekleyen');
etiketAyarYukle();
$('etiketBekleyen').onclick = () => etiketYazdir(sonEtiketler);
$('etiket').onclick = async () => {
  try { const d = await post('api/etiket', { sessionId: session, keys: selected() }); await etiketYazdir(d.etiketler); }
  catch (e) { $('llDurum').innerHTML = warnHtml([e.message]); }
};

function selected() { return [...document.querySelectorAll('.pick:checked')].map((c) => c.value); }
function update() {
  const n = selected().length;
  $('export').disabled = !n; $('excel').disabled = !n; $('etiket').disabled = !n;
  $('etiket').textContent = n ? '🏷️ Etiketleri yazdır (' + n + ')' : '🏷️ Etiketleri yazdır';
  $('export').textContent = n ? "Excel'e aktar (" + n + ')' : "Excel'e aktar";
  $('excel').textContent = n ? 'Excel indir (' + n + ')' : 'Excel indir';
  $('count').textContent = mode === 'bekleyen' && session ? orders.length + ' bekleyen sipariş' : '';
  $('all').checked = n && n === orders.length;
  document.querySelectorAll('.pick').forEach((c) => c.closest('tr').classList.toggle('off', !c.checked));
}
$('all').onchange = (e) => { document.querySelectorAll('.pick').forEach((c) => c.checked = e.target.checked); update(); };

function addrCell(o) {
  const bad = mode === 'bekleyen' && o.adresDurum === 'sorunlu';
  return '<td class="addr' + (bad ? ' bad' : '') + '">' + esc(o.adres) +
    (bad ? '<span class="anote">⚠ ' + esc(o.adresNot) + '</span>' : '') +
    (o.adresDuzeltildi ? '<span class="afix">✓ adres otomatik düzeltildi</span>' : '') + '</td>';
}

function render() {
  siraGoster();
  if (!orders.length) {
    $('rows').innerHTML = empty(mode === 'bekleyen' ? "Excel'e aktarılmamış sipariş yok." : 'Bu aktarımdaki siparişler bulunamadı.');
    return update();
  }
  // Seçimler korunur (yeniden sıralamada işaretler kaybolmaz)
  const secili = new Set(document.querySelectorAll('.pick').length ? selected() : orders.map((o) => o.key));
  $('rows').innerHTML = sirala(orders).map((o) =>
    '<tr><td><input type="checkbox" class="pick" value="' + esc(o.key) + '"' + (secili.has(o.key) ? ' checked' : '') + '></td>' +
    '<td class="nowrap">' + esc(KANAL_AD[o.kanal] || o.kanal) + '</td><td class="nowrap"><b>' + esc(o.siparisNo) + '</b></td><td class="nowrap">' + when(o.tarih) + '</td>' +
    '<td>' + esc(o.musteri) + '</td><td class="nowrap">' + esc(o.telefon) + '</td>' + addrCell(o) +
    '<td>' + o.skus.map((x) => '<span class="sku">' + esc(x.sku) + (x.adet > 1 ? ' ×' + esc(x.adet) : '') + '</span>').join('') + '</td>' +
    '<td class="num">' + money(o.tutar) + '</td><td class="odeme" title="' + esc(o.odemeTipi) + '">' + esc(o.odeme || o.odemeTipi) + '</td>' +
    '<td class="mono">' + esc(o.kargoAnahtari) + '</td><td>' + esc(o.durum) + '</td><td class="tags">' +
    String(o.etiketler || '').split(',').map((t) => t.trim()).filter(Boolean).map((t) => '<span class="tag">' + esc(t) + '</span>').join('') + '</td></tr>'
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
const warnHtml = (w) => w && w.length ? '<div class="warn">' + w.map(esc).join('<br>') + '</div>' : '';

async function load(keepMsg) {
  const btn = $('fetch');
  const my = mode;
  btn.disabled = true; btn.textContent = 'Çekiliyor…';
  if (!keepMsg) $('msg').innerHTML = '';
  try {
    const d = await call('api/liste?liste=bekleyen');
    if (my !== mode) return;
    session = d.sessionId; orders = d.orders; $('rows').innerHTML = ''; render();
    $('msg').innerHTML += warnHtml(d.warnings);
  } catch (e) { if (my === mode) { $('rows').innerHTML = empty(''); $('msg').innerHTML += warnHtml([e.message]); } }
  finally { btn.disabled = false; btn.textContent = 'Siparişleri çek'; }
}
$('fetch').onclick = () => load();

$('export').onclick = async () => {
  const btn = $('export');
  btn.disabled = true; btn.textContent = 'Aktarılıyor…';
  try {
    const d = await post('api/excel-aktar', { sessionId: session, keys: selected(), sirali: true });
    $('msg').innerHTML = '<span class="ok"><b>Aktarım ' + esc(d.aktarim.etiket) + '</b>: ' + d.orderCount + " sipariş Excel'e aktarıldı.</span>" +
      (d.sheetUrl ? '<br><a class="dl" href="' + esc(d.sheetUrl) + '" target="_blank" rel="noopener">↗ Google Sheet\\'te aç</a>' : '') +
      warnHtml(d.warnings);
    // Az önce aktarılan siparişlerin etiketleri: düğmeyle Label LIVE'a gönderilir
    sonEtiketler = d.etiketler || null;
    $('etiketBekleyen').classList.toggle('hide', !(sonEtiketler && sonEtiketler.length));
    if (sonEtiketler) $('etiketBekleyen').textContent = '🏷️ Aktarım ' + d.aktarim.etiket.split(' ')[0] + ' etiketlerini yazdır (' + sonEtiketler.length + ')';
    // v17: dosya inmez (kullanıcı isteği); Excel, Önceki aktarımlar › Excel indir ile alınır
    await load(true);
  } catch (e) { $('msg').innerHTML = warnHtml([e.message]); update(); }
};

$('excel').onclick = async () => {
  try {
    const d = await post('api/excel', { sessionId: session, keys: selected(), sirali: true });
    window.location.href = d.download;
  } catch (e) { $('msg').innerHTML = warnHtml([e.message]); }
};

setMode('bekleyen');
</script></body></html>`;

// ======================================================================
// SUNUCU
// ======================================================================
const SURUM = '2026-11-26 · v20 (etiketler doğrudan Argox’a, PPLA)';
const app = express();
app.get('/surum', (_req, res) => res.send(SURUM));

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
  return `indir/${id}`;   // göreli: Dilara panelinin /siparis/ altında da çalışır
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
  if (req.query.liste === 'aktarim') {
    const no = Number(req.query.no);
    if (!Number.isInteger(no) || no < 1) throw new Error('Aktarım numarası geçersiz.');
    return flow.list({ stage: 'aktarim', no });
  }
  return flow.list({ stage: 'bekleyen' });
}));

// Bir günün aktarımları (no, saat, sipariş sayısı)
app.get('/api/aktarimlar', handle(async (req) => {
  if (!isDate(req.query.tarih)) throw new Error('Tarih seçin.');
  return { aktarimlar: await flow.aktarimListesi({ tarih: req.query.tarih }) };
}));

app.post('/api/excel-aktar', handle(exclusive(async (req) => {
  const { sessionId, keys, sirali } = req.body || {};
  const out = await flow.exportExcel({ sessionId, keys, sirali: sirali === true });
  return {
    orderCount: out.orderCount, aktarim: out.aktarim, warnings: out.warnings, etiketler: out.etiketler,
    sheetUrl: out.sheet && out.sheet.url,
  };
})));

app.post('/api/excel', handle(async (req) => {
  const { sessionId, keys, sirali } = req.body || {};
  const out = await flow.download({ sessionId, keys, sirali: sirali === true });
  return { orderCount: out.orderCount, download: keep(out.buffer, out.filename) };
}));

// Etiket ayarları (ölçü + şablonlar)
app.get('/api/etiket-ayar', handle(async () => flow.etiketAyarOku()));
app.post('/api/etiket-ayar', handle(async (req) => flow.etiketAyarYaz(req.body || {})));

// Mac etiket yardımcısının kurulum/kaldırma komutları
app.get('/api/yazici-kurulum', (_req, res) => res.json(yazici.komutlar()));

// Etiket verisi (seçilenler, ekrandaki sırayla)
app.post('/api/etiket', handle(async (req) => {
  const { sessionId, keys } = req.body || {};
  return flow.etiketler({ sessionId, keys });
}));

app.get('/api/oto-durum', (_req, res) => res.json(oto.durum));

app.get('/indir/:id', (req, res) => {
  const f = files.get(req.params.id);
  if (!f) return res.status(404).send('Dosyanın süresi dolmuş, listeyi yeniden oluşturun.');
  res.set({
    'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    'Content-Disposition': `attachment; filename="${f.filename}"`,
  });
  res.send(Buffer.from(f.buffer));
});

app.get('/', (_req, res) => res.type('html').send(PAGE.replace('{{SURUM}}', SURUM)));

app.listen(config.port, () => {
  console.log(`Panel hazır: port ${config.port}`);
  oto.baslat();
});
