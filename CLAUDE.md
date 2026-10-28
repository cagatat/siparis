# Esse Jeffe Otomasyon — Claude Code notları

Esse Jeffe (kadın abiye/ferace e-ticaret markası) için sipariş paneli. Railway'de çalışır,
GitHub `main` dalına her push'ta Railway otomatik yeniden dağıtır. Arayüz ve mesajlar Türkçe.

## Dosya yapısı (kullanıcı tercihi — değiştirme)
- Tüm kod TEK dosyada: `index.js`. Alt klasör yok (`src/` yok), kodu modüllere bölme.
- Diğer dosyalar: `package.json`, `README.md`, `.env.example`, `.gitignore`.
- `index.js` içindeki bölümler (her biri bir IIFE): Ayarlar (config) · Yardımcılar (util) · Shopify ·
  Trendyol · Hepsiburada · Kayıt (store) · Adres kontrolü (adres) · Excel · Google Sheets ·
  E-posta (Resend) · Sipariş akışı (flow) · Otomatik adres kontrolü (oto) · Panel sayfası (PAGE) · Sunucu.
- Her değişiklikte `SURUM` sabitini artır (panel başlığında görünür, dağıtımı doğrulamak için kullanılıyor).

## Panel akışı (tek sayfa — kullanıcı 3 aşamalı sistemi istemedi, geri getirme)
- **Siparişleri çek**: Excel'e aktarılmamış siparişleri listeler, hiçbir şeyi değiştirmez.
- **Excel'e aktar**: seçilenler Excel'e yazılır ve indirilir. Shopify → tek etiket `excel`
  (`SHOPIFY_ETIKET_EXCEL`); Trendyol → `Picking`; Hepsiburada → paketle; ikisi de `DATA_DIR/aktarilanlar.json`.
  Google Sheet ("Siparişler" sayfası, başlık hariç temizlenip yazılır) ayarlıysa yazılır, yoksa atlanır.
  **E-posta gönderilmez** (v13, kullanıcı isteği; Resend modülü kaldırıldı — `MAIL_TO` yalnız adres kontrolü e-postası için okunur).
  Butonlar onay sormaz (kullanıcı isteği).
- **Sıralama (v13)**: üstte "Sırala" seçimi (SKU · ödeme yöntemi, önce kapıda · tarih yeni/eski · etiket, değişim/iade önce) ve
  tıklanabilir sütun başlıkları; seçim tarayıcıda hatırlanır (`localStorage.siralama`), varsayılan SKU. Excel ve Google Sheet
  **ekrandaki sırayla** yazılır (istemci `keys`'i ekran sırasıyla ve `sirali: true` gönderir; gelmezse eskisi gibi `sortBySku`).
- **Ödeme kolonu (ekran)**: `util.odemeYontemi` okunur ad gösterir (Kapıda / Havale/EFT / Kart / Ödendi / Değişim / İade / Pazaryeri;
  Dilara'nın taslak kapıda siparişleri `manual` + ödeme bekliyor → Kapıda; `displayFinancialStatus` okunur). Ham altyapı adı
  fare üstünde görünür. **Excel/Sheet'teki "Ödeme" kolonu değişmedi** (ham ad; etiket programı Sheet'ten okuyor).
- Geniş tablo (v13): etiketler alt alta tam görünür, tablo en az 980px.
- **Sipariş kanalı (v14, ekran)**: `util.siparisKanali` — Dilara'nın etiketlerinden WhatsApp / Instagram / Messenger / Telegram
  (yalnız "dilara" → Dilara), değilse Shopify `sourceName`: web → Web sitesi, shopify_draft_order → Elle girilen; Trendyol/Hepsiburada.
  Kanal kolonu her zaman görünür; **varsayılan sıralama kanala göre** (WhatsApp → Instagram → Messenger → Telegram → Dilara → Web sitesi
  → Elle girilen → pazaryerleri, sonra SKU). Hatırlanan seçim anahtarı `siralama2` (v13'teki bir kez sıfırlandı).
  **Excel/Sheet'teki "Kanal" kolonu değişmedi** (website/trendyol/hepsiburada; etiket programı KANAL değişkenini etikete basıyor).
- Etiketler İngilizce kelimelerse (Instagram, WhatsApp, Cash on Delivery) `toLowerCase()` kullan; `toLocaleLowerCase('tr')`
  "I"yı "ı" yapar ve eşleşme kaçar. Türkçe etiketler (İade, Ödeme bekleniyor) için `tr` doğru.
- Her "Excel'e aktar" numaralı bir **aktarımdır** (#1, #2…; no + zaman + sipariş kimlikleri). Liste Shopify'da
  uygulamanın gizli alanında (AppInstallation metafield `esse_siparis.aktarimlar`, json; Shopify kapalıysa
  `DATA_DIR/aktarimlar.json`). Liste okunamazsa aktarım hiç başlamaz. Excel adı `aktarim-N_…`, Sheet'te "Aktarım" = `#N · tarih saat`.
- **Önceki aktarımlar** (aynı sayfada): gün seçilir → o günün aktarımları (#no · saat · adet) düğme olarak → seçilenin
  siparişleri + Excel indirme; durum değiştirmez. v12'den önceki aktarımlar numarasız olduğu için listede yok.
- Eski etiketler: `drive'a aktarıldı - otomatik` / `etiketi çıkarıldı - otomatik` aktarılmış sayılır;
  `etiket oluşturuldu - otomatik` bekleyen sayılır.

Bir satır = bir sipariş; SKU'lar alt alta; siparişten çıkarılmış/değiştirilmiş kalemler
(`currentQuantity` = 0) hariç. Etiketler kolonu en sonda (bizim otomatik etiketlerimiz gizli).

- Dilara panelinde "📦 Siparişler" sekmesinde de açılır: Dilara `/siparis/*` isteklerini Basic Auth ile buraya iletir.
  Bu yüzden sayfadaki adresler GÖRELİ olmalı (`api/liste`, `indir/…`; başında `/` yok).

## Excel / Sheet çıktısı
- Kanal adları: Shopify → `website`, Trendyol → `trendyol`, Hepsiburada → `hepsiburada`.
- Kargo adları: "Marketplace" kelimesi atılır (Trendyol Express); HepsiJET'in tüm yazımları → `HepsiJet`.
- Kolonlar: (Sheet'te başta "Aktarım") Kanal, Sipariş No, Tarih, Müşteri, Telefon, Adres, Kargo Firması,
  Kargo Anahtarı (Shopify'da siparişin uzun sistem ID'si = legacyResourceId), SKU, Adet, Tutar, Ödeme, Durum, Etiketler.
- SKU sıralı: sipariş içindeki SKU'lar alfabetik (doğal sayı sırası), siparişler ilk SKU'ya göre (Excel ve Sheet).
- Uzun ID'ler ve telefonlar metin olarak yazılır (Sheet'te `valueInputOption=RAW`).

## Shopify
- Tek uygulama kullanılıyor: "Dilara" (chatbot ile ortak). Ayrı uygulama ÖNERME.
  Kimlik: client credentials (`SHOPIFY_CLIENT_ID` / `SHOPIFY_CLIENT_SECRET`, shpss_…).
- Ödeme formundaki alanlar: "Adres" (address1) = mahalle/sokak/no · "Apartman, daire vb." (address2) = İLÇE ·
  "Şehir" (city) = İL. Province alanı genelde boş.
- `orderUpdate` ile adres güncellerken soyad zorunlu; soyad boşsa tam ad ikiye bölünüyor.
- API sürümü şu an `2026-07`. Shopify 2026-10'da adres güncellemesinde vergi yeniden hesaplanıyor
  (Türkiye'de KDV adrese göre değişmediği için etkisiz). 1 Temmuz'a kadar `SHOPIFY_API_VERSION=2026-10`'a geçilmeli.

## Adres kontrolü — ŞU AN KAPALI
`ADRES_DUZELTME` varsayılan `hayir`: hiçbir yerde kontrol/düzeltme yapılmıyor (hatalı sonuçlar yüzünden
kullanıcı kapattı). Tekrar açmadan önce kullanıcının hatalı örneklerini incele.
Açıldığında kurallar (kullanıcı tarafından belirlendi, değiştirme):
- Sadece emin olunan dört parça standart yazılır ve başa bu sırayla alınır: `Xxx Mah.` → `Xxx Cad.` → `Xxx Sok.` → `Xxx Apt.`
- Geri kalan her şey (No, Kat, Daire, site adı, blok, bulvar, notlar) yazıldığı gibi, aynı sırayla kalır; kısaltma yok.
- Hiçbir bilgi silinmez, uydurulmaz. Tek ekleme: kesin bulunan mahalle (harita sonucu aynı sokak + aynı ilçe
  + resmi listede olmalı). Okul, hastane, üniversite vb. kurum adreslerine dokunulmaz.
- Adresteki tekrar il/ilçe sadece mahalle–ilçe–il uyumluysa silinir; uyumsuzsa satır olduğu gibi kalır.
- Herhangi bir düzeltmede siparişe `adres düzeltildi - otomatik` etiketi eklenir.
- Mahalle listesi: `turkey-neighbourhoods` paketi. Harita: `GOOGLE_MAPS_API_KEY` varsa Google (hata verirse
  OpenStreetMap'e düşer), yoksa OpenStreetMap (saniyede ≤1 istek).

## Durum / yapılacaklar
- Kanallar: şimdilik sadece Shopify açık (`KANALLAR=shopify`).
- Trendyol: kod hazır, gerçek hesapla denenmedi. Açmadan önce Railway volume (`/data`) + `DATA_DIR=/data`.
- Hepsiburada: kullanıcının kendi entegrasyonu için API erişimi talep etmesi gerekiyor (fatura uygulamasının
  entegratör anahtarı kullanılmaz). Alan adları ilk gerçek yanıtta doğrulanmalı.
- Google Maps: Google Cloud 1.500 TL ön ödeme istedi; şimdilik ücretsiz harita. İleride bağlanacak.
- E-posta (Resend) henüz kurulmadı.

## Çalışma şekli
- Gerçek API'lere bağlanmadan test et: `global.fetch`'i taklit eden bir mock ile sunucuyu başlatıp
  `/api/liste`, `/api/panele-cek`, `/api/drive`, `/api/excel` uçlarını çağır.
- Trendyol/Hepsiburada'da durum değiştiren işlemler geri alınamaz; bunları kullanıcı onayı olmadan
  gerçek hesapta deneme.
- Gizli bilgiler (anahtarlar, şifreler) sadece Railway Variables'ta; koda ya da depoya yazma.
