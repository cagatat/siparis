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
  Google Sheet ("Siparişler" sayfası, başlık hariç temizlenip yazılır) ve e-posta ayarlıysa yapılır, yoksa atlanır.
  Butonlar onay sormaz (kullanıcı isteği).
- **Önceki aktarımlar** (aynı sayfada): tarih aralığıyla görüntüleme + Excel indirme, durum değiştirmez.
- Eski etiketler: `drive'a aktarıldı - otomatik` / `etiketi çıkarıldı - otomatik` aktarılmış sayılır;
  `etiket oluşturuldu - otomatik` bekleyen sayılır.

Bir satır = bir sipariş; SKU'lar alt alta; siparişten çıkarılmış/değiştirilmiş kalemler
(`currentQuantity` = 0) hariç. Etiketler kolonu en sonda (bizim otomatik etiketlerimiz gizli).

## Excel / Sheet çıktısı
- Kanal adları: Shopify → `website`, Trendyol → `trendyol`, Hepsiburada → `hepsiburada`.
- Kargo adları: "Marketplace" kelimesi atılır (Trendyol Express); HepsiJET'in tüm yazımları → `HepsiJet`.
- Kolonlar: (Sheet'te başta "Aktarım") Kanal, Sipariş No, Tarih, Müşteri, Telefon, Adres, Kargo Firması,
  Kargo Anahtarı (Shopify'da siparişin uzun sistem ID'si = legacyResourceId), SKU, Adet, Tutar, Ödeme, Durum, Etiketler.
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
