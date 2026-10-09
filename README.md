# Esse Jeffe Otomasyon

Panel üç aşamadan oluşur:

| Aşama | Ne var | Buton | Butonun yaptığı |
|---|---|---|---|
| 1 · Yeni gelen siparişler | Hiç dokunulmamış siparişler | **Panele çek** | Shopify: `etiket oluşturuldu - otomatik` etiketi · Trendyol: "İşleme Alındı" · Hepsiburada: paketlenip "Gönderime Hazır" |
| 2 · Panele çekilenler | 1. aşamadan çekilenler | **Drive'a aktar** (ayrıca **Adresleri düzelt**: seçilenlerin adres kontrolünü yeniden çalıştırır, aşamayı değiştirmez) | Google Sheet'teki "Siparişler" sayfasını temizleyip seçilenleri yazar, e-posta gönderir · Shopify: `drive'a aktarıldı - otomatik` etiketi · Trendyol/Hepsiburada: sistem kaydına alınır |
| 3 · Etiket oluşturulanlar | Drive'a aktarılanlar (tarih aralığıyla) | **Excel indir** | Hiçbir şeyi değiştirmez |

**Shopify adres kontrolü**: yeni gelen siparişlerde 5 dakikada bir otomatik çalışır; "Panele çek" ve 2. sekmedeki "Adresleri düzelt" sırasında da tekrar kontrol eder. Düzeltme hangi aşamada yapılırsa yapılsın siparişe `adres düzeltildi - otomatik` etiketi eklenir. Sorunlu adresler 1. ve 2. sekmede kırmızı görünür. Ödeme formundaki alanlar şöyle okunur:
"Adres" = mahalle, sokak, kapı no · "Apartman, daire vb." = ilçe · "Şehir" = il.
- "Apartman, daire vb." alanına (ilçe) dokunulmaz. Boşsa ve ilçe adres satırında ya da Şehir alanında yazıyorsa oraya taşınır;
  Şehir alanında "Adana Yüreğir" gibi fazlalık varsa sadece il adı bırakılır.
- İl ve ilçe geçerli mi, ilçe o ile mi ait, adreste yazan mahalle o ilçede gerçekten var mı kontrol edilir.
- Adreste mahalle yoksa sokak + ilçe + il ile haritada aranır; bulunan mahalle o ilçenin resmi
  listesinde varsa adresin başına "X Mah." olarak eklenir.
- Adres satırında tekrar yazılmış il ve ilçe adları sadece mahalle–ilçe–il uyumluysa silinir; uyumsuzluk varsa satıra dokunulmaz.
- Adres satırı standart sıraya dizilir: `Xxx Mah. Xxx Cad. Xxx Sok. No: 3 Kat: 2 Daire: 5 Xxx Apt. Xxx Sitesi A Blok`.
  Satırdaki her kelime bu parçalardan birine oturmuyorsa (not, tarif vb.) satır olduğu gibi bırakılır; hiçbir bilgi silinmez,
  uydurma bilgi eklenmez. Haritadan bulunan mahalle ancak sonuç aynı sokak + aynı ilçeye düşüyorsa eklenir.
- Okul, hastane, üniversite gibi kurum adreslerine dokunulmaz.
- Google Maps hata verirse (faturalandırma, kota vb.) arama otomatik olarak OpenStreetMap'e geçer.
- Değişiklik yapılan siparişlere `adres düzeltildi - otomatik` etiketi eklenir; 2. sekmede yeşil not olarak görünür.
- Mahalle bulunamayan ya da il/ilçe/mahalle uyuşmayan adresler 2. sekmede kırmızı görünür, sebebi altında yazar.
- Mahalle listesi: `turkey-neighbourhoods` paketi. Harita: `GOOGLE_MAPS_API_KEY` varsa Google Maps, yoksa OpenStreetMap.

Her sekmede siparişler tek satırdır; SKU'lar alt alta yazılır, siparişten çıkarılmış ya da
değiştirilmiş ürünler listelenmez. Adres ve Shopify etiketleri (notlarınız) da görünür.
Her sekmede "Excel indir" ile seçilenleri durum değiştirmeden Excel olarak alabilirsiniz.

Önceki sürümde `etiketi çıkarıldı - otomatik` etiketi almış siparişler 3. aşamada görünür.

Butonlar onay sormadan çalışır. Trendyol ve Hepsiburada'daki durum değişiklikleri geri
alınamaz; basmadan önce seçimleri kontrol edin.

Şimdilik yalnızca **Shopify** açık (`KANALLAR=shopify`).

## Kurulum (bir kerelik)

### 1. GitHub
1. github.com'da ücretsiz hesap açın.
2. **New repository** → adı `esse-jeffe-otomasyon`, **Private** seçin.
3. **uploading an existing file** linkine tıklayıp bu klasördeki tüm dosyaları sürükleyin
   (`.env` dosyası yoksa sorun yok; şifreler GitHub'a hiç girmeyecek).

### 2. Railway
1. railway.com'da GitHub hesabınızla giriş yapın.
2. **New Project → Deploy from GitHub repo →** `esse-jeffe-otomasyon`.
3. Proje açılınca **Variables** sekmesine `.env.example` dosyasındaki değişkenleri girin.
4. **Settings → Networking → Generate Domain** ile panel adresini alın.
5. Adrese girince tarayıcı kullanıcı adı/şifre soracak: `PANEL_USER` / `PANEL_PASSWORD`.

### 3. Kalıcı kayıt (volume)
Trendyol ve Hepsiburada'da etiket olmadığı için, Drive'a aktarılan siparişlerin kaydı
sunucuda bir dosyada tutulur. Güncellemelerde silinmemesi için:
1. Railway'de servise sağ tıklayın (ya da Ctrl/Cmd + K) → **Attach volume / Volume ekle**.
2. Mount path olarak `/data` yazın.
3. Variables'a `DATA_DIR=/data` ekleyin.
Sadece Shopify kullanırken bu kayıt kullanılmaz; yine de şimdiden açmanız önerilir.

### 4. API bilgileri

**Shopify** (Ocak 2026'dan beri yeni uygulamalar Dev Dashboard'dan oluşturuluyor)
1. Shopify admin → **Apps → Develop apps → Build apps in Dev Dashboard**.
2. Yeni uygulama oluşturun, Admin API izinlerinden `read_orders` ve `write_orders` seçin
   (`write_orders` etiket eklemek için gerekli).
3. **Protected customer data** bölümünden isim, telefon ve adres erişimini açın;
   açılmazsa müşteri adı/telefon kolonları boş gelir.
4. Uygulamayı mağazanıza kurun; **Client ID** ve **Client secret** değerlerini Railway'e girin.

**Trendyol:** Satıcı Paneli → Hesap Bilgilerim → Entegrasyon Bilgileri → Satıcı ID, API Key, API Secret.

**Hepsiburada:** Satıcı panelinden API bilgileri. Panelde yoksa Yardım Merkezi → Satıcı
Destek Talep Formu → API Entegrasyon üzerinden talep edilir. Merchant ID, API şifresi ve
size verilen entegratör kullanıcı adı (`HB_USER_AGENT`) gerekiyor.

**Google Sheets (servis hesabı):**
1. console.cloud.google.com → yeni proje oluşturun.
2. "APIs & Services → Library" → **Google Sheets API**'yi bulup **Enable**.
3. "IAM & Admin → Service Accounts" → **Create service account** (ad: esse-otomasyon), rol vermeden bitirin.
4. Hesaba girip **Keys → Add key → Create new key → JSON**. Bir .json dosyası iner.
5. Bu dosyayı metin düzenleyiciyle açıp içeriğinin tamamını `GOOGLE_SERVICE_ACCOUNT_JSON` değişkenine yapıştırın.
6. Siparişlerin yazılacağı Google Sheet'i açın → **Paylaş** → dosyadaki `client_email`
   adresini **Düzenleyen** olarak ekleyin.
7. Tablonun adresindeki `/d/` ile `/edit` arasındaki kodu `GOOGLE_SHEET_ID`'ye yazın.

**E-posta (Resend):** resend.com'da hesap açın, `essejeffe.com` alan adını doğrulayın
(DNS'e birkaç kayıt eklenir), API anahtarını `RESEND_API_KEY`'e girin.

## Klasördeki dosyalar
- `index.js` – uygulamanın tamamı (ayarlar, Shopify/Trendyol/Hepsiburada bağlantıları, Excel, e-posta, panel)
- `package.json` – Railway'in hangi kütüphaneleri kuracağını ve uygulamayı nasıl başlatacağını söyler; ayrı durması zorunlu
- `.env.example` – Railway'e girilecek ayarların listesi
- `.gitignore` – şifre dosyasının ve kütüphanelerin GitHub'a gitmesini engeller
- `README.md` – bu dosya

## Bilgisayarda deneme
```
npm install
npm start      # paneli http://localhost:3000 adresinde açar (.env gerekir)
```
