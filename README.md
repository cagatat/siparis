# Esse Jeffe Otomasyon

Şu an çalışan modül: **Sipariş paneli.**

**Yeni siparişler sekmesi**
1. "Siparişleri getir" → kargoya verilmemiş ve henüz etiketlenmemiş siparişler tabloda görünür.
   Her satır bir sipariş; SKU'lar alt alta yazılır, siparişten çıkarılmış ya da değiştirilmiş
   ürünler listelenmez. En sağdaki "Etiketler" kolonunda siparişe yazdığınız etiketler/notlar görünür.
2. Listeye girmesini istemediğiniz siparişlerin işaretini kaldırın (hepsi seçili gelir).
3. "Excel oluştur" → seçilenler Google Sheet'e yeni bir sekme olarak yazılır, Excel olarak
   panelden indirilebilir, e-postanıza (Sheet linki + Excel eki) gönderilir ve Shopify'da
   `etiketi çıkarıldı - otomatik` etiketi eklenir. Etiketli siparişler bir daha bu sekmede çıkmaz.

**Geçmiş (etiketlenmiş) sekmesi**
Tarih aralığı seçip daha önce etiketlenmiş siparişleri görebilir, seçtiklerinizi tekrar
listeye alabilirsiniz. Bu sekme etiketlere dokunmaz.

Bir siparişin tekrar "Yeni siparişler"e düşmesini isterseniz Shopify'da etiketini silmeniz yeterli.

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

### 3. API bilgileri

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
