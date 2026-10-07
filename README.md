# Esse Jeffe Otomasyon

Şu an çalışan modül: **Sipariş paneli.**

**Yeni siparişler sekmesi**
1. "Siparişleri getir" → kargoya verilmemiş ve henüz etiketlenmemiş siparişler tabloda görünür.
2. Listeye girmesini istemediğiniz siparişlerin işaretini kaldırın (hepsi seçili gelir).
3. "Excel oluştur" → seçilenler Excel'e alınır, e-postanıza gönderilir (kutucuk işaretliyse),
   panelden indirilebilir ve Shopify'da `etiketi çıkarıldı - otomatik` etiketi eklenir.
   Etiketli siparişler bir daha bu sekmede çıkmaz; seçmedikleriniz bir sonraki listede yine görünür.

**Geçmiş (etiketlenmiş) sekmesi**
Tarih aralığı seçip daha önce etiketlenmiş siparişleri görebilir, seçtiklerinizi tekrar
Excel'e alabilirsiniz. Bu sekme etiketlere dokunmaz.

Bir siparişin tekrar "Yeni siparişler"e düşmesini isterseniz Shopify'da etiketini silmeniz yeterli.

Şimdilik yalnızca **Shopify** açık (`KANALLAR=shopify`). Trendyol ve Hepsiburada'nın kodu
hazır; açmak için `KANALLAR=shopify,trendyol,hepsiburada` yapmanız yeterli.

Sıradaki modül: DHL teslim edilemeyen paket uyarıcısı (DHL API dokümanı bekleniyor).

## Excel'de ne var?

Her satır bir ürün. Sayfalar: **Tümü**, **Shopify**, **Trendyol**, **Hepsiburada**; bir kanal
hata verirse diğerleri yine listelenir ve hata **Uyarılar** sayfasına yazılır.

| Kanal | Hangi siparişler | Kargo anahtarı |
|---|---|---|
| Shopify | Gönderilmemiş / kısmi gönderilmiş, iptal edilmemiş | Siparişin uzun sistem ID'si |
| Trendyol | Created, Picking, Invoiced (son 14 gün) | Trendyol'un kargo takip numarası |
| Hepsiburada | Paketlenecek kalemler + paketlenip kargoya verilmemiş paketler | Paket barkodu (paketlenmemişlerde boş) |

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
