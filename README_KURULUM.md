# ACEP Saha Operasyon Dashboard — GitHub üzerinden paylaşılan veri kurulumu

Bu paket, dashboard'un **herkes için aynı güncel veriyi** göstermesini sağlar.
Mantık: Excel dosyası repoya push'landığında, GitHub Actions onu otomatik
olarak `data/dashboard-data.json`'a çevirir. Dashboard sayfası bu JSON'u
her açılışta çeker. Hiçbir token veya şifre tarayıcı tarafında görünmez.

## 1) Dosyaları repoya yerleştir

Bu paketteki klasör yapısını repo'nun kök dizinine **birebir aynı yollarla**
kopyala:

```
repo-kökü/
├── acep_saha_operasyon_dashboard.html   ← GitHub Pages'ten açılan sayfa
├── saha_durum_operasyon_tablosu.xlsx    ← kaynak Excel (güncellediğinde bunu değiştir)
├── data/
│   └── dashboard-data.json              ← otomatik üretilir, ilk sürüm dahil
├── scripts/
│   └── parse_saha_durum.py              ← Excel → JSON dönüştürücü
└── .github/
    └── workflows/
        └── update-dashboard-data.yml    ← otomatik çalışan GitHub Actions
```

> Eğer dashboard'u repo'nun başka bir alt klasöründe (ör. `docs/`) tutuyorsan,
> `data/dashboard-data.json` yolunu hem HTML içindeki `SHARED_DATA_URL`
> değişkeninde hem de workflow/script çağrılarında ona göre güncelle.

## 2) Repo ayarlarında yazma izni aç (bir kerelik)

GitHub Actions'ın `data/dashboard-data.json`'ı commit'leyebilmesi için:

1. Repo → **Settings → Actions → General**
2. **Workflow permissions** bölümünde **"Read and write permissions"** seçeneğini işaretle
3. **Save**

Bu adım atlanırsa Actions çalışır ama commit adımında yetki hatası alır.

## 3) GitHub Pages'i aç (henüz açık değilse)

Repo → **Settings → Pages** → Source: `main` branch, `/ (root)` (veya
dashboard'u koyduğun klasör) → **Save**.

## 4) Veriyi güncellemek istediğinde

Artık tek yapman gereken: güncel Excel'i **aynı isimle**
(`saha_durum_operasyon_tablosu.xlsx`) repo'ya push'lamak veya GitHub web
arayüzünden "Add file → Upload files" ile üzerine yazmak.

- Push algılanır → Actions otomatik tetiklenir (bkz. **Actions** sekmesi)
- ~30-60 saniye içinde `data/dashboard-data.json` güncellenir ve commit'lenir
- Dashboard sayfasını (yeniden) açan **herkes** güncel veriyi görür

İstersen elle de tetikleyebilirsin: **Actions → "Dashboard Verisini Güncelle" →
Run workflow**.

## 5) Dashboard'daki "Yerel Önizle" butonu ne işe yarar?

Bu buton paylaşılan veriyi **değiştirmez**. Sadece o an açtığın tarayıcıda,
elindeki bir Excel dosyasını yayınlamadan önce nasıl görüneceğini kontrol
etmeni sağlar (ör. push'lamadan önce hızlı bir göz atma). Sayfayı
yenilediğinde veya "↺ Paylaşılan Veriye Dön" dediğinde tekrar gerçek/paylaşılan
veriye döner.

## Notlar

- `MMC_PLAN` (3.000) sabiti hem HTML/JS içinde hem de gerekirse ayrıca
  Python tarafında yorumla belirtilmiştir; Excel'de ayrı bir hücre olarak
  yer almadığından sabit tanımlıdır. Değişirse ilgili script/HTML'de
  güncellenmeli.
- `scripts/parse_saha_durum.py`'yi elle de çalıştırabilirsin:
  ```
  pip install openpyxl
  python scripts/parse_saha_durum.py saha_durum_operasyon_tablosu.xlsx data/dashboard-data.json
  ```
