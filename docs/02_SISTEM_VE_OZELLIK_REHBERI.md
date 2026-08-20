# PDGM — Sistem ve Özellik Rehberi

Bu belge PDGM web uygulamasının **ne olduğunu, kimlerin nasıl kullandığını ve hangi özelliklerin bulunduğunu** anlatır. Teknik detaya girmeden, gerektiğinde dosya/route referansı verir.

---

## 1. PDGM nedir?

**PDGM**, baskı / dizgi atölyesinde iş kartlarının durumunu takip eden bir **intranet web uygulamasıdır**.

### Hangi problemi çözer?

Kaynak üretim planı genelde büyük bir Excel dosyasında tutulur. Atölyede ise kartların anlık durumu (plana alındı, dizgide, hazır, teslim) operatörler ve duvar monitörleri tarafından görülmelidir. PDGM:

1. Kaynak Excel’i güvenli şekilde içeri alır
2. Kartları web ekranlarında gösterir
3. Operatör işlemlerini kaydeder
4. Audit (kim ne yaptı) tutar
5. Yedek ve rapor sunar

### Kimler kullanır?

| Rol | Tipik kullanıcı |
|-----|-----------------|
| **admin** | Planlama / sistem sorumlusu |
| **operator** | Dizgi operatörü |
| **gozlemci** | Sadece izleyen kişi / duvar ekranı hesabı |

### Excel’in rolü

Kartların **asıl kaydı** Excel’dedir (`data/kartlar.xlsx`). Veritabanı yoktur. Web uygulaması bu dosyayı okur, bellekte işler, her önemli işlemde tekrar yazar.

### Web uygulamasının rolü

Tarayıcıdan güvenli giriş, rol bazlı ekranlar, butonlarla durum geçişi, canlı panel/monitör, yönetim ve rapor.

```text
Kaynak Excel (MAKİNE sayfası)
        │
        ▼
   PDGM Import
        │
        ▼
 Üretim kartları (kartlar.xlsx)
        │
        ├── Operatör işlemleri
        ├── Panel / Monitor
        └── Yönetim / Özet / Rapor
        │
        ▼
   Teslim + audit log
```

---

## 2. Roller ve yetkiler

| Özellik | admin | operator | gozlemci |
|---------|:-----:|:--------:|:--------:|
| Giriş / çıkış | ✓ | ✓ | ✓ |
| Panel | ✓ | ✓ | ✓ |
| Monitor | ✓ | ✓ | ✓ |
| Özet | ✓ | ✓ | ✓ |
| Operatör ekranı | ✓ | ✓ | — |
| Kart aksiyonları (dizgi, adet, hazır, teslim, not) | ✓ | ✓ | — |
| Yönetim | ✓ | — | — |
| Excel yükleme / yedek / rapor / kart düzenleme | ✓ | — | — |

Yetkisiz erişimde `yetkisiz.html` (403) veya girişe yönlendirme görülür.

---

## 3. Ana ekranlar

### `/giris` — Giriş

- Kullanıcı adı + şifre
- Çok fazla hatalı deneme → kısa süreli engel (IP bazlı)
- Başarılı girişten sonra istenen sayfaya veya varsayılan ana sayfaya gider

### `/` — Ana yönlendirme

Oturum yoksa girişe; varsa role göre panel / operatör / yönetim.

### `/panel` — Panel

- Dizgide / hazır / plana alındı kartları ve tablolar
- Arama, filtre, teslim edilenler
- ~30 sn’de otomatik yenileme (canlı görünüm)
- **Kim:** admin, operator, gozlemci

### `/operator` — Operatör

- Aktif üretim kartları
- Butonlar: Dizgiye Al, Adet Bitir, Hazıra Al, Teslim, Not
- Arama + durum filtreleri; kaydırma/filtre `sessionStorage` ile korunur
- Her aksiyondan sonra sayfa yenilenir
- **Kim:** admin, operator

### `/monitor` — Atölye monitörü

- Büyük ekran / kiosk tarzı
- Hazır / dizgide / plana listeleri, saat, kart rotasyonu
- ~30 sn yenileme
- Salt okunur (aksiyon yok)
- **Kim:** tüm roller

### `/yonetim` — Yönetim (yalnız admin)

- Excel yükleme
- Kart listesi / düzenleme / gizleme / geri getirme
- Durumu eksik kartlar
- Yedek listesi ve geri yükleme
- Yükleme geçmişi + son işlem logları
- Rapor ve Excel dosyası indirme
- Kart dosyasını yeniden okuma

### `/ozet` — Özet

- Genel sayılar (plana / dizgi / hazır / teslim / gecikme)
- Bu hafta / ay / yıl teslim KPI’ları
- Haftalık planlanan vs teslim grafiği (basit çubuklar)
- **Kim:** tüm roller

---

## 4. Kart yaşam döngüsü

```text
PLANA ALINDI
      │  Dizgiye Al
      ▼
   DİZGİDE
      │  Adet Bitir (kısmi veya tam)
      │  (tam bitince hâlâ DİZGİDE kalır)
      │  Hazıra Al  (yalnız adet tamamsa)
      ▼
    HAZIR
      │  Teslim Edildi
      ▼
 TESLİM EDİLDİ
```

### PLANA ALINDI

- İş henüz dizgiye alınmamış
- Operatör “Dizgiye Al” ile başlatır (başlangıç adedi seçilebilir)

### DİZGİDE

- Üretim devam ediyor
- “Adet Bitir”: kısmi veya kalan kadar
- Adet tam bitince sistem **otomatik HAZIR yapmaz**; ayrıca “Hazıra Al” gerekir
- Mesaj: üretim bittiğinde “Kart HAZIR'a otomatik alınmadı.”

### HAZIR

- Üretim adedi tamam, teslime hazır
- Yalnız HAZIR’dan “Teslim Edildi” yapılabilir

### TESLİM EDİLDİ

- İş kapanmış sayılır
- Gerçekleşen teslim tarihi kaydedilir

### Durumu boş / geçersiz

- Operatör ekranında **görünmez**
- Yönetim’de “DURUMU EKSİK” olarak listelenir; admin durum atayabilir
- Ham Excel değeri `Excel Durumu` alanında saklanır

### Geriye dönüş

Operatör akışında geri adım (ör. HAZIR → DİZGİDE) standart butonlarla yoktur. Admin düzenleme ekranı durum değiştirmeye izin verir (kurallara bağlı).

---

## 5. Kart alanları (kullanıcıya görünen)

| Alan | Anlamı |
|------|--------|
| Talep NO | İş / talep numarası |
| Kart Stok No | Stok kimliği |
| Talep Sahibi | Talep eden |
| Toplam Adet | Hedef üretim adedi |
| Tamamlanan Adet | Bitirilen adet |
| Kalan | Toplam − tamamlanan (ekranda hesaplanır) |
| Plan Başlangıç / Plan Teslim | Plan tarihleri |
| Gerçekleşen Teslim | Teslim tarihi |
| Durum | Workflow durumu |
| Not / Açıklama | Serbest not |
| Operatör | Son işlem yapan |
| PCB, Plan Haftası, Adet Metni | Kaynak Excel’den gelen ek bilgiler |

### Teknik alanlar (ekranda kısmen)

| Alan | Anlamı |
|------|--------|
| `anahtar` | `TalepNO\|StokNo` — import eşleştirmesi |
| `kaynak` | EXCEL / MANUEL vb. |
| `source_active` | Kaynak Excel’de hâlâ var mı |
| `admin_gizli` | Listeden gizlendi mi (silinmedi) |
| `excel_durum` | Kaynak Excel’deki ham DURUM metni |

---

## 6. Excel import sistemi

**Kim:** yalnız admin — Yönetim → Excel yükle.

### Akış (kullanıcı gözü)

1. `.xlsx` / `.xlsm` seçilir (max ~25 MB)
2. Dosya güvenli isimle `data/yuklenen_exceller/` altına kaydedilir
3. Sistem Microsoft Excel (COM) ile **değer-only snapshot** alır
4. `MAKİNE` sayfasını okur
5. Talep NO + Stok No olan satırları kart sayar
6. Mevcut kartlarla eşleştirir; yenileri ekler
7. Operatörün girdiği durum/adetleri **ezmez** (workflow korunur)
8. Çakışma varsa (ör. HAZIR kartın adedi Excel ile uyuşmuyorsa) **tüm import iptal** edilir
9. Başarıda yedek alınır; özet mesaj gösterilir

### Neden doğrudan formüllü dosya okunmuyor?

Kaynak Excel formül, gizli kolon ve dış bağlantı içerebilir. COM snapshot görünür kolonların **hesaplanmış değerlerini** alır; makro ve link güncellemesi kapalıdır.

### Kaynakta olmayan kartlar

Excel’de artık yoksa kart silinmez; `source_active = 0` olur (kaynakta yok uyarısı).

### Başarılı yükleme sonrası

Eski yüklenen dosyalardan en yeni 20’si tutulur; gerisi silinir.

---

## 7. Yönetim özellikleri

| Özellik | Ne yapar |
|---------|----------|
| Excel aktar | Yukarıdaki import |
| Kart ekle / düzenle | Manuel kart veya alan güncelleme |
| Kart gizle | Listeden çıkarır; dosyadan silmez |
| Geri getir | Gizlenen kartı gösterir |
| Durum ata | Durumu eksik kartlara durum verir |
| Yeniden oku | `kartlar.xlsx` diskten validate edilerek yüklenir |
| Rapor indir | Özet + kart + log Excel’i |
| Kartlar / Log / Yüklemeler Excel | Ham dosya indirme |
| Yedek listesi | Anlık ve günlük yedekler |
| Geri yükle | Seçilen yedekten kartları geri alır (önce koruma yedeği alınır; audit log geri sarılmaz) |

---

## 8. Operatör özellikleri

| Buton / işlem | Endpoint (özet) | Koşul |
|---------------|-----------------|--------|
| Dizgiye Al | `POST /api/basla` | Durum = PLANA ALINDI |
| Adet Bitir | `POST /api/bitir` | Durum = DİZGİDE; adet ≤ kalan |
| Hazıra Al | `POST /api/hazirla` | DİZGİDE ve tamamlanan = toplam |
| Teslim Edildi | `POST /api/teslim-et` | Durum = HAZIR |
| Not | `POST /api/not` | Görünür kart |

Filtre örnekleri: aktif, dizgide, hazır, plana, geciken vb. (ekrandaki chip’ler).

---

## 9. Monitor

- Atölye duvar ekranı için tasarlandı
- Hazır kartlar öne çıkar; dizgi ve plan listeleri destekler
- Saat + periyodik yenileme + kart carousel
- İşlem yapılmaz; oturum açıktır

---

## 10. Özet / raporlama

`/ozet` ve admin rapor indirme:

- Durum dağılımı
- Bu hafta / ay / yıl teslim sayısı, zamanında %, ortalama sapma
- Son haftalarda planlanan vs teslim
- Geciken açık kartlar

Hesaplar aktif, gizlenmemiş kartlar üzerinden yapılır (`app.py → ozet_hesapla()`).

---

## 11. Veri dosyaları

```text
data/
├── kartlar.xlsx          ← Source of truth (kartlar)
├── islem_logu.xlsx       ← Audit (kim ne yaptı)
├── yuklemeler.xlsx       ← Import geçmişi
├── kullanicilar.json     ← Kullanıcılar (hash’li şifre)
├── gizli.key             ← Session secret
├── sunucu.lock           ← Tek process kilidi
├── uygulama.log          ← Uygulama logu
├── BASLATMA_HATASI.txt   ← Açılış hatası talimatı (varsa)
├── yuklenen_exceller/    ← Yüklenen kaynak Excel’ler
└── yedekler/             ← Günlük + anlık yedekler
```

| Dosya | Rol |
|-------|-----|
| kartlar.xlsx | Ana iş verisi |
| islem_logu.xlsx | Denetim izi |
| yuklemeler.xlsx | Import özeti |
| kullanicilar.json | Kimlik |

---

## 12. Güvenlik (kullanıcı dili)

- Şifreler düz metin saklanmaz (hash)
- Her sayfada rol kontrolü
- Değiştiren işlemlerde CSRF jetonu
- Oturum çerezi HttpOnly; HTTPS açıksa Secure
- Excel yüklemede güvenli dosya adı
- Yedek/indirme yollarında dizin dışına çıkma engeli
- Excel import’ta makro ve dış link güncellemesi kapalı
- Aynı anda ikinci sunucu process’i engellenir

---

## 13. Veri güvenliği ve yedekleme

- Her kart işlemi diske yazılmadan kullanıcıya “tamam” dönmez
- Yazma: geçici dosya → diske zorla → atomik isim değiştirme → hata olursa geri al
- Günlük kart yedeği + işlem öncesi anlık yedek
- Import ve geri yükleme öncesi snapshot
- Eski anlık yedekler sınırlanır (yaklaşık son 30); günlük ~90 gün
- İsteğe bağlı: `yedek_disari_kopyala.bat` ile ağ paylaşımına gece kopya

**Geri yükleme notu:** Kartlar yedekten döner; işlem logu kasıtlı olarak geri sarılmaz (geçmiş silinmesin diye).

---

## 14. Hata ve recovery (kullanıcı ne görür?)

| Durum | Davranış |
|-------|----------|
| Excel dosyası Excel’de açık | “Excel’de açık olabilir” benzeri mesaj |
| Bozuk log / yükleme dosyası | Kenara alınır, uyarı, sistem açılır |
| Bozuk kartlar.xlsx | Açılmaz; `BASLATMA_HATASI.txt` + yedek talimatı |
| İkinci sunucu | Red; kilit mesajı |
| Import çakışması | Hiçbir değişiklik uygulanmaz; hata mesajı |
| Hatalı şifre | Genel hata metni; çok denemede bekleme |

---

## 15. Kullanıcı yönetimi (CLI)

Web’den kullanıcı ekleme yoktur (bilinçli). Sunucuda:

```text
python kullanici_yonet.py listele
python kullanici_yonet.py ekle ali operator "Ali Operatör"
python kullanici_yonet.py parola ali
python kullanici_yonet.py rol ali admin
python kullanici_yonet.py pasif ali
python kullanici_yonet.py aktif ali
```

- Parola en az 8 karakter, iki kez sorulur
- Son aktif admin pasife alınamaz / rolü düşürülemez
- Değişiklik sunucu restart istemez (dosya yeniden okunur)

İlk kurulumda `.env` içindeki `PDGM_*_PASSWORD` değerleri bir kez okunup hash’lenir; sonra `.env`’den silinmelidir.

---

## 16. Çalıştırma özeti

```text
Tarayıcı
   │
   ▼
(isteğe bağlı HTTPS / Caddy)
   │
   ▼
Waitress (Python)
   │
   ▼
Flask (app.py)
   │
   ▼
depo.py + Excel dosyaları
```

- Varsayılan port: `5001` (`PDGM_PORT`)
- Windows’ta Excel COM için uygulamayı **açık kullanıcı oturumunda** çalıştırın (Windows Service önerilmez)
- `run_pdgm.bat` yardımcı başlatıcıdır

---

## 17. Uçtan uca senaryolar

### Yeni iş sisteme geliyor

1. Admin Yönetim’den kaynak Excel’i yükler  
2. Kartlar PLANA ALINDI (veya Excel’deki durum) ile oluşur  
3. Operatör ekranında görünür  
4. Operatör Dizgiye Al → üretim

### Operatör üretim yapıyor

1. PLANA → DİZGİDE  
2. Birkaç kez kısmi Adet Bitir  
3. Son adet bitince hâlâ DİZGİDE  
4. Hazıra Al → HAZIR

### Teslim

1. HAZIR kartta Teslim Edildi  
2. Kart TESLİM EDİLDİ; özet/raporlara yansır

### Admin düzeltme

1. Durumu eksik kartı Yönetim’de görür  
2. Durum atar veya kartı düzenler  
3. Gerekirse Excel’i yeniden okur

### Yedekten dönüş

1. Yönetim → yedek seç → Geri Yükle  
2. Önce koruma yedeği alınır  
3. Kartlar eski haline döner; log geçmişi silinmez

---

## 18. Sık sorulanlar

**Şifremi unuttum?**  
Sunucuda `python kullanici_yonet.py parola <kullanici>`.

**Kart silindi mi?**  
Hayır; gizlenmiştir. Yönetim’den geri getirilebilir.

**İki bilgisayarda sunucu açılır mı?**  
Aynı `data` klasöründe hayır; process kilidi engeller.

**Mac’te Excel import?**  
Production hedefi Windows + Microsoft Excel’dir. COM olmadan import çalışmaz; diğer ekranlar geliştirme için açılabilir.

---

Bu rehberi okuyan biri PDGM’nin amacını, rollerini, ekranlarını, kart yaşam döngüsünü ve Excel’in yerini anlayabilir. Kod seviyesinde mimari için `03_DEVELOPER_HANDBOOK.md` dosyasına bakın.

---

## 19. Ekran detayları (genişletilmiş)

### Panel içeriği

Panel (`/panel`) şu blokları gösterir:

- **Sayaçlar:** Plana alındı, dizgide, hazır, teslim, gecikme adedi
- **Dizgide listesi:** Üretimdeki kartlar
- **Hazır listesi:** Teslim bekleyenler
- **Plana alındı:** En fazla 12 kart (özet görünüm)
- **Son teslim edilenler:** En fazla 6 kart (teslim zamanına göre)

Veri `app.py → _pano_verisi()` üzerinden gelir; yalnız operasyonda görünür kartlar (`depo.kartlari_getir()`).

Otomatik yenileme yaklaşık 30 saniyedir; manuel yenile butonu da vardır.

### Monitor içeriği

Monitor, panelden daha büyük tipografi ve kart odaklıdır:

- Saat / tarih alanı
- Hazır kartların vurgulu gösterimi (teslim bekleyen iş)
- Dizgide ve plana alındı listeleri
- Kartlar arasında otomatik rotasyon (carousel)
- Periyodik tam sayfa yenileme

Monitor’de buton yoktur; yanlışlıkla işlem yapılamaz. Atölyede ayrı bir gozlemci hesabıyla açık bırakmak yaygındır.

### Operatör filtreleri

Operatör ekranındaki filtre chip’leri:

| Filtre | Anlamı |
|--------|--------|
| Aktif | Plana / Dizgi / Hazır (teslim hariç tipik aktif iş) |
| Plana Alındı | Yalnız PLANA ALINDI |
| Dizgide | Yalnız DİZGİDE |
| Hazır | Yalnız HAZIR |
| Teslim Edildi | Yalnız TESLİM EDİLDİ |
| Hepsi | Tüm görünür kartlar |

Arama kutusu talep no, stok no, sahip, not gibi alanlarda metin arar (`data-arama` özniteliği).

### Yönetim ekranı bölümleri

Yönetim sayfası tek sayfada birden fazla işlev barındırır:

1. **Excel aktarım formu** — dosya seç + aktar
2. **Durumu eksik kartlar** — hızlı durum atama
3. **Kart tablosu** — arama/filtre, düzenle, gizle
4. **Gizlenen kartlar** — geri getirme
5. **Yedekler** — liste + geri yükle
6. **Yükleme geçmişi** — son importlar
7. **İşlem logu özeti** — son ~25 kayıt
8. **Dosya indirme linkleri** — kartlar / log / yüklemeler / rapor

---

## 20. Renk ve rozet dili

Kartlar durum ve plana göre renk alır (`depo.durum_bilgisi`):

| Renk sınıfı | Tipik anlam |
|-------------|-------------|
| iyi | Planında / zamanında |
| uyari | Dikkat (son gün, üretim bitti hazıra alın, durum eksik) |
| kotu | Süre aşıldı / geç teslim |
| notr | Nötr |

Örnek rozet metinleri:

- `PLANINDA (3 gün var)`
- `SÜRE AŞILDI (2 gün)`
- `ÜRETİM BİTTİ · HAZIRA ALIN`
- `HAZIR · TESLİM BEKLİYOR`
- `ZAMANINDA TESLİM` / `GEÇ TESLİM (+1 gün)`
- `DURUMU EKSİK`

Bu metinler operatörün “ne yapmalıyım?” sorusuna görsel cevap verir.

---

## 21. Kim hangi ana sayfaya düşer?

Giriş sonrası `/` yönlendirmesi:

| Rol | Varsayılan hedef |
|-----|------------------|
| admin | `/yonetim` |
| operator | `/operator` |
| gozlemci | `/panel` |

İstenirse giriş URL’sine `?devam=/monitor` gibi göreli yol eklenebilir (open redirect korumalı).

---

## 22. Kısmi üretim (partial production) ayrıntısı

Operatör “Adet Bitir” dediğinde:

1. Sistem kalan adedi hesaplar: `toplam − tamamlanan`
2. Girilen adet 1…kalan aralığında olmalıdır
3. Tamamlanan artar; durum **DİZGİDE kalır**
4. Kalan sıfırlandığında:
   - Log: `ÜRETİM ADEDİ TAMAMLANDI`
   - Mesaj: otomatik HAZIR olmadığı belirtilir
   - UI genelde “Hazıra alınsın mı?” diye sorar; operatör onaylarsa ayrıca Hazıra Al çağrılır

Bu iki adımlı tasarım, “adet bitti = fiziksel olarak hazır” varsayımını operatör onayına bağlar.

---

## 23. Admin kart düzenleme (özet)

Admin, Yönetim’den bir kartın:

- durumunu
- toplam / tamamlanan adedi
- plan tarihlerini
- gerçekleşen teslimi
- notunu

değiştirebilir. Durum PLANA ALINDI’ya çekilirse üretim sayaçları ve zaman damgaları sıfırlanır. HAZIR/TESLİM için adet kuralları sunucuda yeniden uygulanır. Bu, operatör butonlarından daha güçlü bir “düzeltme” aracıdır; yanlış kullanım audit log’a yazılır.

---

## 24. Gizleme vs silme

“Kart sil” aslında:

```text
admin_gizli = 1
```

Kart `kartlar.xlsx` içinde kalır. Operatör/panel listelerinde görünmez. Yönetim’de “gizlenenler”den geri getirilebilir. Gerçek satır silme özelliği yoktur; bu bilinçli bir veri koruma tercihidir.

---

## 25. Ağ ve erişim modeli

Tipik kurulum:

- Bir Windows PC sunucu olur
- Atölyedeki tablet/PC’ler `http://SUNUCU-IP:5001` ile bağlanır
- İsteğe bağlı HTTPS (Caddy vb.) ile parola sniffing riski azaltılır

Aynı `data` klasörünü iki sunucu process’inin açması engellenir. İkinci deneme hata mesajı alır.

---

## 26. Ne bu sistem değildir?

Anlaşılır sınırlar:

- Muhasebe / stok ERP’si değildir
- Çok şubeli bulut SaaS değildir
- Mobil native uygulama değildir (tarayıcı yeterlidir)
- Otomatik HAZIR/TESLİM robotu değildir; insan onayı vardır
- Excel’i “gösteren” bir viewer değildir; iş kurallı bir takip sistemidir

---

Bu ek bölümlerle birlikte rehber; ekran içerikleri, renk dili, partial production ve admin düzeltme sınırlarını da kapsar.
