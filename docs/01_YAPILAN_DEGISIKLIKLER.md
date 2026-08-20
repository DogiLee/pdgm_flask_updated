# PDGM Flask — Yapılan Değişiklikler

Bu belge, PDGM İş Takip Sistemi üzerinde yapılan **hardening (dayanıklılık ve güvenlik güçlendirme)** çalışmalarını özetler. Amaç: “Bu projede ne değiştirdik ve neden?” sorusuna net cevap vermektir.

**Kaynak:** Güncel repository (2026-08 itibarıyla). Git geçmişi yok; karşılaştırma orijinal `codebase.md` dump’ı ile yapılmıştır.  
**Kapsam dışı:** Templates ve `stil.css` değiştirilmemiştir. Uygulama yeniden yazılmamış, mimari aynı bırakılmıştır.

---

## 1. Genel Bakış

### Başlangıç durumu

Sistem, Flask + Waitress ile çalışan, Excel dosyalarını veritabanı gibi kullanan bir intranet uygulamasıydı. Temel mimari zaten sağlamdı:

- Atomik Excel yazımı (`temp` + `os.replace` + rollback)
- Process kilidi (`sunucu.lock`)
- CSRF, rol kontrolü, path traversal koruması
- Excel COM ile makro/link güvenliği
- Workflow korumalı Excel import

### Neden iyileştirme yapıldı?

Üretim atölyesinde elektrik kesintisi, bozuk Excel satırı, çökme sonrası kilitli kalan sunucu veya uzun süreli Excel import gibi olaylar **sistemi açılamaz veya veri kaybına açık** hale getirebiliyordu. Amaç mimariyi değiştirmek değil; bu boşlukları kapatmaktı.

### Genel hedef

> Çalışan sistemi yeniden yazmadan daha güvenli, dayanıklı ve sürdürülebilir hale getirmek.

### Mimari ve iş kuralları

| Soru | Cevap |
|------|--------|
| Mimari değiştirildi mi? | Hayır (Flask + Jinja + Excel + Waitress + tek process) |
| Veritabanına geçildi mi? | Hayır |
| Workflow değiştirildi mi? | Hayır |
| Kart kimliği / Excel şeması değişti mi? | Hayır |

---

## 2. Yapılan Değişikliklerin Özeti

### 2.1 Veri güvenliği — fsync

**Ne vardı?**  
Kart işlemi sonrası Excel geçici dosyaya yazılıyor, ardından `os.replace` ile yerine konuyordu. Dosya içeriğinin diske gerçekten “inmesi” (fsync) garanti edilmiyordu.

**Ne değişti?**  
`depo.py → _diske_zorla()` eklendi; temp workbook kapatıldıktan sonra `os.fsync` çağrılıyor.

**Neden?**  
Elektrik kesintisinde page cache’te kalan yarım Excel dosyası riski.

**Kullanıcı açısından:**  
Normal kullanımda fark edilmez; yazma birkaç milisaniye uzayabilir.

**Teknik kazanç:**  
Atomic write zincirinin son halkası tamamlandı.

---

### 2.2 Transaction yedek koruması

**Ne vardı?**  
Yazma başarısız olup geri alma (rollback) da başarısız olursa, `.txn.bak` dosyası yine de siliniyordu.

**Ne değişti?**  
Yalnız başarılı commit sonrası yedekler temizleniyor; başarısız rollback’te `.txn.bak` diskte kalıyor.

**Neden?**  
Kurtarma kopyasının yanlışlıkla silinmesini önlemek.

**Kullanıcı açısından:**  
Nadir hata durumunda manuel kurtarma şansı artar (diskte `.txn.bak` kalabilir).

---

### 2.3 Geçersiz durum (DURUM) — boot brick düzeltmesi

**Ne vardı?**  
`kartlar.xlsx` içinde tanınmayan bir DURUM değeri (ör. elle yazılmış `IPTAL`) uygulama açılışını tamamen çökertiyordu.

**Ne değişti?**  
`_durum_normalize` tanımadığı değeri `None` yapıyor. Kart Yönetim ekranında “DURUMU EKSİK” olarak görünüyor; operatör ekranında görünmüyor.

**Neden?**  
Tek yanlış hücre yüzünden tüm atölyenin sistemi kullanamaması kabul edilemezdi.

**Kullanıcı açısından:**  
Admin “Durum Ata” ile düzeltebilir; sistem açılır.

---

### 2.4 Process kilidi — stale lock recovery

**Ne vardı?**  
Uygulama çökerse veya elektrik kesilirse `data/sunucu.lock` kalıyordu; sunucu bilerek tekrar açılmayı reddediyordu (manuel silme gerekiyordu).

**Ne değişti?**

1. PID ölüyse kilit **devralınır** (uyarı basılır).
2. PID canlı ama süreç Python değilse (PID reuse) yine devralınır.
3. PID canlı ve Python ise ikinci sunucu **reddedilir**.
4. Reclaim atomik: `os.replace` ile stale dosya taşınır, sonra `O_EXCL` ile yeni kilit oluşturulur (çift process yarışı kapatıldı).

**Neden?**  
Vardiya başında “sistem açılmıyor” duruşunu azaltmak; aynı anda iki sunucunun veri bozmasını engellemek.

**Kullanıcı açısından:**  
Crash sonrası çoğu zaman elle `sunucu.lock` silmeden yeniden başlatılabilir.

---

### 2.5 Startup recovery

**Ne vardı?**  
Bozuk `islem_logu.xlsx` veya `yuklemeler.xlsx` tüm uygulamayı açılmaz yapıyordu.

**Ne değişti?**

- Log / yükleme geçmişi bozuksa: dosya `*.bozuk_Tarih` olarak kenara alınır, uyarı yazılır, boş dosya ile devam edilir.
- `kartlar.xlsx` bozuksa: Türkçe talimat + `data/BASLATMA_HATASI.txt` + uygulama kapanır (sessiz boş kartla açılmaz).

**Neden?**  
Kart verisi source of truth; audit dosyası yüzünden UI’ya hiç girememek yanlış öncelikti. Kart kaybı ise kabul edilemez.

---

### 2.6 Yedekleme — retention ve offsite

**Ne vardı?**  
Anlık/günlük yedekler sınırsız birikiyordu.

**Ne değişti?**

- Son **30** anlık yedek klasörü (`YYYYMMDD_HHMMSS_...` pattern)
- Günlük `YYYYMMDD_kartlar.xlsx` için **90 gün**
- Pattern dışı klasörler silinmez
- `yedek_disari_kopyala.bat`: ağ paylaşımına robocopy (locale-safe tarih; robocopy hata kodu Task Scheduler’a iletilir)

**Kullanıcı açısından:**  
Disk dolması riski azalır; gece offsite kopya kurulursa disk arızasına karşı ek koruma.

---

### 2.7 Excel COM / import

**Ne vardı?**  
COM referansları tam serbest bırakılmayabiliyordu; başarısız import temp dosya bırakabiliyordu; iki admin aynı anda import edince iki Excel instance açılabiliyordu.

**Ne değişti?**

- COM ref’ler `None` + `gc.collect`
- Başarısız snapshot’ta temp silinir
- `_import_kilidi` ile import serileştirilir
- Windows dışı ortamda `pywin32` yoksa uygulama import edilebilir (COM çağrısı net hata verir)

**Kullanıcı açısından:**  
Import sonrası zombi `EXCEL.EXE` birikimi azalır; çift tıklama/çift yükleme daha güvenli.

---

### 2.8 Performans (düşük risk)

**Ne değişti?**

- Yönetim log listesi: önce dilimle, sonra kopyala
- Kart işleminde log rollback: gereksiz deepcopy yerine `del` (yalnız append yolu)
- Log arşivi: önce arşiv dosyası yaz, sonra RAM kırp; hata olursa RAM geri alınır

**Kullanıcı açısından:**  
Yönetim sayfası ve buton basışları büyük log hacminde daha az “donar”.

---

### 2.9 Güvenlik / audit DoS

**Ne değişti?**

- Hatalı giriş artık Excel audit’e yazılmaz → `uygulama.log`
- Rapor ve kayıt dosyası indirme (GET) Excel mutasyonu yapmaz → uygulama logu
- Başarılı giriş hâlâ Excel audit’te

**Neden?**  
Kimlik doğrulanmamış veya GET isteklerinin tüm log workbook’unu yeniden yazması operatörleri kilitleyebiliyordu.

**Kullanıcı / güvenlik açısından:**  
Hatalı giriş kaydı Excel’de değil log dosyasında aranmalı (bilinçli politika değişimi).

---

### 2.10 Kullanıcı yönetimi

**Ne değişti?** (`kullanici_yonet.py`)

- `parola <kullanici>` — parola sıfırlama
- `rol <kullanici> <rol>` — rol değiştirme
- Son aktif admin’i pasife alma / rol düşürme engeli
- Yazmada `fsync`

**Kullanıcı açısından:**  
JSON’u elle düzenlemeden parola/rol yönetimi mümkün.

---

### 2.11 Deployment hijyeni

**Ne eklendi?**

- `.gitignore` (`.env`, `data/`, venv)
- `.env.example`
- `requirements.txt` (pin’li sürümler)
- `run_pdgm.bat`
- `PDGM_BIND` (Waitress dinleme adresi)
- `data/.gitkeep`

**Not:** Canlı `.env` içindeki bootstrap parolalarının silinmesi operasyonel adımdır; kod otomatik silmez.

---

### 2.12 UI / UX

**Değişiklik yok.**  
Operatör ekranı hâlâ aksiyon sonrası sayfa yeniliyor (`location.reload`). Bu bilinçli olarak hardening kapsamı dışında bırakıldı.

---

## 3. Önce / Sonra

### Excel yazımı

```text
ÖNCE
Kart işlemi → temp Excel → os.replace

SONRA
Kart işlemi → temp Excel → fsync → os.replace
(başarısız rollback → .txn.bak korunur)
```

### Elektrik kesintisi / crash

```text
ÖNCE
sunucu.lock kalır → uygulama açılmaz (manuel sil)

SONRA
stale PID doğrulanır → atomik reclaim → uygulama açılabilir
(canlı ikinci Python instance hâlâ reddedilir)
```

### Geçersiz DURUM hücresi

```text
ÖNCE
IPTAL vb. → uygulama hiç açılmaz

SONRA
None / DURUMU EKSİK → Yönetim’den düzeltilebilir
```

### Hatalı giriş

```text
ÖNCE
Excel islem_logu yeniden yazılır (kilit tutar)

SONRA
yalnız uygulama.log’a uyarı
```

### Import

```text
ÖNCE
Eşzamanlı COM mümkün; fail temp kalabilir; COM ref sızıntısı

SONRA
Import kilidi; temp cleanup; COM ref release
```

---

## 4. Dosya Bazında Değişiklikler

| Dosya | Önemli değişiklikler |
|-------|----------------------|
| `depo.py` | fsync, txn.bak koruma, durum normalize, stale lock (atomik reclaim), quarantine startup, retention, log perf/archive |
| `app.py` | Başlatma hatası bildirimi, JSON user cache fallback, audit DoS, upload prune (başarı sonrası), `PDGM_BIND` |
| `excel_araclari.py` | Soft COM import, cleanup, orphan temp, `_import_kilidi` |
| `kullanici_yonet.py` | `parola`, `rol`, son-admin koruması, fsync |
| Templates / CSS | **Değişmedi** |
| `.gitignore`, `.env.example`, `requirements.txt` | Yeni |
| `run_pdgm.bat`, `yedek_disari_kopyala.bat` | Yeni |

---

## 5. Özellikle Korunan Davranışlar

Koddan doğrulanmış; değiştirilmedi:

- Workflow: `PLANA ALINDI → DİZGİDE → HAZIR → TESLİM EDİLDİ`
- `kart_bitir` sonrası **otomatik HAZIR yok** (operatör onayı gerekir)
- Import’ta `plan` sonra `workflow` overlay sırası
- HAZIR/TESLİM kartlarında adet çakışmasında import’un tamamen iptali
- Kimlik: `anahtar = talep_no|stok_no`
- `KART_ALANLARI` başlık metinleri
- Kart “silme” = `admin_gizli` (gerçek delete yok)
- Global `_kilit` (RLock) ve Excel yazımının kilit içinde kalması
- `_coklu_yaz` sırası: temp → backup → replace → ters rollback
- CSRF + `@yetki` + her request’te aktif kullanıcı kontrolü
- Excel COM güvenlik bayrakları (AutomationSecurity, UpdateLinks, EnableEvents, ReadOnly, …)
- Path traversal korumaları (yedek/indirme)

---

## 6. Şu Anki Durum

| Boyut | Değerlendirme |
|-------|----------------|
| Reliability | Crash sonrası açılma ve bozuk yardımcı dosya recovery iyileşti |
| Data safety | fsync + txn.bak koruma ile yazma güvenliği arttı |
| Recoverability | Quarantine + Türkçe başlatma hatası + yedek retention |
| Security | CSRF/yetki aynı; GET mutasyon ve hatalı giriş DoS azaltıldı |
| Maintainability | CLI parola/rol; pin’li requirements; docs eklenecek |
| UX | Aynı; reload tabanlı operatör akışı duruyor |

Bu, “risk sıfır” demek değildir. Excel-as-database ölçek limiti, Windows+Excel COM bağımlılığı ve tek process modeli mimari kısıt olarak durur.

---

## 7. Hâlâ Yapılabilecek İyileştirmeler

Bilinen, bilinçli ertelenenler (zorunlu değil):

- Operatör aksiyonlarında sayfa yenileme yerine DOM güncelleme (dikkatli JS testi gerekir)
- Log yazımını debounce etmek (audit penceresi trade-off)
- Idle session timeout / yedek geri yüklemede parola tekrarı
- Formula injection korumasını `+`/`-`/`@` ile genişletme
- Proxy arkasında `ProxyFix` (yalnız TLS terminator + XFF gerekiyorsa)
- Teslim edilmiş eski kartların arşivlenmesi (ölçek için uzun vadeli)

---

## Kısa sonuç

Sistem yeniden yazılmadı. Dayanıklılık (fsync, lock reclaim, startup), operasyon (CLI, retention, deployment dosyaları) ve düşük riskli performans güçlendirildi. İş kuralları ve arayüz davranışı bilinçli olarak korunmuştur.
