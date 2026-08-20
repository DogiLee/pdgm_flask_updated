# PDGM Flask Uygulaması — Codebase Teknik Rehberi

> **Kaynak sınırı:** Bu doküman yalnızca verilen `codebase_export.md` dump’ındaki koda dayanır.  
> Dump’ta görünmeyen deployment bileşenleri, ağ topolojisi, Task Scheduler tanımları, Caddy konfigürasyonu, gerçek Excel örnekleri veya gerçek kullanıcı verileri hakkında varsayım yapılmamıştır.  
> Bir davranış koddan doğrulanamıyorsa açıkça **“dump’ta görünmüyor”** denmiştir.

---

# 1) Sistem nedir?

## 1.1 PDGM uygulamasının amacı

Kodun kendi docstring’ine göre sistem:

- **PDGM Baskı Dizgi Atölyesi İş Takip Sistemi**
- **Flask + Excel tabanlı bir intranet uygulaması**
- Üretim kartlarının durumunu ve adet ilerlemesini takip ediyor.
- Temel workflow dört gerçek durumdan oluşuyor:

```text
PLANA ALINDI
    ↓
DİZGİDE
    ↓
HAZIR
    ↓
TESLİM EDİLDİ
```

Buradaki “kart”, bir üretim / iş takip kaydıdır. Kartın kimliği uygulamada yalnız sayısal `id` ile değil, iş açısından ayrıca:

```text
Talep NO + Kart Stok No
```

kombinasyonuyla kuruluyor.

Kodda bu iş anahtarı:

```python
anahtar = f"{talep_no}|{stok_no}"
```

şeklinde saklanıyor.

---

## 1.2 Kimler kullanıyor?

Sistemde üç rol var:

| Rol | Temel kullanım |
|---|---|
| `admin` | Tüm ekranlar, kart yönetimi, Excel import, yedek geri yükleme, rapor ve kayıt dosyaları |
| `operator` | Operatör ekranı, pano, monitör, özet; üretim workflow işlemleri |
| `gozlemci` | Pano, monitör ve özet gibi salt-okunur ekranlar |

Rol kontrolü Flask route’larında `@yetki(...)` decorator’ı ile yapılıyor.

Örneğin:

```python
@app.route("/operator")
@yetki("admin", "operator")
def operator():
    ...
```

Gözlemci bu route’a yetkili değil.

---

## 1.3 Excel’in rolü nedir?

Bu uygulamada klasik bir SQL veritabanı yok.

Kod açıkça şunu söylüyor:

> `kartlar.xlsx` uygulamanın **source of truth** dosyasıdır.

**Source of truth**, sistemde bir verinin “esas / otoritatif kaynağı” anlamına gelir.

Ana kalıcı dosyalar:

```text
data/
├── kartlar.xlsx
├── islem_logu.xlsx
├── yuklemeler.xlsx
├── kullanicilar.json
├── gizli.key
├── uygulama.log
├── sunucu.lock
├── yedekler/
└── yuklenen_exceller/
```

Bunların bazıları ilk çalıştırmada oluşur.

### `kartlar.xlsx`

Ana operasyon verisi burada saklanır.

Örneğin:

- Talep NO
- stok no
- toplam adet
- workflow durumu
- tamamlanan adet
- başlangıç / bitiş / teslim zamanları
- operatör
- not
- kaynak bilgisi
- görünürlük bayrakları

### `islem_logu.xlsx`

Audit log yani işlem geçmişidir.

Örneğin:

- kim yaptı?
- hangi rolle yaptı?
- ne yaptı?
- hangi talep / stok üzerinde yaptı?
- kaç adet?
- açıklama neydi?

### `yuklemeler.xlsx`

Excel import geçmişini saklar.

Örneğin:

- yükleyen kullanıcı
- dosya
- okunan satır
- yeni kart
- güncellenen kart
- kaynakta olmayan kart
- uyarı sayısı

---

## 1.4 Web uygulamasının rolü nedir?

Flask uygulaması Excel dosyalarının üzerine bir iş mantığı ve kullanıcı arayüzü koyuyor.

Web uygulaması:

1. Kullanıcıyı doğruluyor.
2. Rolünü kontrol ediyor.
3. Kartları `depo.py` üzerinden okuyor.
4. Kullanıcının yaptığı üretim işlemlerini validate ediyor.
5. RAM’deki state’i değiştiriyor.
6. Kart + audit log değişikliklerini Excel’e yazıyor.
7. Pano / operatör / yönetim ekranlarını Jinja template’leriyle oluşturuyor.
8. Kaynak Excel upload’ını `excel_araclari.py` üzerinden parse edip ana kayıtlarla merge ediyor.

Bu nedenle Flask burada yalnızca “Excel görüntüleyici” değildir. Asıl iş kurallarının önemli bölümü backend’de uygulanıyor.

---

# 2) Mimari harita

## 2.1 Normal uygulama akışı

```text
┌──────────────────────┐
│       Browser        │
│ HTML + JS + CSS      │
└──────────┬───────────┘
           │ HTTP
           ▼
┌──────────────────────┐
│       app.py         │
│ Flask routes         │
│ auth / role / CSRF   │
│ request/response     │
└──────────┬───────────┘
           │ Python function calls
           ▼
┌──────────────────────┐
│       depo.py        │
│ business rules       │
│ RAM state + RLock    │
│ validation           │
│ atomic-like writes   │
│ backups              │
└──────────┬───────────┘
           │ openpyxl / filesystem
           ▼
┌──────────────────────┐
│     data/*.xlsx      │
│ kartlar.xlsx         │
│ islem_logu.xlsx      │
│ yuklemeler.xlsx      │
└──────────────────────┘
```

Browser doğrudan Excel’e erişmiyor.

`app.py`, veri katmanına doğrudan dosya manipülasyonu yapmak yerine büyük ölçüde `depo.py` fonksiyonlarını çağırıyor.

---

## 2.2 Excel import akışı

Kaynak Excel importunda ayrı bir yol var:

```text
Browser
   │
   │ multipart/form-data upload
   ▼
app.py
/yonetim/yukle
   │
   ▼
excel_araclari.excelden_aktar()
   │
   ├─ import Lock
   │
   ├─ Microsoft Excel COM
   │
   ├─ MAKİNE sheet
   │
   ├─ visible columns only
   │
   ├─ values-only snapshot
   │
   ▼
geçici .xlsx snapshot
   │
   ▼
openpyxl parse
   │
   ├─ header mapping
   ├─ adet parsing
   ├─ date parsing
   ├─ durum parsing
   ├─ duplicate check
   ▼
parsed satırlar
   │
   ▼
depo.excel_import_uygula()
   │
   ├─ mevcut kartlarla merge
   ├─ workflow koruma
   ├─ validation
   ├─ import öncesi yedek
   ├─ audit log
   └─ kartlar/log/yuklemeler write
```

Önemli ayrım:

- `excel_araclari.py` **kaynağı okur ve parse eder**.
- `depo.py` **mevcut uygulama state’iyle birleştirme ve commit kararını verir**.

---

# 3) Repository haritası

Dump toplam **18 dosya** gösteriyor.

## `app.py`

Flask uygulamasının giriş noktasıdır.

Burada:

- Flask instance oluşturulur.
- `.env` okunur.
- process lock alınır.
- session secret hazırlanır.
- cookie/security ayarları yapılır.
- kullanıcı dosyası bootstrap edilir.
- logging kurulur.
- `depo.kur()` çağrılır.
- authentication / authorization / CSRF decorator’ları tanımlanır.
- tüm route ve API handler’ları bulunur.
- Waitress ile sunucu başlatılır.

`app.py`, HTTP ile business/storage katmanı arasındaki ana adaptördür.

---

## `depo.py`

Uygulamanın en kritik dosyasıdır.

Görevleri:

- RAM state tutmak
- thread synchronization
- kart doğrulama
- workflow kuralları
- kart görünüm hesapları
- Excel read/write
- rollback
- backup
- process lock
- Excel import merge
- audit log
- restore

Kodun mimari güvenlik ağı büyük ölçüde bu dosyadadır.

---

## `excel_araclari.py`

İki ana işi var:

1. Kaynak Excel’den güvenli bir **values-only snapshot** oluşturup parse etmek.
2. Rapor workbook’u üretmek.

Kaynak Excel’i doğrudan openpyxl ile okumak yerine önce Microsoft Excel COM kullanılıyor.

Snapshot alındıktan sonra parsing openpyxl ile yapılıyor.

---

## `kullanici_yonet.py`

Terminalden kullanıcı yönetmek için küçük bir CLI aracıdır.

Desteklenen işler:

- kullanıcı listeleme
- kullanıcı ekleme
- parola değiştirme
- rol değiştirme
- pasife alma
- tekrar aktif etme

Parolalar plaintext olarak `kullanicilar.json` içine yazılmaz; Werkzeug password hash kullanılır.

Ayrıca son aktif admin’in yanlışlıkla düşürülmesini veya pasife alınmasını engelleyen kontrol vardır.

---

## `templates/*`

Jinja HTML template’leridir.

### `templates/base.html`

Ortak layout:

- navbar
- rol bazlı menüler
- logout formu
- flash mesajları
- CSRF meta tag
- `pdgmFetch()`
- toast helper’ları

### `templates/giris.html`

Login formu.

`base.html`’i extend etmiyor; kendi bağımsız HTML yapısı var.

### `templates/operator.html`

Operasyonun aktif olarak değiştirildiği ana ekran.

- Dizgiye Al
- Adet Bitir
- Hazıra Al
- Teslim Edildi
- Not

işlemleri buradan API’lere gider.

### `templates/panel.html`

Salt-okunur pano görünümü.

Kartları durumlarına göre gösterir, filtreler ve 30 saniyede bir full-page reload yapar.

### `templates/monitor.html`

Atölye monitörü için büyük ekran görünümüdür.

Dump’taki mevcut template yalnız:

- `DİZGİDE`
- `PLANA ALINDI`

gruplarını render ediyor.

Workflow’ta `HAZIR` olmasına rağmen mevcut `monitor()` route’u ve `monitor.html` içinde HAZIR listesi görünmüyor.

### `templates/yonetim.html`

Admin operasyon ekranıdır.

Buradan:

- kaynak Excel upload
- kart dosyası reload
- yedek restore
- gizli kart restore
- manuel kart ekleme
- kart düzenleme
- kart gizleme
- rapor indirme
- kayıt dosyalarını indirme

yapılabiliyor.

### `templates/ozet.html`

`ozet_hesapla()` tarafından hazırlanan istatistikleri gösterir.

### `templates/yetkisiz.html`

Yetkisiz rol erişiminde 403 ekranıdır.

---

## `static/stil.css`

Uygulamanın tek CSS dosyasıdır.

Burada:

- temel renk değişkenleri
- kartlar
- durum rozetleri
- navbar
- dialog/modal
- operator layout
- panel layout
- monitor layout
- yönetim layout
- responsive kurallar

bulunuyor.

Template’lerde çok sayıda semantic class bu CSS’e bağlanıyor.

### Logo notu

`base.html` ve `giris.html` şu dosyaya referans veriyor:

```text
static/pdgm_logo.png
```

Ancak dump’ın dosya listesinde yalnız `static/stil.css` görünüyor.

Dolayısıyla `pdgm_logo.png` dosyasının gerçekten repository’de bulunup bulunmadığı **bu dump’tan doğrulanamıyor**.

---

## `run_pdgm.bat`

Windows’ta uygulamayı başlatır.

Önemli davranışlar:

- script klasörüne `cd` eder.
- `.venv\Scripts\python.exe` var mı kontrol eder.
- yoksa hata verir.
- varsa:

```bat
.venv\Scripts\python.exe app.py
```

çalıştırır.

Yorumlarda Task Scheduler için:

```text
Run only when user is logged on
```

ifadesi bulunuyor ve bunun Excel COM için zorunlu olduğu belirtiliyor.

Task Scheduler’ın gerçek konfigürasyonu dump’ta yok.

---

## `yedek_disari_kopyala.bat`

`data/` klasörünün dış bir network share’e günlük kopyasını almak için tasarlanmış.

Varsayılan hedef:

```text
\\dosyasunucu\yedek\pdgm
```

olarak yazılmış ve yorumda kendi network path’inizle değiştirilmesi isteniyor.

Her gün:

```text
HEDEF\YYYYMMDD\
```

altına `data` klasörü `robocopy` ile kopyalanıyor.

60 günden eski tarih klasörleri PowerShell ile siliniyor.

---

## `requirements.txt`

Dump’taki bağımlılıklar:

```text
Flask==3.0.3
Werkzeug==3.0.6
openpyxl==3.1.5
waitress==3.0.0
python-dotenv==1.0.1
pywin32==308; sys_platform == "win32"
```

COM entegrasyonu için Windows’ta `pywin32` kullanılıyor.

---

## `.env.example`

Desteklenen environment ayarları:

```text
PDGM_PORT
PDGM_HTTPS
PDGM_BIND
PDGM_ADMIN_PASSWORD
PDGM_OPERATOR_PASSWORD
PDGM_VIEWER_PASSWORD
```

Varsayılan port `5001`.

Varsayılan bind:

```text
0.0.0.0
```

Yorumda TLS terminator olarak Caddy örneği geçiyor fakat repository dump’ında Caddy config dosyası yok.

---

# 4) Uygulama ayağa kalkınca ne oluyor?

Startup sırasını anlamak için önemli bir detay var:

`app.py` içinde bazı işlemler **fonksiyon çağrısından önce, module import edilirken** gerçekleşiyor.

---

## 4.1 `.env` yüklenir

```python
KOK = os.path.dirname(os.path.abspath(__file__))
load_dotenv(os.path.join(KOK, ".env"))
```

Ardından:

```python
SUNUCU_PORTU = int(os.environ.get("PDGM_PORT", "5001"))
DINLENEN_ADRES = os.environ.get("PDGM_BIND", "0.0.0.0")
```

okunur.

---

## 4.2 Flask instance oluşturulur

```python
app = Flask(__name__)
```

---

## 4.3 Process lock hemen alınır

Daha route’lar tanımlanmadan:

```python
depo.process_kilidi_al()
```

çağrılır.

Amaç aynı `data/` klasörünü kullanan ikinci PDGM Python process’inin çalışmasını engellemektir.

Bu kritik bir mimari varsayımdır.

---

## 4.4 Session secret hazırlanır

`_anahtar()` fonksiyonu:

```text
data/gizli.key
```

dosyasını kullanır.

Dosya yoksa:

```python
secrets.token_hex(32)
```

ile secret üretir.

Ardından best-effort olarak:

```python
os.chmod(yol, 0o600)
```

çağrılır.

Dosyadaki anahtar 32 karakterden kısaysa startup `RuntimeError` ile durur.

Bu secret Flask session cookie imzasında kullanılır.

---

## 4.5 Flask session / upload config ayarlanır

```python
MAX_CONTENT_LENGTH = 25 * 1024 * 1024
SESSION_COOKIE_HTTPONLY = True
SESSION_COOKIE_SAMESITE = "Lax"
SESSION_COOKIE_SECURE = PDGM_HTTPS == "1"
PERMANENT_SESSION_LIFETIME = 12 saat
```

Sonuç:

- upload body sınırı 25 MB
- JavaScript session cookie’ye doğrudan erişemez
- SameSite=Lax
- `PDGM_HTTPS=1` ise Secure cookie
- login sonrası session permanent yapıldığı için 12 saatlik lifetime devreye girer

---

## 4.6 Security response header’ları eklenir

`@app.after_request` ile:

```text
X-Content-Type-Options: nosniff
X-Frame-Options: DENY
Referrer-Policy: same-origin
Permissions-Policy: camera=(), microphone=(), geolocation=()
```

ekleniyor.

Static dışındaki response’larda:

```text
Cache-Control: no-store
```

ekleniyor.

Dump’ta:

- Content-Security-Policy
- Strict-Transport-Security

header’ları görünmüyor.

---

## 4.7 Kullanıcı dosyası bootstrap edilir

Module seviyesinde:

```python
_kullanicilari_yukle()
```

çağrılıyor.

### Eğer `data/kullanicilar.json` varsa

JSON okunup kullanılıyor.

### Dosya yoksa

Üç başlangıç hesabı oluşturuluyor:

```text
admin
operator
gozlemci
```

ve şifreler environment’tan alınıyor:

```text
PDGM_ADMIN_PASSWORD
PDGM_OPERATOR_PASSWORD
PDGM_VIEWER_PASSWORD
```

Herhangi biri yoksa startup `RuntimeError` ile başarısız olur.

Parolalar:

```python
generate_password_hash(parola)
```

ile hash’lenir.

Sonuç önce:

```text
kullanicilar.json.yeni
```

dosyasına yazılır, sonra:

```python
os.replace(...)
```

ile gerçek dosyanın yerine geçirilir.

---

## 4.8 Uygulama file logging kurulur

`_gunluk_dosya_logu_kur()`:

```text
data/uygulama.log
```

için `RotatingFileHandler` kuruyor.

Ayar:

```text
maxBytes = 2,000,000
backupCount = 5
```

Bu audit Excel’den farklıdır.

- `uygulama.log`: teknik uygulama logu
- `islem_logu.xlsx`: iş/audit işlem geçmişi

---

## 4.9 `depo.kur()` çalışır

Module seviyesinde:

```python
try:
    depo.kur()
except depo.VeriDogrulamaHatasi:
    ...
```

çağrılıyor.

`depo.kur()`:

1. `data/` klasörünü oluşturur.
2. `kartlar.xlsx` varsa okur.
3. kartları normalize eder.
4. kart listesini validate eder.
5. `islem_logu.xlsx` okur.
6. `yuklemeler.xlsx` okur.
7. eksik dosyaları boş workbook olarak oluşturur.

### Ana kart dosyası bozuksa

`kartlar.xlsx` doğrulanamazsa sistem startup’ı durdurur.

`_baslatma_hatasi_bildir()`:

- anlaşılır hata mesajı üretir
- son yedekleri listeler
- `data/BASLATMA_HATASI.txt` yazar
- sonunda `SystemExit(1)` oluşur

### Audit/yükleme dosyası bozuksa

Bunlar için farklı davranış var:

`_bozuk_dosyayi_kenara_al()`

bozuk dosyayı:

```text
<dosya>.bozuk_YYYYMMDD_HHMMSS
```

adıyla kenara almaya çalışır ve boş listeyle devam eder.

Kart dosyasına aynı tolerans uygulanmaz.

---

## 4.10 `calistir()` tekrar lock ve depo kurulumunu çağırır

`if __name__ == "__main__":`

altında:

```python
calistir()
```

çağrılır.

`calistir()` tekrar:

```python
depo.process_kilidi_al()
depo.kur()
_gunluk_dosya_logu_kur()
```

çağırıyor.

Process lock aynı PID’e aitse `process_kilidi_al()` hemen return eder.

Bu yüzden aynı process içinde ikinci lock çağrısı reddedilmez.

`depo.kur()` da state’i yeniden diskten yükler.

---

## 4.11 Waitress başlatılır

Önce:

```python
from waitress import serve
```

denenir.

Başarılıysa:

```python
serve(
    app,
    host=DINLENEN_ADRES,
    port=SUNUCU_PORTU,
    threads=8,
)
```

kullanılır.

Yani dump’taki production server modeli:

```text
1 process
8 Waitress thread
```

şeklindedir.

Waitress import edilemezse fallback:

```python
app.run(
    host=...,
    port=...,
    debug=False,
    threaded=True,
)
```

çalışır.

---

# 5) Kimlik doğrulama ve yetki

# 5.1 Login

Route:

```text
GET/POST /giris
```

Handler:

```python
giris()
```

### GET

Sadece:

```python
render_template("giris.html")
```

döner.

### POST

Form alanları:

```text
kullanici
sifre
```

---

## 5.2 Login rate limit

Uygulamada basit in-memory brute-force limiti var.

```python
_GIRIS_LIMIT = 8
_GIRIS_PENCERE_SN = 5 * 60
```

Yani aynı IP için son 5 dakika içinde 8 başarısız deneme varsa yeni login:

```text
HTTP 429
```

ile reddediliyor.

State:

```python
_giris_basarisiz: dict[str, deque[float]]
```

içinde RAM’de tutuluyor.

Bu nedenle:

- process restart edilirse sayaç sıfırlanır
- çok process olsaydı process’ler arasında ortak olmazdı

Ancak mimari zaten tek process olarak tasarlanmış.

IP:

```python
request.remote_addr
```

ile alınıyor.

Reverse proxy IP handling / `ProxyFix` dump’ta görünmüyor.

---

## 5.3 Login input limitleri

Aşırı büyük girişleri hash fonksiyonuna sokmamak için:

```python
if len(kullanici) > 128 or len(sifre) > 512:
```

kontrolü var.

Bu durumda standart:

```text
Kullanıcı adı veya şifre hatalı.
```

cevabı veriliyor.

---

## 5.4 Password doğrulama

Kullanıcı kaydı:

```python
kayit = _kullanicilari_al().get(kullanici)
```

ile bulunuyor.

Şartlar:

1. kullanıcı var
2. `aktif=True`
3. `check_password_hash(...)` başarılı

ise login olur.

---

## 5.5 Başarılı login sonrası session

Önce:

```python
session.clear()
```

yapılıyor.

Ardından:

```python
session.permanent = True

session.update(
    kullanici=kullanici,
    rol=kayit["rol"],
    ad=kayit["ad"],
)
```

CSRF token da üretiliyor.

Sonra audit log’a:

```text
GİRİŞ YAPILDI
```

yazılmaya çalışılıyor.

Audit Excel yazılamazsa login iptal edilmiyor; uygulama loguna exception yazılıyor.

---

## 5.6 Safe redirect

Login URL’sinde örneğin:

```text
/giris?devam=/operator
```

olabilir.

`_guvenli_devam_hedefi()` yalnız local relative path kabul etmeye çalışıyor.

Reddedilen örnek türleri:

- scheme içeren URL
- host içeren URL
- `//...`
- backslash içeren hedef
- `/` ile başlamayan hedef

Amaç open redirect riskini azaltmak.

---

# 5.7 Logout

Route:

```text
POST /cikis
```

Decorator:

```python
@yetki("admin", "operator", "gozlemci")
@csrf_koru
```

Önce audit log’a:

```text
ÇIKIŞ YAPILDI
```

yazılmaya çalışılıyor.

Her durumda:

```python
finally:
    session.clear()
```

çalışıyor.

Yani audit log hatası logout’u engellemiyor.

---

# 5.8 Her request’te kullanıcı kontrolü

`@app.before_request`:

```python
_oturum_kullanici_kontrol()
```

mevcut session kullanıcısını `kullanicilar.json` ile yeniden kontrol ediyor.

Kullanıcı:

- dosyada yoksa
- veya pasifse

session temizleniyor.

API request ise:

```text
401
{"hata": "Oturum sonlandırıldı. Tekrar giriş yapın."}
```

normal sayfa ise login’e redirect yapılıyor.

Ayrıca:

```python
session["rol"] = ...
session["ad"] = ...
```

her request’te güncelleniyor.

Bu sayede kullanıcı dosyasındaki rol/ad değişiklikleri restart olmadan etkili olabilir.

---

# 5.9 Kullanıcı dosyası cache’i

`_kullanicilari_al()` JSON dosyasının `mtime` değerini izliyor.

Dosya değişmişse yeniden okuyor.

Eğer yeni JSON bozuksa ve RAM’de daha önce okunmuş iyi bir cache varsa:

- eski iyi cache kullanılmaya devam ediyor
- uygulama loguna hata yazılıyor

İlk kez okunuyorsa ve iyi cache yoksa hata yukarı taşınabilir.

---

# 5.10 Yetki decorator’ı

```python
def yetki(*roller):
```

temel route authorization mekanizmasıdır.

### Session yoksa

```python
redirect(url_for("giris", devam=request.path))
```

### Rol uygun değilse

```python
render_template("yetkisiz.html"), 403
```

### Rol uygunsa

gerçek handler çağrılır.

### İnce nokta

API endpoint’leri de aynı decorator’ı kullanıyor.

Bu nedenle tamamen session’sız bir API isteğinde decorator JSON 401 yerine login redirect’i üretebilir.

Bu davranış koddan görülebiliyor.

---

# 5.11 CSRF

**CSRF (Cross-Site Request Forgery)**, kullanıcının browser session’ını kötüye kullanarak istem dışı mutation request’i yaptırma saldırısıdır.

Token:

```python
_csrf_token_uret()
```

ile session’da tutulur.

Üretim:

```python
secrets.token_urlsafe(32)
```

### Browser’a nasıl gider?

`base.html`:

```html
<meta name="csrf-token" content="{{ csrf_token }}">
```

Normal POST form’larında:

```html
<input type="hidden" name="_csrf_token" value="{{ csrf_token }}">
```

### Backend nasıl kontrol eder?

`csrf_koru`:

```python
request.headers.get("X-CSRF-Token")
```

veya:

```python
request.form.get("_csrf_token")
```

alır.

Ardından:

```python
secrets.compare_digest(beklenen, gelen)
```

kullanır.

Başarısız API:

```text
403
{"hata": "Geçersiz veya eksik CSRF token."}
```

Başarısız normal form:

- flash mesajı
- referrer veya ana sayfaya redirect

### Login CSRF

`/giris` POST route’unda `@csrf_koru` decorator’ı yoktur.

Bu nedenle login formu CSRF token istemiyor.

Bu, dump’taki mevcut davranıştır.

---

# 5.12 Parola hash’i

Başlangıç kullanıcıları da CLI ile eklenen kullanıcılar da:

```python
generate_password_hash(...)
```

ile saklanır.

Login:

```python
check_password_hash(...)
```

kullanır.

`kullanicilar.json` formatı yaklaşık olarak:

```json
{
  "kullanici_adi": {
    "sifre_hash": "...",
    "rol": "operator",
    "ad": "Görünen Ad",
    "aktif": true
  }
}
```

Gerçek hash formatı / algoritma seçimi Werkzeug’e bırakılmış.

---

# 5.13 `kullanici_yonet.py` CLI

Kullanım mesajındaki komutlar:

```bash
python kullanici_yonet.py listele

python kullanici_yonet.py ekle <kullanici> <admin|operator|gozlemci> <Görünen Ad>

python kullanici_yonet.py parola <kullanici>

python kullanici_yonet.py rol <kullanici> <admin|operator|gozlemci>

python kullanici_yonet.py pasif <kullanici>

python kullanici_yonet.py aktif <kullanici>
```

### CLI parola kuralı

CLI üzerinden yeni parola en az:

```text
8 karakter
```

olmalı.

Başlangıç `.env` parolaları için `_kullanicilari_yukle()` içinde aynı minimum uzunluk kontrolü görünmüyor.

---

## Son aktif admin koruması

CLI:

- son aktif admin’in rolünü düşürmeyi
- son aktif admin’i pasife almayı

reddediyor.

Amaç sistemin admin’siz kalmasını engellemek.

---

# 6) Kart yaşam döngüsü

Ana workflow:

```text
PLANA ALINDI
    ↓
DİZGİDE
    ↓
HAZIR
    ↓
TESLİM EDİLDİ
```

Normal operator akışı bu sırayı zorunlu tutuyor.

Admin düzenleme API’si ise kontrollü şekilde state’i doğrudan düzeltebiliyor.

---

# 6.1 PLANA ALINDI → DİZGİDE

## UI butonu

`operator.html`:

```text
Dizgiye Al
```

yalnız:

```python
k.durum == "PLANA ALINDI"
```

ise gösteriliyor.

---

## API

```text
POST /api/basla
```

Handler:

```python
api_basla()
```

Çağırdığı depo fonksiyonu:

```python
depo.kart_baslat(...)
```

---

## Girdiler

JSON:

```json
{
  "kart_id": 12,
  "adet": 100,
  "not": "Opsiyonel açıklama"
}
```

---

## Önkoşullar

`kart_baslat()` kontrol ediyor:

1. kart var mı?
2. operasyon ekranında görünür mü?
3. durum gerçekten `PLANA ALINDI` mı?
4. adet sayıya çevrilebilir mi?
5. adet:

```text
1 <= adet <= toplam_adet
```

mi?

`adet` `None` veya boşsa backend toplam adedi default kabul ediyor.

UI ise başlangıç dialogunu `kalan_adet` ile dolduruyor.

---

## Değişen alanlar

Kart:

```python
durum = DIZGIDE
baslangic_adet = adet
baslama_zamani = mevcutsa koru, yoksa simdi()
bitis_zamani = None
teslim_zamani = None
gerceklesen_teslim = None
operator = kullanici
aciklama = yeni not varsa yeni not, yoksa eski not
guncelleme = simdi()
```

---

## Audit log

```text
DİZGİYE ALINDI
```

Adet ve açıklama kaydedilir.

---

# 6.2 DİZGİDE → kısmi üretim

## UI butonu

```text
Adet Bitir
```

şu durumda görünür:

```python
k.durum == "DİZGİDE" and k.kalan_adet > 0
```

---

## API

```text
POST /api/bitir
```

Handler:

```python
api_bitir()
```

Depo:

```python
depo.kart_bitir(...)
```

---

## Önkoşullar

1. kart var
2. durum `DİZGİDE`
3. kalan adet > 0
4. girilen adet integer
5. adet:

```text
1 <= adet <= kalan
```

---

## Kısmi üretimde değişen alanlar

```python
tamamlanan_adet += adet
operator = kullanici
aciklama = ...
guncelleme = simdi()
```

Üretim tamamen bitmediyse:

```python
bitis_zamani
```

eski haliyle kalır.

---

## Audit log

Kısmi ise:

```text
KISMİ ÜRETİM
```

Detay default:

```text
yeni_toplam/toplam_adet adet tamamlandı
```

---

## API response

```json
{
  "tamam": true,
  "uretim_bitti": false,
  "mesaj": "60/100 adet tamamlandı.",
  "kart": { "...": "..." }
}
```

---

# 6.3 Son adet tamamlanınca ne olur?

En kritik davranışlardan biri budur.

`kart_bitir()` son adedi de tamamladığında:

```python
uretim_bitti = True
```

olur.

Kartta:

```python
tamamlanan_adet = toplam_adet
bitis_zamani = simdi()
```

yazılır.

Audit:

```text
ÜRETİM ADEDİ TAMAMLANDI
```

olur.

Ancak:

```python
durum
```

**DİZGİDE olarak kalır.**

Backend mesajı açıkça:

```text
Üretim adedi tamamlandı. Kart HAZIR'a otomatik alınmadı.
```

der.

## Sonuç

### `kart_bitir` sonrası otomatik HAZIR olur mu?

**Hayır.**

---

# 6.4 UI neden bazen otomatik gibi görünebilir?

Operator JavaScript’i `/api/bitir` response’unda:

```javascript
if (data.uretim_bitti) {
    const hazir = window.confirm(...)
}
```

yapıyor.

Kullanıcı onay verirse frontend ikinci bir API çağrısı daha yapıyor:

```text
POST /api/hazirla
```

Bu nedenle kullanıcı açısından tek akış gibi görünebilir ama backend’de iki ayrı transaction-benzeri mutation vardır:

```text
/api/bitir
      ↓
kart_bitir()
      ↓
DİZGİDE + tamamlanan=toplam

Kullanıcı onayı

/api/hazirla
      ↓
kart_hazirla()
      ↓
HAZIR
```

Kullanıcı “Hayır” derse kart:

```text
DİZGİDE
```

kalır.

UI daha sonra:

```text
Hazıra Al
```

butonunu gösterir.

---

# 6.5 DİZGİDE → HAZIR

## UI butonu

```text
Hazıra Al
```

şu durumda görünür:

```python
k.durum == "DİZGİDE" and k.kalan_adet == 0
```

---

## API

```text
POST /api/hazirla
```

Depo:

```python
kart_hazirla()
```

---

## Önkoşullar

1. kart var
2. durum `DİZGİDE`
3. `tamamlanan_adet == toplam_adet`

aksi halde:

```text
Kart HAZIR yapılmadan önce üretim adedinin tamamı bitirilmelidir.
```

---

## Değişen alanlar

```python
durum = HAZIR
bitis_zamani = mevcutsa koru, yoksa simdi()
operator = kullanici
aciklama = ...
guncelleme = simdi()
```

---

## Audit

```text
HAZIR OLARAK İŞARETLENDİ
```

---

# 6.6 HAZIR → TESLİM EDİLDİ

## UI butonu

```text
Teslim Edildi
```

yalnız `HAZIR` kartta görünür.

---

## API

```text
POST /api/teslim-et
```

Depo:

```python
kart_teslim_et()
```

---

## Önkoşullar

1. kart var
2. durum `HAZIR`

---

## Değişen alanlar

```python
durum = TESLIM_EDILDI
gerceklesen_teslim = bugun()
teslim_zamani = simdi()
operator = kullanici
aciklama = ...
guncelleme = teslim_ani
```

---

## Audit

```text
TESLİM EDİLDİ
```

Default detay:

```text
Teslim tarihi: YYYY-MM-DD
```

---

# 6.7 Not güncelleme

Route:

```text
POST /api/not
```

Depo:

```python
kart_not_guncelle()
```

Kart durumunu değiştirmez.

Değişenler:

```text
aciklama
guncelleme
```

Audit:

```text
NOT GÜNCELLENDİ
```

Not boşaltılmışsa:

```text
Not temizlendi
```

detayı yazılır.

---

# 7) Veri modeli

# 7.1 `kartlar.xlsx` içinde saklanan alanlar

`KART_ALANLARI` doğrudan workbook kolonlarını tanımlar.

| Excel başlığı | Python alanı | Anlam |
|---|---|---|
| ID | `id` | Dahili sayısal kart kimliği |
| Sıra | `sira` | Kaynak / plan sıra bilgisi |
| Talep NO | `talep_no` | Talep numarası |
| Kart Stok No | `stok_no` | Kart / malzeme stok numarası |
| Talep Sahibi | `talep_sahibi` | Talep sahibi |
| Toplam Adet | `toplam_adet` | Üretilecek toplam adet |
| Adet Metni | `adet_metin` | Kaynak Excel’deki adet metni |
| Plan Haftası | `plan_hafta` | Plan haftası / kaynak plan alanı |
| Plan Başlangıç | `plan_baslama` | Planlanan başlangıç tarihi |
| Plan Teslim | `plan_teslim` | Planlanan teslim tarihi |
| Gerçekleşen Teslim | `gerceklesen_teslim` | Gerçek teslim tarihi |
| Excel Durumu | `excel_durum` | Kaynak Excel’de gelen ham durum |
| PCB | `pcb` | PCB bilgisi |
| Durum | `durum` | Uygulamanın gerçek workflow durumu |
| Başlangıç Adedi | `baslangic_adet` | Dizgiye alınırken kaydedilen başlangıç adedi |
| Tamamlanan Adet | `tamamlanan_adet` | Üretimi tamamlanmış adet |
| Başlama Zamanı | `baslama_zamani` | Gerçek işlem başlangıç timestamp’i |
| Üretim Bitiş Zamanı | `bitis_zamani` | Üretim bitiş timestamp’i |
| Teslim Zamanı | `teslim_zamani` | Teslim işlem timestamp’i |
| Operatör | `operator` | Son işlem yapan kullanıcı |
| Not | `aciklama` | Kart notu |
| Son Güncelleme | `guncelleme` | Son update timestamp’i |
| Listede | `aktif` | Genel aktiflik bayrağı |
| Kaynakta Aktif | `source_active` | Son kaynak Excel’de bulunup bulunmadığı |
| Admin Gizli | `admin_gizli` | Soft-delete / gizleme bayrağı |
| Kaynak | `kaynak` | `EXCEL` veya `MANUEL` |
| Anahtar | `anahtar` | `talep_no|stok_no` iş anahtarı |

---

# 7.2 Zorunlu alanlar

Diskteki ana kart dosyası okunurken zorunlu kolonlar:

```python
ZORUNLU_KART_ALANLARI = {
    "id",
    "talep_no",
    "stok_no",
    "toplam_adet",
    "tamamlanan_adet",
    "anahtar",
}
```

Bu kolonlar yoksa `VeriDogrulamaHatasi` oluşur.

---

# 7.3 Sayısal alanlar

`SAYISAL_ALANLAR` içinde:

- `id`
- `sira`
- `toplam_adet`
- `baslangic_adet`
- `tamamlanan_adet`
- `aktif`
- `source_active`
- `admin_gizli`

ve log/import sayaçları yer alıyor.

---

# 7.4 Ekranda hesaplanan alanlar

Bunlar `kartlar.xlsx` kolonları değildir.

`kart_gorunumu()` ve `durum_bilgisi()` üretir.

### `rozet`

UI’da gösterilecek açıklayıcı durum.

Örnek:

```text
PLANINDA (3 gün var)
SON 1 GÜN
SÜRE AŞILDI (2 gün)
ÜRETİM BİTTİ · HAZIRA ALIN
ZAMANINDA TESLİM
GEÇ TESLİM (+4 gün)
```

### `renk`

UI durum sınıfı:

```text
notr
uyari
iyi
kotu
```

### `sapma`

Plan ile gerçekleşen / mevcut gecikme arasındaki gün farkı.

### `kalan`

`DİZGİDE` kartta teslim planına kalan gün.

Bu alan **`kalan_adet` ile aynı değildir**.

### `zaman_yuzde`

Plan süresine göre hesaplanan progress metriği.

Üretim adedi yüzdesi değildir.

### `plan_gun`

Plan teslim - plan başlangıç gün farkı.

### `kalan_adet`

```python
max(
    0,
    toplam_adet - tamamlanan_adet
)
```

### `adet_yuzde`

```python
min(
    100,
    round(tamamlanan_adet / toplam_adet * 100)
)
```

### `gorunur`

`_operasyonda_gorunur_mu()` sonucu.

### `kaynakta_yok`

```python
source_active != 1
```

### `is_durumu`

Durum varsa durum, yoksa:

```text
DURUMU EKSİK
```

### `kaynak_durumu`

Kaynak Excel’in ham `excel_durum` alanının temizlenmiş hali.

---

# 7.5 Kartın iş anahtarı

İş anahtarı:

```python
f"{talep_no}|{stok_no}"
```

Örnek:

```text
TLP-2026-001|ABC-12345
```

Import sırasında duplicate detection bu anahtar üzerinden yapılıyor.

Manuel kart eklemede de aynı anahtar zaten varsa kart oluşturulması reddediliyor.

---

# 7.6 `id` ile `anahtar` farkı

### `id`

Uygulama içi sayısal kimlik.

Yeni ID:

```python
max(mevcut id) + 1
```

şeklinde oluşturulur.

### `anahtar`

İş kimliğidir:

```text
Talep NO + Kart Stok No
```

Kaynak Excel’den aynı işi sonraki importta bulmak için asıl merge anahtarı budur.

---

# 7.7 `source_active` önemli ayrımı

Importta bir Excel kaynak kartı yeni dosyada görünmezse:

```python
source_active = 0
```

yapılıyor.

Ancak `_operasyonda_gorunur_mu()` şu alanları kontrol ediyor:

```text
aktif == 1
admin_gizli != 1
durum geçerli workflow durumlarından biri
```

`source_active` burada şart değildir.

Dolayısıyla kaynak Excel’de artık görünmeyen fakat sistemde açık workflow’u olan kart tamamen silinmez ve sadece bu bayrak nedeniyle operasyon ekranından otomatik düşmez.

UI’da:

```text
Kaynak Excel'de yok
```

uyarısı gösterilebilir.

---

# 8) `depo.py` çekirdek mekanizmalar

Bu bölüm sistemin en kritik teknik bölümüdür.

---

# 8.1 RAM state + `RLock`

Global state:

```python
_kilit = threading.RLock()

_kartlar = []
_loglar = []
_yuklemeler = []
```

## `RLock` nedir?

`RLock`, **reentrant lock** yani aynı thread’in aynı lock’u tekrar almasına izin veren mutex türüdür.

Bu kodda bazı üst seviye fonksiyonlar lock içindeyken alt yardımcılar da lock kullanabildiği için `RLock` normal `Lock`’tan daha uygundur.

---

## Amaç

Aynı process içindeki birden fazla Waitress thread’in:

- aynı anda kart state’ini değiştirmesini
- aynı Excel dosyasına yarışarak yazmasını
- birbirinin in-memory mutation’ını bozmasını

engellemeye çalışır.

---

## Ne zaman kullanılır?

Neredeyse tüm kritik depo fonksiyonlarında:

```python
with _kilit:
```

vardır.

Örnek:

- `kur()`
- `kartlari_getir()`
- `kart_baslat()`
- `kart_bitir()`
- `kart_hazirla()`
- `kart_teslim_et()`
- `admin_kart_duzenle()`
- `excel_import_uygula()`
- `yedekten_geri_yukle()`

---

## Neyi garanti etmeye çalışır?

**Tek Python process içinde** read-modify-write bütünlüğü sağlamaya çalışır.

Ancak bu lock process dışındaki bir Python instance’ını durduramaz.

Bu nedenle ayrıca `sunucu.lock` process lock sistemi vardır.

---

## Bozulursa ne olur?

Thread synchronization kaldırılırsa iki request aynı RAM state’i aynı anda değiştirip:

- lost update
- yanlış adet
- log / kart uyumsuzluğu
- aynı anda Excel write

gibi problemler yaratabilir.

---

## Değiştirirken dikkat

Bu uygulama:

```text
tek process + çok thread
```

varsayımına göre tasarlanmıştır.

Multi-process deployment’a geçerken yalnız `RLock` yeterli olmaz.

---

# 8.2 Normalize mekanizması

## `_durum_normalize(deger)`

### Amaç

Durum text’ini dört canonical workflow durumundan birine çevirmek.

### Ne yapar?

Türkçe karakterleri sadeleştirip uppercase eder.

Örnek:

```text
DİZGİDE
DIZGIDE
```

aynı canonical değere dönüşür:

```python
DIZGIDE
```

sabitinin değeri:

```text
DİZGİDE
```

olur.

Boş input:

```python
None
```

döner.

Geçersiz text de:

```python
None
```

döner.

### Değiştirirken dikkat

Geçersiz durumun `None` olması sistemde bilinçli bir davranıştır.

Bu kartlar kaybolmaz; Yönetim ekranında “Durumu Eksik” olarak görülebilir.

---

## `tarih_coz(deger)`

### Amaç

Farklı tarih biçimlerini:

```text
YYYY-MM-DD
```

formatına normalize etmek.

### Kabul ettiği kaynak tipleri

- Python `datetime`
- Python `date`
- Excel serial number
- çeşitli string formatları

Örnek formatlar:

```text
YYYY-MM-DD HH:MM:SS
DD.MM.YYYY HH:MM:SS
DD/MM/YYYY
DD-MM-YYYY
YYYY/MM/DD
```

Geçersizse:

```python
None
```

döner.

---

## `_kart_normalize(kart)`

### Amaç

Diskten / başka kaynaktan gelen kart dict’ini canonical internal forma getirmek.

### Yaptıkları

- ID integer
- sıra integer / `None`
- toplam adet minimum fallback
- tamamlanan / başlangıç adet integer
- bayraklar 0/1
- `kaynak` default `EXCEL`
- durum normalize
- tarih alanları normalize

Geçersiz non-empty tarih varsa:

```python
VeriDogrulamaHatasi
```

atar.

---

# 8.3 Kart validation

## `_kart_dogrula(kart)`

### Amaç

Tek kartın iş ve veri invariant’larını kontrol etmek.

### Kontroller

- ID pozitif
- Talep NO dolu
- Stok No dolu
- toplam adet >= 1
- `0 <= tamamlanan <= toplam`
- durum ya `None` ya dört geçerli durumdan biri
- `PLANA ALINDI` veya durumu boş kartta tamamlanan adet 0
- `HAZIR` ve `TESLİM EDİLDİ` kartta tamamlanan == toplam
- `TESLİM EDİLDİ` ise gerçekleşen teslim tarihi zorunlu
- plan başlangıç <= plan teslim
- `anahtar` dolu

---

## `_kart_listesi_dogrula(kartlar)`

Tek tek `_kart_dogrula()` çağırır.

Ek olarak:

- duplicate ID
- duplicate anahtar

kontrol eder.

---

## Neyi garanti etmeye çalışır?

Excel’e yazılacak in-memory state’in sistemin temel iş kurallarını ihlal etmemesini sağlar.

---

## Bozulursa ne olur?

Bu validation katmanı gevşetilirse örneğin:

- HAZIR fakat üretimi yarım kart
- aynı iş anahtarına iki kart
- teslim edilmiş fakat teslim tarihi olmayan kart

oluşabilir.

Bu da UI ve rapor hesaplarının varsayımlarını bozar.

---

# 8.4 `_atomik_kart_islemi`

Bu fonksiyon workflow mutation’larının merkezidir.

```python
def _atomik_kart_islemi(islem):
```

## Amaç

RAM state mutation + validation + Excel commit sürecinde hata olursa RAM’i eski haline döndürmek.

---

## Ne zaman çağrılır?

Örneğin:

- `kart_baslat`
- `kart_bitir`
- `kart_hazirla`
- `kart_teslim_et`
- `kart_not_guncelle`
- `admin_kart_ekle`
- `admin_kart_duzenle`
- `admin_kart_gizle`
- `admin_kart_geri_getir`

---

## Ne yapar?

### 1. Kart state’inin derin kopyasını alır

```python
eski_kartlar = copy.deepcopy(_kartlar)
```

Deepcopy özellikle önemli çünkü kart dict’leri in-place mutate ediliyor.

### 2. Log uzunluğunu kaydeder

```python
log_sayisi = len(_loglar)
```

### 3. Mutation callback’ini çalıştırır

```python
sonuc = islem()
```

### 4. Tüm kart listesini doğrular

```python
_kart_listesi_dogrula(_kartlar)
```

### 5. Kart + log commit eder

```python
_kart_log_commit()
```

### 6. Başarılıysa callback sonucunu döner

---

## Hata olursa

```python
_kartlar = eski_kartlar
del _loglar[log_sayisi:]
raise
```

Yani:

- kart RAM state eski hale döner
- işlem sırasında append edilen yeni audit kayıtları silinir
- exception yukarı çıkar

---

## Neyi garanti etmeye çalışır?

Bir iş işlemi:

```text
ya tamamen uygulanmış
ya da RAM açısından uygulanmamış
```

gibi davranmaya çalışır.

Bu gerçek database transaction değildir.

---

## Bozulursa / hata olursa ne olur?

Örneğin Excel write hata verirse:

- in-memory kart değişikliği rollback edilir
- yeni audit append’i rollback edilir
- kullanıcıya hata döner

Disk tarafındaki rollback ayrıca `_coklu_yaz()` tarafından denenir.

---

## Değiştirirken dikkat

Yeni mutation fonksiyonu yazıldığında doğrudan:

```python
_kartlar
```

değiştirip `_yaz()` çağırmak yerine mevcut transaction-benzeri pattern korunmalıdır.

---

# 8.5 `_kart_log_commit()`

```python
def _kart_log_commit():
```

### Amaç

Kart ve audit log’u birlikte persist etmeye çalışmak.

### Adımlar

1. Günlük kart yedeği:
   ```python
   _gunluk_yedek(KARTLAR_DOSYA)
   ```

2. Çoklu write:
   ```python
   _coklu_yaz([
       kartlar.xlsx,
       islem_logu.xlsx,
   ])
   ```

### Dikkat

Bu iki dosya aynı gerçek ACID transaction içinde değildir.

`_coklu_yaz()` bunun için rollback mekanizması kurar.

---

# 8.6 `_temp_yaz()`

## Amaç

Hedef Excel’i doğrudan overwrite etmek yerine önce tam yeni workbook üretmek.

### Adımlar

1. Random temp adı:
   ```text
   hedef.xlsx.<uuid>.yeni
   ```
2. workbook oluştur
3. temp’e save
4. workbook close
5. `_diske_zorla(temp)`

Başarıdan sonra temp path döner.

---

# 8.7 `_diske_zorla()`

## Amaç

Temp dosyanın yalnız OS page cache’te kalma ihtimalini azaltmak.

### Ne yapar?

Dosyayı read-only fd ile açar:

```python
os.open(...)
```

sonra:

```python
os.fsync(fd)
```

çağırır.

### Hata olursa

`OSError` çoğu durumda yutulur.

Yani fsync başarısı mutlak zorunlu bir invariant olarak ele alınmıyor.

---

# 8.8 `_coklu_yaz()`

Bu dosya güvenliği açısından en önemli fonksiyonlardan biridir.

```python
def _coklu_yaz(dosyalar):
```

## Amaç

Bir veya daha fazla Excel dosyasını:

```text
temp → backup → replace → rollback
```

yaklaşımıyla yazmak.

---

## Ne zaman çağrılır?

- kart + audit commit
- importta kart + audit + yükleme geçmişi
- restore
- tek dosya yazan `_yaz()` wrapper’ı

---

## Adım 1 — Tüm yeni workbook’ları temp’e hazırla

Her hedef için:

```python
_temp_yaz(...)
```

çalışır.

Henüz gerçek hedef değiştirilmez.

Bu önemli çünkü workbook üretimi sırasında hata varsa mevcut dosyalar korunur.

---

## Adım 2 — Mevcut hedeflerden transaction backup oluştur

Her mevcut hedef için:

```text
<hedef>.<uuid>.txn.bak
```

oluşturulur.

Bu backup’lar normal kullanıcı yedeklerinden farklıdır.

---

## Adım 3 — Temp dosyaları gerçek hedeflere sırayla geçir

```python
os.replace(temp, hedef)
```

kullanılır.

Aynı filesystem içinde `os.replace` hedef file replacement için güçlü bir primitive’dir.

---

## Adım 4 — Bir replace başarısız olursa rollback

Daha önce değiştirilmiş hedefler ters sırada dolaşılır.

Eski backup varsa:

```python
os.replace(backup, hedef)
```

ile geri konur.

Hedef önceden yoksa ve yeni oluşmuşsa:

```python
os.remove(hedef)
```

denenir.

Sonra original exception yeniden fırlatılır.

---

## Adım 5 — Cleanup

Başarısız veya başarılı fark etmeksizin kalan temp dosyalar silinmeye çalışılır.

Commit tamamen başarılıysa `.txn.bak` dosyaları da silinir.

---

## Neyi garanti etmeye çalışır?

Tek request sırasında kart/log/yükleme gibi birden çok workbook’un birbiriyle mümkün olduğunca uyumlu kalmasını sağlamaya çalışır.

---

## Ne garanti etmez?

Kodun kendisi de Excel’in transactional database olmadığını söylüyor.

Özellikle:

- birden çok dosyaya yapılan replace gerçek tek-filesystem transaction değildir
- process/power crash tam ortada olursa ACID database semantiği yoktur
- directory fsync görünmüyor
- rollback sırasında restore OSError’ları `pass` ile geçiliyor

Dolayısıyla tasarım **transaction-benzeri** davranış sağlamaya çalışıyor; database transaction garantisi vermiyor.

---

## Excel açıkken ne olabilir?

Windows’ta hedef workbook başka program tarafından kilitliyse `os.replace` / dosya erişimi `PermissionError` üretebilir.

Flask’ta genel bir `PermissionError` handler’ı vardır ve API request’lerinde 423 döndürmeye çalışır.

---

## Değiştirirken dikkat

Şu patterni bozmak risklidir:

```text
write temp
→ fsync temp
→ backup existing
→ replace
→ rollback
```

Doğrudan:

```python
wb.save(KARTLAR_DOSYA)
```

yaklaşımına dönmek yarım / corrupt write riskini yükseltebilir.

---

# 8.9 Backup mekanizması

Sistemde birden fazla backup türü var.

---

## `_gunluk_yedek(dosya)`

Kart dosyası için günlük yedek üretir.

Format:

```text
data/yedekler/YYYYMMDD_kartlar.xlsx
```

Aynı gün dosya zaten varsa tekrar kopyalamaz.

Bu nedenle “günün ilk pre-write kopyası” gibi davranır.

---

## `anlik_yedek(etiket)`

Kart/log/yükleme dosyalarını timestamp’li bir klasöre kopyalar.

Format:

```text
data/yedekler/
└── YYYYMMDD_HHMMSS_<etiket>/
    ├── kartlar.xlsx
    ├── islem_logu.xlsx
    └── yuklemeler.xlsx
```

Var olan dosyalar kopyalanır.

Import öncesi:

```text
import_oncesi
```

Restore öncesi:

```text
geri_yukleme_oncesi
```

etiketleri kullanılıyor.

---

# 8.10 Backup retention

Sabitler:

```python
ANLIK_YEDEK_SAKLA = 30
GUNLUK_YEDEK_GUN = 90
```

`yedekleri_buda()`:

- timestamp pattern’li anlık backup klasörlerinden en yeni 30 tanesini tutar
- günlük `YYYYMMDD_kartlar.xlsx` yedeklerini 90 günden eskiyse silmeye çalışır

Hata olursa bu pruning işlemi sessizce geçebilir.

---

# 8.11 `yedekleri_getir()`

Yönetim UI’da gösterilecek restore adaylarını üretir.

Hem:

- anlık klasör backup’ları
- günlük tek `kartlar.xlsx` backup’ları

listelenir.

UI default olarak:

```python
depo.yedekleri_getir(12)
```

ile son 12 taneyi gösteriyor.

---

# 8.12 `_yedek_kart_dosyasi_bul()`

## Amaç

Kullanıcının gönderdiği backup adından güvenli kart dosyası bulmak.

### Kontroller

- boş mu?
- basename dışında path içeriyor mu?
- backup root dışına çıkıyor mu?
- klasörse içinde `kartlar.xlsx` var mı?
- tek dosyaysa beklenen günlük naming pattern’ine uyuyor mu?

Bu path traversal riskine karşı koruma sağlar.

---

# 8.13 `yedekten_geri_yukle()`

## Amaç

Seçilen kart backup’ını ana kart state’i haline getirmek.

---

## Ne zaman çağrılır?

Admin:

```text
POST /yonetim/yedek-geri-yukle
```

---

## Adımlar

### 1. Backup kart dosyasını bul

```python
_yedek_kart_dosyasi_bul(...)
```

### 2. Backup kartlarını oku ve normalize et

```python
_oku(...)
_kart_normalize(...)
```

### 3. Boş backup’ı reddet

```text
Seçilen yedekte hiç kart bulunmuyor.
```

### 4. Tüm kartları validate et

```python
_kart_listesi_dogrula(...)
```

### 5. Mevcut state için koruma backup’ı al

```python
anlik_yedek("geri_yukleme_oncesi")
```

Bu restore işleminden geri dönebilmeyi kolaylaştırır.

### 6. RAM eski kart/log kopyalarını al

### 7. `_kartlar = yeni_kartlar`

### 8. Audit log’a yeni kayıt ekle

```text
YEDEKTEN GERİ YÜKLENDİ
```

### 9. Kart + log dosyasını birlikte yaz

### 10. Hata olursa RAM rollback

---

## Çok önemli audit davranışı

`yonetim.html` açıkça:

```text
Yalnız kart verisi geri alınır; audit logu geriye sarılmaz.
```

diyor.

Kodda da backup’taki `islem_logu.xlsx` restore edilmiyor.

Mevcut audit listesine yalnız yeni:

```text
YEDEKTEN GERİ YÜKLENDİ
```

kaydı ekleniyor.

Bu bilinçli bir tasarım davranışıdır.

---

# 8.14 Process lock

Dosya:

```text
data/sunucu.lock
```

---

## Amaç

Aynı `data/` klasörünü iki PDGM process’inin aynı anda kullanmasını engellemek.

Çünkü:

```text
RLock
```

yalnız bir process içindeki thread’leri koordine eder.

---

## Lock içeriği

```text
PID|YYYY-MM-DD HH:MM:SS
```

---

# 8.15 `_process_kilit_sahibi()`

Lock dosyasının ilk parçasını PID olarak okumaya çalışır.

Dosya:

- yoksa
- okunamıyorsa
- PID parse edilemiyorsa

`None` dönebilir.

---

# 8.16 `_pid_calisiyor_mu(pid)`

## Amaç

Lock içindeki PID hâlâ canlı mı kontrol etmek.

### Windows

`tasklist` çağırır.

Tasklist çalıştırılamazsa:

```python
return True
```

yani **fail-closed** davranır.

Başka bir deyişle emin değilse “process yaşıyor olabilir” kabul eder.

### Non-Windows

`os.kill(pid, 0)` kullanır.

---

# 8.17 `_pid_python_mu(pid)`

## Amaç

Canlı PID gerçekten Python process’i mi kontrol etmek.

Bu, PID reuse durumunu ayırmak için kullanılıyor.

Windows’ta `tasklist` output’unda:

```text
python.exe
pythonw.exe
```

aranıyor.

Karar verilemezse yine fail-closed davranış tercih ediliyor.

---

# 8.18 `process_kilidi_al()`

## Amaç

Lock’u atomik şekilde kazanmak.

---

## Durum 1 — Lock zaten bu PID’e ait

```python
if mevcut_pid == os.getpid():
    return
```

---

## Durum 2 — Lock canlı Python PID’e ait

Startup reddedilir:

```text
İkinci sunucu açmayın.
```

---

## Durum 3 — Stale lock

PID:

- artık çalışmıyorsa
- veya PID yeniden kullanılmış ama process Python değilse

lock stale kabul edilir.

Eski lock önce random stale adına:

```python
os.replace(...)
```

ile taşınır.

---

## Yeni lock nasıl atomik oluşturuluyor?

```python
os.O_WRONLY | os.O_CREAT | os.O_EXCL
```

ile `os.open()` kullanılıyor.

`O_EXCL`, başka process aynı anda file oluşturmuşsa ikinci tarafın kazanmasını engeller.

Lock file yazıldıktan sonra:

```python
flush()
os.fsync()
```

yapılır.

---

## Atexit cleanup

Process kapanırken:

- lock hâlâ kendi PID’ine aitse
- dosya silinmeye çalışılır

---

## Değiştirirken dikkat

Bu lock uygulamanın:

```text
tek process
```

varsayımını teknik olarak enforce eden parçalardan biridir.

Waitress’i 2–4 worker process’e çıkarıp bu lock’u olduğu gibi bırakırsanız ikinci worker başlatılamaz.

---

# 8.19 Startup recovery gerçekte ne yapıyor?

Kodda “startup recovery” adıyla otomatik olarak:

```text
bozuk kartlar.xlsx → en son yedek
```

dönüşü yapan bir fonksiyon yok.

Mevcut davranış:

### Kart dosyası bozuksa

- startup durur
- hata açıklaması yazılır
- backup listesi gösterilir
- kullanıcıya manuel düzeltme / backup kopyalama talimatı verilir

### `islem_logu.xlsx` veya `yuklemeler.xlsx` bozuksa

- dosya `.bozuk_...` olarak kenara alınabilir
- boş listeyle devam edilir

### Eksik dosya varsa

- boş workbook oluşturulur

Bu nedenle otomatik kart restore mekanizması **dump’ta görünmüyor**.

---

# 9) Excel import sistemi

# 9.1 Neden COM var?

`excel_araclari.py` docstring’i açıkça:

```text
Microsoft Excel COM ile values-only snapshot oluşturulur.
```

diyor.

Akış:

```text
Orijinal kaynak Excel
        ↓
Microsoft Excel COM
        ↓
geçici values-only .xlsx
        ↓
openpyxl parser
```

Bu tasarımın koddan doğrulanabilen sonucu şudur:

- parser doğrudan kaynak workbook’un formül yapısına bağlı kalmaz
- COM üzerinden Excel’in hücre `Value2` değerleri alınır
- hedef snapshot yeni `.xlsx` workbook’tur

“COM neden özellikle seçildi?” konusunda bunun ötesinde ürün kararı açıklaması dump’ta yok.

---

# 9.2 COM kullanılabilirlik şartı

Import fonksiyonu:

```python
if pythoncom is None or win32com is None:
```

ise:

```text
Excel COM aktarımı yalnızca Windows + Microsoft Excel ortamında çalışır
(pywin32 gerekli).
```

hatası verir.

Bu, import yolunun Windows + Microsoft Excel bağımlılığını açıkça doğruluyor.

---

# 9.3 Import serialization

Global:

```python
_import_kilidi = threading.Lock()
```

var.

Public fonksiyon:

```python
excelden_aktar()
```

şunu yapıyor:

```python
with _import_kilidi:
    return _excelden_aktar(...)
```

Yani aynı process içinde iki Excel COM import işlemi aynı anda çalıştırılmıyor.

---

# 9.4 COM thread initialization

Her snapshot çağrısında:

```python
pythoncom.CoInitialize()
```

başta, sonunda:

```python
pythoncom.CoUninitialize()
```

çağrılıyor.

Bu, COM işleminin request thread’i üzerinde initialize edilmesi için önemlidir.

---

# 9.5 Excel Application güvenlik ayarları

Yeni Excel instance:

```python
win32com.client.DispatchEx("Excel.Application")
```

ile açılıyor.

Ardından:

```python
Visible = False
DisplayAlerts = False
AskToUpdateLinks = False
EnableEvents = False
AutomationSecurity = 3
```

ayarlanıyor.

`3` sabiti:

```python
MSO_AUTOMATION_SECURITY_FORCE_DISABLE
```

adıyla tanımlı.

Bu macro automation security’yi force-disable etmeye çalışır.

Workbook ayrıca:

```python
UpdateLinks=0
ReadOnly=True
IgnoreReadOnlyRecommended=True
AddToMru=False
```

ile açılıyor.

Kod açıkça external link update çağırmıyor ve full calculation çağrısı da yapmıyor.

---

# 9.6 Yalnız `MAKİNE` sheet kullanılır

Sheet adı:

```python
KAYNAK_SAYFA_ADI = "MAKİNE"
```

Sheet karşılaştırmasında Türkçe karakter sadeleştirme kullanılıyor.

Bulunamazsa dosyadaki sheet listesi hata mesajında gösteriliyor.

---

# 9.7 Hidden kolonlar neden yok?

COM snapshot loop’u her source column için:

```python
if bool(kaynak_ws.Columns(kaynak_sutun).Hidden):
    continue
```

yapıyor.

Yani hidden kolonlar snapshot’a hiç alınmıyor.

Hidden satırlar için aynı skip kodu yok.

Dolayısıyla docstring’le uyumlu olarak hidden satırlar korunuyor.

---

# 9.8 Values-only snapshot

Her visible column range için:

```python
hedef_aralik.Value2 = kaynak_aralik.Value2
```

yapılıyor.

Bu, formül nesnesini / formatting’i kopyalamak yerine hücre değerlerini hedef workbook’a taşıyor.

Sonra:

```python
SaveAs(..., FileFormat=51)
```

ile xlsx kaydediliyor.

Snapshot geçici dosyadır ve parsing bittikten sonra silinmeye çalışılır.

---

# 9.9 Header mapping

Başlıklar `_sadelestir()` ile normalize edilir.

Eşlenen temel başlıklar:

| Kaynak Excel | Internal alan |
|---|---|
| NO / Sıra | `sira` |
| Talep NO | `talep_no` |
| Talep Sahibi | `talep_sahibi` |
| Kart Stok No | `stok_no` |
| Kart Üretim Adet / Üretim Adet / Adet | `adet_metin` |
| Planlanan Başlangıç T. | `plan_hafta` |
| Dizgi Başlama Tarihi | `plan_baslama` |
| Planlanan Teslim T. | `plan_teslim` |
| Gerçekleşen Teslim T. | `gerceklesen_teslim` |
| DURUM | `excel_durum` |
| PCB | `pcb` |

---

# 9.10 Header satırı nasıl bulunuyor?

İlk 10 satır taranıyor.

Bir satırda:

```text
TALEP NO
```

veya:

```text
KART STOK NO
```

varsa header satırı kabul ediliyor.

---

# 9.11 Zorunlu kaynak kolonlar

Import için kesin gerekli:

```text
Talep NO
Kart Stok No
Kart Üretim Adet
```

Bunlardan biri yoksa import reddedilir.

---

# 9.12 Duplicate visible source column kontrolü

Normalize edilmiş bir business alanına birden fazla görünür Excel kolonu map olursa import reddedilir.

Bu, belirsiz header mapping’i engellemeye çalışır.

---

# 9.13 Satır ne zaman kart kabul edilir?

Satır tamamen boşsa geçilir.

Sonra:

```python
talep_no
stok_no
```

okunur.

İkisinden biri yoksa satır kart kabul edilmez ve skip edilir.

---

# 9.14 Anahtar duplicate kontrolü

Her parsed kart:

```python
anahtar = f"{talep_no}|{stok_no}"
```

oluşturur.

Aynı upload içinde aynı anahtar ikinci kez görülürse import reddedilir.

---

# 9.15 `adet_coz()`

Üretim adedini pozitif integer’a çevirmeye çalışır.

Kabul örnekleri kod docstring’inde:

```text
400
400.0
"400 ADET"
"1.500 ADET"
"1,500 ADET"
"1 500 ADET"
```

Reddedilen örnekler:

```text
400.5
"400.5 ADET"
"400,5 ADET"
"-5 ADET"
"400 / 500 ADET"
```

En az:

```text
1
```

olmalı.

---

# 9.16 Tarih doğrulama

Kaynak date parse edilemiyorsa import error verir.

Ayrıca:

```text
Dizgi Başlama Tarihi > Planlanan Teslim Tarihi
```

ise satır / import reddedilir.

---

# 9.17 Kaynak DURUM parse

Mapping:

```text
PLANA ALINDI
DİZGİDE
HAZIR
TESLİM EDİLDİ
```

DURUM:

- boş
- veya farklı / geçersiz

ise:

```python
ilk_durum = None
```

ve uyarı sayılır.

Bu kart parse’dan atılmaz.

---

# 9.18 Gerçekleşen teslim ile durum çelişkisi

### Kaynak `TESLİM EDİLDİ`, gerçekleşen teslim boşsa

Parser:

```python
ilk_durum = None
```

yapar ve uyarı sayısını artırır.

Yani bu satır otomatik teslim edilmiş kabul edilmez.

### Kaynak TESLİM EDİLDİ değil ama gerçekleşen teslim doluysa

Uyarı sayısı artırılır.

Yeni kart oluşturulurken gerçekleşen teslim yalnız durum gerçekten `TESLİM EDİLDİ` ise karta yazılır.

---

# 9.19 Parse sonucu

Her kart `depo.excel_import_uygula()` fonksiyonuna yaklaşık şu yapıyla geçer:

```python
{
    "anahtar": "...",
    "talep_no": "...",
    "stok_no": "...",
    "plan": {
        "sira": ...,
        "talep_sahibi": ...,
        "toplam_adet": ...,
        "adet_metin": ...,
        "plan_hafta": ...,
        "plan_baslama": ...,
        "plan_teslim": ...,
        "excel_durum": ...,
        "pcb": ...,
    },
    "gerceklesen_teslim": ...,
    "ilk_durum": ...
}
```

---

# 9.20 Import öncesi yedek

`excel_import_uygula()` ilk olarak:

```python
anlik_yedek("import_oncesi")
```

çağırıyor.

Bu import rollback’ından ayrı bir operasyonel koruma kopyasıdır.

---

# 9.21 Mevcut kartla merge

Mevcut kartlar:

```python
mevcut_harita = {
    kart["anahtar"]: kart
}
```

ile map ediliyor.

Aynı anahtar geldiyse mevcut kart update edilir.

---

# 9.22 Workflow neden ezilmiyor?

Mevcut kart için önce şu alanlar ayrı dict’e alınır:

```python
workflow = {
    "durum": ...,
    "baslangic_adet": ...,
    "tamamlanan_adet": ...,
    "baslama_zamani": ...,
    "bitis_zamani": ...,
    "teslim_zamani": ...,
    "operator": ...,
    "aciklama": ...,
    "gerceklesen_teslim": ...,
}
```

Sonra:

```python
mevcut.update(plan)
mevcut.update(workflow)
```

yapılır.

Yani kaynak Excel’den gelen plan/source alanları update edilirken uygulamanın yaşayan workflow state’i geri konur.

Korunan önemli alanlar:

- durum
- başlangıç adedi
- tamamlanan adet
- gerçek başlangıç zamanı
- üretim bitiş zamanı
- teslim zamanı
- operatör
- kullanıcı notu
- gerçekleşen teslim

---

# 9.23 HAZIR / TESLİM conflict ne zaman import’u reddeder?

Bu kısım çok kritiktir.

### Genel kural

Yeni kaynak toplam adedi:

```text
mevcut tamamlanan adet
```

değerinden küçükse:

```text
import iptal
```

Örnek:

```text
sistemde tamamlanan = 80
Excel yeni toplam = 60
```

reddedilir.

---

## HAZIR / TESLİM EDİLDİ özel kuralı

Kart:

```text
HAZIR
```

veya:

```text
TESLİM EDİLDİ
```

ise ve:

```python
yeni_toplam != tamamlanan
```

olursa import reddedilir.

Bu iki durumda normal validation’a göre zaten:

```text
tamamlanan == eski toplam
```

olmalıdır.

Dolayısıyla tamamlanmış / teslim edilmiş kartın toplam adedini import yoluyla değiştirmek engellenmiş olur.

Hata:

```text
admin kontrolü gerekir
```

mesajı verir.

---

# 9.24 DİZGİDE kartta toplam adet değişebilir mi?

Koddan çıkan sonuç:

Yeni toplam:

```text
tamamlanan_adet
```

değerinden küçük olmadığı sürece ve kart HAZIR/TESLİM değilse import tarafından değiştirilebilir.

Workflow state’i korunur.

Bu nedenle DİZGİDE bir kartın toplamı örneğin 100’den 120’ye çıkabilir.

---

# 9.25 Yeni kart nasıl oluşturuluyor?

Anahtar mevcut değilse yeni ID verilir.

Yeni kartın başlangıç state’i kaynak `DURUM` üzerinden oluşturulur.

### PLANA ALINDI

```text
tamamlanan = 0
baslangic_adet = 0
```

### DİZGİDE

```text
tamamlanan = 0
baslangic_adet = toplam
```

### HAZIR

```text
tamamlanan = toplam
baslangic_adet = toplam
```

### TESLİM EDİLDİ

```text
tamamlanan = toplam
baslangic_adet = toplam
gerceklesen_teslim = kaynak gerçekleşen teslim
```

`operator` geçerli source durumlarında:

```text
Excel
```

olarak yazılır.

Timestamp alanları yeni import kartında `None` bırakılıyor.

---

# 9.26 Kaynakta artık olmayan kart

Import sonunda tüm `EXCEL` kaynaklı mevcut kartlar dolaşılır.

Yeni kaynakta anahtarı görülmeyen ve daha önce `source_active=1` olan kart:

```python
source_active = 0
```

yapılır.

Kart silinmez.

Workflow korunur.

---

# 9.27 Import audit ve geçmiş

`yuklemeler.xlsx` için kayıt:

- zaman
- kullanıcı
- dosya
- satır
- yeni
- güncellenen
- kaynakta olmayan
- uyarı

Audit log:

```text
EXCEL YÜKLENDİ
```

Detay içinde:

- satır
- yeni
- güncellenen
- workflow korundu
- kaynakta yok
- uyarı
- backup adı

yer alır.

---

# 9.28 Import commit

Üç dosya birlikte `_coklu_yaz()` ile yazılır:

```text
kartlar.xlsx
islem_logu.xlsx
yuklemeler.xlsx
```

RAM’de hata olursa:

```text
_kartlar
_loglar
_yuklemeler
```

eski deep copy’lere döndürülür.

---

# 10) Route ve API rehberi

## 10.1 Tüm önemli route’lar

| Route | Method | Rol | Handler | Ne yapar | İş verisi mutasyonu? |
|---|---|---|---|---|---|
| `/giris` | GET | Public | `giris` | Login formu | Hayır |
| `/giris` | POST | Public | `giris` | Authentication, session oluşturma, login audit | Evet: session + audit |
| `/cikis` | POST | admin/operator/gozlemci | `cikis` | Logout, session clear, audit | Evet |
| `/` | GET | Public/Session yönlendirme | `ana` | Role göre ekran redirect | Hayır |
| `/panel` | GET | admin/operator/gozlemci | `panel` | Pano render | Hayır |
| `/monitor` | GET | admin/operator/gozlemci | `monitor` | Büyük ekran monitor render | Hayır |
| `/operator` | GET | admin/operator | `operator` | Operasyon kart ekranı | Hayır |
| `/yonetim` | GET | admin | `yonetim` | Yönetim ekranı | Hayır |
| `/ozet` | GET | admin/operator/gozlemci | `ozet` | İstatistik özeti | Hayır |
| `/api/veriler` | GET | admin/operator/gozlemci | `api_veriler` | `_pano_verisi()` JSON | Hayır |
| `/api/basla` | POST | admin/operator | `api_basla` | Kartı DİZGİDE yapar | Evet |
| `/api/bitir` | POST | admin/operator | `api_bitir` | Tamamlanan üretim adedi ekler | Evet |
| `/api/hazirla` | POST | admin/operator | `api_hazirla` | DİZGİDE → HAZIR | Evet |
| `/api/teslim-et` | POST | admin/operator | `api_teslim_et` | HAZIR → TESLİM EDİLDİ | Evet |
| `/api/not` | POST | admin/operator | `api_not` | Kart notu günceller | Evet |
| `/api/admin/kart-ekle` | POST | admin | `api_kart_ekle` | Manuel PLANA ALINDI kart oluşturur | Evet |
| `/api/admin/duzenle` | POST | admin | `api_duzenle` | Admin kart/workflow düzeltmesi | Evet |
| `/api/admin/kart-sil` | POST | admin | `api_kart_sil` | Kartı soft-delete/gizli yapar | Evet |
| `/yonetim/kart-geri-getir` | POST | admin | `kart_geri_getir` | Gizli kartı geri getirir | Evet |
| `/yonetim/yedek-geri-yukle` | POST | admin | `yedek_geri_yukle` | Kart state’ini backup’tan restore eder | Evet |
| `/yonetim/yukle` | POST | admin | `yukle` | Kaynak Excel upload + import | Evet |
| `/yonetim/yeniden-oku` | POST | admin | `yeniden_oku` | `kartlar.xlsx` doğrulayıp RAM’e yeniden yükler, audit yazar | Evet: RAM + audit |
| `/yonetim/kayit-dosyasi/<hangi>` | GET | admin | `kayit_dosyasi` | kart/log/yükleme workbook download | İş state’i hayır; app log yazar |
| `/yonetim/rapor` | GET | admin | `rapor_indir` | RAM verilerinden rapor xlsx üretir | Kalıcı iş state’i hayır |

---

# 10.2 `_api_kart_islemi()`

Operatör/admin JSON API error wrapper’ıdır.

Şu exception’ları map eder:

| Exception | HTTP |
|---|---:|
| `KartBulunamadi` | 404 |
| `IsKuralHatasi` | 409 |
| `VeriDogrulamaHatasi` | 400 |
| `TypeError` | 400 |
| `ValueError` | 400 |

Response biçimi:

```json
{
  "hata": "Açıklama"
}
```

`PermissionError` burada catch edilmiyor; genel Flask error handler’a bırakılıyor.

---

# 10.3 `/api/basla`

Request:

```json
{
  "kart_id": 15,
  "adet": 80,
  "not": "İlk seri"
}
```

Başarı:

```json
{
  "tamam": true,
  "mesaj": "Kart DİZGİDE durumuna alındı.",
  "kart": {
    "...": "kart görünümü"
  }
}
```

---

# 10.4 `/api/bitir`

Request:

```json
{
  "kart_id": 15,
  "adet": 25,
  "not": "25 adet tamamlandı"
}
```

Kısmi response:

```json
{
  "tamam": true,
  "uretim_bitti": false,
  "mesaj": "25/80 adet tamamlandı.",
  "kart": {
    "...": "..."
  }
}
```

Son adet tamamlanırsa:

```json
{
  "tamam": true,
  "uretim_bitti": true,
  "mesaj": "Üretim adedi tamamlandı. Kart HAZIR'a otomatik alınmadı.",
  "kart": {
    "...": "..."
  }
}
```

---

# 10.5 `/api/hazirla`

Request:

```json
{
  "kart_id": 15,
  "not": "Üretim kontrol edildi"
}
```

Operator otomatik follow-up akışında yalnız:

```json
{
  "kart_id": 15
}
```

gönderiyor.

Başarı:

```json
{
  "tamam": true,
  "mesaj": "Kart HAZIR durumuna alındı.",
  "kart": {
    "...": "..."
  }
}
```

---

# 10.6 `/api/teslim-et`

Request:

```json
{
  "kart_id": 15,
  "not": "Teslim edildi"
}
```

Başarı:

```json
{
  "tamam": true,
  "mesaj": "Kart TESLİM EDİLDİ olarak kaydedildi.",
  "kart": {
    "...": "..."
  }
}
```

---

# 10.7 `/api/not`

Request:

```json
{
  "kart_id": 15,
  "not": "Yeni kart notu"
}
```

Başarı:

```json
{
  "tamam": true,
  "kart": {
    "...": "..."
  }
}
```

---

# 10.8 `/api/admin/kart-ekle`

Request alanları:

```json
{
  "sira": 1,
  "talep_no": "TLP-001",
  "talep_sahibi": "Kullanıcı",
  "stok_no": "STK-001",
  "toplam_adet": 100,
  "plan_hafta": "34",
  "plan_baslama": "2026-08-20",
  "plan_teslim": "2026-08-25",
  "gerceklesen_teslim": null,
  "pcb": "HBT",
  "not": "Manuel kart"
}
```

Başarı:

```text
HTTP 201
```

ve:

```json
{
  "tamam": true,
  "kart": { "...": "..." }
}
```

Yeni kart daima:

```text
PLANA ALINDI
```

başlar.

---

# 10.9 `/api/admin/duzenle`

Request:

```json
{
  "kart_id": 15,
  "durum": "DİZGİDE",
  "tamamlanan_adet": 40,
  "toplam_adet": 100,
  "plan_hafta": "34",
  "plan_baslama": "2026-08-20",
  "plan_teslim": "2026-08-25",
  "gerceklesen_teslim": "",
  "not": "Admin düzeltmesi"
}
```

Bu endpoint normal workflow step’lerini taklit etmek zorunda değildir.

Backend seçilen duruma göre tutarlı alanları kendisi ayarlamaya çalışır.

Örneğin:

### `PLANA ALINDI`

- tamamlanan = 0
- başlangıç adedi = 0
- timestamps reset
- gerçekleşen teslim reset

### `HAZIR`

- tamamlanan = toplam
- başlangıç / bitiş zamanı yoksa oluştur
- teslim zamanı reset

### `TESLİM EDİLDİ`

- tamamlanan = toplam
- gerçekleşen teslim yoksa bugün
- teslim zamanı yoksa şimdi

Audit:

```text
ADMİN DÜZENLEDİ
```

---

# 10.10 `/api/admin/kart-sil`

İsmi “sil” olsa da gerçek hard delete yapmıyor.

Depo fonksiyonu:

```python
admin_kart_gizle()
```

yalnız:

```python
admin_gizli = 1
```

yapar.

Audit:

```text
KART LİSTEDEN GİZLENDİ
```

Detay açıkça:

```text
Kart silinmedi; admin gizli olarak işaretlendi.
```

---

# 10.11 `/yonetim/yukle`

Form:

```text
multipart/form-data
```

Alan:

```text
dosya
```

CSRF hidden field var.

Kabul edilen uzantılar:

```text
.xlsx
.xlsm
```

Dosya adı:

```python
secure_filename(...)
```

ile temizleniyor.

Kayıt adı:

```text
YYYYMMDD_HHMMSS_<8hex>_<guvenli_ad>
```

şeklinde oluşturuluyor.

Upload dosyaları `data/yuklenen_exceller` altında tutuluyor.

En yeni 20 yükleme dışındakiler `_yuklenen_exceleri_buda()` ile silinmeye çalışılıyor.

---

# 11) Frontend nasıl bağlanıyor?

# 11.1 Ortak `pdgmFetch()`

`base.html` tüm extended sayfalara şu helper’ı sağlıyor:

```javascript
async function pdgmFetch(url, options = {}) {
    ...
}
```

---

## CSRF header

Her çağrıda:

```javascript
headers.set("X-CSRF-Token", csrfToken());
```

yapılıyor.

Token `<meta>` tag’den okunuyor.

---

## JSON Content-Type

Body varsa ve FormData değilse:

```text
Content-Type: application/json
```

otomatik ekleniyor.

---

## Cookie gönderimi

```javascript
credentials: "same-origin"
```

kullanılıyor.

---

## Response parsing

Content-Type JSON ise:

```javascript
await response.json()
```

aksi halde text alınır.

Non-2xx response’ta:

```javascript
throw new Error(...)
```

yapılır.

---

# 11.2 Toast sistemi

Başarı:

```javascript
toast(...)
```

Hata:

```javascript
hataMesaji(hata)
```

ile bottom/overlay toast UI’a yansır.

Toast yaklaşık 3.2 saniye sonra kaldırılıyor.

---

# 11.3 Operator ekranı

Operator JS tamamen `operator.html` içinde.

---

## Filtre / arama

State `sessionStorage` içinde tutuluyor:

```text
pdgm-op-filtre
pdgm-op-arama
pdgm-op-scroll
```

Default filtre:

```text
AKTIF
```

Aktif tanımı frontend’de:

```text
PLANA ALINDI
DİZGİDE
HAZIR
```

---

## Dizgiye Al

Button:

```text
data-baslat
```

dialog açar.

Submit:

```text
POST /api/basla
```

gönderir.

Başarı sonrası:

- toast
- dialog kapanır
- arama ve scroll state saklanır
- 350 ms sonra page reload

---

## Adet Bitir

Button:

```text
data-bitir
```

dialog açar.

Submit:

```text
POST /api/bitir
```

---

## Üretim tamamen bittiyse

Frontend kullanıcıya confirm gösterir.

Kabul:

```text
POST /api/hazirla
```

Reddetme:

- backend’den gelen “otomatik alınmadı” mesajı gösterilir
- kart DİZGİDE kalır

---

## Hazıra Al

`data-hazirla` button:

```text
POST /api/hazirla
```

---

## Teslim

`data-teslim`:

```text
POST /api/teslim-et
```

Önce physical delivery confirm’i sorulur.

---

## Not

`data-not` dialogu:

```text
POST /api/not
```

---

# 11.4 Panel ekranı

`panel.html` backend mutation yapmıyor.

JS yalnız:

- arama
- durum filtresi
- scroll koruma
- manual reload
- periyodik reload

işlerini yapıyor.

State:

```text
pdgm-panel-filtre
pdgm-panel-arama
pdgm-panel-scroll
```

30 saniye sonunda:

```javascript
window.location.reload()
```

çağrılıyor.

Yani panel `/api/veriler` ile incremental update yapmıyor.

---

# 11.5 Monitor ekranı

Monitor JS:

### Saat

Her 1 saniyede browser local clock ile:

```javascript
saatiGuncelle()
```

çalışıyor.

### Kart rotasyonu

Her grup:

```text
data-monitor-grup
data-sayfa-boyutu
```

ile page’leniyor.

#### DİZGİDE

```text
sayfa boyutu = 3
```

3’ten fazla kart varsa her:

```text
5 saniye
```

sonraki üçlü görünür.

#### PLANA ALINDI

```text
sayfa boyutu = 6
```

6’dan fazla kart varsa her 5 saniye sonraki grup görünür.

### Server data refresh

Her:

```text
30 saniye
```

full-page reload var.

### HAZIR

Mevcut dump’taki monitor route/template HAZIR grubunu render etmiyor.

---

# 11.6 Yönetim ekranı

Burada iki frontend yaklaşımı bir arada kullanılıyor.

---

## Normal HTML POST form’ları

Şunlar klasik form submit:

- Excel upload
- kart dosyası yeniden oku
- gizlenen kart geri getir
- yedek geri yükle

Her formda:

```html
<input type="hidden" name="_csrf_token" ...>
```

var.

Sonuç Flask `flash()` ile sonraki sayfada gösteriliyor.

---

## AJAX / Fetch API işlemleri

`pdgmFetch()` kullanılanlar:

```text
/api/admin/kart-sil
/api/admin/duzenle
/api/admin/kart-ekle
```

Başarı:

- toast
- dialog kapama
- 350ms sonra reload

Hata:

```javascript
hataMesaji()
```

---

## Admin arama / filtre

Client-side.

Backend’e yeni query göndermez.

`data-admin-kart` elementleri hide/show edilir.

---

# 11.7 `/api/veriler` frontend’de kullanılıyor mu?

Route mevcut:

```text
GET /api/veriler
```

Ancak dump’taki template/JS dosyalarında bu endpoint’e bir çağrı görünmüyor.

Dolayısıyla mevcut frontend refresh mekanizması esas olarak full-page reload’dur.

---

# 12) Uçtan uca 5 senaryo

# 12.1 Senaryo 1 — Login

```text
giris.html
   │
   │ POST /giris
   │ kullanici + sifre
   ▼
app.giris()
   │
   ├─ rate limit kontrolü
   ├─ input size kontrolü
   ├─ _kullanicilari_al()
   ├─ check_password_hash()
   │
   ├─ session.clear()
   ├─ session[kullanici, rol, ad]
   ├─ csrf token üret
   │
   └─ depo.log_ekle("GİRİŞ YAPILDI")
            │
            ▼
       islem_logu.xlsx
   │
   ▼
safe redirect
   │
   ├─ admin → /yonetim
   ├─ operator → /operator
   └─ gozlemci → /panel
```

### Log write başarısızsa

Authentication başarılı kalır.

Exception teknik loga yazılır.

---

# 12.2 Senaryo 2 — Dizgiye Al

```text
operator.html
   │
   │ "Dizgiye Al"
   ▼
dialog
   │
   │ POST /api/basla
   │ JSON kart_id/adet/not
   ▼
app.api_basla()
   │
   ├─ @yetki(admin, operator)
   ├─ @csrf_koru
   │
   ▼
depo.kart_baslat()
   │
   ├─ RLock
   ├─ kart var mı
   ├─ görünür mü
   ├─ durum PLANA ALINDI mı
   ├─ adet valid mi
   │
   ▼
_atomik_kart_islemi()
   │
   ├─ RAM backup
   ├─ kart durum=DİZGİDE
   ├─ audit append
   ├─ tüm kart validate
   ├─ günlük yedek
   │
   ▼
_coklu_yaz()
   │
   ├─ kartlar.xlsx
   └─ islem_logu.xlsx
   │
   ▼
JSON success
   │
   ▼
toast
   │
   ▼
350ms sonra reload
```

---

# 12.3 Senaryo 3 — Adet Bitir

## Kısmi üretim

```text
operator.html
   │
   │ POST /api/bitir
   ▼
app.api_bitir()
   │
   ▼
depo.kart_bitir()
   │
   ├─ durum DİZGİDE mi
   ├─ kalan > 0 mı
   ├─ adet <= kalan mı
   │
   ├─ tamamlanan += adet
   ├─ audit = KISMİ ÜRETİM
   ▼
_atomik_kart_islemi
   ▼
kartlar.xlsx + islem_logu.xlsx
   ▼
{
  uretim_bitti: false
}
   ▼
toast + reload
```

---

## Tam üretim

```text
/api/bitir
   │
   ▼
kart_bitir()
   │
   ├─ tamamlanan = toplam
   ├─ bitis_zamani = şimdi
   ├─ audit = ÜRETİM ADEDİ TAMAMLANDI
   └─ durum hâlâ DİZGİDE
   │
   ▼
{
  uretim_bitti: true
}
   │
   ▼
Browser confirm:
"HAZIR durumuna alınsın mı?"
   │
   ├─ Hayır
   │    └─ DİZGİDE kalır
   │
   └─ Evet
        │
        │ POST /api/hazirla
        ▼
      kart_hazirla()
        │
        ├─ tamamlanan == toplam kontrolü
        ├─ durum = HAZIR
        └─ audit
```

Bu iki ayrı backend write işlemidir.

---

# 12.4 Senaryo 4 — Excel import

```text
yonetim.html
   │
   │ select .xlsx/.xlsm
   │ POST /yonetim/yukle
   ▼
app.yukle()
   │
   ├─ CSRF
   ├─ admin role
   ├─ secure_filename
   ├─ extension check
   ├─ upload path oluştur
   └─ dosya.save()
   │
   ▼
excel_araclari.excelden_aktar()
   │
   ├─ _import_kilidi
   ▼
excel_deger_snapshot_olustur()
   │
   ├─ CoInitialize
   ├─ DispatchEx Excel
   ├─ macros disabled
   ├─ update links off
   ├─ read-only open
   ├─ MAKİNE sheet
   ├─ hidden columns skip
   └─ Value2 → temp xlsx
   │
   ▼
openpyxl
   │
   ├─ header detect
   ├─ required cols
   ├─ duplicate key
   ├─ adet parse
   ├─ date parse
   └─ durum parse
   │
   ▼
depo.excel_import_uygula()
   │
   ├─ RLock
   ├─ RAM backups
   ├─ anlik_yedek(import_oncesi)
   ├─ existing merge
   ├─ workflow preserve
   ├─ conflict checks
   ├─ source_active update
   ├─ audit
   ├─ upload history
   ├─ validation
   └─ _coklu_yaz(3 files)
   │
   ▼
result dict
   │
   ▼
app.yukle()
   │
   ├─ flash success / warning
   ├─ app log
   └─ uploaded files retention
   │
   ▼
redirect /yonetim
```

---

# 12.5 Senaryo 5 — Yedekten geri yükleme

```text
yonetim.html
   │
   │ yedek seçimi
   │ confirmation
   │ POST /yonetim/yedek-geri-yukle
   ▼
app.yedek_geri_yukle()
   │
   ├─ admin role
   └─ CSRF
   │
   ▼
depo.yedekten_geri_yukle()
   │
   ├─ backup name/path validate
   ├─ backup kartlar.xlsx read
   ├─ normalize
   ├─ validate
   │
   ├─ anlik_yedek("geri_yukleme_oncesi")
   │
   ├─ RAM old copies
   │
   ├─ _kartlar = backup kartları
   │
   ├─ yeni audit:
   │     YEDEKTEN GERİ YÜKLENDİ
   │
   └─ _coklu_yaz(
         kartlar.xlsx,
         islem_logu.xlsx
       )
   │
   ▼
success result
   │
   ▼
flash:
"Yedek geri yüklendi..."
```

### Audit neden geri sarılmıyor?

Çünkü backup’taki log dosyası ana `islem_logu.xlsx` üzerine konmuyor.

Mevcut audit trail korunup yeni restore event’i ekleniyor.

---

# 13) Güvenlik ve operasyon

# 13.1 Dump’ta görünen güvenlik kontrolleri

## Authentication

- password hash
- active user check
- session
- 12 saat permanent session lifetime

## Authorization

- role decorator
- admin/operator/gozlemci route ayrımı

## CSRF

Mutation endpoint’lerinin büyük bölümünde `@csrf_koru`.

Normal form ve fetch header desteği.

## Login rate limit

- IP bazlı
- 8 failure / 5 dakika
- RAM state

## Input sınırı

Login username/password size limit.

## Upload sınırı

```text
25 MB
```

## Upload filename cleanup

```python
secure_filename()
```

## Upload extension allowlist

```text
.xlsx
.xlsm
```

## Session cookie

- HttpOnly
- SameSite=Lax
- Secure yalnız `PDGM_HTTPS=1` ise

## HTTP security headers

- nosniff
- DENY frame
- same-origin referrer
- camera/mic/geolocation disable
- no-store dynamic responses

## Open redirect koruması

`_guvenli_devam_hedefi()`.

## Excel formula injection koruması

`depo._excel_hucre_yaz()`:

Eğer kullanıcı text’i:

```text
=
```

ile başlıyorsa cell `data_type = "s"` yapılıyor.

Amaç kullanıcı metninin Excel formülü haline gelmesini önlemek.

`excel_araclari.py` rapor writer’ında da aynı isimde helper vardır.

## COM güvenliği

- macro automation security force disable
- external link update kapalı
- events kapalı
- read-only source open

## Backup path validation

Restore path root dışına çıkamıyor.

## Process lock

İkinci PDGM process aynı data klasörüyle çalıştırılmıyor.

---

# 13.2 Dump’ta görünmeyen / doğrulanamayan kontroller

Aşağıdakiler bu dump’ta görünmüyor:

- Caddy config
- gerçek TLS certificate config
- HSTS
- CSP
- reverse proxy trusted-IP / `ProxyFix`
- merkezi SSO / LDAP / Active Directory
- MFA
- database-level authorization
- antivirus upload scanning
- Windows ACL config
- network firewall config
- Task Scheduler XML / gerçek retry ayarı
- offsite backup hedefinin gerçekten erişilebilir / immutable olduğu

`.env.example` yorumunda Caddy geçmesi bunların kurulu olduğunu doğrulamaz.

---

# 13.3 `run_pdgm.bat`

Script mantığı:

```text
repo klasörüne geç
   ↓
.venv python var mı?
   │
   ├─ hayır → hata + exit 1
   │
   └─ evet
        ↓
      python app.py
        ↓
      app exit code’u dışarı ver
```

Yorumda ayrıca Task Scheduler recommendation bulunuyor:

```text
Run only when user is logged on
```

ve failure restart davranışı yorumla belirtilmiş.

Ancak gerçek scheduler task dosyası dump’ta yok.

---

# 13.4 `yedek_disari_kopyala.bat`

Kaynak:

```text
<repo>\data
```

Hedef:

```text
\\dosyasunucu\yedek\pdgm
```

Önce hedef path erişilebilir mi kontrol eder.

Tarihi:

```text
YYYYMMDD
```

alır.

`robocopy`:

```bat
/E /R:2 /W:5 /NFL /NDL
```

ile çalışır.

Log:

```text
<hedef>\robocopy.log
```

altına append edilir.

60 günden eski tarih klasörleri silinir.

`robocopy` return code:

```text
0-7 → başarı/uyarı
8+  → hata
```

olarak yorumlanır.

---

# 13.5 Bilinen mimari sınırlar

Kodun kendisinden çıkan sınırlar:

## 1. Excel database değildir

Docstring açıkça bunu vurguluyor.

## 2. Tek process gerekir

`sunucu.lock` bunu enforce ediyor.

## 3. Çok thread beklenir

Waitress 8 thread kullanıyor.

## 4. Thread safety process içi `RLock` ile

Cross-process lock için file lock var ama shared multi-process execution desteklenmiyor.

## 5. Kaynak Excel importu Windows + Excel COM ister

`pywin32 + Microsoft Excel`.

## 6. COM import sırasında interactive user session ihtiyacı script yorumunda belirtilmiş

`run_pdgm.bat` Task Scheduler için “Run only when user is logged on” diyor.

## 7. Storage local filesystem semantiğine güveniyor

`os.replace`, `copy2`, `fsync`, lock file.

Data klasörünün hangi disk türünde olduğu dump’ta görünmüyor.

## 8. Excel dosyasının başka programda açık olması write’ı engelleyebilir

Buna özel kullanıcı mesajı vardır.

---

# 13.6 Log retention

Application log:

```text
2 MB × mevcut + 5 backup
```

RotatingFileHandler.

Audit Excel için:

```python
LOG_SINIRI = 20_000
LOG_SAKLA = 5_000
```

Log 20.000’i aşarsa eski kayıtlar:

```text
data/yedekler/YYYYMMDD_HHMMSS_islem_logu_arsiv.xlsx
```

dosyasına yazılır.

Ana log listesinde son 5.000 tutulur.

Bu archive dosyaları `yedekleri_buda()` içindeki anlık backup klasörü veya günlük kart backup pattern’i ile aynı retention mekanizmasına açıkça dahil değildir.

---

# 14) Yeni developer için “dokunma” listesi

Aşağıdaki alanlar değişiklik sırasında özellikle hassastır.

---

## 1. `kartlar.xlsx` source-of-truth sözleşmesi

**Kural:** Ana kart state’ini başka bir yerde ikinci gerçek kaynak gibi tutma.

**Neden:** `depo.py` tüm state modelini Excel + RAM senkronizasyonu üzerine kuruyor.

**Bozulursa:** Restart sonrası başka state, runtime’da başka state görülebilir.

---

## 2. Tek process varsayımını fark etmeden değiştirme

**Kural:** Waitress’i multi-process worker mimarisine çevirmeden önce storage lock modelini yeniden tasarla.

**Neden:** `RLock` process-local; `sunucu.lock` ikinci process’i özellikle reddediyor.

**Bozulursa:** Startup engellenir veya lock kaldırılırsa aynı Excel’lere eşzamanlı write yarışları oluşabilir.

---

## 3. `_atomik_kart_islemi()` pattern’ini bypass etme

**Kural:** Yeni workflow mutation’larını mümkün olduğunca bu rollback kalıbına bağla.

**Neden:** RAM + validation + kart/log commit birlikte ele alınıyor.

**Bozulursa:** RAM başarılı görünüp disk başarısız kalabilir veya audit event state ile ayrışabilir.

---

## 4. `_coklu_yaz()` yerine doğrudan workbook save kullanma

**Kural:** Ana dosyalarda temp/backup/replace yaklaşımını koru.

**Neden:** Mevcut yapı yarım write riskini azaltmak için tasarlanmış.

**Bozulursa:** Bir write sırasında hedef Excel’in kendisi partial/corrupt hale gelebilir.

---

## 5. `tamamlanan_adet` invariant’larını gevşetme

**Kural:**

```text
0 <= tamamlanan <= toplam
```

ve:

```text
HAZIR/TESLİM → tamamlanan == toplam
```

korunmalı.

**Neden:** UI ve workflow bu varsayımları kullanıyor.

**Bozulursa:** HAZIR fakat üretimi bitmemiş kart gibi çelişkiler oluşur.

---

## 6. `kart_bitir()` ile `kart_hazirla()`yı fark etmeden birleştirme

**Kural:** Mevcut davranışta üretim tamamlanması ve HAZIR state transition ayrı olaylardır.

**Neden:** Kullanıcıya son üretimden sonra HAZIR’a geçiş seçeneği bırakılmış.

**Bozulursa:** İş süreci semantiği ve iki ayrı audit olayı değişir.

---

## 7. Importta workflow koruma dict’ini kaldırma

**Kural:** Mevcut karta kaynak Excel update edilirken operator workflow state’i korunmalı.

**Neden:** Kaynak plan Excel’i yaşayan üretim state’inin sahibi değildir.

**Bozulursa:** Yeni import:
- kartı geriye taşıyabilir
- tamamlanan adedi sıfırlayabilir
- operatör / timestamp / not bilgisini silebilir

---

## 8. HAZIR/TESLİM toplam adet conflict kontrolünü kaldırma

**Kural:** Tamamlanmış kartın kaynak total’inin sessizce değişmesine izin verme.

**Neden:** Tamamlanmış adet ile yeni toplam ayrışır.

**Bozulursa:** Validation veya geçmiş operasyon gerçeği çelişir.

---

## 9. `anahtar = talep_no|stok_no` semantiğini rastgele değiştirme

**Kural:** Import merge identity değişecekse migration / data compatibility düşün.

**Neden:** Mevcut kartları yeni Excel satırlarıyla eşleyen ana business identity budur.

**Bozulursa:** Eski kart “yeni kart” sanılabilir veya duplicate oluşabilir.

---

## 10. `admin_gizli`yi hard delete sanma

**Kural:** “Kart sil” endpoint’i gerçek delete değildir.

**Neden:** Audit ve geri getirme davranışı soft-delete üzerine kurulu.

**Bozulursa:** Yönetim “Gizlenen Kartlar” ve geri getirme akışı anlamsızlaşır; veri kaybı oluşabilir.

---

## 11. Process lock stale logic’ini basitleştirirken yarış durumunu düşün

**Kural:** `remove(lock) → create(lock)` gibi basit yaklaşıma dönme.

**Neden:** Mevcut kod stale lock’u önce unique path’e `os.replace` ile taşıyor ve yeni lock’u `O_EXCL` ile yaratıyor.

**Bozulursa:** İki process aynı anda stale lock reclaim etmeye çalışırken race oluşabilir.

---

## 12. Restore işleminde audit log’u backup’tan geri sarmayı bilinçsizce ekleme

**Kural:** Mevcut tasarım kart state’ini geri alırken audit geçmişini forward-only tutuyor.

**Neden:** Restore event’i de audit trail’de kalıyor.

**Bozulursa:** “Kim ne zaman restore yaptı?” geçmişi eski backup ile yok olabilir.

---

# 15) Mini sözlük

## Kart

Bir PDGM iş / üretim kaydı.

`kartlar.xlsx` içinde bir satır ve uygulamada bir dict olarak temsil edilir.

---

## Dizgi

Kodda production workflow’un aktif üretim aşamasına karşılık gelir.

Durum:

```text
DİZGİDE
```

---

## Workflow

Kartın süreç durumları arasındaki iş akışı:

```text
PLANA ALINDI → DİZGİDE → HAZIR → TESLİM EDİLDİ
```

---

## Snapshot

Kaynak Excel’in COM ile değer bazlı geçici `.xlsx` kopyası.

Parser asıl source workbook yerine bunu okur.

---

## Source of truth

Bir verinin ana / otoritatif kalıcı kaynağı.

Bu sistemde kart operasyon state’i için:

```text
data/kartlar.xlsx
```

---

## Audit log

Kullanıcının yaptığı önemli işlemlerin append ağırlıklı geçmişi.

Ana dosya:

```text
data/islem_logu.xlsx
```

---

## Application log

Teknik log.

Dosya:

```text
data/uygulama.log
```

Audit log ile aynı şey değildir.

---

## Process lock

Aynı `data/` klasörünü ikinci PDGM Python process’inin açmasını engelleyen file-based lock.

Dosya:

```text
data/sunucu.lock
```

---

## Stale lock

Lock dosyası var fakat sahibi process artık çalışmıyorsa kalan eski lock.

Kod bunu kontrollü şekilde reclaim etmeye çalışır.

---

## RLock

Aynı Python process içindeki thread’lerin shared state’e kontrollü erişmesi için kullanılan reentrant mutex.

---

## Atomic write

Bir dosyayı doğrudan üstüne yazmak yerine yeni temp dosya hazırlayıp final adımda replace ederek kullanıcıya yarım dosya gösterme ihtimalini azaltma yaklaşımı.

Bu sistemde çoklu dosya tarafında rollback de eklenmiştir.

---

## Rollback

Bir mutation / write başarısız olduğunda RAM veya dosya state’ini önceki duruma döndürme denemesi.

---

## Soft delete

Kaydı fiziksel olarak silmeyip görünmez işaretlemek.

Bu sistemde:

```python
admin_gizli = 1
```

ile uygulanıyor.

---

## `source_active`

Kartın son kaynak Excel importunda hâlâ bulunup bulunmadığını gösteren bayrak.

`0` olması kartın fiziksel olarak silindiği anlamına gelmez.

---

## `aktif`

Kart modelindeki “Listede” alanıdır.

Mevcut kodda görünürlük hesabının parçalarından biridir.

Bu alanın tüm yaşam döngüsü kullanım amacı hakkında dump’ta ayrıca bir domain dokümantasyonu yok.

---

## `anahtar`

Kartın business identity değeri:

```text
Talep NO|Kart Stok No
```

---

## COM

Windows’ta Microsoft Excel uygulamasını programatik olarak kontrol etmek için kullanılan Component Object Model arayüzü.

Bu projede `pywin32` üzerinden kullanılıyor.

---

## CSRF

Kullanıcının session’ını başka bir sayfanın kötüye kullanarak mutation request’i göndermesini engellemeye yönelik koruma.

---

## Jinja

Flask tarafında HTML oluşturmak için kullanılan template sistemi.

---

## Waitress

Uygulamanın production HTTP server’ı olarak kullanılan Python WSGI sunucusu.

Dump’ta:

```text
threads=8
```

ile çalıştırılıyor.

---

# Bu sistemde en kritik 10 kavram

> ## 1. `kartlar.xlsx` ana operasyon state’idir
> Sistem SQL database yerine Excel’i source of truth olarak kullanıyor.
>
> ## 2. Mimari tek process + çok thread’dir
> Waitress 8 thread kullanıyor; ikinci process `sunucu.lock` ile engelleniyor.
>
> ## 3. `RLock` thread yarışlarını kontrol eder
> Tüm kritik RAM read-modify-write akışları lock altında yürütülür.
>
> ## 4. `_atomik_kart_islemi()` RAM rollback merkezidir
> İşlem, validation ve kart/audit commit’i aynı kalıpta tutulur.
>
> ## 5. `_coklu_yaz()` temp + backup + replace + rollback kullanır
> Excel database olmadığı için dosya bütünlüğünü olabildiğince korumaya çalışır.
>
> ## 6. Workflow kesin olarak dört gerçek state’tir
> `PLANA ALINDI → DİZGİDE → HAZIR → TESLİM EDİLDİ`.
>
> ## 7. `kart_bitir()` otomatik HAZIR yapmaz
> Son adet tamamlandıktan sonra ayrı `/api/hazirla` işlemi gerekir.
>
> ## 8. Kaynak Excel mevcut workflow’u ezmez
> Import plan/source alanlarını günceller, yaşayan üretim state’ini korur.
>
> ## 9. İş identity’si `Talep NO|Kart Stok No` anahtarıdır
> Import merge ve duplicate kontrollerinin temelidir.
>
> ## 10. Yedek restore kart state’ini geri alır, audit geçmişini geri sarmaz
> Restore öncesinde ayrıca koruma backup’ı alınır ve restore işlemi audit’e yeni olay olarak eklenir.

---

# Ek: Dosyalar arası zihinsel model

Yeni bir developer repository’ye ilk kez giriyorsa dosyaları şu sırayla okumak en anlamlı akışı verir:

```text
1. app.py
   ↓
   HTTP / auth / route contract

2. depo.py
   ↓
   gerçek iş kuralları + storage invariants

3. excel_araclari.py
   ↓
   kaynak Excel sınırı + import parser

4. operator.html
   ↓
   operator workflow’un browser tarafı

5. yonetim.html
   ↓
   admin recovery/import/edit akışları

6. base.html
   ↓
   CSRF + shared JS + navigation

7. panel.html / monitor.html / ozet.html
   ↓
   read-only presentation

8. kullanici_yonet.py
   ↓
   operasyonel user administration

9. run_pdgm.bat / yedek_disari_kopyala.bat
   ↓
   Windows operational wrapper

10. stil.css
    ↓
    görünüm ve responsive davranış
```

Bu sıralama kodun bağımlılık yönünü takip eder:

```text
Browser/template
      ↓
app.py
      ↓
depo.py
      ↓
Excel/filesystem

Excel importunda:
app.py
      ↓
excel_araclari.py
      ↓
depo.py
      ↓
Excel/filesystem
```

---

# Son not: dump’tan kesin çıkarılamayan noktalar

Aşağıdaki konularda kod parçaları veya konfigürasyon dosyaları dump’ta bulunmadığı için kesin hüküm verilemez:

- gerçek production hostname
- gerçek LAN subnet / firewall
- gerçek Windows service / Task Scheduler task tanımı
- TLS gerçekten aktif mi
- Caddy gerçekten kullanılıyor mu
- network backup share gerçekten yapılandırılmış mı
- `static/pdgm_logo.png` repository’de gerçekten mevcut mu
- gerçek kaynak Excel örneklerinde hangi formüllerin / external link’lerin bulunduğu
- kullanıcıların gerçek sayısı
- uygulamanın gerçek atölye operasyon prosedürü
- Windows klasör ACL’leri
- antivirus / EDR davranışı
- data klasörünün local disk mi network filesystem mi olduğu

Bu dokümanda bunların hiçbiri varsayılmamıştır.
