# PDGM Flask — Developer Handbook

Bu belge, repository’yi devralacak bir Python developer’ın mimariye, state yönetimine, request flow’una ve kritik iş kurallarına hâkim olmasını hedefler. README değildir; architecture + code reference’tır.

**Source of truth:** Güncel kod (`app.py`, `depo.py`, `excel_araclari.py`, `kullanici_yonet.py`, templates). Satır numaralarına bağımlı olma; fonksiyon adlarıyla referans verilir.

---

## 3.1 Sistem mimarisi

```text
Browser (Jinja + vanilla JS)
        │
        ▼
Waitress (threads=8) ── veya Flask app.run fallback
        │
        ▼
Flask app.py
  ├── Auth / CSRF / session / rate limit
  ├── Routes + JSON APIs
  └── Templates / static
        │
        ▼
depo.py
  ├── RAM: _kartlar, _loglar, _yuklemeler
  ├── threading.RLock (_kilit)
  ├── Workflow + validation
  ├── Atomic Excel persistence
  └── Backup / process lock
        │
        ▼
data/*.xlsx + kullanicilar.json
```

### Import mimarisi

```text
Upload (yonetim/yukle)
        │
        ▼
excel_araclari.excelden_aktar   ← _import_kilidi
        │
        ├── COM snapshot (DispatchEx)
        ├── openpyxl parse (MAKİNE)
        └── depo.excel_import_uygula  ← _kilit
                │
                ▼
        anlik_yedek + _coklu_yaz (3 dosya)
```

**Lock order (doğrulanmış):** `_import_kilidi` → `depo._kilit`. Ters yol yoktur (`depo` → `excel_araclari` çağırmaz). AB-BA deadlock yok.

---

## 3.2 Repository tree

```text
pdgm_flask/
├── app.py                 HTTP katmanı, auth, özet, Waitress
├── depo.py                Domain + persistence + lock
├── excel_araclari.py      COM snapshot, parse, rapor workbook
├── kullanici_yonet.py     Offline kullanıcı CLI
├── requirements.txt       Pin’li bağımlılıklar
├── .env / .env.example    Config
├── .gitignore
├── run_pdgm.bat           Windows başlatıcı
├── yedek_disari_kopyala.bat
├── templates/             Jinja
├── static/stil.css
├── data/                  Runtime state (gitignore)
└── docs/                  Bu dokümanlar
```

---

## 3.3 Application startup

```text
import app
  ├── load_dotenv
  ├── Flask(__name__)
  ├── depo.process_kilidi_al()     # module import (app.py üst)
  ├── secret: data/gizli.key
  ├── _kullanicilari_yukle()       # yoksa .env → hash → json
  ├── _gunluk_dosya_logu_kur()
  ├── try: depo.kur()              # kartlar hard-fail; log quarantine
  │     except VeriDogrulamaHatasi → BASLATMA_HATASI.txt + SystemExit
  └── route registration

python app.py / calistir()
  ├── process_kilidi_al()          # aynı PID → no-op
  ├── depo.kur()                   # diskten taze reload (kasıtlı 2. çağrı)
  └── waitress serve(PDGM_BIND, PDGM_PORT, threads=8)
```

**Çift `kur()`:** Silinmemelidir. Import-time boot + `calistir` öncesi taze yükleme. İkinci çağrı startup’ta zararsızdır; runtime ortasında çağrılırsa unsaved RAM silinir — public API olarak kötüye kullanma.

---

## 3.4 `app.py` — gruplar

### Config / secret

| Sembol | Açıklama |
|--------|----------|
| `SUNUCU_PORTU` | `PDGM_PORT` default 5001 |
| `DINLENEN_ADRES` | `PDGM_BIND` default `0.0.0.0` |
| `_anahtar()` | `data/gizli.key` üret/oku; session secret |

### Users

| Fonksiyon | Açıklama |
|-----------|----------|
| `_kullanicilari_yukle` | İlk bootstrap veya mevcut json oku |
| `_kullanicilari_al` | mtime cache; JSON hata → son iyi cache |

### Auth decorators

| Fonksiyon | Açıklama |
|-----------|----------|
| `_oturum_kullanici_kontrol` | before_request: aktif/rol tazele |
| `yetki(*roller)` | Session + rol gate |
| `csrf_koru` | Header `X-CSRF-Token` veya form `_csrf_token` |
| `_guvenli_devam_hedefi` | Open redirect koruması |

### Login rate limit

In-memory `_giris_basarisiz` dict; IP başına 8 deneme / 5 dk. Proxy arkasında `remote_addr` tek IP’ye çökebilir (ProxyFix yok — bilinçli).

### Error handlers

`PermissionError` → Excel kilitli mesajı (geniş yakalama; trade-off kabul edilmiş).

### Özet

`ozet_hesapla`, `_donem_ozeti`, `_teslim_tarihi` — Flask’tan bağımsız hesap; route sadece render.

---

## 3.5 Route reference

| Route | Method | Role | Handler | Mutasyon | CSRF |
|-------|--------|------|---------|----------|------|
| `/giris` | GET/POST | public | `giris` | session | hayır |
| `/cikis` | POST | all | `cikis` | session+log | evet |
| `/` | GET | — | `ana` | hayır | — |
| `/panel` | GET | all | `panel` | hayır | — |
| `/monitor` | GET | all | `monitor` | hayır | — |
| `/operator` | GET | admin,operator | `operator` | hayır | — |
| `/yonetim` | GET | admin | `yonetim` | hayır | — |
| `/ozet` | GET | all | `ozet` | hayır | — |
| `/api/veriler` | GET | all | `api_veriler` | hayır | — |
| `/api/basla` | POST | admin,operator | `api_basla` | evet | evet |
| `/api/bitir` | POST | admin,operator | `api_bitir` | evet | evet |
| `/api/hazirla` | POST | admin,operator | `api_hazirla` | evet | evet |
| `/api/teslim-et` | POST | admin,operator | `api_teslim_et` | evet | evet |
| `/api/not` | POST | admin,operator | `api_not` | evet | evet |
| `/api/admin/kart-ekle` | POST | admin | `api_kart_ekle` | evet | evet |
| `/api/admin/duzenle` | POST | admin | `api_duzenle` | evet | evet |
| `/api/admin/kart-sil` | POST | admin | `api_kart_sil` | evet | evet |
| `/yonetim/kart-geri-getir` | POST | admin | `kart_geri_getir` | evet | evet |
| `/yonetim/yedek-geri-yukle` | POST | admin | `yedek_geri_yukle` | evet | evet |
| `/yonetim/yukle` | POST | admin | `yukle` | evet | evet |
| `/yonetim/yeniden-oku` | POST | admin | `yeniden_oku` | evet | evet |
| `/yonetim/kayit-dosyasi/<hangi>` | GET | admin | `kayit_dosyasi` | hayır* | — |
| `/yonetim/rapor` | GET | admin | `rapor_indir` | hayır* | — |

\*Excel audit yazılmaz; yalnız `app.logger`.

### API örnekleri

```json
POST /api/basla
{ "kart_id": 12, "adet": 50, "not": "opsiyonel" }

→ { "tamam": true, "mesaj": "...", "kart": { ...kart_gorunumu } }
```

```json
POST /api/bitir
{ "kart_id": 12, "adet": 10, "not": "" }

→ { "tamam": true, "uretim_bitti": true|false, "mesaj": "...", "kart": {...} }
```

Hatalar `_api_kart_islemi` üzerinden: 404 `KartBulunamadi`, 409 `IsKuralHatasi`, 400 validation.

---

## 3.6 `depo.py` — global state

| Global | İçerik |
|--------|--------|
| `_kartlar` | `list[dict]` kart kayıtları |
| `_loglar` | audit satırları |
| `_yuklemeler` | import geçmişi |
| `_kilit` | `threading.RLock` — tüm public mutasyon/okuma yolları |

Yükleme: `kur()`. Yazma: `_kart_log_commit` / `_coklu_yaz` / `log_ekle`.

---

## 3.7 Data model — `KART_ALANLARI`

| Excel başlık | Alan | Tip (pratik) | Kaynak |
|--------------|------|--------------|--------|
| ID | id | int | sistem |
| Sıra | sira | int? | Excel/plan |
| Talep NO | talep_no | str | Excel |
| Kart Stok No | stok_no | str | Excel |
| Talep Sahibi | talep_sahibi | str | Excel |
| Toplam Adet | toplam_adet | int | Excel/plan |
| Adet Metni | adet_metin | str | Excel |
| Plan Haftası | plan_hafta | str | Excel |
| Plan Başlangıç | plan_baslama | date str | Excel |
| Plan Teslim | plan_teslim | date str | Excel |
| Gerçekleşen Teslim | gerceklesen_teslim | date str | workflow/teslim |
| Excel Durumu | excel_durum | str | ham Excel |
| PCB | pcb | str | Excel |
| Durum | durum | enum/None | workflow |
| Başlangıç Adedi | baslangic_adet | int | workflow |
| Tamamlanan Adet | tamamlanan_adet | int | workflow |
| Başlama/Bitiş/Teslim Zamanı | *_zamani | datetime str | workflow |
| Operatör | operator | str | workflow |
| Not | aciklama | str | not |
| Son Güncelleme | guncelleme | str | sistem |
| Listede | aktif | 0/1 | sistem |
| Kaynakta Aktif | source_active | 0/1 | import |
| Admin Gizli | admin_gizli | 0/1 | admin |
| Kaynak | kaynak | str | EXCEL/MANUEL |
| Anahtar | anahtar | str | `talep\|stok` |

`ZORUNLU_KART_ALANLARI`: id, talep_no, stok_no, toplam_adet, tamamlanan_adet, anahtar.

---

## 3.8 Derived / view fields

`kart_gorunumu` / `durum_bilgisi` üretir; Excel’e yazılmaz:

| Alan | Anlam |
|------|-------|
| kalan_adet | toplam − tamamlanan |
| adet_yuzde | ilerleme % |
| renk / rozet | UI sınıfı (iyi/uyari/kotu/hazir…) |
| gorunur | operasyon görünürlüğü |
| is_durumu | durum veya `DURUMU EKSİK` |
| sapma, zaman_yuzde, … | plan sapması |

Frontend `kart` JSON’unu alır ama şu an DOM patch yapmaz; `location.reload` kullanır.

---

## 3.9 Normalization pipeline

```text
Excel/cell veya form
    │
    ▼
_temiz_metin / tarih_coz / _sayi / durum_coz (import)
    │
    ▼
_durum_normalize   # unknown → None (boot-safe)
    │
    ▼
_kart_normalize
    │
    ▼
_kart_dogrula      # tek kart kuralları
    │
    ▼
_kart_listesi_dogrula  # id/anahtar tekilliği
```

---

## 3.10 Workflow state machine

```text
PLANA_ALINDI
    │ kart_baslat (POST /api/basla)
    ▼
DIZGIDE
    │ kart_bitir (POST /api/bitir)  — partial veya complete; durum aynı kalır
    │ kart_hazirla — yalnız tamamlanan==toplam
    ▼
HAZIR
    │ kart_teslim_et (POST /api/teslim-et)
    ▼
TESLIM_EDILDI
```

| Transition | Precondition | Mutation özeti | Log |
|------------|--------------|----------------|-----|
| baslat | PLANA, operasyonda görünür | durum=DİZGİDE, baslangic_adet | DİZGİYE ALINDI |
| bitir | DİZGİDE, adet≤kalan | tamamlanan+=adet; otomatik HAZIR yok | KISMİ ÜRETİM / ÜRETİM ADEDİ TAMAMLANDI |
| hazirla | DİZGİDE, adet eşit | durum=HAZIR | HAZIR OLARAK İŞARETLENDİ |
| teslim | HAZIR | durum=TESLİM, tarihler | TESLİM EDİLDİ |
| not | görünür | aciklama | NOT GÜNCELLENDİ |

Hepsi `_atomik_kart_islemi` üzerinden.

---

## 3.11 Transaction / rollback

### Kart işlemleri

```text
_atomik_kart_islemi(islem)
  eski_kartlar = deepcopy(_kartlar)
  log_sayisi = len(_loglar)          # append-only varsayımı
  try:
      sonuc = islem()                # in-place kart + append log
      _kart_listesi_dogrula
      _kart_log_commit()             # günlük yedek + _coklu_yaz kart+log
      return sonuc
  except:
      _kartlar = eski_kartlar
      del _loglar[log_sayisi:]
      raise
```

**Kart deepcopy zorunlu** (in-place update). Log `del` yalnızca bu path’te append olduğu için güvenli.

### `_coklu_yaz`

```text
1. Her hedef için _temp_yaz (+ fsync)
2. Mevcut dosyaların .txn.bak kopyası
3. os.replace(temp → hedef) sırayla
4. Hata: degisen listesini ters sırada restore
5. finally: temp temizle; yalnız commit_basarili ise bak sil
```

Başarısız rollback’te `.txn.bak` **bilinçli bırakılır**.

### `log_ekle`

Archive-first: arşiv workbook yaz → RAM kırp → ana log yaz. Hata → full deepcopy rollback. Orphan arşiv dosyası mümkün (kayıp değil, duplikasyon).

---

## 3.12 Locking / concurrency

| Katman | Mekanizma | Amaç |
|--------|-----------|------|
| Thread | `_kilit` RLock | RAM↔disk tutarlılığı |
| Process | `sunucu.lock` O_EXCL + stale reclaim via `os.replace` | Tek Python process |
| Import | `_import_kilidi` | Tek COM/import |

**Stale reclaim:** Canlı python PID → refuse. Ölü veya non-python → `os.replace(lock, lock.stale.uuid)` sonra O_EXCL. İki process yarışırsa yalnız biri replace kazanır.

**Yapma:** Yazmayı `_kilit` dışına çıkarma.

---

## 3.13 Persistence dosyaları

| Dosya | Rol | Yazma |
|-------|-----|-------|
| kartlar.xlsx | SoT | her kart commit |
| islem_logu.xlsx | audit | commit + log_ekle |
| yuklemeler.xlsx | import history | import commit |

Okuma: `openpyxl` read_only/data_only. Yazma: stilli workbook + filter + freeze.

---

## 3.14 Excel import subsystem

### Neden COM?

Formül sonuçları ve hidden kolon filtresi için Excel otomasyonu. openpyxl tek başına calculated values vermez.

### Security settings (`excel_deger_snapshot_olustur`)

| Ayar | Değer | Neden |
|------|-------|-------|
| AutomationSecurity | 3 | Makro kapalı |
| UpdateLinks | 0 | Dış link yok |
| EnableEvents | False | Event makro yok |
| DisplayAlerts | False | UI blok yok |
| ReadOnly | True | Kaynak bozulmasın |
| Visible | False | Headless |

### Lifecycle

CoInitialize → DispatchEx → Open → copy Value2 → SaveAs → Close/Quit → refs=None → gc.collect → CoUninitialize → fail ise temp sil.

### Parse

- Sayfa: `MAKİNE`
- Header: Talep NO / Stok No içeren satır
- `BASLIK_ESLESME` ile kolon map
- Hidden kolonlar atlanır
- `durum_coz`: bilinmeyen → None + uyarı sayacı

---

## 3.15 Workflow preservation (import)

```text
workflow = { durum, adetler, zamanlar, operator, aciklama, gerceklesen_teslim }
mevcut.update(plan)      # Excel plan alanları
mevcut.update(workflow)  # operatör state kazanır
```

**Sıra kritik.** Tersine çevirmek Excel’in boş alanlarıyla üretim adedini siler.

`plan` dict’inde `gerceklesen_teslim` kasıtlı yoktur (parser ayrı tutar).

Conflict:

- `yeni_toplam < tamamlanan` → abort
- HAZIR/TESLİM ve `yeni_toplam != tamamlanan` → abort

---

## 3.16 Identity / idempotency

```python
anahtar = f"{talep_no}|{stok_no}"
```

Import aynı anahtarı günceller; yeniden yükleme duplicate üretmez. Değiştirmek mevcut dosyaları ve eşleşmeyi kırar.

---

## 3.17 Backup architecture

| Tür | Tetik | İsim |
|-----|-------|------|
| Günlük | ilk yazma/gün | `YYYYMMDD_kartlar.xlsx` |
| Anlık | import/restore/manuel etiket | `YYYYMMDD_HHMMSS_etiket/` |
| Retention | `yedekleri_buda` | 30 anlık pattern, 90g günlük |
| Offsite | bat + Task Scheduler | robocopy |

Restore: yalnız kartlar; önce `geri_yukleme_oncesi` anlık; log geri sarılmaz.

---

## 3.18 Recovery scenarios

| Senaryo | Code path |
|---------|-----------|
| Corrupt kartlar | `kur` raise → `_baslatma_hatasi_bildir` → exit |
| Corrupt log/yükleme | `_bozuk_dosyayi_kenara_al` → boş + recreate |
| Excel locked | `PermissionError` handler |
| Stale lock | reclaim veya refuse |
| COM fail | `ExcelAktarimHatasi`; temp silinir |
| Write fail | RAM + disk rollback |
| Disk full | OSError; bak kalabilir |
| Bad users JSON | cache fallback veya 500 |

---

## 3.19 Auth internals

```text
kullanicilar.json:
  user: { sifre_hash, rol, ad, aktif }
```

- Hash: Werkzeug `generate_password_hash`
- Login: `check_password_hash`
- Cache mtime; CLI değişikliği restart’sız
- before_request rol/aktif tazeler
- Inactive → session clear

---

## 3.20 Security model — kaldırırsan ne olur?

| Kontrol | Kaldırılırsa |
|---------|--------------|
| CSRF | Cross-site POST mutasyon |
| yetki | Yetkisiz ekran/API |
| `_oturum_kullanici_kontrol` | Pasif kullanıcı işlem yapar |
| COM AutomationSecurity | Makro riski |
| path commonpath/basename | Path traversal |
| process lock | Çift writer corruption |
| secure_filename+uuid | Overwrite / path issues |

---

## 3.21 `kullanici_yonet.py`

```text
listele | ekle | parola | rol | aktif | pasif
```

Parola ≥8, çift giriş. Son aktif admin koruması. `yaz`: fsync + replace.

---

## 3.22 Templates

| Template | Route | Rol |
|----------|-------|-----|
| base.html | — | CSRF meta, pdgmFetch, toast |
| giris.html | /giris | public |
| panel.html | /panel | all |
| operator.html | /operator | admin/operator |
| monitor.html | /monitor | all |
| yonetim.html | /yonetim | admin |
| ozet.html | /ozet | all |
| yetkisiz.html | 403 | — |

Context: kart listeleri, sayac, csrf_token, session ad/rol.

---

## 3.23 Frontend JS (özet)

`base.html → pdgmFetch`: JSON + CSRF header; 401/403 toast.

`operator.html`: filtre, sessionStorage, confirm, aksiyon → toast → `location.reload`.

`monitor.html`: reload timer + carousel.

`yonetim.html`: formlar + fetch düzenleme/gizleme → reload.

---

## 3.24 CSS

`stil.css`: CSS variables (`--iyi`, `--uyari`, `--kotu`, `--hazir`), layout, kartlar, operatör, monitör, tablolar, dialog, responsive. Framework yok.

---

## 3.25 Exceptions

| Exception | Üreten | Yakalanan | Sonuç |
|-----------|--------|-----------|-------|
| KartBulunamadi | depo | `_api_kart_islemi` | 404 |
| IsKuralHatasi | depo/import | API/flash | 409 / hata mesajı |
| VeriDogrulamaHatasi | depo oku/kur | startup/yeniden-oku | exit veya flash |
| ExcelAktarimHatasi | excel_araclari | yukle | flash |
| DepoHatasi | taban | — | — |

---

## 3.26 Logging / audit

| Kanal | İçerik |
|-------|--------|
| islem_logu.xlsx | Başarılı iş aksiyonları, giriş/çıkış |
| uygulama.log | Hatalı giriş, rapor indirme, exceptions |
| yuklemeler.xlsx | Import özeti |
| yedekler/ | Snapshot’lar |

---

## 3.27 Configuration

| Variable | Default | Required | Açıklama |
|----------|---------|----------|----------|
| PDGM_PORT | 5001 | hayır | Waitress port |
| PDGM_BIND | 0.0.0.0 | hayır | Bind address |
| PDGM_HTTPS | 0 | hayır | Secure cookie |
| PDGM_ADMIN_PASSWORD | — | yalnız ilk bootstrap | |
| PDGM_OPERATOR_PASSWORD | — | yalnız ilk bootstrap | |
| PDGM_VIEWER_PASSWORD | — | yalnız ilk bootstrap | |

---

## 3.28 Deployment

- Python 3.8+ (venv), `pip install -r requirements.txt`
- Windows + Microsoft Excel (COM import için)
- Interactive user session (Office automation; Service önerilmez)
- Waitress production WSGI
- İsteğe bağlı Caddy TLS + `PDGM_BIND=127.0.0.1` + `PDGM_HTTPS=1`
- Task Scheduler: `run_pdgm.bat`, `yedek_disari_kopyala.bat`

---

## 3.29 Developer setup

```text
1. Repo klonla / kopyala
2. python -m venv .venv && source .venv/bin/activate  # Windows: .venv\Scripts\activate
3. pip install -r requirements.txt
4. .env.example → .env; bootstrap parolaları doldur
5. python app.py
6. İlk açılıştan sonra .env parolalarını SİL
7. http://127.0.0.1:5001/giris
```

Mac/Linux: uygulama açılır; Excel import COM olmadan fail eder.

---

## 3.30 Safe development rules

```text
Kural: _coklu_yaz sırasını değiştirme
Neden: atomiklik
Bozulursa: yarım commit / kayıp bak

Kural: Yazmayı _kilit dışına alma
Neden: RAM/disk ayrışması
Bozulursa: race corruption

Kural: Kart deepcopy’yi kaldırma
Neden: in-place mutation
Bozulursa: bozuk rollback

Kural: import plan→workflow sırası
Neden: workflow preservation
Bozulursa: adet/durum silinir

Kural: HAZIR/TESLİM conflict’i gevşetme
Neden: sessiz veri bozulması
Bozulursa: tutarsız kartlar

Kural: anahtar şemasını değiştirme
Neden: idempotent import
Bozulursa: duplicate/yanlış merge

Kural: KART_ALANLARI başlık rename etme
Neden: mevcut xlsx okunamaz
Bozulursa: boot fail

Kural: COM security bayraklarını kaldırma
Neden: makro/link yüzeyi
Bozulursa: güvenlik açığı

Kural: CSRF / yetki / before_request kaldırma
Neden: authz
Bozulursa: yetkisiz mutasyon

Kural: path whitelist/commonpath sadeleştirme
Neden: traversal
Bozulursa: dosya okuma/yazma saldırısı
```

---

## 3.31 Extension guide

| İstek | Dokunulacak yerler |
|-------|-------------------|
| Yeni kart alanı | `KART_ALANLARI`, normalize, template, belki BASLIK_ESLESME |
| Yeni workflow state | sabitler, geçiş fonksiyonları, UI butonları, import eşleme — yüksek risk |
| Yeni API | `app.py` route + `depo` fonksiyon + CSRF/yetki |
| Yeni Excel kolonu | `BASLIK_ESLESME` + plan dict + alan listesi |
| Yeni rol | `ROLLER`, `yetki` çağrıları, CLI, bootstrap |
| Yeni KPI | `ozet_hesapla` + `ozet.html` |

---

## 3.32 Testing / smoke

Otomatik test paketi yok. Manuel minimum:

```text
login (3 rol)
workflow PLANA→DİZGİ→kısmi→tam→HAZIR→TESLİM
import + conflict reject
yedek restore
ikinci process refuse
stale lock reclaim
COM import (Windows)
rapor/kayıt indirme
CLI parola/rol
```

---

## 3.33 Known limitations

- Excel persistence ≈ binlerce aktif kartta yavaşlar
- Tek process; yatay ölçek yok
- Windows + kurulu Excel (import)
- COM interactive session gerektirir
- Concurrent write thrput düşük (global RLock)
- UI her aksiyonda full reload

Bunlar bug değil, mimari kısıt.

---

## 3.34 Debugging guide

| Sorun | İlk bakılacak |
|-------|----------------|
| Login olmuyor | kullanicilar.json, hash, rate limit, uygulama.log |
| Kart görünmüyor | durum None? admin_gizli? source_active? `_operasyonda_gorunur_mu` |
| Import hata | COM/Excel kurulu mu? MAKİNE sayfası? conflict mesajı |
| EXCEL.EXE kaldı | snapshot finally; import kilidi; manuel Task Manager |
| İkinci server | sunucu.lock PID |
| Kart kaydı gitmiyor | PermissionError (Excel açık); uygulama.log |
| Açılmıyor | BASLATMA_HATASI.txt; kartlar.xlsx; lock |

---

## 3.35 End-to-end traces

### Adet Bitir

```text
operator.html → POST /api/bitir + CSRF
→ api_bitir → depo.kart_bitir
→ _atomik_kart_islemi → mutation + log append
→ _kart_log_commit → _coklu_yaz
→ JSON {kart, uretim_bitti, mesaj}
→ toast + location.reload
```

### Login

```text
giris POST → rate limit → check_password_hash
→ session + CSRF token → log_ekle GİRİŞ (Excel)
→ redirect
```

### Excel import

```text
yonetim/yukle → save file
→ excelden_aktar (_import_kilidi)
→ COM snapshot → parse → excel_import_uygula (_kilit)
→ anlik_yedek → mutate → _coklu_yaz ×3
→ flash + prune uploads
```

### Backup restore

```text
POST yedek-geri-yukle
→ yedekten_geri_yukle
→ path guard → anlik koruma → swap _kartlar → log append → yaz
```

### Dizgiye Al

```text
POST /api/basla → kart_baslat → PLANA→DİZGİDE → commit → reload
```

---

## 3.36 Function index (kritik)

| Fonksiyon | Dosya | Amaç |
|-----------|-------|------|
| process_kilidi_al | depo | Process exclusivity |
| kur | depo | Boot load |
| _coklu_yaz | depo | Atomic multi-write |
| _atomik_kart_islemi | depo | RAM txn |
| kart_baslat/bitir/hazirla/teslim_et | depo | Workflow |
| excel_import_uygula | depo | Import commit |
| excelden_aktar | excel_araclari | Import entry |
| excel_deger_snapshot_olustur | excel_araclari | COM |
| yetki / csrf_koru | app | Authz |
| _kullanicilari_al | app | User cache |
| ozet_hesapla | app | KPI |
| parola_degistir / rol_degistir | kullanici_yonet | CLI |

---

## 3.37 Glossary

| Terim | Anlam |
|-------|--------|
| Kart | Tek üretim iş kaydı |
| Anahtar | talep_no\|stok_no |
| Dizgi | Üretim aşaması (DİZGİDE) |
| Snapshot | COM values-only geçici xlsx |
| Source of truth | kartlar.xlsx |
| Anlık yedek | Timestamp’li klasör kopyası |
| Audit log | islem_logu.xlsx |
| Workflow preservation | Import’ta operatör state’inin korunması |
| Stale lock | Ölü process’ten kalan sunucu.lock |

---

## Self-check

Bu handbook ile yeni bir developer: startup sırasını, lock modelini, `_coklu_yaz` / `_atomik_kart_islemi` semantiğini, import merge kurallarını, route/API yüzeyini ve “dokunulmaması gereken” invariant’ları görebilir. Eksik kalan otomatik test suite’i bilinen bir boşluktur; smoke listesi yeterlidir.

---

## 3.38 `kart_baslat` / `kart_bitir` / `kart_hazirla` / `kart_teslim_et` sözleşmeleri

### `kart_baslat(kart_id, kullanici, rol)`

**Önkoşul:** durum == `PLANA ALINDI`  
**Mutasyon:** durum=`DİZGİDE`, `baslama_zamani`, `operator`, `guncelleme`  
**Log:** `DİZGİYE ALINDI`  
**API:** `POST /api/basla`  
**Exception:** `KartBulunamadi`, `IsKuralHatasi`

### `kart_bitir(kart_id, adet, kullanici, rol, aciklama="")`

**Önkoşul:** durum == `DİZGİDE`; `1 <= adet <= kalan`  
**Mutasyon:** `tamamlanan_adet += adet`; durum DİZGİDE kalır; kalan 0 ise `bitis_zamani` set edilebilir  
**Log:** kısmi veya `ÜRETİM ADEDİ TAMAMLANDI`  
**Kritik:** HAZIR’a otomatik geçmez  
**API:** `POST /api/bitir` body `{kart_id, adet}`

### `kart_hazirla(...)`

**Önkoşul:** DİZGİDE **ve** `tamamlanan_adet == toplam_adet`  
**Mutasyon:** durum=`HAZIR`  
**API:** `POST /api/hazirla`

### `kart_teslim_et(...)`

**Önkoşul:** durum == `HAZIR`  
**Mutasyon:** `TESLİM EDİLDİ`, `gerceklesen_teslim=bugun()`, `teslim_zamani`  
**API:** `POST /api/teslim-et`

Hepsi `_atomik_kart_islemi` ile persist edilir.

---

## 3.39 `admin_kart_duzenle` durum yan etkileri

Admin durum değişince zaman alanları yeniden düzenlenir (özet):

| Yeni durum | Yan etki (özet) |
|------------|-----------------|
| PLANA ALINDI | tamamlanan=0, baslangic_adet=0, baslama/bitis/teslim=None, gerceklesen=None |
| DİZGİDE | baslama yoksa set; teslim temizlenebilir; adet tutarlılığı korunur |
| HAZIR | tamamlanan=toplam zorunlu mantığı; bitis set |
| TESLİM EDİLDİ | gerceklesen/teslim set; tamamlanan=toplam |

Plan tarihleri `plan_hafta`, `plan_baslama`, `plan_teslim` opsiyonel form alanlarından gelir. `None` gönderilirse mevcut değer korunur; boş string temizlenebilir (`_tarih_form_degeri`).

---

## 3.40 View-model: `kart_gorunumu` ve `durum_bilgisi`

Persistence alanları Excel’dedir. UI’ya giden dict şunları **ekler**:

| Alan | Kaynak |
|------|--------|
| `kalan_adet` | toplam − tamamlanan |
| `adet_yuzde` | yüzde |
| `renk`, `rozet`, `sapma`, `kalan`, `zaman_yuzde`, `plan_gun` | `durum_bilgisi()` |
| `durum_bilgisi` | nested dict (şablonlarda kullanılır) |

Bunları Excel’e yazmaya çalışma; `_kart_normalize` / yazım yolu bunları persistence şemasına koymaz.

---

## 3.41 Log satırı şeması (`_log_kaydi`)

Tipik alanlar:

- zaman
- kullanici
- rol
- islem
- talep_no
- stok_no
- adet
- aciklama

`log_ekle` append-only yazım + RAM truncate ile çalışır. Başarısız login gibi gürültülü olaylar artık Excel log’a değil `app.logger`’a gider (DoS / şişme önlemi).

---

## 3.42 Exception haritası (geniş)

| Exception | Üreten | HTTP (tipik) |
|-----------|--------|--------------|
| `KartBulunamadi` | depo kart API | 404 JSON |
| `IsKuralHatasi` | workflow/validation | 400 JSON |
| `ExcelAktarimHatasi` | excel_araclari / import | flash + yönetim redirect |
| `DepoHatasi` / IOError | persistence | 500 / startup fail |
| `ProcessKilitHatasi` | `process_kilidi_al` | process exit / mesaj |

API hata gövdesi genelde `{"ok": false, "mesaj": "..."}` şeklindedir (`app.py` error helpers).

---

## 3.43 CSRF ve session detayı

- Session cookie: Flask secret key
- CSRF token: session’da; form hidden + `X-CSRFToken` header
- `csrf_koru` mutating method’larda zorunlu
- `SESSION_COOKIE_HTTPONLY`, `SESSION_COOKIE_SAMESITE=Lax`
- `PDGM_HTTPS=1` ise `SESSION_COOKIE_SECURE`

CSRF’yi kaldırmak CSRF saldırısına kapı açar; SPA refactor’da bile token taşımayı unutma.

---

## 3.44 Auth cache / mtime reload

Kullanıcı dosyası `data/kullanicilar.json`:

1. İlk yüklemede hash’li kullanıcılar belleğe alınır
2. Dosya mtime değişince yeniden okunur
3. Session’daki rol, her request’te dosyadaki güncel rolle hizalanabilir (`yetki` / before_request yolu)

CLI ile rol değişince web process restart gerekmez (mtime path çalışıyorsa). Parola değişince eski session hâlâ açık kalabilir; çıkış + yeniden giriş beklenir.

---

## 3.45 Import conflict — kod seviyesi

`depo.excel_import_uygula` içinde (özet mantık):

```text
for each Excel row:
  anahtar = talep|stok
  if anahtar in mevcut:
    if mevcut.durum in {HAZIR, TESLİM} and qty conflict:
      abort entire import
    plan_fields = Excel plan columns
    workflow_fields = mevcut operator state
    mevcut.update(plan_fields)
    mevcut.update(workflow_fields)   # workflow wins
  else:
    create new card PLANA ALINDI (or normalized status)
```

Abort atomic’tir: kısmi merge diske yazılmaz (transaction path başarısız kalır / RAM rollback).

---

## 3.46 `_coklu_yaz` sırası (dokunulmaz)

```text
1. Her hedef için temp workbook yaz
2. fsync (_diske_zorla)
3. Orijinaller .txn.bak
4. os.replace(temp → final) hepsi
5. Başarı → .txn.bak temizlenebilir / tutulma politikası
6. Hata → reverse order restore from .txn.bak
7. RAM snapshot restore (_atomik_kart_islemi üst katmanı)
```

Sıra değiştirmek split-brain (kartlar yeni, log eski) riski yaratır.

---

## 3.47 Lock order (deadlock)

```text
İzinli:
  _import_kilidi  →  depo._kilit

Yasak:
  depo._kilit tutulurken _import_kilidi almak
```

Process lock (`sunucu.lock`) ayrı katmandır; thread RLock ile karıştırma. Stale reclaim **`os.replace`** ile atomik olmalıdır; `remove`+`O_EXCL` yarışı iki owner üretebilir.

---

## 3.48 COM lifecycle checklist

```text
CoInitialize
→ DispatchEx Excel.Application
→ AutomationSecurity / UpdateLinks / EnableEvents / DisplayAlerts / ReadOnly
→ Open workbook
→ UsedRange.Value → values workbook kaydet
→ Close
→ Quit
→ Release COM refs (None)
→ gc.collect()
→ CoUninitialize
→ orphan temp sil
```

Import aynı anda tek (`_import_kilidi`). macOS’ta pywin32 yoksa soft-fail / açık hata.

---

## 3.49 Frontend–backend sözleşme

Operatör JS (`operator.html`):

- `pdgmFetch(url, {method, body})` → JSON + CSRF header
- Başarıda kart DOM güncellemesi veya `location.reload`
- Filtre `sessionStorage` key: `pdgm-op-filtre`
- Bitir sonrası opsiyonel hazirla zinciri

Backend kart dict alan adları HTML `data-*` ile eşleşmeli; alan rename = frontend kırılır.

---

## 3.50 CSS mantıksal harita (`static/stil.css`)

| Bölüm | Örnek sınıflar |
|-------|----------------|
| Tokens | `:root` değişkenleri |
| Shell | nav, layout, sayfa |
| Kart | `.kart`, durum renkleri |
| Operator | `.operator-*`, filtre |
| Monitor | büyük tipografi, carousel |
| Forms / tables | yönetim |
| Dialog / toast | bildirim |
| Responsive | media queries |

Yeni durum rengi eklerken hem `durum_bilgisi` hem CSS token’ını güncelle.

---

## 3.51 End-to-end: Login

```text
giris.html POST /giris
→ rate limit check
→ check_password_hash
→ session[kullanici,rol]
→ redirect role home
fail → app.logger (Excel audit değil)
```

## 3.52 End-to-end: Dizgiye Al

```text
operator button
→ POST /api/basla {kart_id}
→ yetki(operator|admin) + csrf
→ depo.kart_baslat
→ _atomik_kart_islemi
→ JSON kart_gorunumu
→ DOM update
```

## 3.53 End-to-end: Excel import

```text
yonetim multipart upload
→ secure filename → uploads/
→ anlık yedek
→ excel_araclari.excelden_aktar (COM snapshot)
→ parse/validate
→ depo.excel_import_uygula
→ success → upload history + prune old uploads
→ flash sonuç
```

## 3.54 End-to-end: Backup restore

```text
yonetim restore form
→ path validate (yedekler/ altında)
→ current snapshot backup
→ replace kartlar.xlsx from chosen backup
→ depo reload
→ audit log entry
```

## 3.55 End-to-end: Admin edit

```text
yonetim edit modal/form
→ POST yönetim endpoint
→ depo.admin_kart_duzenle
→ durum yan etkileri
→ atomic persist
→ redirect flash
```

---

## 3.56 Function index (geniş)

| Fonksiyon | Dosya | Amaç |
|-----------|-------|------|
| `kur` | depo | startup load + lock |
| `process_kilidi_al` | depo | single-process |
| `_coklu_yaz` | depo | multi-file atomic write |
| `_atomik_kart_islemi` | depo | RAM+disk txn |
| `_kart_normalize` / `_kart_dogrula` | depo | schema |
| `kart_baslat/bitir/hazirla/teslim_et` | depo | workflow |
| `excel_import_uygula` | depo | merge |
| `yedek_al` / `yedekten_geri_yukle` | depo | backup |
| `kartlari_getir` / `kart_gorunumu` | depo | read API |
| `excelden_aktar` | excel_araclari | COM import |
| `yetki` / `csrf_koru` | app | security |
| `api_bitir` vb. | app | HTTP API |
| CLI `liste/ekle/parola/rol/...` | kullanici_yonet | users |

---

## 3.57 Manuel regression checklist (kod değiştirmeden)

```text
[ ] login admin/operator/gozlemci
[ ] yanlış parola rate limit
[ ] CSRF’siz POST reddi
[ ] PLANA→DİZGİ→partial→complete→HAZIR→TESLİM
[ ] bitir sonrası otomatik HAZIR olmamalı
[ ] HAZIR kartta qty conflict import abort
[ ] gizle / geri getir
[ ] yedek al / restore
[ ] ikinci process lock reddi
[ ] stale lock reclaim (ölü PID)
[ ] COM sonrası EXCEL.EXE kalmamalı (Windows)
[ ] monitor read-only
[ ] ozet KPI sayfası
[ ] rapor indirme
[ ] kullanıcı CLI parola/rol/son admin koruması
```

---

## 3.58 Bilinen mimari sınırlar

- Excel SoT → satır sayısı büyüdükçe yavaşlar
- Tek process zorunlu
- Windows + kurulu Excel + interactive session (COM)
- Horizontal scale yok
- Concurrent write kuyruğu yok; RLock serial eder
- Otomatik test suite yok
- UI çoğu aksiyonda full reload

Bunlar bug değil; seçilmiş trade-off.

---

## 3.59 Debugging hızlı tablo

| Sorun | İlk bakılacak yer |
|-------|-------------------|
| Login olmuyor | `kullanicilar.json`, hash, rate limit, `uygulama.log` |
| Kart görünmüyor | `admin_gizli`, durum, filtre JS |
| Import hata | COM/Excel kurulu mu, `_import_kilidi`, flash mesaj |
| EXCEL.EXE kalıyor | `excel_araclari` cleanup, orphan temp |
| İkinci sunucu açılmıyor | `data/sunucu.lock`, stale reclaim |
| Kart update yok | CSRF, rol, `IsKuralHatasi` mesajı |
| Disk tutarsız | `.txn.bak`, yedekler/, `BASLATMA_HATASI.txt` |
| Kullanıcı dosyası bozuk | JSON syntax; CLI yeniden yaz |

---

## 3.60 Glossary (geniş)

| Terim | Anlam |
|-------|-------|
| Kart | Tek üretim iş kaydı (talep+stok) |
| Anahtar | `talep_no\|stok_no` kimlik |
| Dizgi | Üretim/dizgi aşaması |
| Workflow | Operatör durum/adet/zaman alanları |
| Snapshot | Values-only Excel kopyası |
| Source of truth | `kartlar.xlsx` |
| Anlık yedek | Mutasyon öncesi kopya |
| Audit log | `islem_logu.xlsx` |
| Process lock | `sunucu.lock` |
| Soft delete | `admin_gizli` |
| Partial production | Parçalı adet bitirme |

---

Bu ek bölümler; workflow sözleşmeleri, admin yan etkileri, lock/COM/CSRF invariants, E2E izler ve regression checklist ile handbook’u yeni developer onboarding’i için güçlendirir.
