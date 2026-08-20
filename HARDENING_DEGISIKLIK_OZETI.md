# PDGM Flask — Hardening Değişiklik Özeti

**Tarih:** 2026-08-20  
**Baseline:** `codebase.md` dump’ı (klasörde git reposu yok)  
**Kapsam:** Dump’a göre gerçekten uygulanmış değişiklikler; planlanıp yapılmayanlar ayrı bölümde.

---

## 1. Genel Özet

- **Değişen kaynak dosyalar:** 4 (`depo.py`, `app.py`, `excel_araclari.py`, `kullanici_yonet.py`)
- **Yeni dosyalar:** 5 (`.gitignore`, `.env.example`, `requirements.txt`, `run_pdgm.bat`, `yedek_disari_kopyala.bat`)
- **Hiç dokunulmayan:** tüm `templates/*`, `static/stil.css` (byte-exact dump ile aynı)
- **Ana gruplar:** persistence güvenliği · boot/recovery · process lock · COM/import · düşük riskli perf · audit DoS azaltımı · backup retention · kullanıcı CLI · deploy hygiene
- **En kritik reliability:** fsync, txn.bak koruma, stale lock reclaim, unknown DURUM→None, bozuk log quarantine / bozuk kart loud-fail
- **Security:** GET rapor/kayıt Excel mutasyonu kalktı; hatalı giriş Excel’e yazılmıyor; `.gitignore` / `.env.example`; `PDGM_BIND`
- **Persistence/recovery:** fsync + rollback’te backup saklama + startup recovery
- **Performance:** log slice-önce-deepcopy; atomik işlemde log deepcopy→`del`
- **UI:** hiçbir değişiklik yok (DOM patch / reload / CSS yapılmadı)

Yaklaşık satır (dump’a göre):

| Dosya | +/- |
| ----- | --- |
| `depo.py` | +233 / −29 |
| `app.py` | +116 / −26 |
| `excel_araclari.py` | +46 / −2 |
| `kullanici_yonet.py` | +72 / −0 |

---

## 2. Dosya Bazında Değişiklikler

### `depo.py`

#### Değişiklik 1 — fsync (`_diske_zorla` + `_temp_yaz`)

**Eski davranış:**  
`wb.save(temp)` → `wb.close()` → return (fsync yok)

**Yeni davranış:**  
close sonrası `_diske_zorla(temp)` (`O_RDONLY` + `os.fsync`)

**Neden değiştirildi:**  
Power-loss’ta OS page cache’te kalan yarım xlsx riski.

**Çözdüğü risk:**  
`os.replace` sonrası 0-byte / corrupt `kartlar.xlsx` penceresi.

**Davranış/regression riski:**  
Yazma birkaç ms yavaşlar; semantik aynı.

**Dokunulmayan kritik mantık:**  
`_coklu_yaz` sırası (temp → bak → replace → rollback).

---

#### Değişiklik 2 — txn.bak koruma (`_coklu_yaz`)

**Eski davranış:**  
`finally` her zaman `.txn.bak` siliyordu (rollback fail olsa bile).

**Yeni davranış:**  
`commit_basarili` True ise bak silinir; exception path’te bak kalır; başarılı restore edilen bak `None` yapılır.

**Neden değiştirildi:**  
Rollback fail + bak silme = kurtarma kopyası kaybı.

**Çözdüğü risk:**  
Transaction recovery kopyasının yanlışlıkla silinmesi.

**Davranış/regression riski:**  
Exception sonrası diskte `.txn.bak` kalabilir (bilinçli).

**Dokunulmayan kritik mantık:**  
replace sırası, ters rollback, temp cleanup.

---

#### Değişiklik 3 — `_durum_normalize`

**Eski davranış:**  
`esleme.get(sade, metin)` → unknown ham metin → `_kart_dogrula` → `kur()` patlar.

**Yeni davranış:**  
`esleme.get(sade)` → unknown `None` → Yönetim’de “DURUMU EKSİK”.

**Neden değiştirildi:**  
Elle yazılmış geçersiz DURUM hücresi sistemi boot edilemez yapıyordu.

**Çözdüğü risk:**  
Boot-brick.

**Davranış/regression riski:**  
Geçersiz DURUM artık boot’u öldürmez; operasyonda görünmez (zaten invalid’di).

**Dokunulmayan kritik mantık:**  
`excel_durum`, workflow kuralları, görünürlük ayrımı.

---

#### Değişiklik 4 — Stale lock reclaim

**Eski davranış:**  
Lock dosyası varsa her zaman refuse; manuel sil.

**Yeni davranış:**

1. PID ölü → reclaim  
2. PID canlı + image python değil → reclaim (PID reuse)  
3. PID canlı + python → refuse  
4. tasklist fail → fail-closed  

Lock içeriği: `PID|timestamp` (timestamp reclaim kriteri değil).

**Neden değiştirildi:**  
Crash / power-loss sonrası sunucunun açılamaması.

**Çözdüğü risk:**  
Operasyonel downtime (stale `sunucu.lock`).

**Davranış/regression riski:**  
Orta — yanlış reclaim teorik riski var; fail-closed tercih edildi. Windows’ta test şart.

**Dokunulmayan kritik mantık:**  
`O_EXCL` oluşturma, atexit release.

---

#### Değişiklik 5 — Kurtarılabilir açılış

**Eski davranış:**  
Bozuk `islem_logu` / `yuklemeler` → tüm boot fail.

**Yeni davranış:**  
`_bozuk_dosyayi_kenara_al` → `*.bozuk_TS` + console / `BASLATMA_HATASI.txt` append + boş liste.  
Kartlar hâlâ `VeriDogrulamaHatasi` fırlatır (sessiz boş kart yok).

**Neden değiştirildi:**  
Audit dosyası yüzünden UI’ya hiç girememe.

**Çözdüğü risk:**  
Yardımcı Excel bozulunca toplam kesinti.

**Davranış/regression riski:**  
Audit trail reset olabilir — yüksek görünürlüklü uyarı ile; sessiz değil.

**Dokunulmayan kritik mantık:**  
Kartlar source of truth; bozuk kartta hard-fail.

---

#### Değişiklik 6 — Yedek retention

**Eski davranış:**  
Retention yok.

**Yeni davranış:**  
`yedekleri_buda`: yalnız `^\d{8}_\d{6}_` klasörler (son 30); `^\d{8}_kartlar\.xlsx$` >90 gün; `anlik_yedek` sonunda çağrılır.

**Neden değiştirildi:**  
Sınırsız disk büyümesi.

**Çözdüğü risk:**  
Disk dolması / yedek listesi şişmesi.

**Davranış/regression riski:**  
Düşük — pattern dışı klasörler silinmez.

**Dokunulmayan kritik mantık:**  
Elle konmuş / pattern dışı klasörler.

---

#### Değişiklik 7 — Log perf / archive sırası

**Eski davranış:**

- `loglari_getir`: full deepcopy sonra slice  
- `_atomik_kart_islemi`: log deepcopy  
- `_log_arsivle`: önce truncate sonra arşiv yaz  

**Yeni davranış:**

- slice → deepcopy  
- atomik yolda `log_sayisi` + `del _loglar[n:]`  
- arşiv **önce yaz**, sonra truncate  
- `log_ekle` fail’de **hâlâ full deepcopy rollback**  

**Neden değiştirildi:**  
Kilit tutma süresi / gereksiz CPU; arşiv yazımı fail olursa RAM kaybı riski.

**Çözdüğü risk:**  
Yönetim sayfası ve buton basışlarında gereksiz deepcopy; archive-first veri kaybı önlemi.

**Davranış/regression riski:**  
Düşük (append-only call graph doğrulandı). MD’deki “arsivlendiyse rollback yapma” **reddedildi**.

**Dokunulmayan kritik mantık:**  
Kart deepcopy rollback.

---

### `app.py`

#### Değişiklik 1 — Recoverable startup wrapper

**Eski davranış:**  
Modül / `calistir` içinde çıplak `depo.kur()`.

**Yeni davranış:**  
`try/except VeriDogrulamaHatasi` → `_baslatma_hatasi_bildir` + `SystemExit(1)`.  
Çift `kur()` **silinmedi**.

**Neden / risk:**  
Traceback yerine Türkçe kurtarma + yedek listesi.

**Regression:**  
Düşük.

---

#### Değişiklik 2 — `_kullanicilari_al` JSON fallback

**Eski davranış:**  
Bozuk JSON → her request 500.

**Yeni davranış:**  
Decode/OSError’da son iyi cache varsa onu kullan + error log; cache yoksa raise.

**Neden / risk:**  
Typo ile tam kesinti.

**Regression:**  
Düşük (stale cache kısa süre kalabilir).

---

#### Değişiklik 3 — Hatalı giriş audit

**Eski davranış:**  
`depo.log_ekle("HATALI GİRİŞ")` → Excel full rewrite + global kilit.

**Yeni davranış:**  
Yalnız `app.logger.warning`.

**Neden / risk:**  
Unauth DoS / kilit blokajı.

**Davranış değişimi:**  
Audit semantiği değişti (Excel → uygulama logu) — bilinçli.

---

#### Değişiklik 4 — GET rapor / kayıt dosyası

**Eski davranış:**  
GET + `log_ekle` (state mutation).

**Yeni davranış:**  
`app.logger.info`; Excel audit yok.

**Neden / risk:**  
GET mutasyon / CSRF-DoS vektörü.

---

#### Değişiklik 5 — Upload retention + bind

**Yeni:**  
`_yuklenen_exceleri_buda` N=20; `PDGM_BIND` / `DINLENEN_ADRES` (default `0.0.0.0`).

---

### `excel_araclari.py`

#### Değişiklik 1 — Soft COM import + runtime gate

Mac’te app açılabilsin; Windows production’da COM aynı. Import yoksa `ExcelAktarimHatasi`.

#### Değişiklik 2 — COM cleanup

finally: ref=`None` + `gc.collect` + fail’de orphan temp sil.

#### Değişiklik 3 — Import serialization

`excelden_aktar` → `_import_kilidi` ile `_excelden_aktar`.

**Dokunulmayan:**  
AutomationSecurity / UpdateLinks / ReadOnly / EnableEvents vb.

---

### `kullanici_yonet.py`

**Yeni:**

- `parola` komutu  
- `rol` komutu  
- son-admin koruması (`pasif` / `rol` düşürme)  
- `yaz()` fsync  

**Dokunulmayan:**  
Web UI kullanıcı yönetimi yok (bilinçli).

---

### Yeni dosyalar

| Dosya | Amaç |
| ----- | ---- |
| `.gitignore` | `.env`, `data/`, venv |
| `.env.example` | Port/HTTPS/bind + parola notu |
| `requirements.txt` | Pin’li bağımlılıklar |
| `run_pdgm.bat` | Windows başlatma iskeleti |
| `yedek_disari_kopyala.bat` | Locale-safe tarih + robocopy + 60g retention |

---

## 3. Eski → Yeni Davranış

### Persistence

```text
ESKİ:
Kart işlemi
→ workbook temp dosyaya yazılır
→ os.replace

YENİ:
Kart işlemi
→ workbook temp dosyaya yazılır
→ fsync
→ os.replace

Fail rollback:
ESKİ → .txn.bak silinir
YENİ → .txn.bak korunur
```

### Startup

```text
ESKİ:
bozuk log/kart → traceback, açılmaz

YENİ:
bozuk log/yükleme → quarantine + uyarı + devam
bozuk kart → BASLATMA_HATASI.txt + SystemExit
```

### Process lock

```text
ESKİ:
lock varsa refuse (manuel sil)

YENİ:
ölü / non-python PID → reclaim
canlı python → refuse
```

### Excel import

```text
ESKİ:
COM refs tutulabilir; eşzamanlı COM mümkün; fail temp kalır

YENİ:
ref release + gc; import kilidi; fail temp silinir
```

### Log handling

```text
ESKİ:
deepcopy(tümü) → slice
atomik: log deepcopy rollback
arşiv: truncate → yaz

YENİ:
slice → deepcopy
atomik: del _loglar[n:]
arşiv: yaz → truncate; fail’de full RAM rollback
```

### Login / GET audit

```text
ESKİ:
Hatalı giriş / rapor / kayıt → Excel log_ekle

YENİ:
uygulama logu (app.logger)
```

### UI

```text
ESKİ = YENİ (değişmedi)
```

---

## 4. Kritik Değişikliklerin Teknik Açıklaması

### Persistence

- `_temp_yaz` + `_diske_zorla`: fsync eklendi.
- `_coklu_yaz`: sıra aynı; bak silme koşulu değişti (fail’de korunur).

### Status normalization

- `_durum_normalize`: unknown → `None`.
- Workflow business kuralları değişmedi.

### Process lock

- Stale reclaim: PID + python image.
- Timestamp diagnostik; reclaim kriteri değil.

### Startup

- `kur()`: log/yükleme quarantine; kart hard-fail.
- `app.py`: Türkçe recovery + `BASLATMA_HATASI.txt`.

### Backup

- Pattern-safe retention (30 anlık / 90 gün günlük).
- Offsite: `yedek_disari_kopyala.bat` (operasyonel).

### Excel COM

- Soft import, cleanup, orphan temp, `_import_kilidi`.
- COM security bayrakları aynı.

### Performance

- `loglari_getir` / `yuklemeleri_getir` slice-önce.
- `_atomik_kart_islemi` append-only `del`.
- Login/GET Excel yazımı kalktı.

### User management

- CLI `parola` / `rol` / son-admin / fsync.

### UI

- Uygulanmadı.

---

## 5. Özellikle Dokunulmayan Kritik Alanlar

| Alan | Durum |
| ---- | ----- |
| `_coklu_yaz` temp→backup→replace→rollback **sırası** | Değişmedi (yalnız bak silme koşulu) |
| `_kilit` / RLock kapsamı | Değişmedi |
| Kart deepcopy rollback | Değişmedi |
| `excel_import_uygula` plan→workflow sırası | Değişmedi |
| HAZIR/TESLİM adet conflict | Değişmedi |
| `anahtar = f"{talep_no}\|stok_no}"` | Değişmedi |
| `KART_ALANLARI` başlıkları | Değişmedi |
| Workflow geçiş kuralları | Değişmedi |
| Excel COM security ayarları | Değişmedi |
| CSRF / `@yetki` / `_oturum_kullanici_kontrol` | Değişmedi |
| Path traversal (`basename`+`commonpath`, whitelist) | Değişmedi |
| Templates / CSS / reload davranışı | Değişmedi |

---

## 6. Yapılmayan Öneriler

```text
Öneri: location.reload → DOM patch
Uygulandı mı: Hayır
Neden: Section move / sayaç / data-* stale riski (TOO RISKY)
Şimdilik bırakılmasının riski: UX yavaşlığı devam eder
Daha sonra yapılmalı mı: Evet, ayrı P2 sprint

Öneri: Log debounce / write_only workbook
Uygulandı mı: Hayır
Neden: Audit penceresi / görsel regresyon
Şimdilik risk: Ölçek büyürse yazma maliyeti
Daha sonra: İsteğe bağlı

Öneri: Blueprint / ozet.py split / dead code silme
Uygulandı mı: Hayır
Neden: Gereksiz refactor
Daha sonra: Hayır (acil değil)

Öneri: Formula injection genişletme / idle timeout / ProxyFix / login CSRF
Uygulandı mı: Hayır
Neden: Düşük öncelik / P2
Daha sonra: Ortama göre

Öneri: MD’deki “arsivlendiyse rollback yapma”
Uygulandı mı: Hayır (bilinçli reddedildi)
Neden: Disk/RAM ayrışması ve audit kaybı riski
Yerine: archive-first + full RAM rollback

Öneri: Retention’da tüm isdir silme
Uygulandı mı: Hayır
Neden: Yanlış klasör silme riski
Yerine: pattern-safe retention

Öneri: Salt-PID stale reclaim
Uygulandı mı: Hayır (geliştirildi)
Neden: PID reuse
Yerine: PID + python image kontrolü

Öneri: HTTPS/Caddy gerçek kurulumu + .env parola silme
Uygulandı mı: Kısmen (PDGM_BIND, .env.example, gitignore)
Neden: Operasyonel adım; canlı .env’e dokunulmadı
Daha sonra: Production öncesi zorunlu
```

---

## 7. Potansiyel Regression Noktaları

- Stale lock reclaim (Windows tasklist / PID reuse edge)
- Quarantine sonrası boş log ile devam (audit kaybı farkındalığı)
- Exception sonrası kalan `.txn.bak` dosyaları (beklenen)
- Hatalı giriş / rapor indirme artık Excel audit’te yok
- COM cleanup sonrası import (Windows + Excel)
- Retention’ın yanlışlıkla fazla silmediği (pattern kontrolü)

---

## 8. Test Edilmesi Gerekenler

```text
[ ] Sunucu normal açılıyor
[ ] İkinci instance aynı data/ ile reddediliyor
[ ] Force-kill sonrası (Windows) yeniden açılıyor / reclaim uyarısı
[ ] Bilinmeyen DURUM hücresi ile boot oluyor; Yönetim’de DURUMU EKSİK
[ ] Bozuk islem_logu → quarantine + uyarı + açılış
[ ] Bozuk kartlar.xlsx → SystemExit + BASLATMA_HATASI.txt
[ ] PLANA → DİZGİDE → kısmi bitir → tam bitir → HAZIR → TESLİM
[ ] Excel import workflow korunuyor
[ ] HAZIR/TESLİM adet conflict hâlâ reject
[ ] Kart işleminden sonra .yeni kalıntısı yok; başarıda .txn.bak yok
[ ] Yedek listele + geri yükle (koruma yedeği)
[ ] Hatalı giriş → uygulama.log; islem_logu mtime değişmiyor
[ ] Rapor/kayıt indirme çalışıyor; Excel’e audit satırı yazılmıyor
[ ] kullanici_yonet parola / rol / son-admin engeli
[ ] admin/operator/gozlemci giriş + CSRF mutasyon
[ ] Operator/monitor UI (reload hâlâ var; regressiyon yok)
[ ] (Windows) 3–5 Excel import; EXCEL.EXE kalıntı yok
```

---

## 9. Diff Özeti ve Önerilen Commit Grupları

**Not:** Klasörde git yok; sayılar `codebase.md` dump karşılaştırmasından.

| Dosya | Eklenen | Silinen | Ana değişiklik | Risk |
| ----- | ------: | ------: | -------------- | ---- |
| `depo.py` | ~233 | ~29 | fsync, bak, lock, normalize, startup, retention, perf | Orta (lock/bak) |
| `app.py` | ~116 | ~26 | startup handler, audit, upload, bind, JSON cache | Düşük–orta |
| `excel_araclari.py` | ~46 | ~2 | COM cleanup, import lock, soft import | Düşük–orta |
| `kullanici_yonet.py` | ~72 | ~0 | parola/rol/son-admin/fsync | Düşük |
| templates/CSS | 0 | 0 | — | — |
| `.gitignore` vb. | yeni | — | deploy hygiene | Yok |

Önerilen commit grupları (oluşturulmadı):

```text
fix(storage): fsync + keep txn backups on failed rollback
fix(startup): unknown DURUM=None; quarantine corrupt log/yukleme; loud card fail
fix(lock): reclaim verified stale sunucu.lock
fix(import): COM ref release, orphan temp, serialize imports
perf(storage): slice-before-deepcopy; append-only log rollback; archive-first
fix(audit): failed login + GET downloads to app log
feat(ops): backup retention; upload prune; user parola/rol
chore(deploy): gitignore, env.example, bat runners, PDGM_BIND
```

---

## Sonuç

> **Bu değişikliklerden sonra sistemin davranışı açısından en önemli fark:** Crash/power-loss ve bozuk yardımcı Excel dosyaları artık sistemi kalıcı olarak kapatmak yerine kontrollü recovery’ye yöneliyor; yazma yoluna fsync ve güvenli txn.bak koruması eklendi.

> **Mevcut business logic’in değişip değişmediği: HAYIR** — workflow, import overlay sırası, HAZIR/TESLİM conflict, kimlik şeması, yetki/CSRF ve UI akışı aynı; değişenler dayanıklılık, operasyonel recovery, audit taşıma (hatalı giriş/GET) ve düşük riskli performans.
