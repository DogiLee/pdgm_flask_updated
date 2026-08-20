# Repository Codebase Dump

Olusturulma tarihi: 2026-08-20 23:41:48

Repository: `pdgm_flask`

Toplam dosya: **18**


---


## Dosya Listesi

```text
PYTHON
├── app.py
├── depo.py
├── excel_araclari.py
├── kullanici_yonet.py

TEMPLATES
├── templates/base.html
├── templates/giris.html
├── templates/monitor.html
├── templates/operator.html
├── templates/ozet.html
├── templates/panel.html
├── templates/yetkisiz.html
├── templates/yonetim.html

STATIC
├── static/stil.css

SCRIPTS
├── run_pdgm.bat
├── yedek_disari_kopyala.bat

CONFIG / DIGER
├── .env.example
├── .gitignore
├── requirements.txt

```


---


# 1. PYTHON


## `app.py`


```python
"""PDGM · Baskı Dizgi Atölyesi İş Takip Sistemi.

Flask + Excel tabanlı intranet uygulaması.

Mimari sınırlar:
- kartlar.xlsx uygulamanın source of truth dosyasıdır.
- Tek Python process + çok thread kullanılır.
- Storage concurrency ve atomic write sorumluluğu depo.py'dedir.
- Workflow: PLANA ALINDI -> DİZGİDE -> HAZIR -> TESLİM EDİLDİ.
"""

from __future__ import annotations

import json
import logging
import logging.handlers
import os
import secrets
from collections import OrderedDict,deque
from datetime import date, datetime, timedelta
from functools import wraps
from urllib.parse import urlsplit
from uuid import uuid4
import threading
import time
from flask import (
    Flask,
    flash,
    jsonify,
    redirect,
    render_template,
    request,
    send_file,
    session,
    url_for,
)
from werkzeug.security import check_password_hash, generate_password_hash
from werkzeug.utils import secure_filename

import depo
import excel_araclari as ex
from dotenv import load_dotenv

KOK = os.path.dirname(os.path.abspath(__file__))
load_dotenv(os.path.join(KOK, ".env"))

VERI_KLASORU = os.path.join(KOK, "data")
YUKLEME_KLASORU = os.path.join(VERI_KLASORU, "yuklenen_exceller")
KULLANICI_DOSYASI = os.path.join(VERI_KLASORU, "kullanicilar.json")
LOG_DOSYASI = os.path.join(VERI_KLASORU, "uygulama.log")
SUNUCU_PORTU = int(os.environ.get("PDGM_PORT", "5001"))
DINLENEN_ADRES = os.environ.get("PDGM_BIND", "0.0.0.0")

app = Flask(__name__)
depo.process_kilidi_al()



# ---------------------------------------------------------------------------
# Config ve kullanıcı dosyası
# ---------------------------------------------------------------------------

def _anahtar() -> str:
    """Session secret'ını ilk çalıştırmada üretir ve diskte saklar."""
    yol = os.path.join(VERI_KLASORU, "gizli.key")
    os.makedirs(os.path.dirname(yol), exist_ok=True)

    if not os.path.exists(yol):
        with open(yol, "w", encoding="utf-8") as f:
            f.write(secrets.token_hex(32))
        try:
            os.chmod(yol, 0o600)
        except OSError:
            pass

    with open(yol, encoding="utf-8") as f:
        anahtar = f.read().strip()
    if len(anahtar) < 32:
        raise RuntimeError("data/gizli.key geçersiz veya çok kısa.")
    return anahtar


app.secret_key = _anahtar()
app.config.update(
    MAX_CONTENT_LENGTH=25 * 1024 * 1024,
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE="Lax",
    SESSION_COOKIE_SECURE=os.environ.get("PDGM_HTTPS", "0") == "1",
    PERMANENT_SESSION_LIFETIME=timedelta(hours=12),
)

@app.after_request
def _guvenlik_basliklari(response):
    response.headers.setdefault(
        "X-Content-Type-Options",
        "nosniff",
    )
    response.headers.setdefault(
        "X-Frame-Options",
        "DENY",
    )
    response.headers.setdefault(
        "Referrer-Policy",
        "same-origin",
    )
    response.headers.setdefault(
        "Permissions-Policy",
        "camera=(), microphone=(), geolocation=()",
    )

    if not request.path.startswith("/static/"):
        response.headers.setdefault(
            "Cache-Control",
            "no-store",
        )

    return response



def _gunluk_dosya_logu_kur():
    os.makedirs(VERI_KLASORU, exist_ok=True)
    if any(isinstance(handler, logging.FileHandler) for handler in app.logger.handlers):
        return

    handler = logging.handlers.RotatingFileHandler(
        LOG_DOSYASI,
        maxBytes=2_000_000,
        backupCount=5,
        encoding="utf-8",
    )
    handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s"))
    app.logger.setLevel(logging.INFO)
    app.logger.addHandler(handler)
    logging.getLogger("waitress").setLevel(logging.INFO)


def _kullanicilari_yukle() -> dict[str, dict]:
    """İlk çalışmada .env şifrelerinden hash üretir."""

    os.makedirs(VERI_KLASORU, exist_ok=True)

    if os.path.exists(KULLANICI_DOSYASI):
        with open(KULLANICI_DOSYASI, encoding="utf-8") as f:
            veri = json.load(f)

        if not isinstance(veri, dict) or not veri:
            raise RuntimeError("data/kullanicilar.json geçersiz.")

        return veri


    baslangic = {
        "admin": (
            "admin",
            "Sistem Yöneticisi",
            "PDGM_ADMIN_PASSWORD",
        ),
        "operator": (
            "operator",
            "Dizgi Operatörü",
            "PDGM_OPERATOR_PASSWORD",
        ),
        "gozlemci": (
            "gozlemci",
            "Gözlemci",
            "PDGM_VIEWER_PASSWORD",
        ),
    }


    sonuc = {}

    for kullanici, (rol, ad, env_adi) in baslangic.items():

        parola = os.getenv(env_adi)

        if not parola:
            raise RuntimeError(
                f"{env_adi} .env içinde bulunamadı."
            )

        sonuc[kullanici] = {
            "sifre_hash": generate_password_hash(parola),
            "rol": rol,
            "ad": ad,
            "aktif": True,
        }


    gecici = KULLANICI_DOSYASI + ".yeni"

    with open(gecici, "w", encoding="utf-8") as f:
        json.dump(
            sonuc,
            f,
            ensure_ascii=False,
            indent=2,
        )

    os.replace(
        gecici,
        KULLANICI_DOSYASI,
    )


    print("İlk kullanıcılar oluşturuldu.")
    print("Şifreler hash olarak kaydedildi.")

    return sonuc


_kullanici_mtime: float | None = None
_kullanici_onbellek: dict[str, dict] | None = None


def _kullanicilari_al() -> dict[str, dict]:
    """kullanicilar.json değişmişse restart gerektirmeden yeniden okur."""
    global _kullanici_mtime, _kullanici_onbellek

    try:
        mtime = os.path.getmtime(KULLANICI_DOSYASI)
    except OSError:
        if _kullanici_onbellek is None:
            _kullanici_onbellek = _kullanicilari_yukle()
        return _kullanici_onbellek

    if _kullanici_onbellek is None or _kullanici_mtime != mtime:
        try:
            with open(KULLANICI_DOSYASI, encoding="utf-8") as f:
                veri = json.load(f)
            if not isinstance(veri, dict) or not veri:
                raise RuntimeError("data/kullanicilar.json geçersiz.")
            _kullanici_onbellek = veri
            _kullanici_mtime = mtime
        except (json.JSONDecodeError, OSError, RuntimeError) as exc:
            if _kullanici_onbellek is not None:
                app.logger.error(
                    "kullanicilar.json okunamadı, son iyi önbellek kullanılıyor: %s",
                    exc,
                )
                return _kullanici_onbellek
            raise

    return _kullanici_onbellek


_kullanicilari_yukle()
_gunluk_dosya_logu_kur()


def _baslatma_hatasi_bildir(exc: Exception) -> None:
    """Açılış başarısızsa Python traceback yerine anlaşılır talimat üretir."""
    yedekler = []
    try:
        yedekler = depo.yedekleri_getir(5)
    except Exception:
        pass

    satirlar = [
        "=" * 70,
        "  PDGM İŞ TAKİP SİSTEMİ BAŞLATILAMADI",
        "=" * 70,
        "",
        f"  Hata: {exc}",
        "",
        "  Bu genellikle data/kartlar.xlsx dosyasının elle düzenlenmesi",
        "  sırasında oluşan bir veri hatasıdır.",
        "",
        "  YAPILACAKLAR:",
        "  1) data/kartlar.xlsx dosyasını Excel'de KAPATIN.",
        "  2) Yukarıdaki hata mesajında geçen kart ID / sütunu düzeltin,",
        "     VEYA aşağıdaki yedeklerden birini kartlar.xlsx üzerine kopyalayın.",
        "  3) Sunucuyu tekrar başlatın.",
        "",
    ]

    if yedekler:
        satirlar.append("  KULLANILABİLİR YEDEKLER (data/yedekler/ altında):")
        for y in yedekler:
            satirlar.append(f"    - {y['ad']}   ({y['zaman']}, {y['tip']})")
    else:
        satirlar.append("  UYARI: Kullanılabilir yedek bulunamadı.")

    satirlar += ["", "=" * 70, ""]
    metin = "\n".join(satirlar)

    print(metin)
    try:
        with open(
            os.path.join(VERI_KLASORU, "BASLATMA_HATASI.txt"),
            "w",
            encoding="utf-8",
        ) as f:
            f.write(metin)
    except OSError:
        pass


try:
    depo.kur()
except depo.VeriDogrulamaHatasi as _hata:
    app.logger.critical("Açılışta kart dosyası doğrulanamadı: %s", _hata)
    _baslatma_hatasi_bildir(_hata)
    raise SystemExit(1)


# ---------------------------------------------------------------------------
# Yetki, CSRF ve ortak template verisi
# ---------------------------------------------------------------------------

@app.before_request
def _oturum_kullanici_kontrol():
    kullanici = session.get("kullanici")
    if not kullanici:
        return None

    kayit = _kullanicilari_al().get(kullanici)
    if not kayit or not kayit.get("aktif", True):
        session.clear()
        if request.path.startswith("/api/"):
            return jsonify(hata="Oturum sonlandırıldı. Tekrar giriş yapın."), 401
        flash("Hesabınız pasif veya bulunamadı. Tekrar giriş yapın.", "hata")
        return redirect(url_for("giris"))

    session["rol"] = kayit.get("rol") or session.get("rol")
    session["ad"] = kayit.get("ad") or session.get("ad")
    return None


def yetki(*roller):
    def sarmalayici(fn):
        @wraps(fn)
        def ic(*args, **kwargs):
            if "kullanici" not in session:
                return redirect(url_for("giris", devam=request.path))
            if roller and session.get("rol") not in roller:
                return render_template("yetkisiz.html"), 403
            return fn(*args, **kwargs)

        return ic

    return sarmalayici


def _csrf_token_uret() -> str:
    token = session.get("csrf_token")
    if not token:
        token = secrets.token_urlsafe(32)
        session["csrf_token"] = token
    return token


def csrf_koru(fn):
    @wraps(fn)
    def ic(*args, **kwargs):
        beklenen = session.get("csrf_token")
        gelen = request.headers.get("X-CSRF-Token") or request.form.get("_csrf_token")

        if beklenen and gelen and secrets.compare_digest(beklenen, gelen):
            return fn(*args, **kwargs)

        if request.path.startswith("/api/"):
            return jsonify(hata="Geçersiz veya eksik CSRF token."), 403
        flash("Oturum doğrulaması başarısız. Sayfayı yenileyip tekrar deneyin.", "hata")
        return redirect(request.referrer or url_for("ana"))

    return ic


def _guvenli_devam_hedefi(hedef: str | None) -> bool:
    if not hedef or "\\" in hedef:
        return False

    parca = urlsplit(hedef)

    return (
        not parca.scheme
        and not parca.netloc
        and hedef.startswith("/")
        and not hedef.startswith("//")
    )


@app.context_processor
def genel_degiskenler():
    return {
        "oturum_ad": session.get("ad"),
        "oturum_rol": session.get("rol"),
        "oturum_kullanici": session.get("kullanici"),
        "bugun": date.today().strftime("%d.%m.%Y"),
        "csrf_token": _csrf_token_uret() if "kullanici" in session else "",
    }


@app.template_filter("gun")
def gun_filtresi(deger):
    if not deger:
        return "—"
    parcalar = str(deger)[:10].split("-")
    return f"{parcalar[2]}.{parcalar[1]}.{parcalar[0]}" if len(parcalar) == 3 else str(deger)


@app.errorhandler(PermissionError)
def _excel_kilitli(_hata):
    mesaj = (
        "Bir kayıt Excel dosyası şu anda Excel'de açık olabilir. "
        "Kartlar, işlem logu ve yükleme geçmişi dosyalarını kapatıp "
        "tekrar deneyin."
    )

    if request.path.startswith("/api/"):
        return jsonify(hata=mesaj), 423

    flash(mesaj, "hata")
    return redirect(url_for("ana"))


# ---------------------------------------------------------------------------
# Giriş / çıkış
# ---------------------------------------------------------------------------
_GIRIS_LIMIT = 8
_GIRIS_PENCERE_SN = 5 * 60

_giris_kilit = threading.Lock()
_giris_basarisiz: dict[str, deque[float]] = {}


def _giris_ip() -> str:
    return request.remote_addr or "bilinmeyen"


def _giris_engelli_mi() -> bool:
    ip = _giris_ip()
    an = time.monotonic()

    with _giris_kilit:
        denemeler = _giris_basarisiz.get(ip)

        if not denemeler:
            return False

        while (
            denemeler
            and an - denemeler[0] >= _GIRIS_PENCERE_SN
        ):
            denemeler.popleft()

        if not denemeler:
            _giris_basarisiz.pop(ip, None)
            return False

        return len(denemeler) >= _GIRIS_LIMIT


def _giris_basarisiz_kaydet():
    ip = _giris_ip()
    an = time.monotonic()

    with _giris_kilit:
        denemeler = _giris_basarisiz.setdefault(
            ip,
            deque(),
        )

        while (
            denemeler
            and an - denemeler[0] >= _GIRIS_PENCERE_SN
        ):
            denemeler.popleft()

        denemeler.append(an)


def _giris_limit_temizle():
    with _giris_kilit:
        _giris_basarisiz.pop(
            _giris_ip(),
            None,
        )

@app.route("/giris", methods=["GET", "POST"])
def giris():
    if request.method == "GET":
        return render_template("giris.html")

    if _giris_engelli_mi():
        return render_template(
            "giris.html",
            hata=(
                "Çok fazla başarısız giriş denemesi yapıldı. "
                "Birkaç dakika sonra tekrar deneyin."
            ),
        ), 429

    kullanici = (
        request.form.get("kullanici") or ""
    ).strip()
    sifre = request.form.get("sifre") or ""

    # Aşırı büyük input'u password hash fonksiyonuna ve Excel loguna sokma.
    if len(kullanici) > 128 or len(sifre) > 512:
        _giris_basarisiz_kaydet()
        return render_template(
            "giris.html",
            hata="Kullanıcı adı veya şifre hatalı.",
        ), 401

    kayit = _kullanicilari_al().get(kullanici)

    if (
        kayit
        and kayit.get("aktif", True)
        and check_password_hash(
            kayit.get("sifre_hash", ""),
            sifre,
        )
    ):
        _giris_limit_temizle()

        session.clear()
        session.permanent = True
        session.update(
            kullanici=kullanici,
            rol=kayit["rol"],
            ad=kayit["ad"],
        )
        _csrf_token_uret()

        try:
            depo.log_ekle(
                kullanici,
                kayit["rol"],
                "GİRİŞ YAPILDI",
            )
        except Exception:
            # Audit Excel geçici olarak yazılamasa bile başarılı auth
            # bozulmasın. Fallback olarak uygulama loguna yaz.
            app.logger.exception(
                "Giriş audit kaydı Excel'e yazılamadı: %s",
                kullanici,
            )

        hedef = request.args.get("devam")
        return redirect(
            hedef
            if _guvenli_devam_hedefi(hedef)
            else url_for("ana")
        )

    _giris_basarisiz_kaydet()

    app.logger.warning(
        "HATALI GİRİŞ · kullanici=%s · ip=%s",
        kullanici or "-",
        _giris_ip(),
    )

    return render_template(
        "giris.html",
        hata="Kullanıcı adı veya şifre hatalı.",
    ), 401


@app.route("/cikis", methods=["POST"])
@yetki("admin", "operator", "gozlemci")
@csrf_koru
def cikis():
    kullanici = session.get("kullanici", "-")
    rol = session.get("rol", "")

    try:
        depo.log_ekle(
            kullanici,
            rol,
            "ÇIKIŞ YAPILDI",
        )
    except Exception:
        app.logger.exception(
            "Çıkış audit kaydı Excel'e yazılamadı: %s",
            kullanici,
        )
    finally:
        session.clear()

    return redirect(url_for("giris"))


@app.route("/")
def ana():
    hedefler = {
        "admin": "yonetim",
        "operator": "operator",
        "gozlemci": "panel",
    }
    endpoint = hedefler.get(session.get("rol"))
    return redirect(url_for(endpoint or "giris"))


# ---------------------------------------------------------------------------
# Ekran verileri
# ---------------------------------------------------------------------------

def _pano_verisi():
    kartlar = depo.kartlari_getir()
    sayac = {
        "plana_alindi": sum(k["durum"] == depo.PLANA_ALINDI for k in kartlar),
        "dizgide": sum(k["durum"] == depo.DIZGIDE for k in kartlar),
        "hazir": sum(k["durum"] == depo.HAZIR for k in kartlar),
        "teslim": sum(k["durum"] == depo.TESLIM_EDILDI for k in kartlar),
        "gecikme": sum(
            k["renk"] == "kotu" and k["durum"] != depo.TESLIM_EDILDI
            for k in kartlar
        ),
    }
    return {
        "kartlar": kartlar,
        "sayac": sayac,
        "guncelleme": datetime.now().strftime("%H:%M:%S"),
    }


@app.route("/panel")
@yetki("admin", "operator", "gozlemci")
def panel():
    veri = _pano_verisi()
    kartlar = veri["kartlar"]

    teslim_edilen = sorted(
        [k for k in kartlar if k["durum"] == depo.TESLIM_EDILDI],
        key=lambda k: k.get("teslim_zamani") or k.get("gerceklesen_teslim") or "",
        reverse=True,
    )[:6]

    return render_template(
        "panel.html",
        sayac=veri["sayac"],
        guncelleme=veri["guncelleme"],
        dizgide=[k for k in kartlar if k["durum"] == depo.DIZGIDE],
        hazir=[k for k in kartlar if k["durum"] == depo.HAZIR],
        plana_alindi=[k for k in kartlar if k["durum"] == depo.PLANA_ALINDI][:12],
        teslim_edilen=teslim_edilen,
    )


@app.route("/monitor")
@yetki("admin", "operator", "gozlemci")
def monitor():
    veri = _pano_verisi()
    kartlar = veri["kartlar"]

    return render_template(
        "monitor.html",
        guncelleme=veri["guncelleme"],
        dizgide=[k for k in kartlar if k["durum"] == depo.DIZGIDE],
        plana_alindi=[k for k in kartlar if k["durum"] == depo.PLANA_ALINDI],
    )


@app.route("/operator")
@yetki("admin", "operator")
def operator():
    veri = _pano_verisi()
    return render_template("operator.html", kartlar=veri["kartlar"], sayac=veri["sayac"])


@app.route("/yonetim")
@yetki("admin")
def yonetim():
    kartlar = depo.kartlari_yonetim_getir()
    kaynakta_olmayan = [
        k
        for k in kartlar
        if k.get("kaynakta_yok") and k.get("durum") != depo.TESLIM_EDILDI
    ]

    return render_template(
        "yonetim.html",
        kartlar=kartlar,
        kaynakta_olmayan=kaynakta_olmayan,
        durumu_eksik=depo.durumu_eksik_kartlari_getir(),
        gizlenen_kartlar=depo.gizlenen_kartlari_getir(),
        yedekler=depo.yedekleri_getir(12),
        yuklemeler=depo.yuklemeleri_getir(8),
        loglar=depo.loglari_getir(25),
    )


@app.route("/ozet")
@yetki("admin", "operator", "gozlemci")
def ozet():
    return render_template("ozet.html", **ozet_hesapla())


@app.route("/api/veriler")
@yetki("admin", "operator", "gozlemci")
def api_veriler():
    return jsonify(_pano_verisi())


# ---------------------------------------------------------------------------
# Operatör API'leri
# ---------------------------------------------------------------------------

def _api_kart_islemi(fn):
    try:
        return fn()
    except depo.KartBulunamadi as hata:
        return jsonify(hata=str(hata)), 404
    except depo.IsKuralHatasi as hata:
        return jsonify(hata=str(hata)), 409
    except (depo.VeriDogrulamaHatasi, TypeError, ValueError) as hata:
        return jsonify(hata=str(hata)), 400


@app.route("/api/basla", methods=["POST"])
@yetki("admin", "operator")
@csrf_koru
def api_basla():
    veri = request.get_json(silent=True) or {}

    def islem():
        kart = depo.kart_baslat(
            kart_id=veri.get("kart_id"),
            adet=veri.get("adet"),
            kullanici=session["kullanici"],
            rol=session["rol"],
            aciklama=(veri.get("not") or "").strip(),
        )
        return jsonify(tamam=True, mesaj="Kart DİZGİDE durumuna alındı.", kart=kart)

    return _api_kart_islemi(islem)


@app.route("/api/bitir", methods=["POST"])
@yetki("admin", "operator")
@csrf_koru
def api_bitir():
    veri = request.get_json(silent=True) or {}

    def islem():
        kart, uretim_bitti, mesaj = depo.kart_bitir(
            kart_id=veri.get("kart_id"),
            adet=veri.get("adet"),
            kullanici=session["kullanici"],
            rol=session["rol"],
            aciklama=(veri.get("not") or "").strip(),
        )
        return jsonify(
            tamam=True,
            uretim_bitti=uretim_bitti,
            mesaj=mesaj,
            kart=kart,
        )

    return _api_kart_islemi(islem)


@app.route("/api/hazirla", methods=["POST"])
@yetki("admin", "operator")
@csrf_koru
def api_hazirla():
    veri = request.get_json(silent=True) or {}

    def islem():
        kart = depo.kart_hazirla(
            kart_id=veri.get("kart_id"),
            kullanici=session["kullanici"],
            rol=session["rol"],
            aciklama=(veri.get("not") or "").strip(),
        )
        return jsonify(tamam=True, mesaj="Kart HAZIR durumuna alındı.", kart=kart)

    return _api_kart_islemi(islem)


@app.route("/api/teslim-et", methods=["POST"])
@yetki("admin", "operator")
@csrf_koru
def api_teslim_et():
    veri = request.get_json(silent=True) or {}

    def islem():
        kart = depo.kart_teslim_et(
            kart_id=veri.get("kart_id"),
            kullanici=session["kullanici"],
            rol=session["rol"],
            aciklama=(veri.get("not") or "").strip(),
        )
        return jsonify(tamam=True, mesaj="Kart TESLİM EDİLDİ olarak kaydedildi.", kart=kart)

    return _api_kart_islemi(islem)


@app.route("/api/not", methods=["POST"])
@yetki("admin", "operator")
@csrf_koru
def api_not():
    veri = request.get_json(silent=True) or {}

    def islem():
        kart = depo.kart_not_guncelle(
            kart_id=veri.get("kart_id"),
            aciklama=(veri.get("not") or "").strip(),
            kullanici=session["kullanici"],
            rol=session["rol"],
        )
        return jsonify(tamam=True, kart=kart)

    return _api_kart_islemi(islem)


# ---------------------------------------------------------------------------
# Admin API'leri
# ---------------------------------------------------------------------------

@app.route("/api/admin/kart-ekle", methods=["POST"])
@yetki("admin")
@csrf_koru
def api_kart_ekle():
    veri = request.get_json(silent=True) or {}

    def islem():
        kart = depo.admin_kart_ekle(
            sira=veri.get("sira"),
            talep_no=veri.get("talep_no"),
            talep_sahibi=veri.get("talep_sahibi"),
            stok_no=veri.get("stok_no"),
            toplam_adet=veri.get("toplam_adet"),
            plan_hafta=veri.get("plan_hafta"),
            plan_baslama=veri.get("plan_baslama"),
            plan_teslim=veri.get("plan_teslim"),
            gerceklesen_teslim=veri.get("gerceklesen_teslim"),
            pcb=veri.get("pcb"),
            aciklama=veri.get("not"),
            kullanici=session["kullanici"],
        )
        return jsonify(tamam=True, kart=kart), 201

    return _api_kart_islemi(islem)


@app.route("/api/admin/duzenle", methods=["POST"])
@yetki("admin")
@csrf_koru
def api_duzenle():
    veri = request.get_json(silent=True) or {}

    def islem():
        kart = depo.admin_kart_duzenle(
            kart_id=veri.get("kart_id"),
            durum=veri.get("durum"),
            tamamlanan_adet=veri.get("tamamlanan_adet"),
            toplam_adet=veri.get("toplam_adet"),
            plan_hafta=veri.get("plan_hafta"),
            plan_baslama=veri.get("plan_baslama"),
            plan_teslim=veri.get("plan_teslim"),
            gerceklesen_teslim=veri.get("gerceklesen_teslim"),
            aciklama=veri.get("not"),
            kullanici=session["kullanici"],
        )
        return jsonify(tamam=True, kart=kart)

    return _api_kart_islemi(islem)


@app.route("/api/admin/kart-sil", methods=["POST"])
@yetki("admin")
@csrf_koru
def api_kart_sil():
    veri = request.get_json(silent=True) or {}

    def islem():
        depo.admin_kart_gizle(veri.get("kart_id"), session["kullanici"])
        return jsonify(tamam=True)

    return _api_kart_islemi(islem)


@app.route("/yonetim/kart-geri-getir", methods=["POST"])
@yetki("admin")
@csrf_koru
def kart_geri_getir():
    try:
        kart = depo.admin_kart_geri_getir(
            kart_id=request.form.get("kart_id"),
            kullanici=session["kullanici"],
        )
    except (depo.KartBulunamadi, depo.IsKuralHatasi) as hata:
        flash(str(hata), "hata")
    else:
        flash(
            f"Kart geri getirildi · {kart.get('talep_no') or '—'} · {kart.get('stok_no') or '—'}",
            "basari",
        )
    return redirect(url_for("yonetim"))


@app.route("/yonetim/yedek-geri-yukle", methods=["POST"])
@yetki("admin")
@csrf_koru
def yedek_geri_yukle():
    yedek_adi = (request.form.get("yedek") or "").strip()

    try:
        sonuc = depo.yedekten_geri_yukle(yedek_adi, session["kullanici"])
    except (depo.IsKuralHatasi, depo.VeriDogrulamaHatasi) as hata:
        flash(f"Yedek geri yüklenemedi: {hata}", "hata")
    except OSError as hata:
        app.logger.exception("Yedek geri yükleme dosya hatası")
        flash(f"Yedek geri yüklenirken dosya hatası oluştu: {hata}", "hata")
    else:
        flash(
            f"Yedek geri yüklendi · {sonuc['kart']} kart. "
            f"Önceki durum '{os.path.basename(sonuc['koruma_yedegi'])}' içinde korundu.",
            "basari",
        )
        app.logger.warning("Admin yedek geri yükledi: %s", yedek_adi)

    return redirect(url_for("yonetim"))


# ---------------------------------------------------------------------------
# Excel yükleme / kayıt dosyaları / rapor
# ---------------------------------------------------------------------------

YUKLEME_SAKLA = 20


def _yuklenen_exceleri_buda():
    """En yeni N yüklenmiş Excel'i tutar, gerisini siler."""
    try:
        dosyalar = []
        for ad in os.listdir(YUKLEME_KLASORU):
            yol = os.path.join(YUKLEME_KLASORU, ad)
            if os.path.isfile(yol):
                dosyalar.append((os.path.getmtime(yol), yol))
        dosyalar.sort(reverse=True)
        for _, yol in dosyalar[YUKLEME_SAKLA:]:
            try:
                os.remove(yol)
            except OSError:
                pass
    except OSError:
        pass


@app.route("/yonetim/yukle", methods=["POST"])
@yetki("admin")
@csrf_koru
def yukle():
    dosya = request.files.get("dosya")
    if not dosya or not dosya.filename:
        flash("Dosya seçilmedi.", "hata")
        return redirect(url_for("yonetim"))

    guvenli_ad = secure_filename(dosya.filename)
    if not guvenli_ad.lower().endswith((".xlsx", ".xlsm")):
        flash("Sadece .xlsx veya .xlsm dosyası yükleyin.", "hata")
        return redirect(url_for("yonetim"))

    os.makedirs(YUKLEME_KLASORU, exist_ok=True)
    kayit_adi = f"{datetime.now():%Y%m%d_%H%M%S}_{uuid4().hex[:8]}_{guvenli_ad}"
    yol = os.path.join(YUKLEME_KLASORU, kayit_adi)
    dosya.save(yol)

    try:
        sonuc = ex.excelden_aktar(yol, session["kullanici"])
        mesaj = (
            f"Excel aktarıldı · {sonuc['satir']} satır · {sonuc['yeni']} yeni · "
            f"{sonuc['guncellenen']} güncellendi · {sonuc['workflow_korundu']} workflow korundu."
        )
        if sonuc.get("pasife_alinan"):
            mesaj += f" · {sonuc['pasife_alinan']} kart kaynak Excel'de artık yok."
        if sonuc.get("uyari"):
            mesaj += (
                f" · {sonuc['uyari']} satırda DURUM boş/geçersiz veya başka bir küçük veri uyarısı var; "
                "bu kartlar Yönetim ekranında kontrol edilebilir."
            )
        if sonuc.get("yedek"):
            mesaj += f" · Yedek: {os.path.basename(sonuc['yedek'])}"

        flash(mesaj, "basari")
        app.logger.info("Excel import OK: %s", mesaj)
        _yuklenen_exceleri_buda()
    except (ex.ExcelAktarimHatasi, depo.IsKuralHatasi, depo.VeriDogrulamaHatasi) as hata:
        flash(f"Excel içeriği kabul edilmedi: {hata}", "hata")
        app.logger.warning("Excel import reddedildi: %s", hata)
    except Exception:  # noqa: BLE001
        app.logger.exception("Excel import sırasında beklenmeyen hata")
        flash("Excel okunurken beklenmeyen bir hata oluştu. Uygulama logunu kontrol edin.", "hata")

    return redirect(url_for("yonetim"))


@app.route("/yonetim/yeniden-oku", methods=["POST"])
@yetki("admin")
@csrf_koru
def yeniden_oku():
    try:
        adet = depo.kartlari_diskten_yeniden_yukle()
    except depo.VeriDogrulamaHatasi as hata:
        flash(f"kartlar.xlsx yeniden okunamadı: {hata}", "hata")
        return redirect(url_for("yonetim"))

    depo.log_ekle(
        session["kullanici"],
        "admin",
        "KART DOSYASI YENİDEN OKUNDU",
        detay=f"{adet} kart",
    )
    flash(f"kartlar.xlsx doğrulandı ve yeniden okundu · {adet} kart.", "basari")
    return redirect(url_for("yonetim"))


@app.route("/yonetim/kayit-dosyasi/<hangi>")
@yetki("admin")
def kayit_dosyasi(hangi):
    dosyalar = {
        "kartlar": depo.KARTLAR_DOSYA,
        "log": depo.LOG_DOSYA,
        "yuklemeler": depo.YUKLEME_DOSYA,
    }
    yol = dosyalar.get(hangi)
    if not yol or not os.path.exists(yol):
        flash("Dosya henüz oluşmamış.", "hata")
        return redirect(url_for("yonetim"))

    app.logger.info(
        "KAYIT DOSYASI İNDİRİLDİ · kullanici=%s · hangi=%s",
        session.get("kullanici"),
        hangi,
    )
    return send_file(yol, as_attachment=True, download_name=os.path.basename(yol))


@app.route("/yonetim/rapor")
@yetki("admin")
def rapor_indir():
    ozet = ozet_hesapla()
    ozet_satirlari = [
        ["Rapor tarihi", datetime.now().strftime("%d.%m.%Y %H:%M")],
        ["Toplam kart", ozet["genel"]["toplam"]],
        ["Plana alındı", ozet["genel"]["plana_alindi"]],
        ["Dizgide", ozet["genel"]["dizgide"]],
        ["Hazır", ozet["genel"]["hazir"]],
        ["Teslim edildi", ozet["genel"]["teslim"]],
        ["Süresi aşan açık kart", ozet["genel"]["gecikme"]],
        ["Bu hafta teslim edilen", ozet["donemler"]["Bu hafta"]["kart"]],
        ["Bu ay teslim edilen", ozet["donemler"]["Bu ay"]["kart"]],
        ["Bu yıl teslim edilen", ozet["donemler"]["Bu yıl"]["kart"]],
        ["Bu ay zamanında teslim (%)", ozet["donemler"]["Bu ay"]["zamaninda_yuzde"]],
        ["Bu ay ortalama teslim sapması (gün)", ozet["donemler"]["Bu ay"]["ort_sapma"]],
    ]

    wb = ex.calisma_kitabi_uret(
        depo.kartlari_getir(sadece_gorunen=False),
        depo.loglari_getir(),
        ozet_satirlari,
    )
    app.logger.info(
        "RAPOR İNDİRİLDİ · kullanici=%s",
        session.get("kullanici"),
    )
    return send_file(
        ex.kitap_baytlari(wb),
        as_attachment=True,
        download_name=ex.dosya_adi("PDGM_Rapor"),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# ---------------------------------------------------------------------------
# Özet hesapları
# ---------------------------------------------------------------------------

def _teslim_tarihi(kart) -> str | None:
    deger = kart.get("gerceklesen_teslim") or kart.get("teslim_zamani")
    return str(deger)[:10] if deger else None


def _donem_ozeti(kartlar, baslangic, bitis):
    alt = baslangic.strftime("%Y-%m-%d")
    ust = bitis.strftime("%Y-%m-%d")
    secilen = []

    for kart in kartlar:
        if kart.get("durum") != depo.TESLIM_EDILDI:
            continue
        teslim = _teslim_tarihi(kart)
        if teslim and alt <= teslim <= ust:
            secilen.append(kart)

    sapmalar = [kart["sapma"] for kart in secilen if kart["sapma"] is not None]
    zamaninda = sum(sapma <= 0 for sapma in sapmalar)
    return {
        "kart": len(secilen),
        "adet": sum(kart["tamamlanan_adet"] for kart in secilen),
        "zamaninda": zamaninda,
        "gecikmeli": len(sapmalar) - zamaninda,
        "zamaninda_yuzde": round(zamaninda / len(sapmalar) * 100) if sapmalar else 0,
        "ort_sapma": round(sum(sapmalar) / len(sapmalar), 1) if sapmalar else 0,
    }


def ozet_hesapla():
    kartlar = [
        kart
        for kart in depo.kartlari_getir(sadece_gorunen=False)
        if not kart.get("admin_gizli") and kart.get("aktif", 1) == 1
    ]
    bugun_tarih = date.today()

    donemler = OrderedDict()
    donemler["Bu hafta"] = _donem_ozeti(
        kartlar,
        bugun_tarih - timedelta(days=bugun_tarih.weekday()),
        bugun_tarih,
    )
    donemler["Bu ay"] = _donem_ozeti(kartlar, bugun_tarih.replace(day=1), bugun_tarih)
    donemler["Bu yıl"] = _donem_ozeti(
        kartlar,
        bugun_tarih.replace(month=1, day=1),
        bugun_tarih,
    )

    haftalar = []
    for geri in range(7, -1, -1):
        bas = bugun_tarih - timedelta(days=bugun_tarih.weekday() + geri * 7)
        son = min(bas + timedelta(days=6), bugun_tarih)
        alt, ust = bas.strftime("%Y-%m-%d"), son.strftime("%Y-%m-%d")

        planlanan = sum(
            1
            for kart in kartlar
            if kart.get("plan_teslim") and alt <= kart["plan_teslim"] <= ust
        )
        teslim = sum(
            1
            for kart in kartlar
            if kart.get("durum") == depo.TESLIM_EDILDI
            and (teslim_tarihi := _teslim_tarihi(kart))
            and alt <= teslim_tarihi <= ust
        )
        haftalar.append(
            {
                "etiket": bas.strftime("%d.%m"),
                "hafta_no": bas.isocalendar()[1],
                "planlanan": planlanan,
                "teslim": teslim,
                "sapma": teslim - planlanan,
            }
        )

    en_yuksek = max(
        [hafta["planlanan"] for hafta in haftalar]
        + [hafta["teslim"] for hafta in haftalar]
        + [1]
    )

    genel = {
        "toplam": len(kartlar),
        "plana_alindi": sum(k["durum"] == depo.PLANA_ALINDI for k in kartlar),
        "dizgide": sum(k["durum"] == depo.DIZGIDE for k in kartlar),
        "hazir": sum(k["durum"] == depo.HAZIR for k in kartlar),
        "teslim": sum(k["durum"] == depo.TESLIM_EDILDI for k in kartlar),
        "durumu_eksik": sum(not k.get("durum") for k in kartlar),
        "gecikme": sum(
            k.get("gorunur")
            and k.get("renk") == "kotu"
            and k.get("durum") != depo.TESLIM_EDILDI
            for k in kartlar
        ),
    }
    geciken = [
        k
        for k in kartlar
        if k.get("gorunur")
        and k.get("renk") == "kotu"
        and k.get("durum") != depo.TESLIM_EDILDI
    ]

    return {
        "genel": genel,
        "donemler": donemler,
        "haftalar": haftalar,
        "en_yuksek": en_yuksek,
        "geciken_kartlar": geciken,
    }


# ---------------------------------------------------------------------------
# Çalıştırma
# ---------------------------------------------------------------------------

def calistir():
    depo.process_kilidi_al()
    try:
        depo.kur()
    except depo.VeriDogrulamaHatasi as hata:
        app.logger.critical("Açılışta kart dosyası doğrulanamadı: %s", hata)
        _baslatma_hatasi_bildir(hata)
        raise SystemExit(1)
    _gunluk_dosya_logu_kur()

    print("\n  PDGM İş Takip Sistemi çalışıyor")
    print(f"  Dinlenen adres : {DINLENEN_ADRES}:{SUNUCU_PORTU}")
    print(f"  Bu bilgisayarda : http://127.0.0.1:{SUNUCU_PORTU}")
    print(f"  Ağdaki diğer PC : http://<sunucunun-ip-adresi>:{SUNUCU_PORTU}")
    print(f"  Kayıtlar        : {VERI_KLASORU}")
    print("  Sunucu modeli   : tek process + çok thread")
    print(f"  Uygulama logu   : {LOG_DOSYASI}")
    print("  Durdurmak için  : Ctrl + C\n")
    app.logger.info("Sunucu başladı port=%s", SUNUCU_PORTU)

    try:
        from waitress import serve

        serve(app, host=DINLENEN_ADRES, port=SUNUCU_PORTU, threads=8)
    except ImportError:
        app.run(host=DINLENEN_ADRES, port=SUNUCU_PORTU, debug=False, threaded=True)


if __name__ == "__main__":
    calistir()
```


## `depo.py`


```python
"""PDGM İş Takip Sistemi için Excel tabanlı veri deposu.

Tasarım hedefi:
- kartlar.xlsx uygulamanın source of truth dosyasıdır.
- Tek Python process + çok thread modeli kullanılır.
- Read/modify/write işlemleri RLock ile korunur.
- Yazmalar temp dosya + os.replace ile yapılır.
- Kritik işlemler öncesi yedek alınır.
- Workflow yalnız dört gerçek durumdan oluşur:
    PLANA ALINDI -> DİZGİDE -> HAZIR -> TESLİM EDİLDİ
- Kaynak Excel'de DURUM boşsa kart saklanır fakat operasyon ekranlarında gösterilmez.

Not: Excel transactional database değildir. Aynı data klasörünü birden fazla Python
process'i paylaşmamalıdır. Bu sınır process lock ile açıkça korunur.
"""

from __future__ import annotations

import atexit
import copy
import os
import re
import shutil
import subprocess
import threading
from datetime import date, datetime
from uuid import uuid4

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.utils.datetime import from_excel


# ---------------------------------------------------------------------------
# Dosyalar ve workflow sabitleri
# ---------------------------------------------------------------------------

KOK = os.path.dirname(os.path.abspath(__file__))
VERI_KLASORU = os.path.join(KOK, "data")
KARTLAR_DOSYA = os.path.join(VERI_KLASORU, "kartlar.xlsx")
LOG_DOSYA = os.path.join(VERI_KLASORU, "islem_logu.xlsx")
YUKLEME_DOSYA = os.path.join(VERI_KLASORU, "yuklemeler.xlsx")
YEDEK_KLASORU = os.path.join(VERI_KLASORU, "yedekler")
PROCESS_KILIT = os.path.join(VERI_KLASORU, "sunucu.lock")

PLANA_ALINDI = "PLANA ALINDI"
DIZGIDE = "DİZGİDE"
HAZIR = "HAZIR"
TESLIM_EDILDI = "TESLİM EDİLDİ"

GECERLI_DURUMLAR = {PLANA_ALINDI, DIZGIDE, HAZIR, TESLIM_EDILDI}
AKTIF_DURUMLAR = {PLANA_ALINDI, DIZGIDE, HAZIR}

# Template/rapor kodunun daha okunabilir olması için durum sırası tek yerde tutulur.
SIRALAMA = {
    DIZGIDE: 0,
    HAZIR: 1,
    PLANA_ALINDI: 2,
    TESLIM_EDILDI: 3,
    None: 4,
}


# ---------------------------------------------------------------------------
# Exceptions
# ---------------------------------------------------------------------------

class DepoHatasi(Exception):
    """Depo katmanının temel exception sınıfı."""


class KartBulunamadi(DepoHatasi):
    pass


class IsKuralHatasi(DepoHatasi):
    pass


class VeriDogrulamaHatasi(DepoHatasi):
    pass


# ---------------------------------------------------------------------------
# Excel şemaları
# ---------------------------------------------------------------------------

KART_ALANLARI = [
    ("ID", "id"),
    ("Sıra", "sira"),
    ("Talep NO", "talep_no"),
    ("Kart Stok No", "stok_no"),
    ("Talep Sahibi", "talep_sahibi"),
    ("Toplam Adet", "toplam_adet"),
    ("Adet Metni", "adet_metin"),
    ("Plan Haftası", "plan_hafta"),
    ("Plan Başlangıç", "plan_baslama"),
    ("Plan Teslim", "plan_teslim"),
    ("Gerçekleşen Teslim", "gerceklesen_teslim"),
    ("Excel Durumu", "excel_durum"),
    ("PCB", "pcb"),
    ("Durum", "durum"),
    ("Başlangıç Adedi", "baslangic_adet"),
    ("Tamamlanan Adet", "tamamlanan_adet"),
    ("Başlama Zamanı", "baslama_zamani"),
    ("Üretim Bitiş Zamanı", "bitis_zamani"),
    ("Teslim Zamanı", "teslim_zamani"),
    ("Operatör", "operator"),
    ("Not", "aciklama"),
    ("Son Güncelleme", "guncelleme"),
    ("Listede", "aktif"),
    ("Kaynakta Aktif", "source_active"),
    ("Admin Gizli", "admin_gizli"),
    ("Kaynak", "kaynak"),
    ("Anahtar", "anahtar"),
]

LOG_ALANLARI = [
    ("Zaman", "zaman"),
    ("Kullanıcı", "kullanici"),
    ("Rol", "rol"),
    ("İşlem", "islem"),
    ("Talep NO", "talep_no"),
    ("Kart Stok No", "stok_no"),
    ("Adet", "adet"),
    ("Detay", "detay"),
]

YUKLEME_ALANLARI = [
    ("Zaman", "zaman"),
    ("Kullanıcı", "kullanici"),
    ("Dosya", "dosya"),
    ("Okunan Satır", "satir"),
    ("Yeni Kart", "yeni"),
    ("Güncellenen", "guncellenen"),
    ("Kaynakta Olmayan", "pasife_alinan"),
    ("Uyarı", "uyari"),
]

SAYISAL_ALANLAR = {
    "id",
    "sira",
    "toplam_adet",
    "baslangic_adet",
    "tamamlanan_adet",
    "aktif",
    "source_active",
    "admin_gizli",
    "adet",
    "satir",
    "yeni",
    "guncellenen",
    "pasife_alinan",
    "uyari",
}

ZORUNLU_KART_ALANLARI = {
    "id",
    "talep_no",
    "stok_no",
    "toplam_adet",
    "tamamlanan_adet",
    "anahtar",
}


# ---------------------------------------------------------------------------
# In-memory state
# ---------------------------------------------------------------------------

_kilit = threading.RLock()
_kartlar: list[dict] = []
_loglar: list[dict] = []
_yuklemeler: list[dict] = []

LOG_SINIRI = 20_000
LOG_SAKLA = 5_000

BASLIK_DOLGU = PatternFill("solid", fgColor="0F2027")
BASLIK_YAZI = Font(name="Arial", bold=True, color="FFFFFF", size=11)
GOVDE_YAZI = Font(name="Arial", size=10)


# ---------------------------------------------------------------------------
# Genel yardımcılar
# ---------------------------------------------------------------------------

def simdi() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def bugun() -> str:
    return date.today().strftime("%Y-%m-%d")


def _sayi(deger, varsayilan=0) -> int:
    try:
        return int(float(deger))
    except (TypeError, ValueError):
        return varsayilan


def _temiz_metin(deger) -> str:
    return str(deger or "").strip()


def _durum_normalize(deger):
    """Dört gerçek workflow durumunu normalize eder; boş değer None olarak kalır."""
    metin = _temiz_metin(deger)
    if not metin:
        return None

    sade = (
        metin.upper()
        .replace("İ", "I")
        .replace("Ğ", "G")
        .replace("Ü", "U")
        .replace("Ş", "S")
        .replace("Ö", "O")
        .replace("Ç", "C")
    )
    sade = re.sub(r"\s+", " ", sade).strip()

    esleme = {
        "PLANA ALINDI": PLANA_ALINDI,
        "DIZGIDE": DIZGIDE,
        "HAZIR": HAZIR,
        "TESLIM EDILDI": TESLIM_EDILDI,
    }
    return esleme.get(sade)


def tarih_coz(deger):
    """Excel veya kullanıcı girdisini YYYY-MM-DD biçimine çevirir."""
    if deger is None:
        return None

    if isinstance(deger, datetime):
        return deger.strftime("%Y-%m-%d")
    if isinstance(deger, date):
        return deger.strftime("%Y-%m-%d")

    if isinstance(deger, (int, float)) and not isinstance(deger, bool):
        try:
            sonuc = from_excel(deger)
            if isinstance(sonuc, (datetime, date)):
                return sonuc.strftime("%Y-%m-%d")
        except (TypeError, ValueError, OverflowError):
            return None

    metin = str(deger).strip()
    if not metin or metin.upper() in {"-", "YOK", "N/A", "NONE"}:
        return None

    if re.fullmatch(r"\d+(?:\.\d+)?", metin):
        try:
            sonuc = from_excel(float(metin))
            if isinstance(sonuc, (datetime, date)):
                return sonuc.strftime("%Y-%m-%d")
        except (TypeError, ValueError, OverflowError):
            pass

    for kalip in (
        "%Y-%m-%d %H:%M:%S",
        "%d.%m.%Y %H:%M:%S",
        "%d/%m/%Y %H:%M:%S",
        "%d-%m-%Y %H:%M:%S",
        "%Y-%m-%d",
        "%d.%m.%Y",
        "%d/%m/%Y",
        "%Y/%m/%d",
        "%d-%m-%Y",
    ):
        try:
            return datetime.strptime(metin, kalip).strftime("%Y-%m-%d")
        except ValueError:
            continue
    return None


def gun_farki(a, b):
    """a - b gün farkını döndürür."""
    if not a or not b:
        return None
    try:
        t1 = datetime.strptime(str(a)[:10], "%Y-%m-%d").date()
        t2 = datetime.strptime(str(b)[:10], "%Y-%m-%d").date()
    except ValueError:
        return None
    return (t1 - t2).days


def _kart_ref(kart_id: int):
    return next((kart for kart in _kartlar if kart.get("id") == kart_id), None)


def _operasyonda_gorunur_mu(kart: dict) -> bool:
    return (
        kart.get("aktif", 1) == 1
        and kart.get("admin_gizli", 0) != 1
        and kart.get("durum") in GECERLI_DURUMLAR
    )


def _yonetimde_gorunur_mu(kart: dict) -> bool:
    return kart.get("aktif", 1) == 1 and kart.get("admin_gizli", 0) != 1


# ---------------------------------------------------------------------------
# Dosya okuma/yazma
# ---------------------------------------------------------------------------

def _oku(dosya, alanlar, zorunlu_alanlar=frozenset()):
    if not os.path.exists(dosya):
        return []

    try:
        wb = openpyxl.load_workbook(dosya, data_only=True, read_only=True)
    except Exception as exc:  # noqa: BLE001
        raise VeriDogrulamaHatasi(
            f"'{os.path.basename(dosya)}' açılamadı: {exc}"
        ) from exc

    try:
        ws = wb[wb.sheetnames[0]]
        satirlar = list(ws.iter_rows(values_only=True))
    finally:
        wb.close()

    if not satirlar:
        return []

    basliklar = [str(h or "").strip() for h in satirlar[0]]
    yerlesim = {
        alan: basliklar.index(baslik)
        for baslik, alan in alanlar
        if baslik in basliklar
    }

    eksik = sorted(alan for alan in zorunlu_alanlar if alan not in yerlesim)
    if eksik:
        ters = {alan: baslik for baslik, alan in alanlar}
        eksik_adlar = ", ".join(ters.get(alan, alan) for alan in eksik)
        raise VeriDogrulamaHatasi(
            f"'{os.path.basename(dosya)}' zorunlu sütunları eksik: {eksik_adlar}"
        )

    kayitlar = []
    for satir in satirlar[1:]:
        if not any(hucre not in (None, "") for hucre in satir):
            continue

        kayit = {}
        for _, alan in alanlar:
            index = yerlesim.get(alan)
            deger = satir[index] if index is not None and index < len(satir) else None

            if alan in SAYISAL_ALANLAR:
                kayit[alan] = _sayi(deger, 0) if deger not in (None, "") else None
            elif isinstance(deger, datetime):
                kayit[alan] = deger.strftime("%Y-%m-%d %H:%M:%S")
            else:
                kayit[alan] = str(deger).strip() if deger not in (None, "") else None
        kayitlar.append(kayit)

    return kayitlar

def _excel_hucre_yaz(hucre, deger):
    """'=' ile başlayan kullanıcı metninin Excel formülüne dönüşmesini engeller."""
    hucre.value = deger

    if isinstance(deger, str) and deger.startswith("="):
        hucre.data_type = "s"

def _workbook_uret(alanlar, kayitlar, sayfa_adi):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sayfa_adi
    ws.append([baslik for baslik, _ in alanlar])

    for hucre in ws[1]:
        hucre.fill = BASLIK_DOLGU
        hucre.font = BASLIK_YAZI
        hucre.alignment = Alignment(horizontal="center", vertical="center")

    for kayit in kayitlar:
        satir_no = ws.max_row + 1

        for sutun_no, (_, alan) in enumerate(alanlar, start=1):
            _excel_hucre_yaz(
                ws.cell(row=satir_no, column=sutun_no),
                kayit.get(alan),
            )

    for satir in ws.iter_rows(min_row=2):
        for hucre in satir:
            hucre.font = GOVDE_YAZI

    for sutun, (baslik, alan) in enumerate(alanlar, start=1):
        en = max(
            [len(baslik), 10]
            + [len(str(kayit.get(alan) or "")) for kayit in kayitlar[:300]]
        )
        ws.column_dimensions[get_column_letter(sutun)].width = min(42, en + 3)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    return wb


def _diske_zorla(yol):
    """Temp dosya içeriğinin OS page cache'ten diske inmesini zorlar."""
    try:
        fd = os.open(yol, os.O_RDONLY)
    except OSError:
        return
    try:
        os.fsync(fd)
    except OSError:
        pass
    finally:
        os.close(fd)


def _temp_yaz(hedef, alanlar, kayitlar, sayfa_adi):
    os.makedirs(os.path.dirname(hedef), exist_ok=True)
    temp = f"{hedef}.{uuid4().hex}.yeni"
    wb = _workbook_uret(alanlar, kayitlar, sayfa_adi)
    try:
        wb.save(temp)
    finally:
        wb.close()
    _diske_zorla(temp)
    return temp


def _coklu_yaz(dosyalar):
    """Birden fazla Excel dosyasını temp + replace + rollback yaklaşımıyla yazar."""
    os.makedirs(VERI_KLASORU, exist_ok=True)
    temps = []
    backups = {}
    degisen = []
    commit_basarili = False

    try:
        for hedef, alanlar, kayitlar, sayfa_adi in dosyalar:
            temps.append((hedef, _temp_yaz(hedef, alanlar, kayitlar, sayfa_adi)))

        for hedef, _ in temps:
            if os.path.exists(hedef):
                backup = f"{hedef}.{uuid4().hex}.txn.bak"
                shutil.copy2(hedef, backup)
                backups[hedef] = backup
            else:
                backups[hedef] = None

        for hedef, temp in temps:
            os.replace(temp, hedef)
            degisen.append(hedef)

        commit_basarili = True

    except Exception:
        for hedef in reversed(degisen):
            backup = backups.get(hedef)
            try:
                if backup and os.path.exists(backup):
                    os.replace(backup, hedef)
                    backups[hedef] = None
                elif os.path.exists(hedef):
                    os.remove(hedef)
            except OSError:
                pass
        raise
    finally:
        for _, temp in temps:
            if os.path.exists(temp):
                try:
                    os.remove(temp)
                except OSError:
                    pass
        if commit_basarili:
            for backup in backups.values():
                if backup and os.path.exists(backup):
                    try:
                        os.remove(backup)
                    except OSError:
                        pass


def _yaz(dosya, alanlar, kayitlar, sayfa_adi):
    _coklu_yaz([(dosya, alanlar, kayitlar, sayfa_adi)])


def _gunluk_yedek(dosya):
    if not os.path.exists(dosya):
        return
    os.makedirs(YEDEK_KLASORU, exist_ok=True)
    hedef = os.path.join(
        YEDEK_KLASORU,
        f"{date.today():%Y%m%d}_{os.path.basename(dosya)}",
    )
    if not os.path.exists(hedef):
        shutil.copy2(dosya, hedef)


ANLIK_YEDEK_SAKLA = 30
GUNLUK_YEDEK_GUN = 90
ANLIK_YEDEK_DESEN = re.compile(r"^\d{8}_\d{6}_")


def yedekleri_buda():
    """Yalnız PDGM naming pattern'li yedekleri sınırlar. Hata olursa sessizce geçer."""
    try:
        if not os.path.isdir(YEDEK_KLASORU):
            return

        anlik = []
        for ad in os.listdir(YEDEK_KLASORU):
            yol = os.path.join(YEDEK_KLASORU, ad)

            if os.path.isdir(yol) and ANLIK_YEDEK_DESEN.match(ad):
                try:
                    anlik.append((os.path.getmtime(yol), yol))
                except OSError:
                    pass
                continue

            if re.fullmatch(r"\d{8}_kartlar\.xlsx", ad, flags=re.IGNORECASE):
                try:
                    yas_gun = (
                        datetime.now() - datetime.fromtimestamp(os.path.getmtime(yol))
                    ).days
                    if yas_gun > GUNLUK_YEDEK_GUN:
                        os.remove(yol)
                except OSError:
                    pass

        anlik.sort(reverse=True)
        for _, yol in anlik[ANLIK_YEDEK_SAKLA:]:
            shutil.rmtree(yol, ignore_errors=True)
    except Exception:  # noqa: BLE001
        pass


def anlik_yedek(etiket: str = "once") -> str:
    """Kart/log/yükleme dosyalarının timestamp'li güvenlik kopyasını alır."""
    os.makedirs(YEDEK_KLASORU, exist_ok=True)
    damga = datetime.now().strftime("%Y%m%d_%H%M%S")
    guvenli = re.sub(r"[^0-9A-Za-z_-]+", "_", etiket or "yedek")
    klasor = os.path.join(YEDEK_KLASORU, f"{damga}_{guvenli}")
    os.makedirs(klasor, exist_ok=True)

    for dosya in (KARTLAR_DOSYA, LOG_DOSYA, YUKLEME_DOSYA):
        if os.path.exists(dosya):
            shutil.copy2(dosya, os.path.join(klasor, os.path.basename(dosya)))

    yedekleri_buda()
    return klasor


def yedekleri_getir(adet=12):
    os.makedirs(YEDEK_KLASORU, exist_ok=True)
    sonuc = []

    for ad in os.listdir(YEDEK_KLASORU):
        yol = os.path.join(YEDEK_KLASORU, ad)

        if os.path.isdir(yol):
            kart_dosyasi = os.path.join(yol, os.path.basename(KARTLAR_DOSYA))
            if not os.path.isfile(kart_dosyasi):
                continue
            try:
                mtime = os.path.getmtime(kart_dosyasi)
                boyut = os.path.getsize(kart_dosyasi)
            except OSError:
                continue

            parcalar = ad.split("_", 2)
            etiket = parcalar[2].replace("_", " ") if len(parcalar) >= 3 else "anlık yedek"
            sonuc.append(
                {
                    "ad": ad,
                    "tip": "Anlık",
                    "etiket": etiket,
                    "zaman": datetime.fromtimestamp(mtime).strftime("%d.%m.%Y %H:%M:%S"),
                    "boyut_kb": round(boyut / 1024, 1),
                    "_mtime": mtime,
                }
            )
            continue

        if not os.path.isfile(yol) or not re.fullmatch(
            r"\d{8}_kartlar\.xlsx", ad, flags=re.IGNORECASE
        ):
            continue

        try:
            mtime = os.path.getmtime(yol)
            boyut = os.path.getsize(yol)
        except OSError:
            continue

        sonuc.append(
            {
                "ad": ad,
                "tip": "Günlük",
                "etiket": "günlük otomatik yedek",
                "zaman": datetime.fromtimestamp(mtime).strftime("%d.%m.%Y %H:%M:%S"),
                "boyut_kb": round(boyut / 1024, 1),
                "_mtime": mtime,
            }
        )

    sonuc.sort(key=lambda kayit: kayit["_mtime"], reverse=True)
    for kayit in sonuc:
        kayit.pop("_mtime", None)
    return sonuc if adet is None else sonuc[:adet]


def _yedek_kart_dosyasi_bul(yedek_adi):
    yedek_adi = _temiz_metin(yedek_adi)
    if not yedek_adi:
        raise IsKuralHatasi("Yedek seçilmedi.")
    if os.path.basename(yedek_adi) != yedek_adi:
        raise IsKuralHatasi("Geçersiz yedek adı.")

    yedek_kok = os.path.abspath(YEDEK_KLASORU)
    yol = os.path.abspath(os.path.join(yedek_kok, yedek_adi))
    try:
        if os.path.commonpath([yedek_kok, yol]) != yedek_kok:
            raise IsKuralHatasi("Geçersiz yedek yolu.")
    except ValueError as exc:
        raise IsKuralHatasi("Geçersiz yedek yolu.") from exc

    if os.path.isdir(yol):
        aday = os.path.join(yol, os.path.basename(KARTLAR_DOSYA))
        if os.path.isfile(aday):
            return aday
        raise IsKuralHatasi("Seçilen yedekte kartlar.xlsx bulunamadı.")

    if os.path.isfile(yol) and re.fullmatch(
        r"\d{8}_kartlar\.xlsx", yedek_adi, flags=re.IGNORECASE
    ):
        return yol

    raise IsKuralHatasi("Seçilen yedek kullanılamıyor.")


# ---------------------------------------------------------------------------
# Process lock
# ---------------------------------------------------------------------------



def _process_kilit_sahibi():
    if not os.path.exists(PROCESS_KILIT):
        return None

    try:
        with open(PROCESS_KILIT, encoding="utf-8") as f:
            metin = (f.read() or "").strip()
        ilk = metin.split("|", 1)[0].strip()
        return int(ilk) if ilk else None
    except (OSError, ValueError):
        return None


def _pid_calisiyor_mu(pid: int):
    """PID canlı mı? Karar verilemezse True (fail-closed: kilidi koru)."""
    if pid is None or pid <= 0:
        return False

    if os.name != "nt":
        try:
            os.kill(pid, 0)
            return True
        except ProcessLookupError:
            return False
        except PermissionError:
            return True
        except OSError:
            return True

    try:
        cikti = subprocess.run(
            ["tasklist", "/FI", f"PID eq {pid}", "/NH", "/FO", "CSV"],
            capture_output=True,
            text=True,
            timeout=10,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
    except Exception:
        return True

    if cikti.returncode != 0:
        return True

    return f'"{pid}"' in (cikti.stdout or "")


def _pid_python_mu(pid: int):
    """PID python/pythonw sürecine mi ait? Karar verilemezse True (fail-closed)."""
    if pid is None or pid <= 0:
        return False

    if os.name != "nt":
        try:
            import sys

            if pid == os.getpid():
                return True
            cmdline_yolu = f"/proc/{pid}/cmdline"
            if os.path.exists(cmdline_yolu):
                with open(cmdline_yolu, "rb") as f:
                    cmd = f.read().decode("utf-8", errors="ignore").lower()
                return "python" in cmd
            return True
        except OSError:
            return True

    try:
        cikti = subprocess.run(
            ["tasklist", "/FI", f"PID eq {pid}", "/NH", "/FO", "CSV"],
            capture_output=True,
            text=True,
            timeout=10,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
    except Exception:
        return True

    if cikti.returncode != 0:
        return True

    satir = (cikti.stdout or "").strip().lower()
    if f'"{pid}"' not in satir and str(pid) not in satir:
        return False
    return "python.exe" in satir or "pythonw.exe" in satir


def process_kilidi_al():
    """Aynı data klasörünü ikinci PDGM process'inin açmasını atomik olarak engeller.

    Not:
    - Windows'ta os.kill(pid, 0) yerine tasklist kullanılır.
    - Lock dosyası O_EXCL ile atomik oluşturulur.
    - Stale lock: PID ölüyse veya PID canlı ama python değilse (PID reuse) devralınır.
    - Stale reclaim os.replace ile atomik yapılır (remove+O_EXCL yarışı yok).
    - PID canlı ve python ise reddedilir. tasklist başarısızsa fail-closed.
    """
    os.makedirs(VERI_KLASORU, exist_ok=True)
    mevcut_pid = _process_kilit_sahibi()

    if mevcut_pid == os.getpid():
        return

    if os.path.exists(PROCESS_KILIT):
        canli = mevcut_pid is not None and _pid_calisiyor_mu(mevcut_pid)
        python_sureci = canli and _pid_python_mu(mevcut_pid)

        if canli and python_sureci:
            raise RuntimeError(
                f"data/ klasörü başka bir PDGM process tarafından kilitli "
                f"(PID {mevcut_pid}) ve bu process ŞU AN ÇALIŞIYOR. "
                "İkinci sunucu açmayın."
            )

        sahip = mevcut_pid if mevcut_pid is not None else "bilinmiyor"
        if canli and not python_sureci:
            print(
                f"UYARI: data/sunucu.lock PID {sahip} başka bir uygulamaya ait "
                "(PID reuse). Kalıntı kilit devralınıyor."
            )
        else:
            print(
                f"UYARI: data/sunucu.lock artık çalışmayan bir process'e ait "
                f"(PID {sahip}). Kalıntı kilit devralınıyor."
            )

        # Atomik reclaim: remove+O_EXCL yarışını önlemek için stale lock'u
        # benzersiz bir isme taşı. İki process aynı anda denerse yalnız biri
        # os.replace kazanır; diğeri FileNotFoundError alır ve O_EXCL'de kaybeder.
        stale_yol = f"{PROCESS_KILIT}.stale.{uuid4().hex}"
        try:
            os.replace(PROCESS_KILIT, stale_yol)
        except FileNotFoundError:
            stale_yol = None
        except OSError as exc:
            raise RuntimeError(
                "data/sunucu.lock devralınamadı: "
                f"{exc}. Dosyayı manuel silip tekrar deneyin."
            ) from exc
    else:
        stale_yol = None

    bayraklar = os.O_WRONLY | os.O_CREAT | os.O_EXCL

    try:
        fd = os.open(PROCESS_KILIT, bayraklar)
    except FileExistsError as exc:
        if stale_yol:
            try:
                os.remove(stale_yol)
            except OSError:
                pass
        sahip = _process_kilit_sahibi()
        raise RuntimeError(
            f"data/ klasörü başka bir PDGM process tarafından kilitlendi "
            f"(PID {sahip if sahip is not None else 'bilinmiyor'})."
        ) from exc

    try:
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            f.write(f"{os.getpid()}|{datetime.now():%Y-%m-%d %H:%M:%S}")
            f.flush()
            os.fsync(f.fileno())
    except Exception:
        try:
            os.remove(PROCESS_KILIT)
        except OSError:
            pass
        if stale_yol:
            try:
                os.remove(stale_yol)
            except OSError:
                pass
        raise

    if stale_yol:
        try:
            os.remove(stale_yol)
        except OSError:
            pass

    if getattr(process_kilidi_al, "_atexit_bagli", False):
        return

    def _birak():
        try:
            if _process_kilit_sahibi() == os.getpid():
                os.remove(PROCESS_KILIT)
        except OSError:
            pass

    atexit.register(_birak)
    process_kilidi_al._atexit_bagli = True


# ---------------------------------------------------------------------------
# Kart normalizasyonu ve doğrulama
# ---------------------------------------------------------------------------

def _kart_normalize(kart: dict) -> dict:
    kart = dict(kart)
    kart["id"] = _sayi(kart.get("id"), 0)
    kart["sira"] = _sayi(kart.get("sira"), 0) or None
    kart["toplam_adet"] = _sayi(kart.get("toplam_adet"), 1) or 1
    kart["baslangic_adet"] = _sayi(kart.get("baslangic_adet"), 0)
    kart["tamamlanan_adet"] = _sayi(kart.get("tamamlanan_adet"), 0)
    kart["aktif"] = 0 if kart.get("aktif") == 0 else 1
    kart["source_active"] = 0 if kart.get("source_active") == 0 else 1
    kart["admin_gizli"] = 1 if kart.get("admin_gizli") == 1 else 0
    kart["kaynak"] = _temiz_metin(kart.get("kaynak")) or "EXCEL"
    kart["durum"] = _durum_normalize(kart.get("durum"))

    for alan in ("plan_baslama", "plan_teslim", "gerceklesen_teslim"):
        ham = kart.get(alan)
        if ham in (None, ""):
            kart[alan] = None
            continue
        cozulmus = tarih_coz(ham)
        if not cozulmus:
            raise VeriDogrulamaHatasi(
                f"Kart {kart.get('id') or '?'}: {alan} geçerli bir tarih değil: {ham!r}"
            )
        kart[alan] = cozulmus

    return kart


def _kart_dogrula(kart: dict):
    kart_id = _sayi(kart.get("id"), -1)
    toplam = _sayi(kart.get("toplam_adet"), 0)
    tamam = _sayi(kart.get("tamamlanan_adet"), 0)
    durum = kart.get("durum")

    if kart_id < 1:
        raise VeriDogrulamaHatasi("Kart ID pozitif tam sayı olmalı.")
    if not _temiz_metin(kart.get("talep_no")):
        raise VeriDogrulamaHatasi(f"Kart {kart_id}: Talep NO boş olamaz.")
    if not _temiz_metin(kart.get("stok_no")):
        raise VeriDogrulamaHatasi(f"Kart {kart_id}: Kart Stok No boş olamaz.")
    if toplam < 1:
        raise VeriDogrulamaHatasi(f"Kart {kart_id}: Toplam Adet en az 1 olmalı.")
    if tamam < 0 or tamam > toplam:
        raise VeriDogrulamaHatasi(
            f"Kart {kart_id}: Tamamlanan Adet ({tamam}) 0 ile Toplam Adet ({toplam}) arasında olmalı."
        )
    if durum is not None and durum not in GECERLI_DURUMLAR:
        raise VeriDogrulamaHatasi(
            f"Kart {kart_id}: Geçersiz durum '{durum}'. "
            f"Geçerli durumlar: {', '.join(sorted(GECERLI_DURUMLAR))}."
        )
    if durum in (None, PLANA_ALINDI) and tamam != 0:
        etiket = "Durumu boş" if durum is None else PLANA_ALINDI
        raise VeriDogrulamaHatasi(
            f"Kart {kart_id}: {etiket} kartta Tamamlanan Adet 0 olmalı."
        )
    if durum in (HAZIR, TESLIM_EDILDI) and tamam != toplam:
        raise VeriDogrulamaHatasi(
            f"Kart {kart_id}: {durum} durumunda Tamamlanan Adet Toplam Adet'e eşit olmalı."
        )
    if durum == TESLIM_EDILDI and not kart.get("gerceklesen_teslim"):
        raise VeriDogrulamaHatasi(
            f"Kart {kart_id}: TESLİM EDİLDİ durumunda Gerçekleşen Teslim tarihi zorunlu."
        )
    if kart.get("plan_baslama") and kart.get("plan_teslim"):
        if kart["plan_baslama"] > kart["plan_teslim"]:
            raise VeriDogrulamaHatasi(
                f"Kart {kart_id}: Plan başlangıç tarihi plan teslim tarihinden sonra olamaz."
            )
    if not _temiz_metin(kart.get("anahtar")):
        raise VeriDogrulamaHatasi(f"Kart {kart_id}: Anahtar boş olamaz.")


def _kart_listesi_dogrula(kartlar):
    idler = set()
    anahtarlar = set()

    for kart in kartlar:
        _kart_dogrula(kart)
        if kart["id"] in idler:
            raise VeriDogrulamaHatasi(f"Tekrarlanan kart ID: {kart['id']}")
        if kart["anahtar"] in anahtarlar:
            raise VeriDogrulamaHatasi(f"Tekrarlanan kart anahtarı: {kart['anahtar']}")
        idler.add(kart["id"])
        anahtarlar.add(kart["anahtar"])


# ---------------------------------------------------------------------------
# Başlangıç ve reload
# ---------------------------------------------------------------------------

def _baslatma_uyarisi_yaz(mesaj: str) -> None:
    """Açılış uyarılarını konsola basar ve data/BASLATMA_HATASI.txt'ye ekler."""
    print(mesaj)
    yol = os.path.join(VERI_KLASORU, "BASLATMA_HATASI.txt")
    try:
        with open(yol, "a", encoding="utf-8") as f:
            f.write(f"{datetime.now():%Y-%m-%d %H:%M:%S}  {mesaj}\n")
    except OSError:
        pass


def _bozuk_dosyayi_kenara_al(dosya, alanlar):
    """Okunamayan yardımcı dosyayı yeniden adlandırıp boş liste döner.

    Yalnız islem_logu / yuklemeler için kullanılır. Kart verisi bu yolla sıfırlanmaz.
    """
    try:
        return _oku(dosya, alanlar)
    except VeriDogrulamaHatasi as exc:
        if os.path.exists(dosya):
            damga = datetime.now().strftime("%Y%m%d_%H%M%S")
            bozuk = f"{dosya}.bozuk_{damga}"
            try:
                os.replace(dosya, bozuk)
            except OSError:
                bozuk = dosya
            _baslatma_uyarisi_yaz(
                f"UYARI: '{os.path.basename(dosya)}' okunamadı ({exc}). "
                f"Dosya '{os.path.basename(bozuk)}' olarak kenara alındı; "
                "boş liste ile devam ediliyor. Kart verisi etkilenmedi."
            )
        else:
            _baslatma_uyarisi_yaz(
                f"UYARI: '{os.path.basename(dosya)}' okunamadı ({exc}). "
                "Boş liste ile devam ediliyor."
            )
        return []


def kur():
    global _kartlar, _loglar, _yuklemeler

    with _kilit:
        os.makedirs(VERI_KLASORU, exist_ok=True)

        kartlar = _oku(
            KARTLAR_DOSYA,
            KART_ALANLARI,
            ZORUNLU_KART_ALANLARI if os.path.exists(KARTLAR_DOSYA) else frozenset(),
        )
        kartlar = [_kart_normalize(kart) for kart in kartlar]
        if kartlar:
            _kart_listesi_dogrula(kartlar)

        _kartlar = kartlar
        _loglar = _bozuk_dosyayi_kenara_al(LOG_DOSYA, LOG_ALANLARI)
        _yuklemeler = _bozuk_dosyayi_kenara_al(YUKLEME_DOSYA, YUKLEME_ALANLARI)

        eksikler = []
        if not os.path.exists(KARTLAR_DOSYA):
            eksikler.append((KARTLAR_DOSYA, KART_ALANLARI, _kartlar, "Kartlar"))
        if not os.path.exists(LOG_DOSYA):
            eksikler.append((LOG_DOSYA, LOG_ALANLARI, _loglar, "İşlem Logu"))
        if not os.path.exists(YUKLEME_DOSYA):
            eksikler.append((YUKLEME_DOSYA, YUKLEME_ALANLARI, _yuklemeler, "Yüklemeler"))
        if eksikler:
            _coklu_yaz(eksikler)


def kartlari_diskten_yeniden_yukle():
    """Manuel Excel müdahalesinden sonra kartlar.xlsx'i validate ederek tekrar yükler."""
    global _kartlar

    with _kilit:
        yeni = _oku(KARTLAR_DOSYA, KART_ALANLARI, ZORUNLU_KART_ALANLARI)
        yeni = [_kart_normalize(kart) for kart in yeni]
        _kart_listesi_dogrula(yeni)
        _kartlar = yeni
        return len(_kartlar)


def yeniden_yukle():
    return kartlari_diskten_yeniden_yukle()


def _kartlari_kaydet():
    _gunluk_yedek(KARTLAR_DOSYA)
    _yaz(KARTLAR_DOSYA, KART_ALANLARI, _kartlar, "Kartlar")


# ---------------------------------------------------------------------------
# Log ve yükleme geçmişi
# ---------------------------------------------------------------------------

def _log_kaydi(kullanici, rol, islem, talep_no="", stok_no="", adet=None, detay=""):
    return {
        "zaman": simdi(),
        "kullanici": kullanici,
        "rol": rol,
        "islem": islem,
        "talep_no": talep_no,
        "stok_no": stok_no,
        "adet": adet,
        "detay": detay,
    }


def _log_arsivle_gerekirse():
    global _loglar

    if len(_loglar) <= LOG_SINIRI:
        return

    os.makedirs(YEDEK_KLASORU, exist_ok=True)
    arsivlenecek = _loglar[:-LOG_SAKLA]
    arsiv = os.path.join(
        YEDEK_KLASORU,
        f"{datetime.now():%Y%m%d_%H%M%S}_islem_logu_arsiv.xlsx",
    )
    _yaz(arsiv, LOG_ALANLARI, arsivlenecek, "İşlem Logu")
    _loglar = _loglar[-LOG_SAKLA:]


def log_ekle(kullanici, rol, islem, talep_no="", stok_no="", adet=None, detay=""):
    global _loglar

    with _kilit:
        eski = copy.deepcopy(_loglar)
        try:
            _loglar.append(
                _log_kaydi(kullanici, rol, islem, talep_no, stok_no, adet, detay)
            )
            _log_arsivle_gerekirse()
            _yaz(LOG_DOSYA, LOG_ALANLARI, _loglar, "İşlem Logu")
        except Exception:
            _loglar = eski
            raise


def loglari_getir(adet=None):
    with _kilit:
        secim = list(reversed(_loglar))
        if adet:
            secim = secim[:adet]
        return copy.deepcopy(secim)


def yuklemeleri_getir(adet=None):
    with _kilit:
        secim = list(reversed(_yuklemeler))
        if adet:
            secim = secim[:adet]
        return copy.deepcopy(secim)


# ---------------------------------------------------------------------------
# Görünüm hesapları
# ---------------------------------------------------------------------------

def durum_bilgisi(kart):
    durum = kart.get("durum")
    plan_baslama = kart.get("plan_baslama")
    plan_teslim = kart.get("plan_teslim")
    bugun_iso = bugun()

    bilgi = {
        "rozet": durum or "DURUMU EKSİK",
        "renk": "notr",
        "sapma": None,
        "kalan": None,
        "zaman_yuzde": 0,
        "plan_gun": gun_farki(plan_teslim, plan_baslama),
    }

    if durum is None:
        bilgi["rozet"] = "DURUMU EKSİK"
        bilgi["renk"] = "uyari"
        return bilgi

    if durum == TESLIM_EDILDI:
        teslim = kart.get("gerceklesen_teslim") or str(kart.get("teslim_zamani") or "")[:10]
        sapma = gun_farki(teslim, plan_teslim)
        bilgi["sapma"] = sapma
        bilgi["zaman_yuzde"] = 100
        if sapma is None:
            bilgi["rozet"], bilgi["renk"] = "TESLİM EDİLDİ", "iyi"
        elif sapma > 0:
            bilgi["rozet"], bilgi["renk"] = f"GEÇ TESLİM (+{sapma} gün)", "kotu"
        else:
            bilgi["rozet"], bilgi["renk"] = "ZAMANINDA TESLİM", "iyi"
        return bilgi

    if durum == HAZIR:
        gecikme = gun_farki(bugun_iso, plan_teslim)
        bilgi["zaman_yuzde"] = 100
        if plan_teslim and gecikme is not None and gecikme > 0:
            bilgi["sapma"] = gecikme
            bilgi["rozet"], bilgi["renk"] = f"HAZIR · TESLİM BEKLİYOR (+{gecikme} gün)", "kotu"
        else:
            bilgi["rozet"], bilgi["renk"] = "HAZIR · TESLİM BEKLİYOR", "iyi"
        return bilgi

    if durum == DIZGIDE:
        if kart.get("tamamlanan_adet", 0) >= kart.get("toplam_adet", 1):
            bilgi["rozet"], bilgi["renk"] = "ÜRETİM BİTTİ · HAZIRA ALIN", "uyari"
            bilgi["zaman_yuzde"] = 100
            return bilgi

        kalan = gun_farki(plan_teslim, bugun_iso)
        bilgi["kalan"] = kalan
        baslangic = str(kart.get("baslama_zamani") or plan_baslama or bugun_iso)[:10]
        gecen = gun_farki(bugun_iso, baslangic) or 0
        plan_gun = bilgi["plan_gun"]

        if plan_gun and plan_gun > 0:
            bilgi["zaman_yuzde"] = max(0, min(140, round(gecen / plan_gun * 100)))
        else:
            bilgi["zaman_yuzde"] = 100 if kalan is not None and kalan < 0 else 50

        if kalan is None:
            bilgi["rozet"], bilgi["renk"] = "DİZGİDE", "uyari"
        elif kalan < 0:
            bilgi["sapma"] = -kalan
            bilgi["rozet"], bilgi["renk"] = f"SÜRE AŞILDI ({-kalan} gün)", "kotu"
        elif kalan <= 1:
            bilgi["rozet"] = "SON GÜN" if kalan == 0 else "SON 1 GÜN"
            bilgi["renk"] = "uyari"
        else:
            bilgi["rozet"], bilgi["renk"] = f"PLANINDA ({kalan} gün var)", "iyi"
        return bilgi

    gecikme = gun_farki(bugun_iso, plan_baslama)
    if plan_baslama and gecikme is not None and gecikme > 0:
        bilgi["rozet"], bilgi["renk"], bilgi["sapma"] = (
            f"BAŞLAMADI (+{gecikme} gün)",
            "kotu",
            gecikme,
        )
    elif plan_baslama and gecikme == 0:
        bilgi["rozet"], bilgi["renk"] = "BUGÜN BAŞLAMALI", "uyari"
    else:
        bilgi["rozet"], bilgi["renk"] = PLANA_ALINDI, "notr"
    return bilgi


def kart_gorunumu(kart):
    d = copy.deepcopy(kart)
    d.update(durum_bilgisi(kart))
    d["toplam_adet"] = d.get("toplam_adet") or 1
    d["tamamlanan_adet"] = d.get("tamamlanan_adet") or 0
    d["baslangic_adet"] = d.get("baslangic_adet") or 0
    d["kalan_adet"] = max(0, d["toplam_adet"] - d["tamamlanan_adet"])
    d["adet_yuzde"] = min(100, round(d["tamamlanan_adet"] / d["toplam_adet"] * 100))
    d["gorunur"] = _operasyonda_gorunur_mu(d)
    d["kaynakta_yok"] = d.get("source_active", 1) != 1
    d["is_durumu"] = d.get("durum") or "DURUMU EKSİK"
    d["kaynak_durumu"] = _temiz_metin(d.get("excel_durum"))
    return d


def _kartlari_sirala(kartlar):
    kartlar.sort(
        key=lambda kart: (
            SIRALAMA.get(kart.get("durum"), 9),
            kart.get("plan_baslama") or "9999-12-31",
            kart.get("sira") or 999999,
            kart.get("id") or 0,
        )
    )
    return kartlar


def kartlari_getir(sadece_gorunen=True):
    with _kilit:
        secim = [
            kart for kart in _kartlar
            if not sadece_gorunen or _operasyonda_gorunur_mu(kart)
        ]
        kartlar = [kart_gorunumu(kart) for kart in secim]
    return _kartlari_sirala(kartlar)


def kartlari_yonetim_getir():
    """Admin tablosu için, durumu boş kartlar dahil, gizlenmemiş tüm aktif kartlar."""
    with _kilit:
        kartlar = [
            kart_gorunumu(kart)
            for kart in _kartlar
            if _yonetimde_gorunur_mu(kart)
        ]
    return _kartlari_sirala(kartlar)


def durumu_eksik_kartlari_getir():
    with _kilit:
        kartlar = [
            kart_gorunumu(kart)
            for kart in _kartlar
            if _yonetimde_gorunur_mu(kart) and not kart.get("durum")
        ]
    return _kartlari_sirala(kartlar)


def gizlenen_kartlari_getir():
    with _kilit:
        kartlar = [
            kart_gorunumu(kart)
            for kart in _kartlar
            if kart.get("aktif", 1) == 1 and kart.get("admin_gizli", 0) == 1
        ]
    kartlar.sort(key=lambda kart: (kart.get("guncelleme") or "", kart.get("id") or 0), reverse=True)
    return kartlar


def kart_getir(kart_id):
    kart_id = _sayi(kart_id, -1)
    with _kilit:
        kart = _kart_ref(kart_id)
        return kart_gorunumu(kart) if kart else None


def kart_bul(anahtar):
    with _kilit:
        for kart in _kartlar:
            if kart.get("anahtar") == anahtar:
                return copy.deepcopy(kart)
    return None


def yeni_kimlik():
    with _kilit:
        return max([_sayi(kart.get("id"), 0) for kart in _kartlar] or [0]) + 1


# ---------------------------------------------------------------------------
# Kart + log commit yardımcıları
# ---------------------------------------------------------------------------

def _kart_log_commit():
    _gunluk_yedek(KARTLAR_DOSYA)
    _coklu_yaz(
        [
            (KARTLAR_DOSYA, KART_ALANLARI, _kartlar, "Kartlar"),
            (LOG_DOSYA, LOG_ALANLARI, _loglar, "İşlem Logu"),
        ]
    )


def _atomik_kart_islemi(islem):
    """RAM state rollback kalıbını tek yerde tutar.

    Kartlar in-place mutasyon gördüğü için deepcopy zorunlu.
    Loglara yalnız append yapıldığı için uzunluk + del yeterlidir.
    """
    global _kartlar

    eski_kartlar = copy.deepcopy(_kartlar)
    log_sayisi = len(_loglar)
    try:
        sonuc = islem()
        _kart_listesi_dogrula(_kartlar)
        _kart_log_commit()
        return sonuc
    except Exception:
        _kartlar = eski_kartlar
        del _loglar[log_sayisi:]
        raise


# ---------------------------------------------------------------------------
# Operatör workflow işlemleri
# ---------------------------------------------------------------------------

def kart_baslat(kart_id, adet, kullanici, rol, aciklama=""):
    kart_id = _sayi(kart_id, -1)

    with _kilit:
        kart = _kart_ref(kart_id)
        if not kart:
            raise KartBulunamadi("Kart bulunamadı.")
        if not _operasyonda_gorunur_mu(kart):
            raise IsKuralHatasi("Bu kart operasyon ekranında aktif değil.")
        if kart["durum"] != PLANA_ALINDI:
            raise IsKuralHatasi("Yalnız PLANA ALINDI durumundaki kart DİZGİDE'ye alınabilir.")

        if adet in (None, ""):
            adet = kart["toplam_adet"]
        try:
            adet = int(adet)
        except (TypeError, ValueError) as exc:
            raise ValueError("Adet sayı olmalı.") from exc
        if adet < 1 or adet > kart["toplam_adet"]:
            raise IsKuralHatasi(f"Başlatılacak adet 1 ile {kart['toplam_adet']} arasında olmalı.")

        def islem():
            kart.update(
                durum=DIZGIDE,
                baslangic_adet=adet,
                baslama_zamani=kart.get("baslama_zamani") or simdi(),
                bitis_zamani=None,
                teslim_zamani=None,
                gerceklesen_teslim=None,
                operator=kullanici,
                aciklama=aciklama or kart.get("aciklama"),
                guncelleme=simdi(),
            )
            _loglar.append(
                _log_kaydi(
                    kullanici,
                    rol,
                    "DİZGİYE ALINDI",
                    kart.get("talep_no") or "",
                    kart.get("stok_no") or "",
                    adet,
                    aciklama or f"{adet} adet dizgiye alındı",
                )
            )
            return kart_gorunumu(kart)

        return _atomik_kart_islemi(islem)


def kart_bitir(kart_id, adet, kullanici, rol, aciklama=""):
    kart_id = _sayi(kart_id, -1)

    with _kilit:
        kart = _kart_ref(kart_id)
        if not kart:
            raise KartBulunamadi("Kart bulunamadı.")
        if kart.get("durum") != DIZGIDE:
            raise IsKuralHatasi("Tamamlanan adet yalnız DİZGİDE durumundaki karta girilebilir.")

        kalan = kart["toplam_adet"] - kart["tamamlanan_adet"]
        if kalan <= 0:
            raise IsKuralHatasi("Üretim adedi zaten tamamlandı. Kartı HAZIR durumuna alın.")

        try:
            adet = int(adet)
        except (TypeError, ValueError) as exc:
            raise ValueError("Adet sayı olmalı.") from exc
        if adet < 1 or adet > kalan:
            raise IsKuralHatasi(f"Adet 1 ile {kalan} arasında olmalı.")

        def islem():
            yeni_toplam = kart["tamamlanan_adet"] + adet
            uretim_bitti = yeni_toplam == kart["toplam_adet"]

            kart.update(
                tamamlanan_adet=yeni_toplam,
                bitis_zamani=simdi() if uretim_bitti else kart.get("bitis_zamani"),
                operator=kullanici,
                aciklama=aciklama or kart.get("aciklama"),
                guncelleme=simdi(),
            )
            _loglar.append(
                _log_kaydi(
                    kullanici,
                    rol,
                    "ÜRETİM ADEDİ TAMAMLANDI" if uretim_bitti else "KISMİ ÜRETİM",
                    kart.get("talep_no") or "",
                    kart.get("stok_no") or "",
                    adet,
                    aciklama or f"{yeni_toplam}/{kart['toplam_adet']} adet tamamlandı",
                )
            )

            mesaj = (
                "Üretim adedi tamamlandı. Kart HAZIR'a otomatik alınmadı."
                if uretim_bitti
                else f"{yeni_toplam}/{kart['toplam_adet']} adet tamamlandı."
            )
            return kart_gorunumu(kart), uretim_bitti, mesaj

        return _atomik_kart_islemi(islem)


def kart_hazirla(kart_id, kullanici, rol, aciklama=""):
    kart_id = _sayi(kart_id, -1)

    with _kilit:
        kart = _kart_ref(kart_id)
        if not kart:
            raise KartBulunamadi("Kart bulunamadı.")
        if kart.get("durum") != DIZGIDE:
            raise IsKuralHatasi("Yalnız DİZGİDE durumundaki kart HAZIR yapılabilir.")
        if kart.get("tamamlanan_adet", 0) != kart.get("toplam_adet", 0):
            raise IsKuralHatasi("Kart HAZIR yapılmadan önce üretim adedinin tamamı bitirilmelidir.")

        def islem():
            kart.update(
                durum=HAZIR,
                bitis_zamani=kart.get("bitis_zamani") or simdi(),
                operator=kullanici,
                aciklama=aciklama or kart.get("aciklama"),
                guncelleme=simdi(),
            )
            _loglar.append(
                _log_kaydi(
                    kullanici,
                    rol,
                    "HAZIR OLARAK İŞARETLENDİ",
                    kart.get("talep_no") or "",
                    kart.get("stok_no") or "",
                    kart.get("toplam_adet"),
                    aciklama or "Üretim tamamlandı; teslim alınmayı bekliyor.",
                )
            )
            return kart_gorunumu(kart)

        return _atomik_kart_islemi(islem)


def kart_teslim_et(kart_id, kullanici, rol, aciklama=""):
    kart_id = _sayi(kart_id, -1)

    with _kilit:
        kart = _kart_ref(kart_id)
        if not kart:
            raise KartBulunamadi("Kart bulunamadı.")
        if kart.get("durum") != HAZIR:
            raise IsKuralHatasi("Yalnız HAZIR durumundaki kart TESLİM EDİLDİ yapılabilir.")

        def islem():
            teslim_ani = simdi()
            kart.update(
                durum=TESLIM_EDILDI,
                gerceklesen_teslim=bugun(),
                teslim_zamani=teslim_ani,
                operator=kullanici,
                aciklama=aciklama or kart.get("aciklama"),
                guncelleme=teslim_ani,
            )
            _loglar.append(
                _log_kaydi(
                    kullanici,
                    rol,
                    "TESLİM EDİLDİ",
                    kart.get("talep_no") or "",
                    kart.get("stok_no") or "",
                    kart.get("toplam_adet"),
                    aciklama or f"Teslim tarihi: {bugun()}",
                )
            )
            return kart_gorunumu(kart)

        return _atomik_kart_islemi(islem)


def kart_not_guncelle(kart_id, aciklama, kullanici, rol):
    kart_id = _sayi(kart_id, -1)

    with _kilit:
        kart = _kart_ref(kart_id)
        if not kart:
            raise KartBulunamadi("Kart bulunamadı.")

        def islem():
            kart["aciklama"] = _temiz_metin(aciklama) or None
            kart["guncelleme"] = simdi()
            _loglar.append(
                _log_kaydi(
                    kullanici,
                    rol,
                    "NOT GÜNCELLENDİ",
                    kart.get("talep_no") or "",
                    kart.get("stok_no") or "",
                    detay=kart.get("aciklama") or "Not temizlendi",
                )
            )
            return kart_gorunumu(kart)

        return _atomik_kart_islemi(islem)


# ---------------------------------------------------------------------------
# Admin kart işlemleri
# ---------------------------------------------------------------------------

def _tarih_form_degeri(deger, alan_adi):
    if deger in (None, ""):
        return None
    sonuc = tarih_coz(deger)
    if not sonuc:
        raise IsKuralHatasi(f"{alan_adi} geçerli bir tarih olmalı.")
    return sonuc


def admin_kart_ekle(
    talep_no,
    stok_no,
    toplam_adet,
    kullanici,
    sira=None,
    talep_sahibi="",
    plan_hafta="",
    plan_baslama=None,
    plan_teslim=None,
    gerceklesen_teslim=None,
    pcb="",
    aciklama="",
):
    """Admin panelinden PLANA ALINDI durumunda manuel kart oluşturur."""
    global _kartlar, _loglar

    talep_no = _temiz_metin(talep_no)
    stok_no = _temiz_metin(stok_no)
    if not talep_no:
        raise IsKuralHatasi("Talep NO boş olamaz.")
    if not stok_no:
        raise IsKuralHatasi("Kart Stok No boş olamaz.")

    try:
        toplam = int(toplam_adet)
    except (TypeError, ValueError) as exc:
        raise ValueError("Toplam adet sayı olmalı.") from exc
    if toplam < 1:
        raise IsKuralHatasi("Toplam adet en az 1 olmalı.")

    if sira in (None, ""):
        sira_degeri = None
    else:
        try:
            sira_degeri = int(sira)
        except (TypeError, ValueError) as exc:
            raise ValueError("Sıra tam sayı olmalı.") from exc

    plan_baslama_iso = _tarih_form_degeri(plan_baslama, "Dizgi Başlama Tarihi")
    plan_teslim_iso = _tarih_form_degeri(plan_teslim, "Planlanan Teslim Tarihi")
    gerceklesen_iso = _tarih_form_degeri(gerceklesen_teslim, "Gerçekleşen Teslim Tarihi")

    if plan_baslama_iso and plan_teslim_iso and plan_baslama_iso > plan_teslim_iso:
        raise IsKuralHatasi("Dizgi Başlama Tarihi Planlanan Teslim Tarihinden sonra olamaz.")
    if gerceklesen_iso:
        raise IsKuralHatasi(
            "Yeni kart PLANA ALINDI durumunda başlar; Gerçekleşen Teslim Tarihi başlangıçta boş olmalı."
        )

    anahtar = f"{talep_no}|{stok_no}"

    with _kilit:
        if any(kart.get("anahtar") == anahtar for kart in _kartlar):
            raise IsKuralHatasi(f"Bu Talep NO + Kart Stok No zaten mevcut: {anahtar}")

        def islem():
            kart_id = max([_sayi(kart.get("id"), 0) for kart in _kartlar] or [0]) + 1
            kayit = {
                "id": kart_id,
                "sira": sira_degeri,
                "talep_no": talep_no,
                "stok_no": stok_no,
                "talep_sahibi": _temiz_metin(talep_sahibi) or None,
                "toplam_adet": toplam,
                "adet_metin": f"{toplam} ADET",
                "plan_hafta": _temiz_metin(plan_hafta) or None,
                "plan_baslama": plan_baslama_iso,
                "plan_teslim": plan_teslim_iso,
                "gerceklesen_teslim": None,
                "excel_durum": "MANUEL",
                "pcb": _temiz_metin(pcb) or None,
                "durum": PLANA_ALINDI,
                "baslangic_adet": 0,
                "tamamlanan_adet": 0,
                "baslama_zamani": None,
                "bitis_zamani": None,
                "teslim_zamani": None,
                "operator": None,
                "aciklama": _temiz_metin(aciklama) or None,
                "guncelleme": simdi(),
                "aktif": 1,
                "source_active": 1,
                "admin_gizli": 0,
                "kaynak": "MANUEL",
                "anahtar": anahtar,
            }
            _kartlar.append(kayit)
            _loglar.append(
                _log_kaydi(
                    kullanici,
                    "admin",
                    "MANUEL KART EKLENDİ",
                    talep_no,
                    stok_no,
                    toplam,
                    f"{toplam} adet · Durum: {PLANA_ALINDI}",
                )
            )
            return kart_gorunumu(kayit)

        return _atomik_kart_islemi(islem)


def admin_kart_duzenle(
    kart_id,
    durum,
    tamamlanan_adet,
    toplam_adet,
    aciklama,
    kullanici,
    plan_hafta=None,
    plan_baslama=None,
    plan_teslim=None,
    gerceklesen_teslim=None,
):
    """Admin workflow ve plan alanlarını kontrollü biçimde düzeltir."""
    kart_id = _sayi(kart_id, -1)

    with _kilit:
        kart = _kart_ref(kart_id)
        if not kart:
            raise KartBulunamadi("Kart bulunamadı.")

        yeni_durum = _durum_normalize(durum)
        if yeni_durum not in GECERLI_DURUMLAR:
            raise IsKuralHatasi("Durum PLANA ALINDI, DİZGİDE, HAZIR veya TESLİM EDİLDİ olmalı.")

        try:
            toplam = int(kart["toplam_adet"] if toplam_adet in (None, "") else toplam_adet)
            tamamlanan = int(
                kart["tamamlanan_adet"] if tamamlanan_adet in (None, "") else tamamlanan_adet
            )
        except (TypeError, ValueError) as exc:
            raise ValueError("Adetler sayı olmalı.") from exc

        if toplam < 1 or tamamlanan < 0 or tamamlanan > toplam:
            raise IsKuralHatasi("Adet değerleri tutarsız.")

        yeni_plan_hafta = (
            kart.get("plan_hafta")
            if plan_hafta is None
            else (_temiz_metin(plan_hafta) or None)
        )
        yeni_plan_baslama = (
            kart.get("plan_baslama")
            if plan_baslama is None
            else _tarih_form_degeri(plan_baslama, "Dizgi Başlama Tarihi")
        )
        yeni_plan_teslim = (
            kart.get("plan_teslim")
            if plan_teslim is None
            else _tarih_form_degeri(plan_teslim, "Planlanan Teslim Tarihi")
        )
        yeni_gerceklesen = (
            kart.get("gerceklesen_teslim")
            if gerceklesen_teslim is None
            else _tarih_form_degeri(gerceklesen_teslim, "Gerçekleşen Teslim Tarihi")
        )

        if yeni_plan_baslama and yeni_plan_teslim and yeni_plan_baslama > yeni_plan_teslim:
            raise IsKuralHatasi("Dizgi Başlama Tarihi Planlanan Teslim Tarihinden sonra olamaz.")

        onceki_durum = kart.get("durum") or "DURUMU EKSİK"

        def islem():
            nonlocal tamamlanan, yeni_gerceklesen

            baslama = kart.get("baslama_zamani")
            bitis = kart.get("bitis_zamani")
            teslim_zamani = kart.get("teslim_zamani")
            baslangic_adet = kart.get("baslangic_adet") or 0

            if yeni_durum == PLANA_ALINDI:
                tamamlanan = 0
                baslangic_adet = 0
                baslama = None
                bitis = None
                teslim_zamani = None
                yeni_gerceklesen = None

            elif yeni_durum == DIZGIDE:
                baslama = baslama or simdi()
                baslangic_adet = baslangic_adet or toplam
                bitis = bitis if tamamlanan == toplam else None
                teslim_zamani = None
                yeni_gerceklesen = None

            elif yeni_durum == HAZIR:
                tamamlanan = toplam
                baslama = baslama or simdi()
                bitis = bitis or simdi()
                teslim_zamani = None
                yeni_gerceklesen = None

            elif yeni_durum == TESLIM_EDILDI:
                tamamlanan = toplam
                baslama = baslama or simdi()
                bitis = bitis or simdi()
                yeni_gerceklesen = yeni_gerceklesen or bugun()
                teslim_zamani = teslim_zamani or simdi()

            kart.update(
                durum=yeni_durum,
                toplam_adet=toplam,
                tamamlanan_adet=tamamlanan,
                baslangic_adet=baslangic_adet,
                plan_hafta=yeni_plan_hafta,
                plan_baslama=yeni_plan_baslama,
                plan_teslim=yeni_plan_teslim,
                gerceklesen_teslim=yeni_gerceklesen,
                baslama_zamani=baslama,
                bitis_zamani=bitis,
                teslim_zamani=teslim_zamani,
                aciklama=_temiz_metin(aciklama) if aciklama is not None else kart.get("aciklama"),
                guncelleme=simdi(),
            )

            _loglar.append(
                _log_kaydi(
                    kullanici,
                    "admin",
                    "ADMİN DÜZENLEDİ",
                    kart.get("talep_no") or "",
                    kart.get("stok_no") or "",
                    tamamlanan,
                    f"{onceki_durum} → {yeni_durum} · {tamamlanan}/{toplam} adet",
                )
            )
            return kart_gorunumu(kart)

        return _atomik_kart_islemi(islem)


def admin_kart_gizle(kart_id, kullanici):
    kart_id = _sayi(kart_id, -1)

    with _kilit:
        kart = _kart_ref(kart_id)
        if not kart:
            raise KartBulunamadi("Kart bulunamadı.")

        def islem():
            kart["admin_gizli"] = 1
            kart["guncelleme"] = simdi()
            _loglar.append(
                _log_kaydi(
                    kullanici,
                    "admin",
                    "KART LİSTEDEN GİZLENDİ",
                    kart.get("talep_no") or "",
                    kart.get("stok_no") or "",
                    detay="Kart silinmedi; admin gizli olarak işaretlendi.",
                )
            )

        _atomik_kart_islemi(islem)


def admin_kart_geri_getir(kart_id, kullanici):
    kart_id = _sayi(kart_id, -1)

    with _kilit:
        kart = _kart_ref(kart_id)
        if not kart:
            raise KartBulunamadi("Kart bulunamadı.")
        if kart.get("admin_gizli", 0) != 1:
            raise IsKuralHatasi("Bu kart zaten görünür durumda.")

        def islem():
            kart["admin_gizli"] = 0
            kart["aktif"] = 1
            kart["guncelleme"] = simdi()
            _loglar.append(
                _log_kaydi(
                    kullanici,
                    "admin",
                    "GİZLENEN KART GERİ GETİRİLDİ",
                    kart.get("talep_no") or "",
                    kart.get("stok_no") or "",
                    detay="Kart tekrar yönetim/operasyon listelerine alındı.",
                )
            )
            return kart_gorunumu(kart)

        return _atomik_kart_islemi(islem)


# ---------------------------------------------------------------------------
# Yedekten geri yükleme
# ---------------------------------------------------------------------------

def yedekten_geri_yukle(yedek_adi, kullanici):
    global _kartlar, _loglar

    with _kilit:
        aday = _yedek_kart_dosyasi_bul(yedek_adi)
        yeni_kartlar = _oku(aday, KART_ALANLARI, ZORUNLU_KART_ALANLARI)
        yeni_kartlar = [_kart_normalize(kart) for kart in yeni_kartlar]
        if not yeni_kartlar:
            raise VeriDogrulamaHatasi("Seçilen yedekte hiç kart bulunmuyor.")
        _kart_listesi_dogrula(yeni_kartlar)

        koruma_yedegi = anlik_yedek("geri_yukleme_oncesi")
        eski_kartlar = copy.deepcopy(_kartlar)
        eski_loglar = copy.deepcopy(_loglar)

        try:
            _kartlar = yeni_kartlar
            _loglar.append(
                _log_kaydi(
                    kullanici,
                    "admin",
                    "YEDEKTEN GERİ YÜKLENDİ",
                    detay=(
                        f"Yedek: {yedek_adi} · {len(_kartlar)} kart · "
                        f"koruma: {os.path.basename(koruma_yedegi)}"
                    ),
                )
            )
            _coklu_yaz(
                [
                    (KARTLAR_DOSYA, KART_ALANLARI, _kartlar, "Kartlar"),
                    (LOG_DOSYA, LOG_ALANLARI, _loglar, "İşlem Logu"),
                ]
            )
        except Exception:
            _kartlar = eski_kartlar
            _loglar = eski_loglar
            raise

        return {
            "kart": len(_kartlar),
            "yedek": yedek_adi,
            "koruma_yedegi": koruma_yedegi,
        }


# ---------------------------------------------------------------------------
# Excel import commit
# ---------------------------------------------------------------------------

def excel_import_uygula(dosya_adi, kullanici, satirlar, uyari_sayisi=0):
    """Tamamen parse/validate edilmiş kaynak satırlarını tek transaction-benzeri blokta uygular."""
    global _kartlar, _loglar, _yuklemeler

    with _kilit:
        eski_kartlar = copy.deepcopy(_kartlar)
        eski_loglar = copy.deepcopy(_loglar)
        eski_yuklemeler = copy.deepcopy(_yuklemeler)

        try:
            yedek_klasoru = anlik_yedek("import_oncesi")
            mevcut_harita = {kart.get("anahtar"): kart for kart in _kartlar}
            gorulen = set()
            yeni = 0
            guncellenen = 0
            workflow_korundu = 0
            pasife_listesi = []
            sonraki_id = max([_sayi(kart.get("id"), 0) for kart in _kartlar] or [0]) + 1

            for satir in satirlar:
                anahtar = satir["anahtar"]
                gorulen.add(anahtar)
                plan = satir["plan"]
                mevcut = mevcut_harita.get(anahtar)

                if mevcut:
                    yeni_toplam = int(plan["toplam_adet"])
                    tamamlanan = int(mevcut.get("tamamlanan_adet") or 0)
                    if yeni_toplam < tamamlanan:
                        raise IsKuralHatasi(
                            f"{anahtar}: Excel toplam adedi ({yeni_toplam}), sistemde tamamlanan "
                            f"adetten ({tamamlanan}) küçük olamaz. Import iptal edildi."
                        )
                    if mevcut.get("durum") in (HAZIR, TESLIM_EDILDI) and yeni_toplam != tamamlanan:
                        raise IsKuralHatasi(
                            f"{anahtar}: Kart {mevcut['durum']} ve {tamamlanan} adet tamamlanmış. "
                            f"Yeni Excel toplam adedi {yeni_toplam}; admin kontrolü gerekir."
                        )

                    workflow = {
                        "durum": mevcut.get("durum"),
                        "baslangic_adet": mevcut.get("baslangic_adet"),
                        "tamamlanan_adet": tamamlanan,
                        "baslama_zamani": mevcut.get("baslama_zamani"),
                        "bitis_zamani": mevcut.get("bitis_zamani"),
                        "teslim_zamani": mevcut.get("teslim_zamani"),
                        "operator": mevcut.get("operator"),
                        "aciklama": mevcut.get("aciklama"),
                        "gerceklesen_teslim": mevcut.get("gerceklesen_teslim"),
                    }

                    mevcut.update(plan)
                    mevcut.update(workflow)
                    mevcut["source_active"] = 1
                    mevcut["aktif"] = 1
                    mevcut["kaynak"] = "EXCEL"
                    mevcut["guncelleme"] = simdi()
                    _kart_dogrula(mevcut)
                    guncellenen += 1
                    workflow_korundu += 1
                    continue

                toplam = int(plan["toplam_adet"])
                durum = satir.get("ilk_durum")
                gerceklesen = satir.get("gerceklesen_teslim")

                tamamlanan = toplam if durum in (HAZIR, TESLIM_EDILDI) else 0
                baslangic_adet = (
                    toplam
                    if durum in (DIZGIDE, HAZIR, TESLIM_EDILDI)
                    else 0
                )

                baslama = None
                bitis = None
                teslim_zamani = None

                kayit = {
                    "id": sonraki_id,
                    "anahtar": anahtar,
                    "talep_no": satir["talep_no"],
                    "stok_no": satir["stok_no"],
                    "gerceklesen_teslim": (
                        gerceklesen
                        if durum == TESLIM_EDILDI
                        else None
                    ),
                    "durum": durum,
                    "baslangic_adet": baslangic_adet,
                    "tamamlanan_adet": tamamlanan,
                    "baslama_zamani": None,
                    "bitis_zamani": None,
                    "teslim_zamani": None,
                    "operator": "Excel" if durum in GECERLI_DURUMLAR else None,
                    "aciklama": None,
                    "guncelleme": simdi(),
                    "aktif": 1,
                    "source_active": 1,
                    "admin_gizli": 0,
                    "kaynak": "EXCEL",
                }
                kayit.update(plan)
                _kart_dogrula(kayit)
                _kartlar.append(kayit)
                mevcut_harita[anahtar] = kayit
                sonraki_id += 1
                yeni += 1

            pasife_alinan = 0
            for kart in _kartlar:
                if kart.get("kaynak") != "EXCEL" or kart.get("anahtar") in gorulen:
                    continue
                if kart.get("source_active", 1) == 1:
                    kart["source_active"] = 0
                    kart["guncelleme"] = simdi()
                    pasife_alinan += 1
                    pasife_listesi.append(
                        f"{kart.get('talep_no') or ''}|{kart.get('stok_no') or ''}"
                    )

            _kart_listesi_dogrula(_kartlar)

            _yuklemeler.append(
                {
                    "zaman": simdi(),
                    "kullanici": kullanici,
                    "dosya": dosya_adi,
                    "satir": len(satirlar),
                    "yeni": yeni,
                    "guncellenen": guncellenen,
                    "pasife_alinan": pasife_alinan,
                    "uyari": uyari_sayisi,
                }
            )
            _loglar.append(
                _log_kaydi(
                    kullanici,
                    "admin",
                    "EXCEL YÜKLENDİ",
                    detay=(
                        f"{len(satirlar)} satır · {yeni} yeni · {guncellenen} güncellendi · "
                        f"{workflow_korundu} workflow korundu · {pasife_alinan} kaynakta yok · "
                        f"{uyari_sayisi} uyarı · yedek={os.path.basename(yedek_klasoru)}"
                    ),
                )
            )

            _gunluk_yedek(KARTLAR_DOSYA)
            _coklu_yaz(
                [
                    (KARTLAR_DOSYA, KART_ALANLARI, _kartlar, "Kartlar"),
                    (LOG_DOSYA, LOG_ALANLARI, _loglar, "İşlem Logu"),
                    (YUKLEME_DOSYA, YUKLEME_ALANLARI, _yuklemeler, "Yüklemeler"),
                ]
            )

            return {
                "satir": len(satirlar),
                "yeni": yeni,
                "guncellenen": guncellenen,
                "pasife_alinan": pasife_alinan,
                "pasife_listesi": pasife_listesi[:20],
                "workflow_korundu": workflow_korundu,
                "uyari": uyari_sayisi,
                "yedek": yedek_klasoru,
            }

        except Exception:
            _kartlar = eski_kartlar
            _loglar = eski_loglar
            _yuklemeler = eski_yuklemeler
            raise


# ---------------------------------------------------------------------------
# Legacy küçük yardımcılar
# ---------------------------------------------------------------------------

def kart_guncelle(kart_id, **alanlar):
    global _kartlar

    kart_id = _sayi(kart_id, -1)

    with _kilit:
        index = next(
            (
                i
                for i, kart in enumerate(_kartlar)
                if kart.get("id") == kart_id
            ),
            None,
        )

        if index is None:
            return None

        eski = copy.deepcopy(_kartlar)

        try:
            yeni = copy.deepcopy(_kartlar[index])
            yeni.update(alanlar)
            yeni["guncelleme"] = simdi()

            if "talep_no" in alanlar or "stok_no" in alanlar:
                yeni["anahtar"] = (
                    f"{_temiz_metin(yeni.get('talep_no'))}|"
                    f"{_temiz_metin(yeni.get('stok_no'))}"
                )

            yeni = _kart_normalize(yeni)
            _kartlar[index] = yeni

            _kart_listesi_dogrula(_kartlar)
            _kartlari_kaydet()

            return kart_gorunumu(yeni)

        except Exception:
            _kartlar = eski
            raise


def toplu_kaydet(yeni_kartlar=(), degisen=True):
    global _kartlar

    with _kilit:
        eski = copy.deepcopy(_kartlar)
        try:
            _kartlar.extend(_kart_normalize(kart) for kart in yeni_kartlar)
            _kart_listesi_dogrula(_kartlar)
            if degisen:
                _kartlari_kaydet()
        except Exception:
            _kartlar = eski
            raise
```


## `excel_araclari.py`


```python
"""PDGM kaynak Excel importu ve rapor üretimi.

Kaynak import kuralları:
- Yalnız MAKİNE sayfası kullanılır.
- Hidden kolonlar alınmaz; hidden satırlar korunur.
- Microsoft Excel COM ile values-only snapshot oluşturulur.
- External link güncellemesi ve full recalculation yapılmaz.
- Talep NO + Kart Stok No olmayan satırlar kart sayılmaz.
- Workflow yalnız şu dört durumdan oluşur:
    PLANA ALINDI, DİZGİDE, HAZIR, TESLİM EDİLDİ
- DURUM boş veya farklı bir değer ise kart kaybolmaz; durum None olur ve
  operasyon ekranında gösterilmez. Admin Yönetim ekranında düzeltebilir.
"""

from __future__ import annotations

import gc
import io
import os
import re
import tempfile
import threading
from datetime import datetime
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

try:
    import pythoncom
    import win32com.client
except ImportError:
    pythoncom = None
    win32com = None

import depo

_import_kilidi = threading.Lock()


class ExcelAktarimHatasi(Exception):
    pass


KAYNAK_SAYFA_ADI = "MAKİNE"
MSO_AUTOMATION_SECURITY_FORCE_DISABLE = 3


TURKCE_HARFLER = str.maketrans(
    {
        "ç": "c",
        "Ç": "C",
        "ğ": "g",
        "Ğ": "G",
        "ı": "i",
        "İ": "I",
        "ö": "o",
        "Ö": "O",
        "ş": "s",
        "Ş": "S",
        "ü": "u",
        "Ü": "U",
        "â": "a",
        "Â": "A",
    }
)


def _sadelestir(deger):
    metin = str(deger or "").translate(TURKCE_HARFLER).upper()
    metin = metin.replace(".", " ").replace("_", " ").replace("\xa0", " ")
    return re.sub(r"\s+", " ", metin).strip()


BASLIK_ESLESME = {
    _sadelestir(baslik): alan
    for baslik, alan in {
        "NO": "sira",
        "Talep NO": "talep_no",
        "Talep Sahibi": "talep_sahibi",
        "Kart Stok No": "stok_no",
        "Kart Üretim Adet": "adet_metin",
        "Planlanan Başlangıç T.": "plan_hafta",
        "Dizgi Başlama Tarihi": "plan_baslama",
        "Planlanan Teslim T.": "plan_teslim",
        "Gerçekleşen Teslim T.": "gerceklesen_teslim",
        "DURUM": "excel_durum",
        "PCB": "pcb",
        # Küçük format toleransı
        "Sıra": "sira",
        "Üretim Adet": "adet_metin",
        "Adet": "adet_metin",
    }.items()
}

DURUM_ESLESME = {
    "PLANA ALINDI": depo.PLANA_ALINDI,
    "DIZGIDE": depo.DIZGIDE,
    "HAZIR": depo.HAZIR,
    "TESLIM EDILDI": depo.TESLIM_EDILDI,
}


# ---------------------------------------------------------------------------
# COM snapshot
# ---------------------------------------------------------------------------

def excel_deger_snapshot_olustur(dosya_yolu):
    """MAKİNE sheet'indeki visible kolonları values-only geçici xlsx'e kopyalar."""
    if pythoncom is None or win32com is None:
        raise ExcelAktarimHatasi(
            "Excel COM aktarımı yalnızca Windows + Microsoft Excel ortamında çalışır "
            "(pywin32 gerekli)."
        )

    kaynak = Path(dosya_yolu).resolve()
    temp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
    temp.close()
    hedef = Path(temp.name).resolve()

    pythoncom.CoInitialize()
    excel = None
    kaynak_wb = None
    hedef_wb = None
    kaynak_ws = None
    hedef_ws = None
    used = None
    kaynak_aralik = None
    hedef_aralik = None
    basarili = False

    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.AskToUpdateLinks = False
        excel.EnableEvents = False
        excel.AutomationSecurity = MSO_AUTOMATION_SECURITY_FORCE_DISABLE

        kaynak_wb = excel.Workbooks.Open(
            str(kaynak),
            UpdateLinks=0,
            ReadOnly=True,
            IgnoreReadOnlyRecommended=True,
            AddToMru=False,
        )

        kaynak_ws = next(
            (
                ws
                for ws in kaynak_wb.Worksheets
                if _sadelestir(ws.Name) == _sadelestir(KAYNAK_SAYFA_ADI)
            ),
            None,
        )
        if kaynak_ws is None:
            sayfalar = [ws.Name for ws in kaynak_wb.Worksheets]
            raise ExcelAktarimHatasi(
                f"'{KAYNAK_SAYFA_ADI}' sayfası bulunamadı. "
                f"Dosyadaki sayfalar: {', '.join(sayfalar)}"
            )

        hedef_wb = excel.Workbooks.Add()
        while hedef_wb.Worksheets.Count > 1:
            hedef_wb.Worksheets(hedef_wb.Worksheets.Count).Delete()

        hedef_ws = hedef_wb.Worksheets(1)
        hedef_ws.Name = KAYNAK_SAYFA_ADI

        used = kaynak_ws.UsedRange
        ilk_satir = used.Row
        son_satir = used.Row + used.Rows.Count - 1
        ilk_sutun = used.Column
        son_sutun = used.Column + used.Columns.Count - 1
        hedef_sutun = 1

        for kaynak_sutun in range(ilk_sutun, son_sutun + 1):
            if bool(kaynak_ws.Columns(kaynak_sutun).Hidden):
                continue

            kaynak_aralik = kaynak_ws.Range(
                kaynak_ws.Cells(ilk_satir, kaynak_sutun),
                kaynak_ws.Cells(son_satir, kaynak_sutun),
            )
            hedef_aralik = hedef_ws.Range(
                hedef_ws.Cells(ilk_satir, hedef_sutun),
                hedef_ws.Cells(son_satir, hedef_sutun),
            )
            hedef_aralik.Value2 = kaynak_aralik.Value2
            hedef_sutun += 1

        hedef_wb.SaveAs(str(hedef), FileFormat=51)
        basarili = True
        return str(hedef)

    finally:
        if hedef_wb is not None:
            try:
                hedef_wb.Close(SaveChanges=False)
            except Exception:
                pass
        if kaynak_wb is not None:
            try:
                kaynak_wb.Close(SaveChanges=False)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass

        hedef_aralik = None
        kaynak_aralik = None
        used = None
        hedef_ws = None
        kaynak_ws = None
        hedef_wb = None
        kaynak_wb = None
        excel = None
        gc.collect()

        pythoncom.CoUninitialize()

        if not basarili:
            try:
                os.remove(str(hedef))
            except OSError:
                pass


# ---------------------------------------------------------------------------
# Parser helpers
# ---------------------------------------------------------------------------

def durum_coz(deger):
    """Kaynak durumunu dört gerçek workflow durumundan birine çevirir."""
    temiz = _sadelestir(deger)
    if not temiz:
        return None, True
    durum = DURUM_ESLESME.get(temiz)
    return durum, durum is None


def adet_coz(deger):
    """Üretim adedini pozitif tam sayıya çevirir.

    Kabul:
    - 400
    - 400.0
    - "400 ADET"
    - "1.500 ADET"
    - "1,500 ADET"
    - "1 500 ADET"

    Red:
    - 400.5
    - "400.5 ADET"
    - "400,5 ADET"
    - "-5 ADET"
    - "400 / 500 ADET"
    """
    if deger is None or str(deger).strip() == "":
        raise ExcelAktarimHatasi("Üretim adedi boş olamaz.")

    if isinstance(deger, bool):
        raise ExcelAktarimHatasi("Üretim adedi sayı olmalı.")

    if isinstance(deger, (int, float)):
        if isinstance(deger, float) and not deger.is_integer():
            raise ExcelAktarimHatasi(
                f"Üretim adedi tam sayı olmalı: {deger}"
            )
        adet = int(deger)

    else:
        metin = str(deger).strip()

        eslesmeler = [
            eslesme.strip()
            for eslesme in re.findall(
                r"[-+]?\d[\d.,\s]*",
                metin,
            )
            if eslesme.strip()
        ]

        if len(eslesmeler) != 1:
            raise ExcelAktarimHatasi(
                f"Üretim adedi tek bir sayı içermeli: '{metin}'"
            )

        token = re.sub(r"\s+", "", eslesmeler[0])

        if token.startswith(("+", "-")):
            raise ExcelAktarimHatasi(
                f"Üretim adedi pozitif tam sayı olmalı: '{metin}'"
            )

        if token.isdigit():
            adet = int(token)

        elif re.fullmatch(
            r"\d{1,3}(?:[.,]\d{3})+",
            token,
        ):
            adet = int(re.sub(r"[.,]", "", token))

        else:
            raise ExcelAktarimHatasi(
                f"Üretim adedi tam sayı olmalı: '{metin}'"
            )

    if adet < 1:
        raise ExcelAktarimHatasi(
            f"Üretim adedi en az 1 olmalı: {deger!r}"
        )

    return adet


def _baslik_satiri_bul(ws):
    for satir_no, satir in enumerate(
        ws.iter_rows(min_row=1, max_row=10, values_only=True),
        start=1,
    ):
        temiz = [_sadelestir(hucre) for hucre in satir]
        if "TALEP NO" in temiz or "KART STOK NO" in temiz:
            return satir_no, temiz
    return None, None


def _hucre_al(satir, kolonlar, alan):
    index = kolonlar.get(alan)
    if index is None or index >= len(satir):
        return None
    deger = satir[index]
    return None if deger == "" else deger


def _tarih_coz_ve_dogrula(deger, alan_adi, excel_satir_no):
    sonuc = depo.tarih_coz(deger)
    if deger not in (None, "") and not sonuc:
        raise ExcelAktarimHatasi(
            f"Excel satır {excel_satir_no}: {alan_adi} okunamadı. Gelen değer: {deger!r}"
        )
    return sonuc


def _sira_coz(deger):
    if deger in (None, ""):
        return None, False
    try:
        return int(float(deger)), False
    except (TypeError, ValueError):
        return None, True


# ---------------------------------------------------------------------------
# Import
# ---------------------------------------------------------------------------

def excelden_aktar(dosya_yolu, kullanici):
    with _import_kilidi:
        return _excelden_aktar(dosya_yolu, kullanici)


def _excelden_aktar(dosya_yolu, kullanici):
    snapshot_yolu = None
    wb = None

    try:
        snapshot_yolu = excel_deger_snapshot_olustur(dosya_yolu)
        wb = openpyxl.load_workbook(snapshot_yolu, data_only=True, read_only=True)
    except ExcelAktarimHatasi:
        raise
    except Exception as exc:
        raise ExcelAktarimHatasi(
            f"Dosya geçerli bir Excel çalışma kitabı değil: {exc}"
        ) from exc

    try:
        if KAYNAK_SAYFA_ADI not in wb.sheetnames:
            raise ExcelAktarimHatasi(
                f"'{KAYNAK_SAYFA_ADI}' sayfası bulunamadı. "
                f"Dosyadaki sayfalar: {', '.join(wb.sheetnames)}"
            )

        ws = wb[KAYNAK_SAYFA_ADI]
        baslik_no, basliklar = _baslik_satiri_bul(ws)
        if not baslik_no:
            raise ExcelAktarimHatasi(
                "Başlık satırı bulunamadı. Dosyada 'Talep NO' veya 'Kart Stok No' sütunu olmalı."
            )

        kolonlar = {}
        for index, baslik in enumerate(basliklar):
            alan = BASLIK_ESLESME.get(baslik)
            if not alan:
                continue
            if alan in kolonlar:
                raise ExcelAktarimHatasi(
                    f"'{baslik}' için birden fazla görünür Excel sütunu bulundu."
                )
            kolonlar[alan] = index

        for gerekli in ("talep_no", "stok_no", "adet_metin"):
            if gerekli not in kolonlar:
                ad = {
                    "talep_no": "Talep NO",
                    "stok_no": "Kart Stok No",
                    "adet_metin": "Kart Üretim Adet",
                }[gerekli]
                raise ExcelAktarimHatasi(f"Zorunlu '{ad}' sütunu bulunamadı.")

        parsed = []
        gorulen_anahtarlar = set()
        uyari_sayisi = 0

        for excel_satir_no, satir in enumerate(
            ws.iter_rows(min_row=baslik_no + 1, values_only=True),
            start=baslik_no + 1,
        ):
            if not any(hucre not in (None, "") for hucre in satir):
                continue

            talep_no = str(_hucre_al(satir, kolonlar, "talep_no") or "").strip()
            stok_no = str(_hucre_al(satir, kolonlar, "stok_no") or "").strip()
            if not talep_no or not stok_no:
                continue

            anahtar = f"{talep_no}|{stok_no}"
            if anahtar in gorulen_anahtarlar:
                raise ExcelAktarimHatasi(
                    f"Excel satır {excel_satir_no}: Talep NO + Kart Stok No tekrar ediyor ({anahtar})."
                )
            gorulen_anahtarlar.add(anahtar)

            try:
                toplam_adet = adet_coz(_hucre_al(satir, kolonlar, "adet_metin"))
            except ExcelAktarimHatasi as exc:
                raise ExcelAktarimHatasi(f"Excel satır {excel_satir_no}: {exc}") from exc

            plan_baslama_ham = _hucre_al(satir, kolonlar, "plan_baslama")
            plan_teslim_ham = _hucre_al(satir, kolonlar, "plan_teslim")
            gerceklesen_ham = _hucre_al(satir, kolonlar, "gerceklesen_teslim")

            plan_baslama = _tarih_coz_ve_dogrula(
                plan_baslama_ham, "Dizgi Başlama Tarihi", excel_satir_no
            )
            plan_teslim = _tarih_coz_ve_dogrula(
                plan_teslim_ham, "Planlanan Teslim Tarihi", excel_satir_no
            )
            gerceklesen = _tarih_coz_ve_dogrula(
                gerceklesen_ham, "Gerçekleşen Teslim Tarihi", excel_satir_no
            )

            if plan_baslama and plan_teslim and plan_baslama > plan_teslim:
                raise ExcelAktarimHatasi(
                    f"Excel satır {excel_satir_no}: Dizgi Başlama Tarihi ({plan_baslama}) "
                    f"Planlanan Teslim Tarihinden ({plan_teslim}) sonra olamaz."
                )

            excel_durum_raw = _hucre_al(satir, kolonlar, "excel_durum")
            ilk_durum, durum_uyarisi = durum_coz(excel_durum_raw)
            if durum_uyarisi:
                uyari_sayisi += 1

            if (ilk_durum == depo.TESLIM_EDILDI and not gerceklesen):
                ilk_durum = None
                uyari_sayisi +=1

            elif (ilk_durum != depo.TESLIM_EDILDI and gerceklesen):
                uyari_sayisi += 1

            sira, sira_uyarisi = _sira_coz(_hucre_al(satir, kolonlar, "sira"))
            if sira_uyarisi:
                uyari_sayisi += 1

            plan = {
                "sira": sira,
                "talep_sahibi": str(_hucre_al(satir, kolonlar, "talep_sahibi") or "").strip(),
                "toplam_adet": toplam_adet,
                "adet_metin": str(_hucre_al(satir, kolonlar, "adet_metin") or "").strip(),
                "plan_hafta": str(_hucre_al(satir, kolonlar, "plan_hafta") or "").strip(),
                "plan_baslama": plan_baslama,
                "plan_teslim": plan_teslim,
                "excel_durum": str(excel_durum_raw or "").strip(),
                "pcb": str(_hucre_al(satir, kolonlar, "pcb") or "").strip(),
            }

            parsed.append(
                {
                    "anahtar": anahtar,
                    "talep_no": talep_no,
                    "stok_no": stok_no,
                    "plan": plan,
                    "gerceklesen_teslim": gerceklesen,
                    "ilk_durum": ilk_durum,
                }
            )

        if not parsed:
            raise ExcelAktarimHatasi("Excel'de işlenecek kart satırı bulunamadı.")

    finally:
        if wb is not None:
            try:
                wb.close()
            except Exception:
                pass
        if snapshot_yolu:
            try:
                os.remove(snapshot_yolu)
            except OSError:
                pass

    return depo.excel_import_uygula(
        dosya_adi=os.path.basename(dosya_yolu),
        kullanici=kullanici,
        satirlar=parsed,
        uyari_sayisi=uyari_sayisi,
    )


# ---------------------------------------------------------------------------
# Rapor üretimi
# ---------------------------------------------------------------------------

BASLIK_DOLGU = PatternFill("solid", fgColor="0F2027")
BASLIK_YAZI = Font(name="Arial", bold=True, color="FFFFFF", size=11)
GOVDE_YAZI = Font(name="Arial", size=10)

def _excel_hucre_yaz(hucre, deger):
    hucre.value = deger

    if isinstance(deger, str) and deger.startswith("="):
        hucre.data_type = "s"

def _sayfa_yaz(ws, basliklar, satirlar):
    ws.append(basliklar)
    for hucre in ws[1]:
        hucre.fill = BASLIK_DOLGU
        hucre.font = BASLIK_YAZI
        hucre.alignment = Alignment(horizontal="center", vertical="center")

    for satir in satirlar:
        satir_no = ws.max_row + 1

        for sutun_no, deger in enumerate(satir,start=1):
            _excel_hucre_yaz(
                ws.cell(row=satir_no,column=sutun_no,),deger)


    for sutun in range(1, len(basliklar) + 1):
        en = max(
            [len(str(basliklar[sutun - 1])), 10]
            + [len(str(satir[sutun - 1])) for satir in satirlar[:400]]
        )
        ws.column_dimensions[get_column_letter(sutun)].width = min(40, en + 3)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    for satir in ws.iter_rows(min_row=2):
        for hucre in satir:
            hucre.font = GOVDE_YAZI


def calisma_kitabi_uret(kartlar, loglar, ozet):
    wb = openpyxl.Workbook()

    ws = wb.active
    ws.title = "Kart Durumları"
    _sayfa_yaz(
        ws,
        [
            "ID",
            "Sıra",
            "Talep NO",
            "Talep Sahibi",
            "Kart Stok No",
            "Toplam Adet",
            "Tamamlanan",
            "Kalan",
            "Durum",
            "Kaynak Durumu",
            "Değerlendirme",
            "Sapma (gün)",
            "Plan Başlangıç",
            "Plan Teslim",
            "Üretim Başlangıç",
            "Üretim Bitiş",
            "Teslim Zamanı",
            "Gerçekleşen Teslim",
            "Operatör",
            "Not",
            "PCB",
            "Kaynakta Aktif",
            "Admin Gizli",
        ],
        [
            [
                k.get("id"),
                k.get("sira"),
                k.get("talep_no"),
                k.get("talep_sahibi"),
                k.get("stok_no"),
                k.get("toplam_adet"),
                k.get("tamamlanan_adet"),
                k.get("kalan_adet"),
                k.get("durum") or "DURUMU EKSİK",
                k.get("excel_durum") or "",
                k.get("rozet"),
                k.get("sapma") if k.get("sapma") is not None else "",
                k.get("plan_baslama") or "",
                k.get("plan_teslim") or "",
                k.get("baslama_zamani") or "",
                k.get("bitis_zamani") or "",
                k.get("teslim_zamani") or "",
                k.get("gerceklesen_teslim") or "",
                k.get("operator") or "",
                k.get("aciklama") or "",
                k.get("pcb") or "",
                k.get("source_active", 1),
                k.get("admin_gizli", 0),
            ]
            for k in kartlar
        ],
    )

    ws2 = wb.create_sheet("İşlem Logu")
    _sayfa_yaz(
        ws2,
        ["Zaman", "Kullanıcı", "Rol", "İşlem", "Talep NO", "Kart Stok No", "Adet", "Detay"],
        [
            [
                l.get("zaman"),
                l.get("kullanici"),
                l.get("rol"),
                l.get("islem"),
                l.get("talep_no") or "",
                l.get("stok_no") or "",
                l.get("adet") if l.get("adet") is not None else "",
                l.get("detay") or "",
            ]
            for l in loglar
        ],
    )

    ws3 = wb.create_sheet("Özet")
    _sayfa_yaz(ws3, ["Başlık", "Değer"], ozet)
    return wb


def kitap_baytlari(wb):
    tampon = io.BytesIO()
    wb.save(tampon)
    tampon.seek(0)
    return tampon


def dosya_adi(on_ek):
    return f"{on_ek}_{datetime.now():%Y%m%d_%H%M}.xlsx"
```


## `kullanici_yonet.py`


```python
"""PDGM kullanıcı yönetimi için küçük CLI aracı.

Parolalar plaintext saklanmaz; yalnız Werkzeug hash'i kullanicilar.json'a yazılır.
"""

from __future__ import annotations

import getpass
import json
import os
import sys

from werkzeug.security import generate_password_hash

KOK = os.path.dirname(os.path.abspath(__file__))
DOSYA = os.path.join(KOK, "data", "kullanicilar.json")
ROLLER = {"admin", "operator", "gozlemci"}


def oku():
    if not os.path.exists(DOSYA):
        raise SystemExit("Önce uygulamayı bir kez çalıştırın; data/kullanicilar.json oluşturulsun.")
    with open(DOSYA, encoding="utf-8") as f:
        veri = json.load(f)
    if not isinstance(veri, dict):
        raise SystemExit("kullanicilar.json geçersiz.")
    return veri


def yaz(veri):
    gecici = DOSYA + ".yeni"
    with open(gecici, "w", encoding="utf-8") as f:
        json.dump(veri, f, ensure_ascii=False, indent=2)
        f.flush()
        os.fsync(f.fileno())
    os.replace(gecici, DOSYA)


def _parola_sor():
    parola = getpass.getpass("Yeni parola: ")
    parola_tekrar = getpass.getpass("Yeni parola tekrar: ")
    if parola != parola_tekrar:
        raise SystemExit("Parolalar eşleşmiyor.")
    if len(parola) < 8:
        raise SystemExit("Parola en az 8 karakter olmalı.")
    return parola


def _aktif_admin_sayisi(veri, haric=None):
    return sum(
        1
        for k, b in veri.items()
        if k != haric and b.get("rol") == "admin" and b.get("aktif", True)
    )


def ekle(kullanici, rol, gorunen_ad):
    kullanici = kullanici.strip()
    rol = rol.strip().lower()
    gorunen_ad = gorunen_ad.strip()

    if not kullanici:
        raise SystemExit("Kullanıcı adı boş olamaz.")
    if rol not in ROLLER:
        raise SystemExit(f"Rol şu değerlerden biri olmalı: {', '.join(sorted(ROLLER))}")

    veri = oku()
    if kullanici in veri:
        raise SystemExit("Bu kullanıcı adı zaten var.")

    parola = getpass.getpass("Parola: ")
    parola_tekrar = getpass.getpass("Parola tekrar: ")
    if parola != parola_tekrar:
        raise SystemExit("Parolalar eşleşmiyor.")
    if len(parola) < 8:
        raise SystemExit("Parola en az 8 karakter olmalı.")

    veri[kullanici] = {
        "sifre_hash": generate_password_hash(parola),
        "rol": rol,
        "ad": gorunen_ad or kullanici,
        "aktif": True,
    }
    yaz(veri)
    print(f"Kullanıcı eklendi: {kullanici} ({rol})")


def parola_degistir(kullanici):
    kullanici = kullanici.strip()
    veri = oku()
    if kullanici not in veri:
        raise SystemExit("Kullanıcı bulunamadı.")

    veri[kullanici]["sifre_hash"] = generate_password_hash(_parola_sor())
    yaz(veri)
    print(f"{kullanici}: parola güncellendi.")
    print("Sunucuyu yeniden başlatmanıza gerek yok.")


def rol_degistir(kullanici, yeni_rol):
    kullanici = kullanici.strip()
    yeni_rol = yeni_rol.strip().lower()

    if yeni_rol not in ROLLER:
        raise SystemExit(f"Rol şu değerlerden biri olmalı: {', '.join(sorted(ROLLER))}")

    veri = oku()
    if kullanici not in veri:
        raise SystemExit("Kullanıcı bulunamadı.")

    eski_rol = veri[kullanici].get("rol", "-")

    if eski_rol == "admin" and yeni_rol != "admin":
        if _aktif_admin_sayisi(veri, haric=kullanici) == 0:
            raise SystemExit(
                "Bu son aktif admin. Rolü düşürürseniz sistemi yönetemezsiniz. "
                "Önce başka bir admin oluşturun."
            )

    veri[kullanici]["rol"] = yeni_rol
    yaz(veri)
    print(f"{kullanici}: {eski_rol} -> {yeni_rol}")
    print("Değişiklik ilk request'te etkili olur; yeniden başlatma gerekmez.")


def aktiflik(kullanici, aktif):
    veri = oku()
    if kullanici not in veri:
        raise SystemExit("Kullanıcı bulunamadı.")

    if not aktif and veri[kullanici].get("rol") == "admin":
        if _aktif_admin_sayisi(veri, haric=kullanici) == 0:
            raise SystemExit(
                "Bu son aktif admin. Pasife alırsanız sisteme admin olarak "
                "giremezsiniz. Önce başka bir admin oluşturun."
            )

    veri[kullanici]["aktif"] = aktif
    yaz(veri)
    print(f"{kullanici}: {'aktif' if aktif else 'pasif'}")


def listele():
    for kullanici, bilgi in oku().items():
        print(
            f"{kullanici:20} "
            f"{bilgi.get('rol', '-'):10} "
            f"{bilgi.get('ad', '-'):30} "
            f"aktif={bilgi.get('aktif', True)}"
        )


def main():
    if len(sys.argv) < 2:
        raise SystemExit(
            "Kullanım:\n"
            "  python kullanici_yonet.py listele\n"
            "  python kullanici_yonet.py ekle <kullanici> <admin|operator|gozlemci> <Görünen Ad>\n"
            "  python kullanici_yonet.py parola <kullanici>\n"
            "  python kullanici_yonet.py rol <kullanici> <admin|operator|gozlemci>\n"
            "  python kullanici_yonet.py pasif <kullanici>\n"
            "  python kullanici_yonet.py aktif <kullanici>"
        )

    komut = sys.argv[1].lower()
    if komut == "listele":
        listele()
    elif komut == "ekle" and len(sys.argv) >= 5:
        ekle(sys.argv[2], sys.argv[3], " ".join(sys.argv[4:]))
    elif komut == "parola" and len(sys.argv) == 3:
        parola_degistir(sys.argv[2])
    elif komut == "rol" and len(sys.argv) == 4:
        rol_degistir(sys.argv[2], sys.argv[3])
    elif komut == "pasif" and len(sys.argv) == 3:
        aktiflik(sys.argv[2], False)
    elif komut == "aktif" and len(sys.argv) == 3:
        aktiflik(sys.argv[2], True)
    else:
        raise SystemExit("Geçersiz komut veya eksik argüman.")


if __name__ == "__main__":
    main()
```


# 2. TEMPLATES


## `templates/base.html`


```html
<!doctype html>
<html lang="tr">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <meta name="csrf-token" content="{{ csrf_token }}">
    <title>{% block title %}PDGM İş Takip{% endblock %}</title>
    <link rel="stylesheet" href="{{ url_for('static', filename='stil.css') }}">
    {% block head %}{% endblock %}
</head>
<body class="{% block body_class %}{% endblock %}">
<header class="ust-cubuk">
    <a href="{{ url_for('ana') }}" class="marka-link">
        <span class="marka-isaret">
            <img src="{{ url_for('static', filename='pdgm_logo.png') }}" alt="PDGM" class="marka-logo">
        </span>
        <span class="marka-metin">
            <strong>İş Takip Sistemi</strong>
            <small>Baskı Dizgi Atölyesi</small>
        </span>
    </a>

    <nav class="ana-nav" aria-label="Ana menü">
        {% if oturum_rol in ["admin", "operator", "gozlemci"] %}
            <a class="{% if request.endpoint == 'panel' %}aktif{% endif %}" href="{{ url_for('panel') }}">Pano</a>
            <a class="{% if request.endpoint == 'ozet' %}aktif{% endif %}" href="{{ url_for('ozet') }}">Özet</a>
        {% endif %}

        {% if oturum_rol in ["admin", "operator"] %}
            <a class="{% if request.endpoint == 'operator' %}aktif{% endif %}" href="{{ url_for('operator') }}">Operatör</a>
        {% endif %}

        {% if oturum_rol in ["admin", "operator", "gozlemci"] %}
            <a class="{% if request.endpoint == 'monitor' %}aktif{% endif %}" href="{{ url_for('monitor') }}">Monitör</a>
        {% endif %}

        {% if oturum_rol == "admin" %}
            <a class="{% if request.endpoint == 'yonetim' %}aktif{% endif %}" href="{{ url_for('yonetim') }}">Yönetim</a>
        {% endif %}
    </nav>

    <div class="oturum">
        <div class="oturum-metin">
            <strong>{{ oturum_ad or oturum_kullanici }}</strong>
            <small>{{ oturum_rol|upper }}</small>
        </div>
        {% if oturum_kullanici %}
        <form method="post" action="{{ url_for('cikis') }}">
            <input type="hidden" name="_csrf_token" value="{{ csrf_token }}">
            <button class="buton buton-hayalet buton-kucuk" type="submit">Çıkış</button>
        </form>
        {% endif %}
    </div>
</header>

<main class="sayfa">
    {% with mesajlar = get_flashed_messages(with_categories=true) %}
        {% if mesajlar %}
        <section class="bildirimler" aria-live="polite">
            {% for kategori, mesaj in mesajlar %}
                <div class="bildirim {{ kategori }}">{{ mesaj }}</div>
            {% endfor %}
        </section>
        {% endif %}
    {% endwith %}

    {% block content %}{% endblock %}
</main>

<div id="toast-alani" class="toast-alani" aria-live="polite"></div>

<script>
function csrfToken() {
    return document.querySelector('meta[name="csrf-token"]')?.content || "";
}

async function pdgmFetch(url, options = {}) {
    const headers = new Headers(options.headers || {});
    headers.set("X-CSRF-Token", csrfToken());

    if (options.body && !(options.body instanceof FormData) && !headers.has("Content-Type")) {
        headers.set("Content-Type", "application/json");
    }

    const response = await fetch(url, {credentials: "same-origin", ...options, headers});
    const contentType = response.headers.get("content-type") || "";
    const data = contentType.includes("application/json")
        ? await response.json()
        : {hata: await response.text()};

    if (!response.ok) {
        throw new Error(data.hata || data.detail || "İşlem tamamlanamadı.");
    }
    return data;
}

function toast(mesaj, tip = "basari") {
    const alan = document.getElementById("toast-alani");
    if (!alan) return;

    const kutu = document.createElement("div");
    kutu.className = `toast ${tip}`;
    kutu.textContent = mesaj;
    alan.appendChild(kutu);

    requestAnimationFrame(() => kutu.classList.add("goster"));
    window.setTimeout(() => {
        kutu.classList.remove("goster");
        window.setTimeout(() => kutu.remove(), 200);
    }, 3200);
}

function hataMesaji(hata) {
    toast(hata?.message || "Beklenmeyen bir hata oluştu.", "hata");
}

function dialogKapat(dialog) {
    if (dialog?.open) dialog.close();
}

document.addEventListener("click", (event) => {
    const kapat = event.target.closest("[data-dialog-kapat]");
    if (kapat) dialogKapat(kapat.closest("dialog"));
});
</script>
{% block scripts %}{% endblock %}
</body>
</html>
```


## `templates/giris.html`


```html
<!doctype html>
<html lang="tr">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>Giriş · PDGM İş Takip</title>
    <link rel="stylesheet" href="{{ url_for('static', filename='stil.css') }}">
</head>
<body class="giris-sayfa">
<main class="giris-kutu">
    <div class="giris-logo">
        <img src="{{ url_for('static', filename='pdgm_logo.png') }}" alt="PDGM Logo">
    </div>

    <h1>PDGM İŞ TAKİP</h1>
    <p class="alt">Baskı Dizgi Atölyesi</p>

    {% if hata %}<div class="bildirim hata">{{ hata }}</div>{% endif %}

    <form method="post" autocomplete="off" class="giris-form">
        <div class="alan">
            <label for="kullanici">Kullanıcı adı</label>
            <input type="text" id="kullanici" name="kullanici" autocomplete="username" autofocus required>
        </div>
        <div class="alan">
            <label for="sifre">Şifre</label>
            <input type="password" id="sifre" name="sifre" autocomplete="current-password" required>
        </div>
        <button class="buton buton-ana tam-genislik" type="submit">Giriş Yap</button>
    </form>

    <p class="giris-dipnot">Bu sistem yalnızca yetkili PDGM personeli içindir.</p>
</main>
</body>
</html>
```


## `templates/monitor.html`


```html
{% extends "base.html" %}

{% block title %}Monitör · PDGM İş Takip{% endblock %}
{% block body_class %}monitor-body{% endblock %}

{% block content %}
<div class="monitor-sayfa">

    <section class="monitor-ust">
        <div>
            <p class="ust-etiket">ATÖLYE CANLI GÖRÜNÜMÜ</p>
            <h1>Üretim Monitörü</h1>
            <p class="monitor-aciklama">
                Aktif üretim ve üretim sırası.
            </p>
        </div>

        <div class="monitor-zaman">
            <span id="monitor-saat">{{ guncelleme }}</span>
            <small>
                {{ bugun }} · Son veri yenileme {{ guncelleme }}
            </small>
        </div>
    </section>

    <section class="monitor-grid">

        <!-- ============================================================
             DİZGİDE
        ============================================================ -->

        <article class="panel-kutu monitor-bolum monitor-dizgide">

            <div class="panel-baslik monitor-bolum-baslik">
                <div>
                    <p class="ust-etiket">AKTİF ÜRETİM</p>
                    <h2>DİZGİDE</h2>
                </div>

                <span class="sayi-rozet">
                    {{ dizgide|length }}
                </span>
            </div>

            {% if dizgide %}

            <div class="monitor-liste monitor-dizgide-liste {% if dizgide|length >= 3 %}uc-dizgi{% endif %}"
                data-monitor-grup="dizgide"
                data-sayfa-boyutu="3">

                {% for k in dizgide %}

                <article class="monitor-kart {{ k.renk }}" data-monitor-kart>

                    <div class="monitor-kart-ust">

                        <div class="monitor-talep">
                            <span>TALEP</span>
                            <strong>
                                {{ k.talep_no or "—" }}
                            </strong>
                        </div>

                        <span class="durum-rozet {{ k.renk }}">
                            {{ k.rozet }}
                        </span>

                    </div>

                    <h3>
                        {{ k.stok_no or "Stok no yok" }}
                    </h3>

                    <p class="monitor-sahip">
                        {{ k.talep_sahibi or "Talep sahibi belirtilmemiş" }}
                    </p>

                    <div class="monitor-adet">
                        <strong>
                            {{ k.tamamlanan_adet }}
                            /
                            {{ k.toplam_adet }}
                        </strong>

                        <span>
                            adet tamamlandı
                        </span>
                    </div>

                    <div class="ilerleme monitor-ilerleme">

                        <div class="ilerleme-ust">
                            <span>
                                Üretim ilerlemesi
                            </span>

                            <strong>
                                %{{ k.adet_yuzde }}
                            </strong>
                        </div>

                        <div class="ilerleme-ray">
                            <div
                                class="ilerleme-dolgu"
                                style="width: {{ [k.adet_yuzde, 100]|min }}%">
                            </div>
                        </div>

                    </div>

                    <div class="monitor-meta">

                        <span>
                            <small>Plan teslim</small>
                            <strong>
                                {{ k.plan_teslim|gun }}
                            </strong>
                        </span>

                    </div>

                </article>

                {% endfor %}

            </div>

            {% else %}

            <div class="bos-durum monitor-bos">
                Şu anda dizgide kart bulunmuyor.
            </div>

            {% endif %}

        </article>


        <!-- ============================================================
             PLANA ALINDI
        ============================================================ -->

        <article class="panel-kutu monitor-bolum monitor-plana">

            <div class="panel-baslik monitor-bolum-baslik">

                <div>
                    <p class="ust-etiket">ÜRETİM KUYRUĞU</p>
                    <h2>PLANA ALINDI</h2>
                </div>

                <span class="sayi-rozet">
                    {{ plana_alindi|length }}
                </span>

            </div>

            {% if plana_alindi %}

            <div class="monitor-liste monitor-plan-liste"
                data-monitor-grup="plan"
                data-sayfa-boyutu="6">

                {% for k in plana_alindi %}

                <article class="monitor-plan-kart {{ k.renk }}" data-monitor-kart>

                    <div class="monitor-plan-sira">
                        {{ k.sira or loop.index }}
                    </div>

                    <div class="monitor-plan-icerik">

                        <div class="monitor-plan-ust">

                            <div>
                                <span>TALEP</span>
                                <strong>
                                    {{ k.talep_no or "—" }}
                                </strong>
                            </div>

                            <span class="durum-rozet {{ k.renk }}">
                                {{ k.rozet }}
                            </span>

                        </div>

                        <h3>
                            {{ k.stok_no or "Stok no yok" }}
                        </h3>

                        <p>
                            {{ k.talep_sahibi or "Talep sahibi belirtilmemiş" }}
                        </p>

                        <div class="monitor-plan-meta">

                            <span>
                                <small>Adet</small>
                                <strong>
                                    {{ k.toplam_adet }}
                                </strong>
                            </span>

                            <span>
                                <small>Başlangıç</small>
                                <strong>
                                    {{ k.plan_baslama|gun }}
                                </strong>
                            </span>

                            <span>
                                <small>Teslim</small>
                                <strong>
                                    {{ k.plan_teslim|gun }}
                                </strong>
                            </span>

                        </div>

                    </div>

                </article>

                {% endfor %}

            </div>

            {% else %}

            <div class="bos-durum monitor-bos">
                Plana alınmış bekleyen kart yok.
            </div>

            {% endif %}

        </article>

    </section>

</div>
{% endblock %}


{% block scripts %}
<script>
(() => {
    const saat = document.getElementById("monitor-saat");

    function saatiGuncelle() {
        if (!saat) return;

        saat.textContent = new Intl.DateTimeFormat("tr-TR", {
            hour: "2-digit",
            minute: "2-digit",
            second: "2-digit"
        }).format(new Date());
    }

    function rotasyonBaslat(grup) {
        const kartlar = [...grup.querySelectorAll("[data-monitor-kart]")];
        const sayfaBoyutu = Number(grup.dataset.sayfaBoyutu || 0);

        if (!sayfaBoyutu || kartlar.length <= sayfaBoyutu) {
            return;
        }

        const sayfaSayisi = Math.ceil(
            kartlar.length / sayfaBoyutu
        );

        let aktifSayfa = 0;

        function sayfayiGoster() {
            const baslangic = aktifSayfa * sayfaBoyutu;
            const bitis = baslangic + sayfaBoyutu;

            kartlar.forEach((kart, index) => {
                kart.hidden = !(
                    index >= baslangic &&
                    index < bitis
                );
            });
        }

        sayfayiGoster();

        window.setInterval(() => {
            aktifSayfa = (aktifSayfa + 1) % sayfaSayisi;
            sayfayiGoster();
        }, 5000);
    }

    saatiGuncelle();

    window.setInterval(
        saatiGuncelle,
        1000
    );

    document
        .querySelectorAll("[data-monitor-grup]")
        .forEach(rotasyonBaslat);

    window.setTimeout(
        () => window.location.reload(),
        30000
    );
})();
</script>
{% endblock %}
```


## `templates/operator.html`


```html
{% extends "base.html" %}
{% block title %}Operatör · PDGM İş Takip{% endblock %}

{% block content %}
<div class="sayfa-shell operator-sayfa">
    <section class="sayfa-baslik sayfa-hero">
        <div>
            <p class="ust-etiket">OPERATÖR EKRANI</p>
            <h1>Kart İşlemleri</h1>
            <p class="soluk">Üretimi başlatın, tamamlanan adedi kaydedin, HAZIR ve TESLİM EDİLDİ geçişlerini yönetin.</p>
        </div>
    </section>

    <section class="istatistik-grid" aria-label="Durum özeti">
        <article class="istatistik kart-plana"><span>Plana Alındı</span><strong>{{ sayac.plana_alindi }}</strong></article>
        <article class="istatistik kart-dizgide"><span>Dizgide</span><strong>{{ sayac.dizgide }}</strong></article>
        <article class="istatistik kart-hazir"><span>Hazır</span><strong>{{ sayac.hazir }}</strong></article>
        <article class="istatistik kart-teslim"><span>Teslim Edildi</span><strong>{{ sayac.teslim }}</strong></article>
    </section>

    <section class="arac-cubugu panel-kutu operator-arac-cubugu">
        <div class="arama">
            <label for="kart-ara">Kart ara</label>
            <input id="kart-ara" type="search" placeholder="Talep no, stok no, talep sahibi, PCB..." autocomplete="off">
        </div>

        <div class="operator-filtre-alani">
            <div class="filtreler" role="group" aria-label="Durum filtresi">
                <button class="filtre aktif" data-filtre="AKTIF" type="button">Aktif</button>
                <button class="filtre" data-filtre="PLANA ALINDI" type="button">Plana Alındı</button>
                <button class="filtre" data-filtre="DİZGİDE" type="button">Dizgide</button>
                <button class="filtre" data-filtre="HAZIR" type="button">Hazır</button>
                <button class="filtre" data-filtre="TESLİM EDİLDİ" type="button">Teslim Edildi</button>
                <button class="filtre" data-filtre="HEPSI" type="button">Hepsi</button>
            </div>
            <div class="operator-filtre-meta">
                <span id="operator-sonuc"></span>
                <button id="operator-temizle" class="operator-temizle" type="button" hidden>Filtreleri temizle</button>
            </div>
        </div>
    </section>

    <section id="operator-kartlari" class="operator-grid">
    {% for k in kartlar %}
        <article class="operator-kart {{ k.renk }}"
                 data-kart
                 data-durum="{{ k.durum }}"
                 data-arama="{{ ((k.talep_no or '') ~ ' ' ~ (k.stok_no or '') ~ ' ' ~ (k.talep_sahibi or '') ~ ' ' ~ (k.pcb or '') ~ ' ' ~ (k.operator or '') ~ ' ' ~ (k.aciklama or ''))|lower }}">

            <div class="operator-kart-ust">
                <div>
                    <span class="ust-etiket">{{ k.talep_no or "TALEP YOK" }}</span>
                    <h2>{{ k.stok_no or "Stok no yok" }}</h2>
                    <p>{{ k.talep_sahibi or "Talep sahibi yok" }}</p>
                </div>
                <div class="operator-rozetler">
                    <span class="durum-rozet {{ k.renk }}">{{ k.rozet }}</span>
                    {% if k.kaynakta_yok %}<span class="durum-rozet uyari">Kaynak Excel'de yok</span>{% endif %}
                </div>
            </div>

            <div class="durum-satiri">
                <span><b>Durum:</b> {{ k.durum }}</span>
                <span><b>Kaynak:</b> {{ k.kaynak_durumu or "—" }}</span>
            </div>

            <div class="bilgi-grid">
                <div><span>Toplam</span><strong>{{ k.toplam_adet }}</strong></div>
                <div><span>Tamamlanan</span><strong>{{ k.tamamlanan_adet }}</strong></div>
                <div><span>Kalan</span><strong>{{ k.kalan_adet }}</strong></div>
                <div><span>Plan Teslim</span><strong>{{ k.plan_teslim|gun }}</strong></div>
            </div>

            <div class="ilerleme">
                <div class="ilerleme-ust">
                    <span>Üretim ilerlemesi</span>
                    <strong>%{{ k.adet_yuzde }}</strong>
                </div>
                <div class="ilerleme-ray">
                    <div class="ilerleme-dolgu" style="width: {{ [k.adet_yuzde, 100]|min }}%"></div>
                </div>
            </div>

            {% if k.aciklama %}<p class="kart-not"><b>Not:</b> {{ k.aciklama }}</p>{% endif %}

            <div class="kart-ek-bilgi">
                {% if k.pcb %}<span>PCB: {{ k.pcb }}</span>{% endif %}
                {% if k.operator %}<span>Son işlem: {{ k.operator }}</span>{% endif %}
                {% if k.bitis_zamani %}<span>Üretim bitiş: {{ k.bitis_zamani|gun }}</span>{% endif %}
            </div>

            <div class="kart-aksiyonlar">
                {% if k.durum == "PLANA ALINDI" %}
                    <button class="buton buton-ana" type="button" data-baslat
                            data-id="{{ k.id }}" data-kalan="{{ k.kalan_adet }}"
                            data-talep="{{ k.talep_no or '' }}" data-stok="{{ k.stok_no or '' }}">
                        Dizgiye Al
                    </button>
                {% elif k.durum == "DİZGİDE" and k.kalan_adet > 0 %}
                    <button class="buton buton-ana" type="button" data-bitir
                            data-id="{{ k.id }}" data-kalan="{{ k.kalan_adet }}"
                            data-talep="{{ k.talep_no or '' }}" data-stok="{{ k.stok_no or '' }}">
                        Adet Bitir
                    </button>
                {% elif k.durum == "DİZGİDE" and k.kalan_adet == 0 %}
                    <button class="buton buton-basari" type="button" data-hazirla
                            data-id="{{ k.id }}" data-talep="{{ k.talep_no or '' }}">
                        Hazıra Al
                    </button>
                {% elif k.durum == "HAZIR" %}
                    <button class="buton buton-basari" type="button" data-teslim
                            data-id="{{ k.id }}" data-talep="{{ k.talep_no or '' }}">
                        Teslim Edildi
                    </button>
                {% endif %}

                <button class="buton buton-hayalet" type="button" data-not
                        data-id="{{ k.id }}" data-not-mevcut="{{ k.aciklama or '' }}"
                        data-talep="{{ k.talep_no or '' }}">Not</button>
            </div>
        </article>
    {% else %}
        <div class="bos-durum tam-satir">Görüntülenecek kart yok.</div>
    {% endfor %}
    </section>

    <div id="operator-bos" class="bos-durum" hidden>Arama ve filtreye uyan kart bulunamadı.</div>
</div>

<dialog id="baslat-dialog" class="modal">
    <form method="dialog" class="modal-kutu" id="baslat-form">
        <div class="modal-baslik">
            <div><p class="ust-etiket">DİZGİYE AL</p><h2 id="baslat-baslik">Kart</h2></div>
            <button class="ikon-buton" type="button" data-dialog-kapat aria-label="Kapat">×</button>
        </div>
        <input type="hidden" id="baslat-id">
        <div class="alan">
            <label for="baslat-adet">Dizgiye alınacak adet</label>
            <input type="number" id="baslat-adet" min="1" required>
            <small id="baslat-kalan"></small>
        </div>
        <div class="alan">
            <label for="baslat-not">Not <span class="soluk">(opsiyonel)</span></label>
            <textarea id="baslat-not" rows="3"></textarea>
        </div>
        <div class="modal-aksiyon">
            <button type="button" class="buton buton-hayalet" data-dialog-kapat>Vazgeç</button>
            <button type="submit" class="buton buton-ana">Dizgiye Al</button>
        </div>
    </form>
</dialog>

<dialog id="bitir-dialog" class="modal">
    <form method="dialog" class="modal-kutu" id="bitir-form">
        <div class="modal-baslik">
            <div><p class="ust-etiket">ÜRETİM ADEDİ</p><h2 id="bitir-baslik">Kart</h2></div>
            <button class="ikon-buton" type="button" data-dialog-kapat aria-label="Kapat">×</button>
        </div>
        <input type="hidden" id="bitir-id">
        <div class="alan">
            <label for="bitir-adet">Tamamlanan adet</label>
            <input type="number" id="bitir-adet" min="1" required>
            <small id="bitir-kalan"></small>
        </div>
        <div class="alan">
            <label for="bitir-not">Not <span class="soluk">(opsiyonel)</span></label>
            <textarea id="bitir-not" rows="3"></textarea>
        </div>
        <div class="modal-aksiyon">
            <button type="button" class="buton buton-hayalet" data-dialog-kapat>Vazgeç</button>
            <button type="submit" class="buton buton-ana">Kaydet</button>
        </div>
    </form>
</dialog>

<dialog id="not-dialog" class="modal">
    <form method="dialog" class="modal-kutu" id="not-form">
        <div class="modal-baslik">
            <div><p class="ust-etiket">KART NOTU</p><h2 id="not-baslik">Not Düzenle</h2></div>
            <button class="ikon-buton" type="button" data-dialog-kapat aria-label="Kapat">×</button>
        </div>
        <input type="hidden" id="not-id">
        <div class="alan">
            <label for="not-metin">Not</label>
            <textarea id="not-metin" rows="5"></textarea>
        </div>
        <div class="modal-aksiyon">
            <button type="button" class="buton buton-hayalet" data-dialog-kapat>Vazgeç</button>
            <button type="submit" class="buton buton-ana">Kaydet</button>
        </div>
    </form>
</dialog>
{% endblock %}

{% block scripts %}
<script>
(() => {
    const kartAra = document.getElementById("kart-ara");
    const filtreButonlari = [...document.querySelectorAll("[data-filtre]")];
    const kartlar = [...document.querySelectorAll("[data-kart]")];
    const sonuc = document.getElementById("operator-sonuc");
    const temizle = document.getElementById("operator-temizle");
    const bos = document.getElementById("operator-bos");

    const FILTRE_KEY = "pdgm-op-filtre";
    const ARAMA_KEY = "pdgm-op-arama";
    const SCROLL_KEY = "pdgm-op-scroll";
    let aktifFiltre = sessionStorage.getItem(FILTRE_KEY) || "AKTIF";

    function aktifDurum(durum) {
        return ["PLANA ALINDI", "DİZGİDE", "HAZIR"].includes(durum);
    }

    function durumUygunMu(durum) {
        if (aktifFiltre === "HEPSI") return true;
        if (aktifFiltre === "AKTIF") return aktifDurum(durum);
        return durum === aktifFiltre;
    }

    function filtrele() {
        const arama = (kartAra.value || "").trim().toLocaleLowerCase("tr-TR");
        let gorunen = 0;

        for (const kart of kartlar) {
            const uygun = durumUygunMu(kart.dataset.durum) && (!arama || kart.dataset.arama.includes(arama));
            kart.hidden = !uygun;
            if (uygun) gorunen += 1;
        }

        sonuc.textContent = `${gorunen} kart gösteriliyor`;
        bos.hidden = gorunen !== 0;
        temizle.hidden = !arama && aktifFiltre === "AKTIF";
    }

    function filtreSec(deger) {
        aktifFiltre = deger;
        sessionStorage.setItem(FILTRE_KEY, deger);
        filtreButonlari.forEach((buton) => {
            const secili = buton.dataset.filtre === deger;
            buton.classList.toggle("aktif", secili);
            buton.setAttribute("aria-pressed", secili ? "true" : "false");
        });
        filtrele();
    }

    function durumuSaklaVeYenile() {
        sessionStorage.setItem(ARAMA_KEY, kartAra.value || "");
        sessionStorage.setItem(SCROLL_KEY, String(window.scrollY));
        window.setTimeout(() => window.location.reload(), 350);
    }

    async function kartAksiyonu(url, kartId, onayMesaji, basariMesaji) {
        if (onayMesaji && !window.confirm(onayMesaji)) return false;
        const data = await pdgmFetch(url, {
            method: "POST",
            body: JSON.stringify({kart_id: Number(kartId)})
        });
        toast(data.mesaj || basariMesaji);
        return true;
    }

    kartAra.value = sessionStorage.getItem(ARAMA_KEY) || "";
    kartAra.addEventListener("input", () => {
        sessionStorage.setItem(ARAMA_KEY, kartAra.value || "");
        filtrele();
    });

    filtreButonlari.forEach((buton) => buton.addEventListener("click", () => filtreSec(buton.dataset.filtre)));
    temizle.addEventListener("click", () => {
        kartAra.value = "";
        sessionStorage.removeItem(ARAMA_KEY);
        filtreSec("AKTIF");
        kartAra.focus();
    });

    const gecerliFiltre = filtreButonlari.some((buton) => buton.dataset.filtre === aktifFiltre);
    filtreSec(gecerliFiltre ? aktifFiltre : "AKTIF");

    const kayitliScroll = sessionStorage.getItem(SCROLL_KEY);
    if (kayitliScroll !== null) {
        sessionStorage.removeItem(SCROLL_KEY);
        window.scrollTo(0, Number(kayitliScroll) || 0);
    }

    document.querySelectorAll("[data-baslat]").forEach((buton) => {
        buton.addEventListener("click", () => {
            const kalan = Number(buton.dataset.kalan || 1);
            document.getElementById("baslat-id").value = buton.dataset.id;
            document.getElementById("baslat-adet").value = kalan;
            document.getElementById("baslat-adet").max = kalan;
            document.getElementById("baslat-not").value = "";
            document.getElementById("baslat-kalan").textContent = `En fazla ${kalan} adet`;
            document.getElementById("baslat-baslik").textContent = `${buton.dataset.talep || "Kart"} · ${buton.dataset.stok || ""}`;
            document.getElementById("baslat-dialog").showModal();
        });
    });

    document.querySelectorAll("[data-bitir]").forEach((buton) => {
        buton.addEventListener("click", () => {
            const kalan = Number(buton.dataset.kalan || 1);
            document.getElementById("bitir-id").value = buton.dataset.id;
            document.getElementById("bitir-adet").value = kalan;
            document.getElementById("bitir-adet").max = kalan;
            document.getElementById("bitir-not").value = "";
            document.getElementById("bitir-kalan").textContent = `Kalan ${kalan} adet`;
            document.getElementById("bitir-baslik").textContent = `${buton.dataset.talep || "Kart"} · ${buton.dataset.stok || ""}`;
            document.getElementById("bitir-dialog").showModal();
        });
    });

    document.querySelectorAll("[data-not]").forEach((buton) => {
        buton.addEventListener("click", () => {
            document.getElementById("not-id").value = buton.dataset.id;
            document.getElementById("not-metin").value = buton.dataset.notMevcut || "";
            document.getElementById("not-baslik").textContent = `${buton.dataset.talep || "Kart"} · Not`;
            document.getElementById("not-dialog").showModal();
        });
    });

    document.querySelectorAll("[data-hazirla]").forEach((buton) => {
        buton.addEventListener("click", async () => {
            try {
                const tamam = await kartAksiyonu(
                    "/api/hazirla",
                    buton.dataset.id,
                    `${buton.dataset.talep || "Bu kart"} HAZIR durumuna alınsın mı?`,
                    "Kart HAZIR durumuna alındı."
                );
                if (tamam) durumuSaklaVeYenile();
            } catch (hata) {
                hataMesaji(hata);
            }
        });
    });

    document.querySelectorAll("[data-teslim]").forEach((buton) => {
        buton.addEventListener("click", async () => {
            try {
                const tamam = await kartAksiyonu(
                    "/api/teslim-et",
                    buton.dataset.id,
                    `${buton.dataset.talep || "Bu kart"} fiziksel olarak teslim edildi mi?`,
                    "Teslim kaydedildi."
                );
                if (tamam) durumuSaklaVeYenile();
            } catch (hata) {
                hataMesaji(hata);
            }
        });
    });

    document.getElementById("baslat-form").addEventListener("submit", async (event) => {
        event.preventDefault();
        const submit = event.submitter;
        submit.disabled = true;
        try {
            const data = await pdgmFetch("/api/basla", {
                method: "POST",
                body: JSON.stringify({
                    kart_id: Number(document.getElementById("baslat-id").value),
                    adet: Number(document.getElementById("baslat-adet").value),
                    not: document.getElementById("baslat-not").value.trim()
                })
            });
            toast(data.mesaj || "Kart DİZGİDE durumuna alındı.");
            dialogKapat(document.getElementById("baslat-dialog"));
            durumuSaklaVeYenile();
        } catch (hata) {
            hataMesaji(hata);
        } finally {
            submit.disabled = false;
        }
    });

    document.getElementById("bitir-form").addEventListener("submit", async (event) => {
        event.preventDefault();
        const submit = event.submitter;
        submit.disabled = true;
        try {
            const kartId = Number(document.getElementById("bitir-id").value);
            const data = await pdgmFetch("/api/bitir", {
                method: "POST",
                body: JSON.stringify({
                    kart_id: kartId,
                    adet: Number(document.getElementById("bitir-adet").value),
                    not: document.getElementById("bitir-not").value.trim()
                })
            });
            dialogKapat(document.getElementById("bitir-dialog"));

            if (data.uretim_bitti) {
                const hazir = window.confirm(
                    "Üretim adedinin tamamı bitti. Kart HAZIR durumuna alınsın mı?\n\n" +
                    "Hayır derseniz kart DİZGİDE kalır ve daha sonra 'Hazıra Al' butonuyla geçiş yapılabilir."
                );
                if (hazir) {
                    await pdgmFetch("/api/hazirla", {
                        method: "POST",
                        body: JSON.stringify({kart_id: kartId})
                    });
                    toast("Üretim tamamlandı ve kart HAZIR durumuna alındı.");
                } else {
                    toast(data.mesaj, "uyari");
                }
            } else {
                toast(data.mesaj);
            }
            durumuSaklaVeYenile();
        } catch (hata) {
            hataMesaji(hata);
        } finally {
            submit.disabled = false;
        }
    });

    document.getElementById("not-form").addEventListener("submit", async (event) => {
        event.preventDefault();
        const submit = event.submitter;
        submit.disabled = true;
        try {
            await pdgmFetch("/api/not", {
                method: "POST",
                body: JSON.stringify({
                    kart_id: Number(document.getElementById("not-id").value),
                    not: document.getElementById("not-metin").value.trim()
                })
            });
            toast("Not kaydedildi.");
            dialogKapat(document.getElementById("not-dialog"));
            durumuSaklaVeYenile();
        } catch (hata) {
            hataMesaji(hata);
        } finally {
            submit.disabled = false;
        }
    });
})();
</script>
{% endblock %}
```


## `templates/ozet.html`


```html
{% extends "base.html" %}
{% block title %}Özet · PDGM İş Takip{% endblock %}

{% block content %}
<div class="sayfa-shell ozet-sayfa">
    <section class="sayfa-baslik sayfa-hero">
        <div>
            <p class="ust-etiket">PERFORMANS ÖZETİ</p>
            <h1>Üretim ve Teslim Özeti</h1>
            <p class="soluk">Workflow dağılımı, zamanında teslim oranı ve son sekiz haftalık plan/teslim görünümü.</p>
        </div>
    </section>

    <section class="istatistik-grid">
        <article class="istatistik kart-plana"><span>Plana Alındı</span><strong>{{ genel.plana_alindi }}</strong></article>
        <article class="istatistik kart-dizgide"><span>Dizgide</span><strong>{{ genel.dizgide }}</strong></article>
        <article class="istatistik kart-hazir"><span>Hazır</span><strong>{{ genel.hazir }}</strong></article>
        <article class="istatistik kart-teslim"><span>Teslim Edildi</span><strong>{{ genel.teslim }}</strong></article>
    </section>

    {% if genel.gecikme or genel.durumu_eksik %}
    <section class="ozet-uyari-satiri">
        {% if genel.gecikme %}<span class="durum-rozet kotu">{{ genel.gecikme }} geciken açık kart</span>{% endif %}
        {% if genel.durumu_eksik %}<span class="durum-rozet uyari">{{ genel.durumu_eksik }} durumu eksik kart</span>{% endif %}
    </section>
    {% endif %}

    <section class="donem-grid">
    {% for ad, d in donemler.items() %}
        <article class="panel-kutu ozet-donem-karti">
            <p class="ust-etiket">{{ ad|upper }}</p>
            <h2>{{ d.kart }} teslim</h2>
            <div class="donem-metrik"><span>Teslim edilen adet</span><strong>{{ d.adet }}</strong></div>
            <div class="donem-metrik"><span>Zamanında teslim</span><strong>%{{ d.zamaninda_yuzde }}</strong></div>
            <div class="donem-metrik"><span>Ort. teslim sapması</span><strong>{{ d.ort_sapma }} gün</strong></div>
        </article>
    {% endfor %}
    </section>

    <section class="panel-kutu ozet-grafik-panel">
        <div class="panel-baslik">
            <div>
                <p class="ust-etiket">8 HAFTALIK GÖRÜNÜM</p>
                <h2>Planlanan Teslim / Gerçekleşen Teslim</h2>
            </div>
        </div>
        <div class="grafik-aciklama">
            <span><i class="lejant plan"></i>Planlanan</span>
            <span><i class="lejant teslim"></i>Teslim Edilen</span>
        </div>
        <div class="hafta-grafik" role="img" aria-label="Son sekiz hafta planlanan ve teslim edilen kart grafiği">
        {% for h in haftalar %}
            {% set plan_yuzde = (h.planlanan / en_yuksek * 100) if en_yuksek else 0 %}
            {% set teslim_yuzde = (h.teslim / en_yuksek * 100) if en_yuksek else 0 %}
            <div class="hafta-sutun">
                <div class="sutun-alani">
                    <div class="sutun plan" style="height: {{ plan_yuzde }}%"><span>{{ h.planlanan }}</span></div>
                    <div class="sutun teslim" style="height: {{ teslim_yuzde }}%"><span>{{ h.teslim }}</span></div>
                </div>
                <strong>H{{ h.hafta_no }}</strong>
                <small>{{ h.etiket }}</small>
            </div>
        {% endfor %}
        </div>
    </section>

    <section class="panel-kutu ozet-geciken-panel">
        <div class="panel-baslik">
            <div>
                <p class="ust-etiket">DİKKAT GEREKTİREN</p>
                <h2>Geciken Açık Kartlar</h2>
            </div>
            <span class="sayi-rozet">{{ geciken_kartlar|length }}</span>
        </div>

        {% if geciken_kartlar %}
        <div class="tablo-kapsayici">
            <table>
                <thead><tr><th>Talep NO</th><th>Stok No</th><th>Durum</th><th>Plan Başlangıç</th><th>Plan Teslim</th><th>İlerleme</th><th>Değerlendirme</th></tr></thead>
                <tbody>
                {% for k in geciken_kartlar %}
                    <tr>
                        <td>{{ k.talep_no or "—" }}</td>
                        <td>{{ k.stok_no or "—" }}</td>
                        <td>{{ k.durum }}</td>
                        <td>{{ k.plan_baslama|gun }}</td>
                        <td>{{ k.plan_teslim|gun }}</td>
                        <td>{{ k.tamamlanan_adet }}/{{ k.toplam_adet }}</td>
                        <td><span class="durum-rozet kotu">{{ k.rozet }}</span></td>
                    </tr>
                {% endfor %}
                </tbody>
            </table>
        </div>
        {% else %}
        <div class="bos-durum">Şu anda geciken açık kart bulunmuyor.</div>
        {% endif %}
    </section>
</div>
{% endblock %}
```


## `templates/panel.html`


```html
{% extends "base.html" %}
{% block title %}Pano · PDGM İş Takip{% endblock %}

{% block content %}
<div class="sayfa-shell panel-sayfa">
    <section class="sayfa-baslik sayfa-hero panel-hero">
        <div>
            <p class="ust-etiket">CANLI ÜRETİM PANOSU</p>
            <h1>İş Durumu</h1>
            <div class="panel-canli-satir">
                <span class="panel-canli">Güncel pano</span>
                <span class="panel-guncelleme">Son görüntüleme: {{ guncelleme }} · {{ bugun }}</span>
                {% if sayac.gecikme %}<span class="durum-rozet kotu">{{ sayac.gecikme }} geciken açık kart</span>{% endif %}
            </div>
        </div>
        <button id="panel-yenile" class="buton buton-hayalet" type="button">Yenile</button>
    </section>

    <section class="istatistik-grid" aria-label="İş durumu özeti">
        <article class="istatistik kart-plana"><span>Plana Alındı</span><strong>{{ sayac.plana_alindi }}</strong></article>
        <article class="istatistik kart-dizgide"><span>Dizgide</span><strong>{{ sayac.dizgide }}</strong></article>
        <article class="istatistik kart-hazir"><span>Hazır</span><strong>{{ sayac.hazir }}</strong></article>
        <article class="istatistik kart-teslim"><span>Teslim Edildi</span><strong>{{ sayac.teslim }}</strong></article>
    </section>

    <section class="arac-cubugu panel-kutu panel-arac-cubugu">
        <div class="arama panel-arama-alani">
            <label for="panel-kart-ara">Kart ara</label>
            <input id="panel-kart-ara" type="search" placeholder="Talep no, stok no, talep sahibi, operatör..." autocomplete="off">
        </div>
        <div class="panel-filtre-alani">
            <div class="filtreler" role="group" aria-label="Durum filtresi">
                <button class="filtre" data-panel-filtre="AKTIF" type="button">Aktif</button>
                <button class="filtre" data-panel-filtre="PLANA ALINDI" type="button">Plana Alındı</button>
                <button class="filtre" data-panel-filtre="DİZGİDE" type="button">Dizgide</button>
                <button class="filtre" data-panel-filtre="HAZIR" type="button">Hazır</button>
                <button class="filtre" data-panel-filtre="TESLİM EDİLDİ" type="button">Teslim Edildi</button>
                <button class="filtre aktif" data-panel-filtre="HEPSI" type="button">Hepsi</button>
            </div>
            <div class="panel-filtre-alt">
                <span id="panel-sonuc-sayisi"></span>
                <button id="panel-temizle" class="metin-buton" type="button" hidden>Filtreleri temizle</button>
            </div>
        </div>
    </section>

    <section id="panel-aktif-grid" class="pano-grid panel-aktif-grid">
        <article class="panel-kutu panel-bolum" data-panel-bolum data-durum="DİZGİDE">
            <div class="panel-baslik">
                <div>
                    <p class="ust-etiket">AKTİF ÜRETİM</p>
                    <h2>Dizgide</h2>
                    <p class="panel-aciklama">Üretimi devam eden kartlar.</p>
                </div>
                <span class="sayi-rozet">{{ dizgide|length }}</span>
            </div>

            {% if dizgide %}
            <div class="kart-listesi">
                {% for k in dizgide %}
                <article class="is-karti panel-is-karti {{ k.renk }}" data-panel-kart data-durum="DİZGİDE"
                         data-arama="{{ ((k.talep_no or '') ~ ' ' ~ (k.stok_no or '') ~ ' ' ~ (k.talep_sahibi or '') ~ ' ' ~ (k.operator or '') ~ ' ' ~ (k.aciklama or '') ~ ' ' ~ (k.pcb or ''))|lower }}">
                    <div class="is-karti-ust">
                        <div>
                            <strong>{{ k.talep_no or "Talep yok" }}</strong>
                            <span>{{ k.stok_no or "Stok no yok" }}</span>
                        </div>
                        <span class="durum-rozet {{ k.renk }}">{{ k.rozet }}</span>
                    </div>
                    <div class="kart-bilgiler">
                        <span><b>Sahibi:</b> {{ k.talep_sahibi or "—" }}</span>
                        <span><b>Plan teslim:</b> {{ k.plan_teslim|gun }}</span>
                        <span><b>Operatör:</b> {{ k.operator or "—" }}</span>
                    </div>
                    <div class="ilerleme">
                        <div class="ilerleme-ust"><span>{{ k.tamamlanan_adet }}/{{ k.toplam_adet }} adet</span><strong>%{{ k.adet_yuzde }}</strong></div>
                        <div class="ilerleme-ray"><div class="ilerleme-dolgu" style="width: {{ [k.adet_yuzde, 100]|min }}%"></div></div>
                    </div>
                    {% if k.aciklama %}<p class="kart-not">{{ k.aciklama }}</p>{% endif %}
                </article>
                {% endfor %}
            </div>
            {% else %}
            <div class="bos-durum">Şu anda dizgide iş yok.</div>
            {% endif %}
        </article>

        <article class="panel-kutu panel-bolum panel-hazir-kutu" data-panel-bolum data-durum="HAZIR">
            <div class="panel-baslik">
                <div>
                    <p class="ust-etiket">TESLİM BEKLİYOR</p>
                    <h2>Hazır</h2>
                    <p class="panel-aciklama">Üretimi bitmiş, teslim alınmayı bekleyen kartlar.</p>
                </div>
                <span class="sayi-rozet">{{ hazir|length }}</span>
            </div>

            {% if hazir %}
            <div class="mini-liste">
                {% for k in hazir %}
                <div class="mini-satir panel-mini-satir" data-panel-kart data-durum="HAZIR"
                     data-arama="{{ ((k.talep_no or '') ~ ' ' ~ (k.stok_no or '') ~ ' ' ~ (k.talep_sahibi or '') ~ ' ' ~ (k.operator or '') ~ ' ' ~ (k.aciklama or ''))|lower }}">
                    <div>
                        <strong>{{ k.talep_no or "—" }} · {{ k.stok_no or "—" }}</strong>
                        <span>{{ k.talep_sahibi or "Talep sahibi yok" }}</span>
                        <small>{{ k.toplam_adet }} adet · Plan teslim {{ k.plan_teslim|gun }}</small>
                    </div>
                    <div class="mini-sag">
                        <span class="durum-rozet {{ k.renk }}">{{ k.rozet }}</span>
                        <small>Üretim bitiş {{ k.bitis_zamani|gun }}</small>
                    </div>
                </div>
                {% endfor %}
            </div>
            {% else %}
            <div class="bos-durum">Teslim bekleyen hazır kart yok.</div>
            {% endif %}
        </article>
    </section>

    <section class="panel-kutu panel-bolum" data-panel-bolum data-durum="PLANA ALINDI">
        <div class="panel-baslik">
            <div>
                <p class="ust-etiket">ÜRETİM SIRASI</p>
                <h2>Plana Alınan İşler</h2>
                <p class="panel-aciklama">Dizgiye alınmayı bekleyen ilk 12 kart.</p>
            </div>
            <span class="sayi-rozet">{{ plana_alindi|length }}</span>
        </div>

        {% if plana_alindi %}
        <div class="plan-liste">
            {% for k in plana_alindi %}
            <div class="plan-satir" data-panel-kart data-durum="PLANA ALINDI"
                 data-arama="{{ ((k.talep_no or '') ~ ' ' ~ (k.stok_no or '') ~ ' ' ~ (k.talep_sahibi or '') ~ ' ' ~ (k.pcb or ''))|lower }}">
                <div class="plan-sira">{{ k.sira or loop.index }}</div>
                <div>
                    <strong>{{ k.talep_no or "—" }}</strong>
                    <span>{{ k.stok_no or "—" }}</span>
                </div>
                <span>{{ k.talep_sahibi or "—" }}</span>
                <span>{{ k.toplam_adet }} adet</span>
                <span>{{ k.plan_baslama|gun }}</span>
                <span class="durum-rozet {{ k.renk }}">{{ k.rozet }}</span>
            </div>
            {% endfor %}
        </div>
        {% else %}
        <div class="bos-durum">Plana alınmış bekleyen kart yok.</div>
        {% endif %}
    </section>

    <section class="panel-kutu panel-bolum" data-panel-bolum data-durum="TESLİM EDİLDİ">
        <div class="panel-baslik">
            <div>
                <p class="ust-etiket">SON HAREKETLER</p>
                <h2>Son Teslim Edilenler</h2>
                <p class="panel-aciklama">Yakın zamanda fiziksel teslimi tamamlanan kartlar.</p>
            </div>
            <span class="sayi-rozet">{{ teslim_edilen|length }}</span>
        </div>

        {% if teslim_edilen %}
        <div class="tablo-kapsayici">
            <table>
                <thead><tr><th>Talep No</th><th>Stok No</th><th>Adet</th><th>Teslim</th><th>Operatör</th><th>Değerlendirme</th></tr></thead>
                <tbody>
                {% for k in teslim_edilen %}
                    <tr data-panel-kart data-durum="TESLİM EDİLDİ"
                        data-arama="{{ ((k.talep_no or '') ~ ' ' ~ (k.stok_no or '') ~ ' ' ~ (k.talep_sahibi or '') ~ ' ' ~ (k.operator or '') ~ ' ' ~ (k.aciklama or ''))|lower }}">
                        <td><strong>{{ k.talep_no or "—" }}</strong></td>
                        <td>{{ k.stok_no or "—" }}</td>
                        <td>{{ k.toplam_adet }}</td>
                        <td>{{ (k.gerceklesen_teslim or k.teslim_zamani)|gun }}</td>
                        <td>{{ k.operator or "—" }}</td>
                        <td><span class="durum-rozet {{ k.renk }}">{{ k.rozet }}</span></td>
                    </tr>
                {% endfor %}
                </tbody>
            </table>
        </div>
        {% else %}
        <div class="bos-durum">Henüz teslim edilen kart yok.</div>
        {% endif %}
    </section>

    <div id="panel-arama-bos" class="bos-durum" hidden>Arama ve filtreye uyan kart bulunamadı.</div>
</div>
{% endblock %}

{% block scripts %}
<script>
(() => {
    const arama = document.getElementById("panel-kart-ara");
    const filtreler = [...document.querySelectorAll("[data-panel-filtre]")];
    const kartlar = [...document.querySelectorAll("[data-panel-kart]")];
    const bolumler = [...document.querySelectorAll("[data-panel-bolum]")];
    const sonuc = document.getElementById("panel-sonuc-sayisi");
    const temizle = document.getElementById("panel-temizle");
    const bos = document.getElementById("panel-arama-bos");
    const aktifGrid = document.getElementById("panel-aktif-grid");

    const FILTRE_KEY = "pdgm-panel-filtre";
    const ARAMA_KEY = "pdgm-panel-arama";
    const SCROLL_KEY = "pdgm-panel-scroll";
    let aktifFiltre = sessionStorage.getItem(FILTRE_KEY) || "HEPSI";

    function durumUygun(durum) {
        if (aktifFiltre === "HEPSI") return true;
        if (aktifFiltre === "AKTIF") return ["PLANA ALINDI", "DİZGİDE", "HAZIR"].includes(durum);
        return durum === aktifFiltre;
    }

    function filtrele() {
        const metin = (arama.value || "").trim().toLocaleLowerCase("tr-TR");
        let gorunen = 0;

        kartlar.forEach((kart) => {
            const uygun = durumUygun(kart.dataset.durum) && (!metin || kart.dataset.arama.includes(metin));
            kart.hidden = !uygun;
            if (uygun) gorunen += 1;
        });

        bolumler.forEach((bolum) => {
            const bolumKartlari = [...bolum.querySelectorAll("[data-panel-kart]")];
            const durumUygunMu = durumUygun(bolum.dataset.durum);
            bolum.hidden = !durumUygunMu || (bolumKartlari.length > 0 && !bolumKartlari.some((kart) => !kart.hidden));
        });

        if (aktifGrid) {
            const gorunenBolumler = [...aktifGrid.querySelectorAll("[data-panel-bolum]")].filter((b) => !b.hidden);
            aktifGrid.hidden = gorunenBolumler.length === 0;
            [...aktifGrid.querySelectorAll("[data-panel-bolum]")].forEach((b) => {
                b.style.gridColumn = gorunenBolumler.length === 1 && !b.hidden ? "1 / -1" : "";
            });
        }

        sonuc.textContent = gorunen ? `${gorunen} kart gösteriliyor` : "Sonuç bulunamadı";
        bos.hidden = gorunen !== 0;
        temizle.hidden = !metin && aktifFiltre === "HEPSI";
    }

    function filtreSec(deger) {
        aktifFiltre = deger;
        sessionStorage.setItem(FILTRE_KEY, deger);
        filtreler.forEach((buton) => {
            const secili = buton.dataset.panelFiltre === deger;
            buton.classList.toggle("aktif", secili);
            buton.setAttribute("aria-pressed", secili ? "true" : "false");
        });
        filtrele();
    }

    function scrollKaydet() {
        sessionStorage.setItem(SCROLL_KEY, String(window.scrollY));
    }

    arama.value = sessionStorage.getItem(ARAMA_KEY) || "";
    arama.addEventListener("input", () => {
        sessionStorage.setItem(ARAMA_KEY, arama.value || "");
        filtrele();
    });
    filtreler.forEach((buton) => buton.addEventListener("click", () => filtreSec(buton.dataset.panelFiltre)));
    temizle.addEventListener("click", () => {
        arama.value = "";
        sessionStorage.removeItem(ARAMA_KEY);
        filtreSec("HEPSI");
        arama.focus();
    });
    document.getElementById("panel-yenile").addEventListener("click", () => {
        scrollKaydet();
        window.location.reload();
    });

    const kayitliScroll = sessionStorage.getItem(SCROLL_KEY);
    if (kayitliScroll !== null) {
        sessionStorage.removeItem(SCROLL_KEY);
        window.scrollTo(0, Number(kayitliScroll) || 0);
    }

    filtreSec(filtreler.some((b) => b.dataset.panelFiltre === aktifFiltre) ? aktifFiltre : "HEPSI");
    window.setTimeout(() => {
        scrollKaydet();
        window.location.reload();
    }, 30000);
})();
</script>
{% endblock %}
```


## `templates/yetkisiz.html`


```html
{% extends "base.html" %}
{% block title %}Yetkisiz · PDGM İş Takip{% endblock %}
{% block content %}
<div class="yetkisiz-sayfa">
    <section class="durum-sayfasi">
        <p class="ust-etiket">ERİŞİM KISITLI</p>
        <div class="durum-ikon">403</div>
        <h1>Bu sayfaya erişim yetkiniz yok.</h1>
        <p>Mevcut hesabınız bu işlemi gerçekleştirmek için gerekli role sahip değil.</p>
        <a class="buton buton-ana" href="{{ url_for('ana') }}">Ana Sayfaya Dön</a>
    </section>
</div>
{% endblock %}
```


## `templates/yonetim.html`


```html
{% extends "base.html" %}
{% block title %}Yönetim · PDGM İş Takip{% endblock %}

{% block content %}
<div class="sayfa-shell yonetim-sayfa">
    <section class="sayfa-baslik sayfa-hero">
        <div>
            <p class="ust-etiket">YÖNETİCİ PANELİ</p>
            <h1>Sistem Yönetimi</h1>
            <p class="soluk">İlk Excel aktarımı, manuel kart yönetimi, kartlar.xlsx bakımı, yedekler ve audit geçmişi.</p>
        </div>
        <div class="baslik-aksiyon">
            <a class="buton buton-hayalet" href="{{ url_for('rapor_indir') }}">Rapor İndir</a>
            <a class="buton buton-ana" href="{{ url_for('panel') }}">Canlı Panoyu Aç</a>
        </div>
    </section>

    <section class="yonetim-grid">
        <article class="panel-kutu yonetim-yukleme-karti">
            <div class="panel-baslik">
                <div>
                    <p class="ust-etiket">VERİ AKTARIMI</p>
                    <h2>Plan Excel'ini Aktar</h2>
                    <p class="panel-aciklama">MAKİNE sayfası okunur; hidden kolonlar atlanır.</p>
                </div>
            </div>
            <form method="post" action="{{ url_for('yukle') }}" enctype="multipart/form-data" class="yukleme-form">
                <input type="hidden" name="_csrf_token" value="{{ csrf_token }}">
                <label class="dosya-sec">
                    <span>Excel dosyası seçin</span>
                    <small>.xlsx veya .xlsm</small>
                    <input type="file" name="dosya" accept=".xlsx,.xlsm" required>
                </label>
                <button class="buton buton-ana" type="submit">Excel'i Aktar</button>
            </form>
            <p class="yardim">
                Kaynak Excel ilk kurulumda kartları oluşturmak için kullanılabilir. Sonrasında günlük operasyonun source of truth'u
                <code>data/kartlar.xlsx</code> olur. Tekrar import yapılırsa mevcut workflow, tamamlanan adet ve operatör işlemleri korunur;
                yalnız plan/source alanları güncellenir.
            </p>
        </article>

        <article class="panel-kutu yonetim-bakim-karti">
            <div class="panel-baslik">
                <div>
                    <p class="ust-etiket">BAKIM</p>
                    <h2>Kayıt Dosyaları</h2>
                    <p class="panel-aciklama">Excel'e doğrudan müdahale gerektiğinde kullanılacak dosyalar.</p>
                </div>
            </div>
            <div class="buton-grup dikey">
                <a class="buton buton-hayalet" href="{{ url_for('kayit_dosyasi', hangi='kartlar') }}">Kartlar Excel</a>
                <a class="buton buton-hayalet" href="{{ url_for('kayit_dosyasi', hangi='log') }}">İşlem Logu</a>
                <a class="buton buton-hayalet" href="{{ url_for('kayit_dosyasi', hangi='yuklemeler') }}">Yükleme Geçmişi</a>
                <form method="post" action="{{ url_for('yeniden_oku') }}"
                      onsubmit="return confirm('Diskteki kartlar.xlsx doğrulanarak yeniden okunacak. Devam edilsin mi?')">
                    <input type="hidden" name="_csrf_token" value="{{ csrf_token }}">
                    <button class="buton buton-uyari tam-genislik" type="submit">Kart Dosyasını Yeniden Oku</button>
                </form>
            </div>
            <p class="yardim">Manuel Excel düzenlemesi yaparken önce dosyayı kaydedip Excel'de kapatın; ardından "Kart Dosyasını Yeniden Oku" kullanın.</p>
        </article>
    </section>

    {% if durumu_eksik %}
    <section class="panel-kutu yonetim-durum-eksik">
        <div class="panel-baslik">
            <div>
                <p class="ust-etiket">KONTROL GEREKİYOR</p>
                <h2>Durumu Eksik Kartlar</h2>
                <p class="panel-aciklama">Kaynak Excel'de DURUM boş veya geçersiz olduğu için Pano ve Operatör ekranında gösterilmezler.</p>
            </div>
            <span class="sayi-rozet">{{ durumu_eksik|length }}</span>
        </div>

        <div class="mini-liste yonetim-durum-eksik-liste">
            {% for k in durumu_eksik %}
            <div class="mini-satir">
                <div>
                    <strong>{{ k.talep_no or "—" }} · {{ k.stok_no or "—" }}</strong>
                    <span>{{ k.talep_sahibi or "Talep sahibi yok" }} · {{ k.toplam_adet }} adet</span>
                    <small>Kaynak DURUM: {{ k.excel_durum or "boş" }}</small>
                </div>
                <div class="mini-sag">
                    <span class="durum-rozet uyari">DURUMU EKSİK</span>
                    <button class="buton buton-kucuk buton-ana" type="button" data-admin-duzenle
                            data-id="{{ k.id }}" data-talep="{{ k.talep_no or '' }}" data-stok="{{ k.stok_no or '' }}"
                            data-durum="" data-toplam="{{ k.toplam_adet }}" data-tamamlanan="{{ k.tamamlanan_adet }}"
                            data-plan-hafta="{{ k.plan_hafta or '' }}" data-plan-baslama="{{ k.plan_baslama or '' }}"
                            data-plan-teslim="{{ k.plan_teslim or '' }}" data-gerceklesen-teslim="{{ k.gerceklesen_teslim or '' }}"
                            data-not="{{ k.aciklama or '' }}">Durum Ata</button>
                </div>
            </div>
            {% endfor %}
        </div>
    </section>
    {% endif %}

    {% if kaynakta_olmayan %}
    <section class="panel-kutu yonetim-uyari-kutu">
        <div class="panel-baslik">
            <div>
                <p class="ust-etiket">KAYNAK KONTROLÜ</p>
                <h2>Son Excel'de Olmayan Açık Kartlar</h2>
                <p class="panel-aciklama">Kartlar silinmez; mevcut operasyon state'i korunur.</p>
            </div>
            <span class="sayi-rozet">{{ kaynakta_olmayan|length }}</span>
        </div>
        <div class="mini-liste">
            {% for k in kaynakta_olmayan %}
            <div class="mini-satir">
                <div>
                    <strong>{{ k.talep_no or "—" }} · {{ k.stok_no or "—" }}</strong>
                    <span>{{ k.is_durumu }} · {{ k.tamamlanan_adet }}/{{ k.toplam_adet }} adet</span>
                </div>
                <div class="mini-sag"><span class="durum-rozet uyari">Kaynak Excel'de yok</span></div>
            </div>
            {% endfor %}
        </div>
    </section>
    {% endif %}

    <section class="yonetim-grid">
        <article class="panel-kutu">
            <div class="panel-baslik">
                <div>
                    <p class="ust-etiket">KART YÖNETİMİ</p>
                    <h2>Gizlenen Kartlar</h2>
                    <p class="panel-aciklama">Gizlemek silme işlemi değildir; kart kartlar.xlsx içinde kalır.</p>
                </div>
                <span class="sayi-rozet">{{ gizlenen_kartlar|length }}</span>
            </div>

            {% if gizlenen_kartlar %}
            <div class="mini-liste">
                {% for k in gizlenen_kartlar %}
                <div class="mini-satir">
                    <div>
                        <strong>{{ k.talep_no or "—" }} · {{ k.stok_no or "—" }}</strong>
                        <span>{{ k.is_durumu }} · {{ k.tamamlanan_adet }}/{{ k.toplam_adet }} adet</span>
                    </div>
                    <div class="mini-sag">
                        <span class="durum-rozet">Gizli</span>
                        <form method="post" action="{{ url_for('kart_geri_getir') }}"
                              onsubmit="return confirm('Bu kart tekrar aktif listelere alınsın mı?')">
                            <input type="hidden" name="_csrf_token" value="{{ csrf_token }}">
                            <input type="hidden" name="kart_id" value="{{ k.id }}">
                            <button type="submit" class="buton buton-kucuk buton-basari">Geri Getir</button>
                        </form>
                    </div>
                </div>
                {% endfor %}
            </div>
            {% else %}
            <div class="bos-durum">Gizlenmiş kart bulunmuyor.</div>
            {% endif %}
        </article>

        <article class="panel-kutu">
            <div class="panel-baslik">
                <div>
                    <p class="ust-etiket">KURTARMA</p>
                    <h2>Yedekten Geri Yükle</h2>
                    <p class="panel-aciklama">Yalnız kart verisi geri alınır; audit logu geriye sarılmaz.</p>
                </div>
            </div>

            {% if yedekler %}
            <div class="mini-liste">
                {% for y in yedekler %}
                <div class="mini-satir">
                    <div>
                        <strong>{{ y.zaman }}</strong>
                        <span>{{ y.tip }} · {{ y.etiket }}</span>
                        <small>{{ y.boyut_kb }} KB</small>
                    </div>
                    <div class="mini-sag">
                        <form method="post" action="{{ url_for('yedek_geri_yukle') }}"
                              onsubmit="return confirm('Mevcut kart verileri seçilen yedekle değiştirilecek. Mevcut durum önce ayrıca yedeklenecek. Devam edilsin mi?')">
                            <input type="hidden" name="_csrf_token" value="{{ csrf_token }}">
                            <input type="hidden" name="yedek" value="{{ y.ad }}">
                            <button type="submit" class="buton buton-kucuk buton-uyari">Geri Yükle</button>
                        </form>
                    </div>
                </div>
                {% endfor %}
            </div>
            {% else %}
            <div class="bos-durum">Henüz kullanılabilir yedek bulunmuyor.</div>
            {% endif %}
        </article>
    </section>

    <section class="panel-kutu yonetim-kart-panel">
        <div class="panel-baslik">
            <div>
                <p class="ust-etiket">KART YÖNETİMİ</p>
                <h2>Kartlar</h2>
                <p class="panel-aciklama">
                    Gizlenmemiş tüm aktif kayıtlar; durumu eksik kayıtlar dahil.
                </p>
            </div>

            <div class="baslik-aksiyon">
                <button id="admin-yeni-ac"
                        class="buton buton-ana"
                        type="button">
                    + Yeni Kart
                </button>

                <div class="arama kucuk">
                    <label class="sr-only" for="admin-kart-ara">Kart ara</label>
                    <input id="admin-kart-ara"
                           type="search"
                           placeholder="Talep, stok, kişi..."
                           autocomplete="off">
                </div>
            </div>
        </div>

        <div class="admin-filtre-cubugu">
            <div class="filtreler"
                 role="group"
                 aria-label="Yönetim kart durum filtresi">

                <button class="filtre aktif"
                        data-admin-filtre="HEPSI"
                        type="button">
                    Hepsi
                </button>

                <button class="filtre"
                        data-admin-filtre="AKTIF"
                        type="button">
                    Açık İşler
                </button>

                <button class="filtre"
                        data-admin-filtre="PLANA ALINDI"
                        type="button">
                    Plana Alındı
                </button>

                <button class="filtre"
                        data-admin-filtre="DİZGİDE"
                        type="button">
                    Dizgide
                </button>

                <button class="filtre"
                        data-admin-filtre="HAZIR"
                        type="button">
                    Hazır
                </button>

                <button class="filtre"
                        data-admin-filtre="TESLİM EDİLDİ"
                        type="button">
                    Teslim Edildi
                </button>

                <button class="filtre"
                        data-admin-filtre="DURUMU EKSİK"
                        type="button">
                    Durumu Eksik
                </button>
            </div>

            <span id="admin-sonuc" class="admin-filtre-sonuc">
                {{ kartlar|length }} kart
            </span>
        </div>

        <div class="tablo-kapsayici yonetim-kart-scroll">
            <table id="admin-kart-tablosu">
                <thead>
                    <tr>
                        <th>ID</th>
                        <th>Talep NO</th>
                        <th>Stok No</th>
                        <th>Talep Sahibi</th>
                        <th>Adet</th>
                        <th>Durum</th>
                        <th>Kaynak Durumu</th>
                        <th>Plan Teslim</th>
                        <th>Operatör</th>
                        <th>İşlem</th>
                    </tr>
                </thead>

                <tbody>
                {% for k in kartlar %}
                    <tr data-admin-kart
                        data-durum="{{ k.durum or 'DURUMU EKSİK' }}"
                        data-arama="{{ (
                            (k.id|string)
                            ~ ' ' ~ (k.talep_no or '')
                            ~ ' ' ~ (k.stok_no or '')
                            ~ ' ' ~ (k.talep_sahibi or '')
                            ~ ' ' ~ (k.durum or 'DURUMU EKSİK')
                            ~ ' ' ~ (k.excel_durum or '')
                        )|lower }}">

                        <td>{{ k.id }}</td>

                        <td>
                            <strong>{{ k.talep_no or "—" }}</strong>
                        </td>

                        <td>{{ k.stok_no or "—" }}</td>

                        <td>{{ k.talep_sahibi or "—" }}</td>

                        <td>
                            {{ k.tamamlanan_adet }}/{{ k.toplam_adet }}
                        </td>

                        <td>
                            {% if k.durum %}
                                <span class="durum-rozet {{ k.renk }}">
                                    {{ k.durum }}
                                </span>
                            {% else %}
                                <span class="durum-rozet uyari">
                                    DURUMU EKSİK
                                </span>
                            {% endif %}

                            {% if k.kaynakta_yok %}
                                <div class="satir-alt-rozet">
                                    <span class="durum-rozet uyari">
                                        Kaynakta yok
                                    </span>
                                </div>
                            {% endif %}
                        </td>

                        <td>{{ k.kaynak_durumu or "—" }}</td>

                        <td>{{ k.plan_teslim|gun }}</td>

                        <td>{{ k.operator or "—" }}</td>

                        <td>
                            <div class="satir-aksiyon">
                                <button class="buton buton-kucuk buton-hayalet"
                                        type="button"
                                        data-admin-duzenle
                                        data-id="{{ k.id }}"
                                        data-talep="{{ k.talep_no or '' }}"
                                        data-stok="{{ k.stok_no or '' }}"
                                        data-durum="{{ k.durum or '' }}"
                                        data-toplam="{{ k.toplam_adet }}"
                                        data-tamamlanan="{{ k.tamamlanan_adet }}"
                                        data-plan-hafta="{{ k.plan_hafta or '' }}"
                                        data-plan-baslama="{{ k.plan_baslama or '' }}"
                                        data-plan-teslim="{{ k.plan_teslim or '' }}"
                                        data-gerceklesen-teslim="{{ k.gerceklesen_teslim or '' }}"
                                        data-not="{{ k.aciklama or '' }}">
                                    Düzenle
                                </button>

                                <button class="buton buton-kucuk buton-tehlike"
                                        type="button"
                                        data-admin-gizle
                                        data-id="{{ k.id }}"
                                        data-talep="{{ k.talep_no or '' }}">
                                    Gizle
                                </button>
                            </div>
                        </td>
                    </tr>
                {% else %}
                    <tr>
                        <td colspan="10" class="bos-durum">
                            Kart bulunmuyor.
                        </td>
                    </tr>
                {% endfor %}
                </tbody>
            </table>
        </div>

        <div id="admin-kart-bos"
             class="bos-durum admin-kart-bos"
             hidden>
            Arama ve filtreye uyan kart bulunamadı.
        </div>
    </section>

    <section class="yonetim-grid yonetim-gecmis">
        <article class="panel-kutu">
            <div class="panel-baslik"><div><p class="ust-etiket">SON YÜKLEMELER</p><h2>Excel Geçmişi</h2></div></div>
            {% if yuklemeler %}
            <div class="mini-liste">
                {% for y in yuklemeler %}
                <div class="mini-satir">
                    <div><strong>{{ y.dosya or "—" }}</strong><span>{{ y.zaman or "—" }} · {{ y.kullanici or "—" }}</span></div>
                    <div class="mini-sag"><strong>{{ y.satir or 0 }} satır</strong><small>{{ y.yeni or 0 }} yeni · {{ y.guncellenen or 0 }} güncel · {{ y.uyari or 0 }} uyarı</small></div>
                </div>
                {% endfor %}
            </div>
            {% else %}<div class="bos-durum">Henüz Excel yüklenmemiş.</div>{% endif %}
        </article>

        <article class="panel-kutu">
            <div class="panel-baslik"><div><p class="ust-etiket">AUDIT</p><h2>Son İşlemler</h2></div></div>
            {% if loglar %}
            <div class="log-liste">
                {% for l in loglar %}
                <div class="log-satir">
                    <div><strong>{{ l.islem or "İşlem" }}</strong><span>{{ l.kullanici or "—" }} · {{ l.zaman or "—" }}</span></div>
                    <small>{% if l.talep_no %}{{ l.talep_no }}{% endif %}{% if l.stok_no %} · {{ l.stok_no }}{% endif %}{% if l.detay %} · {{ l.detay }}{% endif %}</small>
                </div>
                {% endfor %}
            </div>
            {% else %}<div class="bos-durum">Henüz işlem kaydı yok.</div>{% endif %}
        </article>
    </section>
</div>

<dialog id="admin-dialog" class="modal">
    <form method="dialog" class="modal-kutu" id="admin-form">
        <div class="modal-baslik">
            <div><p class="ust-etiket">ADMİN MÜDAHALESİ</p><h2 id="admin-baslik">Kart Düzenle</h2></div>
            <button class="ikon-buton" type="button" data-dialog-kapat aria-label="Kapat">×</button>
        </div>
        <input type="hidden" id="admin-id">

        <div class="iki-kolon">
            <div class="alan">
                <label for="admin-durum">Durum</label>
                <select id="admin-durum" required>
                    <option value="" disabled>Durum seçin</option>
                    <option value="PLANA ALINDI">PLANA ALINDI</option>
                    <option value="DİZGİDE">DİZGİDE</option>
                    <option value="HAZIR">HAZIR</option>
                    <option value="TESLİM EDİLDİ">TESLİM EDİLDİ</option>
                </select>
            </div>
            <div class="alan">
                <label for="admin-toplam">Toplam adet</label>
                <input type="number" id="admin-toplam" min="1" required>
            </div>
        </div>

        <div class="alan">
            <label for="admin-tamamlanan">Tamamlanan adet</label>
            <input type="number" id="admin-tamamlanan" min="0" required>
            <small>HAZIR ve TESLİM EDİLDİ seçildiğinde sistem otomatik olarak toplam adede eşitler.</small>
        </div>

        <div class="alan">
            <label for="admin-plan-hafta">Plan Haftası</label>
            <input type="text" id="admin-plan-hafta" placeholder="Örn. 34. hafta (17.08 haftası)">
        </div>

        <div class="iki-kolon">
            <div class="alan"><label for="admin-plan-baslama">Dizgi Başlama Tarihi</label><input type="date" id="admin-plan-baslama"></div>
            <div class="alan"><label for="admin-plan-teslim">Planlanan Teslim Tarihi</label><input type="date" id="admin-plan-teslim"></div>
        </div>

        <div class="alan">
            <label for="admin-gerceklesen-teslim">Gerçekleşen Teslim Tarihi</label>
            <input type="date" id="admin-gerceklesen-teslim">
            <small>TESLİM EDİLDİ durumunda boş bırakılırsa bugünün tarihi kullanılır.</small>
        </div>

        <div class="alan"><label for="admin-not">Not</label><textarea id="admin-not" rows="4"></textarea></div>

        <div class="modal-aksiyon">
            <button type="button" class="buton buton-hayalet" data-dialog-kapat>Vazgeç</button>
            <button type="submit" class="buton buton-ana">Kaydet</button>
        </div>
    </form>
</dialog>

<dialog id="admin-yeni-dialog" class="modal">
    <form method="dialog" class="modal-kutu" id="admin-yeni-form">
        <div class="modal-baslik">
            <div><p class="ust-etiket">KART YÖNETİMİ</p><h2>Yeni Kart</h2><p class="soluk">Yeni kart PLANA ALINDI durumunda oluşturulur.</p></div>
            <button class="ikon-buton" type="button" data-dialog-kapat aria-label="Kapat">×</button>
        </div>

        <div class="iki-kolon">
            <div class="alan"><label for="yeni-sira">NO / Sıra</label><input type="number" id="yeni-sira" step="1"></div>
            <div class="alan"><label for="yeni-talep-no">Talep NO *</label><input type="text" id="yeni-talep-no" required></div>
        </div>

        <div class="alan"><label for="yeni-talep-sahibi">Talep Sahibi</label><input type="text" id="yeni-talep-sahibi"></div>

        <div class="iki-kolon">
            <div class="alan"><label for="yeni-stok-no">Kart Stok No *</label><input type="text" id="yeni-stok-no" required></div>
            <div class="alan"><label for="yeni-toplam">Toplam Adet *</label><input type="number" id="yeni-toplam" min="1" required></div>
        </div>

        <div class="alan"><label for="yeni-plan-hafta">Plan Haftası</label><input type="text" id="yeni-plan-hafta" placeholder="Örn. 34. hafta (17.08 haftası)"></div>

        <div class="iki-kolon">
            <div class="alan"><label for="yeni-plan-baslama">Dizgi Başlama Tarihi</label><input type="date" id="yeni-plan-baslama"></div>
            <div class="alan"><label for="yeni-plan-teslim">Planlanan Teslim Tarihi</label><input type="date" id="yeni-plan-teslim"></div>
        </div>

        <div class="alan"><label for="yeni-pcb">PCB</label><input type="text" id="yeni-pcb" placeholder="Örn. HBT"></div>
        <div class="alan"><label for="yeni-not">Not</label><textarea id="yeni-not" rows="3"></textarea></div>

        <div class="modal-aksiyon">
            <button type="button" class="buton buton-hayalet" data-dialog-kapat>Vazgeç</button>
            <button type="submit" class="buton buton-ana">Kartı Oluştur</button>
        </div>
    </form>
</dialog>
{% endblock %}

{% block scripts %}
<script>
(() => {
    const adminAra = document.getElementById("admin-kart-ara");
    const adminSatirlar = [...document.querySelectorAll("[data-admin-kart]")];
    const adminFiltreler = [...document.querySelectorAll("[data-admin-filtre]")];
    const adminSonuc = document.getElementById("admin-sonuc");
    const adminBos = document.getElementById("admin-kart-bos");

    const adminDialog = document.getElementById("admin-dialog");
    const adminForm = document.getElementById("admin-form");
    const yeniDialog = document.getElementById("admin-yeni-dialog");
    const yeniForm = document.getElementById("admin-yeni-form");

    let adminAktifFiltre = "HEPSI";

    function adminDurumUygun(durum) {
        if (adminAktifFiltre === "HEPSI") {
            return true;
        }

        if (adminAktifFiltre === "AKTIF") {
            return ["PLANA ALINDI", "DİZGİDE", "HAZIR"].includes(durum);
        }

        return durum === adminAktifFiltre;
    }

    function adminKartlariFiltrele() {
        const arama = adminAra.value
            .trim()
            .toLocaleLowerCase("tr-TR");

        let gorunen = 0;

        adminSatirlar.forEach((satir) => {
            const durumUygun = adminDurumUygun(satir.dataset.durum);
            const aramaUygun = !arama || satir.dataset.arama.includes(arama);
            const uygun = durumUygun && aramaUygun;

            satir.hidden = !uygun;

            if (uygun) {
                gorunen += 1;
            }
        });

        adminSonuc.textContent =
            gorunen === adminSatirlar.length
                ? `${gorunen} kart`
                : `${gorunen} / ${adminSatirlar.length} kart`;

        adminBos.hidden = gorunen !== 0;
    }

    adminAra.addEventListener("input", adminKartlariFiltrele);

    adminFiltreler.forEach((buton) => {
        buton.addEventListener("click", () => {
            adminAktifFiltre = buton.dataset.adminFiltre;

            adminFiltreler.forEach((filtre) => {
                const aktif = filtre === buton;
                filtre.classList.toggle("aktif", aktif);
                filtre.setAttribute(
                    "aria-pressed",
                    aktif ? "true" : "false"
                );
            });

            adminKartlariFiltrele();
        });
    });

adminKartlariFiltrele();

    function duzenlemeDialogunuAc(buton) {
        document.getElementById("admin-id").value = buton.dataset.id;
        document.getElementById("admin-durum").value = buton.dataset.durum || "";
        document.getElementById("admin-toplam").value = buton.dataset.toplam;
        document.getElementById("admin-tamamlanan").value = buton.dataset.tamamlanan;
        document.getElementById("admin-plan-hafta").value = buton.dataset.planHafta || "";
        document.getElementById("admin-plan-baslama").value = buton.dataset.planBaslama || "";
        document.getElementById("admin-plan-teslim").value = buton.dataset.planTeslim || "";
        document.getElementById("admin-gerceklesen-teslim").value = buton.dataset.gerceklesenTeslim || "";
        document.getElementById("admin-not").value = buton.dataset.not || "";
        document.getElementById("admin-baslik").textContent = `${buton.dataset.talep || "Kart"} · ${buton.dataset.stok || ""}`;
        adminDialog.showModal();
    }

    document.querySelectorAll("[data-admin-duzenle]").forEach((buton) => {
        buton.addEventListener("click", () => duzenlemeDialogunuAc(buton));
    });

    document.querySelectorAll("[data-admin-gizle]").forEach((buton) => {
        buton.addEventListener("click", async () => {
            if (!confirm(`${buton.dataset.talep || "Bu kart"} listeden gizlensin mi?\n\nKart silinmez; kartlar.xlsx içinde tutulur.`)) return;

            buton.disabled = true;

            try {
                await pdgmFetch("/api/admin/kart-sil", {
                    method: "POST",
                    body: JSON.stringify({kart_id: Number(buton.dataset.id)})
                });
                toast("Kart listeden gizlendi.");
                window.setTimeout(() => window.location.reload(), 350);
            } catch (hata) {
                hataMesaji(hata);
                buton.disabled = false;
            }
        });
    });

    function durumAlanlariniAyarla() {
        const durum = document.getElementById("admin-durum").value;
        const toplam = Number(document.getElementById("admin-toplam").value || 0);
        const tamamlanan = document.getElementById("admin-tamamlanan");

        if (durum === "PLANA ALINDI") {
            tamamlanan.value = 0;
        } else if (["HAZIR", "TESLİM EDİLDİ"].includes(durum) && toplam > 0) {
            tamamlanan.value = toplam;
        }
    }

    document.getElementById("admin-durum").addEventListener("change", durumAlanlariniAyarla);
    document.getElementById("admin-toplam").addEventListener("input", durumAlanlariniAyarla);

    adminForm.addEventListener("submit", async (event) => {
        event.preventDefault();
        const submit = event.submitter;
        submit.disabled = true;

        try {
            await pdgmFetch("/api/admin/duzenle", {
                method: "POST",
                body: JSON.stringify({
                    kart_id: Number(document.getElementById("admin-id").value),
                    durum: document.getElementById("admin-durum").value,
                    toplam_adet: Number(document.getElementById("admin-toplam").value),
                    tamamlanan_adet: Number(document.getElementById("admin-tamamlanan").value),
                    plan_hafta: document.getElementById("admin-plan-hafta").value,
                    plan_baslama: document.getElementById("admin-plan-baslama").value,
                    plan_teslim: document.getElementById("admin-plan-teslim").value,
                    gerceklesen_teslim: document.getElementById("admin-gerceklesen-teslim").value,
                    not: document.getElementById("admin-not").value
                })
            });

            toast("Kart güncellendi.");
            dialogKapat(adminDialog);
            window.setTimeout(() => window.location.reload(), 350);
        } catch (hata) {
            hataMesaji(hata);
        } finally {
            submit.disabled = false;
        }
    });

    document.getElementById("admin-yeni-ac").addEventListener("click", () => {
        yeniForm.reset();
        yeniDialog.showModal();
        window.setTimeout(() => document.getElementById("yeni-talep-no").focus(), 50);
    });

    yeniForm.addEventListener("submit", async (event) => {
        event.preventDefault();
        const submit = event.submitter;
        submit.disabled = true;

        try {
            const sira = document.getElementById("yeni-sira").value;

            await pdgmFetch("/api/admin/kart-ekle", {
                method: "POST",
                body: JSON.stringify({
                    sira: sira ? Number(sira) : null,
                    talep_no: document.getElementById("yeni-talep-no").value,
                    talep_sahibi: document.getElementById("yeni-talep-sahibi").value,
                    stok_no: document.getElementById("yeni-stok-no").value,
                    toplam_adet: Number(document.getElementById("yeni-toplam").value),
                    plan_hafta: document.getElementById("yeni-plan-hafta").value,
                    plan_baslama: document.getElementById("yeni-plan-baslama").value,
                    plan_teslim: document.getElementById("yeni-plan-teslim").value,
                    pcb: document.getElementById("yeni-pcb").value,
                    not: document.getElementById("yeni-not").value
                })
            });

            toast("Yeni kart PLANA ALINDI durumunda oluşturuldu.");
            dialogKapat(yeniDialog);
            window.setTimeout(() => window.location.reload(), 350);
        } catch (hata) {
            hataMesaji(hata);
        } finally {
            submit.disabled = false;
        }
    });
})();
</script>
{% endblock %}
```


# 3. STATIC


## `static/stil.css`


```css
:root {
    --zemin: #f4f6f8;
    --kart: #ffffff;
    --yazi: #172027;
    --soluk: #66717a;
    --cizgi: #dfe5e8;
    --ana: #0f2027;
    --ana-2: #203a43;
    --iyi: #237a57;
    --iyi-bg: #e9f7f0;
    --uyari: #9a6500;
    --uyari-bg: #fff5dc;
    --kotu: #b33a3a;
    --kotu-bg: #fdecec;
    --notr: #64717a;
    --notr-bg: #eef2f4;
    --hazir: #216b78;
    --hazir-bg: #e8f5f7;
    --golge: 0 12px 34px rgba(15, 32, 39, .08);
    --radius: 16px;
}

* {
    box-sizing: border-box;
}

html {
    color-scheme: light;
}

body {
    position: relative;
    min-height: 100vh;
    margin: 0;
    background: linear-gradient(180deg, #f8fafb 0%, var(--zemin) 100%);
    color: var(--yazi);
    font-family: Arial, Helvetica, sans-serif;
}

button,
input,
select,
textarea {
    font: inherit;
}

button,
a {
    -webkit-tap-highlight-color: transparent;
}

a {
    color: inherit;
}

[hidden] {
    display: none !important;
}

.sr-only {
    position: absolute;
    width: 1px;
    height: 1px;
    padding: 0;
    margin: -1px;
    overflow: hidden;
    clip: rect(0, 0, 0, 0);
    white-space: nowrap;
    border: 0;
}

.ust-cubuk,
.sayfa {
    position: relative;
    z-index: 1;
}

/* --------------------------------------------------------------------------
   Header
---------------------------------------------------------------------------- */

.ust-cubuk {
    position: sticky;
    top: 0;
    z-index: 50;
    min-height: 72px;
    padding: 10px clamp(18px, 4vw, 56px);
    display: grid;
    grid-template-columns: minmax(240px, 1fr) auto minmax(220px, 1fr);
    align-items: center;
    gap: 24px;
    background: rgba(15, 32, 39, .97);
    color: #fff;
    box-shadow: 0 8px 30px rgba(0, 0, 0, .13);
}

.marka-link {
    display: inline-flex;
    align-items: center;
    gap: 12px;
    width: fit-content;
    transform: translateX(-35px);
    text-decoration: none;
}

.marka-isaret {
    width: 72px;
    height: 52px;
    padding: 3px;
    display: flex;
    align-items: center;
    justify-content: center;
    flex-shrink: 0;
}

.marka-logo {
    display: block;
    width: 100%;
    height: 100%;
    object-fit: contain;
}

.marka-metin {
    display: grid;
    gap: 2px;
}

.marka-metin small {
    color: #b9c7cd;
}

.ana-nav {
    display: flex;
    align-items: center;
    gap: 4px;
}

.ana-nav a {
    padding: 10px 13px;
    border-radius: 10px;
    color: #d8e1e5;
    text-decoration: none;
    font-size: 14px;
    font-weight: 700;
}

.ana-nav a:hover,
.ana-nav a.aktif {
    background: rgba(255, 255, 255, .10);
    color: #fff;
}

.oturum {
    display: flex;
    justify-content: flex-end;
    align-items: center;
    gap: 12px;
}

.oturum-metin {
    display: grid;
    gap: 2px;
    text-align: right;
    font-size: 13px;
}

.oturum-metin small {
    color: #aebdc3;
    font-size: 10px;
    letter-spacing: .7px;
}

/* --------------------------------------------------------------------------
   Genel layout
---------------------------------------------------------------------------- */

.sayfa {
    width: min(1520px, calc(100% - 36px));
    margin: 0 auto;
    padding: 30px 0 60px;
}

.sayfa-shell {
    display: grid;
    gap: 20px;
}

.sayfa-baslik {
    display: flex;
    justify-content: space-between;
    align-items: flex-start;
    gap: 24px;
    margin-bottom: 22px;
}

.sayfa-shell > .sayfa-baslik {
    margin-bottom: 0;
}

.sayfa-hero {
    align-items: center;
}

.sayfa-baslik h1,
.panel-baslik h2,
.operator-kart h2,
.ozet-donem-karti h2 {
    margin: 4px 0 6px;
}

.sayfa-baslik h1 {
    font-size: clamp(28px, 4vw, 42px);
    letter-spacing: -.8px;
}

.sayfa-hero .soluk {
    max-width: 760px;
    margin: 4px 0 0;
    line-height: 1.55;
}

.ust-etiket {
    margin: 0;
    color: #66757e;
    font-size: 11px;
    font-weight: 800;
    letter-spacing: 1.2px;
}

.soluk {
    color: var(--soluk);
}

.baslik-aksiyon,
.buton-grup,
.satir-aksiyon,
.kart-aksiyonlar {
    display: flex;
    align-items: center;
    gap: 9px;
    flex-wrap: wrap;
}

.buton-grup.dikey {
    display: grid;
}

.panel-kutu {
    margin: 0;
    padding: 20px;
    background: var(--kart);
    border: 1px solid var(--cizgi);
    border-radius: var(--radius);
    box-shadow: var(--golge);
}

.panel-baslik {
    display: flex;
    justify-content: space-between;
    align-items: center;
    gap: 16px;
    padding-bottom: 13px;
    margin-bottom: 15px;
    border-bottom: 1px solid #edf1f2;
}

.panel-baslik h2 {
    font-size: 20px;
}

.panel-aciklama {
    margin: 2px 0 0;
    color: var(--soluk);
    font-size: 11px;
    line-height: 1.5;
}

.sayi-rozet {
    min-width: 34px;
    height: 34px;
    padding: 0 8px;
    display: grid;
    place-items: center;
    border-radius: 999px;
    background: var(--notr-bg);
    font-weight: 800;
}

/* --------------------------------------------------------------------------
   Butonlar ve formlar
---------------------------------------------------------------------------- */

.buton {
    min-height: 40px;
    padding: 10px 15px;
    display: inline-flex;
    align-items: center;
    justify-content: center;
    gap: 6px;
    border: 1px solid transparent;
    border-radius: 10px;
    cursor: pointer;
    text-decoration: none;
    font-weight: 700;
    transition: transform .12s ease, opacity .12s ease, background .12s ease;
}

.buton:hover {
    transform: translateY(-1px);
}

.buton:disabled {
    cursor: wait;
    opacity: .55;
    transform: none;
}

.buton-ana {
    background: var(--ana);
    color: #fff;
}

.buton-basari {
    background: var(--iyi);
    color: #fff;
}

.buton-uyari {
    background: var(--uyari-bg);
    color: #785000;
    border-color: #e8c777;
}

.buton-tehlike {
    background: var(--kotu-bg);
    color: var(--kotu);
    border-color: #efb7b7;
}

.buton-hayalet {
    background: #fff;
    color: var(--yazi);
    border-color: var(--cizgi);
}

.ust-cubuk .buton-hayalet {
    background: rgba(255, 255, 255, .08);
    color: #fff;
    border-color: rgba(255, 255, 255, .14);
}

.buton-kucuk {
    min-height: 34px;
    padding: 7px 10px;
    font-size: 12px;
}

.tam-genislik {
    width: 100%;
}

.alan,
.arama {
    display: grid;
    gap: 7px;
}

.alan label,
.arama label {
    font-size: 12px;
    font-weight: 700;
}

.alan small,
.dosya-sec small {
    color: var(--soluk);
    font-size: 11px;
}

input,
select,
textarea {
    width: 100%;
    min-height: 42px;
    padding: 9px 11px;
    border: 1px solid #cfd8dc;
    border-radius: 9px;
    background: #fff;
    color: var(--yazi);
    outline: 0;
}

textarea {
    resize: vertical;
}

input:focus,
select:focus,
textarea:focus,
button:focus-visible,
a:focus-visible {
    outline: none;
    border-color: #5c7885;
    box-shadow: 0 0 0 3px rgba(32, 58, 67, .16);
}

.arac-cubugu {
    display: flex;
    justify-content: space-between;
    align-items: end;
    gap: 18px;
}

.arama {
    min-width: min(360px, 100%);
}

.arama.kucuk {
    min-width: 250px;
}

.filtreler {
    display: flex;
    gap: 7px;
    flex-wrap: wrap;
}

.filtre {
    min-height: 40px;
    padding: 8px 12px;
    display: inline-flex;
    align-items: center;
    border: 1px solid var(--cizgi);
    border-radius: 999px;
    background: #fff;
    color: #59666d;
    cursor: pointer;
    font-weight: 700;
    font-size: 12px;
    transition: transform .14s ease, background .14s ease, border-color .14s ease;
}

.filtre:hover {
    transform: translateY(-1px);
}

.filtre.aktif {
    background: var(--ana);
    color: #fff;
    border-color: var(--ana);
}

.metin-buton,
.operator-temizle {
    padding: 0;
    border: 0;
    background: transparent;
    color: var(--ana-2);
    cursor: pointer;
    font-size: 11px;
    font-weight: 800;
}

.metin-buton:hover,
.operator-temizle:hover {
    text-decoration: underline;
}

/* --------------------------------------------------------------------------
   Durum renkleri ve KPI
---------------------------------------------------------------------------- */

.istatistik-grid {
    display: grid;
    grid-template-columns: repeat(4, minmax(0, 1fr));
    gap: 14px;
}

.istatistik {
    min-height: 108px;
    padding: 18px;
    display: grid;
    gap: 8px;
    background: var(--kart);
    border: 1px solid var(--cizgi);
    border-top: 4px solid var(--notr);
    border-radius: 14px;
    box-shadow: var(--golge);
    transition: transform .15s ease, box-shadow .15s ease;
}

.istatistik:hover {
    transform: translateY(-2px);
    box-shadow: 0 15px 34px rgba(15, 32, 39, .10);
}

.istatistik span {
    color: var(--soluk);
    font-size: 13px;
    font-weight: 700;
}

.istatistik strong {
    font-size: 32px;
    line-height: 1;
}

.kart-plana {
    border-top-color: #788890;
}

.kart-dizgide {
    border-top-color: #d18a00;
}

.kart-hazir {
    border-top-color: var(--hazir);
}

.kart-teslim {
    border-top-color: var(--iyi);
}

.durum-rozet {
    width: fit-content;
    max-width: 100%;
    padding: 5px 9px;
    display: inline-flex;
    align-items: center;
    border-radius: 999px;
    background: var(--notr-bg);
    color: var(--notr);
    font-size: 10px;
    line-height: 1.2;
    font-weight: 800;
    white-space: nowrap;
}

.durum-rozet.iyi {
    background: var(--iyi-bg);
    color: var(--iyi);
}

.durum-rozet.uyari {
    background: var(--uyari-bg);
    color: var(--uyari);
}

.durum-rozet.kotu {
    background: var(--kotu-bg);
    color: var(--kotu);
}

.satir-alt-rozet {
    margin-top: 5px;
}

/* --------------------------------------------------------------------------
   Kartlar ve listeler
---------------------------------------------------------------------------- */

.pano-grid,
.yonetim-grid {
    display: grid;
    grid-template-columns: 1.3fr .9fr;
    gap: 20px;
}

.kart-listesi,
.mini-liste,
.log-liste {
    display: grid;
    gap: 10px;
}

.is-karti,
.operator-kart {
    background: #fff;
    border: 1px solid var(--cizgi);
    border-left: 5px solid var(--notr);
    border-radius: 13px;
}

.is-karti {
    padding: 15px;
}

.is-karti.iyi,
.operator-kart.iyi {
    border-left-color: var(--iyi);
}

.is-karti.uyari,
.operator-kart.uyari {
    border-left-color: var(--uyari);
}

.is-karti.kotu,
.operator-kart.kotu {
    border-left-color: var(--kotu);
}

.is-karti-ust,
.operator-kart-ust {
    display: flex;
    justify-content: space-between;
    align-items: flex-start;
    gap: 16px;
}

.is-karti-ust > div {
    display: grid;
    gap: 4px;
}

.is-karti-ust span:not(.durum-rozet) {
    color: var(--soluk);
}

.kart-bilgiler,
.kart-ek-bilgi {
    display: flex;
    flex-wrap: wrap;
    gap: 7px 15px;
    margin-top: 12px;
    color: var(--soluk);
    font-size: 12px;
}

.operator-rozetler {
    display: grid;
    gap: 6px;
    justify-items: end;
}

.durum-satiri {
    display: flex;
    flex-wrap: wrap;
    gap: 10px 18px;
    margin: 12px 0 4px;
    color: #44525a;
    font-size: 13px;
}

.ilerleme {
    margin-top: 14px;
}

.ilerleme-ust {
    display: flex;
    justify-content: space-between;
    gap: 10px;
    margin-bottom: 7px;
    font-size: 12px;
}

.ilerleme-ray {
    height: 8px;
    overflow: hidden;
    border-radius: 999px;
    background: #e8edef;
}

.ilerleme-dolgu {
    height: 100%;
    border-radius: inherit;
    background: linear-gradient(90deg, var(--ana-2), #2c5364);
}

.kart-not {
    margin: 12px 0 0;
    padding: 9px 11px;
    border-left: 3px solid #d7e0e4;
    border-radius: 9px;
    background: #f7f9fa;
    color: #4f5b62;
    font-size: 12px;
}

.mini-satir {
    padding: 12px 0;
    display: flex;
    justify-content: space-between;
    gap: 16px;
    border-bottom: 1px solid #edf0f2;
}

.mini-satir:last-child {
    border-bottom: 0;
}

.mini-satir > div:first-child,
.mini-sag {
    display: grid;
    gap: 4px;
}

.mini-satir span,
.mini-satir small {
    color: var(--soluk);
    font-size: 11px;
}

.mini-sag {
    justify-items: end;
    text-align: right;
}

.bos-durum {
    padding: 28px 16px;
    text-align: center;
    color: var(--soluk);
    border: 1px dashed #ccd5da;
    border-radius: 12px;
    background: rgba(255, 255, 255, .72);
}

.tam-satir {
    grid-column: 1 / -1;
}

/* --------------------------------------------------------------------------
   Operator
---------------------------------------------------------------------------- */

.operator-arac-cubugu,
.panel-arac-cubugu {
    padding: 17px 18px;
    background: linear-gradient(135deg, rgba(255,255,255,.99), rgba(247,250,251,.96));
}

.operator-sayfa .arama,
.panel-arama-alani {
    flex: 1 1 340px;
    max-width: 540px;
}

.operator-filtre-alani,
.panel-filtre-alani {
    display: grid;
    gap: 7px;
    justify-items: end;
}

.operator-filtre-meta,
.panel-filtre-alt {
    min-height: 20px;
    display: flex;
    justify-content: flex-end;
    align-items: center;
    gap: 10px;
    color: var(--soluk);
    font-size: 11px;
}

.operator-grid {
    display: grid;
    grid-template-columns: repeat(3, minmax(0, 1fr));
    gap: 16px;
    align-items: start;
}

.operator-kart {
    padding: 18px;
    box-shadow: var(--golge);
    transition: transform .16s ease, box-shadow .16s ease;
}

.operator-kart:hover {
    transform: translateY(-2px);
    box-shadow: 0 15px 35px rgba(15, 32, 39, .10);
}

.operator-kart-ust {
    padding-bottom: 13px;
    border-bottom: 1px solid #edf1f2;
}

.operator-kart-ust h2 {
    margin-top: 5px;
    margin-bottom: 4px;
    font-size: 19px;
}

.operator-kart-ust p {
    margin: 0;
    color: var(--soluk);
    font-size: 12px;
}

.bilgi-grid {
    display: grid;
    grid-template-columns: repeat(4, 1fr);
    gap: 8px;
    margin-top: 16px;
}

.bilgi-grid > div {
    padding: 10px;
    display: grid;
    gap: 4px;
    border: 1px solid #edf1f2;
    border-radius: 10px;
    background: #fafbfc;
}

.bilgi-grid span {
    color: var(--soluk);
    font-size: 10px;
    font-weight: 700;
}

.bilgi-grid strong {
    font-size: 14px;
}

.kart-aksiyonlar {
    margin-top: 15px;
    padding-top: 14px;
    border-top: 1px solid #edf1f2;
}

.kart-aksiyonlar .buton {
    min-height: 44px;
    flex: 1 1 120px;
    padding: 12px 16px;
}

/* --------------------------------------------------------------------------
   Panel
---------------------------------------------------------------------------- */

.panel-canli-satir {
    display: flex;
    align-items: center;
    flex-wrap: wrap;
    gap: 8px 12px;
    margin-top: 5px;
}

.panel-canli {
    display: inline-flex;
    align-items: center;
    gap: 7px;
    color: #526169;
    font-size: 12px;
    font-weight: 700;
}

.panel-canli::before {
    content: "";
    width: 8px;
    height: 8px;
    border-radius: 50%;
    background: var(--iyi);
    box-shadow: 0 0 0 4px rgba(35, 122, 87, .11);
}

.panel-guncelleme {
    color: var(--soluk);
    font-size: 12px;
}

.panel-aktif-grid {
    align-items: start;
}

.panel-is-karti {
    transition: transform .14s ease, box-shadow .14s ease;
}

.panel-is-karti:hover {
    transform: translateY(-1px);
    box-shadow: 0 9px 22px rgba(15, 32, 39, .07);
}

.panel-hazir-kutu {
    background: linear-gradient(180deg, #fff, #fbfefe);
}

.panel-mini-satir {
    margin: 0 -4px;
    padding: 12px 6px;
    border-radius: 9px;
}

.panel-mini-satir:hover {
    background: #f7f9fa;
}

.plan-liste {
    display: grid;
}

.plan-satir {
    display: grid;
    grid-template-columns: 44px minmax(180px, 1.1fr) minmax(160px, 1fr) 100px 120px minmax(150px, auto);
    align-items: center;
    gap: 14px;
    padding: 12px 4px;
    border-bottom: 1px solid #edf0f2;
    font-size: 12px;
}

.plan-satir:last-child {
    border-bottom: 0;
}

.plan-satir > div:nth-child(2) {
    display: grid;
    gap: 3px;
}

.plan-satir > div:nth-child(2) span,
.plan-satir > span:not(.durum-rozet) {
    color: var(--soluk);
}

.plan-sira {
    width: 32px;
    height: 32px;
    display: grid;
    place-items: center;
    border-radius: 9px;
    background: var(--notr-bg);
    font-weight: 800;
}

/* --------------------------------------------------------------------------
   Monitör
---------------------------------------------------------------------------- */
.monitor-body {
    height: 100vh;
    overflow: hidden;
}

.monitor-body .sayfa {
    width: min(1900px, calc(100% - 28px));
    height: calc(100vh - 72px);
    min-height: 0;
    padding-top: 12px;
    padding-bottom: 12px;
    overflow: hidden;
}


.monitor-sayfa {
    height: 100%;
    min-height: 0;
    display: grid;
    grid-template-rows: auto minmax(0, 1fr);
    gap: 10px;
}

.monitor-ust {
    min-height: 58px;
    padding: 0 3px;
    display: flex;
    justify-content: space-between;
    align-items: center;
    gap: 24px;
}

.monitor-ust h1 {
    margin: 2px 0;
    font-size: clamp(26px, 2.2vw, 36px);
}
.monitor-aciklama {
    margin: 0;
    color: var(--soluk);
    font-size: 12px;
}

.monitor-zaman {
    display: grid;
    justify-items: end;
    gap: 2px;
    text-align: right;
}

.monitor-zaman span {
    color: var(--ana);
    font-size: clamp(30px, 2.7vw, 42px);
    line-height: 1;
    font-weight: 800;
    letter-spacing: -1px;
}

.monitor-zaman small {
    color: var(--soluk);
    font-size: 10px;
}

.monitor-grid {
    min-height: 0;
    height: 100%;
    display: grid;
    grid-template-columns: minmax(0, 1fr) minmax(0, 1.15fr);
    gap: 14px;
    align-items: stretch;
}

.monitor-bolum {
    min-width: 0;
    min-height: 0;
    height: 100%;
    padding: 12px;
    display: grid;
    grid-template-rows: auto minmax(0, 1fr);
    align-content: stretch;
    overflow: hidden;
    border-top-width: 4px;
    box-shadow: 0 8px 22px rgba(15, 32, 39, .06);
}

.monitor-dizgide {
    border-top-color: #d18a00;
}

.monitor-hazir {
    border-top-color: var(--hazir);
}

.monitor-plana {
    border-top-color: #788890;
}

.monitor-bolum-baslik {
    margin-bottom: 9px;
    padding-bottom: 10px;
}

.monitor-bolum-baslik h2 {
    margin: 2px 0 0;
    font-size: clamp(20px, 1.5vw, 27px);
    letter-spacing: -.4px;
}

.monitor-bolum-baslik .sayi-rozet {
    min-width: 34px;
    height: 34px;
    font-size: 14px;
}

/*
   İç scroll kaldırıldı.
   Kartlar section içinde doğal olarak yerleşir.
*/
.monitor-liste {
    min-height: 0;
    height: 100%;
    padding: 0;
    display: grid;
    align-content: start;
    gap: 8px;
    overflow: hidden;
}

.monitor-dizgide-liste {
    min-height: 0;
    height: 100%;
    display: grid;
    gap: 8px;
    align-content: start;
}

/* Tam 3 kart varsa mevcut alanı üç eşit parçaya böl */
.monitor-dizgide-liste.uc-dizgi {
    grid-template-rows: repeat(3, minmax(0, 1fr));
    align-content: stretch;
}

.monitor-dizgide-liste.uc-dizgi .monitor-kart {
    min-height: 0;
    height: 100%;
    padding: 10px 13px;
    display: grid;
    align-content: space-between;
    overflow: hidden;
}

.monitor-dizgide-liste.uc-dizgi .monitor-kart-ust {
    gap: 6px;
}

.monitor-dizgide-liste.uc-dizgi .monitor-talep strong {
    font-size: clamp(19px, 1.25vw, 24px);
}

.monitor-dizgide-liste.uc-dizgi .monitor-kart h3 {
    margin: 6px 0 2px;
    font-size: clamp(17px, 1.1vw, 22px);
}

.monitor-dizgide-liste.uc-dizgi .monitor-sahip {
    font-size: 11px;
}

.monitor-dizgide-liste.uc-dizgi .monitor-adet {
    margin-top: 5px;
}

.monitor-dizgide-liste.uc-dizgi .monitor-adet strong {
    font-size: clamp(19px, 1.25vw, 24px);
}

.monitor-dizgide-liste.uc-dizgi .monitor-adet span {
    font-size: 9px;
}

.monitor-dizgide-liste.uc-dizgi .monitor-ilerleme {
    margin-top: 5px;
}

.monitor-dizgide-liste.uc-dizgi .monitor-ilerleme .ilerleme-ust {
    margin-bottom: 3px;
    font-size: 9px;
}

.monitor-dizgide-liste.uc-dizgi .monitor-ilerleme .ilerleme-ray {
    height: 6px;
}

.monitor-dizgide-liste.uc-dizgi .monitor-meta {
    margin-top: 5px;
    padding-top: 5px;
}

.monitor-dizgide-liste.uc-dizgi .monitor-meta strong {
    font-size: 12px;
}



/* --------------------------------------------------------------------------
   DİZGİDE / HAZIR
---------------------------------------------------------------------------- */

.monitor-kart {
    padding: 15px 16px;
    background: #fff;
    border: 1px solid #dde4e7;
    border-left: 4px solid #d18a00;
    border-radius: 11px;
}

.monitor-hazir .monitor-kart {
    border-left-color: var(--hazir);
}

.monitor-kart.kotu,
.monitor-plan-kart.kotu {
    border-left-color: var(--kotu);
}

.monitor-kart-ust {
    display: flex;
    justify-content: space-between;
    align-items: flex-start;
    gap: 9px;
}

.monitor-talep {
    display: grid;
    gap: 1px;
}

.monitor-talep span,
.monitor-plan-ust > div span {
    color: var(--soluk);
    font-size: 9px;
    font-weight: 800;
    letter-spacing: .9px;
}

.monitor-talep strong {
    color: var(--ana);
    font-size: clamp(21px, 1.45vw, 28px);
    line-height: 1;
}

.monitor-kart .durum-rozet {
    padding: 4px 7px;
    font-size: 9px;
}

.monitor-kart h3 {
    margin: 11px 0 3px;
    overflow-wrap: anywhere;
    font-size: clamp(19px, 1.35vw, 26px);
    line-height: 1.08;
}

.monitor-sahip {
    margin: 0;
    color: var(--soluk);
    font-size: 12px;
    font-weight: 700;
}

.monitor-adet {
    margin-top: 11px;
    display: flex;
    align-items: baseline;
    gap: 6px;
}

.monitor-adet strong {
    font-size: clamp(22px, 1.5vw, 28px);
}

.monitor-adet span {
    color: var(--soluk);
    font-size: 10px;
    font-weight: 700;
}

.monitor-ilerleme {
    margin-top: 9px;
}

.monitor-ilerleme .ilerleme-ust {
    margin-bottom: 5px;
    font-size: 10px;
}

.monitor-ilerleme .ilerleme-ray {
    height: 8px;
}

.monitor-meta {
    margin-top: 10px;
    padding-top: 9px;
    display: flex;
    gap: 15px;
    border-top: 1px solid #edf1f2;
}

.monitor-meta span,
.monitor-hazir-bilgi span,
.monitor-plan-meta span {
    display: grid;
    gap: 2px;
}

.monitor-meta small,
.monitor-hazir-bilgi small,
.monitor-plan-meta small {
    color: var(--soluk);
    font-size: 8px;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: .45px;
}

.monitor-meta strong {
    font-size: 14px;
}

/* --------------------------------------------------------------------------
   HAZIR
---------------------------------------------------------------------------- */

.monitor-hazir-mesaj {
    margin-top: 11px;
    padding: 8px 10px;
    border: 1px solid #c9e0e4;
    border-radius: 8px;
    background: var(--hazir-bg);
    color: #175965;
    font-size: 10px;
    font-weight: 800;
    letter-spacing: .25px;
}

.monitor-hazir-bilgi {
    margin-top: 9px;
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 7px;
}

.monitor-hazir-bilgi span {
    padding: 8px 9px;
    border: 1px solid #edf1f2;
    border-radius: 8px;
    background: #fafcfc;
}

.monitor-hazir-bilgi strong {
    font-size: 14px;
}

/* --------------------------------------------------------------------------
   PLANA ALINDI
---------------------------------------------------------------------------- */

/*
   Plan kuyruğu özellikle daha yoğun.
   Büyük ekranda iki mini kart yan yana.
*/
.monitor-plan-liste {
    grid-template-columns: repeat(2, minmax(0, 1fr));
    grid-auto-rows: max-content;
    gap: 12px;
}
.monitor-plan-kart {
    min-width: 0;
    padding: 14px 15px;
    display: grid;
    grid-template-columns: 38px minmax(0, 1fr);
    gap: 12px;
    background: #fff;
    border: 1px solid #dfe5e8;
    border-left: 4px solid #788890;
    border-radius: 11px;
}

.monitor-plan-sira {
    width: 36px;
    height: 36px;
    display: grid;
    place-items: center;
    border-radius: 9px;
    background: var(--notr-bg);
    color: #4e5b62;
    font-size: 13px;
    font-weight: 800;
}

.monitor-plan-ust {
    display: flex;
    justify-content: space-between;
    align-items: flex-start;
    gap: 8px;
}

.monitor-plan-ust > div {
    min-width: 0;
    display: grid;
    gap: 2px;
}

.monitor-plan-ust > div span {
    font-size: 10px;
}

.monitor-plan-ust strong {
    color: var(--ana);
    font-size: 19px;
    line-height: 1.05;
}

.monitor-plan-kart .durum-rozet {
    padding: 4px 8px;
    font-size: 9px;
}

.monitor-plan-kart h3 {
    margin: 9px 0 4px;
    overflow-wrap: anywhere;
    font-size: 17px;
    line-height: 1.08;
}

.monitor-plan-kart p {
    margin: 0;
    overflow: hidden;
    color: var(--soluk);
    font-size: 12px;
    font-weight: 700;
    line-height: 1.25;
    text-overflow: ellipsis;
    white-space: nowrap;
}

.monitor-plan-meta {
    margin-top: 10px;
    padding-top: 9px;
    display: grid;
    grid-template-columns: .6fr 1fr 1fr;
    gap: 8px;
    border-top: 1px solid #edf1f2;
}

.monitor-plan-meta small {
    font-size: 9px;
}

.monitor-plan-meta strong {
    font-size: 12px;
    line-height: 1.2;
}

/* --------------------------------------------------------------------------
   Empty state
---------------------------------------------------------------------------- */

.monitor-bos {
    min-height: 110px;
    display: grid;
    place-items: center;
    align-self: start;
    padding: 18px;
    font-size: 12px;
    line-height: 1.45;
}

/* --------------------------------------------------------------------------
   Monitör responsive
---------------------------------------------------------------------------- */

@media (max-width: 1050px) {
    .monitor-grid {
        grid-template-columns: 1fr;
    }

    .monitor-plan-liste {
        grid-template-columns: repeat(3, minmax(0, 1fr));
    }
}


@media (max-width: 820px) {
    .monitor-grid {
        grid-template-columns: 1fr;
    }

    .monitor-plan-liste {
        grid-template-columns: repeat(2, minmax(0, 1fr));
    }

    .monitor-ust {
        align-items: flex-start;
    }

    .monitor-zaman {
        flex-shrink: 0;
    }
}

@media (max-width: 520px) {
    .monitor-body .sayfa {
        width: min(100% - 20px, 1520px);
        padding-top: 14px;
    }

    .monitor-ust {
        display: grid;
        gap: 10px;
    }

    .monitor-zaman {
        justify-items: start;
        text-align: left;
    }

    .monitor-zaman span {
        font-size: 30px;
    }

    .monitor-bolum {
        padding: 12px;
    }

    .monitor-plan-liste {
        grid-template-columns: 1fr;
    }

    .monitor-hazir-bilgi {
        grid-template-columns: 1fr 1fr;
    }
}

/* --------------------------------------------------------------------------
   Yönetim
---------------------------------------------------------------------------- */

.yonetim-yukleme-karti {
    background: linear-gradient(145deg, #fff, #f9fbfc);
}

.yukleme-form {
    display: grid;
    gap: 12px;
}

.dosya-sec {
    min-height: 115px;
    padding: 16px;
    display: grid;
    place-content: center;
    gap: 9px;
    border: 2px dashed #aab9c0;
    border-radius: 12px;
    background: #f8fafb;
    cursor: pointer;
    transition: border-color .15s ease, background .15s ease;
}

.dosya-sec:hover {
    border-color: #708890;
    background: #f3f7f8;
}

.dosya-sec span {
    font-size: 14px;
    font-weight: 700;
}

.dosya-sec input {
    padding: 7px;
    background: #fff;
}

.yardim {
    margin-bottom: 0;
    color: var(--soluk);
    font-size: 11px;
    line-height: 1.65;
}

.yardim code {
    padding: 1px 6px;
    border-radius: 6px;
    background: var(--notr-bg);
    font-size: 12px;
}

.yonetim-durum-eksik {
    border-color: #e9ca7c;
    background: linear-gradient(135deg, #fffdf6, #fff9e9);
}

.yonetim-durum-eksik-liste {
    max-height: 320px;
    padding-right: 7px;
    overflow-y: auto;
    overscroll-behavior: contain;
}

.yonetim-durum-eksik-liste .mini-satir:first-child {
    padding-top: 4px;
}

.yonetim-durum-eksik-liste .mini-satir:last-child {
    padding-bottom: 4px;
}

.yonetim-uyari-kutu {
    border-color: #ecd497;
    background: linear-gradient(135deg, #fffdf6, #fffaf0);
}

.yonetim-kart-panel .panel-baslik {
    align-items: end;
}

.yonetim-gecmis .mini-satir,
.yonetim-gecmis .log-satir {
    padding-left: 6px;
    padding-right: 6px;
    border-radius: 8px;
}

.yonetim-gecmis .mini-satir:hover,
.yonetim-gecmis .log-satir:hover {
    background: #f7f9fa;
}

.log-satir {
    padding: 10px 0;
    display: grid;
    gap: 4px;
    border-bottom: 1px solid #edf0f2;
}

.log-satir:last-child {
    border-bottom: 0;
}

.log-satir > div {
    display: flex;
    justify-content: space-between;
    gap: 12px;
}

.log-satir span,
.log-satir small {
    color: var(--soluk);
    font-size: 11px;
}

.admin-filtre-cubugu {
    margin: -2px 0 13px;
    display: flex;
    justify-content: space-between;
    align-items: center;
    gap: 14px;
}

.admin-filtre-sonuc {
    flex-shrink: 0;
    color: var(--soluk);
    font-size: 11px;
    font-weight: 700;
}

.yonetim-kart-scroll {
    max-height: 470px;
    overflow: auto;
    overscroll-behavior: contain;
    scrollbar-width: thin;
    scrollbar-color: #b8c3c8 transparent;
}

/*
   tablo-kapsayici overflow oluşturduğu için mevcut sticky th
   burada header'ı sabit tutar.
*/
.yonetim-kart-scroll th {
    top: 0;
    z-index: 2;
}

.yonetim-kart-scroll::-webkit-scrollbar {
    width: 8px;
    height: 8px;
}

.yonetim-kart-scroll::-webkit-scrollbar-track {
    background: transparent;
}

.yonetim-kart-scroll::-webkit-scrollbar-thumb {
    border-radius: 999px;
    background: #b8c3c8;
}

.yonetim-kart-scroll::-webkit-scrollbar-thumb:hover {
    background: #929fa5;
}

.admin-kart-bos {
    margin-top: 12px;
}

/* --------------------------------------------------------------------------
   Özet ve grafik
---------------------------------------------------------------------------- */

.ozet-uyari-satiri {
    display: flex;
    gap: 8px;
    flex-wrap: wrap;
}

.donem-grid {
    display: grid;
    grid-template-columns: repeat(3, 1fr);
    gap: 16px;
}

.ozet-donem-karti {
    transition: transform .15s ease, box-shadow .15s ease;
}

.ozet-donem-karti:hover {
    transform: translateY(-2px);
}

.ozet-donem-karti h2 {
    margin-bottom: 17px;
}

.donem-metrik {
    min-height: 42px;
    padding: 10px 0;
    display: flex;
    justify-content: space-between;
    align-items: center;
    gap: 12px;
    border-top: 1px solid #edf0f2;
    font-size: 12px;
}

.donem-metrik span {
    color: var(--soluk);
}

.donem-metrik strong {
    color: var(--ana);
    font-size: 14px;
}

.grafik-aciklama {
    display: flex;
    gap: 18px;
    margin-bottom: 14px;
    color: var(--soluk);
    font-size: 11px;
}

.grafik-aciklama span {
    display: flex;
    align-items: center;
    gap: 6px;
}

.lejant {
    width: 10px;
    height: 10px;
    display: inline-block;
    border-radius: 3px;
}

.lejant.plan,
.sutun.plan {
    background: #aab7bd;
}

.lejant.teslim,
.sutun.teslim {
    background: #264653;
}

.hafta-grafik {
    min-height: 310px;
    padding: 20px 8px 0;
    display: grid;
    grid-template-columns: repeat(8, minmax(62px, 1fr));
    align-items: end;
    gap: 12px;
    overflow-x: auto;
}

.hafta-sutun {
    height: 270px;
    display: grid;
    grid-template-rows: 1fr auto auto;
    gap: 5px;
    text-align: center;
    font-size: 11px;
}

.sutun-alani {
    min-height: 220px;
    display: flex;
    align-items: end;
    justify-content: center;
    gap: 4px;
    border-bottom: 1px solid #ccd5da;
}

.sutun {
    position: relative;
    width: min(28px, 42%);
    min-height: 2px;
    border-radius: 6px 6px 0 0;
}

.sutun span {
    position: absolute;
    top: -18px;
    left: 50%;
    transform: translateX(-50%);
    font-size: 10px;
    font-weight: 700;
}

.hafta-sutun small {
    color: var(--soluk);
}

/* --------------------------------------------------------------------------
   Tablolar
---------------------------------------------------------------------------- */

.tablo-kapsayici {
    max-width: 100%;
    overflow: auto;
    border: 1px solid #edf1f2;
    border-radius: 11px;
}

table {
    width: 100%;
    min-width: 760px;
    border-collapse: collapse;
}

th,
td {
    padding: 11px 12px;
    text-align: left;
    vertical-align: middle;
    border-bottom: 1px solid #e9edef;
    font-size: 12px;
}

th {
    position: sticky;
    top: 0;
    z-index: 1;
    background: #f6f8f9;
    color: #59666d;
    font-size: 10px;
    text-transform: uppercase;
    letter-spacing: .5px;
}

tbody tr:hover {
    background: #fafbfb;
}

/* --------------------------------------------------------------------------
   Modal
---------------------------------------------------------------------------- */

.modal {
    width: min(640px, calc(100% - 24px));
    max-width: none;
    max-height: calc(100vh - 32px);
    padding: 0;
    overflow: visible;
    border: 0;
    background: transparent;
}

.modal[open] {
    display: block;
}

.modal::backdrop {
    background: rgba(9, 19, 24, .62);
    backdrop-filter: blur(2px);
}

.modal-kutu {
    width: 100%;
    max-height: calc(100vh - 40px);
    padding: 20px;
    overflow-y: auto;
    border-radius: 16px;
    background: #fff;
    box-shadow: 0 28px 70px rgba(0, 0, 0, .24);
}

.modal-baslik {
    display: flex;
    justify-content: space-between;
    gap: 15px;
    margin-bottom: 18px;
}

.modal-baslik h2 {
    margin: 4px 0 0;
}

.ikon-buton {
    width: 36px;
    height: 36px;
    border: 1px solid var(--cizgi);
    border-radius: 9px;
    background: #fff;
    cursor: pointer;
    font-size: 22px;
}

.modal-aksiyon {
    display: flex;
    justify-content: flex-end;
    gap: 8px;
    margin-top: 18px;
}

.iki-kolon {
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 12px;
}

.modal .alan + .alan,
.modal .iki-kolon + .alan,
.modal .alan + .iki-kolon,
.modal .iki-kolon + .iki-kolon {
    margin-top: 12px;
}

/* --------------------------------------------------------------------------
   Bildirim / toast
---------------------------------------------------------------------------- */

.bildirimler {
    margin-bottom: 16px;
}

.bildirim {
    padding: 11px 13px;
    border: 1px solid transparent;
    border-radius: 10px;
    font-size: 13px;
}

.bildirim.basari {
    background: var(--iyi-bg);
    border-color: #b7dfca;
    color: #175d41;
}

.bildirim.hata {
    background: var(--kotu-bg);
    border-color: #efb9b9;
    color: #922f2f;
}

.bildirim.uyari {
    background: var(--uyari-bg);
    border-color: #ecd08a;
    color: #775200;
}

.toast-alani {
    position: fixed;
    right: 20px;
    bottom: 20px;
    z-index: 200;
    display: grid;
    gap: 8px;
    pointer-events: none;
}

.toast {
    max-width: 380px;
    padding: 12px 14px;
    border-radius: 11px;
    background: var(--ana);
    color: #fff;
    box-shadow: 0 14px 40px rgba(0, 0, 0, .16);
    opacity: 0;
    transform: translateY(8px);
    transition: .18s ease;
    font-size: 13px;
}

.toast.goster {
    opacity: 1;
    transform: translateY(0);
}

.toast.hata {
    background: var(--kotu);
}

.toast.uyari {
    background: #8b5f06;
}

/* --------------------------------------------------------------------------
   Login / 403
---------------------------------------------------------------------------- */

.giris-sayfa {
    position: relative;
    min-height: 100vh;
    padding: 24px;
    display: grid;
    place-items: center;
    overflow: hidden;
    background:
        radial-gradient(circle at 20% 20%, rgba(44, 83, 100, .18), transparent 30%),
        linear-gradient(145deg, #0f2027, #203a43 50%, #2c5364);
}

.giris-sayfa::before {
    content: "";
    position: fixed;
    top: 50%;
    left: 50%;
    width: min(900px, 78vw);
    height: min(900px, 78vw);
    transform: translate(-50%, -50%) rotate(-6deg);
    background: url("/static/pdgm_logo.png") center / contain no-repeat;
    opacity: .025;
    filter: grayscale(100%);
    pointer-events: none;
}

.giris-kutu {
    position: relative;
    z-index: 1;
    width: min(420px, 100%);
    padding: 34px;
    border: 1px solid rgba(255,255,255,.55);
    border-radius: 20px;
    background: rgba(255, 255, 255, .98);
    box-shadow: 0 30px 80px rgba(0, 0, 0, .28);
    backdrop-filter: blur(4px);
}

.giris-logo {
    width: 170px;
    height: 92px;
    margin: 0 auto 16px;
    display: flex;
    align-items: center;
    justify-content: center;
}

.giris-logo img {
    display: block;
    width: 100%;
    height: 100%;
    object-fit: contain;
}

.giris-kutu h1 {
    margin: 0;
    text-align: center;
    letter-spacing: -.5px;
    font-size: 27px;
}

.giris-kutu .alt {
    margin: 7px 0 24px;
    text-align: center;
    color: var(--soluk);
}

.giris-form {
    display: grid;
    gap: 15px;
}

.giris-form input,
.giris-form .buton {
    min-height: 46px;
}

.giris-dipnot {
    margin: 20px 0 0;
    text-align: center;
    color: #89949a;
    font-size: 10px;
}

.yetkisiz-sayfa {
    min-height: calc(100vh - 170px);
    display: grid;
    place-items: center;
}

.durum-sayfasi {
    width: min(580px, 100%);
    padding: clamp(32px, 6vw, 54px);
    text-align: center;
    border: 1px solid var(--cizgi);
    border-radius: 18px;
    background: linear-gradient(145deg, rgba(255,255,255,.99), rgba(248,250,251,.98));
    box-shadow: var(--golge);
}

.durum-ikon {
    margin: 12px 0;
    color: #d7e0e4;
    font-size: clamp(60px, 11vw, 90px);
    line-height: 1;
    font-weight: 900;
}

.durum-sayfasi h1 {
    margin: 5px 0 10px;
    font-size: clamp(24px, 4vw, 32px);
}

.durum-sayfasi p:not(.ust-etiket) {
    max-width: 440px;
    margin: 0 auto 24px;
    color: var(--soluk);
    line-height: 1.6;
}

/* --------------------------------------------------------------------------
   Responsive
---------------------------------------------------------------------------- */

@media (max-width: 1100px) {
    .operator-grid {
        grid-template-columns: repeat(2, minmax(0, 1fr));
    }

    .ust-cubuk {
        grid-template-columns: 1fr auto;
    }

    .ana-nav {
        grid-column: 1 / -1;
        order: 3;
        overflow-x: auto;
    }

    .plan-satir {
        grid-template-columns: 44px minmax(180px, 1fr) 110px 120px minmax(150px, auto);
    }

    .plan-satir > span:nth-of-type(1) {
        display: none;
    }

    .monitor-grid {
        grid-template-columns: repeat(2, minmax(0, 1fr));
    }

    .monitor-plana {
        grid-column: 1 / -1;
    }


}

@media (max-width: 820px) {
    .sayfa-shell {
        gap: 15px;
    }

    .istatistik-grid {
        grid-template-columns: repeat(2, 1fr);
    }

    .pano-grid,
    .yonetim-grid,
    .donem-grid,
    .operator-grid,
    .monitor-grid {
        grid-template-columns: 1fr;
    }

    .monitor-plana {
        grid-column: auto;
    }

    .monitor-ust {
        align-items: flex-start;
    }

    .monitor-zaman {
        flex-shrink: 0;
    }



    .sayfa-baslik,
    .arac-cubugu {
        align-items: stretch;
        flex-direction: column;
    }

    .baslik-aksiyon {
        width: 100%;
    }

    .baslik-aksiyon .buton {
        flex: 1;
    }

    .operator-filtre-alani,
    .panel-filtre-alani {
        width: 100%;
        justify-items: stretch;
    }

    .operator-filtre-meta,
    .panel-filtre-alt {
        justify-content: space-between;
    }

    .filtreler {
        width: 100%;
    }

    .filtre {
        flex: 1 1 auto;
        justify-content: center;
    }

    .bilgi-grid {
        grid-template-columns: repeat(2, 1fr);
    }

    .oturum-metin {
        display: none;
    }

    .ana-nav {
        width: 100%;
        gap: 2px;
    }

    .ana-nav a {
        flex: 1 1 auto;
        padding: 12px 10px;
        text-align: center;
    }

    .plan-satir {
        grid-template-columns: 38px minmax(160px, 1fr) 100px minmax(140px, auto);
    }

    .plan-satir > span:nth-of-type(1),
    .plan-satir > span:nth-of-type(2) {
        display: none;
    }
    .admin-filtre-cubugu {
        align-items: stretch;
        flex-direction: column;
    }

    .admin-filtre-cubugu .filtreler {
        width: 100%;
    }

    .admin-filtre-cubugu .filtre {
        flex: 1 1 auto;
        justify-content: center;
    }

    .admin-filtre-sonuc {
        text-align: right;
    }
}

@media (max-width: 520px) {
    .sayfa,
    .monitor-body .sayfa {
        width: min(100% - 20px, 1520px);
        padding-top: 18px;
    }

    .ust-cubuk {
        padding: 9px 12px;
        gap: 10px;
    }

    .marka-metin small {
        display: none;
    }

    .marka-isaret {
        width: 44px;
        height: 40px;
    }

    .panel-kutu,
    .operator-arac-cubugu,
    .panel-arac-cubugu {
        padding: 14px;
    }

    .istatistik-grid {
        gap: 8px;
    }

    .istatistik {
        min-height: 94px;
        padding: 14px;
    }

    .istatistik strong {
        font-size: 28px;
    }

    .iki-kolon {
        grid-template-columns: 1fr;
    }

    .modal-kutu {
        padding: 16px;
    }

    .is-karti-ust,
    .operator-kart-ust {
        display: grid;
    }

    .operator-rozetler {
        justify-items: start;
    }

    .durum-rozet {
        white-space: normal;
    }

    .operator-filtre-alani .filtre,
    .panel-filtre-alani .filtre {
        flex: 1 1 calc(50% - 7px);
    }

    .giris-kutu {
        padding: 28px 22px;
    }

    .giris-logo {
        width: 145px;
        height: 80px;
    }

    .plan-satir {
        grid-template-columns: 34px 1fr;
        gap: 10px;
    }

    .plan-satir > span {
        display: none !important;
    }

    .monitor-ust {
        display: grid;
        gap: 14px;
    }

    .monitor-zaman {
        justify-items: start;
        text-align: left;
    }

    .monitor-zaman span {
        font-size: 32px;
    }

    .monitor-bolum {
        padding: 14px;
    }

    .monitor-liste {
        max-height: none;
        padding-right: 0;
        overflow: visible;
    }

    .monitor-kart-ust,
    .monitor-plan-ust {
        display: grid;
    }

    .monitor-hazir-bilgi {
        grid-template-columns: 1fr;
    }

    .monitor-plan-kart {
        grid-template-columns: 36px minmax(0, 1fr);
        gap: 10px;
    }

    .monitor-plan-sira {
        width: 34px;
        height: 34px;
    }

    .monitor-plan-meta {
        grid-template-columns: 1fr 1fr;
    }

    .monitor-plan-meta span:first-child {
        grid-column: 1 / -1;
    }

    .yonetim-durum-eksik-liste {
        max-height: 360px;
    }

    .yonetim-durum-eksik-liste .mini-satir {
        display: grid;
    }

    .yonetim-durum-eksik-liste .mini-sag {
        justify-items: start;
        text-align: left;
    }
    .admin-filtre-cubugu .filtre {
        flex: 1 1 calc(50% - 7px);
    }

    .yonetim-kart-scroll {
        max-height: 430px;
    }
}
```


# 4. SCRIPTS


## `run_pdgm.bat`


```bat
@echo off
REM PDGM sunucusunu baslatir.
REM Task Scheduler: "Run only when user is logged on" (Excel COM icin zorunlu)
REM On failure restart: her 1 dk, en fazla 3 kez.

cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
    echo HATA: .venv bulunamadi. Once: python -m venv .venv ^& .venv\Scripts\pip install -r requirements.txt
    exit /b 1
)

.venv\Scripts\python.exe app.py
exit /b %ERRORLEVEL%
```


## `yedek_disari_kopyala.bat`


```bat
@echo off
REM ============================================================
REM  PDGM - Gecelik yedek kopyalama
REM  Task Scheduler ile her gece calistirin.
REM  HEDEF yolunu kendi ag paylasiminizla degistirin.
REM  Robocopy: 0-7 basari/uyari, 8+ hata.
REM ============================================================

set KAYNAK=%~dp0data
set HEDEF=\\dosyasunucu\yedek\pdgm

if not exist "%HEDEF%" (
    echo HATA: Hedef erisilemiyor: %HEDEF%
    exit /b 1
)

for /f %%I in ('powershell -NoProfile -Command "Get-Date -Format yyyyMMdd"') do set BUGUN=%%I

robocopy "%KAYNAK%" "%HEDEF%\%BUGUN%" /E /R:2 /W:5 /NFL /NDL /LOG+:"%HEDEF%\robocopy.log"
set RC=%ERRORLEVEL%

powershell -NoProfile -Command ^
  "Get-ChildItem -Path '%HEDEF%' -Directory | Where-Object { $_.Name -match '^\d{8}$' -and $_.LastWriteTime -lt (Get-Date).AddDays(-60) } | Remove-Item -Recurse -Force"

if %RC% GEQ 8 (
    echo HATA: robocopy basarisiz, kod=%RC%
    exit /b %RC%
)

exit /b 0
```


# 5. CONFIG / DIGER


## `.env.example`


```dotenv
PDGM_PORT=5001
PDGM_HTTPS=0

# Varsayilan: tum ag arayuzleri. TLS terminator (Caddy) arkasinda 127.0.0.1 yapin.
# LAN'da dogrudan dinlemek bilincli bir karardir.
PDGM_BIND=0.0.0.0

# İlk kurulumda kullanicilar.json yokken okunur; bootstrap sonrası SİLİN.
# PDGM_ADMIN_PASSWORD=
# PDGM_OPERATOR_PASSWORD=
# PDGM_VIEWER_PASSWORD=
```


## `.gitignore`


```gitignore
.env
.venv/
venv/
__pycache__/
*.pyc
*.pyo
.DS_Store

data/
!data/.gitkeep

*.xlsx.bozuk_*
data/BASLATMA_HATASI.txt
```


## `requirements.txt`


```text
Flask==3.0.3
Werkzeug==3.0.6
openpyxl==3.1.5
waitress==3.0.0
python-dotenv==1.0.1
pywin32==308; sys_platform == "win32"
```

