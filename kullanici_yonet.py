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
