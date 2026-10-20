from __future__ import annotations
import os
import json
from pathlib import Path
from typing import Optional

# python-dotenv opsiyonel
try:
    from dotenv import load_dotenv, set_key
    DOTENV_KULLANILABILIR = True
except ImportError:
    DOTENV_KULLANILABILIR = False


PROJE_KOK = Path(__file__).parent.resolve()
ENV_DOSYASI = PROJE_KOK / ".env"
SUTUN_DOSYASI = PROJE_KOK / "kalem_sutunlari.json"



# .env DOSYA YÖNETİMİ

def env_yukle() -> None:
    """Proje kökündeki .env dosyasını ortam değişkenlerine yükler."""
    if not ENV_DOSYASI.exists():
        return

    if DOTENV_KULLANILABILIR:
        load_dotenv(ENV_DOSYASI, override=False)
        return

    # Basit fallback parser
    for satir in ENV_DOSYASI.read_text(encoding="utf-8").splitlines():
        satir = satir.strip()
        if not satir or satir.startswith("#") or "=" not in satir:
            continue
        anahtar, _, deger = satir.partition("=")
        anahtar = anahtar.strip()
        deger = deger.strip().strip('"').strip("'")
        if anahtar and anahtar not in os.environ:
            os.environ[anahtar] = deger


def env_yaz(anahtar: str, deger: str) -> None:
    """.env dosyasına anahtar=değer yazar."""
    if DOTENV_KULLANILABILIR:
        ENV_DOSYASI.touch(exist_ok=True)
        set_key(str(ENV_DOSYASI), anahtar, deger, quote_mode="never")
        os.environ[anahtar] = deger
        return

    satirlar = []
    bulundu = False
    if ENV_DOSYASI.exists():
        for s in ENV_DOSYASI.read_text(encoding="utf-8").splitlines():
            if s.strip().startswith(f"{anahtar}="):
                satirlar.append(f"{anahtar}={deger}")
                bulundu = True
            else:
                satirlar.append(s)
    if not bulundu:
        satirlar.append(f"{anahtar}={deger}")

    ENV_DOSYASI.write_text("\n".join(satirlar) + "\n", encoding="utf-8")
    os.environ[anahtar] = deger



# API ANAHTARI ALMA

def api_anahtari_al(saglayici: str) -> Optional[str]:
    env_yukle()
    eslesme = {
        "claude": "CLAUDE_API_KEY",
        "anthropic": "CLAUDE_API_KEY",
        "gemini": "GEMINI_API_KEY",
        "google": "GEMINI_API_KEY",
    }
    env_anahtari = eslesme.get(saglayici.lower())
    if not env_anahtari:
        return None
    deger = os.environ.get(env_anahtari, "").strip()
    return deger if deger else None


def api_anahtari_kaydet(saglayici: str, deger: str) -> None:
    eslesme = {
        "claude": "CLAUDE_API_KEY",
        "gemini": "GEMINI_API_KEY",
    }
    env_anahtari = eslesme.get(saglayici.lower())
    if not env_anahtari:
        raise ValueError(f"Bilinmeyen sağlayıcı: {saglayici}")
    env_yaz(env_anahtari, deger.strip())



# VARSAYILAN MODELLER

VARSAYILAN_MODELLER = {
    "claude": "claude-haiku-4-5-20251001",
    "gemini": "gemini-2.5-flash",  # 2.0-flash ücretsiz kotadan kaldırıldı (Haz 2026)
}


def varsayilan_model(saglayici: str) -> str:
    return VARSAYILAN_MODELLER.get(saglayici.lower(), "")


def hangi_saglayicilar_hazir() -> dict:
    return {
        "claude": api_anahtari_al("claude") is not None,
        "gemini": api_anahtari_al("gemini") is not None,
    }



# KALEM SÜTUNLARI YÖNETİMİ (YENİ)

def kalem_sutunlarini_al() -> list[str]:
    """
    Kayıtlı kalem sütunlarını döner.

    Returns:
        Sütun isimleri listesi. Hiç kayıt yoksa boş liste.
    """
    if not SUTUN_DOSYASI.exists():
        return []

    try:
        veri = json.loads(SUTUN_DOSYASI.read_text(encoding="utf-8"))
        if isinstance(veri, list):
            # Geriye uyumluluk: eskiden direkt liste olabilir
            return [str(s).strip() for s in veri if str(s).strip()]
        if isinstance(veri, dict) and "sutunlar" in veri:
            return [str(s).strip() for s in veri["sutunlar"] if str(s).strip()]
    except (json.JSONDecodeError, OSError):
        pass

    return []


def kalem_sutunlarini_kaydet(sutunlar: list[str]) -> None:
    """
    Kalem sütunlarını JSON dosyasına yazar.

    Args:
        sutunlar: Sütun isimleri listesi (sırası önemli)
    """
    # Temizle: boşları at, boşlukları kırp
    temiz_sutunlar = [str(s).strip() for s in sutunlar if str(s).strip()]

    veri = {
        "sutunlar": temiz_sutunlar,
    }
    SUTUN_DOSYASI.write_text(
        json.dumps(veri, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


if __name__ == "__main__":
    env_yukle()
    durumlar = hangi_saglayicilar_hazir()
    print(f".env yolu:           {ENV_DOSYASI}")
    print(f"Sütun config yolu:   {SUTUN_DOSYASI}")
    print(f"Claude API key:      {'✓ ayarlı' if durumlar['claude'] else '✗ yok'}")
    print(f"Gemini API key:      {'✓ ayarlı' if durumlar['gemini'] else '✗ yok'}")
    sutunlar = kalem_sutunlarini_al()
    if sutunlar:
        print(f"Kayıtlı sütunlar ({len(sutunlar)}):")
        for i, s in enumerate(sutunlar, 1):
            print(f"  {i}. {s}")
    else:
        print("Henüz sütun kaydedilmemiş.")