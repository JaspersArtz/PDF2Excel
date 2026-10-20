

from __future__ import annotations
import json
import logging
import re
from abc import ABC, abstractmethod
from pathlib import Path
from typing import Optional, Union
from dataclasses import dataclass

import ayarlar


logger = logging.getLogger("ihale_otomasyonu")



# VERİ YAPISI - LLM yanıtı

@dataclass
class LLMYaniti:
    """LLM'den dönen yanıtı standart formatta tutar."""
    basarili: bool
    veri: dict                       # Çıkarılan alanlar (boşsa {})
    ham_yanit: str = ""              # Debugging için orijinal LLM çıktısı
    hata: str = ""                   # Hata varsa açıklaması
    saglayici: str = ""              # "claude" / "gemini"
    model: str = ""                  # Kullanılan modelin adı
    girdi_token: int = 0             # Maliyet takibi
    cikti_token: int = 0



# JSON ÇIKARMA YARDIMCI FONKSİYONU

def _yanittan_json_cikar(metin: str) -> Optional[dict]:
    """
    LLM yanıtından JSON nesnesini güvenle çıkarır.
    Markdown ```json fence'leri, etrafındaki açıklamaları temizler.
    """
    if not metin:
        return None

    # Markdown fence'leri temizle
    temiz = re.sub(r"```(?:json)?\s*", "", metin)
    temiz = re.sub(r"```", "", temiz)
    temiz = temiz.strip()

    # Önce direkt parse dene
    try:
        return json.loads(temiz)
    except json.JSONDecodeError:
        pass

    # Metnin içindeki ilk { ... } bloğunu yakala (greedy)
    eslesme = re.search(r"\{.*\}", temiz, re.DOTALL)
    if eslesme:
        try:
            return json.loads(eslesme.group(0))
        except json.JSONDecodeError:
            pass

    return None



# SOYUT SAĞLAYICI ARAYÜZÜ

class LLMSaglayici(ABC):
    """Tüm LLM sağlayıcıları için ortak arayüz."""

    ad: str = ""  # "claude" / "gemini"

    def __init__(self, api_anahtari: str, model: Optional[str] = None):
        if not api_anahtari:
            raise ValueError(f"{self.ad} için API anahtarı gerekli.")
        self.api_anahtari = api_anahtari
        self.model = model or ayarlar.varsayilan_model(self.ad)

    @abstractmethod
    def cikar(
        self,
        sistem_prompt: str,
        kullanici_prompt: str,
        pdf_yolu: Optional[Path] = None,
    ) -> LLMYaniti:
        """
        LLM'e prompt gönderip yapılandırılmış JSON yanıtını alır.

        Args:
            sistem_prompt: Modelin rolünü/davranışını tanımlar
            kullanici_prompt: Asıl talimat ve PDF metni
            pdf_yolu: Eğer dolu PDF göndermek istenirse (taranmış PDF için)

        Returns:
            LLMYaniti nesnesi
        """
        ...



# CLAUDE SAĞLAYICI

class ClaudeSaglayici(LLMSaglayici):
    ad = "claude"

    def __init__(self, api_anahtari: str, model: Optional[str] = None):
        super().__init__(api_anahtari, model)
        try:
            from anthropic import Anthropic
        except ImportError:
            raise ImportError(
                "anthropic paketi kurulu değil. Kurmak için:\n"
                "  pip install anthropic"
            )
        self._istemci = Anthropic(api_key=self.api_anahtari)

    def cikar(
        self,
        sistem_prompt: str,
        kullanici_prompt: str,
        pdf_yolu: Optional[Path] = None,
    ) -> LLMYaniti:
        # Mesaj içeriğini hazırla
        if pdf_yolu and Path(pdf_yolu).exists():
            # PDF'i base64 olarak gönder (Claude native PDF desteği)
            import base64
            pdf_b64 = base64.standard_b64encode(Path(pdf_yolu).read_bytes()).decode("ascii")
            icerik = [
                {
                    "type": "document",
                    "source": {
                        "type": "base64",
                        "media_type": "application/pdf",
                        "data": pdf_b64,
                    },
                },
                {"type": "text", "text": kullanici_prompt},
            ]
        else:
            icerik = kullanici_prompt

        try:
            yanit = self._istemci.messages.create(
                model=self.model,
                max_tokens=4096,
                system=sistem_prompt,
                messages=[{"role": "user", "content": icerik}],
            )
            ham_metin = "".join(
                blok.text for blok in yanit.content if hasattr(blok, "text")
            )
            veri = _yanittan_json_cikar(ham_metin)

            return LLMYaniti(
                basarili=veri is not None,
                veri=veri or {},
                ham_yanit=ham_metin,
                hata="" if veri else "JSON parse edilemedi",
                saglayici=self.ad,
                model=self.model,
                girdi_token=getattr(yanit.usage, "input_tokens", 0),
                cikti_token=getattr(yanit.usage, "output_tokens", 0),
            )
        except Exception as e:
            logger.exception(f"Claude API hatası")
            return LLMYaniti(
                basarili=False,
                veri={},
                hata=f"Claude API hatası: {e}",
                saglayici=self.ad,
                model=self.model,
            )



# GEMINI SAĞLAYICI

class GeminiSaglayici(LLMSaglayici):
    ad = "gemini"

    def __init__(self, api_anahtari: str, model: Optional[str] = None):
        super().__init__(api_anahtari, model)
        try:
            import google.generativeai as genai
        except ImportError:
            raise ImportError(
                "google-generativeai paketi kurulu değil. Kurmak için:\n"
                "  pip install google-generativeai"
            )
        genai.configure(api_key=self.api_anahtari)
        self._genai = genai

    def cikar(
        self,
        sistem_prompt: str,
        kullanici_prompt: str,
        pdf_yolu: Optional[Path] = None,
    ) -> LLMYaniti:
        try:
            # Gemini'de sistem talimatı modele constructor'da verilir
            model_nesnesi = self._genai.GenerativeModel(
                model_name=self.model,
                system_instruction=sistem_prompt,
                generation_config={
                    "temperature": 0.0,        # Deterministik çıkarım için
                    "response_mime_type": "application/json",  # Saf JSON döndür
                },
            )

            if pdf_yolu and Path(pdf_yolu).exists():
                # PDF'i inline gönder (küçük dosyalar için, <20MB)
                pdf_part = {
                    "mime_type": "application/pdf",
                    "data": Path(pdf_yolu).read_bytes(),
                }
                yanit = model_nesnesi.generate_content([pdf_part, kullanici_prompt])
            else:
                yanit = model_nesnesi.generate_content(kullanici_prompt)

            ham_metin = yanit.text or ""
            veri = _yanittan_json_cikar(ham_metin)

            # Token kullanım bilgisi
            girdi_t = 0
            cikti_t = 0
            try:
                kullanim = yanit.usage_metadata
                girdi_t = getattr(kullanim, "prompt_token_count", 0)
                cikti_t = getattr(kullanim, "candidates_token_count", 0)
            except AttributeError:
                pass

            return LLMYaniti(
                basarili=veri is not None,
                veri=veri or {},
                ham_yanit=ham_metin,
                hata="" if veri else "JSON parse edilemedi",
                saglayici=self.ad,
                model=self.model,
                girdi_token=girdi_t,
                cikti_token=cikti_t,
            )
        except Exception as e:
            logger.exception(f"Gemini API hatası")
            return LLMYaniti(
                basarili=False,
                veri={},
                hata=f"Gemini API hatası: {e}",
                saglayici=self.ad,
                model=self.model,
            )



# FACTORY: Sağlayıcı seçimi

def saglayici_olustur(
    saglayici_adi: str,
    api_anahtari: Optional[str] = None,
    model: Optional[str] = None,
) -> LLMSaglayici:
    """
    İsmine göre uygun LLMSaglayici nesnesini üretir.

    saglayici_adi: "claude" veya "gemini"
    api_anahtari: None ise ayarlar.api_anahtari_al() kullanılır
    """
    saglayici_adi = saglayici_adi.lower().strip()

    if api_anahtari is None:
        api_anahtari = ayarlar.api_anahtari_al(saglayici_adi)
        if api_anahtari is None:
            raise RuntimeError(
                f"'{saglayici_adi}' için API anahtarı bulunamadı.\n"
                f".env dosyasına ekleyin veya GUI'den ayarlayın."
            )

    if saglayici_adi in ("claude", "anthropic"):
        return ClaudeSaglayici(api_anahtari, model)
    elif saglayici_adi in ("gemini", "google"):
        return GeminiSaglayici(api_anahtari, model)
    else:
        raise ValueError(
            f"Desteklenmeyen sağlayıcı: '{saglayici_adi}'. "
            f"Desteklenenler: claude, gemini"
        )
