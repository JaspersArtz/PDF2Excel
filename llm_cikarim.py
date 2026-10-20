

from __future__ import annotations
import json
import logging
import re
from typing import Optional
from pathlib import Path

from llm_saglayici import LLMSaglayici, LLMYaniti


logger = logging.getLogger("fatura_kalem_otomasyonu")



# SİSTEM PROMPT

SISTEM_PROMPT = """Sen Türkçe e-faturalardan KALEM TABLOSUNU (ürün/hizmet listesini) \
çıkaran uzman bir veri çıkarım asistanısın.

Tek görevin: Faturanın kalem satırlarını (her bir ürün/hizmet) JSON dizisi olarak döndürmek.

KURALLAR:
1. Sadece JSON yanıt ver, hiçbir açıklama yapma
2. Her kalem ayrı bir nesne olmalı
3. Bulamadığın alan için boş string "" döndür - ASLA tahmin etme
4. Sayıları orijinal formatında bırak (örn: "1.234,56" veya "1234.56")
5. KDV oranını "%18" veya "18" formatında ver, sayıyı koru
6. Para birimi sembollerini yazma, sadece rakamları al
7. Açıklama metnini olduğu gibi al, kısaltma
8. SADECE faturanın kalem tablosundaki satırları al
9. Genel toplam, KDV toplam gibi ÖZET SATIRLARINI ALMA
10. Fatura başlık bilgilerini (satıcı, alıcı, fatura no vs.) ALMA"""



# KULLANICI PROMPT'U

KALEM_PROMPT_SABLONU = """Aşağıdaki Türkçe e-faturadan kalem tablosunu (ürün/hizmet listesini) çıkar.

Excel tablomda şu sütunlar var (bu sırayla):
{sutun_listesi}

Her kalem için bu alanları içeren JSON dizisi döndür.

JSON şeması:
{{
  "kalemler": [
    {{
{ornek_alanlar}
    }},
    ...
  ]
}}

ÖNEMLİ:
- Yalnızca kalem satırlarını dahil et (toplam, KDV toplam, genel toplam SATIRI dahil etme)
- Sıra No yoksa, kalemleri 1'den başlatıp sırala
- Bulamadığın alanlar için boş string "" döndür
- Kalemler tablosunda 5 satır varsa, sen de 5 kalem döndür

BELGE METNİ:
---
{belge_metni}
---

Sadece JSON döndür:"""


def _yanittan_json_cikar(metin: str) -> Optional[dict]:
    """LLM yanıtından JSON nesnesi/dizisini çıkarır."""
    if not metin:
        return None
    temiz = re.sub(r"```(?:json)?\s*", "", metin)
    temiz = re.sub(r"```", "", temiz)
    temiz = temiz.strip()
    try:
        return json.loads(temiz)
    except json.JSONDecodeError:
        pass
    eslesme = re.search(r"\{.*\}", temiz, re.DOTALL)
    if eslesme:
        try:
            return json.loads(eslesme.group(0))
        except json.JSONDecodeError:
            pass
    return None



# ANA FONKSİYON

def kalemleri_cikar(
    belge_metni: str,
    excel_sutunlari: list[str],
    saglayici: LLMSaglayici,
    pdf_yolu: Optional[Path] = None,
) -> tuple[list[dict], LLMYaniti]:
    """
    PDF metninden kalem tablosunu çıkarır.

    Args:
        belge_metni: PDF'in tüm metni
        excel_sutunlari: Excel'deki sütun başlıkları (sıralı liste)
                         örn: ["Sıra No", "Açıklama", "Miktar", "Birim", ...]
        saglayici: LLM sağlayıcı (Claude veya Gemini)
        pdf_yolu: Metin yetersizse PDF'i doğrudan göndermek için

    Returns:
        (kalemler, llm_yaniti) tuple'ı
        kalemler: [{"Sıra No": "1", "Açıklama": "...", ...}, ...]
    """
    # Sütun listesini prompt için hazırla
    sutun_metni = "\n".join(f"  {i+1}. {s}" for i, s in enumerate(excel_sutunlari))

    # JSON şeması için örnek alanlar (sütunlardan üretilir)
    ornek = []
    for sutun in excel_sutunlari:
        ornek.append(f'      "{sutun}": "string"')
    ornek_metni = ",\n".join(ornek)

    kullanici_prompt = KALEM_PROMPT_SABLONU.format(
        sutun_listesi=sutun_metni,
        ornek_alanlar=ornek_metni,
        belge_metni=belge_metni[:30000],
    )

    # PDF'i mi metni mi gönder?
    pdf_gonder = pdf_yolu and len(belge_metni.strip()) < 200

    yanit = saglayici.cikar(
        sistem_prompt=SISTEM_PROMPT,
        kullanici_prompt=kullanici_prompt,
        pdf_yolu=pdf_yolu if pdf_gonder else None,
    )

    if not yanit.basarili:
        logger.warning(f"LLM çıkarım başarısız: {yanit.hata}")
        return [], yanit

    # JSON'dan "kalemler" listesini al
    kalemler_ham = yanit.veri.get("kalemler", [])
    if not isinstance(kalemler_ham, list):
        logger.warning("LLM 'kalemler' anahtarı liste değil")
        return [], yanit

    # Her kalemi normalize et - sadece istenen sütunları al
    kalemler = []
    for k in kalemler_ham:
        if not isinstance(k, dict):
            continue
        kalem = {}
        for sutun in excel_sutunlari:
            deger = k.get(sutun, "")
            # Sayı veya başka tip gelirse string'e çevir
            if deger is None:
                deger = ""
            elif not isinstance(deger, str):
                deger = str(deger)
            kalem[sutun] = deger.strip()
        kalemler.append(kalem)

    logger.info(
        f"✓ {len(kalemler)} kalem çıkarıldı "
        f"({saglayici.ad}/{saglayici.model}, "
        f"girdi:{yanit.girdi_token}, çıktı:{yanit.cikti_token} token)"
    )
    return kalemler, yanit
