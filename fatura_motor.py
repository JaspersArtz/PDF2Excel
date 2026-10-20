from __future__ import annotations
import re
import logging
from datetime import datetime
from pathlib import Path
from typing import Optional

import pdfplumber
from openpyxl import load_workbook, Workbook

import ayarlar

# LLM opsiyonel
try:
    from llm_cikarim import kalemleri_cikar
    from llm_saglayici import LLMSaglayici
    LLM_KULLANILABILIR = True
except ImportError:
    LLM_KULLANILABILIR = False
    LLMSaglayici = None  # type: ignore

# OCR opsiyonel (taranmış PDF'ler için)
try:
    import pytesseract
    from pdf2image import convert_from_path # type: ignore
    OCR_KULLANILABILIR = True

    import os
    import sys
    if sys.platform == "win32":
        olası_yollar = [
            r"C:\Program Files\Tesseract-OCR\tesseract.exe",
            r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
            os.path.expanduser(r"~\AppData\Local\Programs\Tesseract-OCR\tesseract.exe"),
            os.path.expanduser(r"~\AppData\Local\Tesseract-OCR\tesseract.exe"),
        ]
        for yol in olası_yollar:
            if os.path.exists(yol):
                pytesseract.pytesseract.tesseract_cmd = yol
                break
except ImportError:
    OCR_KULLANILABILIR = False


logger = logging.getLogger("fatura_kalem_otomasyonu")
logger.setLevel(logging.INFO)
if not logger.handlers:
    handler = logging.StreamHandler()
    handler.setFormatter(
        logging.Formatter(
            "%(asctime)s [%(levelname)s] %(message)s",
            datefmt="%Y-%m-%d %H:%M:%S",
        )
    )
    logger.addHandler(handler)



# PDF METNİ OKUMA

def pdf_metni_oku(pdf_yolu: Path) -> tuple[str, bool]:

    metin = ""
    # 1. Önce pdfplumber dene
    try:
        with pdfplumber.open(pdf_yolu) as pdf:
            for sayfa in pdf.pages:
                metin += (sayfa.extract_text() or "") + "\n"
    except Exception as e:
        logger.warning(f"pdfplumber hatası: {e}")

    if len(metin.strip()) >= 200:
        return metin, False

    # 2. Metin yetersiz - OCR dene
    if not OCR_KULLANILABILIR:
        logger.warning(
            "PDF taranmış görünüyor ama OCR kurulu değil. "
            "Tesseract ve pdf2image kurun veya LLM kullanın "
            "(LLM PDF'i kendi okuyacak)."
        )
        return metin, False

    logger.info(f"PDF metni yetersiz, OCR çalışıyor: {pdf_yolu.name}")
    try:
        ocr_metni = ocr_ile_oku(pdf_yolu)
        if ocr_metni.strip():
            return ocr_metni, True
    except Exception as e:
        logger.warning(f"OCR hatası: {e}")

    return metin, False


def ocr_ile_oku(pdf_yolu: Path) -> str:
    """PDF sayfalarını görsele çevirip Tesseract ile OCR yapar."""
    metin = ""
    sayfalar = convert_from_path(str(pdf_yolu), dpi=300)
    for i, sayfa in enumerate(sayfalar, 1):
        try:
            # Türkçe + İngilizce
            sayfa_metni = pytesseract.image_to_string(sayfa, lang="tur+eng")
        except pytesseract.TesseractError:
            # Türkçe dil paketi yoksa İngilizce dene
            sayfa_metni = pytesseract.image_to_string(sayfa, lang="eng")
        metin += f"\n--- Sayfa {i} ---\n{sayfa_metni}\n"
    return metin



# REGEX YEDEK ÇIKARIM (LLM yoksa)

def _regex_ile_kalemleri_cikar(pdf_yolu: Path, sutunlar: list[str]) -> list[dict]:
    """Yedek mantık: pdfplumber ile tabloları çıkar, sütunlara eşle."""
    kalemler = []
    try:
        with pdfplumber.open(pdf_yolu) as pdf:
            for page in pdf.pages:
                tablolar = page.extract_tables() or []
                for tablo in tablolar:
                    if not tablo or len(tablo) < 2:
                        continue
                    baslik_satiri = [
                        (str(c).strip().lower() if c else "") for c in tablo[0]
                    ]
                    eslesme_sayisi = sum(
                        1 for b in baslik_satiri
                        if any(k in b for k in ["açıklama", "miktar", "tutar", "birim"])
                    )
                    if eslesme_sayisi < 2:
                        continue
                    for satir in tablo[1:]:
                        if not satir or all(not c for c in satir):
                            continue
                        kalem = {}
                        for i, sutun_adi in enumerate(sutunlar):
                            if i < len(satir) and satir[i]:
                                kalem[sutun_adi] = str(satir[i]).strip()
                            else:
                                kalem[sutun_adi] = ""
                        if any(v for v in kalem.values()):
                            kalemler.append(kalem)
                    if kalemler:
                        return kalemler
    except Exception as e:
        logger.warning(f"Regex tablo çıkarım hatası: {e}")
    return kalemler



# YENİ SAYFA ADI ÜRETME

def yeni_sayfa_adi_uret(wb, tarih: Optional[str] = None) -> str:
    """
    Bugün için boş bir sayfa adı bulur.
    İlk: Kalemler_YYYY-MM-DD
    Çakışırsa: Kalemler2_YYYY-MM-DD, Kalemler3_YYYY-MM-DD ...
    """
    if tarih is None:
        tarih = datetime.now().strftime("%Y-%m-%d")

    aday = f"Kalemler_{tarih}"
    if aday not in wb.sheetnames:
        return aday

    n = 2
    while True:
        aday = f"Kalemler{n}_{tarih}"
        if aday not in wb.sheetnames:
            return aday
        n += 1



# EXCEL'E YAZMA

def kalemleri_excele_yaz(
    excel_yolu: Path,
    kalemler: list[dict],
    sutunlar: list[str],
    toplu_mod: bool = False,
    hedef_sayfa: Optional[str] = None,
) -> str:
    """
    Excel'e kalemleri yazar.

    Args:
        excel_yolu: Hedef Excel
        kalemler: Yazılacak kalemler
        sutunlar: Sütun isimleri
        toplu_mod: True ise bugünün ilk "Kalemler_TARIH" sayfasına ekler
                   (yoksa oluşturur). False ise her seferinde yeni sayfa açar.
        hedef_sayfa: Belirli bir sayfaya yazmak isteniyorsa adı (klasör için)

    Returns:
        Yazılan sayfanın adı
    """
    # Excel yoksa oluştur
    if excel_yolu.exists():
        wb = load_workbook(excel_yolu)
    else:
        excel_yolu.parent.mkdir(parents=True, exist_ok=True)
        wb = Workbook()
        if "Sheet" in wb.sheetnames:
            del wb["Sheet"]

    from openpyxl.styles import Font

    # Sayfa belirleme
    if hedef_sayfa and hedef_sayfa in wb.sheetnames:
        # Klasör işleminde 2. ve sonraki PDF'ler: kilitli sayfaya ekleme
        sayfa_adi = hedef_sayfa
        ws = wb[sayfa_adi]
        baslangic_satir = ws.max_row + 1
    elif toplu_mod:
        # Klasör işleminin ilk PDF'i: HER ZAMAN YENİ sayfa aç
        # (önceki tekli işlemlerin sayfasına karışmasın)
        sayfa_adi = yeni_sayfa_adi_uret(wb)
        ws = wb.create_sheet(sayfa_adi)
        for i, sutun_adi in enumerate(sutunlar, start=1):
            hucre = ws.cell(row=1, column=i, value=sutun_adi)
            hucre.font = Font(bold=True)
        baslangic_satir = 2
    else:
        # Tekli mod: her zaman yeni sayfa
        sayfa_adi = yeni_sayfa_adi_uret(wb)
        ws = wb.create_sheet(sayfa_adi)
        # Başlık satırı
        for i, sutun_adi in enumerate(sutunlar, start=1):
            hucre = ws.cell(row=1, column=i, value=sutun_adi)
            hucre.font = Font(bold=True)
        baslangic_satir = 2

    # Kalemleri yaz
    for offset, kalem in enumerate(kalemler):
        satir_idx = baslangic_satir + offset
        for sut_idx, sutun_adi in enumerate(sutunlar, start=1):
            deger = kalem.get(sutun_adi, "")
            ws.cell(row=satir_idx, column=sut_idx, value=deger)

    # Sütun genişlikleri (mevcut + yeni içeriği göz önüne alarak)
    for i, sutun_adi in enumerate(sutunlar, start=1):
        col_letter = ws.cell(row=1, column=i).column_letter
        mevcut_genislik = ws.column_dimensions[col_letter].width or len(sutun_adi)
        max_uzunluk = max(mevcut_genislik, len(sutun_adi))
        for kalem in kalemler:
            deger = kalem.get(sutun_adi, "")
            max_uzunluk = max(max_uzunluk, len(str(deger)))
        ws.column_dimensions[col_letter].width = min(max_uzunluk + 2, 50)

    wb.save(excel_yolu)
    wb.close()
    return sayfa_adi



# ANA İŞLEM AKIŞI

def faturayi_isle(
    pdf_yolu: Path,
    excel_yolu: Path,
    saglayici: Optional["LLMSaglayici"] = None, # type: ignore
    sutunlar: Optional[list[str]] = None,
    toplu_mod: bool = False,
    hedef_sayfa: Optional[str] = None,
) -> dict:
    """
    Tek bir fatura PDF'ini işler.

    Args:
        pdf_yolu: Fatura PDF dosyası
        excel_yolu: Çıktı Excel dosyası
        saglayici: LLM sağlayıcı (None ise regex'le dener)
        sutunlar: Kullanılacak sütun isimleri (None ise ayarlar'dan okunur)
        toplu_mod: True ise bugünün ortak sayfasına ekler (klasör işlemi için)
        hedef_sayfa: Belirli bir sayfaya yazmak için (klasör işleminde hep aynı sayfa)

    Returns:
        İşlem sonucu dict
    """
    pdf_yolu = Path(pdf_yolu)
    excel_yolu = Path(excel_yolu)

    if not pdf_yolu.exists():
        return {"basarili": False, "hata": f"PDF bulunamadı: {pdf_yolu}"}

    # 1. Sütun listesini al
    if sutunlar is None:
        sutunlar = ayarlar.kalem_sutunlarini_al()

    if not sutunlar:
        return {
            "basarili": False,
            "hata": "Kalem sütunları tanımlı değil.\n"
                    "Lütfen '📋 Sütunlar' butonundan sütun isimlerini girin."
        }

    logger.info(f"▶ İşleniyor: {pdf_yolu.name}")
    logger.info(f"  Sütunlar: {', '.join(sutunlar)}")

    # 2. PDF metnini oku (pdfplumber + OCR fallback)
    metin, ocr_kullanildi = pdf_metni_oku(pdf_yolu)
    if ocr_kullanildi:
        logger.info(f"  OCR kullanıldı (taranmış PDF)")

    # 3. Kalemleri çıkar
    kalemler = []
    llm_kullanildi = False
    llm_hata = ""  # GUI'ye bilgi vermek için

    if saglayici is not None and LLM_KULLANILABILIR:
        try:
            kalemler, yanit = kalemleri_cikar(
                belge_metni=metin,
                excel_sutunlari=sutunlar,
                saglayici=saglayici,
                pdf_yolu=pdf_yolu,
            )
            if yanit.basarili and kalemler:
                llm_kullanildi = True
            elif yanit.basarili and not kalemler:
                llm_hata = "LLM yanıt verdi ama 'kalemler' listesi boş geldi (PDF'te tablo bulunamadı)."
                logger.warning(llm_hata)
            else:
                llm_hata = f"LLM çağrı hatası: {yanit.hata}"
                logger.warning(llm_hata)
        except Exception as e:
            llm_hata = f"LLM beklenmedik hata: {e}"
            logger.exception(llm_hata)

    # LLM yoksa veya başarısızsa regex yedek
    if not kalemler:
        if saglayici is not None:
            logger.info("LLM başarısız oldu, regex yedeğine geçiliyor...")
        else:
            logger.info("LLM kapalı, regex tablo çıkarımı kullanılıyor...")
        kalemler = _regex_ile_kalemleri_cikar(pdf_yolu, sutunlar)

    if not kalemler:
        return {
            "basarili": False,
            "hata": "Hiç kalem çıkarılamadı. PDF'te tablo yok olabilir veya format tanınmıyor.",
            "llm_kullanildi": llm_kullanildi,
        }

    # 4. Excel'e yaz
    try:
        sayfa_adi = kalemleri_excele_yaz(
            excel_yolu=excel_yolu,
            kalemler=kalemler,
            sutunlar=sutunlar,
            toplu_mod=toplu_mod,
            hedef_sayfa=hedef_sayfa,
        )
        logger.info(f"✓ {len(kalemler)} kalem '{sayfa_adi}' sayfasına yazıldı")
    except Exception as e:
        return {"basarili": False, "hata": f"Excel'e yazma hatası: {e}"}

    return {
        "basarili": True,
        "sayfa_adi": sayfa_adi,
        "kalem_sayisi": len(kalemler),
        "excel": str(excel_yolu),
        "pdf": pdf_yolu.name,
        "llm_kullanildi": llm_kullanildi,
        "llm_hata": llm_hata,  # Eğer LLM denendi ama başarısız olduysa açıklama
        "ocr_kullanildi": ocr_kullanildi,
    }


def klasoru_isle(
    klasor: Path,
    excel_yolu: Path,
    saglayici: Optional["LLMSaglayici"] = None, # type: ignore
    sutunlar: Optional[list[str]] = None,
) -> list[dict]:
    """
    Bir klasördeki tüm PDF'leri sırayla işler.

    Tüm PDF'lerin kalemleri TEK SAYFAYA alt alta yazılır.
    Bugün için sayfa zaten varsa (önceki işlem), onun sonuna eklenir.
    """
    klasor = Path(klasor)
    pdf_dosyalari = sorted(klasor.glob("*.pdf"))

    if not pdf_dosyalari:
        logger.warning(f"Klasörde PDF bulunamadı: {klasor}")
        return []

    # Sütunları bir kez al
    if sutunlar is None:
        sutunlar = ayarlar.kalem_sutunlarini_al()

    if not sutunlar:
        return [{
            "basarili": False,
            "pdf": pdf.name,
            "hata": "Kalem sütunları tanımlı değil. '📋 Sütunlar' butonundan girin.",
        } for pdf in pdf_dosyalari]

    logger.info(f"{len(pdf_dosyalari)} PDF bulundu, hepsi TEK SAYFAYA işleniyor...")

    # Tüm klasör için tek hedef sayfa adı belirle
    # Eğer bugünün "Kalemler_TARIH" sayfası varsa kullan, yoksa kalemleri_excele_yaz
    # otomatik oluşturacak
    hedef_sayfa_adi = None
    sonuclar = []

    for pdf in pdf_dosyalari:
        sonuc = faturayi_isle(
            pdf, excel_yolu,
            saglayici=saglayici,
            sutunlar=sutunlar,
            toplu_mod=True,
            hedef_sayfa=hedef_sayfa_adi,
        )
        sonuc["pdf"] = pdf.name
        sonuclar.append(sonuc)

        # İlk başarılı işlemden sonra sayfa adını kilitle
        # (sonraki PDF'ler de aynı sayfaya gitsin)
        if sonuc.get("basarili") and not hedef_sayfa_adi:
            hedef_sayfa_adi = sonuc.get("sayfa_adi")

    return sonuclar



# KOMUT SATIRI

if __name__ == "__main__":
    import sys
    if len(sys.argv) < 3:
        print("Kullanım: python fatura_motor.py <pdf> <excel> [--llm <claude|gemini>]")
        print("\nNot: Sütunları önce arayüzden veya kalem_sutunlari.json ile tanımlayın.")
        sys.exit(1)

    pdf_yolu = Path(sys.argv[1])
    excel_yolu = Path(sys.argv[2])
    saglayici = None

    if "--llm" in sys.argv:
        from llm_saglayici import saglayici_olustur
        idx = sys.argv.index("--llm")
        saglayici_adi = sys.argv[idx + 1] if len(sys.argv) > idx + 1 else "claude"
        try:
            saglayici = saglayici_olustur(saglayici_adi)
            print(f"LLM aktif: {saglayici.ad}/{saglayici.model}")
        except Exception as e:
            print(f"LLM kurulamadı: {e}")

    if pdf_yolu.is_dir():
        sonuclar = klasoru_isle(pdf_yolu, excel_yolu, saglayici=saglayici)
        basarili = sum(1 for s in sonuclar if s["basarili"])
        print(f"\nÖzet: {basarili}/{len(sonuclar)} başarılı")
    else:
        sonuc = faturayi_isle(pdf_yolu, excel_yolu, saglayici=saglayici)
        print(f"\nSonuç: {sonuc}")