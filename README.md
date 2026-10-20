
### 1. Python
Python 3.10+ gerekli. [python.org](https://www.python.org/downloads/)

### 2. Kütüphaneler
```
> pip install -r requirements.txt
> py -m pip install pdfplumber openpyxl anthropic google-generativeai python-dotenv pytesseract pdf2image Pillow

```

### Hangi alanlar regex, hangi alanlar LLM?

**Faturada regex (net format, %100 doğru):**
- VKN, IBAN, fatura no, tarihler, tutarlar

**Faturada LLM (serbest metin, regex yetersiz):**
- Satıcı firma adı, alıcı/idare, ihale konusu, ödeme şekli

**İhale ilanında regex:**
- İhale Kayıt No, tarihler, yaklaşık maliyet

**İhale ilanında LLM:**
- İdare adı, ihale konusu, ihale yeri, ihale usulü, iletişim

> Bu paylaşım maliyeti %70 düşürür ve hızı 2-3 kat artırır.



