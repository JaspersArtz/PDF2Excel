
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

<img width="957" height="748" alt="image" src="https://github.com/user-attachments/assets/76b01abb-6b6b-45e8-a42d-1dd2b15625b6" />
<img width="959" height="746" alt="image" src="https://github.com/user-attachments/assets/0e902ee0-c14f-4668-a6b3-9c78e86d225e" />



