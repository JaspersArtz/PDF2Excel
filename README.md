
# Invoice & Tender Data Extraction System

A hybrid document processing setup combining **Regex** for structured fields and **LLMs** for unstructured text extraction.

### 1. Python
Python 3.10+ required. [python.org](https://www.python.org/downloads/)

### 2. Libraries
```bash
pip install -r requirements.txt
py -m pip install pdfplumber openpyxl anthropic google-generativeai python-dotenv pytesseract pdf2image Pillow
```

### Which fields use regex, which fields use LLM?

**Regex in invoice (strict format, 100% accurate):**
- Tax ID (VKN), IBAN, invoice number, dates, amounts

**LLM in invoice (free text, regex insufficient):**
- Seller company name, buyer/authority, tender subject, payment method

**Regex in tender notice:**
- Tender Registration No, dates, estimated cost

**LLM in tender notice:**
- Authority name, tender subject, tender venue, tender procedure, contact info

> This split reduces cost by 70% and increases speed by 2-3x.



