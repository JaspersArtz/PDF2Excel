import threading
from pathlib import Path
from tkinter import (
    Tk, Toplevel, Frame, Label, Button, Text, Scrollbar, Entry,
    filedialog, messagebox, StringVar, BooleanVar, END, DISABLED, NORMAL,
    PhotoImage, font as tkfont,
)
from tkinter.ttk import Progressbar, Style, Combobox

# Logo
try:
    from PIL import Image
    from PIL.ImageTk import PhotoImage as PILPhotoImage
    PIL_KULLANILABILIR = True
except ImportError:
    PIL_KULLANILABILIR = False

from fatura_motor import faturayi_isle, OCR_KULLANILABILIR

# LLM 
try:
    import ayarlar
    from llm_saglayici import saglayici_olustur
    LLM_KULLANILABILIR = True
except ImportError:
    LLM_KULLANILABILIR = False

RENK_BG = "#F5F7FA"
RENK_KART = "#FFFFFF"
RENK_BIRINCIL = "#1F4E78"
RENK_VURGU = "#2E86AB"
RENK_BASARI = "#2E7D32"
RENK_HATA = "#C62828"
RENK_UYARI = "#EF6C00"
RENK_METIN = "#212121"
RENK_IKINCIL_METIN = "#616161"
RENK_LLM = "#7B1FA2"
RENK_SUTUN = "#00838F"  


class OtomasyonApp:
    def __init__(self, root: Tk):
        self.root = root
        self.root.title("ISKI OTOMASYON")
        self.root.geometry("960x720")
        self.root.configure(bg=RENK_BG)
        self.root.minsize(820, 620)
        self._center_window(960, 720)

        # Excel yolu
        script_klasor = Path(__file__).parent.resolve()
        varsayilan = script_klasor / "tables" / "FaturaTakip.xlsx"
        if not varsayilan.exists():
            varsayilan = script_klasor / "FaturaTakip.xlsx"
        self.fatura_excel_yolu = StringVar(value=str(varsayilan))

        # LLM ayarları
        self.llm_secim = StringVar(value="kapali")
        self.llm_durum = StringVar(value="LLM: Kapalı")
        self._aktif_saglayici = None

        # Yazı tipleri
        self.f_baslik = tkfont.Font(family="Segoe UI", size=18, weight="bold")
        self.f_buton = tkfont.Font(family="Segoe UI", size=11, weight="bold")
        self.f_buton_kucuk = tkfont.Font(family="Segoe UI", size=9)
        self.f_log = tkfont.Font(family="Consolas", size=9)

        # İkonlar (asset/ klasöründen) - garbage collection'a karşı self'te tut
        self._ikonlar: dict[tuple[str, int], "PILPhotoImage"] = {}

        self._arayuzu_kur()
        self._baslangic_kontrolleri()

    def _ikon_yukle(self, dosya_adi: str, boyut: int = 16):
        """asset/ klasöründen bir ikon yükler ve cache'ler. PIL yoksa None döner."""
        if not PIL_KULLANILABILIR:
            return None
        anahtar = (dosya_adi, boyut)
        if anahtar in self._ikonlar:
            return self._ikonlar[anahtar]
        try:
            yol = Path(__file__).resolve().parent / "asset" / dosya_adi
            if not yol.exists():
                return None
            img = Image.open(str(yol))
            img.thumbnail((boyut, boyut), Image.Resampling.LANCZOS)
            tk_img = PILPhotoImage(image=img)
            self._ikonlar[anahtar] = tk_img
            return tk_img
        except Exception:
            return None

    def _center_window(self, width: int, height: int):
        self.root.update_idletasks()
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        self.root.geometry(f"{width}x{height}+{x}+{y}")

    def _arayuzu_kur(self):
        # İçerik alanı
        icerik = Frame(self.root, bg=RENK_BG, padx=20, pady=18)
        icerik.pack(fill="both", expand=True)

        # Bilgi kartı
        bilgi_kart = Frame(icerik, bg=RENK_KART, padx=14, pady=10,
                            highlightbackground="#E0E0E0", highlightthickness=1)
        bilgi_kart.pack(fill="x", pady=(0, 10))

        bilgi_satir = Frame(bilgi_kart, bg=RENK_KART)
        bilgi_satir.pack(fill="x")

        Label(
            bilgi_satir,
            text="Fatura PDF'inden Exele yeni sayfa olarak eklenir.",
            font=("Segoe UI", 9), bg=RENK_KART, fg=RENK_IKINCIL_METIN,
            justify="left",
        ).pack(side="left", anchor="w")

        # Sütun durumu göstergesi
        self.sutun_durum = StringVar(value="")
        Label(
            bilgi_satir, textvariable=self.sutun_durum,
            font=("Segoe UI", 9, "bold"), bg=RENK_KART, fg=RENK_SUTUN,
        ).pack(side="right", padx=10)

        # LLM kontrol satırı
        llm_kart = Frame(icerik, bg=RENK_KART, padx=12, pady=10,
                         highlightbackground="#E0E0E0", highlightthickness=1)
        llm_kart.pack(fill="x", pady=(0, 10))

        Label(llm_kart, text="LLM Seçin:",
              font=("Segoe UI", 10, "bold"),
              bg=RENK_KART, fg=RENK_METIN).pack(side="left", padx=(0, 10))

        self.llm_combo = Combobox(
            llm_kart, textvariable=self.llm_secim,
            values=["kapali", "claude", "gemini"],
            state="readonly", width=15,
        )
        self.llm_combo.pack(side="left", padx=(0, 10))
        self.llm_combo.bind("<<ComboboxSelected>>", self._llm_secim_degisti)

        Label(llm_kart, textvariable=self.llm_durum,
              font=("Segoe UI", 9), bg=RENK_KART, fg=RENK_IKINCIL_METIN
              ).pack(side="left", padx=10)

        # Excel kartı
        excel_satiri = Frame(icerik, bg=RENK_BG)
        excel_satiri.pack(fill="x", pady=(0, 12))

        self._excel_karti_olustur(
            excel_satiri, "Seçilen Exel", self.fatura_excel_yolu
        )

        # Ana eylem butonu
        eylem = Frame(icerik, bg=RENK_BG)
        eylem.pack(fill="x", pady=(0, 12))

        self.pdf_buton = Button(
            eylem, text="PDF Seç ve İşle",
            font=self.f_buton, bg=RENK_BIRINCIL, fg="white",
            relief="flat", padx=20, pady=14, cursor="hand2",
            command=self._pdf_sec_ve_isle,
        )
        self.pdf_buton.pack(fill="x", expand=True)

        # İlerleme
        stil = Style()
        stil.theme_use("default")
        stil.configure("ince.Horizontal.TProgressbar",
                       troughcolor=RENK_BG, background=RENK_VURGU,
                       borderwidth=0, thickness=4)

        self.ilerleme = Progressbar(
            icerik, mode="indeterminate",
            style="ince.Horizontal.TProgressbar",
        )
        self.ilerleme.pack(fill="x", pady=(0, 8))

        # Log alanı
        log_kart = Frame(icerik, bg=RENK_KART,
                         highlightbackground="#E0E0E0", highlightthickness=1)
        log_kart.pack(fill="both", expand=True)

        log_baslik = Frame(log_kart, bg=RENK_KART, padx=12, pady=8)
        log_baslik.pack(fill="x")
        Label(log_baslik, text="Yapılan İşlemler:",
              font=("Segoe UI", 10, "bold"),
              bg=RENK_KART, fg=RENK_METIN).pack(side="left")
        Button(
            log_baslik, text="Temizle",
            font=self.f_buton_kucuk, bg="#EEEEEE", fg=RENK_METIN,
            relief="flat", padx=10, pady=2, cursor="hand2",
            command=self._log_temizle,
        ).pack(side="right")

        log_alan = Frame(log_kart, bg=RENK_KART, padx=12, pady=12)
        log_alan.pack(fill="both", expand=True)

        self.log_text = Text(
            log_alan, height=10, font=self.f_log,
            bg="#FAFAFA", fg=RENK_METIN, relief="flat",
            wrap="word", padx=8, pady=8, state=DISABLED,
        )
        kaydirma = Scrollbar(log_alan, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=kaydirma.set)
        self.log_text.pack(side="left", fill="both", expand=True)
        kaydirma.pack(side="right", fill="y")

        self.log_text.tag_configure("basari", foreground=RENK_BASARI, font=("Consolas", 9, "bold"))
        self.log_text.tag_configure("hata", foreground=RENK_HATA, font=("Consolas", 9, "bold"))
        self.log_text.tag_configure("uyari", foreground=RENK_UYARI)
        self.log_text.tag_configure("bilgi", foreground=RENK_IKINCIL_METIN)
        self.log_text.tag_configure("fatura", foreground="#7B1FA2", font=("Consolas", 9, "bold"))
        self.log_text.tag_configure("llm", foreground=RENK_LLM)
        self.log_text.tag_configure("sutun", foreground=RENK_SUTUN, font=("Consolas", 9, "bold"))

        # Alt bilgi ve kontroller
        alt = Frame(self.root, height=90)
        alt.pack(fill="x", side="bottom")
        alt.pack_propagate(False)

        # Sol taraf: Logo
        logo_path = Path(__file__).resolve().parent / "asset" / "Iski-logo.png"
        if PIL_KULLANILABILIR and logo_path.exists():
            try:
                logo_img = Image.open(str(logo_path))
                logo_img.thumbnail((60, 60), Image.Resampling.LANCZOS)
                self.logo_image = PILPhotoImage(image=logo_img)
                Label(
                    alt, image=self.logo_image,
                ).pack(side="left", padx=15, pady=12)
            except Exception:
                pass

        # Orta taraf: Durum bilgileri
        durum_frame = Frame(alt)
        durum_frame.pack(side="left", padx=10)

        # LLM durumu (ikon + metin yan yana)
        llm_satir = Frame(durum_frame)
        llm_satir.pack(anchor="w")
        llm_ikon = self._ikon_yukle(
            "check.png" if LLM_KULLANILABILIR else "warning.png", 14
        )
        if llm_ikon:
            Label(llm_satir, image=llm_ikon).pack(side="left", padx=(0, 4))
        Label(
            llm_satir,
            text="LLM destekli" if LLM_KULLANILABILIR else "LLM pasif",
            font=("Segoe UI", 8), fg=RENK_BIRINCIL,
        ).pack(side="left")

        # OCR durumu
        ocr_satir = Frame(durum_frame)
        ocr_satir.pack(anchor="w")
        ocr_ikon = self._ikon_yukle(
            "check.png" if OCR_KULLANILABILIR else "warning.png", 14
        )
        if ocr_ikon:
            Label(ocr_satir, image=ocr_ikon).pack(side="left", padx=(0, 4))
        Label(
            ocr_satir,
            text="OCR destekli" if OCR_KULLANILABILIR else "OCR pasif",
            font=("Segoe UI", 8), fg=RENK_BIRINCIL,
        ).pack(side="left")

        # Sağ taraf: Butonlar (Sütunlar + Ayarlar)
        # Önce Ayarlar (sağa yapışsın), sonra Sütunlar (onun solu)
        Button(
            alt, text="Ayarlar",
            font=self.f_buton_kucuk, bg=RENK_IKINCIL_METIN, fg="white",
            relief="flat", padx=14, pady=8, cursor="hand2",
            command=self._ayarlar_penceresi,
        ).pack(side="right", padx=(0, 15), pady=12)

        # Sütun Tanımla butonu (table.png ikonlu)
        sutun_btn = Button(
            alt, text=" Sütun Tanımla",
            font=self.f_buton_kucuk, bg=RENK_SUTUN, fg="white",
            relief="flat", padx=14, pady=8, cursor="hand2",
            command=self._sutun_penceresi,
        )
        sutun_btn.pack(side="right", padx=(0, 8), pady=12)

    def _excel_karti_olustur(self, parent, baslik, degisken):
        kart = Frame(parent, bg=RENK_KART, padx=12, pady=10,
                     highlightbackground="#E0E0E0", highlightthickness=1)
        kart.pack(fill="x", expand=True)

        Label(kart, text=baslik,
              font=("Segoe UI", 10, "bold"),
              bg=RENK_KART, fg=RENK_METIN).pack(anchor="w")

        satir = Frame(kart, bg=RENK_KART)
        satir.pack(fill="x", pady=(6, 0))

        Label(
            satir, textvariable=degisken,
            font=("Segoe UI", 8), bg="#F5F5F5", fg=RENK_IKINCIL_METIN,
            anchor="w", padx=8, pady=4,
        ).pack(side="left", fill="x", expand=True)

        Button(
            satir, text="...",
            font=self.f_buton_kucuk, bg=RENK_VURGU, fg="white",
            relief="flat", padx=10, pady=2, cursor="hand2",
            command=lambda: self._excel_sec(degisken),
        ).pack(side="right", padx=(6, 0))

    def _baslangic_kontrolleri(self):
        self._log(
            "Hazır. PDF seçin — kalem tablosu Excel'e yeni bir sayfa olarak eklenecek.",
            "bilgi"
        )

        # Sütun durumunu kontrol et
        self._sutun_durumu_guncelle()

        fatura_excel = Path(self.fatura_excel_yolu.get())
        if not fatura_excel.exists():
            self._log(
                f"ℹ Excel dosyası henüz yok: {fatura_excel.name}\n"
                "   İşleme başladığınızda otomatik oluşturulacak.",
                "bilgi"
            )

        # LLM durumunu kontrol et
        if LLM_KULLANILABILIR:
            durumlar = ayarlar.hangi_saglayicilar_hazir()
            hazir = [a for a, h in durumlar.items() if h]
            if hazir:
                self._log(f"LLM API anahtarları hazır: {', '.join(hazir)}", "llm")
                self.llm_secim.set(hazir[0])
                self._llm_secim_degisti(None)
            else:
                self._log("ℹ LLM API anahtarı tanımlı değil. Ayarlar'dan ekleyebilirsiniz.", "bilgi")
        else:
            self._log("⚠ LLM kütüphaneleri kurulu değil.", "uyari")

    def _sutun_durumu_guncelle(self):
        """Üstteki sütun göstergesini günceller."""
        sutunlar = ayarlar.kalem_sutunlarini_al()
        if sutunlar:
            self.sutun_durum.set(f" {len(sutunlar)} sütun tanımlı")
            self._log(f"Tanımlı sütunlar: {' → '.join(sutunlar)}", "sutun")
        else:
            self.sutun_durum.set("⚠ Sütun tanımlı değil")
            self._log(
                "⚠ Sütun tanımlı değil. 'Sütunlar' butonundan sütun isimlerini girin.",
                "uyari"
            )

    def _llm_secim_degisti(self, event):
        secim = self.llm_secim.get()

        if secim == "kapali":
            self._aktif_saglayici = None
            self.llm_durum.set("LLM: Kapalı (regex yedek modu)")
            self._log("LLM kapatıldı. Regex tablo çıkarımı kullanılacak.", "bilgi")
            return

        if not LLM_KULLANILABILIR:
            self.llm_secim.set("kapali")
            self.llm_durum.set("LLM: Kütüphaneler kurulu değil")
            messagebox.showerror(
                "LLM Kullanılamıyor",
                "LLM kütüphaneleri kurulu değil.\n\n"
                "Komut satırında çalıştırın:\n"
                "  pip install anthropic google-generativeai python-dotenv"
            )
            return

        try:
            self._aktif_saglayici = saglayici_olustur(secim)
            self.llm_durum.set(f"LLM: {secim} / {self._aktif_saglayici.model}")
            self._log(f"✓ LLM aktif: {secim} ({self._aktif_saglayici.model})", "llm")
        except RuntimeError:
            self.llm_secim.set("kapali")
            self._aktif_saglayici = None
            self.llm_durum.set("LLM: API anahtarı yok")
            cevap = messagebox.askyesno(
                "API Anahtarı Gerekli",
                f"{secim} için API anahtarı tanımlı değil.\n\n"
                f"Ayarlar penceresini açıp anahtarı girmek ister misiniz?"
            )
            if cevap:
                self._ayarlar_penceresi()
        except Exception as e:
            self.llm_secim.set("kapali")
            self._aktif_saglayici = None
            self.llm_durum.set("LLM: Hata")
            messagebox.showerror("LLM Hatası", f"LLM kurulamadı:\n\n{e}")

   
    # SÜTUN PENCERESİ (YENİ)
   
    def _sutun_penceresi(self):
        """Sütun isimlerini düzenlemek için pencere açar."""
        pencere = Toplevel(self.root)
        pencere.title("Kalem Sütunları")
        pencere.geometry("520x600")
        pencere.configure(bg=RENK_BG)
        pencere.transient(self.root)
        pencere.grab_set()

        pencere.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() - 520) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - 600) // 2
        pencere.geometry(f"520x600+{x}+{y}")

        # Başlık
        Label(pencere, text=" Kalem Sütunları",
              font=("Segoe UI", 13, "bold"),
              bg=RENK_BG, fg=RENK_METIN).pack(pady=(15, 5))

        Label(
            pencere,
            text="Excel'de kalemlerin yazılacağı sütun isimlerini sırasıyla girin.\n"
                 "Bu sıra Exel'de yerleştirilen kalem sıraları olacak.",
            font=("Segoe UI", 8), bg=RENK_BG, fg=RENK_IKINCIL_METIN,
            justify="center",
        ).pack(pady=(0, 10))

        # Ekleme satırı
        ekleme_kart = Frame(pencere, bg=RENK_KART, padx=12, pady=10,
                             highlightbackground="#E0E0E0", highlightthickness=1)
        ekleme_kart.pack(fill="x", padx=20, pady=(0, 10))

        Label(ekleme_kart, text="Yeni sütun adı:",
              font=("Segoe UI", 9), bg=RENK_KART, fg=RENK_METIN
              ).pack(side="left", padx=(0, 8))

        yeni_sutun_var = StringVar()
        yeni_entry = Entry(ekleme_kart, textvariable=yeni_sutun_var,
                           font=("Segoe UI", 10), width=25)
        yeni_entry.pack(side="left", fill="x", expand=True, padx=(0, 8))
        yeni_entry.focus_set()

        # Liste alanı (scrollable)
        liste_cerceve = Frame(pencere, bg=RENK_KART,
                              highlightbackground="#E0E0E0", highlightthickness=1)
        liste_cerceve.pack(fill="both", expand=True, padx=20, pady=(0, 10))

        liste_baslik = Frame(liste_cerceve, bg=RENK_KART, padx=12, pady=8)
        liste_baslik.pack(fill="x")
        Label(liste_baslik, text="Tanımlı Sütunlar (sıra önemli):",
              font=("Segoe UI", 10, "bold"),
              bg=RENK_KART, fg=RENK_METIN).pack(side="left")

        # Kaydırılabilir liste
        canvas_cerceve = Frame(liste_cerceve, bg=RENK_KART)
        canvas_cerceve.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        from tkinter import Canvas
        canvas = Canvas(canvas_cerceve, bg=RENK_KART, highlightthickness=0)
        kaydirma_y = Scrollbar(canvas_cerceve, orient="vertical", command=canvas.yview)
        ic_cerceve = Frame(canvas, bg=RENK_KART)

        ic_cerceve.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        canvas.create_window((0, 0), window=ic_cerceve, anchor="nw")
        canvas.configure(yscrollcommand=kaydirma_y.set)
        canvas.pack(side="left", fill="both", expand=True)
        kaydirma_y.pack(side="right", fill="y")

        # Mouse wheel için
        def on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        canvas.bind_all("<MouseWheel>", on_mousewheel)

        # Sütun listesi state
        sutunlar = ayarlar.kalem_sutunlarini_al()
        # Liste boşsa varsayılan öneriler (kullanıcı silebilir)
        if not sutunlar:
            sutunlar = []

        def listeyi_yeniden_ciz():
            """İç çerçeveyi temizleyip sütunları yeniden çizer."""
            for w in ic_cerceve.winfo_children():
                w.destroy()

            if not sutunlar:
                Label(
                    ic_cerceve,
                    text="Henüz sütun yok. Yukarıdaki kutudan ekleyin.",
                    font=("Segoe UI", 9, "italic"),
                    bg=RENK_KART, fg=RENK_IKINCIL_METIN,
                ).pack(pady=20)
                return

            for i, ad in enumerate(sutunlar):
                satir = Frame(ic_cerceve, bg="#FAFAFA", padx=8, pady=6)
                satir.pack(fill="x", pady=2)

                # Sıra numarası
                Label(satir, text=f"{i+1}.", font=("Segoe UI", 9, "bold"),
                      bg="#FAFAFA", fg=RENK_BIRINCIL, width=3
                      ).pack(side="left")

                # Sütun adı (düzenlenebilir)
                ad_var = StringVar(value=ad)
                ad_entry = Entry(satir, textvariable=ad_var,
                                  font=("Segoe UI", 10), bg="white",
                                  relief="flat", bd=1)
                ad_entry.pack(side="left", fill="x", expand=True, padx=(0, 8))

                # Değişiklik olunca listeyi güncelle
                def degisiklik_olunca(idx=i, var=ad_var):
                    sutunlar[idx] = var.get().strip()
                ad_var.trace_add("write", lambda *a, idx=i, var=ad_var: degisiklik_olunca(idx, var))

                # ↑ Yukarı
                def yukari(idx=i):
                    if idx > 0:
                        sutunlar[idx], sutunlar[idx-1] = sutunlar[idx-1], sutunlar[idx]
                        listeyi_yeniden_ciz()
                Button(satir, text="↑", font=("Segoe UI", 9),
                       bg="#E0E0E0", relief="flat", cursor="hand2",
                       width=2, command=yukari).pack(side="left", padx=1)

                # ↓ Aşağı
                def asagi(idx=i):
                    if idx < len(sutunlar) - 1:
                        sutunlar[idx], sutunlar[idx+1] = sutunlar[idx+1], sutunlar[idx]
                        listeyi_yeniden_ciz()
                Button(satir, text="↓", font=("Segoe UI", 9),
                       bg="#E0E0E0", relief="flat", cursor="hand2",
                       width=2, command=asagi).pack(side="left", padx=1)

                # Sil
                def sil(idx=i):
                    sutunlar.pop(idx)
                    listeyi_yeniden_ciz()
                sil_ikon = self._ikon_yukle("x.png", 18)
                if sil_ikon:
                    Button(satir, image=sil_ikon,
                           bg=RENK_HATA, relief="flat", cursor="hand2",
                           command=sil).pack(side="left", padx=(4, 0))
                else:
                    Button(satir, text="✕", font=("Segoe UI", 9, "bold"),
                           bg=RENK_HATA, fg="white", relief="flat", cursor="hand2",
                           width=2, command=sil).pack(side="left", padx=(4, 0))

        def yeni_ekle(event=None):
            ad = yeni_sutun_var.get().strip()
            if not ad:
                return
            if ad in sutunlar:
                messagebox.showwarning("Tekrar", f"'{ad}' zaten listede.", parent=pencere)
                return
            sutunlar.append(ad)
            yeni_sutun_var.set("")
            listeyi_yeniden_ciz()
            # Otomatik en alta kaydır
            pencere.after(100, lambda: canvas.yview_moveto(1.0))

        Button(ekleme_kart, text="+ Ekle",
               font=("Segoe UI", 9, "bold"),
               bg=RENK_BASARI, fg="white",
               relief="flat", padx=12, pady=4, cursor="hand2",
               command=yeni_ekle).pack(side="left")

        # Enter tuşu ile de ekle
        yeni_entry.bind("<Return>", yeni_ekle)

        # İlk çizim
        listeyi_yeniden_ciz()

        # Alt butonlar
        butonlar = Frame(pencere, bg=RENK_BG)
        butonlar.pack(side="bottom", pady=15)

        def kaydet():
            # Boşları at, baş/son boşlukları kırp
            temiz = [s.strip() for s in sutunlar if s.strip()]
            if not temiz:
                cevap = messagebox.askyesno(
                    "Boş Liste",
                    "Sütun listesi boş. Boş kayıt yapılsın mı?\n"
                    "(Boş liste programın çalışmamasına yol açar.)",
                    parent=pencere,
                )
                if not cevap:
                    return
            try:
                ayarlar.kalem_sutunlarini_kaydet(temiz)
                messagebox.showinfo(
                    "Kaydedildi",
                    f"{len(temiz)} sütun kaydedildi.",
                    parent=pencere,
                )
                pencere.destroy()
                self._sutun_durumu_guncelle()
            except Exception as e:
                messagebox.showerror("Hata", f"Kaydetme başarısız:\n{e}", parent=pencere)

        Button(butonlar, text="Kaydet",
               font=self.f_buton, bg=RENK_BASARI, fg="white",
               relief="flat", padx=20, pady=8, cursor="hand2",
               command=kaydet).pack(side="left", padx=5)

        Button(butonlar, text="İptal",
               font=self.f_buton, bg="#9E9E9E", fg="white",
               relief="flat", padx=20, pady=8, cursor="hand2",
               command=pencere.destroy).pack(side="left", padx=5)

        # Pencere kapanırken mouse wheel bind'ını temizle
        def on_close():
            canvas.unbind_all("<MouseWheel>")
            pencere.destroy()
        pencere.protocol("WM_DELETE_WINDOW", on_close)

   
    # AYARLAR PENCERESİ (API key)
   
    def _ayarlar_penceresi(self):
        if not LLM_KULLANILABILIR:
            messagebox.showwarning(
                "Kütüphaneler Eksik",
                "LLM kütüphaneleri kurulu değil.\n\n"
                "pip install anthropic google-generativeai python-dotenv"
            )
            return

        pencere = Toplevel(self.root)
        pencere.title("Ayarlar - API Anahtarları")
        pencere.geometry("520x300")
        pencere.configure(bg=RENK_BG)
        pencere.transient(self.root)
        pencere.grab_set()

        pencere.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() - 520) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - 300) // 2
        pencere.geometry(f"520x300+{x}+{y}")

        Label(pencere, text="LLM API Anahtarları",
              font=("Segoe UI", 13, "bold"),
              bg=RENK_BG, fg=RENK_METIN).pack(pady=(20, 5))

        Label(pencere,
              text="Anahtarlar proje kökündeki .env dosyasına kaydedilir.",
              font=("Segoe UI", 8), bg=RENK_BG, fg=RENK_IKINCIL_METIN).pack(pady=(0, 15))

        mevcut = {
            "claude": ayarlar.api_anahtari_al("claude") or "",
            "gemini": ayarlar.api_anahtari_al("gemini") or "",
        }

        kayit_alanlari = {}
        for saglayici_adi in ["claude", "gemini"]:
            satir = Frame(pencere, bg=RENK_BG, padx=20)
            satir.pack(fill="x", pady=8)

            Label(satir, text=f"{saglayici_adi.capitalize()} API Key:",
                  font=("Segoe UI", 10), bg=RENK_BG, fg=RENK_METIN,
                  width=18, anchor="w").pack(side="left")

            entry = Entry(satir, font=("Consolas", 9), show="•", width=40)
            entry.insert(0, mevcut[saglayici_adi])
            entry.pack(side="left", fill="x", expand=True, padx=(0, 5))
            kayit_alanlari[saglayici_adi] = entry

            def goster_gizle(e=entry):
                e.config(show="" if e.cget("show") == "•" else "•")

            Button(satir, text="👁", font=("Segoe UI", 9),
                   bg="#EEEEEE", relief="flat", cursor="hand2", width=3,
                   command=goster_gizle).pack(side="right")

        Label(
            pencere,
            text="Anahtar almak için: console.anthropic.com / aistudio.google.com",
            font=("Segoe UI", 8), bg=RENK_BG, fg=RENK_IKINCIL_METIN,
        ).pack(pady=(15, 0))

        butonlar = Frame(pencere, bg=RENK_BG)
        butonlar.pack(side="bottom", pady=15)

        def kaydet():
            try:
                for saglayici_adi, entry in kayit_alanlari.items():
                    deger = entry.get().strip()
                    if deger:
                        ayarlar.api_anahtari_kaydet(saglayici_adi, deger)
                messagebox.showinfo("Başarılı", "API anahtarları kaydedildi.")
                pencere.destroy()
                if self.llm_secim.get() != "kapali":
                    self._llm_secim_degisti(None)
                self._log("API anahtarları güncellendi.", "llm")
            except Exception as e:
                messagebox.showerror("Hata", f"Kaydetme başarısız:\n{e}")

        Button(butonlar, text="Kaydet",
               font=self.f_buton, bg=RENK_BASARI, fg="white",
               relief="flat", padx=20, pady=8, cursor="hand2",
               command=kaydet).pack(side="left", padx=5)

        Button(butonlar, text="İptal",
               font=self.f_buton, bg="#9E9E9E", fg="white",
               relief="flat", padx=20, pady=8, cursor="hand2",
               command=pencere.destroy).pack(side="left", padx=5)

   
    # İŞLEM AKIŞI
   
    def _excel_sec(self, degisken: StringVar):
        yol = filedialog.askopenfilename(
            title="Excel dosyasını seçin",
            filetypes=[("Excel dosyaları", "*.xlsx *.xlsm"), ("Tümü", "*.*")],
        )
        if yol:
            degisken.set(yol)
            self._log(f"Excel değiştirildi: {Path(yol).name}", "bilgi")

    def _pdf_sec_ve_isle(self):
        
        yollar = filedialog.askopenfilenames(
            title="Fatura PDF dosyasını/dosyalarını seçin "
                  "(birden fazla seçmek için Ctrl ile tıklayın)",
            filetypes=[("PDF dosyaları", "*.pdf"), ("Tümü", "*.*")],
        )
        if not yollar:
            return
        # alfabetik sırala
        pdf_listesi = sorted([Path(y) for y in yollar], key=lambda p: p.name.lower())
        self._isle_arka_planda(pdf_listesi=pdf_listesi)

    def _isle_arka_planda(self, pdf_listesi: list = None):
        fatura_excel = Path(self.fatura_excel_yolu.get())

        # Sütun kontrolü
        sutunlar = ayarlar.kalem_sutunlarini_al()
        if not sutunlar:
            messagebox.showerror(
                "Sütun Yok",
                "Kalem sütunları tanımlı değil.\n\n"
                "Lütfen önce ' Sütunlar' butonundan sütun isimlerini girin."
            )
            return

        if not pdf_listesi:
            return

        self._butonlari_kilitle(True)
        self.ilerleme.start(10)

        saglayici = self._aktif_saglayici
        coklu_mod = len(pdf_listesi) > 1

        def calistir():
            try:
                if not coklu_mod:
                    # Tek PDF: kendi sayfası
                    pdf = pdf_listesi[0]
                    self._log(f"▶ İşleniyor: {pdf.name}", "bilgi")
                    sonuc = faturayi_isle(
                        pdf,
                        excel_yolu=fatura_excel,
                        saglayici=saglayici,
                        sutunlar=sutunlar,
                        toplu_mod=False,
                    )
                    sonuc["pdf"] = pdf.name
                    self._sonucu_logla(pdf.name, sonuc)
                else:
                    # Çoklu PDF: hepsi tek sayfaya alt alta
                    self._log(
                        f"▶ {len(pdf_listesi)} PDF işleniyor (hepsi tek sayfada toplanacak)",
                        "bilgi"
                    )
                    sonuclar = []
                    hedef_sayfa_adi = None
                    for pdf in pdf_listesi:
                        self._log(f"  ▶ {pdf.name}", "bilgi")
                        sonuc = faturayi_isle(
                            pdf,
                            excel_yolu=fatura_excel,
                            saglayici=saglayici,
                            sutunlar=sutunlar,
                            toplu_mod=True,
                            hedef_sayfa=hedef_sayfa_adi,
                        )
                        sonuc["pdf"] = pdf.name
                        sonuclar.append(sonuc)
                        self._sonucu_logla(pdf.name, sonuc)
                        # İlk başarılı işlemden sonra sayfa adını kilitle
                        if sonuc.get("basarili") and not hedef_sayfa_adi:
                            hedef_sayfa_adi = sonuc.get("sayfa_adi")

                    basarili = sum(1 for s in sonuclar if s["basarili"])
                    toplam_kalem = sum(
                        s.get("kalem_sayisi", 0) for s in sonuclar if s["basarili"]
                    )
                    sayfa_adi = next(
                        (s.get("sayfa_adi") for s in sonuclar if s.get("basarili")),
                        "?"
                    )
                    self._log(
                        f"İşlemler -> {basarili}/{len(sonuclar)} başarılı, "
                        f"toplam {toplam_kalem} kalem '{sayfa_adi}' sayfasına yazıldı",
                        "basari" if basarili == len(sonuclar) else "uyari"
                    )
            except Exception as e:
                self._log(f"✗ Beklenmeyen hata: {e}", "hata")
            finally:
                self.root.after(0, lambda: (
                    self.ilerleme.stop(),
                    self._butonlari_kilitle(False),
                ))

        threading.Thread(target=calistir, daemon=True).start()

    def _sonucu_logla(self, pdf_adi: str, sonuc: dict):
        if not sonuc.get("basarili"):
            self._log(f"✗ {pdf_adi}  →  HATA: {sonuc.get('hata', 'bilinmeyen')}", "hata")
            return

        sayfa = sonuc.get("sayfa_adi", "?")
        kalem_sayisi = sonuc.get("kalem_sayisi", 0)
        llm_kullanildi = sonuc.get("llm_kullanildi", False)
        llm_hata = sonuc.get("llm_hata", "")
        ocr_kullanildi = sonuc.get("ocr_kullanildi", False)

        # LLM denendi ama başarısız oldu mu?
        if llm_hata:
            self._log(f"⚠ LLM başarısız: {llm_hata}", "uyari")
            self._log("   → Regex yedek devreye girdi", "bilgi")

        etiketler = []
        etiketler.append("LLM" if llm_kullanildi else "regex")
        if ocr_kullanildi:
            etiketler.append("OCR")
        etiket = f" [{' + '.join(etiketler)}]"

        self._log(
            f"✓ {pdf_adi}  →  '{sayfa}' sayfasına {kalem_sayisi} kalem yazıldı{etiket}",
            "fatura"
        )

    def _log(self, metin: str, tag: str = "bilgi"):
        def yaz():
            self.log_text.configure(state=NORMAL)
            self.log_text.insert(END, metin + "\n", tag)
            self.log_text.see(END)
            self.log_text.configure(state=DISABLED)
        self.root.after(0, yaz)

    def _log_temizle(self):
        self.log_text.configure(state=NORMAL)
        self.log_text.delete("1.0", END)
        self.log_text.configure(state=DISABLED)

    def _butonlari_kilitle(self, kilitli: bool):
        durum = DISABLED if kilitli else NORMAL
        self.pdf_buton.configure(state=durum)


def main():
    root = Tk()
    app = OtomasyonApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()