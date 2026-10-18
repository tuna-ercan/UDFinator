# -*- coding: utf-8 -*-
"""
UDFinator — basit arayuz (Turkce).

Buyukanne icin olabildigince sade. Dort secenek:
    * UDF  -> Word (.docx)
    * UDF  -> PDF
    * Word -> UDF
    * PDF  -> UDF

Her dugme bir dosya sectirir, donusturur ve sonucu ayni klasore kaydeder.

Arka plan resmi: image.jpg derleme sirasinda exe'nin icine gomulur. Exe'nin
yanina image.jpg (ya da arkaplan.jpg / background.jpg) koyulursa o kullanilir;
boylece resmi degistirmek icin yeniden derlemek gerekmez. Resim yoksa program
duz renkli acilir.
"""

import os
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox

import converters

try:
    from PIL import Image, ImageDraw, ImageTk
    _HAS_PIL = True
except Exception:
    _HAS_PIL = False


# Kucuk ekranlara da sigsin (1366x768 ve %150 olcekli 1080p dahil)
WINDOW_W = 580
WINDOW_H = 620

HEADER_H = 100
CARD_Y0 = 125
CARD_Y1 = WINDOW_H - 20
QUESTION_Y = CARD_Y0 + 35
BTN_Y0 = CARD_Y0 + 100
BTN_STEP = 72
BTN_COUNT = 4
STATUS_Y = BTN_Y0 + BTN_COUNT * BTN_STEP + 4
HINT_Y = CARD_Y1 - 22

BG_COLOR = "#2b3a67"
HEADER_COLOR = "#243156"
CARD_COLOR = "#f4f6fb"
TEXT_DARK = "#1b1b1b"
TEXT_MUTED = "#555555"

# Dugme renkleri
BLUE = "#1f6feb"
BLUE_H = "#1a5fd0"
RED = "#c0392b"
RED_H = "#a93224"
PURPLE = "#7d3c98"
PURPLE_H = "#6c3483"
GREEN = "#1e8449"
GREEN_H = "#176c3c"

BG_NAMES = [
    "image.jpg", "image.jpeg", "image.png",
    "arkaplan.jpg", "arkaplan.jpeg", "arkaplan.png",
    "background.jpg", "background.jpeg", "background.png",
]


def app_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def unique_path(path):
    """Ayni adli dosya varsa uzerine yazmak yerine 'ad (1).uzanti' gibi yeni bir
    ad bulur; boylece daha once duzenlenmis bir belge kaybolmaz."""
    if not os.path.exists(path):
        return path
    base, ext = os.path.splitext(path)
    n = 1
    while os.path.exists("%s (%d)%s" % (base, n, ext)):
        n += 1
    return "%s (%d)%s" % (base, n, ext)


def find_background():
    """Once exe'nin yanindaki resme (degistirmek kolay olsun), sonra exe'nin
    icine gomulu resme bakar."""
    places = [app_dir()]
    bundled = getattr(sys, "_MEIPASS", None)
    if bundled:
        places.append(bundled)
    for base in places:
        for name in BG_NAMES:
            path = os.path.join(base, name)
            if os.path.exists(path):
                return path
    return None


class ConverterApp:
    def __init__(self, root):
        self.root = root
        root.title("UDFinator")
        root.geometry("%dx%d" % (WINDOW_W, WINDOW_H))
        root.resizable(False, False)
        root.configure(bg=BG_COLOR)
        self._center()

        self._bg_photo = None
        self._buttons = []
        self._busy = False

        self.canvas = tk.Canvas(
            root, width=WINDOW_W, height=WINDOW_H,
            highlightthickness=0, bg=BG_COLOR,
        )
        self.canvas.pack(fill="both", expand=True)

        self.has_photo = self._draw_background()
        if self.has_photo:
            self.fg_main, self.fg_muted = "white", "#eeeeee"
            self.fg_ok, self.fg_err = "#8ef0b6", "#ffb3b3"
        else:
            self.fg_main, self.fg_muted = TEXT_DARK, TEXT_MUTED
            self.fg_ok, self.fg_err = "#1a7f37", "#b42318"
        self._build_ui()

    def _center(self):
        self.root.update_idletasks()
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        x = max(0, (sw - WINDOW_W) // 2)
        y = max(0, (sh - WINDOW_H) // 3)
        self.root.geometry("+%d+%d" % (x, y))

    def _draw_background(self):
        """Fotograf varsa cizer ve True dondurur; yoksa duz tasarimi cizer."""
        bg_path = find_background()
        if bg_path and _HAS_PIL:
            try:
                img = Image.open(bg_path).convert("RGB")
                img = self._cover(img, WINDOW_W, WINDOW_H)
                img = self._add_text_panels(img)
                self._bg_photo = ImageTk.PhotoImage(img)
                self.canvas.create_image(0, 0, anchor="nw", image=self._bg_photo)
                return True
            except Exception:
                self._bg_photo = None
        self.canvas.create_rectangle(0, 0, WINDOW_W, HEADER_H, fill=HEADER_COLOR, outline="")
        self.canvas.create_rectangle(40, CARD_Y0, WINDOW_W - 40, CARD_Y1,
                                     fill=CARD_COLOR, outline="")
        return False

    @staticmethod
    def _cover(img, target_w, target_h):
        """Resmi pencereyi tamamen kaplayacak sekilde olcekler ve kirpar."""
        src_w, src_h = img.size
        scale = max(target_w / src_w, target_h / src_h)
        new_w, new_h = int(src_w * scale), int(src_h * scale)
        img = img.resize((new_w, new_h), Image.LANCZOS)
        left = (new_w - target_w) // 2
        top = (new_h - target_h) // 2
        return img.crop((left, top, left + target_w, top + target_h))

    @staticmethod
    def _add_text_panels(img):
        """Yazilarin arkasina yari saydam koyu paneller ekler; fotografin geri
        kalani parlak kalir, beyaz yazi her fotografta okunur."""
        overlay = Image.new("RGBA", img.size, (0, 0, 0, 0))
        draw = ImageDraw.Draw(overlay)
        draw.rectangle((0, 0, WINDOW_W, HEADER_H), fill=(0, 0, 0, 120))
        draw.rounded_rectangle((70, QUESTION_Y - 23, WINDOW_W - 70, QUESTION_Y + 23),
                               radius=14, fill=(0, 0, 0, 135))
        draw.rounded_rectangle((60, STATUS_Y - 24, WINDOW_W - 60, CARD_Y1 - 6),
                               radius=14, fill=(0, 0, 0, 135))
        return Image.alpha_composite(img.convert("RGBA"), overlay).convert("RGB")

    def _build_ui(self):
        cx = WINDOW_W // 2

        self.canvas.create_text(
            cx, HEADER_H // 2, text="UDFinator",
            font=("Segoe UI", 24, "bold"), fill="white",
        )
        self.canvas.create_text(
            cx, QUESTION_Y, text="Ne yapmak istiyorsunuz?",
            font=("Segoe UI", 16, "bold"), fill=self.fg_main,
        )

        specs = [
            ("📄  UDF → Word (.docx)", BLUE, BLUE_H, self.do_udf_to_docx),
            ("📕  UDF → PDF", RED, RED_H, self.do_udf_to_pdf),
            ("📝  Word (.docx) → UDF", PURPLE, PURPLE_H, self.do_docx_to_udf),
            ("🔄  PDF → UDF", GREEN, GREEN_H, self.do_pdf_to_udf),
        ]
        y = BTN_Y0
        for text, color, hover, command in specs:
            btn = tk.Button(
                self.root, text=text,
                font=("Segoe UI", 15, "bold"),
                bg=color, fg="white",
                activebackground=hover, activeforeground="white",
                relief="flat", bd=0, cursor="hand2",
                padx=16, pady=8, command=command,
            )
            self.canvas.create_window(cx, y, window=btn, width=380)
            btn.bind("<Enter>", lambda e, b=btn, h=hover: self._hover(b, h))
            btn.bind("<Leave>", lambda e, b=btn, c=color: self._hover(b, c))
            self._buttons.append((btn, color))
            y += BTN_STEP

        # Durum ve ipucu: tuval yazisi (arka plani saydam, fotografi kapatmaz)
        self.status_item = self.canvas.create_text(
            cx, STATUS_Y, text="Hazır.",
            font=("Segoe UI", 12), fill=self.fg_muted,
            width=420, justify="center",
        )
        self.canvas.create_text(
            cx, HINT_Y, text="Yeni dosya, kaynak dosyanın yanına kaydedilir.",
            font=("Segoe UI", 10), fill=self.fg_muted, justify="center",
        )

    def _status(self, text, color):
        self.canvas.itemconfig(self.status_item, text=text, fill=color)

    def _hover(self, btn, color):
        if not self._busy:
            btn.config(bg=color)

    def _set_busy(self, busy):
        self._busy = busy
        state = "disabled" if busy else "normal"
        for btn, color in self._buttons:
            btn.config(state=state, bg=color)

    # --- Donusum islemleri -------------------------------------------------

    def do_udf_to_docx(self):
        self._run(
            title="UDF dosyası seçin",
            filetypes=[("UDF dosyaları", "*.udf"), ("Tüm dosyalar", "*.*")],
            out_ext=".docx", fn=converters.udf_to_docx,
            out_name="Word belgesi",
        )

    def do_udf_to_pdf(self):
        self._run(
            title="UDF dosyası seçin",
            filetypes=[("UDF dosyaları", "*.udf"), ("Tüm dosyalar", "*.*")],
            out_ext=".pdf", fn=converters.udf_to_pdf,
            out_name="PDF dosyası",
        )

    def do_docx_to_udf(self):
        self._run(
            title="Word dosyası seçin",
            filetypes=[("Word dosyaları", "*.docx"), ("Tüm dosyalar", "*.*")],
            out_ext=".udf", fn=converters.docx_to_udf,
            out_name="UDF dosyası",
        )

    def do_pdf_to_udf(self):
        self._run(
            title="PDF dosyası seçin",
            filetypes=[("PDF dosyaları", "*.pdf"), ("Tüm dosyalar", "*.*")],
            out_ext=".udf", fn=converters.pdf_to_udf,
            out_name="UDF dosyası",
        )

    def _run(self, title, filetypes, out_ext, fn, out_name):
        if self._busy:
            return
        path = filedialog.askopenfilename(title=title, filetypes=filetypes)
        if not path:
            return
        self._set_busy(True)
        self._status("Dönüştürülüyor, lütfen bekleyin...", self.fg_muted)
        threading.Thread(
            target=self._worker, args=(path, out_ext, fn, out_name), daemon=True
        ).start()

    def _worker(self, path, out_ext, fn, out_name):
        out_path = unique_path(os.path.splitext(path)[0] + out_ext)
        try:
            fn(path, out_path)
        except converters.UdfError as exc:
            self.root.after(0, self._on_error, str(exc))
        except Exception as exc:
            self.root.after(0, self._on_error, "Beklenmeyen hata: %s" % exc)
        else:
            self.root.after(0, self._on_success, out_path, out_name)

    def _on_success(self, out_path, out_name):
        self._set_busy(False)
        self._status("Tamamlandı! ✅", self.fg_ok)
        folder = os.path.dirname(out_path)
        name = os.path.basename(out_path)
        if messagebox.askyesno(
            "Başarılı",
            "%s hazır:\n\n%s\n\nDosyanın bulunduğu klasörü açmak ister misiniz?"
            % (out_name, name),
        ):
            try:
                os.startfile(folder)
            except Exception:
                pass

    def _on_error(self, message):
        self._set_busy(False)
        self._status("Dönüştürülemedi.", self.fg_err)
        messagebox.showerror("Hata", "Dosya dönüştürülemedi.\n\n%s" % message)


def _cli_convert(argv):
    """Gizli komut satiri modu: UDFinator.exe --convert MOD GIRDI CIKTI
    (MOD: udf2docx | udf2pdf | docx2udf | pdf2udf). Basari: cikis kodu 0."""
    modes = {
        "udf2docx": converters.udf_to_docx,
        "udf2pdf": converters.udf_to_pdf,
        "docx2udf": converters.docx_to_udf,
        "pdf2udf": converters.pdf_to_udf,
    }
    if len(argv) != 4 or argv[1] not in modes:
        return 2
    try:
        modes[argv[1]](argv[2], argv[3])
    except Exception:
        return 1
    return 0


def main():
    if len(sys.argv) > 1 and sys.argv[1] == "--convert":
        sys.exit(_cli_convert(sys.argv[1:]))
    root = tk.Tk()
    ConverterApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
