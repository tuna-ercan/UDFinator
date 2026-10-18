# -*- coding: utf-8 -*-
"""
UDFinator — basit arayuz (Turkce).

Buyukanne icin olabildigince sade. Uc secenek:
    * UDF  -> Word (.docx)
    * UDF  -> PDF
    * PDF  -> UDF

Her dugme bir dosya sectirir, donusturur ve sonucu ayni klasore kaydeder.

Arka plan resmi: Programin (exe'nin) yanina su adlardan biriyle bir resim
koyarsaniz otomatik arka plan olur:
    arkaplan.png / arkaplan.jpg / arkaplan.jpeg / background.png / background.jpg
Simdilik resim koymadik; program duz renkli acilir.
"""

import os
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox

import converters

try:
    from PIL import Image, ImageTk
    _HAS_PIL = True
except Exception:
    _HAS_PIL = False


WINDOW_W = 580
WINDOW_H = 740

BG_COLOR = "#2b3a67"
CARD_COLOR = "#f4f6fb"
TEXT_DARK = "#1b1b1b"
TEXT_MUTED = "#555555"

# Dugme renkleri
BLUE = "#1f6feb"
BLUE_H = "#1a5fd0"
RED = "#c0392b"
RED_H = "#a93224"
GREEN = "#1e8449"
GREEN_H = "#176c3c"

BG_NAMES = [
    "arkaplan.png", "arkaplan.jpg", "arkaplan.jpeg",
    "background.png", "background.jpg", "background.jpeg",
]


def app_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def find_background():
    base = app_dir()
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

        self._draw_background()
        self._build_ui()

    def _center(self):
        self.root.update_idletasks()
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        x = (sw - WINDOW_W) // 2
        y = max(0, (sh - WINDOW_H) // 3)
        self.root.geometry("+%d+%d" % (x, y))

    def _draw_background(self):
        bg_path = find_background()
        if bg_path and _HAS_PIL:
            try:
                img = Image.open(bg_path).convert("RGB")
                img = self._cover(img, WINDOW_W, WINDOW_H)
                self._bg_photo = ImageTk.PhotoImage(img)
                self.canvas.create_image(0, 0, anchor="nw", image=self._bg_photo)
                return
            except Exception:
                pass
        self.canvas.create_rectangle(0, 0, WINDOW_W, 130, fill="#243156", outline="")

    @staticmethod
    def _cover(img, target_w, target_h):
        src_w, src_h = img.size
        scale = max(target_w / src_w, target_h / src_h)
        new_w, new_h = int(src_w * scale), int(src_h * scale)
        img = img.resize((new_w, new_h), Image.LANCZOS)
        left = (new_w - target_w) // 2
        top = (new_h - target_h) // 2
        return img.crop((left, top, left + target_w, top + target_h))

    def _build_ui(self):
        cx = WINDOW_W // 2

        self.canvas.create_text(
            cx, 70, text="UDFinator",
            font=("Segoe UI", 26, "bold"), fill="white",
        )

        card_x0, card_y0 = 40, 170
        card_x1, card_y1 = WINDOW_W - 40, WINDOW_H - 50
        self.canvas.create_rectangle(
            card_x0, card_y0, card_x1, card_y1,
            fill=CARD_COLOR, outline="",
        )

        self.canvas.create_text(
            cx, card_y0 + 42,
            text="Ne yapmak istiyorsunuz?",
            font=("Segoe UI", 16, "bold"), fill=TEXT_DARK,
        )

        # Uc buyuk dugme
        specs = [
            ("📄  UDF → Word (.docx)", BLUE, BLUE_H, self.do_udf_to_docx),
            ("📕  UDF → PDF", RED, RED_H, self.do_udf_to_pdf),
            ("🔄  PDF → UDF", GREEN, GREEN_H, self.do_pdf_to_udf),
        ]
        y = card_y0 + 110
        for text, color, hover, command in specs:
            btn = tk.Button(
                self.root, text=text,
                font=("Segoe UI", 15, "bold"),
                bg=color, fg="white",
                activebackground=hover, activeforeground="white",
                relief="flat", bd=0, cursor="hand2",
                padx=16, pady=14, command=command,
            )
            self.canvas.create_window(cx, y, window=btn, width=380)
            btn.bind("<Enter>", lambda e, b=btn, h=hover: self._hover(b, h))
            btn.bind("<Leave>", lambda e, b=btn, c=color: self._hover(b, c))
            self._buttons.append((btn, color))
            y += 78

        # Durum yazisi
        self.status_var = tk.StringVar(value="Hazır.")
        self.status_label = tk.Label(
            self.root, textvariable=self.status_var,
            font=("Segoe UI", 12), bg=CARD_COLOR, fg=TEXT_MUTED,
            wraplength=420, justify="center",
        )
        self.canvas.create_window(cx, y + 6, window=self.status_label)

        # Alt ipucu
        self.canvas.create_text(
            cx, card_y1 - 26,
            text="Yeni dosya, kaynak dosyanın yanına kaydedilir.",
            font=("Segoe UI", 10), fill=TEXT_MUTED, justify="center",
        )

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
        self.status_var.set("Dönüştürülüyor, lütfen bekleyin...")
        self.status_label.config(fg=TEXT_MUTED)
        threading.Thread(
            target=self._worker, args=(path, out_ext, fn, out_name), daemon=True
        ).start()

    def _worker(self, path, out_ext, fn, out_name):
        out_path = os.path.splitext(path)[0] + out_ext
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
        self.status_var.set("Tamamlandı! ✅")
        self.status_label.config(fg="#1a7f37")
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
        self.status_var.set("Dönüştürülemedi.")
        self.status_label.config(fg="#b42318")
        messagebox.showerror("Hata", "Dosya dönüştürülemedi.\n\n%s" % message)


def main():
    root = tk.Tk()
    ConverterApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
