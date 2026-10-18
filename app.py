# -*- coding: utf-8 -*-
"""
UDFinator — basit arayuz (Turkce).

Akis:
  1. Ana menu: dort donusumden biri secilir.
  2. Dosya ekrani: dosya pencereye surukleyip birakilir ya da
     "Klasorden Sec" ile secilir. Donusum gecici bir dosyaya yapilir.
  3. Sonuc ekrani: "Kaydet" (orijinal dosyanin yanina) ya da
     "Farkli Kaydet" (yeri ve adi kullanici secer).

Program DPI'a duyarlidir: tum olculer mantiksal (96 dpi) birimlerdir ve
ekran olcegiyle carpilir; boylece %150 gibi olceklerde bulanik degil keskin
gorunur.

Arka plan resmi: image.jpg derleme sirasinda exe'nin icine gomulur. Exe'nin
yanina image.jpg (ya da arkaplan.jpg / background.jpg) koyulursa o kullanilir.
"""

import atexit
import json
import os
import shutil
import subprocess
import sys
import tempfile
import threading
import tkinter as tk
from tkinter import filedialog, messagebox

import converters

try:
    from PIL import Image, ImageDraw, ImageTk
    _HAS_PIL = True
except Exception:
    _HAS_PIL = False

try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    _HAS_DND_LIB = True
except Exception:
    _HAS_DND_LIB = False


# Mantiksal boyutlar (96 dpi). Kucuk ekranlara da sigar.
WINDOW_W = 600
WINDOW_H = 640
HEADER_H = 96
ZONE = (50, 122, WINDOW_W - 50, 470)        # surukle-birak / sonuc paneli
STATUS_Y = 505
BOTTOM_Y = 582

BG_COLOR = "#2b3a67"
HEADER_COLOR = "#243156"
CARD_COLOR = "#f4f6fb"
ZONE_COLOR = "#e7ebf5"
TEXT_DARK = "#1b1b1b"
TEXT_MUTED = "#555555"

BLUE, BLUE_H = "#1f6feb", "#1a5fd0"
RED, RED_H = "#c0392b", "#a93224"
PURPLE, PURPLE_H = "#7d3c98", "#6c3483"
GREEN, GREEN_H = "#1e8449", "#176c3c"
GRAY, GRAY_H = "#5c6680", "#4c5570"

MODES = [
    dict(key="udf2docx", icon="📄", label="UDF → Word (.docx)", color=BLUE, hover=BLUE_H,
         in_exts=(".udf",), in_label="UDF", out_ext=".docx", out_label="Word belgesi",
         fn=converters.udf_to_docx),
    dict(key="udf2pdf", icon="📕", label="UDF → PDF", color=RED, hover=RED_H,
         in_exts=(".udf",), in_label="UDF", out_ext=".pdf", out_label="PDF dosyası",
         fn=converters.udf_to_pdf),
    dict(key="docx2udf", icon="📝", label="Word (.docx) → UDF", color=PURPLE, hover=PURPLE_H,
         in_exts=(".docx",), in_label="Word (.docx)", out_ext=".udf", out_label="UDF dosyası",
         fn=converters.docx_to_udf),
    dict(key="pdf2udf", icon="🔄", label="PDF → UDF", color=GREEN, hover=GREEN_H,
         in_exts=(".pdf",), in_label="PDF", out_ext=".udf", out_label="UDF dosyası",
         fn=converters.pdf_to_udf),
]

BG_NAMES = [
    "image.jpg", "image.jpeg", "image.png",
    "arkaplan.jpg", "arkaplan.jpeg", "arkaplan.png",
    "background.jpg", "background.jpeg", "background.png",
]


def enable_dpi_awareness():
    """Windows'un pencereyi bitmap gibi buyutup bulaniklastirmasini engeller."""
    try:
        import ctypes
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass


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


def make_root():
    """Surukle-birak destekli pencere; kutuphane yuklenemezse normal pencere."""
    if _HAS_DND_LIB:
        try:
            return TkinterDnD.Tk(), True
        except Exception:
            pass
    return tk.Tk(), False


class ConverterApp:
    def __init__(self, root, dnd=False):
        self.root = root
        self.dnd = dnd
        self.scale = max(root.winfo_fpixels("1i") / 96.0, 1.0)
        root.title("UDFinator")
        root.geometry("%dx%d" % (self.px(WINDOW_W), self.px(WINDOW_H)))
        root.resizable(False, False)
        root.configure(bg=BG_COLOR)
        self._center()

        self.mode = None
        self.source = None          # secilen dosya
        self.temp_out = None        # donusturulmus gecici dosya
        self.saved_path = None
        self.busy = False
        self._widgets = []
        self._bg_photo = None
        self._tmpdir = tempfile.mkdtemp(prefix="udfinator_")
        atexit.register(self._cleanup)
        root.protocol("WM_DELETE_WINDOW", self._close)

        self.canvas = tk.Canvas(root, width=self.px(WINDOW_W), height=self.px(WINDOW_H),
                                highlightthickness=0, bg=BG_COLOR)
        self.canvas.pack(fill="both", expand=True)
        self._register_drop(self.canvas)

        self._photo_src = self._load_photo()
        self.has_photo = self._photo_src is not None
        if self.has_photo:
            self.fg_main, self.fg_muted = "white", "#eeeeee"
            self.fg_ok, self.fg_err = "#8ef0b6", "#ffb3b3"
            self.dash, self.dash_hot = "#ffffff", "#ffd166"
        else:
            self.fg_main, self.fg_muted = TEXT_DARK, TEXT_MUTED
            self.fg_ok, self.fg_err = "#1a7f37", "#b42318"
            self.dash, self.dash_hot = "#8a94b3", BLUE
        self.show_menu()

    # --- yardimcilar -------------------------------------------------------

    def px(self, v):
        return int(round(v * self.scale))

    def _center(self):
        self.root.update_idletasks()
        sw, sh = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
        x = max(0, (sw - self.px(WINDOW_W)) // 2)
        y = max(0, (sh - self.px(WINDOW_H)) // 3)
        self.root.geometry("+%d+%d" % (x, y))

    def _load_photo(self):
        path = find_background()
        if not (path and _HAS_PIL):
            return None
        try:
            img = Image.open(path).convert("RGB")
            w, h = self.px(WINDOW_W), self.px(WINDOW_H)
            s = max(w / img.width, h / img.height)
            img = img.resize((int(img.width * s) + 1, int(img.height * s) + 1), Image.LANCZOS)
            left, top = (img.width - w) // 2, (img.height - h) // 2
            return img.crop((left, top, left + w, top + h))
        except Exception:
            return None

    def _render_bg(self, panels):
        """panels: [(tur, (x0, y0, x1, y1))], tur: header | card | zone | pill.
        Fotografta yazilarin arkasina yari saydam koyu paneller, fotografsiz
        tasarimda duz renkli kartlar cizilir; tum cizim ekran cozunurlugunde."""
        w, h = self.px(WINDOW_W), self.px(WINDOW_H)
        r = self.px(14)
        if not _HAS_PIL:
            for kind, rect in panels:
                fill = HEADER_COLOR if kind == "header" else CARD_COLOR
                self.canvas.create_rectangle(*[self.px(v) for v in rect], fill=fill, outline="")
            return
        if self.has_photo:
            img = self._photo_src.convert("RGBA")
            overlay = Image.new("RGBA", (w, h), (0, 0, 0, 0))
            d = ImageDraw.Draw(overlay)
            for kind, rect in panels:
                box = tuple(self.px(v) for v in rect)
                if kind == "header":
                    d.rectangle(box, fill=(0, 0, 0, 120))
                elif kind in ("zone", "pill"):
                    d.rounded_rectangle(box, radius=r, fill=(0, 0, 0, 140))
            img = Image.alpha_composite(img, overlay).convert("RGB")
        else:
            img = Image.new("RGB", (w, h), BG_COLOR)
            d = ImageDraw.Draw(img)
            for kind, rect in panels:
                box = tuple(self.px(v) for v in rect)
                if kind == "header":
                    d.rectangle(box, fill=HEADER_COLOR)
                elif kind == "card":
                    d.rounded_rectangle(box, radius=r, fill=CARD_COLOR)
                elif kind == "zone":
                    d.rounded_rectangle(box, radius=r, fill=ZONE_COLOR)
        self._bg_photo = ImageTk.PhotoImage(img)
        self.canvas.create_image(0, 0, anchor="nw", image=self._bg_photo)

    def _clear(self):
        self.canvas.delete("all")
        for wdg in self._widgets:
            wdg.destroy()
        self._widgets = []

    def _text(self, y, text, size, bold=False, fill=None, width=None, x=None):
        return self.canvas.create_text(
            self.px(x if x is not None else WINDOW_W // 2), self.px(y), text=text,
            font=("Segoe UI", size, "bold" if bold else "normal"),
            fill=fill or self.fg_main, justify="center",
            width=self.px(width) if width else 0)

    def _button(self, y, text, color, hover, command, width=400, size=15, x=None):
        btn = tk.Button(self.root, text=text, font=("Segoe UI", size, "bold"),
                        bg=color, fg="white", activebackground=hover,
                        activeforeground="white", disabledforeground="#d9d9d9",
                        relief="flat", bd=0, cursor="hand2",
                        padx=self.px(16), pady=self.px(8), command=command)
        self.canvas.create_window(self.px(x if x is not None else WINDOW_W // 2), self.px(y),
                                  window=btn, width=self.px(width))
        btn.bind("<Enter>", lambda e: btn.config(bg=hover) if str(btn["state"]) == "normal" else None)
        btn.bind("<Leave>", lambda e: btn.config(bg=color))
        btn._base_color = color
        self._register_drop(btn)
        self._widgets.append(btn)
        return btn

    def _dashed_rect(self, x0, y0, x1, y1, color, tag):
        """Kesik cizgili cerceve (Windows'ta kalin kesik cizgi desteklenmedigi
        icin kisa parcalarla cizilir)."""
        dash, gap, width = self.px(14), self.px(9), max(2, self.px(2))
        x0, y0, x1, y1 = (self.px(v) for v in (x0, y0, x1, y1))
        for a, b, fixed, horizontal in ((x0, x1, y0, True), (x0, x1, y1, True),
                                        (y0, y1, x0, False), (y0, y1, x1, False)):
            p = a
            while p < b:
                q = min(p + dash, b)
                coords = (p, fixed, q, fixed) if horizontal else (fixed, p, fixed, q)
                self.canvas.create_line(*coords, fill=color, width=width, tags=tag)
                p = q + gap

    def _header(self, title):
        self._text(HEADER_H // 2, title, 22, bold=True, fill="white")

    def _status(self, text, color=None):
        if getattr(self, "status_item", None):
            self.canvas.itemconfig(self.status_item, text=text, fill=color or self.fg_muted)

    # --- ekran 1: ana menu -------------------------------------------------

    def show_menu(self):
        self._clear()
        self.mode = None
        self._discard_temp()
        if self.has_photo:
            self._render_bg([("header", (0, 0, WINDOW_W, HEADER_H)),
                             ("pill", (70, 126, WINDOW_W - 70, 174)),
                             ("pill", (60, WINDOW_H - 72, WINDOW_W - 60, WINDOW_H - 22))])
        else:
            self._render_bg([("header", (0, 0, WINDOW_W, HEADER_H)),
                             ("card", (40, 118, WINDOW_W - 40, WINDOW_H - 20))])
        self._header("UDFinator")
        self._text(150, "Ne yapmak istiyorsunuz?", 16, bold=True)
        for i, m in enumerate(MODES):
            self._button(228 + i * 74, "%s  %s" % (m["icon"], m["label"]), m["color"], m["hover"],
                         lambda m=m: self.show_drop(m))
        self.status_item = None
        self._text(WINDOW_H - 47, "Bir dönüştürme seçin, sonra dosyanızı verin.", 11,
                   fill=self.fg_muted)

    # --- ekran 2: dosya verme (surukle-birak / klasorden sec) ---------------

    def show_drop(self, mode):
        self._clear()
        self.mode = mode
        self.source = None
        self._discard_temp()
        x0, y0, x1, y1 = ZONE
        if self.has_photo:
            self._render_bg([("header", (0, 0, WINDOW_W, HEADER_H)), ("zone", ZONE),
                             ("pill", (60, STATUS_Y - 22, WINDOW_W - 60, STATUS_Y + 22))])
        else:
            self._render_bg([("header", (0, 0, WINDOW_W, HEADER_H)),
                             ("card", (40, 112, WINDOW_W - 40, WINDOW_H - 20)),
                             ("zone", ZONE)])
        self._header("%s  %s" % (mode["icon"], mode["label"]))
        self._dashed_rect(x0 + 12, y0 + 12, x1 - 12, y1 - 12, self.dash, "dash")

        if self.dnd:
            self._text(205, "⬇", 34, fill=self.fg_main)
            self._text(263, "Dosyayı buraya sürükleyip bırakın", 16, bold=True, width=440)
            self._text(297, "(%s dosyası)" % mode["in_label"], 11, fill=self.fg_muted)
            self._text(333, "veya", 12, fill=self.fg_muted)
        else:
            self._text(220, "📂", 34, fill=self.fg_main)
            self._text(285, "Dosyanızı seçmek için aşağıdaki düğmeye tıklayın", 15,
                       bold=True, width=440)
        self.pick_btn = self._button(392, "📂  Klasörden Seç", mode["color"], mode["hover"],
                                     self.pick_file, width=300)
        self.status_item = self._text(STATUS_Y, "Dönüştürmek için bir dosya verin.", 12, width=470)
        self.back_btn = self._button(BOTTOM_Y, "←  Ana Menü", GRAY, GRAY_H, self.show_menu,
                                     width=220, size=13)

    def _register_drop(self, widget):
        if not self.dnd:
            return
        try:
            widget.drop_target_register(DND_FILES)
            widget.dnd_bind("<<Drop>>", self._on_drop)
            widget.dnd_bind("<<DropEnter>>", self._on_drag_enter)
            widget.dnd_bind("<<DropLeave>>", self._on_drag_leave)
        except Exception:
            pass

    def _drop_ready(self):
        return self.mode is not None and self.temp_out is None and not self.busy

    def _on_drag_enter(self, event):
        if self._drop_ready():
            self.canvas.itemconfig("dash", fill=self.dash_hot)
        return event.action

    def _on_drag_leave(self, event):
        self.canvas.itemconfig("dash", fill=self.dash)
        return event.action

    def _on_drop(self, event):
        self.canvas.itemconfig("dash", fill=self.dash)
        if self._drop_ready():
            files = self.root.tk.splitlist(event.data)
            if files:
                self.start_conversion(files[0])
        return event.action

    def pick_file(self):
        m = self.mode
        patterns = " ".join("*" + e for e in m["in_exts"])
        path = filedialog.askopenfilename(
            title="%s dosyası seçin" % m["in_label"],
            filetypes=[("%s dosyaları" % m["in_label"], patterns), ("Tüm dosyalar", "*.*")])
        if path:
            self.start_conversion(path)

    def start_conversion(self, path):
        m = self.mode
        path = os.path.normpath(path)  # surukle-birak 'C:/...' bicimini Windows bicimine
        ext = os.path.splitext(path)[1].lower()
        if os.path.isdir(path) or not os.path.isfile(path):
            self._status("Lütfen bir klasör değil, bir dosya verin.", self.fg_err)
            return
        if ext not in m["in_exts"]:
            if ext == ".doc" and ".docx" in m["in_exts"]:
                msg = "Bu eski bir Word (.doc) dosyası. Word'de açıp .docx olarak kaydedin."
            else:
                msg = "Bu bir %s dosyası değil. Lütfen %s dosyası verin." % (
                    m["in_label"], m["in_label"])
            self._status(msg, self.fg_err)
            return
        self.source = path
        self.busy = True
        for b in (self.pick_btn, self.back_btn):
            b.config(state="disabled", bg=b._base_color)
        self._status("Dönüştürülüyor, lütfen bekleyin...", self.fg_muted)
        out = os.path.join(self._tmpdir, os.path.splitext(os.path.basename(path))[0] + m["out_ext"])
        self._result = None
        threading.Thread(target=self._worker, args=(m["fn"], path, out), daemon=True).start()
        self.root.after(100, self._poll)

    def _worker(self, fn, src, out):
        """Arka planda calisir; Tk'ye dokunmaz, sonucu yalnizca birakir."""
        try:
            fn(src, out)
        except converters.UdfError as exc:
            self._result = ("error", str(exc))
        except Exception as exc:
            self._result = ("error", "Beklenmeyen hata: %s" % exc)
        else:
            self._result = ("ok", out)

    def _poll(self):
        """Ana is parcaciginda sonucu bekler (Tk yalnizca buradan guncellenir)."""
        if self._result is None:
            self.root.after(100, self._poll)
            return
        kind, value = self._result
        self._result = None
        if kind == "ok":
            self._on_converted(value)
        else:
            self._on_error(value)

    def _on_error(self, message):
        self.busy = False
        for b in (self.pick_btn, self.back_btn):
            b.config(state="normal")
        self._status("Dönüştürülemedi.", self.fg_err)
        messagebox.showerror("Hata", "Dosya dönüştürülemedi.\n\n%s" % message)

    def _on_converted(self, out):
        self.busy = False
        self.temp_out = out
        self.saved_path = None
        self.show_result()

    # --- ekran 3: sonuc (Kaydet / Farkli Kaydet) ----------------------------

    def show_result(self):
        self._clear()
        if self.has_photo:
            self._render_bg([("header", (0, 0, WINDOW_W, HEADER_H)), ("zone", ZONE),
                             ("pill", (60, STATUS_Y - 30, WINDOW_W - 60, STATUS_Y + 30))])
        else:
            self._render_bg([("header", (0, 0, WINDOW_W, HEADER_H)),
                             ("card", (40, 112, WINDOW_W - 40, WINDOW_H - 20)),
                             ("zone", ZONE)])
        self._header("✅  Dönüştürme Tamamlandı")
        self._text(160, "Yeni dosyanız hazır:", 13, fill=self.fg_muted)
        self._text(193, os.path.basename(self.temp_out), 13, bold=True, width=460)
        self.save_btn = self._button(262, "💾  Kaydet", GREEN, GREEN_H, self.save, width=380)
        self._text(302, "Orijinal dosyanın yanına kaydeder", 10, fill=self.fg_muted)
        self._button(360, "📁  Farklı Kaydet…", BLUE, BLUE_H, self.save_as, width=380)
        self._text(400, "Kaydedilecek yeri ve adı siz seçin", 10, fill=self.fg_muted)
        self.status_item = self._text(STATUS_Y, "Henüz kaydedilmedi.", 12, width=470)
        self.open_btn = self._button(BOTTOM_Y, "📂  Klasörü Aç", GRAY, GRAY_H, self.open_folder,
                                     width=230, size=13, x=WINDOW_W // 2 - 125)
        self.open_btn.config(state="disabled")
        self._button(BOTTOM_Y, "🔁  Yeni Dönüştürme", GRAY, GRAY_H, self.show_menu,
                     width=230, size=13, x=WINDOW_W // 2 + 125)

    def _copy_to(self, dest):
        if os.path.abspath(dest).lower() == os.path.abspath(self.source).lower():
            messagebox.showerror("Hata", "Kaynak dosyanın üzerine kaydedilemez.\n"
                                         "Lütfen başka bir ad seçin.")
            return False
        try:
            shutil.copyfile(self.temp_out, dest)
        except Exception as exc:
            messagebox.showerror("Hata", "Dosya kaydedilemedi.\n\n%s" % exc)
            return False
        self.saved_path = dest
        self.open_btn.config(state="normal")
        folder = os.path.dirname(dest)
        parts = folder.split(os.sep)
        if len(folder) > 48 and len(parts) > 2:
            folder = "…" + os.sep + os.sep.join(parts[-2:])
        self._status("✔  Kaydedildi: %s\nKlasör: %s" % (os.path.basename(dest), folder), self.fg_ok)
        return True

    def save(self):
        """Orijinal dosyanin yanina; ayni adda dosya varsa 'ad (1)' olarak."""
        dest = unique_path(os.path.join(os.path.dirname(self.source),
                                        os.path.basename(self.temp_out)))
        if self._copy_to(dest):
            self.save_btn.config(state="disabled", text="✔  Kaydedildi", bg=GREEN)

    def save_as(self):
        m = self.mode
        dest = filedialog.asksaveasfilename(
            title="Farklı Kaydet",
            initialdir=os.path.dirname(self.source),
            initialfile=os.path.basename(self.temp_out),
            defaultextension=m["out_ext"],
            filetypes=[(m["out_label"], "*" + m["out_ext"]), ("Tüm dosyalar", "*.*")])
        if dest:
            self._copy_to(dest)

    def open_folder(self):
        if not self.saved_path:
            return
        try:
            subprocess.Popen(["explorer", "/select,", os.path.normpath(self.saved_path)])
        except Exception:
            try:
                os.startfile(os.path.dirname(self.saved_path))
            except Exception:
                pass

    # --- temizlik ----------------------------------------------------------

    def _discard_temp(self):
        if self.temp_out and os.path.exists(self.temp_out):
            try:
                os.remove(self.temp_out)
            except Exception:
                pass
        self.temp_out = None

    def _cleanup(self):
        shutil.rmtree(self._tmpdir, ignore_errors=True)

    def _close(self):
        if self.temp_out and not self.saved_path and not self.busy:
            if not messagebox.askyesno(
                    "Çıkış", "Dönüştürülen dosya henüz kaydedilmedi.\nYine de çıkılsın mı?"):
                return
        self._cleanup()
        self.root.destroy()


def _cli_convert(argv):
    """Gizli komut satiri modu: UDFinator.exe --convert MOD GIRDI CIKTI
    (MOD: udf2docx | udf2pdf | docx2udf | pdf2udf). Basari: cikis kodu 0."""
    modes = {m["key"]: m["fn"] for m in MODES}
    if len(argv) != 4 or argv[1] not in modes:
        return 2
    try:
        modes[argv[1]](argv[2], argv[3])
    except Exception:
        return 1
    return 0


def _selftest(out_path):
    """Gizli mod: UDFinator.exe --selftest SONUC.json — paketlenmis exe'de
    surukle-birak, PIL ve arka plan resminin calistigini dogrulamak icin."""
    enable_dpi_awareness()
    root, dnd = make_root()
    root.withdraw()
    info = {"dnd": dnd, "pil": _HAS_PIL, "background": find_background(),
            "scale": root.winfo_fpixels("1i") / 96.0}
    root.destroy()
    with open(out_path, "w", encoding="utf-8") as fh:
        json.dump(info, fh)
    return 0


def main():
    if len(sys.argv) > 1 and sys.argv[1] == "--convert":
        sys.exit(_cli_convert(sys.argv[1:]))
    if len(sys.argv) > 2 and sys.argv[1] == "--selftest":
        sys.exit(_selftest(sys.argv[2]))
    enable_dpi_awareness()
    root, dnd = make_root()
    ConverterApp(root, dnd)
    root.mainloop()


if __name__ == "__main__":
    main()
