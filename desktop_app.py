from __future__ import annotations

import os
import threading
import tkinter as tk
from datetime import date, datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from app import (
    OUTPUT_DIR,
    ProcessCancelled,
    RAW_HEADERS,
    build_category_workbook,
    build_olah_data_workbook,
    build_workbook,
    convert_pdfs_to_excel,
    read_category_detail,
    clear_process_logs,
    db_record_count,
    db_recent_months,
    detect_dates_in_records,
    detect_months_in_records,
    extract_categories,
    extract_records,
    duplicate_category_files_to_remove,
    find_duplicate_categories,
    load_process_logs,
    load_records_from_db,
    reset_database,
    save_process_log,
    save_records_to_db,
)

AUTO_MONTHLY_MIN_DAYS = 20


def format_rupiah(value: float | int | str | None) -> str:
    try:
        amount = float(value or 0)
    except (TypeError, ValueError):
        amount = 0
    return f"Rp {amount:,.0f}".replace(",", ".")


def duplicate_booking_warning_text(duplicates: dict[str, list[dict[str, object]]]) -> str:
    if not duplicates:
        return ""

    total_rows = sum(len(rows) for rows in duplicates.values())
    lines = [
        f"Ditemukan {len(duplicates)} Booking ID duplikat ({total_rows} baris).",
        "Semua baris tetap dihitung dalam rekap. Periksa jika ada nominal yang memang perlu dikoreksi.",
        "",
    ]
    for booking_id, rows in list(duplicates.items())[:10]:
        amounts = ", ".join(format_rupiah(row.get("revenue")) for row in rows)
        courts = sorted({str(row.get("court") or "-") for row in rows})
        times = sorted({str(row.get("start_time") or "-") for row in rows})
        lines.append(
            f"- {booking_id}: {len(rows)} baris, nominal {amounts}, "
            f"court {', '.join(courts)}, jam {', '.join(times)}"
        )
    if len(duplicates) > 10:
        lines.append(f"- ... dan {len(duplicates) - 10} Booking ID lainnya.")
    return "\n".join(lines)


def infer_report_period(
    records: list[dict[str, object]],
    mode_choice: str = "Otomatis",
) -> tuple[date | None, date | None, list[date] | None, str, str]:
    detected_dates = detect_dates_in_records(records)
    if not detected_dates:
        raise ValueError("Tanggal booking tidak terbaca dari kolom Date of Booking.")

    detected_months = detect_months_in_records(records)
    if len(detected_months) > 1:
        months = ", ".join(month.strftime("%Y-%m") for month in detected_months)
        raise ValueError(f"File berisi lebih dari satu bulan ({months}). Pisahkan file per bulan dulu.")

    if len(detected_dates) == 1:
        selected = detected_dates[0]
        return selected, None, None, "Harian", f"Harian otomatis: {selected:%Y-%m-%d}"

    selected_month = detected_months[0]
    if mode_choice == "Sebulan":
        return None, selected_month, None, "Sebulan", f"Sebulan dipilih: {selected_month:%Y-%m}"
    if mode_choice == "Multi Hari":
        return None, selected_month, detected_dates, "Multi Hari", f"Multi Hari: {detected_dates[0]:%Y-%m-%d} s/d {detected_dates[-1]:%Y-%m-%d}"
    if len(detected_dates) >= AUTO_MONTHLY_MIN_DAYS:
        return None, selected_month, None, "Sebulan", f"Sebulan otomatis: {selected_month:%Y-%m}"
    return None, selected_month, detected_dates, "Multi Hari", f"Multi Hari otomatis: {detected_dates[0]:%Y-%m-%d} s/d {detected_dates[-1]:%Y-%m-%d}"


def format_period_label(
    report_date: date | None,
    report_month: date | None,
    report_dates: list[date] | None = None,
) -> str:
    if report_dates:
        selected = sorted(report_dates)
        if len(selected) == 1:
            return selected[0].isoformat()
        return f"{selected[0].isoformat()} s/d {selected[-1].isoformat()}"
    if report_date:
        return report_date.isoformat()
    if report_month:
        return f"{report_month:%Y-%m}"
    return "-"


def preview_summary(
    source: dict[str, object],
    report_date: date | None,
    report_month: date | None,
    report_dates: list[date] | None,
    feature: str,
) -> dict[str, str]:
    records = source["records"]
    if not isinstance(records, list):
        records = []

    dates = detect_dates_in_records(records)
    bk_rows = 0
    mn_rows = 0
    courts: set[str] = set()
    for record in records:
        booking_id = str(record.get("Booking ID") or "").upper()
        if booking_id.startswith("BK/"):
            bk_rows += 1
        elif booking_id.startswith("MN/"):
            mn_rows += 1
        court = str(record.get("Court") or "").strip()
        if court:
            courts.add(court)

    date_range = "-"
    if dates:
        date_range = dates[0].isoformat() if len(dates) == 1 else f"{dates[0].isoformat()} s/d {dates[-1].isoformat()}"

    return {
        "mode": feature,
        "period": format_period_label(report_date, report_month, report_dates),
        "rows": str(len(records)),
        "dates": date_range,
        "channels": f"AYO {bk_rows} | Walk In {mn_rows}",
        "courts": str(len(courts)),
    }


# ── Palet warna UI ──────────────────────────────────────────────────────
COL_BG = "#f5f7fa"          # background aplikasi
COL_SURFACE = "#ffffff"     # kartu / panel putih
COL_SOFT = "#f7f9fb"        # panel lembut
COL_HEADER = "#103b46"      # deep court teal
COL_HEADER_SUB = "#b9d9dc"  # teks subtitle di header
COL_BRAND = "#0f766e"       # court teal
COL_BRAND_HOVER = "#0d8a7f"
COL_BRAND_PRESS = "#095c56"
COL_BLUE = "#2563a6"
COL_BLUE_DARK = "#1d4f85"
COL_AMBER = "#c88716"
COL_GREEN = "#4f8a5b"
COL_RED = "#ef4444"
COL_SIDEBAR = "#101b2b"
COL_SIDEBAR_DARK = "#09121f"
COL_SIDEBAR_SOFT = "#1b2b40"
COL_CONTENT = "#f5f7fa"
COL_TEXT = "#172033"
COL_MUTED = "#667085"
COL_BORDER = "#e2e8f0"
COL_BORDER_STRONG = "#cbd5e1"
COL_DANGER = "#e11d48"
COL_DANGER_HOVER = "#be123c"
DEFAULT_APP_GEOMETRY = "1120x700"
DEFAULT_APP_MINSIZE = (980, 620)

FONT = "Segoe UI"
FONT_SB = "Segoe UI Semibold"


def compact_text(value: str, limit: int = 28) -> str:
    text = str(value or "")
    if len(text) <= limit:
        return text
    return text[: max(0, limit - 3)].rstrip() + "..."


def match_window_size(target: tk.Misc, source: tk.Misc | None) -> None:
    if source is None:
        return
    try:
        source.update_idletasks()
        if source.state() == "zoomed":
            target.state("zoomed")
        else:
            target.state("normal")
            geometry = source.winfo_geometry()
            if "x" in geometry:
                target.geometry(geometry)
    except tk.TclError:
        pass


def configure_styles(root: tk.Misc) -> None:
    style = ttk.Style(root)
    try:
        style.theme_use("clam")
    except tk.TclError:
        pass

    # Frames
    style.configure("App.TFrame", background=COL_BG)
    style.configure("Shell.TFrame", background=COL_CONTENT)
    style.configure("Sidebar.TFrame", background=COL_SIDEBAR)
    style.configure("Content.TFrame", background=COL_CONTENT)
    style.configure("Card.TFrame", background=COL_SURFACE, relief="flat")
    style.configure(
        "Elevated.TFrame",
        background=COL_SURFACE,
        relief="solid",
        borderwidth=1,
        bordercolor=COL_BORDER,
        lightcolor=COL_BORDER,
        darkcolor=COL_BORDER,
    )
    style.configure("SoftCard.TFrame", background=COL_SOFT, relief="flat")
    style.configure("Header.TFrame", background=COL_HEADER)

    # Header text
    style.configure("Title.TLabel", background=COL_HEADER, foreground="#ffffff", font=(FONT_SB, 20))
    style.configure("Subtitle.TLabel", background=COL_HEADER, foreground=COL_HEADER_SUB, font=(FONT, 10))
    style.configure("PageTitle.TLabel", background=COL_CONTENT, foreground=COL_TEXT, font=(FONT_SB, 22))
    style.configure("PageSub.TLabel", background=COL_CONTENT, foreground=COL_MUTED, font=(FONT, 10))
    style.configure("SidebarTitle.TLabel", background=COL_SIDEBAR, foreground="#ffffff", font=(FONT_SB, 15))
    style.configure("SidebarSub.TLabel", background=COL_SIDEBAR, foreground="#9ca3af", font=(FONT, 9))

    # Labels umum
    style.configure("Section.TLabel", background=COL_SOFT, foreground=COL_TEXT, font=(FONT, 9, "bold"))
    style.configure("Body.TLabel", background=COL_SURFACE, foreground=COL_TEXT, font=(FONT_SB, 11))
    style.configure("Hint.TLabel", background=COL_SURFACE, foreground=COL_MUTED, font=(FONT, 9))
    style.configure("PanelHint.TLabel", background=COL_CONTENT, foreground=COL_MUTED, font=(FONT, 9))
    style.configure("AutoHint.TLabel", background=COL_SOFT, foreground=COL_MUTED, font=(FONT, 9))
    style.configure("MetricTitle.TLabel", background=COL_SOFT, foreground=COL_MUTED, font=(FONT, 8, "bold"))
    style.configure("MetricValue.TLabel", background=COL_SOFT, foreground=COL_TEXT, font=(FONT_SB, 15))
    style.configure("PreviewValue.TLabel", background=COL_SOFT, foreground=COL_TEXT, font=(FONT_SB, 10))
    style.configure("CardTitle.TLabel", background=COL_SURFACE, foreground=COL_TEXT, font=(FONT_SB, 12))
    style.configure("CardSub.TLabel", background=COL_SURFACE, foreground=COL_MUTED, font=(FONT, 9))

    # Tombol flat + efek hover
    def flat_button(name, bg, fg, hover, press, disabled_bg, disabled_fg="#eef2f6",
                    font=(FONT_SB, 10), pad=(18, 10)):
        style.configure(name, font=font, padding=pad, background=bg, foreground=fg,
                        borderwidth=0, relief="flat", focusthickness=0, focuscolor=bg,
                        lightcolor=bg, darkcolor=bg, bordercolor=bg)
        style.map(name,
                  background=[("disabled", disabled_bg), ("pressed", press), ("active", hover)],
                  foreground=[("disabled", disabled_fg)],
                  relief=[("pressed", "flat"), ("active", "flat")])

    flat_button("Primary.TButton", COL_BRAND, "#ffffff", COL_BRAND_HOVER, COL_BRAND_PRESS, "#9fb6b3")
    flat_button("PrimaryLarge.TButton", COL_BRAND, "#ffffff", COL_BRAND_HOVER, COL_BRAND_PRESS, "#9fb6b3",
                font=(FONT_SB, 11), pad=(20, 12))
    flat_button("Blue.TButton", COL_BLUE, "#ffffff", "#2f73bb", COL_BLUE_DARK, "#a9bdd3")
    flat_button("Gold.TButton", COL_AMBER, "#ffffff", "#dc9b2d", "#9d6710", "#d8c6a7")
    flat_button("Danger.TButton", COL_DANGER, "#ffffff", "#f43f5e", COL_DANGER_HOVER, "#f3b6c2")
    flat_button("Secondary.TButton", COL_SURFACE, COL_TEXT, "#f8fafc", "#eef2f6", "#f1f5f9",
                disabled_fg="#94a3b8", font=(FONT, 10), pad=(14, 9))
    style.configure("Secondary.TButton", borderwidth=1, bordercolor=COL_BORDER_STRONG,
                    lightcolor=COL_SURFACE, darkcolor=COL_SURFACE)
    style.map("Secondary.TButton", bordercolor=[("active", COL_BRAND), ("disabled", COL_BORDER)])

    # Tombol "ghost" untuk header gelap
    style.configure("Ghost.TButton", font=(FONT, 10), padding=(14, 8), background=COL_HEADER,
                    foreground="#d7ece8", borderwidth=1, relief="flat", bordercolor="#2c6a6f",
                    lightcolor=COL_HEADER, darkcolor=COL_HEADER, focusthickness=0)
    style.map("Ghost.TButton",
              background=[("active", "#155e64"), ("pressed", "#0c3f44")],
              foreground=[("active", "#ffffff")],
              bordercolor=[("active", "#4f8f8f")])

    # Input
    style.configure("TEntry", fieldbackground="#ffffff", bordercolor=COL_BORDER_STRONG,
                    borderwidth=1, relief="flat", padding=6)
    style.configure("TCombobox", fieldbackground="#ffffff", bordercolor=COL_BORDER_STRONG,
                    borderwidth=1, padding=5)

    # Tabel
    style.configure("Treeview", background=COL_SURFACE, fieldbackground=COL_SURFACE,
                    foreground=COL_TEXT, rowheight=30, font=(FONT, 9), borderwidth=0)
    style.map("Treeview", background=[("selected", COL_BRAND)], foreground=[("selected", "#ffffff")])
    style.configure("Treeview.Heading", background=COL_SOFT, foreground=COL_MUTED,
                    font=(FONT, 9, "bold"), relief="flat", padding=(6, 7))
    style.map("Treeview.Heading", background=[("active", COL_BORDER)])

    # Progress bar
    style.configure("Horizontal.TProgressbar", background=COL_BRAND, troughcolor=COL_BORDER,
                    borderwidth=0, thickness=8)

    # Dashboard
    style.configure("MenuCard.TFrame", background=COL_SURFACE)
    style.configure("MenuTag.TLabel", background=COL_SURFACE, foreground=COL_BRAND, font=(FONT, 9, "bold"))
    style.configure("MenuTitle.TLabel", background=COL_SURFACE, foreground=COL_TEXT, font=(FONT_SB, 17))
    style.configure("MenuDesc.TLabel", background=COL_SURFACE, foreground=COL_MUTED, font=(FONT, 10))
    style.configure("DashFoot.TLabel", background=COL_BG, foreground=COL_MUTED, font=(FONT, 9))


def apply_button_cursors(parent: tk.Misc) -> None:
    for widget in parent.winfo_children():
        if isinstance(widget, ttk.Button):
            widget.configure(cursor="hand2")
        apply_button_cursors(widget)


class ModernButton(tk.Canvas):
    """Tombol canvas modern dengan warna konsisten dan state kompatibel."""

    PALETTES = {
        "primary": (COL_BRAND, COL_BRAND_HOVER, COL_BRAND_PRESS, "#ffffff", COL_BRAND),
        "blue": (COL_BLUE, "#3277bd", COL_BLUE_DARK, "#ffffff", COL_BLUE),
        "danger": (COL_DANGER, "#f05252", "#c92f3f", "#ffffff", COL_DANGER),
        "warning": ("#f1a52b", "#f5b548", "#c98012", "#ffffff", "#f1a52b"),
        "secondary": ("#e7eef6", "#dbe6f1", "#ccd9e7", COL_TEXT, "#cbd8e6"),
        "ghost_dark": ("#1b5965", "#246c79", "#0c3440", "#ffffff", "#397480"),
    }

    def __init__(
        self,
        master,
        text: str,
        command=None,
        variant: str = "primary",
        state: str = "normal",
        height: int = 44,
        width: int = 180,
        font=(FONT_SB, 10),
        surface: str = COL_SURFACE,
    ) -> None:
        super().__init__(
            master,
            height=height,
            width=width,
            bg=surface,
            highlightthickness=0,
            bd=0,
            relief="flat",
            takefocus=1,
        )
        self._text = text
        self._command = command
        self._variant = variant
        self._button_state = state
        self._visual_state = "normal"
        self._font = font
        self._radius = 11
        self.bind("<Configure>", self._redraw)
        self.bind("<Enter>", self._on_enter)
        self.bind("<Leave>", self._on_leave)
        self.bind("<ButtonPress-1>", self._on_press)
        self.bind("<ButtonRelease-1>", self._on_release)
        self.bind("<Return>", self._on_keyboard)
        self.bind("<space>", self._on_keyboard)
        self.bind("<FocusIn>", self._redraw)
        self.bind("<FocusOut>", self._redraw)
        self._sync_cursor()

    def configure(self, cnf=None, **kwargs):
        if cnf:
            kwargs.update(cnf)
        redraw = False
        if "state" in kwargs:
            self._button_state = kwargs.pop("state")
            self._visual_state = "normal"
            redraw = True
        if "text" in kwargs:
            self._text = kwargs.pop("text")
            redraw = True
        if "command" in kwargs:
            self._command = kwargs.pop("command")
        if "variant" in kwargs:
            self._variant = kwargs.pop("variant")
            redraw = True
        result = super().configure(**kwargs) if kwargs else None
        if redraw:
            self._sync_cursor()
            self._redraw()
        return result

    config = configure

    def cget(self, key):
        if key == "state":
            return self._button_state
        if key == "text":
            return self._text
        return super().cget(key)

    def _sync_cursor(self) -> None:
        super().configure(cursor="hand2" if self._button_state != "disabled" else "arrow")

    def _colors(self) -> tuple[str, str, str]:
        normal, hover, press, foreground, border = self.PALETTES[self._variant]
        if self._button_state == "disabled":
            disabled_palettes = {
                "primary": ("#dcefeb", "#76a9a1", "#c9e2dd"),
                "blue": ("#e5eef8", "#8aa4bf", "#d5e2ef"),
                "danger": ("#fbe6eb", "#c18491", "#f3d2d9"),
                "warning": ("#faedda", "#bd9662", "#f1dfc5"),
            }
            return disabled_palettes.get(self._variant, ("#edf1f5", "#98a2b3", "#e2e8f0"))
        fill = {"normal": normal, "hover": hover, "press": press}.get(self._visual_state, normal)
        return fill, foreground, border

    def _rounded_rect(self, x1: int, y1: int, x2: int, y2: int, radius: int, **kwargs) -> None:
        points = [
            x1 + radius, y1, x2 - radius, y1, x2, y1, x2, y1 + radius,
            x2, y2 - radius, x2, y2, x2 - radius, y2, x1 + radius, y2,
            x1, y2, x1, y2 - radius, x1, y1 + radius, x1, y1,
        ]
        self.create_polygon(points, smooth=True, splinesteps=24, **kwargs)

    def _redraw(self, _event=None) -> None:
        if not self.winfo_exists():
            return
        self.delete("all")
        width = max(2, self.winfo_width())
        height = max(2, self.winfo_height())
        fill, foreground, border = self._colors()
        focus_border = "#7dd3fc" if self.focus_get() == self and self._button_state != "disabled" else border
        self._rounded_rect(1, 1, width - 1, height - 1, self._radius, fill=focus_border, outline="")
        self._rounded_rect(2, 2, width - 2, height - 2, self._radius - 1, fill=fill, outline="")
        self.create_text(
            width // 2,
            height // 2,
            text=self._text,
            fill=foreground,
            font=self._font,
            anchor="center",
        )

    def _on_enter(self, _event=None) -> None:
        if self._button_state != "disabled":
            self._visual_state = "hover"
            self._redraw()

    def _on_leave(self, _event=None) -> None:
        self._visual_state = "normal"
        self._redraw()

    def _on_press(self, _event=None) -> None:
        if self._button_state != "disabled":
            self.focus_set()
            self._visual_state = "press"
            self._redraw()

    def _on_release(self, event) -> None:
        if self._button_state == "disabled":
            return
        inside = 0 <= event.x <= self.winfo_width() and 0 <= event.y <= self.winfo_height()
        self._visual_state = "hover" if inside else "normal"
        self._redraw()
        if inside and self._command:
            self._command()

    def _on_keyboard(self, _event=None) -> str:
        if self._button_state != "disabled" and self._command:
            self._command()
        return "break"


class RekapExcelApp(tk.Toplevel):
    def __init__(self, master: tk.Misc | None = None) -> None:
        super().__init__(master)
        self.title("Omset Lapangan — Rekap Excel Booking")
        self.protocol("WM_DELETE_WINDOW", self.back_to_dashboard)
        self.geometry(DEFAULT_APP_GEOMETRY)
        self.minsize(*DEFAULT_APP_MINSIZE)

        self.file_var = tk.StringVar()
        self.mode_choice_var = tk.StringVar(value="Otomatis")
        self.mode_buttons: dict[str, ModernButton] = {}
        self.status_var = tk.StringVar(value="Siap memproses rekap.")
        self.output_path: Path | None = None
        self.log_file_by_item: dict[str, str] = {}
        self.cancel_event = threading.Event()
        self.preview_source: dict[str, object] | None = None
        self.preview_path: Path | None = None
        self.preview_report_date: date | None = None
        self.preview_report_month: date | None = None
        self.preview_report_dates: list[date] | None = None
        self.preview_feature: str = ""

        self.metric_vars = {
            "database": tk.StringVar(value=str(db_record_count())),
            "data": tk.StringVar(value="-"),
            "ayo": tk.StringVar(value="-"),
            "walk_in": tk.StringVar(value="-"),
            "total": tk.StringVar(value="-"),
            "tanggal": tk.StringVar(value="-"),
        }
        self.preview_vars = {
            "mode": tk.StringVar(value="-"),
            "period": tk.StringVar(value="-"),
            "rows": tk.StringVar(value="-"),
            "dates": tk.StringVar(value="-"),
            "channels": tk.StringVar(value="-"),
            "courts": tk.StringVar(value="-"),
        }

        self.configure(bg=COL_CONTENT)
        self._configure_style()
        self._build_ui()
        apply_button_cursors(self)
        self.refresh_logs()

    def _configure_style(self) -> None:
        configure_styles(self)

    def back_to_dashboard(self) -> None:
        if self.master is not None:
            match_window_size(self.master, self)
        self.withdraw()
        if self.master is not None:
            self.master.deiconify()
            self.master.lift()
            self.master.focus_force()

    def _build_ui(self) -> None:
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)

        header = ttk.Frame(self, style="Header.TFrame", padding=(26, 18))
        header.grid(row=0, column=0, sticky="ew")
        header.columnconfigure(0, weight=1)
        ttk.Label(header, text="Omset Lapangan", style="Title.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(
            header,
            text="Upload data harian atau sebulan, periode dibaca otomatis, lalu workbook rekap dibuat lengkap.",
            style="Subtitle.TLabel",
        ).grid(row=1, column=0, sticky="w", pady=(4, 0))
        ModernButton(
            header,
            text="←  Kembali ke Dashboard",
            command=self.back_to_dashboard,
            variant="ghost_dark",
            height=40,
            width=205,
            surface=COL_HEADER,
        ).grid(row=0, column=1, rowspan=2, sticky="e")

        main = ttk.Frame(self, style="App.TFrame", padding=18)
        main.grid(row=1, column=0, sticky="nsew")
        main.columnconfigure(0, weight=1)
        main.columnconfigure(1, weight=1)
        main.rowconfigure(1, weight=1)

        input_card = ttk.Frame(main, style="Elevated.TFrame", padding=18)
        input_card.grid(row=0, column=0, sticky="nsew", padx=(0, 9), pady=(0, 14))
        input_card.columnconfigure(1, weight=1)

        ttk.Label(input_card, text="1. Pilih Data Booking", style="CardTitle.TLabel").grid(
            row=0, column=0, columnspan=3, sticky="w"
        )
        file_box = tk.Frame(input_card, bg="#f6f9fd", highlightthickness=1, highlightbackground="#cbdced")
        file_box.grid(row=1, column=0, columnspan=3, sticky="ew", pady=(10, 10))
        file_box.columnconfigure(1, weight=1)
        tk.Label(
            file_box,
            text="XLSX",
            bg="#e4eef9",
            fg=COL_BLUE,
            font=(FONT_SB, 9),
            padx=10,
            pady=7,
        ).grid(row=0, column=0, padx=(10, 8), pady=10)
        ttk.Entry(file_box, textvariable=self.file_var).grid(row=0, column=1, sticky="ew", pady=10)
        ModernButton(
            file_box,
            text="Pilih File Excel",
            command=self.choose_file,
            variant="blue",
            height=40,
            width=150,
            surface="#f6f9fd",
        ).grid(row=0, column=2, sticky="e", padx=10, pady=7)
        ttk.Label(
            input_card,
            text="Tanpa memilih file, aplikasi akan memakai data terakhir yang tersimpan di database.",
            style="Hint.TLabel",
        ).grid(
            row=2, column=0, columnspan=3, sticky="w", pady=(0, 12)
        )

        ttk.Label(input_card, text="2. Atur Jenis Output", style="CardTitle.TLabel").grid(row=3, column=0, sticky="w")
        auto_box = tk.Frame(input_card, bg="#f8fafc", bd=0, highlightthickness=1, highlightbackground="#dbe3e8")
        auto_box.grid(row=4, column=0, columnspan=3, sticky="ew", pady=(7, 14))
        ttk.Label(
            auto_box,
            text="Tanggal, bulan, dan tahun dibaca dari kolom Date of Booking.",
            style="MetricTitle.TLabel",
        ).pack(anchor="w", padx=12, pady=(10, 2))
        ttk.Label(
            auto_box,
            text="Otomatis: 1 tanggal = harian, beberapa tanggal = multi hari, data banyak = sebulan.",
            style="AutoHint.TLabel",
        ).pack(anchor="w", padx=12, pady=(0, 10))
        mode_line = tk.Frame(auto_box, bg="#f8fafc")
        mode_line.pack(fill="x", padx=12, pady=(0, 12))
        tk.Label(
            mode_line, text="JENIS OUTPUT", bg="#f8fafc", fg=COL_MUTED, font=(FONT_SB, 8)
        ).pack(anchor="w", pady=(0, 7))
        choice_line = tk.Frame(mode_line, bg="#f8fafc")
        choice_line.pack(fill="x")
        for value in ("Otomatis", "Multi Hari", "Sebulan"):
            choice_line.columnconfigure(len(self.mode_buttons), weight=1)
            button = ModernButton(
                choice_line,
                text=value,
                command=lambda selected=value: self.set_mode_choice(selected),
                variant="primary" if value == self.mode_choice_var.get() else "secondary",
                height=38,
                width=120,
                surface="#f8fafc",
                font=(FONT_SB, 9),
            )
            button.grid(row=0, column=len(self.mode_buttons), sticky="ew", padx=(0 if not self.mode_buttons else 6, 0))
            self.mode_buttons[value] = button

        ttk.Label(input_card, text="3. Periksa Preview", style="CardTitle.TLabel").grid(row=5, column=0, sticky="w")
        preview_box = tk.Frame(input_card, bg="#f8fafc", bd=0, highlightthickness=1, highlightbackground="#dbe3e8")
        preview_box.grid(row=6, column=0, columnspan=3, sticky="ew", pady=(7, 14))
        for col_idx in range(3):
            preview_box.columnconfigure(col_idx, weight=1)
        preview_items = [
            ("Mode", "mode"),
            ("Periode", "period"),
            ("Data", "rows"),
            ("Tanggal", "dates"),
            ("Channel", "channels"),
            ("Court", "courts"),
        ]
        for index, (label, key) in enumerate(preview_items):
            row = index // 3
            col = index % 3
            cell = tk.Frame(preview_box, bg="#f8fafc")
            cell.grid(row=row, column=col, sticky="ew", padx=12, pady=(10 if row == 0 else 4, 10))
            ttk.Label(cell, text=label.upper(), style="MetricTitle.TLabel").pack(anchor="w")
            ttk.Label(cell, textvariable=self.preview_vars[key], style="PreviewValue.TLabel").pack(anchor="w", pady=(2, 0))

        actions = ttk.Frame(input_card, style="Card.TFrame")
        actions.grid(row=7, column=0, columnspan=3, sticky="ew")
        actions.columnconfigure((0, 1, 2), weight=1)
        self.process_button = ModernButton(
            actions, text="Buat Rekap Excel", command=self.process_file, variant="primary", height=46
        )
        self.process_button.grid(row=0, column=0, sticky="ew", padx=(0, 8), pady=(0, 8))
        self.cancel_button = ModernButton(
            actions, text="Batalkan", command=self.cancel_processing, variant="danger", state="disabled", height=46
        )
        self.cancel_button.grid(row=0, column=1, sticky="ew", padx=(0, 8), pady=(0, 8))
        self.open_button = ModernButton(
            actions, text="Buka File Hasil", command=self.open_output, variant="blue", state="disabled", height=46
        )
        self.open_button.grid(row=0, column=2, sticky="ew", pady=(0, 8))
        self.output_folder_button = ModernButton(
            actions, text="Buka Folder Output", command=self.open_output_folder, variant="secondary", height=42
        )
        self.output_folder_button.grid(row=1, column=0, columnspan=2, sticky="ew", padx=(0, 8))
        self.reset_button = ModernButton(
            actions, text="Reset Database", command=self.reset_app_data, variant="warning", height=42
        )
        self.reset_button.grid(row=1, column=2, sticky="ew")

        result_card = ttk.Frame(main, style="Elevated.TFrame", padding=18)
        result_card.grid(row=0, column=1, sticky="nsew", padx=(9, 0), pady=(0, 14))
        result_card.columnconfigure((0, 1, 2), weight=1)
        ttk.Label(result_card, text="Ringkasan Output", style="CardTitle.TLabel").grid(
            row=0, column=0, columnspan=3, sticky="w"
        )
        tk.Label(
            result_card,
            text="LIVE",
            bg="#e8f7f3",
            fg=COL_BRAND,
            font=(FONT_SB, 8),
            padx=9,
            pady=4,
        ).grid(row=0, column=2, sticky="e")

        metrics = [
            ("Tanggal", "tanggal", COL_BLUE, "#f2f7fc"),
            ("Data Masuk", "data", "#7657b5", "#f7f4fc"),
            ("Database", "database", "#64748b", "#f7f9fb"),
            ("AYO", "ayo", "#2878c7", "#f2f7fc"),
            ("Walk In", "walk_in", COL_AMBER, "#fff8ec"),
            ("Total", "total", COL_BRAND, "#eef9f6"),
        ]
        for index, (label, key, accent, soft_bg) in enumerate(metrics):
            outer = tk.Frame(result_card, bg=COL_SURFACE)
            outer.grid(row=1 + index // 3, column=index % 3, sticky="nsew", padx=5, pady=9)
            outer.columnconfigure(1, weight=1)
            tk.Frame(outer, bg=accent, width=4).grid(row=0, column=0, sticky="ns")
            inner = tk.Frame(outer, bg=soft_bg, highlightthickness=1, highlightbackground=COL_BORDER)
            inner.grid(row=0, column=1, sticky="nsew")
            tk.Label(inner, text=label.upper(), bg=soft_bg, fg=COL_MUTED, font=(FONT_SB, 8)).pack(
                anchor="w", padx=12, pady=(10, 3)
            )
            tk.Label(inner, textvariable=self.metric_vars[key], bg=soft_bg, fg=COL_TEXT, font=(FONT_SB, 16)).pack(
                anchor="w", padx=12, pady=(0, 11)
            )

        status_frame = tk.Frame(
            result_card, bg="#eff9f7", highlightthickness=1, highlightbackground="#c9e8e1"
        )
        status_frame.grid(row=3, column=0, columnspan=3, sticky="ew", padx=5, pady=(12, 0))
        status_frame.columnconfigure(0, weight=1)
        tk.Label(
            status_frame, text="STATUS PROSES", bg="#eff9f7", fg=COL_BRAND, font=(FONT_SB, 8)
        ).grid(row=0, column=0, sticky="w", padx=14, pady=(12, 3))
        tk.Label(
            status_frame,
            textvariable=self.status_var,
            bg="#eff9f7",
            fg=COL_TEXT,
            font=(FONT, 9),
            anchor="w",
            justify="left",
        ).grid(row=1, column=0, sticky="ew", padx=14)
        self.progress_bar = ttk.Progressbar(status_frame, mode="indeterminate")
        self.progress_bar.grid(row=2, column=0, sticky="ew", padx=14, pady=(10, 13))
        tk.Label(
            result_card,
            text="Workbook hasil tersimpan otomatis dan dapat dibuka kembali dari log di bawah.",
            bg=COL_SURFACE,
            fg=COL_MUTED,
            font=(FONT, 9),
            anchor="w",
        ).grid(row=4, column=0, columnspan=3, sticky="ew", padx=5, pady=(12, 0))

        result_card.rowconfigure(5, weight=1)
        workbook_box = tk.Frame(
            result_card, bg="#f8fafc", highlightthickness=1, highlightbackground=COL_BORDER
        )
        workbook_box.grid(row=5, column=0, columnspan=3, sticky="nsew", padx=5, pady=(14, 0))
        workbook_box.columnconfigure((0, 1), weight=1)
        tk.Label(
            workbook_box,
            text="ISI WORKBOOK HASIL",
            bg="#f8fafc",
            fg=COL_MUTED,
            font=(FONT_SB, 8),
        ).grid(row=0, column=0, columnspan=2, sticky="w", padx=14, pady=(12, 7))
        for index, (label, accent) in enumerate(
            (("Summary Omset", COL_BRAND), ("Walk In", COL_AMBER), ("AYO Booking", COL_BLUE), ("Data Keseluruhan", "#7657b5"))
        ):
            chip = tk.Frame(workbook_box, bg=COL_SURFACE, highlightthickness=1, highlightbackground=COL_BORDER)
            chip.grid(
                row=1 + index // 2,
                column=index % 2,
                sticky="ew",
                padx=(14 if index % 2 == 0 else 5, 5 if index % 2 == 0 else 14),
                pady=(0, 7 if index < 2 else 12),
            )
            tk.Label(chip, text="●", bg=COL_SURFACE, fg=accent, font=(FONT, 9)).pack(side="left", padx=(10, 7), pady=8)
            tk.Label(chip, text=label, bg=COL_SURFACE, fg=COL_TEXT, font=(FONT_SB, 9)).pack(
                side="left", pady=8
            )

        log_card = ttk.Frame(main, style="Elevated.TFrame", padding=18)
        log_card.grid(row=1, column=0, columnspan=2, sticky="nsew")
        log_card.columnconfigure(0, weight=1)
        log_card.rowconfigure(1, weight=1)

        log_header = ttk.Frame(log_card, style="Card.TFrame")
        log_header.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        log_header.columnconfigure(0, weight=1)
        ttk.Label(log_header, text="Log Rekap Selesai", style="Body.TLabel").grid(row=0, column=0, sticky="w")
        ModernButton(
            log_header, text="Muat Ulang", command=self.refresh_logs, variant="secondary", height=38, width=115
        ).grid(row=0, column=1, sticky="e", padx=(0, 8))
        ModernButton(
            log_header, text="Buka File Terpilih", command=self.open_selected_log, variant="blue", height=38, width=155
        ).grid(row=0, column=2, sticky="e")
        ModernButton(
            log_header, text="Hapus Log", command=self.clear_logs, variant="danger", height=38, width=110
        ).grid(row=0, column=3, sticky="e", padx=(8, 0))

        columns = ("processed_at", "feature", "period", "source", "included_rows", "ayo", "walk_in", "total", "file")
        self.log_tree = ttk.Treeview(log_card, columns=columns, show="headings", selectmode="browse")
        headings = {
            "processed_at": "Waktu Proses",
            "feature": "Mode",
            "period": "Periode",
            "source": "File Sumber",
            "included_rows": "Data",
            "ayo": "AYO",
            "walk_in": "Walk In",
            "total": "Total",
            "file": "File",
        }
        widths = {
            "processed_at": 145,
            "feature": 80,
            "period": 95,
            "source": 220,
            "included_rows": 65,
            "ayo": 65,
            "walk_in": 75,
            "total": 130,
            "file": 210,
        }
        for column in columns:
            self.log_tree.heading(column, text=headings[column])
            self.log_tree.column(column, width=widths[column], minwidth=60, anchor="w")
        self.log_tree.grid(row=1, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(log_card, orient="vertical", command=self.log_tree.yview)
        scrollbar.grid(row=1, column=1, sticky="ns")
        self.log_tree.configure(yscrollcommand=scrollbar.set)

    def choose_file(self) -> None:
        path = filedialog.askopenfilename(
            title="Pilih file Excel mentah",
            filetypes=[("Excel files", "*.xlsx *.xlsm"), ("All files", "*.*")],
        )
        if path:
            self.file_var.set(path)
            self.start_preview(Path(path))

    def reset_preview(self) -> None:
        self.preview_source = None
        self.preview_path = None
        self.preview_report_date = None
        self.preview_report_month = None
        self.preview_report_dates = None
        self.preview_feature = ""
        for variable in self.preview_vars.values():
            variable.set("-")

    def set_mode_choice(self, value: str) -> None:
        self.mode_choice_var.set(value)
        for choice, button in self.mode_buttons.items():
            button.configure(variant="primary" if choice == value else "secondary")
        self.on_mode_choice_changed()

    def on_mode_choice_changed(self, _event=None) -> None:
        path_text = self.file_var.get().strip()
        if path_text and Path(path_text).exists():
            self.start_preview(Path(path_text))

    def start_preview(self, path: Path) -> None:
        if not path.exists():
            return
        self.cancel_event.clear()
        self.reset_preview()
        self.status_var.set("Membaca preview file...")
        self.set_busy_state(True)
        mode_choice = self.mode_choice_var.get()
        thread = threading.Thread(target=self._preview_worker, args=(path, mode_choice), daemon=True)
        thread.start()

    def _preview_worker(self, path: Path, mode_choice: str) -> None:
        try:
            with path.open("rb") as file_obj:
                source = extract_records(file_obj, should_cancel=self.cancel_event.is_set)
            report_date, report_month, report_dates, feature, _ = infer_report_period(source["records"], mode_choice)
            preview = preview_summary(source, report_date, report_month, report_dates, feature)
            self.after(
                0,
                lambda: self.finish_preview_success(
                    path, source, report_date, report_month, report_dates, feature, preview
                ),
            )
        except ProcessCancelled:
            self.after(0, self.finish_cancelled)
        except Exception as exc:
            message = str(exc)
            self.after(0, lambda message=message: self.finish_preview_error(message))

    def finish_preview_success(
        self,
        path: Path,
        source: dict[str, object],
        report_date: date | None,
        report_month: date | None,
        report_dates: list[date] | None,
        feature: str,
        preview: dict[str, str],
    ) -> None:
        self.preview_path = path
        self.preview_source = source
        self.preview_report_date = report_date
        self.preview_report_month = report_month
        self.preview_report_dates = report_dates
        self.preview_feature = feature
        for key, value in preview.items():
            self.preview_vars[key].set(value)
        self.status_var.set("Preview siap. Periksa periode dan jumlah data sebelum proses.")
        self.set_busy_state(False)

    def finish_preview_error(self, message: str) -> None:
        self.reset_preview()
        self.status_var.set(message)
        self.set_busy_state(False)
        messagebox.showerror("Preview gagal", message)

    def can_process_current_input(self) -> bool:
        path_text = self.file_var.get().strip()
        if path_text:
            path = Path(path_text)
            return self.preview_path == path and self.preview_source is not None
        return db_record_count() > 0

    def process_file(self) -> None:
        path_text = self.file_var.get().strip()
        if path_text and not Path(path_text).exists():
            messagebox.showerror("File tidak ditemukan", "Path file Excel tidak ditemukan.")
            return
        if path_text and not self.can_process_current_input():
            self.start_preview(Path(path_text))
            self.status_var.set("Preview file belum valid. Tunggu preview sukses sebelum proses rekap.")
            return
        if not path_text and db_record_count() == 0:
            messagebox.showerror("Database kosong", "Upload file Excel mentah terlebih dahulu.")
            return

        self.cancel_event.clear()
        self.set_busy_state(True)
        self.status_var.set("Membaca tanggal dari file dan membuat workbook...")
        thread = threading.Thread(target=self._process_worker, args=(path_text,), daemon=True)
        thread.start()

    def _process_worker(self, path_text: str) -> None:
        try:
            uploaded_records = []
            report_date = None
            report_month = None
            report_dates = None
            feature = ""
            period_message = ""
            if path_text:
                path = Path(path_text)
                if self.preview_path == path and self.preview_source is not None:
                    uploaded_source = self.preview_source
                    report_date = self.preview_report_date
                    report_month = self.preview_report_month
                    report_dates = self.preview_report_dates
                    feature = self.preview_feature
                    period_message = f"{feature}: {format_period_label(report_date, report_month, report_dates)}"
                else:
                    self.after(0, lambda: self.status_var.set("Membaca file Excel..."))
                    with path.open("rb") as file_obj:
                        uploaded_source = extract_records(file_obj, should_cancel=self.cancel_event.is_set)
                    report_date, report_month, report_dates, feature, period_message = infer_report_period(
                        uploaded_source["records"],
                        self.mode_choice_var.get(),
                    )
                uploaded_records = uploaded_source["records"]
                if self.cancel_event.is_set():
                    raise ProcessCancelled("Proses dibatalkan.")
                self.after(0, lambda: self.status_var.set("Menyimpan data ke database lokal..."))
                save_records_to_db(uploaded_records, path.name, should_cancel=self.cancel_event.is_set)
            else:
                recent_months = db_recent_months(1)
                if not recent_months:
                    raise ValueError("Belum ada data bulan yang tersimpan di database.")
                report_month = recent_months[0]
                feature = "Sebulan"
                period_message = f"Menggunakan bulan terbaru dari database: {report_month:%Y-%m}"

            if self.cancel_event.is_set():
                raise ProcessCancelled("Proses dibatalkan.")
            self.after(0, lambda: self.status_var.set("Membuat workbook hasil rekap..."))
            records = load_records_from_db(report_date, report_month, report_dates)
            output, stats = build_workbook(
                {"headers": RAW_HEADERS, "records": records},
                report_date=report_date,
                report_month=report_month,
                report_dates=report_dates,
                should_cancel=self.cancel_event.is_set,
            )
            stats["database_rows"] = db_record_count()
            stats["period_message"] = period_message
            stats["source_file"] = Path(path_text).name if path_text else "Database lokal"
            stats["feature"] = feature or ("Harian" if report_date else "Sebulan")
            stats["period_label"] = format_period_label(report_date, report_month, report_dates)
            save_process_log(stats)
            self.output_path = output
            self.after(0, lambda: self.finish_success(stats))
        except ProcessCancelled:
            self.after(0, self.finish_cancelled)
        except Exception as exc:
            message = str(exc)
            self.after(0, lambda message=message: self.finish_error(message))

    def cancel_processing(self) -> None:
        self.cancel_event.set()
        self.cancel_button.configure(state="disabled")
        self.status_var.set("Membatalkan proses...")

    def set_busy_state(self, is_busy: bool) -> None:
        if is_busy:
            self.process_button.configure(state="disabled")
            self.open_button.configure(state="disabled")
            self.reset_button.configure(state="disabled")
            self.output_folder_button.configure(state="disabled")
            self.cancel_button.configure(state="normal")
            self.progress_bar.start(12)
            return

        self.progress_bar.stop()
        self.progress_bar.configure(value=0)
        self.process_button.configure(state="normal" if self.can_process_current_input() else "disabled")
        self.reset_button.configure(state="normal")
        self.output_folder_button.configure(state="normal")
        self.cancel_button.configure(state="disabled")
        self.open_button.configure(state="normal" if self.output_path and self.output_path.exists() else "disabled")

    def finish_success(self, stats: dict[str, object]) -> None:
        self.metric_vars["database"].set(str(stats["database_rows"]))
        self.metric_vars["data"].set(str(stats["included_rows"]))
        self.metric_vars["ayo"].set(str(stats["ayo_rows"]))
        self.metric_vars["walk_in"].set(str(stats["walk_in_rows"]))
        self.metric_vars["total"].set(format_rupiah(stats["total_revenue"]))
        self.metric_vars["tanggal"].set(str(stats.get("period_label") or stats["selected_date"] or stats["month"]))
        status = f"{stats.get('period_message', 'Selesai')}. File dibuat: {stats['filename']}"
        duplicates = stats.get("duplicate_booking_ids") or {}
        if duplicates:
            status += f". Peringatan: {len(duplicates)} Booking ID duplikat tetap dihitung."
        self.status_var.set(status)
        self.set_busy_state(False)
        self.refresh_logs()
        warning_text = duplicate_booking_warning_text(duplicates)
        if warning_text:
            messagebox.showwarning("Booking ID duplikat", warning_text)

    def finish_error(self, message: str) -> None:
        self.status_var.set(message)
        self.set_busy_state(False)
        messagebox.showerror("Gagal memproses", message)

    def finish_cancelled(self) -> None:
        self.status_var.set("Proses dibatalkan.")
        self.set_busy_state(False)

    def clear_logs(self) -> None:
        deleted = clear_process_logs()
        self.refresh_logs()
        self.status_var.set(f"Log rekap dikosongkan. {deleted} log dihapus.")

    def reset_app_data(self) -> None:
        confirmed = messagebox.askyesno(
            "Reset data",
            "Reset akan menghapus data upload tersimpan dan semua log rekap. File hasil di folder outputs tidak dihapus. Lanjutkan?",
        )
        if not confirmed:
            return

        deleted = reset_database()
        self.file_var.set("")
        self.output_path = None
        self.reset_preview()
        self.open_button.configure(state="disabled")
        self.metric_vars["database"].set("0")
        self.metric_vars["data"].set("-")
        self.metric_vars["ayo"].set("-")
        self.metric_vars["walk_in"].set("-")
        self.metric_vars["total"].set("-")
        self.metric_vars["tanggal"].set("-")
        self.refresh_logs()
        self.status_var.set(
            f"Reset selesai. {deleted['bookings']} data upload dan {deleted['logs']} log dihapus."
        )

    def refresh_logs(self) -> None:
        for item in self.log_tree.get_children():
            self.log_tree.delete(item)
        self.log_file_by_item.clear()

        for log in load_process_logs(50):
            period = log.get("period_label") or log.get("report_date") or log.get("report_month") or "-"
            item = self.log_tree.insert(
                "",
                "end",
                values=(
                    log["processed_at"],
                    log.get("feature") or "-",
                    period,
                    log.get("source_file") or "-",
                    log["included_rows"],
                    log["ayo_rows"],
                    log["walk_in_rows"],
                    format_rupiah(log["total_revenue"]),
                    log["output_file"],
                ),
            )
            self.log_file_by_item[item] = log["output_file"]

    def open_output(self) -> None:
        if self.output_path and self.output_path.exists():
            os.startfile(self.output_path)

    def open_output_folder(self) -> None:
        OUTPUT_DIR.mkdir(exist_ok=True)
        os.startfile(OUTPUT_DIR)

    def open_selected_log(self) -> None:
        selected = self.log_tree.selection()
        if not selected:
            messagebox.showinfo("Pilih log", "Pilih salah satu baris log terlebih dahulu.")
            return
        filename = self.log_file_by_item.get(selected[0])
        if not filename:
            return
        path = OUTPUT_DIR / Path(filename).name
        if not path.exists():
            messagebox.showerror("File tidak ditemukan", f"File hasil tidak ditemukan:\n{path}")
            return
        os.startfile(path)


class CategoryUploadDialog(tk.Toplevel):
    def __init__(self, master: tk.Misc | None = None) -> None:
        super().__init__(master)
        self.title("Omset Perkategori Olsera")
        self.protocol("WM_DELETE_WINDOW", self.back_to_dashboard)
        self.geometry(DEFAULT_APP_GEOMETRY)
        self.minsize(*DEFAULT_APP_MINSIZE)
        self.configure(bg=COL_CONTENT)

        self.file_paths: list[Path] = []
        self.cancel_event = threading.Event()
        self.status_var = tk.StringVar(value="Tambahkan file Excel kategori, lalu tekan Cek Group.")
        self.last_results: list[dict[str, object]] = []
        self.last_duplicates: dict[str, list[str]] = {}
        self.last_paths: list[Path] = []
        self.check_passed = False
        self.preview_item_sources: dict[str, str] = {}
        self.preview_context_item: str | None = None
        self.preview_columns: tuple[str, ...] = ()
        self.preview_headings: dict[str, str] = {}
        self.preview_sort_column: str | None = None
        self.preview_sort_desc = False
        self.category_stat_vars = {
            "files": tk.StringVar(value="0"),
            "groups": tk.StringVar(value="-"),
            "rows": tk.StringVar(value="-"),
            "amount": tk.StringVar(value="-"),
        }

        self._build_ui()
        apply_button_cursors(self)

    def back_to_dashboard(self) -> None:
        if self.master is not None:
            match_window_size(self.master, self)
        self.withdraw()
        if self.master is not None:
            self.master.deiconify()
            self.master.lift()
            self.master.focus_force()

    def _build_ui(self) -> None:
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)

        header = ttk.Frame(self, style="Header.TFrame", padding=(22, 18))
        header.grid(row=0, column=0, sticky="ew")
        header.columnconfigure(0, weight=1)
        ttk.Label(header, text="Omset Perkategori Olsera", style="Title.TLabel").grid(
            row=0, column=0, sticky="w"
        )
        ttk.Label(
            header,
            text="Validasi item group dan duplikat, lalu buat workbook laporan bulanan yang siap digunakan.",
            style="Subtitle.TLabel",
        ).grid(row=1, column=0, sticky="w", pady=(4, 0))
        ttk.Button(
            header, text="←  Kembali ke Dashboard", style="Ghost.TButton", command=self.back_to_dashboard
        ).grid(row=0, column=1, rowspan=2, sticky="e")

        shell = ttk.Frame(self, style="Content.TFrame", padding=(24, 20))
        shell.grid(row=1, column=0, sticky="nsew")
        shell.columnconfigure(0, weight=1)
        shell.rowconfigure(1, weight=1)

        stats = ttk.Frame(shell, style="Content.TFrame")
        stats.grid(row=0, column=0, sticky="ew", pady=(0, 16))
        for col in range(4):
            stats.columnconfigure(col, weight=1, uniform="category_stats")
        self._category_stat_card(stats, 0, "File", self.category_stat_vars["files"], "dipilih", COL_BLUE)
        self._category_stat_card(stats, 1, "Kategori", self.category_stat_vars["groups"], "hasil cek", COL_GREEN)
        self._category_stat_card(stats, 2, "Data", self.category_stat_vars["rows"], "baris valid", COL_AMBER)
        self._category_stat_card(stats, 3, "Nominal", self.category_stat_vars["amount"], "hasil cek", COL_BRAND)

        workspace = ttk.Frame(shell, style="Content.TFrame")
        workspace.grid(row=1, column=0, sticky="nsew")
        workspace.columnconfigure(0, weight=0)
        workspace.columnconfigure(1, weight=1)
        workspace.rowconfigure(0, weight=1)

        side = ttk.Frame(workspace, style="Card.TFrame", padding=16)
        side.grid(row=0, column=0, sticky="ns", padx=(0, 14))
        side.grid_propagate(False)
        side.configure(width=320)
        side.columnconfigure(0, weight=1)
        side.rowconfigure(2, weight=1)

        ttk.Label(side, text="File Kategori", style="CardTitle.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(
            side,
            text="Pilih export Olsera yang punya item group. File SEWA RAKET akan dipisah dari item name.",
            style="CardSub.TLabel",
            wraplength=270,
            justify="left",
        ).grid(row=1, column=0, sticky="w", pady=(4, 12))

        list_card = tk.Frame(side, bg=COL_SOFT, highlightthickness=1, highlightbackground=COL_BORDER)
        list_card.grid(row=2, column=0, sticky="nsew", pady=(0, 12))
        list_card.columnconfigure(0, weight=1)
        list_card.rowconfigure(0, weight=1)
        self.file_listbox = tk.Listbox(
            list_card,
            height=12,
            font=(FONT, 9),
            bg=COL_SOFT,
            fg=COL_TEXT,
            relief="flat",
            highlightthickness=0,
            activestyle="none",
            selectbackground=COL_BLUE,
            selectforeground="#ffffff",
        )
        self.file_listbox.grid(row=0, column=0, sticky="nsew", padx=8, pady=8)
        list_scrollbar = ttk.Scrollbar(list_card, orient="vertical", command=self.file_listbox.yview)
        list_scrollbar.grid(row=0, column=1, sticky="ns")
        self.file_listbox.configure(yscrollcommand=list_scrollbar.set)

        file_actions = ttk.Frame(side, style="Card.TFrame")
        file_actions.grid(row=3, column=0, sticky="ew", pady=(0, 12))
        file_actions.columnconfigure(0, weight=1)
        file_actions.columnconfigure(1, weight=1)
        self.add_button = ttk.Button(file_actions, text="+  Pilih File", style="Blue.TButton", command=self.add_files)
        self.add_button.grid(row=0, column=0, sticky="ew", padx=(0, 6))
        self.remove_button = ttk.Button(
            file_actions, text="Hapus Pilihan", style="Secondary.TButton", command=self.remove_selected_file
        )
        self.remove_button.grid(row=0, column=1, sticky="ew", padx=(6, 0))

        process_actions = ttk.Frame(side, style="Card.TFrame")
        process_actions.grid(row=4, column=0, sticky="ew")
        process_actions.columnconfigure(0, weight=1)
        self.check_button = ttk.Button(
            process_actions, text="1. Periksa Item Group", style="Blue.TButton", command=self.check_group, state="disabled"
        )
        self.check_button.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        self.delete_duplicates_button = ttk.Button(
            process_actions,
            text="Hapus Semua Duplikat",
            style="Danger.TButton",
            command=self.delete_all_duplicate_files,
            state="disabled",
        )
        self.delete_duplicates_button.grid(row=1, column=0, sticky="ew", pady=(0, 8))
        self.olah_button = ttk.Button(
            process_actions,
            text="2. Buat Workbook Excel",
            style="PrimaryLarge.TButton",
            command=self.olah_data,
            state="disabled",
        )
        self.olah_button.grid(row=2, column=0, sticky="ew", pady=(0, 8))
        self.download_button = ttk.Button(
            process_actions,
            text="Simpan Salinan Summary",
            style="Secondary.TButton",
            command=self.download_result,
            state="disabled",
        )
        self.download_button.grid(row=3, column=0, sticky="ew", pady=(0, 8))
        self.cancel_button = ttk.Button(
            process_actions, text="Batalkan Proses", style="Danger.TButton", command=self.cancel_processing, state="disabled"
        )
        self.cancel_button.grid(row=4, column=0, sticky="ew")

        status_card = ttk.Frame(side, style="Card.TFrame")
        status_card.grid(row=5, column=0, sticky="ew", pady=(14, 0))
        status_card.columnconfigure(0, weight=1)
        ttk.Label(status_card, textvariable=self.status_var, style="Hint.TLabel", wraplength=270).grid(
            row=0, column=0, sticky="w"
        )
        self.progress_bar = ttk.Progressbar(status_card, mode="indeterminate")
        self.progress_bar.grid(row=1, column=0, sticky="ew", pady=(10, 0))

        tree_card = ttk.Frame(workspace, style="Card.TFrame", padding=16)
        tree_card.grid(row=0, column=1, sticky="nsew")
        tree_card.columnconfigure(0, weight=1)
        tree_card.rowconfigure(1, weight=1)
        table_header = ttk.Frame(tree_card, style="Card.TFrame")
        table_header.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        table_header.columnconfigure(0, weight=1)
        ttk.Label(table_header, text="Preview Validasi Kategori", style="CardTitle.TLabel").grid(
            row=0, column=0, sticky="w"
        )
        ttk.Label(
            table_header,
            text="Hapus massal mempertahankan file pertama tiap group; klik kanan untuk memilih manual.",
            style="CardSub.TLabel",
        ).grid(row=1, column=0, sticky="w")

        columns = ("no", "file", "kolom", "kategori", "tanggal", "jumlah", "nominal", "status")
        self.preview_columns = columns
        self.tree = ttk.Treeview(tree_card, columns=columns, show="headings", selectmode="browse")
        headings = {
            "no": "No",
            "file": "File",
            "kolom": "Kolom",
            "kategori": "Group",
            "tanggal": "Tanggal",
            "jumlah": "Jumlah Data",
            "nominal": "Nominal",
            "status": "Status",
        }
        self.preview_headings = headings
        widths = {
            "no": 46,
            "file": 180,
            "kolom": 90,
            "kategori": 210,
            "tanggal": 170,
            "jumlah": 90,
            "nominal": 120,
            "status": 90,
        }
        for column in columns:
            self.tree.heading(
                column,
                text=headings[column],
                command=lambda selected_column=column: self.sort_preview_column(selected_column),
            )
            anchor = "center" if column in ("no", "jumlah", "nominal", "status") else "w"
            self.tree.column(column, width=widths[column], minwidth=50, anchor=anchor, stretch=False)
        self.tree.grid(row=1, column=0, sticky="nsew")
        y_scrollbar = ttk.Scrollbar(tree_card, orient="vertical", command=self.tree.yview)
        y_scrollbar.grid(row=1, column=1, sticky="ns")
        x_scrollbar = ttk.Scrollbar(tree_card, orient="horizontal", command=self.tree.xview)
        x_scrollbar.grid(row=2, column=0, sticky="ew")
        self.tree.configure(yscrollcommand=y_scrollbar.set, xscrollcommand=x_scrollbar.set)
        self.tree.tag_configure("dup", background="#fef3c7", foreground="#92400e")
        self.tree.tag_configure("invalid", background="#fee2e2", foreground="#991b1b")
        self.tree.bind("<Button-3>", self.show_duplicate_context_menu)

        self.duplicate_menu = tk.Menu(self, tearoff=0)
        self.duplicate_menu.add_command(
            label="Hapus file duplikat dari olah data",
            command=self.delete_context_duplicate_file,
        )

    def _category_stat_card(
        self,
        parent,
        column: int,
        title: str,
        variable: tk.StringVar,
        caption: str,
        accent: str,
    ) -> None:
        card = tk.Frame(parent, bg=COL_SURFACE, highlightthickness=1, highlightbackground=COL_BORDER)
        card.grid(row=0, column=column, sticky="nsew", padx=(0 if column == 0 else 10, 0))
        card.columnconfigure(1, weight=1)
        tk.Frame(card, bg=accent, width=5).grid(row=0, column=0, rowspan=3, sticky="ns")
        tk.Label(card, text=title.upper(), bg=COL_SURFACE, fg=COL_MUTED, font=(FONT, 8, "bold")).grid(
            row=0, column=1, sticky="w", padx=14, pady=(12, 2)
        )
        tk.Label(
            card,
            textvariable=variable,
            bg=COL_SURFACE,
            fg=COL_TEXT,
            font=(FONT_SB, 16),
            anchor="w",
            justify="left",
            wraplength=190,
        ).grid(
            row=1, column=1, sticky="ew", padx=14
        )
        tk.Label(card, text=caption, bg=COL_SURFACE, fg=COL_MUTED, font=(FONT, 9)).grid(
            row=2, column=1, sticky="w", padx=14, pady=(2, 12)
        )

    def update_category_stats(self, results: list[dict[str, object]] | None = None) -> None:
        self.category_stat_vars["files"].set(str(len(self.file_paths)))
        if not results:
            self.category_stat_vars["groups"].set("-")
            self.category_stat_vars["rows"].set("-")
            self.category_stat_vars["amount"].set("-")
            return

        categories: set[str] = set()
        total_rows = 0
        total_amount = 0.0
        for result in results:
            result_categories = result.get("categories") if isinstance(result.get("categories"), dict) else {}
            result_amounts = result.get("amounts") if isinstance(result.get("amounts"), dict) else {}
            for category, count in result_categories.items():
                categories.add(str(category))
                total_rows += int(count)
                total_amount += float(result_amounts.get(category, 0) or 0)
        self.category_stat_vars["groups"].set(str(len(categories)))
        self.category_stat_vars["rows"].set(str(total_rows))
        self.category_stat_vars["amount"].set(format_rupiah(total_amount))

    @staticmethod
    def invalid_files_for_results(results: list[dict[str, object]]) -> dict[str, list[str]]:
        invalid_files = {}
        for result in results:
            source_groups = result.get("source_groups") or result.get("categories") or {}
            if len(source_groups) > 1:
                invalid_files[str(result["filename"])] = sorted(source_groups)
        return invalid_files

    @staticmethod
    def result_matches_source(result: dict[str, object], source_path: str, filename: str) -> bool:
        if source_path:
            return str(result.get("_source_path") or "") == source_path
        return str(result.get("filename") or "") == filename

    @staticmethod
    def path_matches_source(path: Path, source_path: str, filename: str) -> bool:
        if source_path:
            return str(path) == source_path
        return path.name == filename

    def refresh_file_listbox(self) -> None:
        self.file_listbox.delete(0, "end")
        for path in self.file_paths:
            self.file_listbox.insert("end", path.name)

    @staticmethod
    def parse_sort_number(value: object) -> int:
        text = str(value or "")
        sign = -1 if text.strip().startswith("-") else 1
        digits = "".join(ch for ch in text if ch.isdigit())
        return sign * int(digits or "0")

    def preview_sort_value(self, item: str, column: str):
        values = self.tree.item(item, "values")
        try:
            column_index = self.preview_columns.index(column)
        except ValueError:
            return ""
        value = values[column_index] if column_index < len(values) else ""
        if column in {"no", "jumlah", "nominal"}:
            return self.parse_sort_number(value)
        return str(value or "").casefold()

    def update_preview_sort_headings(self) -> None:
        for column, label in self.preview_headings.items():
            suffix = ""
            if column == self.preview_sort_column:
                suffix = " ↓" if self.preview_sort_desc else " ↑"
            self.tree.heading(
                column,
                text=f"{label}{suffix}",
                command=lambda selected_column=column: self.sort_preview_column(selected_column),
            )

    def apply_preview_sort(self) -> None:
        if not self.preview_sort_column:
            return
        items = list(self.tree.get_children(""))
        items.sort(
            key=lambda item: self.preview_sort_value(item, self.preview_sort_column or ""),
            reverse=self.preview_sort_desc,
        )
        for index, item in enumerate(items):
            self.tree.move(item, "", index)
        self.update_preview_sort_headings()

    def sort_preview_column(self, column: str) -> None:
        if self.preview_sort_column == column:
            self.preview_sort_desc = not self.preview_sort_desc
        else:
            self.preview_sort_column = column
            self.preview_sort_desc = False
        self.apply_preview_sort()

    def show_duplicate_context_menu(self, event) -> None:
        item = self.tree.identify_row(event.y)
        if not item:
            return
        values = self.tree.item(item, "values")
        status = str(values[7]) if len(values) > 7 else ""
        if status != "Duplikat":
            return
        self.preview_context_item = item
        self.tree.selection_set(item)
        self.tree.focus(item)
        self.duplicate_menu.tk_popup(event.x_root, event.y_root)
        self.duplicate_menu.grab_release()

    def delete_context_duplicate_file(self) -> None:
        item = self.preview_context_item
        if not item:
            return
        values = self.tree.item(item, "values")
        if len(values) < 8 or str(values[7]) != "Duplikat":
            return

        filename = str(values[1])
        source_path = self.preview_item_sources.get(item, "")
        self.remove_file_from_checked_results(source_path, filename)

    def remove_file_from_checked_results(self, source_path: str, filename: str) -> None:
        before_count = len(self.file_paths)
        self.file_paths = [
            path for path in self.file_paths if not self.path_matches_source(path, source_path, filename)
        ]
        self.last_paths = [
            path for path in self.last_paths if not self.path_matches_source(path, source_path, filename)
        ]
        self.last_results = [
            result for result in self.last_results if not self.result_matches_source(result, source_path, filename)
        ]
        self.refresh_file_listbox()

        duplicates = find_duplicate_categories(self.last_results)
        invalid_files = self.invalid_files_for_results(self.last_results)
        self.last_duplicates = duplicates
        self.check_passed = bool(self.last_results) and not invalid_files
        self.populate_preview(self.last_results, duplicates, set(invalid_files))
        self.update_category_stats(self.last_results)
        self.set_busy_state(False)

        removed = before_count - len(self.file_paths)
        if not self.last_results:
            self.status_var.set("File terakhir dihapus dari olah data. Tambahkan file lagi untuk lanjut.")
        elif duplicates:
            self.status_var.set(
                f"{removed} file dihapus dari olah data. {len(duplicates)} group duplikat masih tersisa."
            )
        else:
            self.status_var.set(f"{removed} file dihapus dari olah data. Duplikat sudah bersih.")

    def delete_all_duplicate_files(self) -> None:
        targets = duplicate_category_files_to_remove(self.last_results)
        if not targets:
            self.last_duplicates = find_duplicate_categories(self.last_results)
            self.set_busy_state(False)
            self.status_var.set("Tidak ada file duplikat yang perlu dihapus.")
            return

        def is_target(path: Path) -> bool:
            return any(self.path_matches_source(path, source_path, filename) for source_path, filename in targets)

        def result_is_target(result: dict[str, object]) -> bool:
            return any(
                self.result_matches_source(result, source_path, filename)
                for source_path, filename in targets
            )

        before_count = len(self.file_paths)
        self.file_paths = [path for path in self.file_paths if not is_target(path)]
        self.last_paths = [path for path in self.last_paths if not is_target(path)]
        self.last_results = [result for result in self.last_results if not result_is_target(result)]
        self.refresh_file_listbox()

        duplicates = find_duplicate_categories(self.last_results)
        invalid_files = self.invalid_files_for_results(self.last_results)
        self.last_duplicates = duplicates
        self.check_passed = bool(self.last_results) and not invalid_files
        self.populate_preview(self.last_results, duplicates, set(invalid_files))
        self.update_category_stats(self.last_results)
        self.set_busy_state(False)

        removed = before_count - len(self.file_paths)
        if duplicates:
            self.status_var.set(
                f"{removed} file duplikat dihapus. {len(duplicates)} group duplikat masih tersisa."
            )
        else:
            self.status_var.set(
                f"{removed} file duplikat dihapus sekaligus. File pertama tiap group dipertahankan."
            )

    def add_files(self) -> None:
        paths = filedialog.askopenfilenames(
            title="Pilih file Excel kategori",
            filetypes=[("Excel files", "*.xlsx *.xlsm"), ("All files", "*.*")],
        )
        if not paths:
            return
        added = 0
        for raw_path in paths:
            path = Path(raw_path)
            if path in self.file_paths:
                continue
            self.file_paths.append(path)
            self.file_listbox.insert("end", path.name)
            added += 1
        if added:
            self.invalidate_check()
            self.update_category_stats()
            self.check_button.configure(state="normal")
            self.status_var.set(f"{len(self.file_paths)} file siap dicek. Tambah lagi atau klik Cek Group.")

    def remove_selected_file(self) -> None:
        selection = self.file_listbox.curselection()
        if not selection:
            return
        for index in reversed(selection):
            self.file_listbox.delete(index)
            del self.file_paths[index]
        self.invalidate_check()
        self.update_category_stats()
        if not self.file_paths:
            self.check_button.configure(state="disabled")
            self.status_var.set("Tambahkan file Excel kategori, lalu tekan Cek Group.")
        else:
            self.status_var.set(f"{len(self.file_paths)} file siap dicek. Tambah lagi atau klik Cek Group.")

    def invalidate_check(self) -> None:
        """Setiap kali daftar file berubah, hasil cek lama tidak berlaku lagi."""
        self.check_passed = False
        self.last_results = []
        self.last_duplicates = {}
        self.last_paths = []
        self.preview_item_sources.clear()
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.olah_button.configure(state="disabled")
        self.download_button.configure(state="disabled")
        self.delete_duplicates_button.configure(state="disabled")

    def check_group(self) -> None:
        if not self.file_paths:
            return
        self.cancel_event.clear()
        self.set_busy_state(True)
        paths = list(self.file_paths)
        self.status_var.set(f"Cek group file 1/{len(paths)}: {paths[0].name}...")
        thread = threading.Thread(target=self._worker, args=(paths,), daemon=True)
        thread.start()

    def cancel_processing(self) -> None:
        self.cancel_event.set()
        self.cancel_button.configure(state="disabled")
        self.status_var.set("Membatalkan proses...")

    def _worker(self, paths: list[Path]) -> None:
        try:
            results = []
            total = len(paths)
            for index, path in enumerate(paths, start=1):
                self.after(0, lambda i=index, name=path.name: self.status_var.set(f"Cek group file {i}/{total}: {name}..."))
                with path.open("rb") as file_obj:
                    result = extract_categories(file_obj, path.name, should_cancel=self.cancel_event.is_set)
                    result["_source_path"] = str(path)
                    results.append(result)
            duplicates = find_duplicate_categories(results)
            invalid_files = self.invalid_files_for_results(results)
            self.after(0, lambda: self.finish_check(paths, results, duplicates, invalid_files))
        except ProcessCancelled:
            self.after(0, self.finish_cancelled)
        except Exception as exc:
            message = str(exc)
            self.after(0, lambda message=message: self.finish_error(message))

    def finish_check(
        self,
        paths: list[Path],
        results: list[dict[str, object]],
        duplicates: dict[str, list[str]],
        invalid_files: dict[str, list[str]],
    ) -> None:
        self.last_paths = paths
        self.last_results = results
        self.last_duplicates = duplicates
        self.check_passed = not invalid_files
        self.populate_preview(results, duplicates, set(invalid_files))
        self.update_category_stats(results)

        if invalid_files:
            lines = [f"- {name}: {', '.join(groups)}" for name, groups in invalid_files.items()]
            messagebox.showerror(
                "Group ganda ditemukan",
                "File berikut punya lebih dari satu group, jadi tidak bisa diolah.\n"
                "Pisahkan jadi satu group per file dulu:\n\n" + "\n".join(lines),
            )
            self.status_var.set(f"Gagal: {len(invalid_files)} file punya group ganda. Perbaiki dulu sebelum Olah Data.")
        elif duplicates:
            lines = [f"- {category}: {', '.join(files)}" for category, files in duplicates.items()]
            messagebox.showwarning(
                "Kategori duplikat antar file",
                "Group berikut muncul di lebih dari satu file (boleh lanjut, tapi cek lagi):\n\n" + "\n".join(lines),
            )
            self.status_var.set(f"Cek OK. {len(duplicates)} group duplikat antar file. Siap Olah Data.")
        else:
            self.status_var.set("Cek OK. Semua file 1 group. Siap Olah Data.")
        self.set_busy_state(False)

    def download_result(self) -> None:
        if not self.check_passed or not self.last_results:
            messagebox.showinfo("Cek dulu", "Jalankan Cek Group dan pastikan semua file 1 group sebelum download.")
            return

        downloads_dir = Path.home() / "Downloads"
        initial_dir = str(downloads_dir if downloads_dir.exists() else Path.home())
        path = filedialog.asksaveasfilename(
            title="Simpan Omset Keseluruhan Per Kategori",
            defaultextension=".xlsx",
            initialfile="Omset Keseluruhan PERKATEGORI.xlsx",
            initialdir=initial_dir,
            filetypes=[("Excel files", "*.xlsx")],
        )
        if not path:
            return

        try:
            output_path = build_category_workbook(self.last_results, self.last_duplicates, Path(path))
        except Exception as exc:
            messagebox.showerror("Gagal menyimpan", str(exc))
            return

        self.status_var.set(f"File tersimpan: {output_path.name}")
        if messagebox.askyesno("Download selesai", f"File tersimpan di:\n{output_path}\n\nBuka file sekarang?"):
            os.startfile(output_path)

    def olah_data(self) -> None:
        if not self.check_passed or not self.last_paths:
            messagebox.showinfo(
                "Cek dulu",
                "Jalankan Cek Group dan pastikan semua file 1 group sebelum Olah Data.",
            )
            return

        downloads_dir = Path.home() / "Downloads"
        initial_dir = str(downloads_dir if downloads_dir.exists() else Path.home())
        path = filedialog.asksaveasfilename(
            title="Simpan workbook olahan",
            defaultextension=".xlsx",
            initialfile="Omset Keseluruhan PERKATEGORI.xlsx",
            initialdir=initial_dir,
            filetypes=[("Excel files", "*.xlsx")],
        )
        if not path:
            return

        self.cancel_event.clear()
        self.set_busy_state(True)
        self.status_var.set("Mengolah semua data jadi 1 workbook...")
        thread = threading.Thread(
            target=self._olah_worker,
            args=(list(self.last_paths), Path(path)),
            daemon=True,
        )
        thread.start()

    def _olah_worker(self, paths: list[Path], output_path: Path) -> None:
        try:
            details = []
            total = len(paths)
            for index, path in enumerate(paths, start=1):
                self.after(0, lambda i=index, name=path.name: self.status_var.set(f"Olah data {i}/{total}: {name}..."))
                with path.open("rb") as file_obj:
                    details.append(read_category_detail(file_obj, path.name, should_cancel=self.cancel_event.is_set))
            output = build_olah_data_workbook(details, output_path)
            self.after(0, lambda: self.finish_olah(output))
        except ProcessCancelled:
            self.after(0, self.finish_cancelled)
        except Exception as exc:
            message = str(exc)
            self.after(0, lambda message=message: self.finish_error(message))

    def finish_olah(self, output_path: Path) -> None:
        self.set_busy_state(False)
        self.status_var.set(f"Workbook dibuat: {output_path.name}")
        if messagebox.askyesno("Olah data selesai", f"Workbook tersimpan di:\n{output_path}\n\nBuka file sekarang?"):
            os.startfile(output_path)

    def finish_error(self, message: str) -> None:
        self.status_var.set("Gagal membaca file.")
        self.set_busy_state(False)
        messagebox.showerror("Gagal membaca kategori", message)

    def finish_cancelled(self) -> None:
        self.status_var.set("Proses dibatalkan.")
        self.set_busy_state(False)

    def populate_preview(
        self,
        results: list[dict[str, object]],
        duplicates: dict[str, list[str]],
        invalid_filenames: set[str] | None = None,
    ) -> None:
        invalid_filenames = invalid_filenames or set()
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.preview_item_sources.clear()
        duplicate_keys = {category.upper() for category in duplicates}
        no = 1
        for result in results:
            first_date = str(result.get("first_date") or "")
            last_date = str(result.get("last_date") or "")
            if first_date and last_date and first_date != last_date:
                date_label = f"{first_date} s/d {last_date}"
            else:
                date_label = first_date or last_date or "-"
            amounts = result.get("amounts") if isinstance(result.get("amounts"), dict) else {}
            file_invalid = str(result["filename"]) in invalid_filenames
            for category, count in sorted(result["categories"].items()):
                if file_invalid:
                    status, tag = "Group Ganda", "invalid"
                elif category.upper() in duplicate_keys:
                    status, tag = "Duplikat", "dup"
                else:
                    status, tag = "OK", ""
                amount = amounts.get(category, 0) if isinstance(amounts, dict) else 0
                item = self.tree.insert(
                    "",
                    "end",
                    values=(
                        no,
                        result["filename"],
                        result.get("group_column") or "-",
                        category,
                        date_label,
                        count,
                        format_rupiah(amount),
                        status,
                    ),
                    tags=(tag,) if tag else (),
                )
                self.preview_item_sources[item] = str(result.get("_source_path") or "")
                no += 1
        self.apply_preview_sort()

    def set_busy_state(self, is_busy: bool) -> None:
        if is_busy:
            self.add_button.configure(state="disabled")
            self.remove_button.configure(state="disabled")
            self.check_button.configure(state="disabled")
            self.olah_button.configure(state="disabled")
            self.download_button.configure(state="disabled")
            self.delete_duplicates_button.configure(state="disabled")
            self.cancel_button.configure(state="normal")
            self.progress_bar.start(12)
        else:
            self.progress_bar.stop()
            self.progress_bar.configure(value=0)
            self.add_button.configure(state="normal")
            self.remove_button.configure(state="normal")
            self.cancel_button.configure(state="disabled")
            self.check_button.configure(state="normal" if self.file_paths else "disabled")
            ready = "normal" if self.check_passed else "disabled"
            self.olah_button.configure(state=ready)
            self.download_button.configure(state=ready)
            self.delete_duplicates_button.configure(state="normal" if self.last_duplicates else "disabled")


class PdfToExcelDialog(tk.Toplevel):
    def __init__(self, master: tk.Misc | None = None) -> None:
        super().__init__(master)
        self.title("PDF TO EXCEL")
        self.protocol("WM_DELETE_WINDOW", self.back_to_dashboard)
        self.geometry(DEFAULT_APP_GEOMETRY)
        self.minsize(*DEFAULT_APP_MINSIZE)
        self.configure(bg=COL_CONTENT)
        configure_styles(self)

        self.file_paths: list[Path] = []
        self.output_path: Path | None = None
        self.cancel_event = threading.Event()
        self.status_var = tk.StringVar(value="Pilih satu atau beberapa file PDF untuk mulai.")
        self.metric_vars = {
            "files": tk.StringVar(value="0"),
            "pages": tk.StringVar(value="-"),
            "tables": tk.StringVar(value="-"),
            "rows": tk.StringVar(value="-"),
        }
        self._build_ui()
        apply_button_cursors(self)

    def back_to_dashboard(self) -> None:
        if self.master is not None:
            match_window_size(self.master, self)
        self.withdraw()
        if self.master is not None:
            self.master.deiconify()
            self.master.lift()
            self.master.focus_force()

    def _build_ui(self) -> None:
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)

        header = ttk.Frame(self, style="Header.TFrame", padding=(22, 18))
        header.grid(row=0, column=0, sticky="ew")
        header.columnconfigure(0, weight=1)
        ttk.Label(header, text="PDF TO EXCEL", style="Title.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(
            header,
            text="Ekstrak tabel dan teks dari beberapa PDF menjadi satu workbook Excel yang rapi.",
            style="Subtitle.TLabel",
        ).grid(row=1, column=0, sticky="w", pady=(4, 0))
        ttk.Button(header, text="←  Kembali ke Dashboard", style="Ghost.TButton", command=self.back_to_dashboard).grid(
            row=0, column=1, rowspan=2, sticky="e"
        )

        main = ttk.Frame(self, style="Content.TFrame", padding=20)
        main.grid(row=1, column=0, sticky="nsew")
        main.columnconfigure(0, weight=3)
        main.columnconfigure(1, weight=2)
        main.rowconfigure(0, weight=1)

        files_card = ttk.Frame(main, style="Card.TFrame", padding=18)
        files_card.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        files_card.columnconfigure(0, weight=1)
        files_card.rowconfigure(2, weight=1)
        ttk.Label(files_card, text="File PDF", style="CardTitle.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(
            files_card,
            text="Urutan file di bawah akan dipakai pada workbook hasil.",
            style="CardSub.TLabel",
        ).grid(row=1, column=0, sticky="w", pady=(3, 12))

        columns = ("name", "folder", "size")
        self.file_tree = ttk.Treeview(files_card, columns=columns, show="headings", selectmode="extended")
        self.file_tree.heading("name", text="Nama File")
        self.file_tree.heading("folder", text="Lokasi")
        self.file_tree.heading("size", text="Ukuran")
        self.file_tree.column("name", width=260, minwidth=160, anchor="w")
        self.file_tree.column("folder", width=270, minwidth=160, anchor="w")
        self.file_tree.column("size", width=80, minwidth=70, anchor="e")
        self.file_tree.grid(row=2, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(files_card, orient="vertical", command=self.file_tree.yview)
        scrollbar.grid(row=2, column=1, sticky="ns")
        self.file_tree.configure(yscrollcommand=scrollbar.set)

        file_actions = ttk.Frame(files_card, style="Card.TFrame")
        file_actions.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(14, 0))
        file_actions.columnconfigure((0, 1, 2), weight=1)
        self.add_button = ttk.Button(
            file_actions, text="+  Pilih PDF", style="Blue.TButton", command=self.choose_files
        )
        self.add_button.grid(row=0, column=0, sticky="ew", padx=(0, 8))
        self.remove_button = ttk.Button(
            file_actions, text="Hapus Pilihan", style="Secondary.TButton", command=self.remove_selected, state="disabled"
        )
        self.remove_button.grid(row=0, column=1, sticky="ew", padx=(0, 8))
        self.clear_button = ttk.Button(
            file_actions, text="Kosongkan", style="Secondary.TButton", command=self.clear_files, state="disabled"
        )
        self.clear_button.grid(row=0, column=2, sticky="ew")

        side = ttk.Frame(main, style="Card.TFrame", padding=18)
        side.grid(row=0, column=1, sticky="nsew", padx=(10, 0))
        side.columnconfigure((0, 1), weight=1)
        side.rowconfigure(5, weight=1)
        ttk.Label(side, text="Hasil Konversi", style="CardTitle.TLabel").grid(
            row=0, column=0, columnspan=2, sticky="w"
        )
        ttk.Label(
            side,
            text="Setiap halaman dibuat sebagai sheet terpisah. Sheet Ringkasan memuat indeks seluruh hasil.",
            style="CardSub.TLabel",
            wraplength=330,
            justify="left",
        ).grid(row=1, column=0, columnspan=2, sticky="w", pady=(4, 14))

        metric_items = [("FILE", "files"), ("HALAMAN", "pages"), ("TABEL", "tables"), ("BARIS", "rows")]
        for index, (label, key) in enumerate(metric_items):
            box = tk.Frame(side, bg=COL_SOFT, highlightthickness=1, highlightbackground=COL_BORDER)
            box.grid(
                row=2 + index // 2,
                column=index % 2,
                sticky="nsew",
                padx=(0 if index % 2 == 0 else 7, 7 if index % 2 == 0 else 0),
                pady=(0, 7),
            )
            tk.Label(box, text=label, bg=COL_SOFT, fg=COL_MUTED, font=(FONT, 8, "bold")).pack(
                anchor="w", padx=12, pady=(10, 2)
            )
            tk.Label(box, textvariable=self.metric_vars[key], bg=COL_SOFT, fg=COL_TEXT, font=(FONT_SB, 17)).pack(
                anchor="w", padx=12, pady=(0, 10)
            )

        note = tk.Frame(side, bg="#eef8f6", highlightthickness=1, highlightbackground="#c8e5df")
        note.grid(row=4, column=0, columnspan=2, sticky="ew", pady=(8, 14))
        tk.Label(note, text="CATATAN", bg="#eef8f6", fg=COL_BRAND, font=(FONT, 8, "bold")).pack(
            anchor="w", padx=12, pady=(10, 3)
        )
        tk.Label(
            note,
            text="PDF scan berupa gambar memerlukan OCR terlebih dahulu.",
            bg="#eef8f6",
            fg=COL_MUTED,
            font=(FONT, 9),
            justify="left",
            wraplength=300,
        ).pack(anchor="w", padx=12, pady=(0, 10))

        action_area = ttk.Frame(side, style="Card.TFrame")
        action_area.grid(row=5, column=0, columnspan=2, sticky="sew")
        action_area.columnconfigure(0, weight=1)
        self.process_button = ttk.Button(
            action_area, text="Buat File Excel", style="PrimaryLarge.TButton", command=self.start_conversion, state="disabled"
        )
        self.process_button.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        self.cancel_button = ttk.Button(
            action_area, text="Batalkan", style="Danger.TButton", command=self.cancel_processing, state="disabled"
        )
        self.cancel_button.grid(row=1, column=0, sticky="ew", pady=(0, 8))
        self.open_button = ttk.Button(
            action_area, text="Buka File Excel", style="Secondary.TButton", command=self.open_output, state="disabled"
        )
        self.open_button.grid(row=2, column=0, sticky="ew", pady=(0, 8))
        ttk.Button(action_area, text="Buka Folder Output", style="Secondary.TButton", command=self.open_output_folder).grid(
            row=3, column=0, sticky="ew"
        )
        ttk.Label(action_area, textvariable=self.status_var, style="Hint.TLabel", wraplength=330, justify="left").grid(
            row=4, column=0, sticky="w", pady=(14, 7)
        )
        self.progress_bar = ttk.Progressbar(action_area, mode="indeterminate")
        self.progress_bar.grid(row=5, column=0, sticky="ew")

    def choose_files(self) -> None:
        selected = filedialog.askopenfilenames(
            title="Pilih file PDF",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
        )
        if not selected:
            return
        known = {str(path.resolve()).casefold() for path in self.file_paths}
        for raw_path in selected:
            path = Path(raw_path)
            key = str(path.resolve()).casefold()
            if key not in known:
                self.file_paths.append(path)
                known.add(key)
        self.output_path = None
        self.refresh_file_list()
        self.status_var.set(f"{len(self.file_paths)} file PDF siap dikonversi.")

    def refresh_file_list(self) -> None:
        for item in self.file_tree.get_children():
            self.file_tree.delete(item)
        for index, path in enumerate(self.file_paths):
            size_kb = max(1, round(path.stat().st_size / 1024)) if path.exists() else 0
            self.file_tree.insert("", "end", iid=str(index), values=(path.name, str(path.parent), f"{size_kb:,} KB"))
        count = len(self.file_paths)
        self.metric_vars["files"].set(str(count))
        self.process_button.configure(state="normal" if count else "disabled")
        self.remove_button.configure(state="normal" if count else "disabled")
        self.clear_button.configure(state="normal" if count else "disabled")
        self.open_button.configure(state="disabled")

    def remove_selected(self) -> None:
        indexes = sorted((int(item) for item in self.file_tree.selection()), reverse=True)
        if not indexes:
            messagebox.showinfo("Pilih file", "Pilih file yang ingin dihapus dari daftar.")
            return
        for index in indexes:
            self.file_paths.pop(index)
        self.output_path = None
        self.refresh_file_list()
        self.status_var.set(f"{len(self.file_paths)} file PDF tersisa.")

    def clear_files(self) -> None:
        self.file_paths.clear()
        self.output_path = None
        for key in ("pages", "tables", "rows"):
            self.metric_vars[key].set("-")
        self.refresh_file_list()
        self.status_var.set("Daftar dikosongkan. Pilih file PDF untuk mulai.")

    def start_conversion(self) -> None:
        if not self.file_paths:
            return
        self.cancel_event.clear()
        self.set_busy_state(True)
        self.status_var.set("Menyiapkan konversi PDF...")
        threading.Thread(target=self._conversion_worker, args=(list(self.file_paths),), daemon=True).start()

    def _conversion_worker(self, paths: list[Path]) -> None:
        def progress(file_index: int, file_total: int, name: str, page: int, page_total: int) -> None:
            message = f"File {file_index}/{file_total}: {name} — halaman {page}/{page_total}"
            self.after(0, lambda message=message: self.status_var.set(message))

        try:
            output_path, stats = convert_pdfs_to_excel(
                paths,
                should_cancel=self.cancel_event.is_set,
                progress=progress,
            )
            self.output_path = output_path
            self.after(0, lambda: self.finish_success(stats))
        except ProcessCancelled:
            self.after(0, self.finish_cancelled)
        except Exception as exc:
            message = str(exc)
            self.after(0, lambda message=message: self.finish_error(message))

    def set_busy_state(self, is_busy: bool) -> None:
        state = "disabled" if is_busy else "normal"
        self.add_button.configure(state=state)
        self.remove_button.configure(state=state if self.file_paths else "disabled")
        self.clear_button.configure(state=state if self.file_paths else "disabled")
        self.process_button.configure(state=state if self.file_paths else "disabled")
        self.cancel_button.configure(state="normal" if is_busy else "disabled")
        if is_busy:
            self.open_button.configure(state="disabled")
            self.progress_bar.start(12)
        else:
            self.progress_bar.stop()
            self.progress_bar.configure(value=0)
            self.open_button.configure(state="normal" if self.output_path and self.output_path.exists() else "disabled")

    def cancel_processing(self) -> None:
        self.cancel_event.set()
        self.cancel_button.configure(state="disabled")
        self.status_var.set("Membatalkan proses...")

    def finish_success(self, stats: dict[str, object]) -> None:
        self.metric_vars["pages"].set(str(stats["page_count"]))
        self.metric_vars["tables"].set(str(stats["table_count"]))
        self.metric_vars["rows"].set(str(stats["row_count"]))
        self.status_var.set(f"Selesai. Workbook dibuat: {stats['filename']}")
        self.set_busy_state(False)

    def finish_error(self, message: str) -> None:
        self.status_var.set(message)
        self.set_busy_state(False)
        messagebox.showerror("Konversi gagal", message)

    def finish_cancelled(self) -> None:
        self.status_var.set("Konversi dibatalkan.")
        self.set_busy_state(False)

    def open_output(self) -> None:
        if self.output_path and self.output_path.exists():
            os.startfile(self.output_path)

    def open_output_folder(self) -> None:
        OUTPUT_DIR.mkdir(exist_ok=True)
        os.startfile(OUTPUT_DIR)


class DashboardApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("BC Padel Club - Dashboard")
        self.geometry(DEFAULT_APP_GEOMETRY)
        self.minsize(*DEFAULT_APP_MINSIZE)
        self.configure(bg=COL_CONTENT)
        configure_styles(self)
        try:
            self.state("zoomed")
        except tk.TclError:
            pass

        self.lapangan_window: RekapExcelApp | None = None
        self.kategori_window: CategoryUploadDialog | None = None
        self.pdf_window: PdfToExcelDialog | None = None

        self._build_ui()
        apply_button_cursors(self)
        self.refresh_dashboard()

    def _build_ui(self) -> None:
        self.columnconfigure(1, weight=1)
        self.rowconfigure(0, weight=1)

        sidebar = ttk.Frame(self, style="Sidebar.TFrame", padding=(18, 22))
        sidebar.grid(row=0, column=0, sticky="ns")
        sidebar.grid_propagate(False)
        sidebar.configure(width=230)
        sidebar.rowconfigure(8, weight=1)

        brand = tk.Frame(sidebar, bg=COL_SIDEBAR)
        brand.grid(row=0, column=0, sticky="ew", pady=(0, 26))
        logo = tk.Label(
            brand,
            text="BC",
            bg=COL_BLUE,
            fg="#ffffff",
            font=(FONT_SB, 12),
            width=4,
            height=2,
        )
        logo.pack(side="left")
        brand_text = tk.Frame(brand, bg=COL_SIDEBAR)
        brand_text.pack(side="left", padx=(10, 0))
        ttk.Label(brand_text, text="BC Padel Club", style="SidebarTitle.TLabel").pack(anchor="w")
        ttk.Label(brand_text, text="Rekap operasional", style="SidebarSub.TLabel").pack(anchor="w", pady=(2, 0))

        self._nav_item(sidebar, 1, "⌂   Dashboard", True, self.refresh_dashboard)
        self._nav_item(sidebar, 2, "▦   Omset Lapangan", False, self.open_lapangan)
        self._nav_item(sidebar, 3, "▤   Perkategori Olsera", False, self.open_kategori)
        self._nav_item(sidebar, 4, "▣   PDF TO EXCEL", False, self.open_pdf_to_excel)

        footer = tk.Frame(sidebar, bg=COL_SIDEBAR)
        footer.grid(row=9, column=0, sticky="ew")
        tk.Label(
            footer,
            text="Output tersimpan lokal\ndi folder aplikasi.",
            bg=COL_SIDEBAR,
            fg="#9ca3af",
            font=(FONT, 9),
            justify="left",
        ).pack(anchor="w")

        content = ttk.Frame(self, style="Content.TFrame", padding=(26, 22))
        content.grid(row=0, column=1, sticky="nsew")
        content.columnconfigure(0, weight=1)
        content.rowconfigure(3, weight=1)

        top = ttk.Frame(content, style="Content.TFrame")
        top.grid(row=0, column=0, sticky="ew", pady=(0, 18))
        top.columnconfigure(0, weight=1)
        ttk.Label(top, text="Dashboard Operasional", style="PageTitle.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(
            top,
            text="Pilih alur kerja yang dibutuhkan dan kelola seluruh file hasil dari satu layar.",
            style="PageSub.TLabel",
        ).grid(row=1, column=0, sticky="w", pady=(3, 0))
        ttk.Button(top, text="↻  Muat Ulang", style="Secondary.TButton", command=self.refresh_dashboard).grid(
            row=0, column=1, rowspan=2, sticky="e"
        )

        hero = tk.Frame(content, bg=COL_HEADER, bd=0)
        hero.grid(row=1, column=0, sticky="ew", pady=(0, 18))
        hero.columnconfigure(0, weight=1)
        hero_text = tk.Frame(hero, bg=COL_HEADER)
        hero_text.grid(row=0, column=0, sticky="w", padx=22, pady=18)
        tk.Label(
            hero_text,
            text="Data operasional, lebih cepat dan terstruktur",
            bg=COL_HEADER,
            fg="#ffffff",
            font=(FONT_SB, 15),
        ).pack(anchor="w")
        tk.Label(
            hero_text,
            text="Pilih modul di bawah — hasil otomatis tersimpan di folder output",
            bg=COL_HEADER,
            fg=COL_HEADER_SUB,
            font=(FONT, 9, "bold"),
        ).pack(anchor="w", pady=(5, 0))
        tk.Label(
            hero,
            text="SIAP DIGUNAKAN",
            bg=COL_BRAND,
            fg="#ffffff",
            font=(FONT_SB, 10),
            padx=16,
            pady=9,
        ).grid(row=0, column=1, sticky="e", padx=22)

        modules = ttk.Frame(content, style="Content.TFrame")
        modules.grid(row=2, column=0, sticky="ew", pady=(0, 18))
        modules.columnconfigure(0, weight=1, uniform="modules")
        modules.columnconfigure(1, weight=1, uniform="modules")
        modules.columnconfigure(2, weight=1, uniform="modules")
        self._module_card(
            modules,
            0,
            "REKAP BOOKING",
            "Omset Lapangan",
            "Pisahkan AYO dan Walk In, deteksi periode, lalu buat laporan lengkap.",
            "Mulai Rekap  →",
            self.open_lapangan,
            COL_BLUE,
            "Blue.TButton",
        )
        self._module_card(
            modules,
            1,
            "DATA OLSERA",
            "Omset Perkategori Olsera",
            "Periksa item group, duplikat, dan gabungkan laporan per kategori.",
            "Kelola Kategori  →",
            self.open_kategori,
            COL_BRAND,
            "Primary.TButton",
        )
        self._module_card(
            modules,
            2,
            "KONVERSI DOKUMEN",
            "PDF TO EXCEL",
            "Ubah tabel dan teks dari beberapa PDF menjadi workbook terstruktur.",
            "Konversi PDF  →",
            self.open_pdf_to_excel,
            COL_AMBER,
            "Gold.TButton",
        )

        output_card = ttk.Frame(content, style="Card.TFrame", padding=18)
        output_card.grid(row=3, column=0, sticky="nsew")
        output_card.columnconfigure(0, weight=1)
        output_card.rowconfigure(1, weight=1)
        header = ttk.Frame(output_card, style="Card.TFrame")
        header.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        header.columnconfigure(0, weight=1)
        ttk.Label(header, text="Output Terbaru", style="CardTitle.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Button(header, text="Buka Folder Output", style="Secondary.TButton", command=self.open_output_folder).grid(
            row=0, column=1, sticky="e"
        )

        columns = ("name", "modified", "size")
        self.output_tree = ttk.Treeview(output_card, columns=columns, show="headings", selectmode="browse")
        self.output_tree.heading("name", text="File")
        self.output_tree.heading("modified", text="Terakhir Diubah")
        self.output_tree.heading("size", text="Ukuran")
        self.output_tree.column("name", width=520, minwidth=260, anchor="w")
        self.output_tree.column("modified", width=160, minwidth=140, anchor="w")
        self.output_tree.column("size", width=100, minwidth=90, anchor="e")
        self.output_tree.grid(row=1, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(output_card, orient="vertical", command=self.output_tree.yview)
        scrollbar.grid(row=1, column=1, sticky="ns")
        self.output_tree.configure(yscrollcommand=scrollbar.set)
        self.output_tree.bind("<Double-1>", self.open_selected_output)

    @staticmethod
    def _within(widget, x, y) -> bool:
        try:
            wx, wy = widget.winfo_rootx(), widget.winfo_rooty()
            return wx <= x <= wx + widget.winfo_width() and wy <= y <= wy + widget.winfo_height()
        except tk.TclError:
            return False

    def _nav_item(self, parent, row: int, text: str, active: bool, command) -> None:
        bg = COL_BLUE if active else COL_SIDEBAR
        hover = COL_BLUE_DARK if active else COL_SIDEBAR_SOFT
        frame = tk.Frame(parent, bg=bg, bd=0)
        frame.grid(row=row, column=0, sticky="ew", pady=3)
        label = tk.Label(
            frame,
            text=text,
            bg=bg,
            fg="#ffffff" if active else "#d1d5db",
            font=(FONT_SB if active else FONT, 10),
            anchor="w",
            padx=14,
            pady=10,
        )
        label.pack(fill="x")

        def on_enter(_=None):
            frame.configure(bg=hover)
            label.configure(bg=hover, fg="#ffffff")

        def on_leave(_=None):
            frame.configure(bg=bg)
            label.configure(bg=bg, fg="#ffffff" if active else "#d1d5db")

        for widget in (frame, label):
            widget.bind("<Enter>", on_enter, add="+")
            widget.bind("<Leave>", on_leave, add="+")
            widget.bind("<Button-1>", lambda _e: command(), add="+")
            widget.configure(cursor="hand2")

    def _stat_card(self, parent, column: int, title: str, variable: tk.StringVar, caption: str, accent: str) -> None:
        card = tk.Frame(parent, bg=COL_SURFACE, highlightthickness=1, highlightbackground=COL_BORDER)
        card.grid(row=0, column=column, sticky="nsew", padx=(0 if column == 0 else 10, 0))
        card.columnconfigure(1, weight=1)
        tk.Frame(card, bg=accent, width=5).grid(row=0, column=0, rowspan=3, sticky="ns")
        tk.Label(card, text=title.upper(), bg=COL_SURFACE, fg=COL_MUTED, font=(FONT, 8, "bold")).grid(
            row=0, column=1, sticky="w", padx=14, pady=(12, 2)
        )
        tk.Label(
            card,
            textvariable=variable,
            bg=COL_SURFACE,
            fg=COL_TEXT,
            font=(FONT_SB, 16),
            anchor="w",
            justify="left",
            wraplength=190,
        ).grid(
            row=1, column=1, sticky="ew", padx=14
        )
        tk.Label(card, text=caption, bg=COL_SURFACE, fg=COL_MUTED, font=(FONT, 9)).grid(
            row=2, column=1, sticky="w", padx=14, pady=(2, 12)
        )

    def _module_card(
        self,
        parent,
        column: int,
        badge: str,
        title: str,
        desc: str,
        button_text: str,
        command,
        accent: str,
        button_style: str,
    ) -> None:
        outer = tk.Frame(parent, bg=COL_BORDER, bd=0)
        outer.grid(row=0, column=column, sticky="nsew", padx=(0 if column == 0 else 10, 0))
        outer.columnconfigure(0, weight=1)
        inner = tk.Frame(outer, bg=COL_SURFACE)
        inner.grid(row=0, column=0, sticky="nsew", padx=1, pady=1)
        inner.columnconfigure(0, weight=1)
        accent_line = tk.Frame(inner, bg=accent, height=4)
        accent_line.grid(row=0, column=0, sticky="ew")
        badge_label = tk.Label(
            inner,
            text=badge,
            bg="#f8fafc",
            fg=accent,
            font=(FONT_SB, 9),
            padx=8,
            pady=4,
        )
        badge_label.grid(row=1, column=0, sticky="w", padx=16, pady=(14, 8))
        title_label = tk.Label(
            inner, text=title, bg=COL_SURFACE, fg=COL_TEXT, font=(FONT_SB, 13), anchor="w"
        )
        title_label.grid(row=2, column=0, sticky="ew", padx=16, pady=(0, 6))
        desc_label = tk.Label(
            inner,
            text=desc,
            bg=COL_SURFACE,
            fg=COL_MUTED,
            font=(FONT, 10),
            anchor="nw",
            justify="left",
            wraplength=230,
        )
        desc_label.grid(row=3, column=0, sticky="ew", padx=16)
        action_button = ttk.Button(inner, text=button_text, style=button_style, command=command)
        action_button.configure(cursor="hand2")
        action_button.grid(
            row=4, column=0, sticky="ew", padx=16, pady=(16, 16)
        )

        def on_enter(_event=None):
            outer.configure(bg=accent)
            title_label.configure(fg=accent)

        def on_leave(_event=None):
            outer.configure(bg=COL_BORDER)
            title_label.configure(fg=COL_TEXT)

        for widget in (outer, inner, accent_line, badge_label, title_label, desc_label):
            widget.bind("<Enter>", on_enter, add="+")
            widget.bind("<Leave>", on_leave, add="+")
            widget.bind("<Button-1>", lambda _event: command(), add="+")
            widget.configure(cursor="hand2")

    def refresh_dashboard(self) -> None:
        OUTPUT_DIR.mkdir(exist_ok=True)
        output_files = sorted(OUTPUT_DIR.glob("*.xlsx"), key=lambda p: p.stat().st_mtime, reverse=True)

        if hasattr(self, "output_tree"):
            for item in self.output_tree.get_children():
                self.output_tree.delete(item)
            for path in output_files[:20]:
                stat = path.stat()
                size_kb = max(1, round(stat.st_size / 1024))
                modified = datetime.fromtimestamp(stat.st_mtime).strftime("%d %b %Y, %H:%M")
                self.output_tree.insert("", "end", values=(path.name, modified, f"{size_kb:,} KB"))

    def open_output_folder(self) -> None:
        OUTPUT_DIR.mkdir(exist_ok=True)
        os.startfile(OUTPUT_DIR)

    def open_selected_output(self, _event=None) -> None:
        selected = self.output_tree.selection()
        if not selected:
            return
        filename = self.output_tree.item(selected[0], "values")[0]
        path = OUTPUT_DIR / Path(filename).name
        if path.exists():
            os.startfile(path)

    def _open_module(self, attr: str, factory) -> None:
        window = getattr(self, attr)
        if window is None or not window.winfo_exists():
            window = factory()
            setattr(self, attr, window)
        match_window_size(window, self)
        window.deiconify()
        window.lift()
        window.focus_force()
        self.withdraw()  # sembunyikan dashboard supaya tidak menutupi modul

    def open_lapangan(self) -> None:
        self._open_module("lapangan_window", lambda: RekapExcelApp(self))

    def open_kategori(self) -> None:
        self._open_module("kategori_window", lambda: CategoryUploadDialog(self))

    def open_pdf_to_excel(self) -> None:
        self._open_module("pdf_window", lambda: PdfToExcelDialog(self))


if __name__ == "__main__":
    DashboardApp().mainloop()
