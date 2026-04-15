import os
import re
import sys
import zipfile
import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
from pathlib import Path
from typing import List, Dict, Tuple

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter
from playwright.sync_api import sync_playwright

TARGET_COLUMNS = [
    "Identyfikator sprzedawcy",
    "Nazwa sprzedawcy",
    "Nr KSeF",
    "Nr faktury",
    "Data wystawienia",
    "Data zapisania w KSeF",
    "Data otrzymania",
    "Waluta",
    "Netto",
    "Brutto",
    "VAT (PLN)",
]

HEADER_ALIASES = {
    "Identyfikator sprzedawcy": ["identyfikator sprzedawcy", "nip sprzedawcy", "nip", "sprzedawca nip"],
    "Nazwa sprzedawcy": ["nazwa sprzedawcy", "sprzedawca", "nazwa podmiotu", "nazwa"],
    "Nr KSeF": ["nr ksef", "numer ksef", "ksef", "numer referencyjny ksef"],
    "Nr faktury": ["nr faktury", "numer faktury", "faktura"],
    "Data wystawienia": ["data wystawienia", "wystawiono"],
    "Data zapisania w KSeF": ["data zapisania w ksef", "data przyjęcia w ksef", "zapisano w ksef"],
    "Data otrzymania": ["data otrzymania", "otrzymano"],
    "Waluta": ["waluta"],
    "Netto": ["netto", "wartość netto", "kwota netto"],
    "Brutto": ["brutto", "wartość brutto", "kwota brutto"],
    "VAT (PLN)": ["vat (pln)", "vat pln", "vat", "kwota vat pln"],
}


class KsefSimpleSummaryApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("KSeF – Zestawienie FV")
        self.root.geometry("980x700")
        self.root.minsize(920, 660)
        self.root.configure(bg="#f5f7fb")

        self.playwright = None
        self.browser = None
        self.context = None
        self.page = None

        self.base_dir = os.getcwd()
        self.output_dir = os.path.join(self.base_dir, "zestawienia_ksef")
        os.makedirs(self.output_dir, exist_ok=True)

        today = datetime.today()
        month_start = today.replace(day=1)

        self.date_from_var = tk.StringVar(value=month_start.strftime("%Y-%m-%d"))
        self.date_to_var = tk.StringVar(value=today.strftime("%Y-%m-%d"))
        self.status_var = tk.StringVar(value="Gotowe")
        self.summary_var = tk.StringVar(value="Brak zapisanych danych")
        self.file_var = tk.StringVar(value="-")

        self.last_headers: List[str] = []
        self.last_rows: List[Dict] = []
        self.last_pages: int = 0

        self.setup_style()
        self.build_ui()
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def setup_style(self):
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Main.TButton", font=("Segoe UI", 11, "bold"), padding=10)
        style.configure("Ghost.TButton", font=("Segoe UI", 10), padding=8)
        style.configure(
            "Red.Horizontal.TProgressbar",
            troughcolor="#e5e7eb",
            background="#c81f25",
            bordercolor="#e5e7eb",
            lightcolor="#c81f25",
            darkcolor="#c81f25",
        )

    def build_ui(self):
        outer = tk.Frame(self.root, bg="#f5f7fb", padx=18, pady=18)
        outer.pack(fill="both", expand=True)

        header = tk.Frame(outer, bg="#ffffff", bd=1, relief="solid")
        header.pack(fill="x", pady=(0, 12))

        tk.Label(
            header,
            text="KSeF – Zestawienie faktur do Excel",
            font=("Segoe UI", 20, "bold"),
            bg="#ffffff",
            fg="#111827",
            padx=18,
            pady=14,
        ).pack(anchor="w")

        tk.Label(
            header,
            text="1. Otwórz KSeF  2. Zaloguj się  3. Kliknij \"Rozwiń filtry\" i ustaw daty ręcznie  4. Kliknij pobierz zestawienie",
            font=("Segoe UI", 10),
            bg="#ffffff",
            fg="#475569",
            padx=18,
            pady=0,
        ).pack(anchor="w", pady=(0, 14))

        top = tk.Frame(outer, bg="#f5f7fb")
        top.pack(fill="x", pady=(0, 12))

        form = tk.Frame(top, bg="#ffffff", bd=1, relief="solid", padx=16, pady=16)
        form.pack(fill="x")

        tk.Label(form, text="Data od (do nazwy pliku)", font=("Segoe UI", 10, "bold"), bg="#ffffff", fg="#111827").grid(row=0, column=0, sticky="w")
        tk.Entry(form, textvariable=self.date_from_var, font=("Segoe UI", 12), relief="solid", bd=1, width=18).grid(row=1, column=0, sticky="w", padx=(0, 12), ipady=4)

        tk.Label(form, text="Data do (do nazwy pliku)", font=("Segoe UI", 10, "bold"), bg="#ffffff", fg="#111827").grid(row=0, column=1, sticky="w")
        tk.Entry(form, textvariable=self.date_to_var, font=("Segoe UI", 12), relief="solid", bd=1, width=18).grid(row=1, column=1, sticky="w", padx=(0, 16), ipady=4)

        ttk.Button(form, text="Otwórz KSeF", style="Ghost.TButton", command=self.start_browser).grid(row=0, column=2, rowspan=2, sticky="ew", padx=(0, 8))
        ttk.Button(form, text="Pobierz zestawienie", style="Main.TButton", command=self.run_full_export).grid(row=0, column=3, rowspan=2, sticky="ew")
        ttk.Button(form, text="Otwórz folder", style="Ghost.TButton", command=self.open_output_folder).grid(row=0, column=4, rowspan=2, sticky="ew", padx=(8, 0))

        form.grid_columnconfigure(5, weight=1)

        info = tk.Frame(outer, bg="#ffffff", bd=1, relief="solid", padx=16, pady=14)
        info.pack(fill="x", pady=(0, 12))

        tk.Label(info, text="Status", font=("Segoe UI", 10, "bold"), bg="#ffffff", fg="#111827").grid(row=0, column=0, sticky="w")
        tk.Label(info, textvariable=self.status_var, font=("Segoe UI", 11), bg="#ffffff", fg="#334155").grid(row=1, column=0, sticky="w", pady=(4, 0))

        tk.Label(info, text="Wynik", font=("Segoe UI", 10, "bold"), bg="#ffffff", fg="#111827").grid(row=0, column=1, sticky="w", padx=(32, 0))
        tk.Label(info, textvariable=self.summary_var, font=("Segoe UI", 11), bg="#ffffff", fg="#334155").grid(row=1, column=1, sticky="w", padx=(32, 0), pady=(4, 0))

        tk.Label(info, text="Plik", font=("Segoe UI", 10, "bold"), bg="#ffffff", fg="#111827").grid(row=0, column=2, sticky="w", padx=(32, 0))
        tk.Label(info, textvariable=self.file_var, font=("Segoe UI", 11), bg="#ffffff", fg="#334155").grid(row=1, column=2, sticky="w", padx=(32, 0), pady=(4, 0))

        self.progress = ttk.Progressbar(outer, mode="determinate", maximum=100, style="Red.Horizontal.TProgressbar")
        self.progress.pack(fill="x", pady=(0, 12))

        log_card = tk.Frame(outer, bg="#ffffff", bd=1, relief="solid")
        log_card.pack(fill="both", expand=True)

        tk.Label(log_card, text="Log", font=("Segoe UI", 14, "bold"), bg="#ffffff", fg="#111827", padx=16, pady=12).pack(anchor="w")

        self.log_box = tk.Text(
            log_card,
            font=("Consolas", 10),
            bg="#0f172a",
            fg="#e2e8f0",
            insertbackground="#ffffff",
            wrap="word",
            bd=0,
            relief="flat",
            padx=12,
            pady=12,
            height=18,
        )
        self.log_box.pack(fill="both", expand=True, padx=16, pady=(0, 16))
        self.log("[INFO] Aplikacja uruchomiona.")
        self.log("[INFO] Daty ustawiasz ręcznie w KSeF. Program tylko czyta aktualną listę i robi Excel.")

    def pump_ui(self):
        try:
            self.root.update_idletasks()
            self.root.update()
        except tk.TclError:
            pass

    def log(self, text: str):
        self.log_box.configure(state="normal")
        self.log_box.insert("end", text + "\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")
        self.pump_ui()

    def set_status(self, text: str):
        self.status_var.set(text)
        self.pump_ui()

    def update_progress(self, current: int, total: int, text: str):
        total = max(1, total)
        current = max(0, min(current, total))
        self.progress.configure(maximum=total)
        self.progress["value"] = current
        percent = int(current / total * 100)
        self.set_status(f"{text} ({percent}%)")

    def reset_progress(self, text: str):
        self.progress.configure(maximum=100)
        self.progress["value"] = 0
        self.set_status(text)

    def open_output_folder(self):
        try:
            os.makedirs(self.output_dir, exist_ok=True)
            os.startfile(self.output_dir)
        except Exception as e:
            messagebox.showerror("Błąd", f"Nie udało się otworzyć folderu.\n\n{e}")

    def parse_date(self, value: str) -> str:
        value = value.strip()
        for fmt in ("%Y-%m-%d", "%d-%m-%Y", "%d.%m.%Y", "%d/%m/%Y", "%Y/%m/%d"):
            try:
                return datetime.strptime(value, fmt).strftime("%Y-%m-%d")
            except ValueError:
                continue
        raise ValueError(f"Niepoprawna data: {value}")

    def normalize_spaces(self, text: str) -> str:
        return re.sub(r"\s+", " ", (text or "")).strip()

    def normalize_key(self, text: str) -> str:
        return self.normalize_spaces(text).lower()

    def canonical_header(self, header: str) -> str:
        key = self.normalize_key(header)
        for target, aliases in HEADER_ALIASES.items():
            if key == self.normalize_key(target):
                return target
            if key in [self.normalize_key(a) for a in aliases]:
                return target
        return header.strip()

    def start_browser(self):
        try:
            if self.page is not None:
                messagebox.showinfo("Informacja", "Przeglądarka jest już otwarta.")
                return

            self.reset_progress("Uruchamianie przeglądarki")
            self.log("[INFO] Uruchamianie przeglądarki...")

            self.playwright = sync_playwright().start()
            self.browser = self.playwright.chromium.launch(headless=False, slow_mo=80)
            self.context = self.browser.new_context(accept_downloads=True)
            self.page = self.context.new_page()
            self.page.goto("https://ap.ksef.mf.gov.pl/web/invoice-list", timeout=90000)

            self.reset_progress("Zaloguj się do KSeF")
            self.log("[OK] KSeF otwarty.")
            self.log("[INFO] Zaloguj się ręcznie, kliknij 'Rozwiń filtry', ustaw daty i pokaż listę faktur zakupu.")
            messagebox.showinfo(
                "KSeF",
                "KSeF został otwarty.\n\n"
                "1. Zaloguj się ręcznie\n"
                "2. Kliknij 'Rozwiń filtry'\n"
                "3. Ustaw daty ręcznie\n"
                "4. Pokaż listę faktur zakupu\n"
                "5. Kliknij 'Pobierz zestawienie' w programie",
            )
        except Exception as e:
            self.reset_progress("Błąd")
            self.log(f"[BŁĄD] Nie udało się otworzyć przeglądarki: {e}")
            messagebox.showerror("Błąd", f"Nie udało się otworzyć przeglądarki.\n\n{e}")

    def get_table_headers(self) -> List[str]:
        for selector in ["thead th", "[role='columnheader']"]:
            try:
                loc = self.page.locator(selector)
                headers = []
                for i in range(loc.count()):
                    txt = self.normalize_spaces(loc.nth(i).inner_text())
                    if txt:
                        headers.append(txt)
                if headers:
                    return headers
            except Exception:
                pass
        return []

    def extract_item_key(self, text: str) -> str:
        text = self.normalize_spaces(text)
        patterns = [
            r"(\d{10,}-\d{8}-[A-Z0-9]+-[A-Z0-9]+)",
            r"([A-Z0-9/\-]{6,}/\d{4})",
            r"([A-Z0-9]{8,})",
        ]
        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                return match.group(1).lower()
        return text.lower()

    def get_current_page_rows(self) -> List[Dict]:
        rows_data = []
        rows = None
        for selector in ["tbody tr", "table tbody tr", "[role='row']"]:
            try:
                loc = self.page.locator(selector)
                if loc.count() > 0:
                    rows = loc
                    break
            except Exception:
                pass

        if rows is None:
            return rows_data

        for i in range(rows.count()):
            try:
                row = rows.nth(i)
                if not row.is_visible():
                    continue

                cell_loc = row.locator("td, [role='cell']")
                cells = []
                for j in range(cell_loc.count()):
                    txt = self.normalize_spaces(cell_loc.nth(j).inner_text())
                    if txt:
                        cells.append(txt)

                if len(cells) < 4:
                    continue

                row_text = " | ".join(cells)
                rows_data.append({
                    "cells": cells,
                    "row_id": self.extract_item_key(row_text),
                    "row_text": row_text,
                })
            except Exception:
                pass

        return rows_data

    def get_page_signature(self) -> str:
        rows = self.get_current_page_rows()
        if not rows:
            return "EMPTY"
        ids = [row["row_id"] for row in rows[:5]]
        if len(rows) > 5:
            ids.extend(row["row_id"] for row in rows[-3:])
        return "|".join(ids)

    def _wait_for_signature_change(self, before: str, timeout_ms: int = 7000) -> bool:
        waited = 0
        while waited < timeout_ms:
            self.page.wait_for_timeout(250)
            waited += 250
            after = self.get_page_signature()
            if after and after != before:
                return True
            self.pump_ui()
        return False

    def _click_pagination_via_js(self, direction: str) -> bool:
        script = """
        (direction) => {
            const normalize = (v) => (v || '').toLowerCase().replace(/\\s+/g, ' ').trim();
            const isVisible = (el) => {
                const style = window.getComputedStyle(el);
                const rect = el.getBoundingClientRect();
                return style && style.display !== 'none' && style.visibility !== 'hidden' && rect.width > 0 && rect.height > 0;
            };
            const isDisabled = (el) => {
                const cls = normalize(el.className || '');
                return !!el.disabled || el.getAttribute('aria-disabled') === 'true' || cls.includes('disabled');
            };
            const words = direction === 'next'
                ? ['następna', 'następny', 'next', 'dalej', '>', '›', '»']
                : ['poprzednia', 'poprzedni', 'previous', 'wstecz', '<', '‹', '«'];

            const els = [...document.querySelectorAll('button, a, [role="button"]')];
            const candidates = [];

            for (const el of els) {
                if (!isVisible(el) || isDisabled(el)) continue;
                const text = normalize(el.innerText || el.textContent || '');
                const aria = normalize(el.getAttribute('aria-label'));
                const title = normalize(el.getAttribute('title'));
                const combined = [text, aria, title].join(' | ');
                if (words.some(w => combined.includes(w))) {
                    candidates.push(el);
                }
            }

            for (const el of candidates) {
                try {
                    el.scrollIntoView({block: 'center'});
                    el.click();
                    return true;
                } catch (e) {}
            }
            return false;
        }
        """
        try:
            return bool(self.page.evaluate(script, direction))
        except Exception:
            return False

    def go_to_first_page(self, max_steps: int = 50):
        selectors = [
            "button[aria-label*='Poprzednia']",
            "button[title*='Poprzednia']",
            "[role='button'][aria-label*='Poprzednia']",
            "text=Poprzednia",
            "text=Previous",
            "button:has-text('<')",
        ]

        for _ in range(max_steps):
            moved = False
            before = self.get_page_signature()

            for selector in selectors:
                try:
                    loc = self.page.locator(selector)
                    if loc.count() == 0:
                        continue
                    btn = loc.first
                    disabled = btn.get_attribute("disabled")
                    aria_disabled = btn.get_attribute("aria-disabled")
                    classes = (btn.get_attribute("class") or "").lower()
                    if disabled is not None or aria_disabled == "true" or "disabled" in classes:
                        continue
                    btn.click(timeout=2500)
                    if self._wait_for_signature_change(before, 4000):
                        moved = True
                        break
                except Exception:
                    pass

            if not moved:
                if self._click_pagination_via_js("prev") and self._wait_for_signature_change(before, 4000):
                    moved = True

            if not moved:
                break

            self.pump_ui()

    def go_to_next_page(self) -> bool:
        selectors = [
            "button[aria-label*='Następna']",
            "button[title*='Następna']",
            "[role='button'][aria-label*='Następna']",
            "button[aria-label*='Dalej']",
            "button[title*='Dalej']",
            "[role='button'][aria-label*='Dalej']",
            "text=Następna",
            "text=Next",
            "text=Dalej",
            "button:has-text('>')",
            "button:has-text('›')",
            "button:has-text('»')",
        ]

        before = self.get_page_signature()

        for selector in selectors:
            try:
                loc = self.page.locator(selector)
                if loc.count() == 0:
                    continue
                btn = loc.first
                disabled = btn.get_attribute("disabled")
                aria_disabled = btn.get_attribute("aria-disabled")
                classes = (btn.get_attribute("class") or "").lower()
                if disabled is not None or aria_disabled == "true" or "disabled" in classes:
                    continue
                btn.click(timeout=3000)
                if self._wait_for_signature_change(before, 6000):
                    return True
            except Exception:
                pass

        if self._click_pagination_via_js("next") and self._wait_for_signature_change(before, 6000):
            return True

        return False

    def scan_all_pages(self) -> Tuple[List[str], List[Dict], int]:
        self.go_to_first_page()
        self.page.wait_for_timeout(900)

        headers = self.get_table_headers()
        all_rows: List[Dict] = []
        seen_pages = set()
        seen_rows = set()
        pages = 0

        while True:
            self.pump_ui()
            current_rows = self.get_current_page_rows()
            signature = self.get_page_signature()

            if signature in seen_pages:
                self.log("[INFO] Wykryto powrót do tej samej strony. Kończę skanowanie.")
                break

            seen_pages.add(signature)
            pages += 1
            self.log(f"[INFO] Odczyt strony {pages}: {len(current_rows)} wierszy")

            for row in current_rows:
                unique_key = row["row_id"] + "||" + row["row_text"]
                if unique_key in seen_rows:
                    continue
                seen_rows.add(unique_key)
                all_rows.append(row)

            self.update_progress(pages, max(1, pages + 1), "Skanowanie listy")

            if not self.go_to_next_page():
                self.log("[INFO] Nie znaleziono kolejnej strony. Koniec skanowania.")
                break

        self.go_to_first_page()
        self.page.wait_for_timeout(700)
        return headers, all_rows, pages

    def build_raw_headers(self, headers: List[str], rows: List[Dict]) -> List[str]:
        if headers and all(len(r["cells"]) == len(headers) for r in rows[: min(20, len(rows))] if r["cells"]):
            return headers
        max_len = max((len(r["cells"]) for r in rows), default=0)
        return [f"Kolumna {i + 1}" for i in range(max_len)]

    def map_to_target_rows(self, headers: List[str], rows: List[Dict]) -> List[Dict[str, str]]:
        raw_headers = self.build_raw_headers(headers, rows)
        canonical = [self.canonical_header(h) for h in raw_headers]
        mapped = []
        for row in rows:
            row_dict = {}
            for idx, header in enumerate(canonical):
                row_dict[header] = row["cells"][idx] if idx < len(row["cells"]) else ""
            normalized = {col: row_dict.get(col, "") for col in TARGET_COLUMNS}
            if not normalized["Nr KSeF"] and len(row["cells"]) >= 11:
                for idx, target in enumerate(TARGET_COLUMNS):
                    normalized[target] = row["cells"][idx]
            mapped.append(normalized)
        return mapped

    def to_number_if_possible(self, value: str):
        text = "" if value is None else str(value).strip().replace(" ", "")
        if not text:
            return ""
        text = text.replace(",", ".")
        if re.fullmatch(r"-?\d+(\.\d+)?", text):
            try:
                return float(text) if "." in text else int(text)
            except ValueError:
                return value
        return value

    def style_header_row(self, ws, row_no: int = 1):
        fill = PatternFill("solid", fgColor="C81F25")
        font = Font(bold=True, color="FFFFFF")
        border = Border(
            left=Side(style="thin", color="D1D5DB"),
            right=Side(style="thin", color="D1D5DB"),
            top=Side(style="thin", color="D1D5DB"),
            bottom=Side(style="thin", color="D1D5DB"),
        )
        for cell in ws[row_no]:
            cell.fill = fill
            cell.font = font
            cell.border = border
            cell.alignment = Alignment(horizontal="center", vertical="center")

    def auto_fit_worksheet(self, ws):
        for col_cells in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col_cells[0].column)
            for cell in col_cells:
                val = "" if cell.value is None else str(cell.value)
                max_len = max(max_len, len(val))
            ws.column_dimensions[col_letter].width = min(max(max_len + 2, 12), 40)

    def save_excel(self, headers: List[str], rows: List[Dict]) -> str:
        if not rows:
            raise ValueError("Brak danych do zapisania.")

        date_from = self.parse_date(self.date_from_var.get())
        date_to = self.parse_date(self.date_to_var.get())
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"KSeF_zestawienie_{date_from}_{date_to}_{timestamp}.xlsx"
        filepath = os.path.join(self.output_dir, filename)

        wb = Workbook()
        ws_info = wb.active
        ws_info.title = "Info"
        ws_info.append(["Parametr", "Wartość"])
        ws_info.append(["Data od (opis pliku)", date_from])
        ws_info.append(["Data do (opis pliku)", date_to])
        ws_info.append(["Data eksportu", datetime.now().strftime("%Y-%m-%d %H:%M:%S")])
        ws_info.append(["Liczba wierszy", len(rows)])
        ws_info.append(["Liczba stron", self.last_pages])
        self.style_header_row(ws_info)
        self.auto_fit_worksheet(ws_info)

        ws_sum = wb.create_sheet("Zestawienie")
        ws_sum.append(TARGET_COLUMNS)
        for row in self.map_to_target_rows(headers, rows):
            ws_sum.append([
                row["Identyfikator sprzedawcy"],
                row["Nazwa sprzedawcy"],
                row["Nr KSeF"],
                row["Nr faktury"],
                row["Data wystawienia"],
                row["Data zapisania w KSeF"],
                row["Data otrzymania"],
                row["Waluta"],
                self.to_number_if_possible(row["Netto"]),
                self.to_number_if_possible(row["Brutto"]),
                self.to_number_if_possible(row["VAT (PLN)"]),
            ])
        self.style_header_row(ws_sum)
        self.auto_fit_worksheet(ws_sum)
        ws_sum.freeze_panes = "A2"

        wb.save(filepath)
        return filepath

    def run_full_export(self):
        try:
            if self.page is None:
                messagebox.showwarning("Uwaga", "Najpierw kliknij 'Otwórz KSeF'.")
                return

            self.parse_date(self.date_from_var.get())
            self.parse_date(self.date_to_var.get())

            self.reset_progress("Przygotowanie")
            self.log("[INFO] Start zestawienia.")
            self.log("[INFO] Program NIE ustawia dat. Czytam dokładnie to, co aktualnie masz wyświetlone w KSeF.")
            self.log("[INFO] Skanuję wszystkie strony listy...")

            headers, rows, pages = self.scan_all_pages()
            self.last_headers = headers
            self.last_rows = rows
            self.last_pages = pages

            if not rows:
                self.reset_progress("Brak danych")
                self.summary_var.set("Nie znaleziono żadnych wierszy")
                messagebox.showwarning("Brak danych", "Nie znaleziono żadnych wierszy na liście.")
                return

            self.update_progress(2, 3, "Zapisywanie Excel")
            filepath = self.save_excel(headers, rows)
            self.update_progress(3, 3, "Gotowe")

            self.summary_var.set(f"Znaleziono {len(rows)} wierszy na {pages} stronach")
            self.file_var.set(Path(filepath).name)
            self.log(f"[OK] Zapisano plik: {filepath}")
            messagebox.showinfo("Gotowe", f"Zestawienie zapisane.\n\n{filepath}")
        except Exception as e:
            self.reset_progress("Błąd")
            self.log(f"[BŁĄD] {e}")
            messagebox.showerror("Błąd", str(e))

    def on_close(self):
        try:
            if self.context:
                self.context.close()
        except Exception:
            pass
        try:
            if self.browser:
                self.browser.close()
        except Exception:
            pass
        try:
            if self.playwright:
                self.playwright.stop()
        except Exception:
            pass
        self.root.destroy()


if __name__ == "__main__":
    if sys.platform.startswith("win"):
        try:
            from ctypes import windll
            windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass

    root = tk.Tk()
    app = KsefSimpleSummaryApp(root)
    root.mainloop()
