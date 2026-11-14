import calendar
import tkinter as tk
from tkinter import ttk, messagebox
import hashlib
import hmac
import json
import os
import shutil
import sqlite3
import ssl
import sys
import traceback
import urllib.request
from datetime import datetime
from pathlib import Path

import certifi
from openpyxl import Workbook
from openpyxl.styles import Font

# Fixed per-user folder, so the data survives moving the .exe around.
DATA_DIR = Path(os.environ.get("APPDATA", Path.home())) / "OtelYonetim"

# Release builds stamp the tag number here (tag v3 -> "3"); "0" is a test build.
APP_VERSION = "0"
RELEASE_API = "https://api.github.com/repos/Edoomaan/otel-yonetim/releases/latest"
EXE_NAME = "OtelYonetim.exe"
# Python on Windows only sees the certificates Windows has already fetched, which on a fresh
# install is too few to verify github.com. certifi ships a complete list inside the .exe.
SSL_CONTEXT = ssl.create_default_context(cafile=certifi.where())
DATE_FORMAT = "%d.%m.%Y"
PASSWORD_FILE = DATA_DIR / "parola.txt"  # deleting this file makes the program ask for a new password

MONTHS = ["Ocak", "Şubat", "Mart", "Nisan", "Mayıs", "Haziran", "Temmuz", "Ağustos", "Eylül", "Ekim", "Kasım", "Aralık"]
WEEKDAYS = ["Pazartesi", "Salı", "Çarşamba", "Perşembe", "Cuma", "Cumartesi", "Pazar"]
# Same columns as the hotel's own "KAYIT DEFTERİ" workbook, so a day can be copied straight into it.
# The program does not collect RESEPSİYONİST, TELEFON, ÇIKIŞ SAATİ or EFT GÖNDEREN; those stay empty.
REGISTER_HEADERS = {"B": "SIRA NO", "C": "ODA NO", "D": "RESEPSİYONİST", "F": "TC NO", "H": "ADI SOYADI", "J": "TELEFON",
                    "K": "DOĞUM YILI", "L": "ÇIKIŞ SAATİ", "N": "ÜCRET", "R": "EFT"}
REGISTER_SUBHEADERS = {"N": "NAKİT", "O": "FİŞ", "P": "EFT", "Q": "TOPLAM", "R": "GÖNDEREN"}
PAYMENT_COLUMNS = {"Nakit": "N", "Fiş": "O", "EFT": "P"}


def is_valid_tc(tc):
    """Check the two control digits of a Turkish ID number."""
    if not (tc.isdigit() and len(tc) == 11) or tc[0] == "0":
        return False
    d = [int(c) for c in tc]
    # d[0:9:2] = 1st, 3rd, 5th, 7th, 9th digits; d[1:8:2] = 2nd, 4th, 6th, 8th
    tenth = (sum(d[0:9:2]) * 7 - sum(d[1:8:2])) % 10
    eleventh = sum(d[:10]) % 10
    return d[9] == tenth and d[10] == eleventh


def parse_date(text):
    """Return the date for a GG.AA.YYYY text, or None when it is not a real date."""
    try:
        return datetime.strptime(text, DATE_FORMAT)
    except ValueError:
        return None


def hash_password(password, salt):
    # Deliberately slow (200,000 rounds) so guessing passwords one after another is slow too.
    return hashlib.pbkdf2_hmac("sha256", password.encode(), salt, 200_000)


def make_password_record(password):
    """Return "salt:hash" to store on disk; the password itself is never stored."""
    salt = os.urandom(16)
    return f"{salt.hex()}:{hash_password(password, salt).hex()}"


def check_password(password, record):
    try:
        salt, expected = (bytes.fromhex(part) for part in record.strip().split(":"))
    except ValueError:  # damaged file
        return False
    return hmac.compare_digest(hash_password(password, salt), expected)


def backup_database(db_path, backup_dir, today, keep=30):
    """Copy the database once per day into backup_dir and keep only the newest `keep` copies."""
    target = backup_dir / f"hotel_{today:%Y-%m-%d}.db"
    if not db_path.exists() or target.exists():
        return
    backup_dir.mkdir(parents=True, exist_ok=True)
    shutil.copy2(db_path, target)
    # The date in the name sorts as text in time order, so the oldest files come first.
    for old in sorted(backup_dir.glob("hotel_*.db"))[:-keep]:
        old.unlink()


def is_newer(tag, current):
    """True when a release tag like "v3" is a higher number than the running version."""
    try:
        return int(tag.lstrip("v")) > int(current)
    except ValueError:
        return False


def format_date(text):
    """Turn typed digits into GG.AA.YYYY, adding each dot as soon as its part is complete."""
    d = "".join(c for c in text if c.isdigit())[:8]
    if len(d) < 2:
        return d
    if len(d) < 4:
        return f"{d[:2]}.{d[2:]}"
    return f"{d[:2]}.{d[2:4]}.{d[4:]}"


def validate(tc, birth_year, reg_date):
    """Return an error message, or None when the input is fine."""
    if not is_valid_tc(tc):
        return "Geçersiz TC Kimlik No! 11 haneyi kontrol edip tekrar girin."
    if birth_year and not (birth_year.isdigit() and 1900 <= int(birth_year) <= datetime.now().year):
        return "Doğum yılı 4 haneli bir yıl olmalıdır! (örnek: 1990)"
    if not parse_date(reg_date):
        return "Kayıt tarihi GG.AA.YYYY biçiminde olmalıdır! (örnek: 25.11.2026)"
    return None


def load_stays(conn):
    """Group every reservation as {(year, month): {day: {room: [(name, tc, birth_year, payment, price), ...]}}}."""
    stays = {}
    rows = conn.execute(
        "SELECT room_no, guest_name, tc_no, birth_year, check_in, payment_type, price FROM reservations ORDER BY id")
    for room, name, tc, birth, check_in, payment, price in rows:
        day = parse_date(check_in or "")
        if day:
            # shortcut: two separate stays in one room on one day share a block; split them by stay if that ever matters
            stays.setdefault((day.year, day.month), {}).setdefault(day.day, {}).setdefault(room, []).append(
                (name, tc, (birth or "")[-4:], payment, price or 0))  # [-4:]: older records hold a full birth date
    return stays


def save_workbook(wb, path):
    for ws in wb:
        for row in ws.iter_rows():
            for cell in row:
                if cell.data_type == "f":  # text starting with "=" would run as a formula when Excel opens the file
                    cell.data_type = "s"
    wb.save(path)


def export_register(stays, path):
    """Guest register in the layout of the hotel's "KAYIT DEFTERİ": a sheet per month, a block per day."""
    wb = Workbook()
    wb.remove(wb.active)
    bold = Font(bold=True)
    for (year, month), days in sorted(stays.items()):
        ws = wb.create_sheet(f"{MONTHS[month - 1]} {year}")
        row = 2
        for day, rooms in sorted(days.items()):
            ws[f"B{row}"] = datetime(year, month, day)
            ws[f"B{row}"].number_format = "DD.MM.YYYY"
            ws[f"B{row}"].font = bold
            for headers in (REGISTER_HEADERS, REGISTER_SUBHEADERS):
                row += 1
                for column, text in headers.items():
                    ws[f"{column}{row}"] = text
                    ws[f"{column}{row}"].font = bold
            row += 1
            day_totals = dict.fromkeys(PAYMENT_COLUMNS.values(), 0)
            for number, (room, guests) in enumerate(sorted(rooms.items()), 1):
                ws[f"B{row}"], ws[f"C{row}"] = number, int(room)
                paid = dict.fromkeys(PAYMENT_COLUMNS.values(), 0)
                for offset, (name, tc, birth, payment, price) in enumerate(guests):
                    ws[f"F{row + offset}"], ws[f"H{row + offset}"] = tc, name
                    ws[f"K{row + offset}"] = int(birth) if birth.isdigit() else birth
                    paid[PAYMENT_COLUMNS.get(payment, "N")] += price
                for column, amount in paid.items():
                    ws[f"{column}{row}"] = amount or None  # the room's money goes on its first line only
                    day_totals[column] += amount
                ws[f"Q{row}"] = sum(paid.values())
                row += max(2, len(guests))  # the hotel's sheet gives every room two lines
            ws[f"M{row}"] = "TOPLAM"
            for column, amount in day_totals.items():
                ws[f"{column}{row}"] = amount
            ws[f"Q{row}"] = sum(day_totals.values())
            for column in "MNOPQ":
                ws[f"{column}{row}"].font = bold
            row += 3
        for column, width in {"B": 11, "D": 15, "F": 14, "H": 30, "J": 14, "K": 12, "L": 12}.items():
            ws.column_dimensions[column].width = width
    save_workbook(wb, path)


def export_daily(stays, path, rooms):
    """Money per room per day in the layout of the hotel's "GÜNLÜK ALINAN": a sheet per month, a row per day."""
    wb = Workbook()
    wb.remove(wb.active)
    for (year, month), days in sorted(stays.items()):
        ws = wb.create_sheet(f"{MONTHS[month - 1].replace('i', 'İ').upper()} {year}")  # plain upper() would give NISAN
        ws.append(["GÜNLER", *map(int, rooms), "TOPLAM", "GÜNLER", "S.O.S"])
        month_days = calendar.monthrange(year, month)[1]
        room_totals, sold_total = [0] * len(rooms), 0
        for day in range(1, month_days + 1):
            day_rooms = days.get(day, {})
            amounts = [sum(guest[4] for guest in day_rooms.get(room, [])) for room in rooms]
            ws.append([day, *(amount or None for amount in amounts), sum(amounts),
                       WEEKDAYS[datetime(year, month, day).weekday()], len(day_rooms) or None])
            room_totals = [total + amount for total, amount in zip(room_totals, amounts)]
            sold_total += len(day_rooms)
        # Row 33 as in the hotel's sheet, whatever the month length. S.O.S = number of rooms sold that day.
        totals = ["TOPLAM", *room_totals, sum(room_totals), round(sum(room_totals) / month_days), sold_total]
        for column, value in enumerate(totals, 1):
            ws.cell(row=33, column=column, value=value).font = Font(bold=True)
        for cell in ws[1]:
            cell.font = Font(bold=True)
    save_workbook(wb, path)


class HotelApp:
    def __init__(self, root):
        self.root = root
        self.root.title("İzmit My House Butik - Otel Yönetim Sistemi")
        self.root.geometry("1300x780")
        if sys.platform == "win32":
            self.root.state("zoomed")  # start maximized so nothing falls off a small screen
        # A windowed .exe has no console: without this, an error in a button handler would vanish silently.
        self.root.report_callback_exception = lambda *exc_info: messagebox.showerror(
            "Beklenmeyen Hata", "Bir hata oluştu. Lütfen bu pencerenin fotoğrafını gönderin.\n\n"
                                + "".join(traceback.format_exception(*exc_info))[-1500:])

        # Oda tipleri dağılımı
        self.room_types = {
            # 1. Kat
            "101": "Büyük Oda", "102": "Büyük Oda", "103": "Büyük Oda", "104": "Küçük Oda", "105": "Küçük Oda",
            # 2. Kat
            "201": "Büyük Oda", "202": "Büyük Oda", "203": "Büyük Oda", "204": "Küçük Oda", "205": "Küçük Oda",
            # 3. Kat
            "301": "Küçük Oda", "302": "Büyük Oda", "303": "Büyük Oda", "304": "Küçük Oda", "305": "Küçük Oda",
            # 4. Kat
            "401": "Küçük Oda", "402": "Küçük Oda", "403": "Büyük Oda", "404": "Suit Oda"
        }

        self.rooms = list(self.room_types.keys())

        # Fiyat tarifeleri
        self.prices = {
            "Küçük Oda": {"Nakit": 1700, "Diger": 1850},
            "Büyük Oda": {"Nakit": 1800, "Diger": 1950},
            "Suit Oda": {"Nakit": 2050, "Diger": 2150}
        }

        if getattr(sys, "frozen", False):  # leftover from the last update; it could not delete itself while running
            try:
                Path(sys.executable).with_suffix(".old").unlink(missing_ok=True)
            except OSError:
                pass

        # Veritabanı bağlantısı
        DATA_DIR.mkdir(parents=True, exist_ok=True)
        try:  # before connecting, so the file is copied while nothing is writing to it
            backup_database(DATA_DIR / "hotel.db", DATA_DIR / "yedekler", datetime.now())
        except OSError:
            pass  # a failed backup must not stop the reception from working
        self.conn = sqlite3.connect(DATA_DIR / "hotel.db")
        self.cursor = self.conn.cursor()
        self.create_table()

        # Stil
        style = ttk.Style()
        style.theme_use("clam")

        # Sekmeler
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=10)

        self.tab_dashboard = ttk.Frame(self.notebook)
        self.tab_reservations = ttk.Frame(self.notebook)
        self.tab_rooms = ttk.Frame(self.notebook)
        self.tab_finance = ttk.Frame(self.notebook)

        self.notebook.add(self.tab_dashboard, text="🏠 Ana Kontrol Paneli")
        self.notebook.add(self.tab_reservations, text="📋 Rezervasyon & Müşteri Kayıtları")
        self.notebook.add(self.tab_rooms, text="🛏️ Kat ve Oda Durumları")
        self.notebook.add(self.tab_finance, text="💰 Kasa & Finans Raporu")

        self.setup_dashboard()
        self.setup_reservations()
        self.setup_rooms()
        self.setup_finance()
        self.refresh_dashboard()

    def create_table(self):
        self.cursor.execute('''
                            CREATE TABLE IF NOT EXISTS reservations
                            (
                                id
                                INTEGER
                                PRIMARY
                                KEY
                                AUTOINCREMENT,
                                room_no
                                TEXT,
                                guest_name
                                TEXT,
                                tc_no
                                TEXT,
                                birth_year -- shortcut: now holds a full date; renaming needs a migration of existing databases
                                TEXT,
                                check_in
                                TEXT,
                                check_out
                                TEXT,
                                payment_type
                                TEXT,
                                extra_person
                                TEXT,
                                price
                                REAL,
                                status
                                TEXT
                            )
                            ''')
        try:  # databases created by versions 1-3 do not have this column yet
            self.cursor.execute("ALTER TABLE reservations ADD COLUMN extra_income REAL DEFAULT 0")
        except sqlite3.OperationalError:
            pass  # already there
        # Versions 1-4 called this payment type "Kredi Kartı"; the hotel's own books call it "Fiş".
        self.cursor.execute("UPDATE reservations SET payment_type='Fiş' WHERE payment_type='Kredi Kartı'")
        self.conn.commit()

    def setup_dashboard(self):
        lbl = ttk.Label(self.tab_dashboard, text="İzmit My House Butik Yönetim Paneline Hoş Geldiniz",
                        font=("Arial", 16, "bold"))
        lbl.pack(pady=20)

        frame = ttk.Frame(self.tab_dashboard)
        frame.pack(pady=10)

        self.lbl_total_res = ttk.Label(frame, text="Aktif Rezervasyon: 0", font=("Arial", 12))
        self.lbl_total_res.grid(row=0, column=0, padx=20, pady=10)

        btn_refresh = ttk.Button(self.tab_dashboard, text="Verileri Yenile", command=self.refresh_dashboard)
        btn_refresh.pack(pady=20)

        self.btn_update = ttk.Button(self.tab_dashboard, text="🔄 Güncellemeleri Denetle", command=self.check_for_update)
        self.btn_update.pack(pady=(40, 5))
        self.lbl_update = ttk.Label(self.tab_dashboard, text=f"Sürüm: {APP_VERSION}")
        self.lbl_update.pack()

    def check_for_update(self):
        if not getattr(sys, "frozen", False):  # only the PyInstaller .exe can replace itself
            messagebox.showinfo("Güncelleme", "Güncelleme yalnızca .exe sürümünde çalışır.")
            return
        exe = Path(sys.executable)
        new, old = exe.with_suffix(".new"), exe.with_suffix(".old")
        self.btn_update.config(state="disabled")
        try:
            request = urllib.request.Request(RELEASE_API, headers={"User-Agent": "OtelYonetim"})
            with urllib.request.urlopen(request, timeout=15, context=SSL_CONTEXT) as response:
                release = json.load(response)
            if not is_newer(release["tag_name"], APP_VERSION):
                messagebox.showinfo("Güncelleme", "Program güncel.")
                return
            asset = next(a for a in release["assets"] if a["name"] == EXE_NAME)
            if not messagebox.askyesno("Güncelleme", "Yeni sürüm var. Şimdi indirilsin mi?\n\nKayıtlarınız silinmez."):
                return
            sha, done = hashlib.sha256(), 0
            with urllib.request.urlopen(asset["browser_download_url"], timeout=30, context=SSL_CONTEXT) as response, open(new, "wb") as f:
                while chunk := response.read(256 * 1024):
                    f.write(chunk)
                    sha.update(chunk)
                    done += len(chunk)
                    self.lbl_update.config(text=f"İndiriliyor… %{done * 100 // asset['size']}")
                    self.root.update()  # keep the window alive during the download
            digest = asset.get("digest")  # "sha256:<hex>"; a half-downloaded exe must never replace the working one
            if done != asset["size"] or (digest and digest != "sha256:" + sha.hexdigest()):
                raise OSError("download is incomplete or corrupted")
            old.unlink(missing_ok=True)
            exe.rename(old)  # Windows lets a running .exe be renamed, but not overwritten
            try:
                new.rename(exe)
            except OSError:
                old.rename(exe)  # put the working program back
                raise
        except (OSError, KeyError, StopIteration, ValueError) as error:
            new.unlink(missing_ok=True)
            # The detail line is for whoever supports the user: it names the real cause.
            messagebox.showerror("Güncelleme", "Güncelleme yapılamadı. İnternet bağlantısını kontrol edip tekrar deneyin.\n\n"
                                               f"Ayrıntı: {type(error).__name__}: {error}")
            return
        finally:
            self.btn_update.config(state="normal")
            self.lbl_update.config(text=f"Sürüm: {APP_VERSION}")
        messagebox.showinfo("Güncelleme", "Güncelleme tamamlandı. Program şimdi kapanacak; lütfen yeniden açın.")
        self.root.destroy()

    def refresh_dashboard(self):
        self.cursor.execute("SELECT COUNT(*) FROM reservations WHERE status='Aktif'")
        count = self.cursor.fetchone()[0]
        self.lbl_total_res.config(text=f"Toplam Aktif Konaklayan Misafir: {count}")
        self.setup_rooms_grid()
        self.update_finance_report()

    def setup_reservations(self):
        left_frame = self.form_frame = ttk.LabelFrame(self.tab_reservations)
        left_frame.pack(side="left", fill="y", padx=10, pady=10, ipadx=10, ipady=10)

        room_display_list = [f"{room} ({self.room_types[room]})" for room in self.rooms]

        ttk.Label(left_frame, text="Oda No:").grid(row=0, column=0, sticky="w", pady=4)
        self.room_cb = ttk.Combobox(left_frame, values=room_display_list, width=22, state="readonly")
        self.room_cb.grid(row=0, column=1, pady=4)
        self.room_cb.current(0)
        self.room_cb.bind("<<ComboboxSelected>>", self.update_price)

        ttk.Label(left_frame, text="Kalan Adı Soyadı:").grid(row=1, column=0, sticky="w", pady=4)
        self.entry_name = ttk.Entry(left_frame)
        self.entry_name.grid(row=1, column=1, pady=4)

        ttk.Label(left_frame, text="TC Kimlik No:").grid(row=2, column=0, sticky="w", pady=4)
        def digits_only(limit):
            # Reject the keystroke unless the field would still be at most `limit` digits (%P = text after the edit).
            return self.root.register(lambda text: len(text) <= limit and (text.isdigit() or text == "")), "%P"

        self.entry_tc = ttk.Entry(left_frame, validate="key", validatecommand=digits_only(11))
        self.entry_tc.grid(row=2, column=1, pady=4)

        ttk.Label(left_frame, text="Doğum Yılı:").grid(row=3, column=0, sticky="w", pady=4)
        self.entry_birth = ttk.Entry(left_frame, validate="key", validatecommand=digits_only(4))
        self.entry_birth.grid(row=3, column=1, pady=4)

        ttk.Label(left_frame, text="Kayıt Tarihi (GG.AA.YYYY):").grid(row=4, column=0, sticky="w", pady=4)
        self.entry_date = ttk.Entry(left_frame)
        self.entry_date.grid(row=4, column=1, pady=4)
        self.entry_date.bind("<KeyRelease>", self.on_date_key)

        ttk.Label(left_frame, text="Ödeme Şekli:").grid(row=5, column=0, sticky="w", pady=4)
        self.payment_cb = ttk.Combobox(left_frame, values=["Nakit", "Fiş", "EFT"], width=22, state="readonly")
        self.payment_cb.grid(row=5, column=1, pady=4)
        self.payment_cb.current(0)
        self.payment_cb.bind("<<ComboboxSelected>>", self.update_price)

        ttk.Label(left_frame, text="Ek Gelir (TL):").grid(row=6, column=0, sticky="w", pady=4)
        self.entry_extra = ttk.Entry(left_frame)
        self.entry_extra.grid(row=6, column=1, pady=4)
        self.entry_extra.bind("<KeyRelease>", self.update_price)

        ttk.Label(left_frame, text="Toplam Ücret (TL):").grid(row=7, column=0, sticky="w", pady=4)
        self.entry_price = ttk.Entry(left_frame)
        self.entry_price.grid(row=7, column=1, pady=4)
        self.lbl_price_note = ttk.Label(left_frame, foreground="#b35c00", wraplength=300)
        self.lbl_price_note.grid(row=8, column=0, columnspan=2)

        self.btn_save = ttk.Button(left_frame, command=self.save_reservation)
        self.btn_save.grid(row=9, column=0, columnspan=2, pady=(12, 4))
        self.btn_cancel = ttk.Button(left_frame, text="Vazgeç", command=self.reset_form)
        self.btn_cancel.grid(row=10, column=0, columnspan=2)
        self.reset_form()

        right_frame = ttk.Frame(self.tab_reservations)
        right_frame.pack(side="right", fill="both", expand=True, padx=10, pady=10)

        search_frame = ttk.Frame(right_frame)
        search_frame.pack(fill="x", pady=5)
        ttk.Label(search_frame, text="🔍 Müşteri Adı veya Oda No Ara:").pack(side="left", padx=5)
        self.search_entry = ttk.Entry(search_frame, width=30)
        self.search_entry.pack(side="left", padx=5)
        self.search_entry.bind("<KeyRelease>", self.load_reservations)

        cols = ("ID", "Oda No", "Oda Tipi", "Kalan Adı", "TC No", "Doğum Yılı", "Kayıt Tarihi", "Ödeme", "Ek Gelir", "Ücret")
        self.tree = ttk.Treeview(right_frame, columns=cols, show="headings")
        for col in cols:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=80)
        self.tree.bind("<Double-1>", lambda event: self.tree.selection() and self.edit_reservation())

        # Packed before the table and pinned to the bottom: a short window shrinks the table, not the buttons.
        btn_action_frame = ttk.Frame(right_frame)
        btn_action_frame.pack(side="bottom", fill="x", pady=5)
        self.tree.pack(fill="both", expand=True, pady=5)
        btn_checkout = ttk.Button(btn_action_frame, text="🚪 Seçili Odanın Çıkışını Yap",
                                  command=self.checkout_reservation)
        btn_checkout.pack(side="left", padx=5)
        ttk.Button(btn_action_frame, text="✏️ Seçili Kaydı Düzenle", command=self.edit_reservation).pack(side="left", padx=5)
        ttk.Button(btn_action_frame, text="📊 Tüm Kayıtları Excel'e Aktar",
                   command=self.export_to_excel).pack(side="right", padx=5)

        self.load_reservations()

    def on_date_key(self, event):
        entry = event.widget
        formatted = format_date(entry.get())
        if formatted != entry.get() and event.keysym != "BackSpace":  # else a deleted dot would come straight back
            # shortcut: the cursor jumps to the end, so fixing a digit mid-date is clumsy; track the cursor if that annoys
            entry.delete(0, tk.END)
            entry.insert(0, formatted)

    def reset_form(self):
        """Empty the form and put it back in "new guest" mode."""
        self.editing_id = None
        for entry in (self.entry_name, self.entry_tc, self.entry_birth, self.entry_date, self.entry_extra):
            entry.delete(0, tk.END)
        self.entry_date.insert(0, datetime.now().strftime(DATE_FORMAT))
        self.payment_cb.current(0)  # do not carry an edited guest's choices over to the next one
        self.form_frame.config(text="Yeni Rezervasyon / Kayıt Ekle")
        self.btn_save.config(text="Kayıt Ekle")
        self.btn_cancel.grid_remove()
        self.update_price(None)

    def edit_reservation(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("Uyarı", "Lütfen düzenlenecek kaydı listeden seçin.")
            return
        rec_id = self.tree.item(selected[0])["values"][0]
        self.cursor.execute(
            "SELECT room_no, guest_name, tc_no, birth_year, check_in, payment_type, extra_income, price "
            "FROM reservations WHERE id=?", (rec_id,))
        room, name, tc, birth, reg_date, payment, extra, price = self.cursor.fetchone()

        self.reset_form()
        self.editing_id = rec_id
        self.room_cb.current(self.rooms.index(room))
        self.entry_name.insert(0, name)
        self.entry_tc.insert(0, tc)
        self.entry_birth.insert(0, (birth or "")[-4:])  # older records hold a full birth date
        self.entry_date.delete(0, tk.END)
        self.entry_date.insert(0, reg_date or "")
        self.payment_cb.set(payment)
        if extra:
            self.entry_extra.insert(0, f"{extra:g}")
        self.update_price(None)  # refreshes the note about other guests in the room
        self.entry_price.delete(0, tk.END)
        self.entry_price.insert(0, f"{price:g}")  # the saved price, not a recalculated one
        self.form_frame.config(text=f"Kaydı Düzenle (Oda {room})")
        self.btn_save.config(text="Değişiklikleri Kaydet")
        self.btn_cancel.grid()

    def extra_income(self):
        try:
            return float(self.entry_extra.get().replace(",", ".") or 0)
        except ValueError:
            return 0.0

    def update_price(self, event):
        room_no = self.room_cb.get().split(" ")[0]
        room_type = self.room_types.get(room_no, "Küçük Oda")
        price = self.prices[room_type]["Nakit" if self.payment_cb.get() == "Nakit" else "Diger"]

        # A room is paid for once: a guest who joins an occupied room adds only the extra income.
        # When editing, only guests registered before this one count, so the first guest keeps the room price.
        self.cursor.execute("SELECT guest_name FROM reservations WHERE room_no=? AND status='Aktif' AND (? IS NULL OR id < ?)",
                            (room_no, self.editing_id, self.editing_id))
        others = [row[0] for row in self.cursor.fetchall()]
        if others:
            price = 0
        self.lbl_price_note.config(
            text=f"Bu odada {', '.join(others)} kalıyor; oda ücreti tekrar eklenmedi." if others else "")
        price += self.extra_income()

        self.entry_price.delete(0, tk.END)
        self.entry_price.insert(0, f"{price:g}")

    def save_reservation(self):
        room = self.room_cb.get().split(" ")[0]
        name = self.entry_name.get().strip()
        tc = self.entry_tc.get()
        birth = self.entry_birth.get()
        reg_date = self.entry_date.get()
        payment = self.payment_cb.get()
        try:
            price = float(self.entry_price.get().replace(",", "."))
        except ValueError:
            price = 0.0

        if not name or not tc:
            messagebox.showerror("Hata", "Lütfen kalan adı soyadı ve TC Kimlik No alanlarını doldurun!")
            return

        error = validate(tc, birth, reg_date)
        if error:
            messagebox.showerror("Hata", error)
            return

        values = (room, name, tc, birth, reg_date, payment, self.extra_income(), price)
        if self.editing_id:
            self.cursor.execute(
                "UPDATE reservations SET room_no=?, guest_name=?, tc_no=?, birth_year=?, check_in=?, payment_type=?, extra_income=?, price=? WHERE id=?",
                values + (self.editing_id,))
        else:
            self.cursor.execute(
                "INSERT INTO reservations (room_no, guest_name, tc_no, birth_year, check_in, payment_type, extra_income, price, status) VALUES (?, ?, ?, ?, ?, ?, ?, ?, 'Aktif')",
                values)
        self.conn.commit()
        messagebox.showinfo("Başarılı", "Kayıt güncellendi!" if self.editing_id else "Rezervasyon başarıyla eklendi!")
        self.reset_form()
        self.load_reservations()
        self.refresh_dashboard()

    def load_reservations(self, event=None):
        """Fill the table with the guests currently staying that match the search box (empty box = everyone)."""
        query = self.search_entry.get().lower()
        for i in self.tree.get_children():
            self.tree.delete(i)
        self.cursor.execute(
            "SELECT id, room_no, guest_name, tc_no, birth_year, check_in, payment_type, extra_income, price "
            "FROM reservations WHERE status='Aktif'")
        for rec_id, room_no, guest_name, tc_no, birth, reg_date, payment_type, extra, price in self.cursor.fetchall():
            if query in str(room_no).lower() or query in str(guest_name).lower():
                self.tree.insert("", "end", values=(rec_id, room_no, self.room_types.get(room_no, ""), guest_name, tc_no,
                                                    (birth or "")[-4:], reg_date, payment_type,
                                                    f"{extra or 0:g} TL", f"{price:g} TL"))

    def checkout_reservation(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("Uyarı", "Lütfen çıkışı yapılacak odadan bir kaydı seçin.")
            return
        room_no = str(self.tree.item(selected[0])["values"][1])
        self.cursor.execute("SELECT COUNT(*) FROM reservations WHERE room_no=? AND status='Aktif'", (room_no,))
        count = self.cursor.fetchone()[0]

        if messagebox.askyesno("Onay", f"Oda {room_no} boşaltılacak ve odadaki {count} misafirin çıkışı yapılacak. "
                                       "Onaylıyor musunuz?"):
            # Keep the records: the cash report and the Excel export need past guests too.
            self.cursor.execute("UPDATE reservations SET status='Çıkış yaptı' WHERE room_no=? AND status='Aktif'", (room_no,))
            self.conn.commit()
            self.reset_form()  # the form may still hold one of these guests for editing
            self.load_reservations()
            self.refresh_dashboard()
            messagebox.showinfo("Bilgi", f"Çıkış yapıldı. Oda {room_no} artık boş.")

    def export_to_excel(self):
        stays = load_stays(self.conn)
        if not stays:
            messagebox.showinfo("Bilgi", "Aktarılacak kayıt yok.")
            return
        desktop = Path.home() / "Desktop"
        folder = desktop if desktop.is_dir() else Path.home()  # OneDrive can move the Desktop folder
        register = folder / f"Kayit_Defteri_{datetime.now():%Y-%m-%d}.xlsx"
        daily = folder / f"Gunluk_Alinan_{datetime.now():%Y-%m-%d}.xlsx"
        try:
            export_register(stays, register)
            export_daily(stays, daily, self.rooms)
        except PermissionError:
            messagebox.showerror("Hata", f"Dosya yazılamadı. Excel'de açıksa kapatıp tekrar deneyin:\n{register}\n{daily}")
            return
        messagebox.showinfo("Başarılı", f"İki Excel dosyası oluşturuldu:\n{register}\n{daily}")
        if hasattr(os, "startfile"):  # Windows only
            os.startfile(register)
            os.startfile(daily)

    def setup_rooms(self):
        self.rooms_container = ttk.Frame(self.tab_rooms)
        self.rooms_container.pack(fill="both", expand=True, padx=10, pady=10)
        self.setup_rooms_grid()

    def setup_rooms_grid(self):
        for widget in self.rooms_container.winfo_children():
            widget.destroy()

        floors = {
            "1. Kat (101-105)": ["101", "102", "103", "104", "105"],
            "2. Kat (201-205)": ["201", "202", "203", "204", "205"],
            "3. Kat (301-305)": ["301", "302", "303", "304", "305"],
            "4. Kat (401-404)": ["401", "402", "403", "404"]
        }

        self.cursor.execute("SELECT room_no, guest_name FROM reservations WHERE status='Aktif'")
        occupied_rooms = {}
        for room_no, guest_name in self.cursor.fetchall():
            occupied_rooms.setdefault(room_no, []).append(guest_name)

        row_idx = 0
        for floor_name, r_list in floors.items():
            lbl_floor = ttk.Label(self.rooms_container, text=floor_name, font=("Arial", 12, "bold"))
            lbl_floor.grid(row=row_idx, column=0, sticky="w", pady=(10, 5))
            row_idx += 1

            f_frame = ttk.Frame(self.rooms_container)
            f_frame.grid(row=row_idx, column=0, sticky="ew", pady=(0, 10))
            row_idx += 1

            for c_idx, room in enumerate(r_list):
                r_type = self.room_types[room]
                if room in occupied_rooms:
                    txt = f"Oda {room} ({r_type})\n[Dolu]\n" + "\n".join(occupied_rooms[room])
                    bg = "#ffcdd2"
                else:
                    txt = f"Oda {room} ({r_type})\n[Boş]"
                    bg = "#c8e6c9"

                lbl_room = tk.Label(f_frame, text=txt, relief="solid", bd=1, padx=5, pady=5, font=("Arial", 9), bg=bg,
                                    width=18, height=6)
                lbl_room.grid(row=0, column=c_idx, padx=4, pady=4)

    def setup_finance(self):
        finance_frame = ttk.Frame(self.tab_finance, padding=20)
        finance_frame.pack(fill="both", expand=True)

        ttk.Label(finance_frame, text="📊 İzmit My House Butik - Kasa ve Finansal Durum Raporu",
                  font=("Arial", 16, "bold")).pack(pady=10)

        period_frame = ttk.Frame(finance_frame)
        period_frame.pack(anchor="w", pady=10)
        ttk.Label(period_frame, text="Dönem:", font=("Arial", 12)).pack(side="left")
        self.finance_period = ttk.Combobox(period_frame, values=["Bugün", "Bu Ay", "Tümü"], width=10, state="readonly")
        self.finance_period.pack(side="left", padx=8)
        self.finance_period.current(0)
        self.finance_period.bind("<<ComboboxSelected>>", lambda event: self.update_finance_report())
        ttk.Label(period_frame, text="(ücret, kayıt tarihine sayılır)").pack(side="left")

        self.lbl_nakit = ttk.Label(finance_frame, text="Nakit Ödemeler Toplamı: 0 TL", font=("Arial", 12))
        self.lbl_nakit.pack(anchor="w", pady=10)

        self.lbl_kart = ttk.Label(finance_frame, text="Fiş Ödemeleri Toplamı: 0 TL", font=("Arial", 12))
        self.lbl_kart.pack(anchor="w", pady=10)

        self.lbl_eft = ttk.Label(finance_frame, text="EFT Ödemeleri Toplamı: 0 TL", font=("Arial", 12))
        self.lbl_eft.pack(anchor="w", pady=10)

        ttk.Separator(finance_frame, orient="horizontal").pack(fill="x", pady=15)

        self.lbl_genel_toplam = ttk.Label(finance_frame, text="Genel Kasa Toplamı: 0 TL", font=("Arial", 14, "bold"),
                                          foreground="green")
        self.lbl_genel_toplam.pack(anchor="w", pady=10)

        btn_fin_refresh = ttk.Button(finance_frame, text="Raporu Güncelle", command=self.update_finance_report)
        btn_fin_refresh.pack(pady=20)

    def update_finance_report(self):
        now = datetime.now()
        # check_in is stored as GG.AA.YYYY text, so a period is a LIKE pattern on it.
        pattern = {"Bugün": now.strftime(DATE_FORMAT), "Bu Ay": now.strftime("%%.%m.%Y"), "Tümü": "%"}[self.finance_period.get()]
        self.cursor.execute(
            "SELECT payment_type, SUM(price) FROM reservations WHERE COALESCE(check_in, '') LIKE ? GROUP BY payment_type",
            (pattern,))
        totals = {"Nakit": 0.0, "Fiş": 0.0, "EFT": 0.0}
        for row in self.cursor.fetchall():
            p_type, p_sum = row
            if p_type in totals:
                totals[p_type] = p_sum if p_sum else 0.0

        nakit = totals["Nakit"]
        kart = totals["Fiş"]
        eft = totals["EFT"]
        genel = nakit + kart + eft

        self.lbl_nakit.config(text=f"Nakit Ödemeler Toplamı: {nakit:,.2f} TL")
        self.lbl_kart.config(text=f"Fiş Ödemeleri Toplamı: {kart:,.2f} TL")
        self.lbl_eft.config(text=f"EFT Ödemeleri Toplamı: {eft:,.2f} TL")
        self.lbl_genel_toplam.config(text=f"Genel Kasa Toplamı: {genel:,.2f} TL")


def show_login(root, on_success):
    """Ask for the password inside the main window; on the first run, ask to choose one."""
    first_run = not PASSWORD_FILE.exists()
    frame = ttk.Frame(root, padding=30)
    frame.pack(expand=True)
    ttk.Label(frame, text="İzmit My House Butik", font=("Arial", 14, "bold")).pack(pady=(0, 15))
    ttk.Label(frame, text="Yeni bir parola belirleyin (en az 4 karakter):" if first_run else "Parola:").pack(anchor="w")
    entry = ttk.Entry(frame, show="*", width=32)
    entry.pack(pady=5)
    entry.focus_set()
    if first_run:
        ttk.Label(frame, text="Parolayı tekrar yazın:").pack(anchor="w")
        repeat = ttk.Entry(frame, show="*", width=32)
        repeat.pack(pady=5)
        ttk.Label(frame, text="Bu parola programın her açılışında sorulacak.").pack()
    message = ttk.Label(frame, foreground="red")
    message.pack()

    def submit(event=None):
        password = entry.get()
        if first_run:
            if len(password) < 4:
                message.config(text="Parola en az 4 karakter olmalıdır.")
                return
            if password != repeat.get():
                message.config(text="İki parola aynı değil.")
                return
            PASSWORD_FILE.write_text(make_password_record(password))
        elif not check_password(password, PASSWORD_FILE.read_text()):
            message.config(text="Parola yanlış.")
            entry.delete(0, tk.END)
            return
        root.unbind("<Return>")
        frame.destroy()
        on_success()

    ttk.Button(frame, text="Kaydet ve Aç" if first_run else "Giriş", command=submit).pack(pady=10)
    root.bind("<Return>", submit)


if __name__ == "__main__":
    root = tk.Tk()
    root.title("İzmit My House Butik - Otel Yönetim Sistemi")
    root.geometry("460x320")
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    show_login(root, lambda: HotelApp(root))
    root.mainloop()
