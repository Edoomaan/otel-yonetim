import tkinter as tk
from tkinter import ttk, messagebox
import os
import sqlite3
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook

# Fixed per-user folder, so the data survives moving the .exe around.
DATA_DIR = Path(os.environ.get("APPDATA", Path.home())) / "OtelYonetim"

EXPORT_HEADERS = ["Oda No", "Oda Tipi", "Ad Soyad", "TC No", "Doğum Tarihi", "Giriş", "Çıkış", "Ödeme", "3. Kişi", "Ücret", "Durum"]


def is_valid_tc(tc):
    """Check the two control digits of a Turkish ID number."""
    if not (tc.isdigit() and len(tc) == 11) or tc[0] == "0":
        return False
    d = [int(c) for c in tc]
    # d[0:9:2] = 1st, 3rd, 5th, 7th, 9th digits; d[1:8:2] = 2nd, 4th, 6th, 8th
    tenth = (sum(d[0:9:2]) * 7 - sum(d[1:8:2])) % 10
    eleventh = sum(d[:10]) % 10
    return d[9] == tenth and d[10] == eleventh


def is_valid_date(text):
    try:
        datetime.strptime(text, "%d.%m.%Y")
        return True
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


def validate(tc, birth, check_out):
    """Return an error message, or None when the input is fine."""
    if not is_valid_tc(tc):
        return "Geçersiz TC Kimlik No! 11 haneyi kontrol edip tekrar girin."
    if birth and not is_valid_date(birth):
        return "Doğum tarihi GG.AA.YYYY biçiminde olmalıdır! (örnek: 05.03.1990)"
    if not is_valid_date(check_out):
        return "Çıkış tarihi GG.AA.YYYY biçiminde olmalıdır! (örnek: 25.11.2026)"
    return None


def export_excel(conn, path, room_types):
    """Write every reservation (active and checked out) to an .xlsx file."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Kayıtlar"
    ws.append(EXPORT_HEADERS)
    rows = conn.execute(
        "SELECT room_no, guest_name, tc_no, birth_year, check_in, check_out, payment_type, extra_person, price, status "
        "FROM reservations ORDER BY id")
    for row in rows:
        ws.append([row[0], room_types.get(row[0], "")] + list(row[1:]))
    for col in ws.columns:
        ws.column_dimensions[col[0].column_letter].width = 18
    wb.save(path)


class HotelApp:
    def __init__(self, root):
        self.root = root
        self.root.title("İzmit My House Butik - Otel Yönetim Sistemi")
        self.root.geometry("1300x780")

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

        # Veritabanı bağlantısı
        DATA_DIR.mkdir(parents=True, exist_ok=True)
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

    def refresh_dashboard(self):
        self.cursor.execute("SELECT COUNT(*) FROM reservations WHERE status='Aktif'")
        count = self.cursor.fetchone()[0]
        self.lbl_total_res.config(text=f"Toplam Aktif Konaklayan Misafir: {count}")
        self.setup_rooms_grid()
        self.update_finance_report()

    def setup_reservations(self):
        left_frame = ttk.LabelFrame(self.tab_reservations, text="Yeni Rezervasyon / Kayıt Ekle")
        left_frame.pack(side="left", fill="y", padx=10, pady=10, ipadx=10, ipady=10)

        room_display_list = [f"{room} ({self.room_types[room]})" for room in self.rooms]

        ttk.Label(left_frame, text="Oda No:").grid(row=0, column=0, sticky="w", pady=4)
        self.room_cb = ttk.Combobox(left_frame, values=room_display_list, width=22)
        self.room_cb.grid(row=0, column=1, pady=4)
        self.room_cb.current(0)
        self.room_cb.bind("<<ComboboxSelected>>", self.update_price)

        ttk.Label(left_frame, text="Kalan Adı Soyadı:").grid(row=1, column=0, sticky="w", pady=4)
        self.entry_name = ttk.Entry(left_frame)
        self.entry_name.grid(row=1, column=1, pady=4)

        ttk.Label(left_frame, text="TC Kimlik No:").grid(row=2, column=0, sticky="w", pady=4)
        # Reject the keystroke unless the field would still be at most 11 digits (%P = text after the edit).
        only_tc_digits = (self.root.register(lambda text: len(text) <= 11 and (text.isdigit() or text == "")), "%P")
        self.entry_tc = ttk.Entry(left_frame, validate="key", validatecommand=only_tc_digits)
        self.entry_tc.grid(row=2, column=1, pady=4)

        ttk.Label(left_frame, text="Doğum Tarihi (GG.AA.YYYY):").grid(row=3, column=0, sticky="w", pady=4)
        self.entry_birth = ttk.Entry(left_frame)
        self.entry_birth.grid(row=3, column=1, pady=4)
        self.entry_birth.bind("<KeyRelease>", self.on_date_key)

        ttk.Label(left_frame, text="Çıkış Tarihi (GG.AA.YYYY):").grid(row=4, column=0, sticky="w", pady=4)
        self.entry_out = ttk.Entry(left_frame)
        self.entry_out.insert(0, datetime.now().strftime("%d.%m.%Y"))
        self.entry_out.grid(row=4, column=1, pady=4)
        self.entry_out.bind("<KeyRelease>", self.on_date_key)

        ttk.Label(left_frame, text="Ödeme Şekli:").grid(row=5, column=0, sticky="w", pady=4)
        self.payment_cb = ttk.Combobox(left_frame, values=["Nakit", "Kredi Kartı", "EFT"], width=22, state="readonly")
        self.payment_cb.grid(row=5, column=1, pady=4)
        self.payment_cb.current(0)
        self.payment_cb.bind("<<ComboboxSelected>>", self.update_price)

        ttk.Label(left_frame, text="3. Kişi (+500 TL):").grid(row=6, column=0, sticky="w", pady=4)
        self.extra_person_cb = ttk.Combobox(left_frame, values=["Yok", "Var"], width=22, state="readonly")
        self.extra_person_cb.grid(row=6, column=1, pady=4)
        self.extra_person_cb.current(0)
        self.extra_person_cb.bind("<<ComboboxSelected>>", self.update_price)

        ttk.Label(left_frame, text="Toplam Ücret (TL):").grid(row=7, column=0, sticky="w", pady=4)
        self.entry_price = ttk.Entry(left_frame)
        self.entry_price.grid(row=7, column=1, pady=4)

        self.update_price(None)

        btn_add = ttk.Button(left_frame, text="Kayıt Ekle", command=self.add_reservation)
        btn_add.grid(row=8, column=0, columnspan=2, pady=12)

        right_frame = ttk.Frame(self.tab_reservations)
        right_frame.pack(side="right", fill="both", expand=True, padx=10, pady=10)

        search_frame = ttk.Frame(right_frame)
        search_frame.pack(fill="x", pady=5)
        ttk.Label(search_frame, text="🔍 Müşteri Adı veya Oda No Ara:").pack(side="left", padx=5)
        self.search_entry = ttk.Entry(search_frame, width=30)
        self.search_entry.pack(side="left", padx=5)
        self.search_entry.bind("<KeyRelease>", self.filter_reservations)

        cols = ("ID", "Oda No", "Oda Tipi", "Kalan Adı", "TC No", "Doğum", "Çıkış", "Ödeme", "3. Kişi", "Ücret")
        self.tree = ttk.Treeview(right_frame, columns=cols, show="headings")
        for col in cols:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=80)
        self.tree.pack(fill="both", expand=True, pady=5)

        btn_action_frame = ttk.Frame(right_frame)
        btn_action_frame.pack(fill="x", pady=5)
        btn_checkout = ttk.Button(btn_action_frame, text="🚪 Seçili Misafirin Çıkışını Yap (Odayı Boşalt)",
                                  command=self.checkout_reservation)
        btn_checkout.pack(side="left", padx=5)
        ttk.Button(btn_action_frame, text="📊 Tüm Kayıtları Excel'e Aktar",
                   command=self.export_to_excel).pack(side="right", padx=5)

        self.load_reservations()

    def on_date_key(self, event):
        if event.keysym == "BackSpace":  # otherwise a deleted dot would come straight back
            return
        entry = event.widget
        formatted = format_date(entry.get())
        if formatted != entry.get():
            # shortcut: the cursor jumps to the end, so fixing a digit mid-date is clumsy; track the cursor if that annoys
            entry.delete(0, tk.END)
            entry.insert(0, formatted)

    def update_price(self, event):
        selected_text = self.room_cb.get()
        room_no = selected_text.split(" ")[0]
        room_type = self.room_types.get(room_no, "Küçük Oda")
        payment = self.payment_cb.get()
        extra = self.extra_person_cb.get()

        if payment == "Nakit":
            price = self.prices[room_type]["Nakit"]
        else:
            price = self.prices[room_type]["Diger"]

        if extra == "Var":
            price += 500

        self.entry_price.delete(0, tk.END)
        self.entry_price.insert(0, str(price))

    def add_reservation(self):
        selected_text = self.room_cb.get()
        room = selected_text.split(" ")[0]
        name = self.entry_name.get()
        tc = self.entry_tc.get()
        birth = self.entry_birth.get()
        cout = self.entry_out.get()
        payment = self.payment_cb.get()
        extra = self.extra_person_cb.get()
        try:
            price = float(self.entry_price.get())
        except ValueError:
            price = 0.0

        if not name or not tc:
            messagebox.showerror("Hata", "Lütfen kalan adı soyadı ve TC Kimlik No alanlarını doldurun!")
            return

        error = validate(tc, birth, cout)
        if error:
            messagebox.showerror("Hata", error)
            return

        self.cursor.execute("SELECT * FROM reservations WHERE room_no=? AND status='Aktif'", (room,))
        if self.cursor.fetchone():
            messagebox.showerror("Hata", f"Oda {room} zaten dolu! Lütfen önce o odanın çıkışını yapın.")
            return

        self.cursor.execute(
            "INSERT INTO reservations (room_no, guest_name, tc_no, birth_year, check_in, check_out, payment_type, extra_person, price, status) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
            (room, name, tc, birth, datetime.now().strftime("%d.%m.%Y"), cout, payment, extra, price, "Aktif"))
        self.conn.commit()
        messagebox.showinfo("Başarılı", "Rezervasyon başarıyla eklendi!")
        self.load_reservations()
        self.refresh_dashboard()

    def load_reservations(self):
        for i in self.tree.get_children():
            self.tree.delete(i)
        self.cursor.execute(
            "SELECT id, room_no, guest_name, tc_no, birth_year, check_out, payment_type, extra_person, price FROM reservations WHERE status='Aktif'")
        for row in self.cursor.fetchall():
            rec_id, room_no, guest_name, tc_no, birth_year, check_out, payment_type, extra_person, price = row
            room_type = self.room_types.get(room_no, "")
            self.tree.insert("", "end",
                             values=(rec_id, room_no, room_type, guest_name, tc_no, birth_year, check_out, payment_type,
                                     extra_person, f"{price} TL"))

    def filter_reservations(self, event):
        query = self.search_entry.get().lower()
        for i in self.tree.get_children():
            self.tree.delete(i)
        self.cursor.execute(
            "SELECT id, room_no, guest_name, tc_no, birth_year, check_out, payment_type, extra_person, price FROM reservations WHERE status='Aktif'")
        for row in self.cursor.fetchall():
            rec_id, room_no, guest_name, tc_no, birth_year, check_out, payment_type, extra_person, price = row
            room_type = self.room_types.get(room_no, "")
            if query in str(room_no).lower() or query in str(guest_name).lower():
                self.tree.insert("", "end",
                                 values=(rec_id, room_no, room_type, guest_name, tc_no, birth_year, check_out,
                                         payment_type, extra_person, f"{price} TL"))

    def checkout_reservation(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("Uyarı", "Lütfen çıkış yapılıp odası boşaltılacak kaydı seçin.")
            return
        item = self.tree.item(selected)
        rec_id = item['values'][0]
        room_no = item['values'][1]

        confirm = messagebox.askyesno("Onay",
                                      f"Oda {room_no} numaralı misafirin çıkışı yapılacak ve oda boşaltılacak. Onaylıyor musunuz?")
        if confirm:
            # Keep the record: the cash report and the Excel export need past guests too.
            self.cursor.execute("UPDATE reservations SET status='Çıkış yaptı' WHERE id=?", (rec_id,))
            self.conn.commit()
            self.load_reservations()
            self.refresh_dashboard()
            messagebox.showinfo("Bilgi", f"Misafir çıkışı yapıldı. Oda {room_no} artık boş.")

    def export_to_excel(self):
        desktop = Path.home() / "Desktop"
        folder = desktop if desktop.is_dir() else Path.home()  # OneDrive can move the Desktop folder
        path = folder / f"Otel_Kayitlari_{datetime.now():%Y-%m-%d}.xlsx"
        try:
            export_excel(self.conn, path, self.room_types)
        except PermissionError:
            messagebox.showerror("Hata", f"Dosya yazılamadı. Excel'de açıksa kapatıp tekrar deneyin:\n{path}")
            return
        messagebox.showinfo("Başarılı", f"Excel dosyası oluşturuldu:\n{path}")
        if hasattr(os, "startfile"):  # Windows only
            os.startfile(path)

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
        occupied_rooms = {row[0]: row[1] for row in self.cursor.fetchall()}

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
                    txt = f"Oda {room} ({r_type})\n[Dolu]\n{occupied_rooms[room]}"
                    bg = "#ffcdd2"
                else:
                    txt = f"Oda {room} ({r_type})\n[Boş]"
                    bg = "#c8e6c9"

                lbl_room = tk.Label(f_frame, text=txt, relief="solid", bd=1, padx=5, pady=5, font=("Arial", 9), bg=bg,
                                    width=18, height=4)
                lbl_room.grid(row=0, column=c_idx, padx=4, pady=4)

    def setup_finance(self):
        finance_frame = ttk.Frame(self.tab_finance, padding=20)
        finance_frame.pack(fill="both", expand=True)

        ttk.Label(finance_frame, text="📊 İzmit My House Butik - Kasa ve Finansal Durum Raporu",
                  font=("Arial", 16, "bold")).pack(pady=10)

        self.lbl_nakit = ttk.Label(finance_frame, text="Nakit Ödemeler Toplamı: 0 TL", font=("Arial", 12))
        self.lbl_nakit.pack(anchor="w", pady=10)

        self.lbl_kart = ttk.Label(finance_frame, text="Kredi Kartı Ödemeleri Toplamı: 0 TL", font=("Arial", 12))
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
        self.cursor.execute("SELECT payment_type, SUM(price) FROM reservations GROUP BY payment_type")
        totals = {"Nakit": 0.0, "Kredi Kartı": 0.0, "EFT": 0.0}
        for row in self.cursor.fetchall():
            p_type, p_sum = row
            if p_type in totals:
                totals[p_type] = p_sum if p_sum else 0.0

        nakit = totals["Nakit"]
        kart = totals["Kredi Kartı"]
        eft = totals["EFT"]
        genel = nakit + kart + eft

        self.lbl_nakit.config(text=f"Nakit Ödemeler Toplamı: {nakit:,.2f} TL")
        self.lbl_kart.config(text=f"Kredi Kartı Ödemeleri Toplamı: {kart:,.2f} TL")
        self.lbl_eft.config(text=f"EFT Ödemeleri Toplamı: {eft:,.2f} TL")
        self.lbl_genel_toplam.config(text=f"Genel Kasa Toplamı: {genel:,.2f} TL")


if __name__ == "__main__":
    root = tk.Tk()
    app = HotelApp(root)
    root.mainloop()
