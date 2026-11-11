"""Self-check for the logic that does not need a window. Run: python test_otel.py"""
import sqlite3
import tempfile
from pathlib import Path

from openpyxl import load_workbook

from otel import export_excel, validate

assert validate("12345678901", "25.11.2026") is None
assert validate("123", "25.11.2026")            # too short
assert validate("1234567890a", "25.11.2026")    # not a number
assert validate("12345678901", "2026-11-25")    # wrong date format
assert validate("12345678901", "31.02.2026")    # impossible date

conn = sqlite3.connect(":memory:")
conn.execute("CREATE TABLE reservations (id INTEGER PRIMARY KEY, room_no, guest_name, tc_no, birth_year, "
             "check_in, check_out, payment_type, extra_person, price REAL, status)")
conn.execute("INSERT INTO reservations VALUES (1,'101','Ali Veli','12345678901','1990','10.11.2026','11.11.2026',"
             "'Nakit','Yok',1800,'Çıkış yaptı')")
conn.execute("INSERT INTO reservations VALUES (2,'404','Ayşe Öz','10987654321','1985','11.11.2026','12.11.2026',"
             "'EFT','Var',2650,'Aktif')")
with tempfile.TemporaryDirectory() as tmp:
    path = Path(tmp) / "out.xlsx"
    export_excel(conn, path, {"101": "Büyük Oda", "404": "Suit Oda"})
    rows = list(load_workbook(path).active.values)
assert len(rows) == 3, rows                      # header + checked-out guest + active guest
assert rows[1][:4] == ("101", "Büyük Oda", "Ali Veli", "12345678901"), rows[1]  # TC stays text
assert rows[2][-2:] == (2650, "Aktif"), rows[2]
print("ok")
