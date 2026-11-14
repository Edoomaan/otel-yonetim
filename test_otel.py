"""Self-check for the logic that does not need a window. Run: python test_otel.py"""
import sqlite3
import tempfile
from datetime import datetime, timedelta
from pathlib import Path

from openpyxl import load_workbook

from otel import (backup_database, check_password, export_daily, export_register, format_date, is_newer, is_valid_tc,
                  load_stays, make_password_record, validate)

TC = "10000000146"  # 10th digit: (1+0+0+0+1)*7 - 0 = 14 -> 4; 11th: digit sum 6 -> 6
assert is_valid_tc(TC)
assert not is_valid_tc("10000000147")           # wrong 11th digit
assert not is_valid_tc("10000000156")           # wrong 10th digit
assert not is_valid_tc("00000000000")           # cannot start with 0
assert not is_valid_tc("1000000014")            # 10 digits
assert not is_valid_tc("1000000014a")           # not a number

assert validate(TC, "", "25.11.2026") is None             # birth year is optional
assert validate(TC, "1990", "25.11.2026") is None
assert validate("12345678901", "", "25.11.2026")          # fails the checksum
assert validate(TC, "90", "25.11.2026")                   # birth year must be a full year
assert validate(TC, "3000", "25.11.2026")                 # nobody is born in the future
assert validate(TC, "", "2026-11-25")                     # wrong date format
assert validate(TC, "", "31.02.2026")                     # impossible date

assert is_newer("v3", "2") and is_newer("v10", "9")           # numbers, not text: "10" > "9"
assert not is_newer("v2", "2") and not is_newer("v1", "2") and not is_newer("latest", "2")

assert [format_date(x) for x in ("1", "12", "123", "1203", "12031990", "120319901", "12.03.1990")] == \
       ["1", "12.", "12.3", "12.03.", "12.03.1990", "12.03.1990", "12.03.1990"]

# --- Excel exports: two guests share room 101, a third room pays by card, one record is in another month
conn = sqlite3.connect(":memory:")
conn.execute("CREATE TABLE reservations (id INTEGER PRIMARY KEY, room_no, guest_name, tc_no, birth_year, "
             "check_in, payment_type, price REAL, status)")
conn.executemany("INSERT INTO reservations (room_no, guest_name, tc_no, birth_year, check_in, payment_type, price, status) "
                 "VALUES (?, ?, ?, ?, ?, ?, ?, ?)", [
    ("101", "Ali Veli", TC, "1990", "03.11.2026", "Nakit", 1800, "Çıkış yaptı"),
    ("101", "=1+1", "11111111110", "05.03.1985", "03.11.2026", "EFT", 500, "Çıkış yaptı"),   # old full birth date
    ("404", "Ayşe Öz", TC, "", "03.11.2026", "Kredi Kartı", 2150, "Aktif"),
    ("205", "Can Er", TC, "2001", "04.11.2026", "Nakit", 1700, "Aktif"),
    ("302", "Nisan Misafiri", TC, "1970", "30.04.2026", "Nakit", 1800, "Çıkış yaptı"),
    ("303", "Tarihsiz", TC, "1970", None, "Nakit", 999, "Aktif"),                           # no date: left out
])
stays = load_stays(conn)
assert sorted(stays) == [(2026, 4), (2026, 11)] and sorted(stays[(2026, 11)]) == [3, 4]

with tempfile.TemporaryDirectory() as tmp:
    register, daily = Path(tmp) / "register.xlsx", Path(tmp) / "daily.xlsx"
    export_register(stays, register)
    export_daily(stays, daily, ["101", "205", "302", "404"])
    reg_book, daily_book = load_workbook(register), load_workbook(daily)

assert reg_book.sheetnames == ["Nisan 2026", "Kasım 2026"]          # months in calendar order
ws = reg_book["Kasım 2026"]
cell = lambda ref: ws[ref].value
assert cell("B2") == datetime(2026, 11, 3) and (cell("C3"), cell("H3"), cell("O4")) == ("ODA NO", "ADI SOYADI", "FİŞ")
# room 101: both guests under one room number, money on the first line, split by payment type
assert (cell("B5"), cell("C5"), cell("F5"), cell("H5"), cell("K5")) == (1, 101, TC, "Ali Veli", 1990)
assert (cell("N5"), cell("O5"), cell("P5"), cell("Q5")) == (1800, None, 500, 2300)
assert (cell("C6"), cell("H6"), cell("K6"), cell("N6")) == (None, "=1+1", 1985, None)
assert ws["H6"].data_type == "s"                                    # a name starting with "=" stays text, not a formula
# room 404 starts two lines later even though it has one guest; then the day's totals
assert (cell("B7"), cell("C7"), cell("H7"), cell("O7"), cell("Q7")) == (2, 404, "Ayşe Öz", 2150, 2150)
assert (cell("M9"), cell("N9"), cell("O9"), cell("P9"), cell("Q9")) == ("TOPLAM", 1800, 2150, 500, 4450)
assert cell("B12") == datetime(2026, 11, 4) and (cell("C15"), cell("N15")) == (205, 1700)   # next day's block

assert daily_book.sheetnames == ["NİSAN 2026", "KASIM 2026"]        # Turkish upper case: NİSAN, not NISAN
ws = daily_book["KASIM 2026"]
assert [c.value for c in ws[1]] == ["GÜNLER", 101, 205, 302, 404, "TOPLAM", "GÜNLER", "S.O.S"]
assert [c.value for c in ws[4]] == [3, 2300, None, None, 2150, 4450, "Salı", 2]          # 3 Nov 2026 is a Tuesday
assert [c.value for c in ws[5]] == [4, None, 1700, None, None, 1700, "Çarşamba", 1]
assert [c.value for c in ws[2]] == [1, None, None, None, None, 0, "Pazar", None]         # an empty day
assert [c.value for c in ws[33]] == ["TOPLAM", 2300, 1700, 0, 2150, 6150, 205, 3]        # 6150 / 30 days = 205
assert ws["A31"].value == 30 and ws["A32"].value is None                                  # November has 30 days

# --- daily backup
with tempfile.TemporaryDirectory() as tmp:
    db, backups, day = Path(tmp) / "hotel.db", Path(tmp) / "yedekler", datetime(2026, 11, 11)
    backup_database(db, backups, day)                        # no database yet: nothing to do
    assert not backups.exists()
    db.write_text("first")
    backup_database(db, backups, day)
    db.write_text("changed later the same day")
    backup_database(db, backups, day)                        # one copy per day; the morning copy is kept
    assert [(f.name, f.read_text()) for f in backups.iterdir()] == [("hotel_2026-11-11.db", "first")]
    for i in range(1, 5):
        backup_database(db, backups, day + timedelta(days=i), keep=3)
    assert sorted(f.name for f in backups.iterdir()) == \
           ["hotel_2026-11-13.db", "hotel_2026-11-14.db", "hotel_2026-11-15.db"]   # oldest ones pruned

# --- startup password
record = make_password_record("gizli1234")
assert "gizli1234" not in record                              # the password itself is not stored
assert check_password("gizli1234", record) and not check_password("gizli123", record)
assert make_password_record("gizli1234") != record            # a new random salt every time
assert not check_password("gizli1234", "damaged file")
print("ok")
