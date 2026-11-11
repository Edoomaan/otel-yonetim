"""Self-check for the logic that does not need a window. Run: python test_otel.py"""
import sqlite3
import tempfile
from datetime import datetime, timedelta
from pathlib import Path

from openpyxl import load_workbook

from otel import backup_database, check_password, count_nights, export_excel, format_date, is_newer, is_valid_tc, make_password_record, validate

TC = "10000000146"  # 10th digit: (1+0+0+0+1)*7 - 0 = 14 -> 4; 11th: digit sum 6 -> 6
assert is_valid_tc(TC)
assert not is_valid_tc("10000000147")           # wrong 11th digit
assert not is_valid_tc("10000000156")           # wrong 10th digit
assert not is_valid_tc("00000000000")           # cannot start with 0
assert not is_valid_tc("1000000014")            # 10 digits
assert not is_valid_tc("1000000014a")           # not a number

IN = "24.11.2026"
assert validate(TC, "", "25.11.2026", IN) is None             # birth date is optional
assert validate(TC, "05.03.1990", "25.11.2026", IN) is None
assert validate(TC, "", IN, IN) is None                       # same-day check-out is allowed
assert validate("12345678901", "", "25.11.2026", IN)          # fails the checksum
assert validate(TC, "1990", "25.11.2026", IN)                 # birth date must be a full date
assert validate(TC, "", "2026-11-25", IN)                     # wrong date format
assert validate(TC, "", "31.02.2026", IN)                     # impossible date
assert validate(TC, "", "23.11.2026", IN)                     # check-out before check-in

assert count_nights("30.11.2026", "02.12.2026") == 2          # across a month end
assert count_nights(IN, IN) == 1                              # same day still costs one night
assert count_nights(IN, "25.1") == 1                          # date still being typed

assert is_newer("v3", "2") and is_newer("v10", "9")           # numbers, not text: "10" > "9"
assert not is_newer("v2", "2") and not is_newer("v1", "2") and not is_newer("latest", "2")

assert [format_date(x) for x in ("1", "12", "123", "1203", "12031990", "120319901", "12.03.1990")] == \
       ["1", "12.", "12.3", "12.03.", "12.03.1990", "12.03.1990", "12.03.1990"]

conn = sqlite3.connect(":memory:")
conn.execute("CREATE TABLE reservations (id INTEGER PRIMARY KEY, room_no, guest_name, tc_no, birth_year, "
             "check_in, check_out, payment_type, extra_person, price REAL, status)")
conn.execute("INSERT INTO reservations VALUES (1,'101','Ali Veli','12345678901','1990','10.11.2026','11.11.2026',"
             "'Nakit','Yok',1800,'Çıkış yaptı')")
conn.execute("INSERT INTO reservations VALUES (2,'404','=1+1','10987654321','1985','11.11.2026','12.11.2026',"
             "'EFT','Var',2650,'Aktif')")
with tempfile.TemporaryDirectory() as tmp:
    path = Path(tmp) / "out.xlsx"
    export_excel(conn, path, {"101": "Büyük Oda", "404": "Suit Oda"})
    sheet = load_workbook(path).active
    rows = list(sheet.values)
assert (sheet["C3"].value, sheet["C3"].data_type) == ("=1+1", "s")   # a name starting with "=" stays text, not a formula
assert len(rows) == 3, rows                      # header + checked-out guest + active guest
assert rows[1][:4] == ("101", "Büyük Oda", "Ali Veli", "12345678901"), rows[1]  # TC stays text
assert rows[2][-2:] == (2650, "Aktif"), rows[2]
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
record = make_password_record("gizli1234")
assert "gizli1234" not in record                              # the password itself is not stored
assert check_password("gizli1234", record) and not check_password("gizli123", record)
assert make_password_record("gizli1234") != record            # a new random salt every time
assert not check_password("gizli1234", "damaged file")
print("ok")
