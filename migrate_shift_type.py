import sqlite3
from pathlib import Path

db = Path(__file__).parent / "data" / "ctep_archive.db"
conn = sqlite3.connect(db)
cols = {row[1] for row in conn.execute("PRAGMA table_info(reports)")}
if "shift_type" not in cols:
    conn.execute("ALTER TABLE reports ADD COLUMN shift_type TEXT")
conn.execute("""
UPDATE reports
SET shift_type = CASE
    WHEN doc_no LIKE '%D' THEN 'Day'
    WHEN doc_no LIKE '%N' THEN 'Night'
    ELSE ''
END
""")
conn.commit()
print(conn.execute("SELECT doc_no, report_date, shift_type, shift FROM reports ORDER BY report_date").fetchall())
conn.close()
