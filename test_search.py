import sqlite3
from pathlib import Path

db = Path(__file__).parent / "data" / "ctep_archive.db"
conn = sqlite3.connect(db)
for term in ("F07", "GTG", "PSV"):
    rows = conn.execute(
        "SELECT doc_no, report_date, shift FROM report_search WHERE report_search MATCH ?",
        (term,)
    ).fetchall()
    print(term, rows)
print("REPORT COUNT", conn.execute("SELECT COUNT(*) FROM reports").fetchone()[0])
