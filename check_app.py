import json
from urllib.request import urlopen

for path in ("/api/summary", "/api/reports?q=F07", "/api/reports?shift_type=Day", "/api/reports/3"):
    with urlopen("http://127.0.0.1:8765" + path, timeout=5) as response:
        body = response.read().decode("utf-8")
        print(path, response.status, body[:1000].replace("\\n", " "))
