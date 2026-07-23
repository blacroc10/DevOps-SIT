from datetime import datetime, timezone
from pathlib import Path

from flask import Flask

app = Flask(__name__)
DATA_DIR = Path("/app/data")
VISITS_FILE = DATA_DIR / "visits.log"


@app.get("/")
def index():
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now(timezone.utc).isoformat()
    with VISITS_FILE.open("a", encoding="utf-8") as log:
        log.write(f"Visited at {timestamp}\n")

    visits = VISITS_FILE.read_text(encoding="utf-8").splitlines()
    return (
        "<!doctype html><html><head><title>DevOps Docker Demo</title>"
        "<style>body{margin:0;background:#f4f7fb;color:#172033;font-family:Segoe UI,"
        "sans-serif;display:grid;place-items:center;min-height:100vh}.card{background:white;"
        "padding:48px 56px;border-radius:18px;box-shadow:0 16px 50px #25385822;"
        "border-top:8px solid #2496ed}h1{margin-top:0;color:#0d5f9b}</style></head>"
        "<body><main class='card'><h1>Hello, Docker!</h1>"
        "<p>DevOps Lab application by Shubhankar Sarangi.</p>"
        f"<p><strong>Persistent visit count:</strong> {len(visits)}</p>"
        "</main></body></html>"
    )


@app.get("/health")
def health():
    return {"status": "healthy"}


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
