import os

import redis
from flask import Flask

app = Flask(__name__)
cache = redis.Redis(host=os.getenv("REDIS_HOST", "redis"), port=6379)


@app.get("/")
def index():
    count = cache.incr("hits")
    return (
        "<!doctype html><html><head><title>Docker Compose Demo</title>"
        "<style>body{margin:0;background:#eef8f3;color:#19352a;font-family:Segoe UI,"
        "sans-serif;display:grid;place-items:center;min-height:100vh}.card{background:white;"
        "padding:48px 56px;border-radius:18px;box-shadow:0 16px 50px #19352a22;"
        "border-top:8px solid #d82c20}h1{margin-top:0;color:#9f1d16}</style></head>"
        "<body><main class='card'><h1>Docker Compose Demo</h1>"
        f"<p>This multi-container application has been visited <strong>{count}</strong> times.</p>"
        "<p>Flask web service connected to Redis.</p></main></body></html>"
    )


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
