"""Entry point for running the Flask app.

Usage:
    python run.py             # serve on http://0.0.0.0:5000
    PORT=8080 python run.py   # custom port
"""

from __future__ import annotations

import os

from app.web import app


if __name__ == "__main__":
    port = int(os.environ.get("PORT", "5000"))
    host = os.environ.get("HOST", "0.0.0.0")
    debug = os.environ.get("DEBUG", "0") == "1"
    app.run(host=host, port=port, debug=debug)
