# -*- coding: utf-8 -*-
"""本地开发入口：python wsgi.py 或 flask --app wsgi run。"""
from app.app import create_app

app = create_app()

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=False)
