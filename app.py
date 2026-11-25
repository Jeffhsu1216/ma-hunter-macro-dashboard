"""
M&A Hunter 總經儀表板
Macro Dashboard Web Application

部署：Render Web Service
網址：https://ma-hunter-macro-dashboard.onrender.com
"""

from flask import Flask, render_template, jsonify, request, abort
import hmac
from data_fetcher import fetch_all, clear_cache
from datetime import datetime
import pytz
import logging
import os

logging.basicConfig(level=logging.INFO)
app = Flask(__name__)

taipei_tz = pytz.timezone("Asia/Taipei")


@app.route("/")
def dashboard():
    """主頁面 — 總經儀表板"""
    data = fetch_all()
    now = datetime.now(taipei_tz)
    weekday_map = {0: "一", 1: "二", 2: "三", 3: "四", 4: "五", 5: "六", 6: "日"}
    weekday = weekday_map[now.weekday()]
    is_weekend = now.weekday() >= 5

    return render_template(
        "dashboard.html",
        data=data,
        today=now.strftime("%Y/%m/%d"),
        weekday=weekday,
        is_weekend=is_weekend,
    )


@app.route("/api/data")
def api_data():
    """API 端點 — JSON 格式"""
    return jsonify(fetch_all())


def _require_admin():
    """有副作用的端點：只收 POST，且須帶 X-Admin-Token（環境變數 ADMIN_TOKEN）。
    未設定 ADMIN_TOKEN 時端點一律關閉。"""
    expected = os.environ.get("ADMIN_TOKEN", "")
    given = request.headers.get("X-Admin-Token", "")
    if not expected or not hmac.compare_digest(given.encode(), expected.encode()):
        abort(403)


@app.route("/api/refresh", methods=["POST"])
def api_refresh():
    """強制刷新快取"""
    _require_admin()
    clear_cache()
    data = fetch_all()
    return jsonify({"status": "refreshed", "data": data})


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=os.environ.get("FLASK_DEBUG") == "1")
