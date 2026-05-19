"""
M&A Hunter 總經儀表板
Macro Dashboard Web Application

部署：Render Web Service
網址：https://ma-hunter-macro-dashboard.onrender.com
"""

from flask import Flask, render_template, jsonify
from data_fetcher import fetch_all, clear_cache
from notifier import push_daily
from datetime import datetime
from apscheduler.schedulers.background import BackgroundScheduler
import pytz
import logging
import os

logging.basicConfig(level=logging.INFO)
app = Flask(__name__)

# ============================================================
# 排程：每日台北時間 08:00 推送 Telegram
# ============================================================
scheduler = BackgroundScheduler()
taipei_tz = pytz.timezone("Asia/Taipei")
scheduler.add_job(push_daily, "cron", hour=8, minute=0, timezone=taipei_tz, id="daily_push")
scheduler.start()


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


@app.route("/api/refresh")
def api_refresh():
    """強制刷新快取"""
    clear_cache()
    data = fetch_all()
    return jsonify({"status": "refreshed", "data": data})


@app.route("/api/push")
def api_push():
    """手動觸發 Telegram 推送"""
    success = push_daily()
    return jsonify({"status": "sent" if success else "failed"})


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=True)
