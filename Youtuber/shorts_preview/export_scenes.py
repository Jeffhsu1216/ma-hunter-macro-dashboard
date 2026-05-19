#!/usr/bin/env python3
"""
export_scenes.py
把 slides.html 的 9 個 scene 截圖成 1080×1920 PNG
存至 ~/Desktop/Youtuber/Story/
"""

import os, time
from pathlib import Path
from playwright.sync_api import sync_playwright

OUT_DIR = Path.home() / "Desktop" / "Youtuber" / "Story"
OUT_DIR.mkdir(parents=True, exist_ok=True)

N_SCENES = 10
URL = "http://localhost:4891/slides.html"

# 顯示尺寸 405×720 → 1080×1920 (scale=2.666...)
W, H = 405, 720
SCALE = 1080 / W  # 2.6667

def export():
    with sync_playwright() as p:
        browser = p.chromium.launch()
        ctx = browser.new_context(
            viewport={"width": W, "height": H},
            device_scale_factor=SCALE,
        )
        page = ctx.new_page()
        page.goto(URL, wait_until="networkidle")
        page.wait_for_timeout(600)  # 等字型載入

        for i in range(N_SCENES):
            # 切換到第 i 個 scene
            page.evaluate(f"cur = {i}; render();")
            page.wait_for_timeout(250)

            # 只截 #slide 元素
            slide = page.locator("#slide")
            out_path = OUT_DIR / f"{str(i+1).zfill(2)}_MA_Hunter_Story.png"
            slide.screenshot(path=str(out_path))
            print(f"✅ S{i+1:02d} → {out_path.name}")

        browser.close()

    print(f"\n🎉 全部 {N_SCENES} 張 Story 已存至：{OUT_DIR}")

if __name__ == "__main__":
    export()
