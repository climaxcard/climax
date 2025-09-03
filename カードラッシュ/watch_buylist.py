# -*- coding: utf-8 -*-
import time, subprocess, os
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

# ====== ここを「カードラッシュ」配下に更新 ======
EXCEL_PATH = r"C:\Users\user\OneDrive\Desktop\デュエマ買取表\buylist.xlsx"
WORK_DIR   = r"C:\Users\user\OneDrive\Desktop\デュエマ買取表\カードラッシュ"
SHEET_URL  = "https://docs.google.com/spreadsheets/d/1gYYmzLkrtAgNZB6dlwzFEqg0VXT7o0QKlFVPS2ln5Xs/edit?gid=0"
CREDS_PATH = r"C:\Users\user\OneDrive\Desktop\デュエマ買取表\カードラッシュ\credentials.json"
# ============================================

class Handler(FileSystemEventHandler):
    def __init__(self):
        self.last_run = 0

    def on_modified(self, event):
        if os.path.abspath(event.src_path) != os.path.abspath(EXCEL_PATH):
            return
        # 連続保存のチャタリング防止（2秒デバウンス）
        now = time.time()
        if now - self.last_run < 2:
            return
        self.last_run = now
        print("[watch] detected change. syncing...")

        # 1) Excel -> シート1 同期
        subprocess.run([
            "python", "-u", "sync_buylist_to_sheet.py",
            "--sheet-url", SHEET_URL,
            "--sheet-name", "シート1",
            "--creds", CREDS_PATH
        ], cwd=WORK_DIR)

        # 2) 差分比較（E=Exp, F=型番, C=名前, O=価格）
        subprocess.run([
            "python", "-u", "compare_sheets_partial.py",
            "--sheet-url", SHEET_URL,
            "--sheet1", "シート1",
            "--sheet2", "CardRush_DM",
            "--sheet1-name-col", "C",
            "--sheet1-exp-col", "E",
            "--sheet1-model-col", "F",
            "--sheet1-price-col", "O",
            "--sheet2-name-col", "カード名",
            "--sheet2-model-col", "型番",
            "--sheet2-price-col", "C",
            "--out-sheet", "差分比較",
            "--creds", CREDS_PATH
        ], cwd=WORK_DIR)

        print("[watch] done.")

if __name__ == "__main__":
    folder = os.path.dirname(EXCEL_PATH)
    event_handler = Handler()
    observer = Observer()
    observer.schedule(event_handler, folder, recursive=False)
    observer.start()
    print("[watch] watching:", EXCEL_PATH)
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()
