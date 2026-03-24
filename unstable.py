import os
import re
import math
import queue
import threading
import concurrent.futures
import cv2
import numpy as np
import pandas as pd
import pytesseract
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

DEFAULT_TESSERACT = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
if os.path.exists(DEFAULT_TESSERACT):
    pytesseract.pytesseract.tesseract_cmd = DEFAULT_TESSERACT

log_queue = queue.Queue()

NUM_RE = re.compile(r"[-+]?\d+[.,]?\d*")
S11_RANGE = (1.0, 2.0)
S22_RANGE = (1.0, 2.0)
RX_ALLOWED_LEVELS = [40, 25, 10, -5, -20]
RX_TOLERANCE_DB = 4
TX_ALLOWED_LEVELS = [0, -15, -30]
TX_TOLERANCE_DB = 4
RX_BANDS = {1: "1025-1525", 2: "975-1475", 3: "1076-1576"}
TX_BANDS = {1: "1000-1500", 2: "850-1350", 3: "1400-1900", 4: "1500-2000"}
MAX_WORKERS = max(2, min(16, (os.cpu_count() or 2) * 2))


def log(msg):
    log_queue.put(msg)


def read_image_unicode(path):
    with open(path, "rb") as f:
        data = np.frombuffer(f.read(), np.uint8)
    return cv2.imdecode(data, cv2.IMREAD_COLOR)


def preprocess_fast(img):
    h, w = img.shape[:2]
    scale = max(1.2, 1800 / min(h, w))
    img = cv2.resize(img, None, fx=scale, fy=scale)
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    _, th = cv2.threshold(gray, 190, 255, cv2.THRESH_BINARY_INV)
    return th


def split_4(img):
    h, w = img.shape[:2]
    return {
        "S21_amp": img[0:h // 2, 0:w // 2],
        "S21_gvz": img[0:h // 2, w // 2:w],
        "S11": img[h // 2:h, 0:w // 2],
        "S22": img[h // 2:h, w // 2:w],
    }


def normalize_token(token):
    token = token.lower()
    translit = str.maketrans({
        "a": "а", "c": "с", "e": "е", "k": "к", "m": "м", "h": "н", "o": "о", "p": "р", "t": "т", "x": "х",
        "y": "у", "b": "в", "n": "п"
    })
    return token.translate(translit)


def parse_numbers(text):
    return [float(n.replace(",", ".")) for n in NUM_RE.findall(text)]


def in_range(val, rng):
    if val is None:
        return False
    if rng is None:
        return True
    return rng[0] <= val <= rng[1]


def ocr_tokens(img):
    data = pytesseract.image_to_data(
        img,
        config="--psm 6 -l rus+eng",
        output_type=pytesseract.Output.DICT
    )
    tokens = []
    n = len(data["text"])
    for i in range(n):
        raw = data["text"][i].strip()
        if not raw:
            continue
        tokens.append({
            "text": raw,
            "norm": normalize_token(raw),
            "left": data["left"][i],
            "top": data["top"][i],
            "width": data["width"][i],
            "height": data["height"][i],
            "block": data["block_num"][i],
            "par": data["par_num"][i],
            "line": data["line_num"][i],
        })
    return tokens


def extract_metric(tokens, keys, value_range):
    norm_keys = [normalize_token(k) for k in keys]
    best = None
    for tk in tokens:
        if not any(k in tk["norm"] for k in norm_keys):
            continue

        line_candidates = []
        for ot in tokens:
            same_line = (ot["block"] == tk["block"] and ot["par"] == tk["par"] and ot["line"] == tk["line"])
            if not same_line:
                continue
            nums = parse_numbers(ot["text"])
            for n in nums:
                if in_range(n, value_range):
                    dist = abs(ot["left"] - (tk["left"] + tk["width"]))
                    line_candidates.append((dist, n))

        if line_candidates:
            line_candidates.sort(key=lambda x: x[0])
            cand = line_candidates[0][1]
            if best is None:
                best = cand

    if best is not None:
        return best

    # fallback: any token with number in range
    for tk in tokens:
        nums = parse_numbers(tk["text"])
        for n in nums:
            if in_range(n, value_range):
                return n
    return None


def detect_path(p):
    p = p.lower()
    if "tx" in p:
        return "Tx"
    if "rx" in p:
        return "Rx"
    return None


def detect_channel(name):
    n = name.lower()
    if "мейн" in n or "main" in n:
        return "Main"
    if "рез" in n or "reserve" in n:
        return "Reserve"
    return None


def detect_channel_no(name):
    name = name.lower()
    p = re.search(r"(?:канал\D*)(\d+)", name)
    if p:
        return int(p.group(1))
    p = re.search(r"\b(\d+)\s*канал", name)
    return int(p.group(1)) if p else None


def detect_att(name):
    nums = re.findall(r"\b(0|15|30)\b", name)
    if len(nums) == 1:
        return f"ATT{nums[0]}"
    if len(nums) >= 2:
        return f"{nums[0]}-{nums[1]}"
    return None


def metadata(path, file_name):
    p = detect_path(path)
    ch = detect_channel_no(file_name)
    band = RX_BANDS.get(ch) if p == "Rx" else TX_BANDS.get(ch) if p == "Tx" else None
    return {
        "Device": os.path.basename(os.path.dirname(os.path.dirname(path))),
        "Channel": detect_channel(file_name),
        "ChannelNo": ch,
        "Path": p,
        "Config": detect_att(file_name),
        "Band": band,
    }


def conf_with_tol(measured, target, tol_db):
    if measured is None or target is None:
        return 0
    return max(0.0, 1 - abs(measured - target) / tol_db)


def tx_target_from_config(cfg):
    if not cfg:
        return None
    nums = list(map(int, re.findall(r"\d+", cfg)))
    if not nums:
        return None
    total = -sum(nums)
    return min(TX_ALLOWED_LEVELS, key=lambda x: abs(x - total))


def conf_rx(m):
    if m is None:
        return 0
    nearest = min(RX_ALLOWED_LEVELS, key=lambda x: abs(x - m))
    return conf_with_tol(m, nearest, RX_TOLERANCE_DB)


def confidence_for_meta(meta, s21_amp_avg):
    if meta.get("Path") == "Rx":
        return conf_rx(s21_amp_avg)
    if meta.get("Path") == "Tx":
        target = tx_target_from_config(meta.get("Config"))
        return conf_with_tol(s21_amp_avg, target, TX_TOLERANCE_DB)
    return 0


def conf_with_asym_tol(measured, target=1.0, minus_tol=-0.8, plus_tol=5.0):
    if measured is None:
        return 0
    low, high = target + minus_tol, target + plus_tol
    if low <= measured <= high:
        return 1.0
    if measured < low:
        return max(0.0, 1 - (low - measured) / abs(minus_tol))
    return max(0.0, 1 - (measured - high) / abs(plus_tol))


def parse_image(path):
    img = read_image_unicode(path)
    zones = split_4(img)
    tokens = {k: ocr_tokens(preprocess_fast(v)) for k, v in zones.items()}

    return {
        "S21_amp_avg": extract_metric(tokens["S21_amp"], ["усил", "усип", "сред", "cpea", "cped"], (-20, 45)),
        "S21_amp_uneven": extract_metric(tokens["S21_amp"], ["неравн", "нераен", "неревн"], (0.0, 6.5)),
        "S21_gvz_uneven": extract_metric(tokens["S21_gvz"], ["неравн", "нераен", "неревн"], (0.0, 20.0)),
        "S11_avg": extract_metric(tokens["S11"], ["сред", "cpea", "cped"], S11_RANGE),
        "S22_avg": extract_metric(tokens["S22"], ["сред", "cpea", "cped"], S22_RANGE),
    }


def process_file(path):
    vals = parse_image(path)
    meta = metadata(path, os.path.basename(path))
    c = confidence_for_meta(meta, vals["S21_amp_avg"])
    c_uneven = conf_with_asym_tol(vals["S21_amp_uneven"])
    quality = "OK" if c > 0.9 else "CHECK" if c > 0.7 else "BAD"
    return {
        **meta,
        **vals,
        "Confidence": round(c, 3),
        "S21_amp_uneven_conf": round(c_uneven, 3),
        "Quality": quality,
        "SourceImage": path,
    }


def worker(input_dir, output_dir, progress_var):
    files = []
    rows = []

    for root, _, fs in os.walk(input_dir):
        for file in fs:
            if file.lower().endswith((".png", ".jpg", ".jpeg")):
                files.append(os.path.join(root, file))

    total = len(files)
    if total == 0:
        log("Нет файлов изображений для обработки")
        return

    done = 0
    with concurrent.futures.ThreadPoolExecutor(max_workers=MAX_WORKERS) as ex:
        fmap = {ex.submit(process_file, p): p for p in files}
        for fut in concurrent.futures.as_completed(fmap):
            p = fmap[fut]
            try:
                rows.append(fut.result())
                log(f"OK: {os.path.basename(p)}")
            except Exception as e:
                log(f"ERR: {p} -> {e}")
            done += 1
            progress_var.set(done / total * 100)

    rows.sort(key=lambda r: (str(r.get("Device", "")), str(r.get("Path", "")), str(r.get("Band", ""))))
    out = os.path.join(output_dir, "results_unstable.xlsx")
    pd.DataFrame(rows).to_excel(out, index=False)
    log(f"ГОТОВО: {out}")


class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Парсер ВАЦ (UNSTABLE FAST)")

        self.input_dir = ""
        self.output_dir = ""

        tk.Button(root, text="Выбрать папку измерений", command=self.pick_input).pack(fill="x")
        tk.Button(root, text="Выбрать папку сохранения", command=self.pick_output).pack(fill="x")
        tk.Button(root, text="ЗАПУСТИТЬ (FAST)", command=self.start).pack(fill="x")

        self.progress = ttk.Progressbar(root, length=300)
        self.progress.pack(fill="x")

        self.progress_var = tk.DoubleVar()
        self.progress.config(variable=self.progress_var)

        self.log = tk.Text(root, height=20)
        self.log.pack(fill="both", expand=True)

        self.update_log()

    def pick_input(self):
        self.input_dir = filedialog.askdirectory()
        log(f"Input: {self.input_dir}")

    def pick_output(self):
        self.output_dir = filedialog.askdirectory()
        log(f"Output: {self.output_dir}")

    def start(self):
        if not self.input_dir or not self.output_dir:
            messagebox.showerror("Ошибка", "Выбери папки")
            return

        t = threading.Thread(target=worker, args=(self.input_dir, self.output_dir, self.progress_var), daemon=True)
        t.start()

    def update_log(self):
        while not log_queue.empty():
            msg = log_queue.get()
            self.log.insert(tk.END, msg + "\n")
            self.log.see(tk.END)
        self.root.after(100, self.update_log)


if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
