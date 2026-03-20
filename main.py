import os
import re
import cv2
import math
import queue
import threading
import concurrent.futures
import numpy as np
import pandas as pd
import pytesseract
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import openpyxl
# ====== TESSERACT ======
DEFAULT_TESSERACT = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
if os.path.exists(DEFAULT_TESSERACT):
    pytesseract.pytesseract.tesseract_cmd = DEFAULT_TESSERACT

# ====== QUEUE ======
log_queue = queue.Queue()

def log(msg):
    log_queue.put(msg)

# ====== IMAGE ======
def read_image_unicode(path):
    with open(path, 'rb') as f:
        data = np.frombuffer(f.read(), np.uint8)
    return cv2.imdecode(data, cv2.IMREAD_COLOR)

def preprocess(img):
    h, w = img.shape[:2]
    scale = max(1.5, 2500 / min(h, w))
    img = cv2.resize(img, None, fx=scale, fy=scale)

    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    _, th = cv2.threshold(gray, 180, 255, cv2.THRESH_BINARY_INV)

    kernel_h = cv2.getStructuringElement(cv2.MORPH_RECT, (50,1))
    kernel_v = cv2.getStructuringElement(cv2.MORPH_RECT, (1,50))

    th = cv2.subtract(th, cv2.morphologyEx(th, cv2.MORPH_OPEN, kernel_h))
    th = cv2.subtract(th, cv2.morphologyEx(th, cv2.MORPH_OPEN, kernel_v))

    return th

def split_4(img):
    h, w = img.shape[:2]
    return {
        "S21_amp": img[0:h//2, 0:w//2],
        "S21_gvz": img[0:h//2, w//2:w],
        "S11": img[h//2:h, 0:w//2],
        "S22": img[h//2:h, w//2:w]
    }

# ====== OCR ======
def ocr(img):
    return pytesseract.image_to_string(img, config="--psm 6 -l rus+eng").lower()

NUM_RE = re.compile(r"[-+]?\d+[.,]?\d*")
S11_RANGE = (1.0, 2.0)
S22_RANGE = (1.0, 2.0)
RX_ALLOWED_LEVELS = [40, 25, 10, -5, -20]
RX_TOLERANCE_DB = 4
MAX_WORKERS = max(1, min(8, (os.cpu_count() or 1)))

def normalize_token(token):
    token = token.lower()
    translit = str.maketrans({
        "a": "а", "c": "с", "e": "е", "k": "к", "m": "м", "h": "н", "o": "о", "p": "р", "t": "т", "x": "х",
        "y": "у", "b": "в", "n": "п"
    })
    return token.translate(translit)

def parse_num(text):
    nums = NUM_RE.findall(text)
    if not nums:
        return None
    return float(nums[-1].replace(",", "."))

def parse_nums(text):
    return [float(n.replace(",", ".")) for n in NUM_RE.findall(text)]

def in_range(val, value_range):
    if value_range is None:
        return val is not None
    if val is None:
        return False
    lo, hi = value_range
    return lo <= val <= hi

def pick_number(candidates, value_range=None):
    if not candidates:
        return None
    if value_range is None:
        return candidates[-1]
    for n in reversed(candidates):
        if in_range(n, value_range):
            return n
    return None

def extract_from_lines(text, keys, value_range=None):
    lines = text.split("\n")
    norm_keys = [normalize_token(k) for k in keys]
    for i, line in enumerate(lines):
        norm_line = normalize_token(line)
        if any(k in norm_line for k in norm_keys):
            val = pick_number(parse_nums(line), value_range=value_range)
            if val is not None:
                return val
            if i + 1 < len(lines):
                val = pick_number(parse_nums(lines[i + 1]), value_range=value_range)
                if val is not None:
                    return val
    return None

def extract_from_layout(img, keys, value_range=None):
    data = pytesseract.image_to_data(
        img, config="--psm 6 -l rus+eng", output_type=pytesseract.Output.DICT
    )
    norm_keys = [normalize_token(k) for k in keys]
    tokens = []
    for i, raw in enumerate(data["text"]):
        text = raw.strip()
        conf = int(data["conf"][i]) if str(data["conf"][i]).lstrip("-").isdigit() else -1
        if not text or conf < 0:
            continue
        tokens.append({
            "text": text,
            "norm": normalize_token(text),
            "left": data["left"][i],
            "top": data["top"][i],
            "width": data["width"][i],
            "height": data["height"][i],
        })

    for tk in tokens:
        if not any(k in tk["norm"] for k in norm_keys):
            continue
        y = tk["top"]
        x_right = tk["left"] + tk["width"]
        candidates = []
        for other in tokens:
            if abs(other["top"] - y) > max(tk["height"], other["height"]) * 1.6:
                continue
            if other["left"] + other["width"] < x_right - 5:
                continue
            num = parse_num(other["text"])
            if num is None:
                continue
            if not in_range(num, value_range):
                continue
            dist = abs(other["left"] - x_right)
            candidates.append((dist, num))
        if candidates:
            candidates.sort(key=lambda item: item[0])
            return candidates[0][1]
    return None

def extract_metric(img, keys, value_range=None):
    text = ocr(img)
    val = extract_from_lines(text, keys, value_range=value_range)
    if val is not None:
        return val
    return extract_from_layout(img, keys, value_range=value_range)

def parse_image(path):
    img = read_image_unicode(path)
    zones = split_4(img)

    # Подготавливаем каждую зону отдельно, чтобы устойчиво работать
    # при "плавающей" области с текстовыми значениями.
    pre = {k: preprocess(v) for k, v in zones.items()}

    return {
        "S21_amp_avg": extract_metric(pre["S21_amp"], ["сред", "сред:", "cpea", "cped"], value_range=(-10, 60)),
        "S21_gvz_uneven": extract_metric(pre["S21_gvz"], ["неравн", "неравн:", "нераен", "неревн"], value_range=(0, 20)),
        "S11_avg": extract_metric(pre["S11"], ["сред", "сред:", "cpea", "cped"], value_range=S11_RANGE),
        "S22_avg": extract_metric(pre["S22"], ["сред", "сред:", "cpea", "cped"], value_range=S22_RANGE)
    }

# ====== META ======
def detect_path(p):
    p=p.lower()
    if "tx" in p: return "Tx"
    if "rx" in p: return "Rx"
    return None

def detect_channel(n):
    n=n.lower()
    if "мейн" in n or "main" in n: return "Main"
    if "рез" in n or "reserve" in n: return "Reserve"
    return None

def detect_att(n):
    nums = re.findall(r"\b(0|15|30)\b", n)
    if len(nums)==1: return f"ATT{nums[0]}"
    if len(nums)>=2: return f"{nums[0]}-{nums[1]}"
    return None

def detect_band(n):
    m=re.search(r"(\d{3,4})\D+(\d{3,4})", n)
    return f"{m.group(1)}-{m.group(2)}" if m else None

def metadata(path,file):
    return {
        "Device": os.path.basename(os.path.dirname(os.path.dirname(path))),
        "Channel": detect_channel(file),
        "Path": detect_path(path),
        "Config": detect_att(file),
        "Band": detect_band(file)
    }

# ====== PHYSICS ======
def expected(cfg):
    if not cfg: return None
    nums=list(map(int,re.findall(r"\d+",cfg)))
    return 37 - sum(nums)

def conf(m,e):
    if m is None or e is None: return 0
    return math.exp(-abs(m-e)/10)

def conf_rx(m):
    if m is None:
        return 0
    d = min(abs(m - ref) for ref in RX_ALLOWED_LEVELS)
    return max(0.0, 1 - d / RX_TOLERANCE_DB)

def expected_for_meta(meta):
    if meta.get("Path") == "Rx":
        return None
    return expected(meta.get("Config"))

def confidence_for_meta(meta, s21_amp):
    if meta.get("Path") == "Rx":
        return conf_rx(s21_amp)
    return conf(s21_amp, expected(meta.get("Config")))

def quality(vals, amp_conf):
    s11_ok = in_range(vals.get("S11_avg"), S11_RANGE)
    s22_ok = in_range(vals.get("S22_avg"), S22_RANGE)

    if amp_conf > 0.9 and s11_ok and s22_ok:
        return "OK"
    if amp_conf > 0.7 and (s11_ok or s22_ok):
        return "CHECK"
    return "BAD"

def process_file(path):
    vals = parse_image(path)
    meta = metadata(path, os.path.basename(path))
    exp = expected_for_meta(meta)
    c = confidence_for_meta(meta, vals["S21_amp_avg"])
    q = quality(vals, c)
    return {
        **meta, **vals,
        "Expected": exp,
        "Confidence": round(c, 3),
        "Quality": q
    }

# ====== WORKER ======
def worker(input_dir, output_dir, progress_var):
    rows=[]
    files=[]

    for r,_,f in os.walk(input_dir):
        for file in f:
            if file.lower().endswith((".png",".jpg",".jpeg")):
                files.append(os.path.join(r,file))

    total=len(files)
    if total == 0:
        log("Нет файлов изображений для обработки")
        return

    done = 0
    with concurrent.futures.ThreadPoolExecutor(max_workers=MAX_WORKERS) as ex:
        future_map = {ex.submit(process_file, path): path for path in files}

        for future in concurrent.futures.as_completed(future_map):
            path = future_map[future]
            try:
                rows.append(future.result())
                log(f"OK: {os.path.basename(path)}")
            except Exception as e:
                log(f"ERR: {path} -> {e}")
            done += 1
            progress_var.set(done/total*100)

    df=pd.DataFrame(rows)
    out=os.path.join(output_dir,"result.xlsx")
    df.to_excel(out,index=False)

    log("ГОТОВО")

# ====== GUI ======
class App:
    def __init__(self,root):
        self.root=root
        self.root.title("Парсер ВАЦ")

        self.input_dir=""
        self.output_dir=""

        tk.Button(root,text="Выбрать папку измерений",command=self.pick_input).pack(fill="x")
        tk.Button(root,text="Выбрать папку сохранения",command=self.pick_output).pack(fill="x")
        tk.Button(root,text="ЗАПУСТИТЬ",command=self.start).pack(fill="x")

        self.progress=ttk.Progressbar(root, length=300)
        self.progress.pack(fill="x")

        self.progress_var=tk.DoubleVar()
        self.progress.config(variable=self.progress_var)

        self.log=tk.Text(root,height=20)
        self.log.pack(fill="both",expand=True)

        self.update_log()

    def pick_input(self):
        self.input_dir=filedialog.askdirectory()
        log(f"Input: {self.input_dir}")

    def pick_output(self):
        self.output_dir=filedialog.askdirectory()
        log(f"Output: {self.output_dir}")

    def start(self):
        if not self.input_dir or not self.output_dir:
            messagebox.showerror("Ошибка","Выбери папки")
            return

        t=threading.Thread(target=worker,
                           args=(self.input_dir,self.output_dir,self.progress_var),
                           daemon=True)
        t.start()

    def update_log(self):
        while not log_queue.empty():
            msg=log_queue.get()
            self.log.insert(tk.END,msg+"\n")
            self.log.see(tk.END)
        self.root.after(100,self.update_log)

# ====== RUN ======
root=tk.Tk()
app=App(root)
root.mainloop()
