import os
import re
import cv2
import math
import queue
import threading
import numpy as np
import pandas as pd
import pytesseract
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import openpyxl
# ====== TESSERACT ======
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

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

def extract(text, keys):
    lines = text.split("\n")
    for i, line in enumerate(lines):
        for k in keys:
            if k in line:
                nums = re.findall(r"[-+]?\d+[.,]?\d*", line)
                if nums:
                    return float(nums[-1].replace(",", "."))
                if i+1 < len(lines):
                    nums = re.findall(r"[-+]?\d+[.,]?\d*", lines[i+1])
                    if nums:
                        return float(nums[0].replace(",", "."))
    return None

def parse_image(path):
    img = read_image_unicode(path)
    zones = split_4(img)

    return {
        "S21_amp_avg": extract(ocr(preprocess(zones["S21_amp"])), ["сред","cpea"]),
        "S21_gvz_uneven": extract(ocr(preprocess(zones["S21_gvz"])), ["неравн","нераен"]),
        "S11_avg": extract(ocr(preprocess(zones["S11"])), ["сред","cpea"]),
        "S22_avg": extract(ocr(preprocess(zones["S22"])), ["сред","cpea"])
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

# ====== WORKER ======
def worker(input_dir, output_dir, progress_var):
    rows=[]
    files=[]

    for r,_,f in os.walk(input_dir):
        for file in f:
            if file.lower().endswith((".png",".jpg",".jpeg")):
                files.append(os.path.join(r,file))

    total=len(files)

    for i,path in enumerate(files):
        try:
            vals=parse_image(path)
            meta=metadata(path, os.path.basename(path))

            exp=expected(meta["Config"])
            c=conf(vals["S21_amp_avg"],exp)

            q="OK" if c>0.9 else "CHECK" if c>0.7 else "BAD"

            row={**meta,**vals,
                 "Expected":exp,
                 "Confidence":round(c,3),
                 "Quality":q}

            rows.append(row)
            log(f"OK: {os.path.basename(path)}")

        except Exception as e:
            log(f"ERR: {path} -> {e}")

        progress_var.set((i+1)/total*100)

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