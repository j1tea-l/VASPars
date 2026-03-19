import os
import re
import cv2
import numpy as np
import pandas as pd
import easyocr

# ====== CONFIG ======
INPUT_FOLDER = "C:/Users/Adm/Documents/Measure"
OUTPUT_FILE = "result.xlsx"
reader = easyocr.Reader(['ru', 'en'], gpu=True)


# ====== IMAGE PREPROCESSING ======
def read_image_unicode(path):
    with open(path, 'rb') as f:
        data = np.frombuffer(f.read(), np.uint8)
    img = cv2.imdecode(data, cv2.IMREAD_COLOR)
    if img is None:
        raise ValueError(f"Image not found or unreadable: {path}")
    return img


def preprocess_adaptive(img):
    # Масштабируем под любую сторону >= 2000px
    h, w = img.shape[:2]
    scale = max(1.0, 2000 / min(h, w))
    img = cv2.resize(img, None, fx=scale, fy=scale, interpolation=cv2.INTER_CUBIC)

    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    # Adaptive threshold + инвертирование
    th = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
                               cv2.THRESH_BINARY, 15, 4)
    th = cv2.bitwise_not(th)
    # Морфология для очистки
    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (2, 2))
    clean = cv2.morphologyEx(th, cv2.MORPH_OPEN, kernel)
    return clean


# ====== DYNAMIC OCR PARSING ======
def extract_values_dynamic_anyres(img):
    result = reader.readtext(img, detail=1)
    data = {
        "S21_amp_avg": None,
        "S21_gvz_uneven": None,
        "S11_avg": None,
        "S22_avg": None
    }

    boxes = []
    for bbox, text, conf in result:
        boxes.append({"bbox": bbox, "text": text.lower(), "conf": conf})

    for item in boxes:
        text_lower = item["text"].replace("cpea", "сред").replace("cped", "сред")
        if "сред" in text_lower or "неравн" in text_lower:
            # ищем ближайшую цифру
            closest_value = None
            min_dist = float('inf')
            x1, y1 = np.mean(item["bbox"], axis=0)  # центр box
            for other in boxes:
                if other == item:
                    continue
                m = re.search(r"([-+]?\d+[.,]?\d*)", other["text"])
                if m:
                    ox, oy = np.mean(other["bbox"], axis=0)
                    dist = np.sqrt((ox - x1) ** 2 + (oy - y1) ** 2)
                    if dist < min_dist:
                        min_dist = dist
                        closest_value = float(m.group(1).replace(',', '.'))

            if "сред" in text_lower:
                if "s21" in text_lower and data["S21_amp_avg"] is None:
                    data["S21_amp_avg"] = closest_value
                elif "s11" in text_lower:
                    data["S11_avg"] = closest_value
                elif "s22" in text_lower:
                    data["S22_avg"] = closest_value
            elif "неравн" in text_lower:
                if "s21" in text_lower:
                    data["S21_gvz_uneven"] = closest_value
    return data


def parse_image_anyres(path):
    img = read_image_unicode(path)
    pre = preprocess_adaptive(img)
    data = extract_values_dynamic_anyres(pre)
    return data


# ====== METADATA PARSING (как раньше) ======
def detect_path_from_folder(path):
    path_lower = path.lower()
    if "tx" in path_lower:
        return "Tx"
    if "rx" in path_lower:
        return "Rx"
    return None


def detect_channel(name):
    name = name.lower()
    if "main" in name or "мейн" in name:
        return "Main"
    if "reserve" in name or "res" in name or "рез" in name:
        return "Reserve"
    return None


def detect_attenuators(name):
    nums = re.findall(r"\b(0|15|30)\b", name)
    if len(nums) == 1:
        return f"ATT{nums[0]}"
    elif len(nums) >= 2:
        return f"{nums[0]}-{nums[1]}"
    return None


def detect_band(name):
    m = re.search(r"(\d{3,4})\D+(\d{3,4})", name)
    if m:
        return f"{m.group(1)}-{m.group(2)}"
    return None


def parse_metadata(filepath, filename):
    name = os.path.splitext(filename)[0]
    meta = {
        "Device": os.path.basename(os.path.dirname(os.path.dirname(filepath))),
        "Channel": detect_channel(name),
        "Path": detect_path_from_folder(filepath),
        "Config": detect_attenuators(name),
        "Band": detect_band(name)
    }
    return meta


# ====== MAIN ======
rows = []
for root, dirs, files in os.walk(INPUT_FOLDER):
    for file in files:
        if file.lower().endswith((".png", ".jpg", ".jpeg")):
            path = os.path.join(root, file)
            try:
                values = parse_image_anyres(path)
                meta = parse_metadata(path, file)
                row = {**meta, **values}
                print(row)
                rows.append(row)
                print(f"Processed: {file}")
            except Exception as e:
                print(f"Error processing {file}: {e}")

# ====== SAVE ======
df = pd.DataFrame(rows)
columns = [
    "Device", "Channel", "Path", "Config", "Band",
    "S21_amp_avg", "S21_gvz_uneven", "S11_avg", "S22_avg"
]
for col in columns:
    if col not in df.columns:
        df[col] = None
df = df[columns]
df.to_excel(OUTPUT_FILE, index=False)
print(f"Saved to {OUTPUT_FILE}")