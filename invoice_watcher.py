import os
import re
import time
import shutil
import pandas as pd
import pdfplumber
import pytesseract
from PIL import Image
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from collections import defaultdict

# ================== ONEDRIVE PATH ==================
ONEDRIVE = os.environ.get("OneDrive")
if not ONEDRIVE:
    raise RuntimeError("❌ OneDrive not found")

BASE = os.path.join(ONEDRIVE, "Invoices")
INPUT = os.path.join(BASE, "Input")
PROCESSED = os.path.join(BASE, "Processed")
EXCEL = os.path.join(BASE, "Invoice_Data.xlsx")

os.makedirs(INPUT, exist_ok=True)
os.makedirs(PROCESSED, exist_ok=True)

HEADERS = [
    "Sr.No", "Invoice Date", "Invoice No", "Ref No",
    "Particular", "Amount", "TDS (10%)",
    "Clear Amount", "Comment"
]

# ================== EXCEL HANDLING ==================
def load_excel():
    if not os.path.exists(EXCEL):
        df = pd.DataFrame(columns=HEADERS)
        safe_save(df)
        return df

    try:
        df = pd.read_excel(EXCEL)
    except:
        df = pd.DataFrame(columns=HEADERS)

    if list(df.columns) != HEADERS:
        df = pd.DataFrame(columns=HEADERS)
        safe_save(df)

    return df

def safe_save(df):
    # retry if excel is open
    for _ in range(10):
        try:
            df.to_excel(EXCEL, index=False)
            return
        except PermissionError:
            time.sleep(1)

def next_sr(df):
    df2 = df[df["Particular"].str.upper() != "TOTAL"]
    return len(df2) + 1

# ================== OCR ==================
def ocr_file(path):
    text = ""
    if path.lower().endswith(".pdf"):
        with pdfplumber.open(path) as pdf:
            for p in pdf.pages:
                t = p.extract_text()
                if t:
                    text += t + " "
    else:
        text = pytesseract.image_to_string(Image.open(path))

    return re.sub(r"\s+", " ", text)

# ================== PARTICULAR ==================
def extract_particular(text):
    case_types = [
        r"Writ Petition\s*\(C\)", r"CS\s*\(COMM\)",
        r"LPA", r"IPD", r"FAO", r"RFA",
        r"CM", r"ARB\.?P", r"OMP"
    ]

    pattern = re.compile(
        rf"({'|'.join(case_types)})\s*No\.?\s*(\d+)\s*of\s*(\d{{4}}).*?"
        r"before\s+the\s+([A-Za-z\s]+Court(?:\s+at\s+[A-Za-z\s]+)?)",
        re.I
    )

    matches = pattern.findall(text)
    if not matches:
        return ""

    grouped = defaultdict(list)
    for case, num, year, court in matches:
        grouped[(case.upper(), year, court.strip())].append(int(num))

    result = []
    for (case, year, court), nums in grouped.items():
        nums = sorted(set(nums))
        ranges = []
        s = p = nums[0]

        for n in nums[1:]:
            if n == p + 1:
                p = n
            else:
                ranges.append(f"{s}-{p}" if s != p else str(s))
                s = p = n

        ranges.append(f"{s}-{p}" if s != p else str(s))

        result.append(
            f"{case.title()} No. {', '.join(ranges)} of {year} before the {court}"
        )

    return "; ".join(result)

# ================== AMOUNT ==================
def extract_amount(text):
    patterns = [
        r"Total\s*Invoice\s*Value.*?([0-9,]+\.\d{2})",
        r"Grand\s*Total.*?([0-9,]+\.\d{2})",
        r"Total\s*Amount.*?([0-9,]+\.\d{2})"
    ]

    for p in patterns:
        m = re.search(p, text, re.I)
        if m:
            return float(m.group(1).replace(",", ""))

    return 0.0

# ================== FIELD EXTRACTION ==================
def extract_fields(text, sr):
    date = ""
    inv = ""
    ref = ""

    d = re.search(r"\d{1,2}-[A-Za-z]{3}-\d{4}", text)
    if d:
        date = d.group()

    i = re.findall(r"\b[A-Z0-9]{6,}\b", text)
    if i:
        inv = max(i, key=len)

    r = re.search(r"(Our\s*Ref|Ref)\s*[:\-]?\s*([A-Z0-9\/\-]+)", text, re.I)
    if r:
        ref = r.group(2)

    amt = extract_amount(text)
    tds = round(amt * 0.10, 2)

    return [
        sr, date, inv, ref,
        extract_particular(text),
        amt, tds, "", ""
    ]

# ================== TOTAL ROW ==================
def add_total(df):
    df = df[df["Particular"].str.upper() != "TOTAL"]

    total_amt = df["Amount"].astype(float).sum()
    total_tds = df["TDS (10%)"].astype(float).sum()

    total_row = {
        "Sr.No": "",
        "Invoice Date": "",
        "Invoice No": "",
        "Ref No": "",
        "Particular": "TOTAL",
        "Amount": round(total_amt, 2),
        "TDS (10%)": round(total_tds, 2),
        "Clear Amount": "",
        "Comment": ""
    }

    return pd.concat([df, pd.DataFrame([total_row])], ignore_index=True)

# ================== PROCESS FILE ==================
def process_file(path):
    time.sleep(2)

    df = load_excel()
    sr = next_sr(df)

    row = extract_fields(ocr_file(path), sr)
    df.loc[len(df)] = row

    df = add_total(df)
    safe_save(df)

    shutil.move(path, os.path.join(PROCESSED, os.path.basename(path)))
    print("✔ Processed:", os.path.basename(path))

# ================== WATCHER ==================
class Handler(FileSystemEventHandler):
    def on_created(self, e):
        if not e.is_directory:
            process_file(e.src_path)

# ================== MAIN ==================
if __name__ == "__main__":
    print("🚀 Invoice Automation Started")

    # Process existing files on startup
    for f in os.listdir(INPUT):
        path = os.path.join(INPUT, f)
        if os.path.isfile(path):
            process_file(path)

    observer = Observer()
    observer.schedule(Handler(), INPUT, recursive=False)
    observer.start()

    try:
        while True:
            time.sleep(5)
    except KeyboardInterrupt:
        observer.stop()

    observer.join()


