import os
import pandas as pd
from datetime import datetime
import re

# ===================== KONFIGURACJA =====================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

NET_KEYS = ["net", "netto", "subtotal", "základ", "base"]
VAT_KEYS = ["vat", "mwst", "tax", "moms", "dph", "áfa"]

# ===================== FUNKCJE =====================

def load_ocr_texts(ocr_txt_dir):
    texts = {}
    for file in os.listdir(ocr_txt_dir):
        if file.lower().endswith(".txt"):
            with open(os.path.join(ocr_txt_dir, file), "r", encoding="utf-8") as f:
                texts[file] = f.read()
    return texts


def normalize_number(text):
    text = text.replace(" ", "").replace(",", ".")
    nums = re.findall(r"\d+\.\d+", text)
    if nums:
        return float(nums[0])
    return ""


def extract_amount(text, keywords):
    for line in text.splitlines():
        low = line.lower()
        if any(k in low for k in keywords):
            nums = re.findall(r"\d{1,3}(?:[ .]\d{3})*[.,]\d{2}", line)
            if nums:
                return float(nums[-1].replace(" ", "").replace(".", "").replace(",", ".", 1))
    return ""


def extract_brutto(text):
    numbers = re.findall(r"\d{1,3}(?:[ .]\d{3})*[.,]\d{2}", text)
    values = [float(n.replace(" ", "").replace(".", "").replace(",", ".", 1)) for n in numbers]
    return max(values) if values else ""


def extract_amount_martex(text):
    netto = ""
    vat = ""

    for line in text.splitlines():
        if re.search(r"\d{1,2}\s*%", line):
            nums = re.findall(r"\d{1,3}(?:[ .]\d{3})*[.,]\d{2}", line)
            if len(nums) >= 2:
                netto = float(nums[0].replace(" ", "").replace(".", "").replace(",", ".", 1))
                vat   = float(nums[1].replace(" ", "").replace(".", "").replace(",", ".", 1))
                break

    return netto, vat


def extract_invoice_date(text, seller=""):
    # === WROV ===
    if seller == "UNIUNEA NATIONALA A TRANSPORTATORILOR RUTIERI DIN ROMANIA":
        for line in text.splitlines():
            if "data:" in line.lower():
                match = re.search(r"\d{2}-\d{2}-\d{4}", line)
                if match:
                    return match.group()


    # === OGÓLNE ===
    for line in text.splitlines():
        match = re.search(r"\d{2}[./-]\d{2}[./-]\d{4}", line)
        if match:
            return match.group()

    return ""



def extract_seller(text, faktura):
    if re.match(r"^WROV\d{7}$", faktura):
        return "UNIUNEA NATIONALA A TRANSPORTATORILOR RUTIERI DIN ROMANIA"

    if re.match(r"^[A-Za-z][A-Za-z0-9]?/[A-Za-z]{2,3}/202\d/\d{5}$", faktura):
        return "MARTEX SP. Z O.O."

    seller_keywords = ["sprzedawca", "seller", "lieferant"]
    lines = text.splitlines()

    for i, line in enumerate(lines):
        if any(k in line.lower() for k in seller_keywords):
            if i + 1 < len(lines):
                return lines[i + 1].strip().upper()

    fallback = [
        l.strip().upper()
        for l in lines[:15]
        if len(l) > 8 and not any(c.isdigit() for c in l)
    ]

    return fallback[0] if fallback else ""

def extract_amount_razem(text: str):
    """
    Szuka linii typu:
      RAZEM 1 234,56  283,95
    i zwraca (netto, vat).
    """
    # dopuszczamy: "razem", "razem:", "razem do zapłaty" itd.
    razem_re = re.compile(r"\brazem\b", re.IGNORECASE)

    # ta sama logika wyciągania kwot co u Ciebie (PL/EU format)
    num_re = re.compile(r"\d{1,3}(?:[ .]\d{3})*[.,]\d{2}")

    for line in text.splitlines():
        if not razem_re.search(line):
            continue

        # bierzemy fragment ZA słowem "razem" (żeby nie złapać czegoś przed)
        parts = razem_re.split(line, maxsplit=1)
        tail = parts[1] if len(parts) > 1 else line

        nums = num_re.findall(tail)
        if len(nums) >= 2:
            netto = float(nums[0].replace(" ", "").replace(".", "").replace(",", ".", 1))
            vat   = float(nums[1].replace(" ", "").replace(".", "").replace(",", ".", 1))
            return netto, vat

    return "", ""


def generate_xlsx(folder, output_dir):
    # folder może być bazą (...\Faktury) albo ...\Faktury\scans
    base_folder = folder
    if os.path.basename(folder).lower() == "scans":
        base_folder = os.path.dirname(folder)

    ocr_txt_dir = os.path.join(base_folder, "scans", "ocr_txt")
    os.makedirs(output_dir, exist_ok=True)

    if not os.path.isdir(ocr_txt_dir):
        raise FileNotFoundError("Brak folderu ocr_txt – najpierw uruchom OCR")

    files = [f for f in os.listdir(ocr_txt_dir) if f.lower().endswith(".txt")]
    files.sort()

    rows = []

    for filename in files:
        name = os.path.splitext(filename)[0]
        if "_" not in name:
            continue

        faktura_raw, rejestracja = name.split("_", 1)
        faktura = faktura_raw.replace("-", "/").strip()
        rejestracja = rejestracja.strip()

        txt_path = os.path.join(ocr_txt_dir, filename)
        with open(txt_path, "r", encoding="utf-8") as f:
            full_text = f.read()

        seller_name = extract_seller(full_text, faktura)

        if seller_name == "MARTEX SP. Z O.O.":
            netto, vat = extract_amount_martex(full_text)
        else:
            # 1) Najpierw próbujemy "RAZEM netto vat"
            netto, vat = extract_amount_razem(full_text)

            # 2) Fallback: dotychczasowe reguły po słowach kluczach
            if netto == "":
                netto = extract_amount(full_text, NET_KEYS)
            if vat == "":
                vat = extract_amount(full_text, VAT_KEYS)

        rows.append({
            "Nr faktury": faktura,
            "Data wystawienia": extract_invoice_date(full_text, seller_name),
            "Nr rejestracyjny": rejestracja,
            "Sprzedawca": seller_name,
            "Netto": netto,
            "VAT": vat,
            "Brutto": extract_brutto(full_text),
            "Plik": filename
        })

    if not rows:
        raise ValueError(
            "Nie powstały żadne wiersze. Sprawdź nazwy plików TXT: muszą mieć format NRFAKTURY_REJESTRACJA.txt"
        )

    df = pd.DataFrame(rows)

    output_file = os.path.join(
        output_dir,
        f"wynik_{datetime.now().strftime('%Y-%m-%d_%H-%M')}.xlsx"
    )

    print("FOLDER:", folder)
    print("OCR_TXT_DIR:", ocr_txt_dir)
    print("TXT FILES:", len(files), files[:5])

    print("ROWS:", len(rows))
    print("ZAPISUJE:", output_file)

    df.to_excel(output_file, index=False)
    return output_file

if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser()
    parser.add_argument("--folder", required=True)
    parser.add_argument("--output", required=True)
    args = parser.parse_args()
    output = generate_xlsx(args.folder, args.output)
    print(f"Gotowe. Plik zapisany: {output}")

