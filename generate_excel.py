import os
import pandas as pd
from datetime import datetime
import re
import json
from openai import OpenAI

USE_AI = True  # <- jednym ruchem możesz wyłączyć AI

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

def normalize_date(date_str: str) -> str:
    """
    Zamienia datę na format dd.mm.rrrr
    Obsługuje:
    - yyyy-mm-dd
    - dd.mm.yyyy
    - dd/mm/yyyy
    """
    if not date_str:
        return ""

    date_str = date_str.strip()

    # yyyy-mm-dd -> dd.mm.yyyy
    m = re.match(r"(\d{4})-(\d{2})-(\d{2})", date_str)
    if m:
        yyyy, mm, dd = m.groups()
        return f"{dd}.{mm}.{yyyy}"

    # dd.mm.yyyy
    m = re.match(r"(\d{2})\.(\d{2})\.(\d{4})", date_str)
    if m:
        dd, mm, yyyy = m.groups()
        return f"{dd}.{mm}.{yyyy}"

    # dd/mm/yyyy
    m = re.match(r"(\d{2})/(\d{2})/(\d{4})", date_str)
    if m:
        dd, mm, yyyy = m.groups()
        return f"{dd}.{mm}.{yyyy}"

    # jak nie rozpoznaliśmy → zwracamy jak jest
    return date_str


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

def load_invoice_schema():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    schema_path = os.path.join(base_dir, "invoice_schema.json")

    with open(schema_path, "r", encoding="utf-8") as f:
        return json.load(f)

def extract_relevant_lines(text: str) -> str:
    keywords = [
        "razem", "total", "suma", "netto", "vat", "mwst", "tax",
        "sprzedawca", "seller", "firma", "company", "lieferant",
        "brutto", "gross", "do zapłaty", "amount due", "razem do zapłaty"
    ]

    lines = []
    for line in text.splitlines():
        low = line.lower()
        if any(k in low for k in keywords):
            lines.append(line.strip())

    # limit bezpieczeństwa (koszt!)
    return "\n".join(lines[:30])

def ai_extract_fields(text_for_ai: str) -> dict:
    """
    Wywołuje AI i każe mu zwrócić WYŁĄCZNIE JSON zgodny ze schematem invoice_schema.json.
    """
    api_key = os.environ.get("OPENAI_API_KEY")
    if not api_key:
        raise RuntimeError("Brak zmiennej środowiskowej OPENAI_API_KEY")

    client = OpenAI(api_key=api_key)
    schema_pack = load_invoice_schema()  # {"name": "...", "schema": {...}}

    instructions = (
        "Wyciągnij z treści faktury: seller_name, netto, vat, currency. "
        "Zwróć WYŁĄCZNIE JSON zgodny ze schematem. "
        "Jeśli nie masz pewności, ustaw null i daj confidence=low. "
        "W evidence.seller_line wklej linię, z której wziąłeś sprzedawcę. "
        "W evidence.totals_line wklej linię/fragment z sumami netto/VAT."
    )

    # safety: nie wysyłamy za dużo tekstu nawet jeśli ktoś poda śmietnik
    trimmed = text_for_ai[:8000]

    resp = client.responses.create(
        model="gpt-4.1-mini",
        instructions=instructions,
        input=trimmed,
        text={
            "format": {
                "type": "json_schema",
                "name": schema_pack["name"],
                "schema": schema_pack["schema"],
                "strict": True
            }
        },
    )

    return json.loads(resp.output_text)

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

def is_poczta_polska_invoice(faktura: str) -> bool:
    return bool(re.match(r"^F\d{5}G\d{12}P$", faktura))

def extract_invoice_date_poczta(text: str) -> str:
    match = re.search(r"Data wystawienia:\s*(\d{4}-\d{2}-\d{2})", text)
    if match:
        return match.group(1)
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

def looks_like_worth_calling_ai(text: str) -> bool:
    # musi być liczba z groszami
    has_amount = bool(re.search(r"\d+[.,]\d{2}", text))
    # i jakieś słowo-klucz sumy
    has_keyword = any(k in text.lower() for k in ["razem", "total", "suma", "vat", "netto", "mwst", "tax"])
    return has_amount and has_keyword


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

        # === REGUŁA: POCZTA POLSKA ===
        if is_poczta_polska_invoice(faktura):
            seller_name = "POCZTA POLSKA"
            invoice_date = extract_invoice_date_poczta(full_text)
        else:
            seller_name = extract_seller(full_text, faktura)
            invoice_date = extract_invoice_date(full_text, seller_name)

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

        # ================= AI FALLBACK (TANI TRYB) =================
        if USE_AI and (seller_name == "" or netto == "" or vat == ""):
            relevant_text = extract_relevant_lines(full_text)

            ai = None
            if relevant_text and looks_like_worth_calling_ai(relevant_text):
                try:
                    ai = ai_extract_fields(relevant_text)
                except Exception:
                    ai = None

            if ai:
                if seller_name == "" and ai.get("seller_name"):
                    seller_name = ai["seller_name"].strip().upper()

                if netto == "" and ai.get("netto") is not None:
                    netto = ai["netto"]

                if vat == "" and ai.get("vat") is not None:
                    vat = ai["vat"]

        rows.append({
            "Nr faktury": faktura,
            "Data wystawienia": normalize_date(invoice_date),
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

