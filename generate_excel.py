import os
import pandas as pd
from datetime import datetime
import re
import json
from openai import OpenAI

USE_AI = False  # <- jednym ruchem możesz wyłączyć AI

# ===================== KONFIGURACJA =====================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

NET_KEYS = ["net", "netto", "subtotal", "základ", "base"]
VAT_KEYS = ["vat", "mwst", "tax", "moms", "dph", "áfa"]

# ===================== FUNKCJE =====================

def fix_ocr_separators(s: str) -> str:
    """
    OCR czasem myli separator dziesiętny na: -, – , — itp.
    Zamieniamy je na kropkę, ale TYLKO gdy są między cyframi i przed 2 cyframi na końcu tokenu.
    np. 118-11 -> 118.11
    """
    if not s:
        return s
    s = s.strip()
    s = re.sub(r"(?<=\d)[\-–—](?=\d{2}\b)", ".", s)
    return s


def parse_amount(s: str):
    """
    Parsuje kwoty z OCR:
    - 183.57 / 183,57
    - 1 794.48 / 2 207,21
    - 1.234,56
    - 118-11 (po fix_ocr_separators -> 118.11)
    Zwraca float albo "".
    """
    if not s:
        return ""

    s = fix_ocr_separators(s)
    s = s.strip().replace(" ", "")

    if "," in s and "." in s:
        # ostatni separator wygrywa jako dziesiętny
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "")
            s = s.replace(",", ".")
        else:
            s = s.replace(",", "")
            # dot zostaje
    else:
        if "," in s:
            s = s.replace(",", ".")
        # jeśli tylko ".", zostaje

    try:
        return float(s)
    except ValueError:
        return ""



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
    Zamienia datę na format dd.mm.rrrr i waliduje poprawność.
    Obsługuje:
    - yyyy-mm-dd
    - dd.mm.yyyy
    - dd/mm/yyyy
    - dd-mm-yyyy
    Jeśli data jest niepoprawna (np. 41.01.2026) -> zwraca "".
    """
    if not date_str:
        return ""

    s = date_str.strip()

    # yyyy-mm-dd
    m = re.match(r"(\d{4})-(\d{2})-(\d{2})", s)
    if m:
        yyyy, mm, dd = m.groups()
        try:
            datetime(int(yyyy), int(mm), int(dd))
            return f"{dd}.{mm}.{yyyy}"
        except ValueError:
            return ""

    # dd.mm.yyyy
    m = re.match(r"(\d{2})\.(\d{2})\.(\d{4})", s)
    if m:
        dd, mm, yyyy = m.groups()
        try:
            datetime(int(yyyy), int(mm), int(dd))
            return f"{dd}.{mm}.{yyyy}"
        except ValueError:
            return ""

    # dd/mm/yyyy
    m = re.match(r"(\d{2})/(\d{2})/(\d{4})", s)
    if m:
        dd, mm, yyyy = m.groups()
        try:
            datetime(int(yyyy), int(mm), int(dd))
            return f"{dd}.{mm}.{yyyy}"
        except ValueError:
            return ""

    # dd-mm-yyyy
    m = re.match(r"(\d{2})-(\d{2})-(\d{4})", s)
    if m:
        dd, mm, yyyy = m.groups()
        try:
            datetime(int(yyyy), int(mm), int(dd))
            return f"{dd}.{mm}.{yyyy}"
        except ValueError:
            return ""

    return ""


def extract_amount(text, keywords):
    for line in text.splitlines():
        low = line.lower()
        if any(k in low for k in keywords):
            nums = re.findall(r"\d{1,3}(?:[ .]\d{3})*[.,]\d{2}", line)
            if nums:
                return parse_amount(nums[-1])
    return ""



def extract_brutto(text):
    numbers = re.findall(r"\d{1,3}(?:[ .]\d{3})*[.,-]\d{2}", text)
    values = []
    for n in numbers:
        v = parse_amount(n)
        if v != "":
            values.append(v)
    return max(values) if values else ""

def detect_vat_rate(text: str):
    """
    Szuka w tekście stawki VAT w stylu: 23%, 8%, 5%, 0%
    Zwraca ułamek (0.23) albo None jeśli nie znaleziono.
    Heurystyka:
    - jeśli jest kilka stawek, bierzemy najwyższą (często to ta "główna")
    """
    # typowe zapisy: 23%, 23 %, VAT 23%, stawka 23%
    rates = re.findall(r"(?<!\d)(\d{1,2})\s*%", text)
    rates = [int(r) for r in rates if 0 <= int(r) <= 30]  # filtr śmieci
    if not rates:
        return None

    # często faktury mają kilka stawek -> bierzemy największą (np. 23 zamiast 8)
    r = max(rates)
    return r / 100.0

def detect_currency(text: str) -> str:
    """
    Wykrywa walutę z OCR.
    Zwraca: PLN, EUR, CZK, SEK, HUF, LEI, NOK albo "".
    """
    if not text:
        return ""

    t = text.upper()

    patterns = [
        ("PLN", r"\bPLN\b|\bZŁ\b|\bZL\b"),
        ("EUR", r"\bEUR\b|€"),
        ("CZK", r"\bCZK\b|\bKČ\b"),
        ("SEK", r"\bSEK\b"),
        ("HUF", r"\bHUF\b|\bFT\b"),
        ("LEI", r"\bLEI\b|\bRON\b"),  # rumuńska waluta, u Ciebie ma być LEI
        ("NOK", r"\bNOK\b"),
    ]

    for code, pat in patterns:
        if re.search(pat, t):
            return code

    return ""


def extract_totals_by_context(text: str):
    """
    Szuka linii z sumami po kontekście (razem/total/suma/brutto/vat/do zapłaty/wartość dokumentu).
    Zwraca (netto, vat, brutto) - każdy może być "" jeśli brak.
    """
    ctx = [
        "razem", "total", "suma", "sumy",
        "brutto", "netto", "vat", "mwst", "tax",
        "do zapłaty", "amount due", "wartość dokumentu", "wartosc dokumentu"
    ]

    num_re = re.compile(r"\d{1,3}(?:[ .]\d{3})*[.,-]\d{2}")

    best = None  # (score, netto, vat, brutto)

    for line in text.splitlines():
        low = line.lower()
        score = sum(1 for k in ctx if k in low)
        if score == 0:
            continue

        line_fixed = fix_ocr_separators(line)
        nums = num_re.findall(line_fixed)

        if len(nums) < 2:
            continue

        values = [parse_amount(n) for n in nums]
        values = [v for v in values if v != ""]

        if len(values) < 2:
            continue

        netto = vat = brutto = ""

        if len(values) >= 3:
            netto, vat, brutto = values[0], values[1], values[2]
        else:
            netto, brutto = values[0], values[1]
            vat = round(brutto - netto, 2)

        # sanity
        if brutto != "" and netto != "":
            vat_from_diff = round(brutto - netto, 2)
            if vat == "" or abs(vat_from_diff - float(vat)) > 0.50:
                vat = vat_from_diff

        cand = (score, netto, vat, brutto)
        if best is None or cand[0] > best[0]:
            best = cand

    if best:
        _, netto, vat, brutto = best
        return netto, vat, brutto

    return "", "", ""


def calc_netto_vat_from_brutto(brutto: float, rate: float):
    netto = round(brutto / (1.0 + rate), 2)
    vat = round(brutto - netto, 2)
    return netto, vat



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
                netto = parse_amount(nums[0])
                vat   = parse_amount(nums[1])
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

        # === PRIORYTET: Data wystawienia / Invoice date ===
    for line in text.splitlines():
        low = line.lower()
        if "data wystawienia" in low or "invoice date" in low:
            m = re.search(r"\d{4}-\d{2}-\d{2}", line)
            if m:
                return m.group()
            m = re.search(r"\d{2}[./-]\d{2}[./-]\d{4}", line)
            if m:
                return m.group()

        # === OGÓLNE (fallback) ===
    for line in text.splitlines():
        m = re.search(r"\d{4}-\d{2}-\d{2}", line)
        if m:
            return m.group()
        m = re.search(r"\d{2}[./-]\d{2}[./-]\d{4}", line)
        if m:
            return m.group()

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

    lines = [l.strip() for l in text.splitlines() if l.strip()]

    # słowa kluczowe
    seller_headers = ["sprzedawca", "seller", "lieferant"]
    buyer_headers  = ["nabywca", "kupujący", "kupujacy", "buyer", "odbiorca", "recipient", "customer"]

    # rzeczy, które często pojawiają się w pobliżu i psują wybór (Twoje: x-trade transport)
    banned_contains = [
        "x-trade transport",  # <- to chciałaś ignorować
        "x trade transport",
        "x-trade",
        "transport",
    ]

    # linie, które są "śmieciem" zamiast nazwy firmy
    banned_exact = {
        "sprzedawca", "seller", "lieferant",
        "nabywca", "kupujący", "kupujacy", "buyer",
    }

    def is_good_company_line(s: str) -> bool:
        low = s.lower()

        if low in banned_exact:
            return False

        # jeśli trafimy na sekcję kupującego, stop
        if any(h in low for h in buyer_headers):
            return False

        # ignorujemy linie z x-trade transport + podobne
        if any(b in low for b in banned_contains):
            return False

        # nie bierzemy linii z samymi cyframi / datami / zbyt krótkich
        if len(s) < 4:
            return False
        if all(ch.isdigit() or ch in ".,-/" for ch in s):
            return False

        # jeśli wygląda jak NIP/REGON/itp. to nie jest nazwa
        if "nip" in low or "vat" in low or "regon" in low:
            return False

        return True

    # 1) preferowane: po nagłówku sprzedawcy szukamy pierwszej sensownej linijki w kolejnych 6 liniach
    for i, line in enumerate(lines):
        low = line.lower()
        if any(h in low for h in seller_headers):
            for j in range(i + 1, min(i + 7, len(lines))):
                cand = lines[j]
                if is_good_company_line(cand):
                    return cand.upper()
            break

    # 2) fallback: pierwsza sensowna "firma" z początku dokumentu, ale też z filtrami
    for cand in lines[:20]:
        if is_good_company_line(cand) and len(cand) > 8 and not any(c.isdigit() for c in cand):
            return cand.upper()

    return ""


def extract_amount_razem(text: str):
    razem_re = re.compile(r"\brazem\b", re.IGNORECASE)

    # <-- tu dopuszczamy . , albo - jako separator dziesiętny
    num_re = re.compile(r"\d{1,3}(?:[ .]\d{3})*[.,-]\d{2}")

    for line in text.splitlines():
        if not razem_re.search(line):
            continue

        parts = razem_re.split(line, maxsplit=1)
        tail = parts[1] if len(parts) > 1 else line

        # najpierw naprawiamy typowe OCR fuckupy
        tail = fix_ocr_separators(tail)

        nums = num_re.findall(tail)
        if len(nums) >= 2:
            netto = parse_amount(fix_ocr_separators(nums[0]))
            vat   = parse_amount(fix_ocr_separators(nums[1]))
            return netto, vat

    return "", ""



def looks_like_worth_calling_ai(text: str) -> bool:
    # musi być liczba z groszami
    has_amount = bool(re.search(r"\d+[.,]\d{2}", text))
    # i jakieś słowo-klucz sumy
    has_keyword = any(k in text.lower() for k in ["razem", "total", "suma", "vat", "netto", "mwst", "tax"])
    return has_amount and has_keyword

def to_money(val):
    """
    Zwraca float z 2 miejscami po przecinku albo "" jeśli brak wartości
    """
    if val == "" or val is None:
        return ""
    try:
        return round(float(val), 2)
    except (ValueError, TypeError):
        return ""


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

        currency = detect_currency(full_text)

        # === REGUŁA: POTWIERDZENIA POCZTY (00...) ===
        if faktura.startswith("(00)"):
            seller_name = "POCZTA POLSKA"
            invoice_date = ""
            netto = ""
            vat = 0.0
            brutto = ""

            if not currency:
                currency = "PLN"

            rows.append({
                "Nr faktury": faktura,
                "Data wystawienia": "",
                "Nr rejestracyjny": rejestracja,
                "Sprzedawca": seller_name,
                "Netto": netto,
                "VAT": vat,
                "Brutto": brutto,
                "Waluta": currency,
                "Plik": filename
            })
            continue

        # ======= SPRZEDAWCA + DATA (bo później tego używamy) =======
        seller_name = extract_seller(full_text, faktura)
        invoice_date = extract_invoice_date(full_text, seller_name)

        # reguła "FxxxxxG...P" = Poczta Polska
        if is_poczta_polska_invoice(faktura):
            seller_name = "POCZTA POLSKA"
            invoice_date = extract_invoice_date_poczta(full_text)
            if not currency:
                currency = "PLN"

        # ======= KWOTY: globalna logika (bez AI) =======

        brutto = extract_brutto(full_text)

        # 0) wykryj stawkę VAT (do ratunku)
        vat_rate = detect_vat_rate(full_text)  # np. 0.23, 0.08, None

        # 1) Najpierw totals po kontekście (działa nawet jak "Razem" jest zjebane)
        netto, vat, brutto_ctx = extract_totals_by_context(full_text)

        # jeśli kontekst znalazł brutto, a globalne brutto nie
        if brutto == "" and brutto_ctx != "":
            brutto = brutto_ctx

        # 2) Potem klasyczne "Razem: netto vat ..."
        if netto == "" or vat == "":
            n2, v2 = extract_amount_razem(full_text)
            if netto == "":
                netto = n2
            if vat == "":
                vat = v2

        # 3) Potem keywordy (netto/vat)
        if netto == "":
            netto = extract_amount(full_text, NET_KEYS)
        if vat == "":
            vat = extract_amount(full_text, VAT_KEYS)

        # 4) OSTATNIA DESKA: brutto -> netto/vat wg wykrytej stawki (albo 23% jak nie wykryło)
        if brutto != "" and (netto == "" or vat == ""):
            rate = vat_rate if vat_rate is not None else 0.23
            n3, v3 = calc_netto_vat_from_brutto(float(brutto), rate)
            if netto == "":
                netto = n3
            if vat == "":
                vat = v3

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

        netto = to_money(netto)
        vat = to_money(vat)
        brutto = to_money(brutto)

        # === REGUŁA: znak VAT zgodny z netto ===
        if netto != "" and vat != "":
            # netto nieujemne -> VAT nie może być ujemny
            if netto >= 0 and vat < 0:
                vat = abs(vat)

            # faktura minusowa -> VAT też minus (zgodny znak)
            elif netto < 0 and vat > 0:
                vat = -abs(vat)

        rows.append({
            "Nr faktury": faktura,
            "Data wystawienia": normalize_date(invoice_date),
            "Nr rejestracyjny": rejestracja,
            "Sprzedawca": seller_name,
            "Netto": netto,
            "VAT": vat,
            "Brutto": brutto,
            "Waluta": currency,
            "Plik": filename
        })

    if not rows:
        raise ValueError(
            "Nie powstały żadne wiersze. Sprawdź nazwy plików TXT: muszą mieć format NRFAKTURY_REJESTRACJA.txt"
        )

    df = pd.DataFrame(rows)

    df = df[[
        "Nr faktury",
        "Data wystawienia",
        "Netto",
        "VAT",
        "Waluta",
        "Brutto",
        "Nr rejestracyjny",
    ]]

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

