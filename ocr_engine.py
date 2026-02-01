import os
import numpy as np
import cv2
import pytesseract
from pdf2image import convert_from_path

# Ustawienia OCR
LANGS = "pol+eng+deu+swe+ces+slk+hun+ita+ro"
TESS_CONFIG = "--oem 3 --psm 6"

import re

def is_text_readable(text: str, min_chars: int = 30) -> bool:
    if not text:
        return False

    text = text.strip()

    if len(text) < min_chars:
        return False

    # musi zawierać jakieś litery albo cyfry
    if not re.search(r"[A-Za-z0-9]", text):
        return False

    return True



def _fast_preprocess(img_bgr):
    return cv2.cvtColor(img_bgr, cv2.COLOR_BGR2GRAY)


def _ocr_image(img_gray):
    return pytesseract.image_to_string(img_gray, lang=LANGS, config=TESS_CONFIG)


def ocr_folder_pdfs(
    pdf_folder: str,
    out_txt_folder: str,
    poppler_path: str | None = None,
    tesseract_cmd: str | None = None,
    dpi: int = 200,
    first_page_only: bool = True,
    on_progress=None,
) -> dict:

    if tesseract_cmd:
        pytesseract.pytesseract.tesseract_cmd = tesseract_cmd

    os.makedirs(out_txt_folder, exist_ok=True)

    pdf_files = [f for f in os.listdir(pdf_folder) if f.lower().endswith(".pdf")]

    total = len(pdf_files)
    skipped_existing = 0
    skipped_unreadable = 0
    done = 0
    errors = 0
    first_error = ""

    for idx, pdf in enumerate(pdf_files, start=1):
        if on_progress:
            on_progress(pdf, idx, total)  # (filename, current, total)

        name = os.path.splitext(pdf)[0]
        pdf_path = os.path.join(pdf_folder, pdf)
        txt_path = os.path.join(out_txt_folder, name + ".txt")

        if os.path.exists(txt_path):
            skipped_existing += 1
            continue

        try:
            pages = convert_from_path(pdf_path, dpi=dpi, poppler_path=poppler_path)

            text = ""
            pages_to_process = pages[:1] if first_page_only else pages

            for page in pages_to_process:
                img = cv2.cvtColor(np.array(page), cv2.COLOR_RGB2BGR)
                gray = _fast_preprocess(img)
                text += _ocr_image(gray) + "\n"

            if not is_text_readable(text):
                skipped_unreadable += 1
                continue

            with open(txt_path, "w", encoding="utf-8", errors="ignore") as f:
                f.write(text)

            done += 1

        except Exception as e:
            errors += 1
            if not first_error:
                first_error = f"{pdf}: {repr(e)}"

    return {
        "total": total,
        "skipped": skipped_existing + skipped_unreadable,
        "skipped_existing": skipped_existing,
        "skipped_unreadable": skipped_unreadable,
        "done": done,
        "errors": errors,
        "first_error": first_error
    }
