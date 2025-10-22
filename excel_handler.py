# excel_handler.py
import os
from openpyxl import Workbook, load_workbook
from datetime import datetime
from helpers import clean_word_for_compare
from dictionary_helper import get_explanation
from config import EXCEL_FILE, TEMP_EXCEL, SHEET_NAME

# ========================
# 📘 Initialize Workbook
# ========================
def init_workbook():
    """
    Load or create workbook and ensure sheet exists with:
    ["No.", "New Words", "Sentence", "Explanation", "Date/Time"]
    Returns (wb, ws)
    """
    if os.path.exists(EXCEL_FILE):
        wb = load_workbook(EXCEL_FILE)
        if SHEET_NAME in wb.sheetnames:
            ws = wb[SHEET_NAME]
        else:
            ws = wb.create_sheet(SHEET_NAME)
            ws.append(["No.", "New Words", "Sentence", "Explanation", "Date/Time"])
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = SHEET_NAME
        ws.append(["No.", "New Words", "Sentence", "Explanation", "Date/Time"])
    return wb, ws

# ========================
# 📊 Load Existing Words
# ========================
def load_existing_words_set(ws):
    """
    Read 'New Words' column and return a set of cleaned lowercase words.
    """
    s = set()
    for row in ws.iter_rows(min_row=2, min_col=2, max_col=2, values_only=True):
        val = row[0]
        if val:
            cleaned = clean_word_for_compare(str(val))
            if cleaned:
                s.add(cleaned)
    return s

# ========================
# ➕ Append Sentence Rows
# ========================
def append_rows_for_sentence(ws, sentence, new_words):
    """
    Append new rows for each new word in sentence.
    Column 1: auto-increment number
    Column 2: new word
    Column 3: full sentence
    Column 4: explanation from dictionary_helper
    Column 5: timestamp
    """
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    next_no = ws.max_row  # next number after header

    if new_words:
        for w in new_words:
            explanation = get_explanation(w)
            ws.append([next_no, w, sentence, explanation, timestamp])
            next_no += 1
    else:
        ws.append([next_no, "", sentence, "", timestamp])

# ========================
# 💾 Safe Save Workbook
# ========================
def save_workbook_safe(wb):
    """
    Save workbook safely to avoid lock or corruption issues.
    """
    wb.save(TEMP_EXCEL)
    wb.close()
    if os.path.exists(EXCEL_FILE):
        os.remove(EXCEL_FILE)
    os.replace(TEMP_EXCEL, EXCEL_FILE)
