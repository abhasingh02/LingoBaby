# main.py
import tkinter as tk
from tkinter import messagebox
import re
import os
import subprocess

from helpers import set_window_icon, clean_word_for_compare, close_excel_if_open, close_word_if_open
from excel_handler import init_workbook, load_existing_words_set, append_rows_for_sentence, save_workbook_safe
from word_handler import init_document, get_next_sentence_number, append_sentence, save_document
from config import ICON_FILE, EXCEL_FILE, DOC_FILE

# GUI logic
def add_sentence_gui(entry_var):
    sentence = entry_var.get().strip()
    if not sentence:
        messagebox.showwarning("Empty", "Please enter a sentence or word.")
        return

    # Close apps before touching files (helps with locks)
    close_excel_if_open()
    close_word_if_open()

    # Excel operations
    wb, ws = init_workbook()
    existing_words = load_existing_words_set(ws)  # set of cleaned words
    # extract words ignoring bracketed content
    cleaned_sentence = re.sub(r"\[.*?\]", "", sentence)
    tokens = re.findall(r"\b[a-zA-Z']+\b", cleaned_sentence)
    new_words = []
    for t in tokens:
        key = clean_word_for_compare(t)
        if key and key not in existing_words:
            new_words.append(t)
            existing_words.add(key)  # ensure uniqueness across this save

    append_rows_for_sentence(ws, sentence, new_words)
    save_workbook_safe(wb)

    # Word operations
    doc = init_document()
    next_num = get_next_sentence_number(doc)
    # create set of cleaned keys for highlighting in doc
    new_words_keys = set(clean_word_for_compare(w) for w in new_words if clean_word_for_compare(w))
    append_sentence(doc, next_num, sentence, new_words_keys)
    save_document(doc)

    entry_var.set("")
    messagebox.showinfo("Saved", f"Saved. {len(new_words)} new words added." if new_words else "No new words found.")


# ===============================
# 🔍 Search Word Logic
# ===============================
def search_word(entry_var):
    word = entry_var.get().strip()
    if not word:
        messagebox.showwarning("Empty", "Please type a word to search.")
        return

    wb, ws = init_workbook()
    cleaned_search = clean_word_for_compare(word)
    found_row = None

    # Search word in Excel (column 2: New Words)
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row[1]:
            continue
        if clean_word_for_compare(str(row[1])) == cleaned_search:
            found_row = row  # [No., New Words, Sentence, Explanation, Date/Time]
            break

    if found_row:
        msg = (
            f"✅ '{word}' found in your Vocabulary!\n\n"
            f"🕓 Added on: {found_row[4]}\n"
            f"📖 Sentence: {found_row[2]}\n"
            f"💬 Explanation: {found_row[3]}"
        )
        messagebox.showinfo("Word Found", msg)
    else:
        messagebox.showinfo("Not Found", f"❌ '{word}' is new and not yet in your list.")


# ===============================
# 📂 Open Excel and DOC
# ===============================
def open_excel():
    if os.path.exists(EXCEL_FILE):
        subprocess.Popen(["start", EXCEL_FILE], shell=True)
    else:
        messagebox.showwarning("Not found", "Excel file not found.")

def open_doc():
    if os.path.exists(DOC_FILE):
        subprocess.Popen(["start", DOC_FILE], shell=True)
    else:
        messagebox.showwarning("Not found", "Doc file not found.")


# ===============================
# 🪟 GUI Setup
# ===============================
def start_app():
    root = tk.Tk()
    root.title("LingoBaby")
    root.geometry("680x250")
    set_window_icon(root)

    tk.Label(root, text="Enter a sentence or a word:").pack(pady=6)
    entry_var = tk.StringVar()
    entry = tk.Entry(root, textvariable=entry_var, width=90)
    entry.pack(pady=4)

    btn_frame = tk.Frame(root)
    btn_frame.pack(pady=10)

    tk.Button(btn_frame, text="Add Sentence", width=18,
              command=lambda: add_sentence_gui(entry_var)).grid(row=0, column=0, padx=6)
    tk.Button(btn_frame, text="Search Word", width=18,
              command=lambda: search_word(entry_var)).grid(row=0, column=1, padx=6)
    tk.Button(btn_frame, text="📊 View Excel", width=18, command=open_excel).grid(row=2, column=0, padx=6)
    tk.Button(btn_frame, text="📝 View Doc", width=18, command=open_doc).grid(row=2, column=1, padx=6)

    tk.Label(
        root,
        text="Tip: Type a sentence to add new words, or type a single word to search it.",
        fg="gray"
    ).pack(pady=8)

    root.mainloop()

if __name__ == "__main__":
    start_app()
