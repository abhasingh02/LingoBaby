# main.py
import tkinter as tk
from tkinter import messagebox
import re

from helpers import set_window_icon, clean_word_for_compare, close_excel_if_open, close_word_if_open
from excel_handler import init_workbook, load_existing_words_set, append_rows_for_sentence, save_workbook_safe
from word_handler import init_document, get_next_sentence_number, append_sentence, save_document
from config import ICON_FILE, EXCEL_FILE, DOC_FILE

# GUI logic
def add_sentence_gui(entry_var):
    sentence = entry_var.get().strip()
    if not sentence:
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
    messagebox.showinfo("Saved", f"Saved. {len(new_words)} new words added.")

def open_excel():
    import subprocess, os
    if os.path.exists(EXCEL_FILE):
        subprocess.Popen(["start", EXCEL_FILE], shell=True)
    else:
        messagebox.showwarning("Not found", "Excel file not found.")

def open_doc():
    import subprocess, os
    if os.path.exists(DOC_FILE):
        subprocess.Popen(["start", DOC_FILE], shell=True)
    else:
        messagebox.showwarning("Not found", "Doc file not found.")

def start_app():
    root = tk.Tk()
    root.title("Highlighted Notes Tracker")
    root.geometry("650x220")
    set_window_icon(root)

    tk.Label(root, text="Type your sentence (press Add Sentence):").pack(pady=6)
    entry_var = tk.StringVar()
    entry = tk.Entry(root, textvariable=entry_var, width=90)
    entry.pack(pady=4)

    btn_frame = tk.Frame(root)
    btn_frame.pack(pady=8)

    add_btn = tk.Button(btn_frame, text="Add Sentence", width=18,
                        command=lambda: add_sentence_gui(entry_var))
    add_btn.grid(row=0, column=0, padx=6)

    view_xl_btn = tk.Button(btn_frame, text="📊 View Excel", width=18, command=open_excel)
    view_xl_btn.grid(row=0, column=1, padx=6)

    view_doc_btn = tk.Button(btn_frame, text="📝 View Doc", width=18, command=open_doc)
    view_doc_btn.grid(row=0, column=2, padx=6)

    root.mainloop()

if __name__ == "__main__":
    start_app()
