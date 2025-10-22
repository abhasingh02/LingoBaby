# helpers.py
import re
import subprocess
import os
from config import ICON_FILE

def clean_word_for_compare(word: str) -> str:
    """
    Return cleaned lowercased word for comparison.
    Returns empty string if word is bracketed like [abc].
    """
    word = word.strip()
    if re.match(r"^\[.*\]$", word):
        return ""
    # remove .,?,! and other punctuation
    word = re.sub(r"[.?!]", "", word)
    word = re.sub(r"[^\w\s]", "", word)
    return word.lower().strip()

def close_excel_if_open():
    """Close Excel on Windows using taskkill (safe ignore errors)."""
    try:
        # Kill all excel instances so file locks are released
        subprocess.call(["taskkill", "/f", "/im", "EXCEL.EXE"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except Exception:
        pass

def close_word_if_open():
    """Close Word on Windows using taskkill (safe ignore errors)."""
    try:
        subprocess.call(["taskkill", "/f", "/im", "WINWORD.EXE"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except Exception:
        pass

def set_window_icon(root):
    """Set tkinter window icon from ICON_FILE if exists."""
    from config import ICON_FILE
    if os.path.exists(ICON_FILE):
        try:
            root.iconbitmap(ICON_FILE)
        except Exception:
            # Some systems may not support ico or icon may be invalid
            pass
