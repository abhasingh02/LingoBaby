"""
✅ Detects irregular verbs using NLTK + heuristics
✅ Skips base (1st-form) verbs
✅ Adds new irregulars to irregular_verbs_extended.json only if needed
✅ Safe for PyInstaller
"""

import os, json, re
from nltk.stem import WordNetLemmatizer
from nltk.corpus import wordnet

# Auto-download WordNet (safe for first-time users)
import nltk
nltk.download("wordnet", quiet=True)

JSON_PATH = "irregular_verbs_extended.json"
lemmatizer = WordNetLemmatizer()


# -----------------------------
# 🔤 Lemmatization & Forms
# -----------------------------
def lemma(word: str) -> str:
    """Return verb lemma (base form)."""
    return lemmatizer.lemmatize(word.lower(), "v")


def conjugate(base: str, form: str) -> str:
    """Simple heuristic conjugator (regular rules)."""
    base = base.lower()
    if form == "past":
        if base.endswith("e"):
            return base + "d"
        if base.endswith("y") and base[-2] not in "aeiou":
            return base[:-1] + "ied"
        if re.match(r".*[aeiou][^aeiou]$", base):  # short-vowel CVC doubling
            return base + base[-1] + "ed"
        return base + "ed"
    if form == "participle":
        return conjugate(base, "past")
    if form == "ing":
        if base.endswith("ie"):
            return base[:-2] + "ying"
        if base.endswith("e") and base not in ("be", "see"):
            return base[:-1] + "ing"
        if re.match(r".*[aeiou][^aeiou]$", base):
            return base + base[-1] + "ing"
        return base + "ing"
    if form == "s":
        if base.endswith("y") and base[-2] not in "aeiou":
            return base[:-1] + "ies"
        if base.endswith(("s", "sh", "ch", "x", "z")):
            return base + "es"
        return base + "s"
    return base


# -----------------------------
# ⚙️ Irregular Check
# -----------------------------
def is_inflected(word: str) -> bool:
    """True if word isn't its base form."""
    return word.lower() != lemma(word)


def detect_irregular(word: str):
    """Return (is_irregular, forms_dict)."""
    word = word.lower().strip()
    if not is_inflected(word):
        print(f"⏩ Skipped base form: {word}")
        return False, None

    base = lemma(word)
    past = conjugate(base, "past")
    part = conjugate(base, "participle")
    ing = conjugate(base, "ing")
    s_form = conjugate(base, "s")

    expected_regulars = {past, part}
    irregular = word not in expected_regulars and wordnet.synsets(base, pos="v")

    forms = {
        "base": base,
        "past": past,
        "past_participle": part,
        "ing": ing,
        "s": s_form,
    }
    return irregular, forms


# -----------------------------
# 💾 JSON Helpers
# -----------------------------
def load_json(path=JSON_PATH):
    """Load JSON safely."""
    if not os.path.exists(path):
        return []
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception as e:
        print(f"⚠️ Error reading JSON: {e}")
        return []


def save_json(data, path=JSON_PATH):
    """Save JSON with indentation."""
    try:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=4, ensure_ascii=False)
    except Exception as e:
        print(f"⚠️ Failed to save JSON: {e}")


# -----------------------------
# 🧠 Main Entry
# -----------------------------
def add_if_irregular(word: str):
    """Check if irregular, then update JSON only if needed."""
    is_irreg, forms = detect_irregular(word)
    if not is_irreg:
        print(f"➡️ '{word}' is regular or base; skipped.")
        return False

    data = load_json()
    existing = {v["base"].lower(): v for v in data}
    base = forms["base"]

    if base in existing:
        print(f"✅ '{base}' already in JSON (id={existing[base]['id']}).")
        return True

    new_id = max((v.get("id", 0) for v in data), default=0) + 1
    new_entry = {"id": new_id, **forms}
    data.append(new_entry)
    save_json(data)
    print(f"🆕 Added irregular verb: {new_entry}")
    return True


# -----------------------------
# 🧪 Test
# -----------------------------
if __name__ == "__main__":
    test_words = ["went", "cried", "fought", "played", "sang"]
    for w in test_words:
        print("----")
        add_if_irregular(w)
