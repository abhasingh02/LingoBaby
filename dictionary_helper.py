# dictionary_helper.py
import os
os.environ["TYPEGUARD_DISABLE"] = "1"  # 👈 Disable typeguard for PyInstaller build

import inflect
from nltk.corpus import wordnet
import nltk

nltk.download('wordnet', quiet=True)

p = inflect.engine()

def get_explanation(word: str) -> str:
    word = word.lower().strip()
    explanation_parts = []

    plural = p.plural(word)
    if plural and plural != word:
        explanation_parts.append(f"Plural: {plural}")

    # Simple verb forms
    ing_form = word if word.endswith("ing") else word + "ing"
    past_form = word if word.endswith("ed") else word + "ed"
    third_form = word if word.endswith("s") else word + "s"
    explanation_parts.append(f"Verb forms: {word}, {ing_form}, {past_form}, {third_form}")

    synsets = wordnet.synsets(word)
    if synsets:
        explanation_parts.append(f"Meaning: {synsets[0].definition()}")

    return " | ".join(explanation_parts)
