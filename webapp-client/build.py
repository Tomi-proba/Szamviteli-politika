#!/usr/bin/env python3
"""
Összeállítja a teljesen kliens-oldali (böngészőben futó) két oldalt:

  - ugyfel_source.html -> dist_ugyfel.html  : az ügyfélnek KÜLDENDŐ kérdőív.
    Csak az ügyféltől kérdezendő mezőket tartalmazza, a végén egy kódot ad,
    amit az ügyfél e-mailben visszaküld - NEM generál dokumentumot, nem
    tartalmaz semmi utalást az irodai/automatizálási folyamatra.

  - iroda_source.html -> dist_iroda.html    : a BELSŐ, jelszóval védett
    eszköz. Az ügyféltől kapott kódot beillesztve előtölti a válaszokat,
    az iroda kitölti a fennmaradó (szakmai minősítést igénylő) mezőket,
    átnézi, majd legenerálja a Word-vázlatot.

    FONTOS: a jelszó csak visszatartás (statikus oldal, nincs szerver) - ld.
    hash_password.py a cseréjéhez.

Mindkettő ugyanazt a zip.js + docx.js + resolve.js motort és
data/questions.json kérdésbankot használja, a sablon/merged.docx-ot égeti
bele base64-ként (csak az iroda-oldalnak van rá szüksége, de egyszerűbb
mindkettőbe ugyanúgy behelyettesíteni).

Futtatás:
    python3 webapp-client/build.py
"""
import base64
import json
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent
CLIENT = Path(__file__).resolve().parent


def build_one(source_name, out_name, questions_json, template_b64, engine):
    src = (CLIENT / source_name).read_text(encoding="utf-8")
    src = src.replace("__QUESTIONS_JSON__", questions_json)
    src = src.replace("__TEMPLATE_B64__", template_b64)
    src = src.replace("__ENGINE_JS__", engine)
    out_path = CLIENT / out_name
    out_path.write_text(src, encoding="utf-8")
    print(f"Elkészült: {out_path} ({len(src)} byte)")


def main():
    questions = json.loads((BASE / "data" / "questions.json").read_text(encoding="utf-8"))["kerdesek"]
    questions_json = json.dumps(questions, ensure_ascii=False)

    template_bytes = (BASE / "sablon" / "merged.docx").read_bytes()
    template_b64 = base64.b64encode(template_bytes).decode("ascii")

    engine = "\n".join(
        (CLIENT / name).read_text(encoding="utf-8") for name in ("zip.js", "docx.js", "resolve.js")
    )

    build_one("ugyfel_source.html", "dist_ugyfel.html", questions_json, template_b64, engine)
    build_one("iroda_source.html", "dist_iroda.html", questions_json, template_b64, engine)


if __name__ == "__main__":
    main()
