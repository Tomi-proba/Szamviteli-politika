#!/usr/bin/env python3
"""
Összeállítja a teljesen kliens-oldali (böngészőben futó) vázlatgeneráló
oldalt egyetlen önálló HTML fájlba: beleégeti a data/questions.json
kérdésbankot és a sablon/merged.docx tartalmát base64-ként, majd hozzáfűzi
a zip.js + docx.js + resolve.js motort.

Ez NEM a scripts/generate_policy.py-t vagy scripts/webapp.py-t helyettesíti
(azok Python-alapú, helyi futtatást igényelnek) - ez egy harmadik, teljesen
önálló változat, ami böngészőben fut, szerver/Python nélkül, publikálható
weboldalként.

Futtatás:
    python3 webapp-client/build.py
Kimenet:
    webapp-client/dist.html - ezt kell feltölteni/publikálni.
"""
import json
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent
CLIENT = Path(__file__).resolve().parent


def main():
    src = (CLIENT / "page_source.html").read_text(encoding="utf-8")
    questions = json.loads((BASE / "data" / "questions.json").read_text(encoding="utf-8"))["kerdesek"]

    import base64
    template_bytes = (BASE / "sablon" / "merged.docx").read_bytes()
    template_b64 = base64.b64encode(template_bytes).decode("ascii")

    engine = "\n".join(
        (CLIENT / name).read_text(encoding="utf-8") for name in ("zip.js", "docx.js", "resolve.js")
    )

    src = src.replace("__QUESTIONS_JSON__", json.dumps(questions, ensure_ascii=False))
    src = src.replace("__TEMPLATE_B64__", template_b64)
    src = src.replace("__ENGINE_JS__", engine)

    out_path = CLIENT / "dist.html"
    out_path.write_text(src, encoding="utf-8")
    print(f"Elkészült: {out_path} ({len(src)} byte)")


if __name__ == "__main__":
    main()
