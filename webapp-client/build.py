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

A működő JS-kód nagy része egy "application/octet-stream" típusú (tehát a
böngésző által automatikusan nem futtatott) <script id="core"> blokkban van
a *_source.html fájlokban. Ez a build lépés az ilyen blokk tartalmát
elhomályosítja (ld. obfuscate.mjs, igényli: `npm install` a webapp-client
mappában), majd kiszámolja a végeredmény SHA-256 lenyomatát, és azt a
__CORE_HASH__ helyére illeszti be - ezt egy másik, kis "bootstrap" script
ellenőrzi futás közben, mielőtt a core-blokkot valódi <script>-ként
beillesztve lefuttatná. Ld. "Integritás-ellenőrzés" a README-ben.

Futtatás:
    python3 webapp-client/build.py
"""
import base64
import hashlib
import json
import re
import subprocess
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent
CLIENT = Path(__file__).resolve().parent

CORE_BLOCK_RE = re.compile(
    r'(<script type="application/octet-stream" id="core">\n)(.*?)(\n</script>)',
    re.DOTALL,
)


def obfuscate_js(code):
    result = subprocess.run(
        ["node", str(CLIENT / "obfuscate.mjs")],
        input=code.encode("utf-8"),
        capture_output=True,
    )
    if result.returncode != 0:
        raise RuntimeError(f"obfuscate.mjs hiba:\n{result.stderr.decode('utf-8', 'replace')}")
    return result.stdout.decode("utf-8")


def obfuscate_core_and_inject_hash(src):
    match = CORE_BLOCK_RE.search(src)
    if not match:
        raise RuntimeError("Nem található a <script id=\"core\"> blokk.")
    obfuscated = obfuscate_js(match.group(2))
    # A hash a <script id="core"> elem TELJES textContent-jét fedi, tehát a
    # nyitó-/záró-tag melletti sortörést is (ld. match.group(1)/(3) végén a
    # "\n"-t) - ennek pontosan egyeznie kell a futásidejű JS oldali
    # számítással (checkIntegrity), különben hamis negatív (letiltás) lenne.
    core_text_content = "\n" + obfuscated + "\n"
    digest = hashlib.sha256(core_text_content.replace("\r\n", "\n").encode("utf-8")).hexdigest()
    src = src[: match.start()] + match.group(1) + obfuscated + match.group(3) + src[match.end() :]
    return src.replace("__CORE_HASH__", digest)


def build_one(source_name, out_name, questions_json, template_b64, engine):
    src = (CLIENT / source_name).read_text(encoding="utf-8")
    src = src.replace("__QUESTIONS_JSON__", questions_json)
    src = src.replace("__TEMPLATE_B64__", template_b64)
    src = src.replace("__ENGINE_JS__", engine)
    src = obfuscate_core_and_inject_hash(src)
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
