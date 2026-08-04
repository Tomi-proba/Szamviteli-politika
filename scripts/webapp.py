#!/usr/bin/env python3
"""
Egyszerű helyi webes űrlap a Számviteli politika vázlatgenerátorhoz.

Nincs külső függősége (csak a Python standard library-t használja), ezért
telepítés nélkül, bárhol elindítható:

    python3 scripts/webapp.py [port]   # alapértelmezett port: 8000

Utána a böngészőben: http://localhost:8000

A form a data/questions.json alapján épül fel. Beküldéskor a válaszokból
összeáll egy answers.json-nak megfelelő szótár, lefut rajta a
generate_policy.generate(), és a böngésző letölti a kész .docx vázlatot.

Ez egy BELSŐ (irodai) eszköznek készült: a form mindkét kérdéstípust
(ügyfélnek szánt és irodai/szakmai minősítést igénylő kérdést) egy oldalon
mutatja - a kollégák tudják, melyiket kell az ügyféltől bekérni, és
melyiket töltik ki ők maguk.
"""
import html
import json
import sys
import traceback
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs

BASE = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(BASE / "scripts"))
import generate_policy as gp  # noqa: E402

QUESTIONS = json.loads((BASE / "data" / "questions.json").read_text(encoding="utf-8"))["kerdesek"]


def humanize(value):
    return value.replace("_", " ")


def render_field(q):
    kulcs = html.escape(q["kulcs"])
    kerdes = html.escape(q["kerdes"])
    forras_label = "ügyfél" if q["forras"] == "ugyfel" else "iroda"
    forras_class = "forras-ugyfel" if q["forras"] == "ugyfel" else "forras-iroda"

    attrs = ""
    if "csak_ha" in q:
        (dep_key, dep_values), = q["csak_ha"].items()
        attrs = (
            f' data-depends-on="{html.escape(dep_key)}"'
            f' data-depends-values="{html.escape(",".join(dep_values))}"'
        )

    if q["tipus"] == "select":
        opts = "".join(
            f'<option value="{html.escape(o)}">{html.escape(humanize(o))}</option>'
            for o in q["opciok"]
        )
        input_html = f'<select name="{kulcs}"><option value="">-- válasszon --</option>{opts}</select>'
    elif q["tipus"] == "number":
        input_html = f'<input type="text" inputmode="numeric" name="{kulcs}">'
    else:
        input_html = f'<input type="text" name="{kulcs}">'

    return f"""
    <div class="field {forras_class}"{attrs}>
      <label>{kerdes} <span class="forras-tag">[{forras_label}]</span></label>
      {input_html}
    </div>"""


def render_form():
    fields_html = "".join(render_field(q) for q in QUESTIONS)
    return f"""<!doctype html>
<html lang="hu">
<head>
<meta charset="utf-8">
<title>Számviteli politika - vázlatgenerátor</title>
<style>
  body {{ font-family: -apple-system, Segoe UI, Arial, sans-serif; max-width: 820px;
         margin: 2rem auto; padding: 0 1rem; color: #222; }}
  h1 {{ font-size: 1.4rem; }}
  .field {{ margin-bottom: 1rem; padding: 0.6rem 0.8rem; border-radius: 6px; }}
  .forras-ugyfel {{ background: #eef6ff; }}
  .forras-iroda {{ background: #fff6e6; }}
  label {{ display: block; font-weight: 600; margin-bottom: 0.3rem; font-size: 0.92rem; }}
  .forras-tag {{ font-weight: 400; color: #666; font-size: 0.8rem; }}
  input, select {{ width: 100%; padding: 0.4rem; box-sizing: border-box;
                   font-size: 0.95rem; border: 1px solid #bbb; border-radius: 4px; }}
  .legend {{ display: flex; gap: 1.5rem; margin: 1rem 0; font-size: 0.9rem; }}
  .legend span {{ padding: 0.15rem 0.5rem; border-radius: 4px; }}
  button {{ margin-top: 1.5rem; padding: 0.7rem 1.4rem; font-size: 1rem;
            background: #1a5fb4; color: white; border: none; border-radius: 6px;
            cursor: pointer; }}
  button:hover {{ background: #164a8f; }}
  .hidden {{ display: none; }}
</style>
</head>
<body>
  <h1>Számviteli politika - automatikus vázlatgenerátor</h1>
  <p>Töltsd ki az alábbi mezőket (a kék hátterű kérdéseket az ügyféltől kell
     bekérni, a sárga hátterűeket az iroda tölti ki), majd kattints a
     "Vázlat legenerálása" gombra. A letöltött Word-fájl egy VÁZLAT -
     mielőtt kimegy az ügyfélnek, egy kollégának át kell néznie.</p>
  <div class="legend">
    <span class="forras-ugyfel">kék = ügyféltől kérdezendő</span>
    <span class="forras-iroda">sárga = iroda tölti ki</span>
  </div>
  <form method="POST" action="/generate">
    {fields_html}
    <button type="submit">Vázlat legenerálása</button>
  </form>
  <script>
    function applyVisibility() {{
      document.querySelectorAll('[data-depends-on]').forEach(function(el) {{
        var dep = el.getAttribute('data-depends-on');
        var values = el.getAttribute('data-depends-values').split(',');
        var ctrl = document.querySelector('[name="' + dep + '"]');
        var show = ctrl && values.indexOf(ctrl.value) !== -1;
        el.classList.toggle('hidden', !show);
      }});
    }}
    document.addEventListener('change', applyVisibility);
    applyVisibility();
  </script>
</body>
</html>"""


class Handler(BaseHTTPRequestHandler):
    def log_message(self, fmt, *args):
        pass  # csendesebb konzol

    def do_GET(self):
        if self.path != "/":
            self.send_response(404)
            self.end_headers()
            return
        body = render_form().encode("utf-8")
        self.send_response(200)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_POST(self):
        if self.path != "/generate":
            self.send_response(404)
            self.end_headers()
            return
        length = int(self.headers.get("Content-Length", 0))
        raw = self.rfile.read(length).decode("utf-8")
        posted = {k: v[0] for k, v in parse_qs(raw).items()}
        answers = {q["kulcs"]: posted.get(q["kulcs"], "") for q in QUESTIONS}

        tmp_out = BASE / "output" / "webform_generalt_vazlat.docx"
        tmp_out.parent.mkdir(parents=True, exist_ok=True)
        try:
            doc = gp.docx.Document(str(gp.TEMPLATE))
            p = doc.paragraphs
            gp.resolve_nyelv(p, answers)
            gp.resolve_penznem(p, answers)
            gp.resolve_merlegkeszites(p, answers)
            gp.resolve_fordulonap(p, answers)
            gp.resolve_alairo(p, answers)
            gp.resolve_konyvvizsgalat(p, answers)
            gp.resolve_ertekelesi_felelos(p, answers)
            gp.resolve_konyveles_felelos(p, answers)
            gp.resolve_nyilvanossagra_hozatal(p, answers)
            gp.resolve_konszolidacio(p, answers)
            gp.resolve_leltar(p, answers)
            gp.resolve_simple_fills(p, answers)
            gp.resolve_alapitas_atszervezes(p, answers)
            gp.resolve_cegertek(p, answers)
            gp.resolve_ertekhelyesbites(p, answers)
            gp.resolve_celtartalek(p, answers)
            gp.resolve_ert4(p, answers)
            gp.remove_jelolt_instructions(p)
            doc.save(str(tmp_out))
        except Exception:
            self.send_response(500)
            self.send_header("Content-Type", "text/plain; charset=utf-8")
            self.end_headers()
            self.wfile.write(traceback.format_exc().encode("utf-8"))
            return

        data = tmp_out.read_bytes()
        self.send_response(200)
        self.send_header(
            "Content-Type",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
        self.send_header("Content-Disposition", 'attachment; filename="szamviteli_politika_vazlat.docx"')
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)


def main():
    port = int(sys.argv[1]) if len(sys.argv) > 1 else 8000
    server = ThreadingHTTPServer(("0.0.0.0", port), Handler)
    print(f"Fut: http://localhost:{port}")
    server.serve_forever()


if __name__ == "__main__":
    main()
