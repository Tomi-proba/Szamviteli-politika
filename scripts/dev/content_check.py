#!/usr/bin/env python3
"""
Tartalmi ellenőrzés: a legenerált vázlatban tényleg a helyes ág maradt bent,
a behelyettesített válaszok a sablon "kitöltendő hely" jelölésével (sárga
kiemelés + áthúzás + piros betű) szerepelnek, és - ami a legfontosabb - a
sablon SEMMILYEN belső jelölése, szerkesztői utasítása vagy más cégre írt
konkrét adata nem szivárgott ki a kimenetbe.

A paritás-teszt (parity_check.py) csak azt garantálja, hogy a két motor
UGYANAZT adja; ez a teszt azt, hogy a közös kimenet HELYES is.

Futtatás:
    python3 scripts/dev/content_check.py [answers1.json ...]
"""
import json
import sys
import tempfile
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(BASE / "scripts"))

import docx  # noqa: E402

import generate_policy  # noqa: E402

MARK = ("YELLOW (7)", True, "C00000")  # a mark_substituted() jelölése

# Ezeknek SOHA, EGYETLEN válaszkészletnél sem szabad a kimenetben lenniük.
#
#  - a sablon saját jelölői és formázatlan helykitöltői (ügyfélhez sosem
#    mehetnek ki),
#  - a sablon szerzőjének feltételes szerkesztői utasításai,
#  - és a ZalaElektro-vázlatból örökölt, MÁS cégre vonatkozó konkrét adatok
#    (ezek bizonyítják, hogy a D1 mondat-újraírás tényleg megtörtént, és nem
#    maradtak véletlen konstansként a dokumentumban).
TILTOTT = [
    "▶ Adatbékérő",
    "Adatbékérő",
    "innen hiányzik",
    "ezt csak akkor írja",
    "ezt akkor írja csak",
    "ez csak akkor legyen itt",
    "NE ÍRJA",
    "Zala-Elektro",
    "Business Central",
    "Tábor Ildikó",
    "174234",
    "meghalódó",              # a sablon elgépelése, javítva kell legyen
    "A Vállalkozásnak nyilatkoznia kell",   # olvasónak szóló felszólítás
    "A Vállalkozásnak rögzítenie kell",
    "A sablon alapértelmezése",
]


def _is_marked(r):
    color = None
    if r.font.color is not None and r.font.color.type is not None:
        color = str(r.font.color.rgb)
    return (str(r.font.highlight_color), bool(r.font.strike), color) == MARK


def load(answers_path):
    """A legenerált vázlat bekezdés-szövegei és a behelyettesítettként
    jelölt futások."""
    out = Path(tempfile.mkdtemp()) / "o.docx"
    generate_policy.generate(answers_path, out)
    doc = docx.Document(str(out))
    paras = [p.text for p in doc.paragraphs]
    marked = [r.text for p in doc.paragraphs for r in p.runs if _is_marked(r)]
    return paras, marked


def check(answers_path, must_contain=(), must_not_contain=(), must_be_marked=()):
    paras, marked = load(answers_path)
    body = "\n".join(paras)
    name = Path(answers_path).name
    fails = []

    # --- minden válaszkészletre kötelező biztonsági invariánsok ---
    for s in TILTOTT:
        if s in body:
            fails.append(f"KISZIVÁRGOTT a kimenetbe: {s!r}")

    for s in must_contain:
        if s not in body:
            fails.append(f"HIÁNYZIK a kimenetből: {s!r}")
    for s in must_not_contain:
        if s in body:
            fails.append(f"BENNMARADT, pedig törölni kellett volna: {s!r}")
    for s in must_be_marked:
        if s not in marked:
            fails.append(f"NEM behelyettesítettként jelölt (sárga/áthúzott/piros): {s!r}")

    flags = sum(1 for t in paras if "⚠ ELLENŐRIZENDŐ" in t)
    print(f"  {'HIBA' if fails else 'OK  '}  {name}: {len(paras)} bekezdés, "
          f"{len(marked)} behelyettesített run, {flags} ⚠ jelölés")
    for f in fails:
        print(f"        - {f}")
    return not fails


def main():
    extra = [Path(x) for x in sys.argv[1:]]
    a1 = BASE / "data" / "answers.example.json"
    a2 = BASE / "data" / "answers.example2.json"
    ans1 = json.loads(a1.read_text(encoding="utf-8"))
    ans2 = json.loads(a2.read_text(encoding="utf-8"))

    ok = True

    # ---- 1. minta ----
    ok &= check(
        a1,
        must_contain=[
            # egyszerű kitöltések
            ans1["szv1"], ans1["szv2"], ans1["szv6"], ans1["cegjegyzekszam"],
            ans1["gazdalkodasi_forma"], ans1["alapitas_idopontja"],
            ans1["fioktelepek"], ans1["szv7"], ans1["szv8"],
            ans1["beszamolo_alairoja"], ans1["konyvelo_program"],
            ans1["evkozi_zarasok_feladatai"],
            # D1 - a válaszból ÚJRAÍRT mondatok
            "Vállalkozásunk esetében a könyvvezetés a KulcsSoft Számla és "
            "Könyvelő Program segítségével történik.",
            "A könyvvezetésért felelős: Nagy Éva (mérlegképes könyvelő, "
            "regisztrációs száma: 123456).",
            "Mindezek alapján a vállalkozás éves beszámolót készít.",
            "1. sz. melléklet „A” változata szerint készítjük el",
            "módszerek közül az összköltség eljárással készülő",
            "csak az 5. számlaosztályban, költségnemenként számolja el",
            "A jövőbeni költségekre a Vállalkozás nem képez céltartalékot",
            "Az üzleti vagy cégértékkel a Vállalkozás nem rendelkezik.",
            # D2 - a feltételes mondatok helyes ága
            "a Vállalkozásnak jogszabályi kötelezettsége nem áll fenn",
            "Jelenleg a Vállalkozás nem minősül",
            "az arab számmal jelzett tételeket egy-egy csoporton belül nem vonjuk össze",
            # "innen hiányzik" helykitöltők kitöltve
            "a mérleg (beszámoló) forduló napja: december hónap 31. napja",
            "a mérleg (beszámoló) készítésének napja: a tárgyévet követő év "
            "március hónap 31. napja",
            "a beszámoló átadása könyvvizsgálatra: a tárgyévet követő év április 30.",
            "a beszámoló letétbe helyezése: a mérlegfordulónapot követő "
            "ötödik hónap utolsó napja",
            "a beszámoló közzététele: a tárgyévet követő év május 31.",
            "ellenőrzése a gazdasági vezető feladata",
            "A beszámoló közzétételéért felelős személy: a gazdasági vezető",
            # ugyanaz a tény két helyen - nem mondhat mást
            "Társaságunknál a beszámolót a mérlegfordulónapot követő ötödik "
            "hónap utolsó napjáig kell letétbe helyezni.",
            "fordulónapja december hónap 31. napja",
            # écs-küszöbök: a lefedett kategóriák felülírva, a többi 20%
            "Immateriális javak: 15%  |  Tárgyi eszközök: 20%  |  "
            "Befektetett pénzügyi eszközök: 25%  |  Készletek: 30%",
        ],
        must_not_contain=[
            "„B” változata",
            "forgalmi költség eljárással",
            "elsődlegesen az 5. számlaosztályban",
            "az anyavállalat által készítendő összevont",
            # Szv#106: mentesül -> az önköltségszámítási szabályzat MINDKÉT
            # felsorolásból kiesik
            "Önköltségszámítási szabályzat",
            "az Önköltség-számítás rendjére vonatkozó szabályzat",
        ],
        must_be_marked=[
            ans1["beszamolo_alairoja"],
            ans1["szv1"],
            ans1["cegjegyzekszam"],
            " éves beszámolót ",   # az inline (sormintás) behelyettesítés
        ],
    )

    # ---- 2. minta: az ELLENTÉTES ágak ----
    ok &= check(
        a2,
        must_contain=[
            ans2["szv1"], ans2["konyvvizsgalo_adatai"],
            ans2["konyvvizsgalo_nyilvantartasi_szam"],
            ans2["konyvvizsgalatert_felelos_konyvvizsgalo"],
            ans2["ertekelesi_delegalas"],
            ans2["konszolidacio_reszletszabalyok"],
            ans2["beszamolo_honlap_cim"],
            ans2["celtartalek_jelentos_hatar"],
            ans2["bizomanyi_dij_hatar"],
            "Mindezek alapján a vállalkozás egyszerűsített éves beszámolót készít.",
            "1. sz. melléklet „B” változata szerint készítjük el",
            "módszerek közül a forgalmi költség eljárással készülő",
            "elsődlegesen az 5. számlaosztályban költségnemenként",
            "A Vállalkozás a Számviteli törvény szerint anyavállalatnak minősül "
            "és összevont (konszolidált) éves beszámolót készít.",
            "A jövőbeni költségekre a Vállalkozás céltartalékot képez",
            "A vállalkozás üzleti vagy cégértéket mutat ki.",
            "A Vállalkozás üzleti vagy cégértéket mutat ki, amelyet 8 év alatt ír le.",
            "A negatív üzleti vagy cégértéket a Vállalkozás 4 év alatt írja le.",
            "internetes honlapon is elérhető",
            "Vállalkozásunknál az immateriális javakkal kapcsolatos "
            "értékhelyesbítési lehetőséget alkalmazzuk.",
            # a letétbe helyezés a szabályozott piaci ág szerint
            "negyedik hónap utolsó napjáig kell letétbe helyezni",
            "a beszámoló letétbe helyezése: a mérlegfordulónapot követő "
            "negyedik hónap utolsó napja",
            # a mérlegkészítés még a TÁRGYÉVBEN van (június 30. + 60 nap)
            "a tárgyév augusztus hónap 29. napját",
            "ellenőrzése az ügyvezető feladata",
            # Szv#106: van szabályzat -> mindkét felsorolásban benne marad
            "Önköltségszámítási szabályzat",
            "az Önköltség-számítás rendjére vonatkozó szabályzat",
        ],
        must_not_contain=[
            "„A” változata",
            "összköltség eljárással készülő",
            "Jelenleg a Vállalkozás nem minősül",
            "a vállalkozás székhelyén megtekinthető.\n",  # a "nincs honlap" ág
            "csak az 5. számlaosztályban, költségnemenként számolja el",
            "A jövőbeni költségekre a Vállalkozás nem képez céltartalékot",
        ],
        must_be_marked=[
            ans2["szv1"],
            ans2["konyvvizsgalo_adatai"],
            "A beszámoló a vállalkozás székhelyén megtekinthető, továbbá "
            f"a(z) {ans2['beszamolo_honlap_cim']} internetes honlapon is elérhető.",
        ],
    )

    # ---- további (pl. adverzariális) készletek: csak a kötelező invariánsok ----
    for path in extra:
        ok &= check(path)

    print("TARTALMI ELLENŐRZÉS: RENDBEN" if ok else "TARTALMI ELLENŐRZÉS: HIBA")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
