#!/usr/bin/env python3
"""
Számviteli politika - automatikus vázlatgenerátor.

Beolvassa a JELÖLT sablont (sablon/szamviteli_politika_JELOLT.docx) és egy
kitöltött válaszkészletet (JSON, ld. data/answers.example.json), majd
legenerál egy tisztított Word-vázlatot:
  - kitölti a kérdőívben megválaszolt helyeket,
  - a VAGY-alternatívák közül a válaszok alapján kiválasztja a megfelelőt és
    törli a többit,
  - eltávolítja a [Szv N. kérdés] / [Ért N. kérdés] hivatkozásokat és a
    JELÖLT munkapéldány szerkesztői útmutatóját,
  - a fentiek szerint kezelt bekezdéseknél megszünteti a "még döntendő"
    kék/sárga/áthúzott jelölést, hogy jól látható legyen, mi már kész.

FONTOS: ez egy VÁZLATOT készít. A sablon számos, a kérdőívben NEM szereplő
döntési pontot (VAGY-alternatívát) is tartalmaz (pl. eszközértékelési
módszerek a VIII. fejezetben) - ezeket a script szándékosan nem nyúlja
hozzá, mert nincs hozzájuk ügyféladat. A munkatársnak ezeket továbbra is
kézzel kell átnéznie és eldöntenie, a dokumentum kiküldése előtt.
"""
import copy
import json
import re
import sys
from datetime import date, timedelta
from pathlib import Path

import docx

BASE = Path(__file__).resolve().parent.parent
TEMPLATE = BASE / "sablon" / "merged.docx"


# --------------------------------------------------------------------------
# Alacsony szintű run/bekezdés segédfüggvények
# --------------------------------------------------------------------------

def blank_groups(paragraph):
    """A bekezdés futásait (run) logikai 'kitöltendő hely' csoportokba
    rendezi. Egy csoport egy vagy több szomszédos, áthúzott (strike) futásból
    áll, köztük legfeljebb üres futásokkal - ezeket a Word run-daraboló
    viselkedése hozza létre, nem önálló helyek."""
    groups, current = [], []
    for i, r in enumerate(paragraph.runs):
        strike = bool(r.font.strike)
        empty_nonstrike = (not strike) and r.text.strip() == ""
        if strike or empty_nonstrike:
            current.append(i)
        else:
            if current and any(paragraph.runs[j].font.strike for j in current):
                groups.append(current)
            current = []
    if current and any(paragraph.runs[j].font.strike for j in current):
        groups.append(current)
    return groups


def strip_bracket_ref(paragraph):
    """A '[Szv N. kérdés]' / '[Ért N. kérdés]' hivatkozás-futásokat üríti."""
    for r in paragraph.runs:
        if r.font.highlight_color is not None and not r.font.strike and re.search(
            r"\[\s*(Szv|Ért)", r.text
        ):
            r.text = ""


def clear_choice_marks(paragraph):
    """A 'még döntendő' jelöléseket (sárga kiemelés, áthúzás, kék/piros
    betűszín) törli egy már véglegesített bekezdésben."""
    for r in paragraph.runs:
        r.font.highlight_color = None
        r.font.strike = False
        if r.font.color is not None and r.font.color.type is not None:
            r.font.color.rgb = None


def fill_blanks(paragraph, values):
    """Sorrendben kitölti a bekezdés kitöltendő helyeit a megadott
    értékekkel, majd normalizálja a formázást."""
    groups = blank_groups(paragraph)
    for grp, val in zip(groups, values):
        paragraph.runs[grp[0]].text = "" if val is None else str(val)
        for j in grp[1:]:
            paragraph.runs[j].text = ""
    strip_bracket_ref(paragraph)
    clear_choice_marks(paragraph)


def keep(paragraph):
    """Megtartott, de nem kitöltendő bekezdés véglegesítése (pl. egy
    VAGY-alternatíva statikus szövege)."""
    strip_bracket_ref(paragraph)
    clear_choice_marks(paragraph)


def delete(paragraph):
    p = paragraph._p
    p.getparent().remove(p)


def flag_before(paragraph, note):
    """Jól látható, félkövér piros felülvizsgálati jelölést szúr be az adott
    bekezdés elé - olyan döntési pontokhoz, amikhez nincs kérdőív-adat."""
    new_p = paragraph.insert_paragraph_before()
    run = new_p.add_run(f"⚠ ELLENŐRIZENDŐ (nincs kérdőív-adat): {note}")
    run.bold = True
    run.font.color.rgb = docx.shared.RGBColor(0xC0, 0x00, 0x00)


# --------------------------------------------------------------------------
# Segéd: hónapnév -> szám, illusztratív dátum a mérlegkészítési blankokhoz
# --------------------------------------------------------------------------

_MONTHS = {
    "január": 1, "február": 2, "március": 3, "április": 4, "május": 5,
    "június": 6, "július": 7, "augusztus": 8, "szeptember": 9,
    "október": 10, "november": 11, "december": 12,
}


def month_num(value):
    s = str(value).strip().lower()
    if s in _MONTHS:
        return _MONTHS[s]
    m = re.search(r"\d+", s)
    return int(m.group()) if m else 12


def illustrative_date(fordulonap_honap, fordulonap_nap, offset_days):
    """A fordulónap + offset_days illusztratív hónap/nap párja a
    mérlegkészítési bekezdésekhez (nem hivatalos határidő, csak a sablon
    zárójeles példájának kitöltése)."""
    try:
        base = date(2024, month_num(fordulonap_honap), int(fordulonap_nap))
        target = base + timedelta(days=int(offset_days))
        return target.month, target.day
    except Exception:
        return "…", "…"


def is_number(value):
    try:
        float(str(value).strip().replace(",", "."))
        return True
    except (TypeError, ValueError):
        return False


# --------------------------------------------------------------------------
# Fő generálási logika
# --------------------------------------------------------------------------

def remove_jelolt_instructions(p):
    """A JELÖLT munkapéldány szerkesztői útmutatóját (0-3. bekezdés) törli -
    ez sosem kerülhet ki az ügyfélhez. Ezt HÍVJUK ELŐSZÖR, mielőtt bármi
    mást törölnénk/kitöltenénk, különben a paragrafus-indexek elcsúsznak."""
    for i in (0, 1, 2, 3):
        delete(p[i])


def resolve_nyelv(p, a):
    if a.get("van_idegen_nyelvu_beszamolo") == "igen":
        fill_blanks(p[600], [a.get("idegen_nyelv", "")])
        delete(p[597])
        delete(p[599])
    else:
        keep(p[597])
        delete(p[599])
        delete(p[600])


def resolve_penznem(p, a):
    tipus = a.get("penznem_tipus", "forint")
    if tipus == "forint":
        keep(p[616])
        for i in (618, 620, 622, 624, 625, 627, 629, 631):
            delete(p[i])
    elif tipus in ("euro", "usd"):
        keep(p[620])
        if tipus == "euro":
            keep(p[622])
            delete(p[624])
            delete(p[625])
        else:
            delete(p[622])
            delete(p[624])
            keep(p[625])
        fill_blanks(p[627], [a.get("penznem_atteres_datum", "")])
        delete(p[616])
        delete(p[618])
        delete(p[629])
        delete(p[631])
    else:  # egyeb_deviza
        fill_blanks(p[631], [a.get("penznem_egyeb_neve", "")])
        for i in (616, 618, 620, 622, 624, 625, 627, 629):
            delete(p[i])
    keep(p[633])


def resolve_merlegkeszites(p, a):
    nap = a.get("merlegkeszites_nap", "")
    honap = a.get("fordulonap_honap", "")
    fnap = a.get("fordulonap_nap", "")
    ill_honap, ill_nap = illustrative_date(honap, fnap, nap if is_number(nap) else 0)
    if a.get("merlegkeszites_van_kivetel") != "igen":
        fill_blanks(p[662], [nap, ill_honap, ill_nap])
        delete(p[664])
        delete(p[666])
        delete(p[667])
        delete(p[668])
    else:
        fill_blanks(p[666], [nap, ill_honap, ill_nap])
        fill_blanks(p[667], [a.get("merlegkeszites_kivetelek", "")])
        delete(p[662])
        delete(p[664])
        delete(p[668])


def resolve_fordulonap(p, a):
    # 651 = statikus "december 31" alapértelmezés, 653 = VAGY, 655 = kitöltendő
    # egyedi hónap/nap - a kérdőívből mindig van konkrét válasz, ezért mindig
    # az egyedi (kitöltött) változatot tartjuk meg.
    fill_blanks(p[655], [a.get("fordulonap_honap", ""), a.get("fordulonap_nap", "")])
    delete(p[651])
    delete(p[653])


def resolve_alairo(p, a):
    fill_blanks(p[369], [a.get("beszamolo_alairoja", "")])


def resolve_konyvvizsgalat(p, a):
    tipus = a.get("konyvvizsgalat_tipus", "nincs_kotelezettseg")
    adatai = a.get("konyvvizsgalo_adatai", "")
    branches = {"kotelezo_altalanos": 405, "kotelezo_koztartozas_miatt": 409, "onkentes": 413}
    if tipus == "nincs_kotelezettseg":
        keep(p[401])
        for i in (405, 409, 413):
            delete(p[i])
    else:
        chosen = branches[tipus]
        fill_blanks(p[chosen], [adatai])
        for i in (401, 405, 409, 413):
            if i != chosen:
                delete(p[i])
    for i in (403, 407, 411):
        delete(p[i])

    fenn = a.get("fenntarthatosagi_jelentes", "nem_koteles")
    delete(p[415])
    if fenn == "nem_koteles":
        delete(p[417])
        delete(p[419])
        delete(p[421])
    elif fenn == "igen_ugyanaz_a_konyvvizsgalo":
        keep(p[417])
        delete(p[419])
        delete(p[421])
    else:
        delete(p[417])
        delete(p[419])
        keep(p[421])


def resolve_ertekelesi_felelos(p, a):
    if a.get("ertekeles_felelos_tipus") == "egyedi_delegalas":
        fill_blanks(p[338], [a.get("ertekelesi_delegalas", "")])
        delete(p[334])
        delete(p[336])
    else:
        keep(p[334])
        delete(p[336])
        delete(p[338])


def resolve_konyveles_felelos(p, a):
    tipus = a.get("konyveles_felelos_tipus", "mentesseg_20mft_arbevetel_alatt")
    nev = a.get("konyveles_felelos_nev_regszam", "")
    if tipus == "merlegkepes_konyvelo":
        fill_blanks(p[313], [nev])
        keep(p[312])
        delete(p[317])
        delete(p[318])
        delete(p[322])
    elif tipus == "oklev_konyvvizsgalo":
        fill_blanks(p[318], [nev])
        keep(p[317])
        delete(p[312])
        delete(p[313])
        delete(p[322])
    else:
        keep(p[322])
        delete(p[312])
        delete(p[313])
        delete(p[317])
        delete(p[318])
    delete(p[315])
    delete(p[320])
    keep(p[324])
    fill_blanks(p[964], [a.get("konyvvezetesert_felelos_szemely", "")])
    fill_blanks(p[962], [a.get("konyvelo_program", "")])


def resolve_nyilvanossagra_hozatal(p, a):
    keep(p[441])
    keep(p[443])
    if a.get("beszamolo_honlapon") == "igen":
        fill_blanks(p[445], [a.get("beszamolo_honlap_cim", "")])
        delete(p[449])
    else:
        keep(p[449])
        delete(p[445])
    delete(p[447])


def resolve_konszolidacio(p, a):
    delete(p[844])  # "VÁLASZTANI KELL:" - szerkesztői utasítás
    tipus = a.get("konszolidacio_tipus", "nem_erintett")
    reszlet = a.get("konszolidacio_reszletszabalyok", "")
    branch_map = {
        "leany_mentesul_bevonas_alol": ([850], 851),
        "leany_bevonva": ([855], 856),
        "anya_mentesul_folerendelt_anya_miatt": ([860], 861),
        "anya_mentesul_nagysagrend_alapjan": ([865, 868], 869),
        "anya_keszit_konszolidaltat": ([873, 874], 875),
    }
    all_idx = [846, 850, 851, 855, 856, 860, 861, 865, 868, 869, 873, 874, 875]
    if tipus == "nem_erintett":
        keep(p[846])
        for i in all_idx:
            if i != 846:
                delete(p[i])
    else:
        keep_static, fill_idx = branch_map[tipus]
        for i in keep_static:
            keep(p[i])
        fill_blanks(p[fill_idx], [reszlet])
        for i in all_idx:
            if i != fill_idx and i not in keep_static:
                delete(p[i])
    for i in (848, 853, 858, 863, 871):
        delete(p[i])


def resolve_leltar(p, a):
    tipus = a.get("leltar_felelos_tipus", "vezetes_vagy_szv_rendert_felelos")
    if tipus == "vezetes_vagy_szv_rendert_felelos":
        keep(p[342])
        delete(p[346])
        delete(p[350])
    elif tipus == "eszkoz_forras_felelosok":
        keep(p[346])
        delete(p[342])
        delete(p[350])
    else:
        fill_blanks(p[350], [a.get("leltar_delegalas_reszletek", "")])
        delete(p[342])
        delete(p[346])
    delete(p[344])
    delete(p[348])


def resolve_simple_fills(p, a):
    # 831 = "VÁLASZTANI KELL:" szerkesztői utasítás, 833 = kitöltendő változat,
    # 835 = VAGY, 837 = fix %-os alternatíva - a kérdőív mindig konkrét
    # összeghatárt kér, ezért mindig a kitöltött változatot tartjuk meg.
    fill_blanks(p[833], [a.get("jelentos_osszeg_bemutatas", "")])
    delete(p[831])
    delete(p[835])
    delete(p[837])

    # 916 = kitöltendő változat, 918 = VAGY, 920 = fix %-os alternatíva.
    fill_blanks(p[916], [a.get("jelentos_osszeg_ertelmezes", "")])
    delete(p[918])
    delete(p[920])

    fill_blanks(p[908], [a.get("jelentos_mertek_hanyad", "")])
    fill_blanks(p[937], [a.get("kiveteles_nagysag_osszeghatar", "")])
    fill_blanks(p[939], [a.get("kiveteles_nagysag_sajattoke_szazalek", "")])
    fill_blanks(p[941], [
        a.get("kiveteles_nagysag_bevetel_szazalek", ""),
        a.get("kiveteles_nagysag_koltseg_szazalek", ""),
    ])

    # 1054/1058 = két fix eljárási alternatíva, 1062 = kitöltendő, részletes
    # változat - a kérdőív ehhez kér konkrét választ, ezért ezt tartjuk meg.
    fill_blanks(p[1062], [a.get("bizonylat_konyveles_osszhang", "")])
    delete(p[1054])
    delete(p[1056])
    delete(p[1058])
    delete(p[1060])

    fill_blanks(p[1015], [a.get("bizonylatkezeles_szabalyai", "")])
    # 1153-nak KÉT kitöltendő helye van: egy listafolytatás ("...archiválás,
    # ....)") és az alábbiakban részletezendő tényleges válasz - az elsőt egy
    # semleges "stb."-vel zárjuk, mert a kérdőívnek csak a másodikra van
    # konkrét adata.
    fill_blanks(p[1153], ["stb.", a.get("evkozi_zarasok_feladatai", "")])
    fill_blanks(p[823], [a.get("kiegeszito_melleklet_sajatossagok", "")])


def resolve_alapitas_atszervezes(p, a):
    if a.get("alapitas_atszervezes_aktivalja") == "igen" and is_number(
        a.get("alapitas_atszervezes_leiras_ev")
    ):
        fill_blanks(p[1810], [a.get("alapitas_atszervezes_leiras_ev", "")])
        delete(p[1805])
        delete(p[1807])
        delete(p[1809])
    else:
        flag_before(
            p[1805],
            "az alapítás-átszervezés aktiválásáról/leírási idejéről szóló válasz "
            "nem egyértelmű vagy 'nem aktiválunk' - kérjük kézzel pontosítani, "
            "ill. ellenőrizni, hogy a kapcsolódó bekerülésiérték-szabály (VIII. "
            "fejezet) is összhangban van-e.",
        )


def resolve_cegertek(p, a):
    if a.get("van_cegertek") == "igen" and is_number(a.get("cegertek_leiras_ev")):
        fill_blanks(p[1815], [a.get("cegertek_leiras_ev", "")])
        delete(p[1812])
        delete(p[1814])
        delete(p[1817])
        delete(p[1818])
    else:
        flag_before(
            p[1812],
            "nincs üzleti/cégérték, vagy a leírási idő nem egyértelmű - "
            "kérjük kézzel törölni/pontosítani a nem releváns bekezdéseket.",
        )

    if a.get("van_cegertek") == "igen" and a.get("van_negativ_cegertek") == "igen" and is_number(
        a.get("negativ_cegertek_leiras_ev")
    ):
        fill_blanks(p[1699], [a.get("negativ_cegertek_leiras_ev", "")])
    else:
        flag_before(
            p[1699],
            "nincs negatív üzleti/cégérték, vagy a leírási idő nem egyértelmű - "
            "kérjük kézzel törölni/pontosítani.",
        )


def resolve_ertekhelyesbites(p, a):
    specs = [
        ("ertekhelyesbites_immat_van", "ertekhelyesbites_immat_javak", 1274, 1276, 1278),
        ("ertekhelyesbites_targyi_van", "ertekhelyesbites_targyi_eszkoz", 1299, 1301, 1303),
        ("ertekhelyesbites_penzugyi_van", "ertekhelyesbites_penzugyi_eszkoz", 1344, 1346, 1348),
    ]
    for gate_key, text_key, fill_idx, sep_idx, static_idx in specs:
        if a.get(gate_key) == "igen":
            fill_blanks(p[fill_idx], [a.get(text_key, "")])
            delete(p[static_idx])
        else:
            keep(p[static_idx])
            delete(p[fill_idx])
        delete(p[sep_idx])


def resolve_celtartalek(p, a):
    if a.get("celtartalek_kepez") == "igen":
        fill_blanks(p[1608], [""])
        delete(p[1612])
        fill_blanks(p[1614], [a.get("celtartalek_jelentos_hatar", "")])
        fill_blanks(p[1630], [a.get("celtartalek_elteres_szazalek", "")])
    else:
        keep(p[1612])
        delete(p[1608])
        delete(p[1614])
        flag_before(
            p[1630],
            "a vállalkozás nem képez céltartalékot lehetőség alapján - "
            "kérjük ellenőrizni, hogy ez a bekezdés (lényeges eltérés %-a) "
            "továbbra is szükséges-e, vagy törlendő.",
        )
    delete(p[1610])


def resolve_ert4(p, a):
    if a.get("ert4_alkalmazza_elhatarolas") == "igen":
        fill_blanks(p[1494], [a.get("bizomanyi_dij_hatar", "")])
        keep(p[1496])
        delete(p[1500])
    else:
        keep(p[1500])
        delete(p[1494])
        delete(p[1496])
    delete(p[1498])


def generate(answers_path, output_path):
    answers = json.loads(Path(answers_path).read_text(encoding="utf-8"))
    answers = {k: v for k, v in answers.items() if not k.startswith("_")}

    doc = docx.Document(str(TEMPLATE))
    # EGYETLEN alkalommal olvassuk ki a bekezdéslistát - minden lenti
    # resolve_* és a törlés is EBBE a listába indexel. A python-docx
    # Paragraph objektumok a mögöttes XML elemre mutatnak, ezért törlés
    # után is stabilak maradnak az itt rögzített indexek; ha bármelyik
    # függvény újra kiolvasná a doc.paragraphs-t, a korábbi törlések miatt
    # minden utána következő index elcsúszna.
    p = doc.paragraphs

    resolve_nyelv(p, answers)
    resolve_penznem(p, answers)
    resolve_merlegkeszites(p, answers)
    resolve_fordulonap(p, answers)
    resolve_alairo(p, answers)
    resolve_konyvvizsgalat(p, answers)
    resolve_ertekelesi_felelos(p, answers)
    resolve_konyveles_felelos(p, answers)
    resolve_nyilvanossagra_hozatal(p, answers)
    resolve_konszolidacio(p, answers)
    resolve_leltar(p, answers)
    resolve_simple_fills(p, answers)
    resolve_alapitas_atszervezes(p, answers)
    resolve_cegertek(p, answers)
    resolve_ertekhelyesbites(p, answers)
    resolve_celtartalek(p, answers)
    resolve_ert4(p, answers)
    remove_jelolt_instructions(p)

    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(output_path))
    print(f"Elkészült: {output_path}")


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Használat: generate_policy.py <answers.json> <output.docx>")
        sys.exit(1)
    generate(sys.argv[1], sys.argv[2])
