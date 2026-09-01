#!/usr/bin/env python3
"""
Számviteli politika - automatikus vázlatgenerátor.

Beolvassa a sablont (sablon/merged.docx) és egy kitöltött válaszkészletet
(JSON, ld. data/answers.example.json), majd legenerál egy Word-vázlatot:

  - kitölti a sablon "▶ Adatbékérő: Sz.pol. N. kerdes" jelölőit,
  - ahol a sablon a jelölő KÖRÜL már egy konkrét, ügyfélre szabott tényt is
    kimond (pl. "Business Central Microsoft rendszer segítségével történik a
    könyvvezetés"), ott NEM elég a jelölőt kitölteni: az egész mondatot
    újraírja a válaszból, különben a dokumentum önmagának mondana ellent,
  - a sablon szerzőjének saját, feltételes szerkesztői utasításait
    ("ezt csak akkor írja ha ...") kiveszi, és a hozzájuk tartozó mondatot a
    válasz alapján megtartja vagy törli,
  - a formázatlanul hagyott "innen hiányzik" helykitöltőket szöveg szerint
    megkeresi és kitölti,
  - minden behelyettesített választ a sablon eredeti "kitöltendő hely"
    jelölésével (sárga kiemelés, áthúzás, piros betűszín) hagy, hogy a
    felülvizsgáló kolléga elsőre lássa, mi került be automatikusan.

FONTOS: ez egy VÁZLATOT készít. Ahol a sablon beégetett törvényi szöveget
tartalmaz, amibe a szabadszöveges választ nem lehet biztonságosan
behelyettesíteni, a script NEM tippel, hanem "⚠ ELLENŐRIZENDŐ" jelölést
hagy - ld. a README "Miért nincs bekötve..." fejezetét.
"""
import copy
import json
import re
import sys
import unicodedata
from datetime import date, timedelta
from pathlib import Path

import docx
from docx.enum.text import WD_COLOR_INDEX
from docx.oxml.ns import qn
from docx.shared import RGBColor
from docx.table import _Cell

BASE = Path(__file__).resolve().parent.parent
TEMPLATE = BASE / "sablon" / "merged.docx"

# A sablon jelölőinek szövegkezdete - a biztonsági söprés (ld.
# sweep_leftover_markers) és a tesztek is erre keresnek.
MARKER_PREFIX = "▶ Adatbékérő"
# A sablonban formázás NÉLKÜL maradt helykitöltő (ld. fill_text_placeholder).
GAP_PLACEHOLDER = "innen hiányzik"


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


def clear_choice_marks(paragraph, except_runs=()):
    """A 'még döntendő' jelöléseket (sárga kiemelés, áthúzás, kék/piros
    betűszín) törli egy már véglegesített bekezdésben - a except_runs-ban
    megadott run-indexeket kihagyja (pl. amiket épp most jelöltünk meg
    behelyettesítettként)."""
    for i, r in enumerate(paragraph.runs):
        if i in except_runs:
            continue
        r.font.highlight_color = None
        r.font.strike = None
        if r.font.color is not None and r.font.color.type is not None:
            r.font.color.rgb = None


def mark_substituted(run):
    """A sablon eredeti 'kitöltendő hely' jelölését (sárga kiemelés,
    áthúzás, piros betűszín) alkalmazza egy behelyettesített válaszra, hogy
    a felülvizsgáló munkatárs elsőre lássa, mely szövegrészek automatikusan
    kerültek be, és melyeket kell ellenőriznie."""
    run.font.highlight_color = WD_COLOR_INDEX.YELLOW
    run.font.strike = True
    run.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)


def fill_blanks(paragraph, values):
    """Sorrendben kitölti a bekezdés kitöltendő helyeit a megadott
    értékekkel; a behelyettesített szöveget a sablon eredeti "kitöltendő"
    jelölésével (sárga kiemelés, áthúzás, piros betű) hagyja, a bekezdés
    egyéb részein pedig megszünteti a "még döntendő" jelölést.

    FIGYELEM: a csoport TÖBBI futása kiürül. A sablon a sormintában (inline)
    álló jelölők után gyakran külön, áthúzott szóköz-futást tesz - az is a
    csoport része, tehát kiürül. Ilyen helyeken a behelyettesített értéknek
    magának kell hoznia az elválasztó szóközt (ld. inline_value())."""
    groups = blank_groups(paragraph)
    marked_indices = set()
    for grp, val in zip(groups, values):
        first = paragraph.runs[grp[0]]
        first.text = text_value(val)
        mark_substituted(first)
        marked_indices.add(grp[0])
        for j in grp[1:]:
            paragraph.runs[j].text = ""
    strip_bracket_ref(paragraph)
    clear_choice_marks(paragraph, except_runs=marked_indices)


def keep(paragraph):
    """Megtartott, de nem kitöltendő bekezdés véglegesítése."""
    strip_bracket_ref(paragraph)
    clear_choice_marks(paragraph)


def delete(paragraph):
    p = paragraph._p
    parent = p.getparent()
    if parent is not None:
        parent.remove(p)


def flag_before(paragraph, note):
    """Jól látható, félkövér piros felülvizsgálati jelölést szúr be az adott
    bekezdés elé - olyan döntési pontokhoz, amikhez nincs kérdőív-adat, vagy
    amiket a motor szándékosan nem old fel."""
    new_p = paragraph.insert_paragraph_before()
    run = new_p.add_run(f"⚠ ELLENŐRIZENDŐ (nincs kérdőív-adat): {note}")
    run.bold = True
    run.font.color.rgb = docx.shared.RGBColor(0xC0, 0x00, 0x00)


# --------------------------------------------------------------------------
# Futás-szintű szerkesztés - a "D1" mondat-újraíráshoz és a formázatlan
# helykitöltőkhöz.
#
# A régi sablonban minden behelyettesítendő hely előre meg volt formázva
# (sárga/áthúzott/piros), ezért elég volt a blank_groups()/fill_blanks()
# páros. Az új sablonban két olyan eset is van, amit ezzel nem lehet
# lefedni:
#
#   1. a jelölő MELLETT a sablon már kimond egy konkrét, ügyfélre szabott
#      tényt egy sima, formázatlan futásban (pl. "az " névelő az
#      eredménykimutatás-mondatban, vagy a "5. hónap utolsó napjáig"
#      tagmondat) - ezt is a válaszból kell újraírni;
#   2. a sablon néhány helyen csak beírta, hogy "innen hiányzik", minden
#      formázás nélkül - ezeket SZÖVEG szerint kell megtalálni.
#
# Az alábbi primitívek pontos párjai a webapp-client/docx.js azonos nevű
# függvényei; a két motor kimenetének bitre egyeznie kell.
# --------------------------------------------------------------------------

def set_plain_run_text(paragraph, run_idx, text):
    """Egy futás szövegének felülírása a formázás megtartásával, JELÖLÉS
    NÉLKÜL. Nyelvtani ragasztóelemekhez (névelő, tagmondat-kötő szöveg)
    való, ami nem maga a behelyettesített válasz."""
    paragraph.runs[run_idx].text = text


def fill_run(paragraph, run_idx, value):
    """Egy MEGLÉVŐ futás szövegének lecserélése a válaszra, a szokásos
    "kitöltendő hely" jelöléssel (sárga kiemelés, áthúzás, piros betű).

    Azokhoz a helyekhez, ahol a sablon a konkrét, ügyfélre szabott tényt egy
    önálló futásban tartja (pl. a "csak sárga" kiemelésű "december hónap 31.
    napja"), de nincs mellette áthúzott kitöltendő hely, amire a
    fill_blanks() épülhetne."""
    run = paragraph.runs[run_idx]
    run.text = text_value(value)
    mark_substituted(run)


def substitute_in_run(paragraph, run_idx, before, value, after):
    """A megadott futást HÁROM futásra bontja (before / value / after) és a
    középsőt - a behelyettesített értéket - a szokásos "kitöltendő hely"
    jelöléssel látja el; a before/after futások az eredeti formázást
    öröklik.

    Mindig pontosan három futás keletkezik (üres before/after esetén is),
    hogy a Python és a JS motor futás-indexei bitre azonosak maradjanak."""
    original = paragraph.runs[run_idx]._r
    mid = copy.deepcopy(original)
    tail = copy.deepcopy(original)
    original.addnext(tail)
    original.addnext(mid)
    runs = paragraph.runs
    runs[run_idx].text = before
    runs[run_idx + 1].text = text_value(value)
    mark_substituted(runs[run_idx + 1])
    runs[run_idx + 2].text = after
    # A before/after örökölte a "kitöltendő" jelölést, ha az eredeti futáson
    # volt ilyen - ezeket le kell venni, csak a középső marad megjelölve.
    clear_choice_marks(paragraph, except_runs={run_idx + 1})
    return run_idx + 1


def fill_text_placeholder(paragraph, placeholder, value):
    """A sablonban FORMÁZÁS NÉLKÜL hagyott helykitöltő szöveget cseréli le a
    válaszra - szöveg szerint keresve, mert ezeken a helyeken nincs sem
    áthúzott kitöltendő hely, sem "▶ Adatbékérő" jelölő, amire a formázás
    alapú keresés épülhetne (IV.8. fejezet, "innen hiányzik").

    A behelyettesített érték megkapja a szokásos mark_substituted()
    jelölést, a körülötte lévő szöveg érintetlen marad.
    True, ha talált helykitöltőt."""
    for i, r in enumerate(paragraph.runs):
        pos = r.text.find(placeholder)
        if pos >= 0:
            substitute_in_run(
                paragraph, i, r.text[:pos], value, r.text[pos + len(placeholder):]
            )
            return True
    return False


def prepend_sentence(paragraph, text):
    """Behelyettesített mondatot szúr a bekezdés ELEJÉRE.

    Akkor kell, ha a sablon jelölője nyelvtanilag rossz helyen áll (pl. egy
    kettősponttal záruló felvezető mondat ÉS a hozzá tartozó felsorolás
    között), és a válasz csak a felvezető mondat elé illeszthető be
    értelmesen."""
    if not paragraph.runs:
        return False
    substitute_in_run(paragraph, 0, "", text, paragraph.runs[0].text)
    return True


def truncate_at(paragraph, marker):
    """A bekezdés szövegéből a megadott résztől kezdve mindent levág.

    A sablon több helyen az OLVASÓNAK szóló felszólítással zárja a
    bekezdést ("A Vállalkozásnak nyilatkoznia kell arról, hogy ..."), aminek
    a helyére a tényleges nyilatkozat kerül - ez a felszólítás nem mehet ki
    az ügyfélhez. True, ha talált és vágott."""
    found = False
    for r in paragraph.runs:
        if found:
            r.text = ""
            continue
        pos = r.text.find(marker)
        if pos >= 0:
            r.text = r.text[:pos].rstrip()
            found = True
    return found


def drop_leading_word(paragraph, run_idx, word):
    """A megadott futás elejéről töröl egy szót a mögötte álló szóközzel
    együtt - ahhoz kell, ha a válasz maga is tartalmazza a sablon által
    kiírt mértékegységet (pl. "200 eFt" a "... forint egyedi beszerzési
    érték" elé kerülve "200 eFt forint"-ot adna)."""
    run = paragraph.runs[run_idx]
    prefix = word + " "
    if run.text.startswith(prefix):
        run.text = run.text[len(prefix):]
        return True
    return False


# --------------------------------------------------------------------------
# A sablon szerzőjének feltételes utasításai ("D2")
#
# Az új sablon NÉGY helyén a szerző magába a dokumentum szövegébe írta bele,
# hogy az adott mondat csak egy bizonyos válasz esetén maradhat benn
# (pl. "ezt csak akkor írja ha az alábbi 25.ös kérdésre igen a válasz,
# különben meg NE ÍRJA"). Ez az utasítás SOSEM kerülhet ki az ügyfélhez -
# akkor sem, ha a mondat maga marad -, a mondatot pedig a válasz alapján
# meg kell tartani vagy törölni kell.
# --------------------------------------------------------------------------

# A három szerzői utasítás mindegyike az "ez(t) (csak) akkor ..." fordulattal
# kezdődik; a sablon jogi szövegében ez a fordulat sehol máshol nem fordul
# elő (a merged.docx-on ellenőrizve: pontosan 3 találat, mind a három
# szerzői utasítás).
_AUTHOR_INSTRUCTION_RE = re.compile(r"\bez(?:t)?\s+(?:csak\s+)?akkor\b", re.IGNORECASE)


def strip_author_instruction(paragraph):
    """A sablon szerzőjének saját, feltételes szerkesztői utasítását
    tartalmazó futás(oka)t üríti. Mindig lefut, függetlenül attól, hogy a
    hozzá tartozó mondat marad-e."""
    removed = False
    for r in paragraph.runs:
        if _AUTHOR_INSTRUCTION_RE.search(r.text):
            r.text = ""
            removed = True
    return removed


def keep_if(paragraph, condition, note=None):
    """FELTÉTELES BEKEZDÉS-MEGTARTÁS - a sablon szerzői utasításainak
    gépiesítése.

    - a szerzői utasítás szövegét MINDIG kiveszi (ld.
      strip_author_instruction), akkor is, ha a bekezdés marad;
    - ha `condition` igaz, a bekezdés marad (keep()-pel véglegesítve);
    - egyébként a bekezdés törlődik, és - ha `note` meg van adva - a helyére
      egy jól látható ⚠ ELLENŐRIZENDŐ jelölés kerül.

    Visszatérési érték: megmaradt-e a bekezdés."""
    strip_author_instruction(paragraph)
    if condition:
        keep(paragraph)
        return True
    if note:
        flag_before(paragraph, note)
    delete(paragraph)
    return False


# --------------------------------------------------------------------------
# Mondat-újraírás ("D1")
#
# Az új sablon a jelölők közel felénél már kimond egy konkrét, EBBŐL A
# ZALAELEKTRO-VÁZLATBÓL származó tényt közvetlenül a jelölő előtti
# bekezdésben (pl. "A társaság nem képez céltartalékot." / "A könyvvezetésért
# felelős: Tábor Ildikó ... (regisztráció száma: 174234)"). Ezekre nem elég a
# jelölőt kitölteni: a régi mondat ott maradna, és ellentmondana a válasznak.
# --------------------------------------------------------------------------

def rewrite_fact(fact_paragraph, marker_paragraph, text, note, fallback=None):
    """A sablon konkrét, ügyfélre szabott mondatát a válaszból újraírt
    mondatra cseréli.

    Ha van mit beírni: a jelölő helyére kerül a TELJES, újrafogalmazott
    mondat (a szokásos kitöltendő-jelöléssel), a sablon eredeti,
    ügyfélspecifikus mondata pedig törlődik.

    Ha nincs válasz, KÉT viselkedés közül választhatunk:

      - `fallback` megadva: a sablon eredeti mondatát ilyenkor is TÖRÖLJÜK,
        és helyette a semleges, cégfüggetlen `fallback` mondat kerül be
        (⚠ ELLENŐRIZENDŐ jelöléssel). Ezt kell használni mindenhol, ahol a
        sablon mondata egy MÁSIK cég azonosítható adatát tartalmazza (nevet,
        regisztrációs számot, rendszernevet) - ilyen szöveg üres válasz
        esetén sem maradhat a dokumentumban, mert az adatszivárgás lenne.
      - `fallback` nélkül: a jelölő eltűnik (sosem mehet ki az ügyfélhez), a
        sablon eredeti mondata pedig marad - ⚠ ELLENŐRIZENDŐ jelöléssel -,
        mert az egy szakmailag helyes, általános alapértelmezés, amit a
        felülvizsgáló kolléga pontosíthat."""
    text = text_value(text).strip()
    if text:
        fill_blanks(marker_paragraph, [text])
        delete(fact_paragraph)
        return True
    if fallback:
        fill_blanks(marker_paragraph, [fallback])
        delete(fact_paragraph)
        flag_before(marker_paragraph, note)
        return False
    delete(marker_paragraph)
    flag_before(fact_paragraph, note)
    keep(fact_paragraph)
    return False


def fill_or_flag(marker_paragraph, text, note):
    """Önálló bekezdésben álló jelölő kitöltése. Ha nincs válasz, a jelölőt
    töröljük (sosem mehet ki az ügyfélhez), és jelölést hagyunk helyette."""
    text = text_value(text).strip()
    if text:
        fill_blanks(marker_paragraph, [text])
        return True
    flag_before(marker_paragraph, note)
    delete(marker_paragraph)
    return False


def flag_unwired(marker_paragraph, note):
    """A jelölő helyén beégetett törvényi szöveg áll, amibe a szabadszöveges
    választ nem lehet biztonságosan behelyettesíteni (ld. README). A jelölőt
    töröljük, és ⚠ ELLENŐRIZENDŐ jelölést hagyunk a helyén."""
    flag_before(marker_paragraph, note)
    delete(marker_paragraph)


def sweep_leftover_markers(paragraphs):
    """BIZTONSÁGI HÁLÓ a generálás végén: ha bármelyik "▶ Adatbékérő" jelölő
    vagy "innen hiányzik" helykitöltő mégis a dokumentumban maradna (pl. egy
    elcsúszott bekezdés-index miatt), azt itt kivesszük - ügyfélhez sosem
    mehet ki -, és jól látható ⚠ ELLENŐRIZENDŐ jelölést hagyunk a helyén.

    A VÉGLEGES (törlések/beszúrások utáni) bekezdéslistát kell megkapnia,
    nem a generálás elején rögzítettet."""
    count = 0
    for para in paragraphs:
        text = para.text
        if MARKER_PREFIX not in text and GAP_PLACEHOLDER not in text:
            continue
        for r in para.runs:
            r.text = r.text.replace(MARKER_PREFIX, "").replace(GAP_PLACEHOLDER, "")
        flag_before(
            para,
            "a sablon egy kitöltetlen jelölője/helykitöltője maradt itt - a "
            "motor kivette a szövegét, de a bekezdést kérjük kézzel átnézni.",
        )
        count += 1
    return count


# --------------------------------------------------------------------------
# Táblázat-szintű segédfüggvények
#
# Az ÚJ sablonban egyetlen táblázat van (az üzleti jelentés célját ismertető,
# tisztán leíró szövegdoboz), amiben nincs kitöltendő hely - ezért a
# generálás EGYIK táblázat-primitívet sem hívja. A primitívek azért maradnak
# a fájlban, mert a paritás-teszt (scripts/dev/parity_check.py) a
# table_rows()/row_cells() párost használja a táblázatok run-lenyomatához, és
# mert egy jövőbeli, táblázatos sablonhoz újra kellenének. A JS oldali
# párjaik (docx.js getTables()/getRows()/getCells()/cellParagraph()/
# fillCell()/deleteTable()/cloneRow()) ugyanígy megmaradtak.
#
# FONTOS: szándékosan a NYERS w:tr / w:tc gyerekeket járjuk be, és NEM a
# python-docx `table.rows[i].cells[j]` rácsát - a python-docx a vízszintesen
# összevont (gridSpan) cellákat a rácsban többször is felsorolja, a JS oldal
# viszont csak a nyers XML-gyerekeket látja.
# --------------------------------------------------------------------------

def table_rows(table):
    """A táblázat közvetlen w:tr gyerekei."""
    return table._tbl.findall(qn("w:tr"))


def row_cells(tr):
    """Egy sor közvetlen w:tc gyerekei."""
    return tr.findall(qn("w:tc"))


def cell_paragraph(table, row_idx, cell_idx, para_idx=0):
    """Egy táblázatcella közvetlen w:p gyerekei közül a para_idx-edik."""
    tc = row_cells(table_rows(table)[row_idx])[cell_idx]
    return _Cell(tc, table).paragraphs[para_idx]


def fill_cell(paragraph, value):
    """Táblázatcella kitöltése a megadott értékkel, a szokásos
    kitöltendő-jelöléssel."""
    text = text_value(value)
    runs = paragraph.runs
    if runs:
        runs[0].text = text
        mark_substituted(runs[0])
        for r in runs[1:]:
            r.text = ""
        clear_choice_marks(paragraph, except_runs={0})
    else:
        mark_substituted(paragraph.add_run(text))


def delete_table(table):
    """Egy teljes táblázat törlése."""
    tbl = table._tbl
    tbl.getparent().remove(tbl)


def clone_row(table, row_idx):
    """A megadott sor másolatát szúrja be közvetlenül az eredeti után."""
    tr = table_rows(table)[row_idx]
    new_tr = copy.deepcopy(tr)
    tr.addnext(new_tr)
    return new_tr


# --------------------------------------------------------------------------
# Általános segédek
# --------------------------------------------------------------------------

_MONTHS = {
    "január": 1, "február": 2, "március": 3, "április": 4, "május": 5,
    "június": 6, "július": 7, "augusztus": 8, "szeptember": 9,
    "október": 10, "november": 11, "december": 12,
}

_MONTH_NAMES = [
    "január", "február", "március", "április", "május", "június",
    "július", "augusztus", "szeptember", "október", "november", "december",
]


def month_num(value):
    s = str(value).strip().lower()
    if s in _MONTHS:
        return _MONTHS[s]
    m = re.search(r"\d+", s)
    return int(m.group()) if m else 12


def month_name(num):
    """Hónapszám -> magyar hónapnév. Érvénytelen értékre üres sztring, hogy
    a hívó jelölni tudjon."""
    n = strict_int(num)
    if n is None or n < 1 or n > 12:
        return ""
    return _MONTH_NAMES[n - 1]


def strict_int(value):
    """A Python int(str) szigorúságát adja vissza egyetlen helyen, hogy a JS
    oldal (docx.js strictInt()) pontosan tükrözhesse."""
    if isinstance(value, bool):
        return None
    if isinstance(value, int):
        return value
    if isinstance(value, float):
        return int(value) if value.is_integer() else None
    s = str(value).strip()
    if not re.fullmatch(r"[+-]?\d+", s):
        return None
    return int(s)


def illustrative_date(fordulonap_honap, fordulonap_nap, offset_days):
    """A fordulónap + offset_days párja (hónap, nap, ÉV-eltolódás).

    Az év-eltolódás azért kell, mert a mérlegkészítés napja lehet még a
    tárgyévben (pl. június 30-i fordulónap + 60 nap), és lehet a tárgyévet
    követő évben (december 31-i fordulónap + 59 nap) - a sablon mondata
    ("a tárgyévet követő év ...") csak az utóbbira igaz.

    FIGYELEM: a docx.js illustrativeDate() PONTOS párja kell legyen. A Python
    datetime.date szigorúan validál (hónap 1-12, a hónapban létező nap) és
    hibás értékre kivételt dob; a JS Date ezzel szemben csendben túlcsordul
    (2024. 13. hó -> 2025. január), ezért ott ugyanezeket a feltételeket
    kézzel kell ellenőrizni. Ezért itt is a strict_int() szigorú
    egész-értelmezését használjuk (a Python int(31.9) csendben 31-et adna, a
    JS strictInt() viszont null-t), és a hónap/nap tartományt is kézzel
    ellenőrizzük."""
    m = month_num(fordulonap_honap)
    d = strict_int(fordulonap_nap)
    off = strict_int(offset_days)
    if d is None or off is None or m < 1 or m > 12 or d < 1 or d > 31:
        return None, None, None
    try:
        base = date(2024, m, d)
    except ValueError:  # pl. február 30.
        return None, None, None
    target = base + timedelta(days=off)
    if target.year < 1 or target.year > 9999:
        return None, None, None
    return target.month, target.day, target.year - 2024


def text_value(value):
    """Egy válasz szöveggé alakítása behelyettesítés előtt.

    FIGYELEM: a docx.js `textValue()` függvényének PONTOS párja kell legyen.
    A nyers Python `str()` és a JS `String()` NEM ekvivalens minden típusra:
      - `str(20.0)` -> "20.0", a `String(20.0)` viszont -> "20";
      - `str(True)` -> "True", a `String(true)` viszont -> "true".
    """
    if value is None:
        return ""
    if isinstance(value, bool):
        return "true" if value else "false"
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


def is_number(value):
    try:
        float(str(value).strip().replace(",", "."))
        return True
    except (TypeError, ValueError):
        return False


def _deaccent(s):
    return "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")


def answer(a, key, default=""):
    """Egy válasz kiolvasása. A HIÁNYZÓ és az ÜRESEN HAGYOTT mezőt egyaránt
    "nincs válasz"-ként kezeljük.

    FIGYELEM: a docx.js `answer()` PONTOS párja kell legyen. A JS
    `a.kulcs || alapertelmezes` alak NEM ekvivalens a Python
    `a.get(kulcs, alapertelmezes)`-szel."""
    value = a.get(key)
    if value is None:
        return default
    if isinstance(value, str) and value.strip() == "":
        return default
    return value


def txt(a, key):
    """Szabadszöveges válasz kiolvasása, körülvágva."""
    return text_value(answer(a, key, "")).strip()


def hu_article(word):
    """Magyar határozott névelő ("a" / "az") a megadott szó elé.

    Az új sablon több mondatában a névelő közvetlenül a behelyettesített
    érték előtt áll (pl. "... ellenőrzése a innen hiányzik feladata"), ezért
    a névelőt is a válaszhoz kell igazítani. Kizárólag a kezdőbetűt nézzük -
    ez a magyar helyesírás alapszabálya, a kivételek (pl. betűszavak
    kiejtése) kézi ellenőrzést igényelnek."""
    s = text_value(word).strip()
    if not s:
        return "a"
    first = _deaccent(s[0].lower())
    return "az" if first in "aeiou" else "a"


def inline_value(text, prefix="", suffix=" "):
    """Sormintában (inline) álló jelölő értéke.

    A sablon az inline jelölő UTÁN gyakran külön, áthúzott szóköz-futást
    tesz, amit a fill_blanks() kiürít; néhol pedig a jelölő ELŐTT sincs
    szóköz (pl. "a vállalkozás▶ ... készít."). Ezért az elválasztó
    szóközöket magának az értéknek kell hoznia - üres értéknél viszont nem
    szabad lógó szóközt hagyni."""
    s = text_value(text).strip()
    if not s:
        return ""
    return prefix + s + suffix


def join_clauses(parts, sep=" "):
    """Az üres darabokat kihagyva fűz össze mondatrészeket - a több
    válaszból összerakott mondatokhoz."""
    return sep.join(x for x in (text_value(p).strip() for p in parts) if x)


def as_sentence(text):
    """Szabadszöveges válaszból mondatot formál: ha nincs a végén mondatzáró
    írásjel, pontot tesz ki.

    Az ügyfél a kérdőívbe ritkán ír mondatvégi pontot ("havonta",
    "átlagárfolyamot használunk"), a sablonba viszont MONDATKÉNT kerül be a
    válasza - e nélkül két, egymás után fűzött válasz írásjel nélkül folyna
    egybe."""
    s = text_value(text).strip()
    if not s:
        return ""
    return s if s[-1] in ".!?:;" else s + "."


def join_sentences(parts):
    """Több válaszból összerakott bekezdés: minden nem üres darabot
    mondattá formál (as_sentence), majd egyetlen szóközzel fűzi össze."""
    return " ".join(x for x in (as_sentence(p) for p in parts) if x)


def restate(paragraph, text):
    """Egy JELÖLETLEN sablonmondat teljes újraírása a válaszból.

    A sablon néhány helyen UGYANAZT az ügyfélspecifikus tényt még egyszer,
    "▶ Adatbékérő" jelölő nélkül is kimondja (pl. "A vállalkozás üzleti vagy
    cégértékkel nem rendelkezik." a VIII. fejezetben, miközben a XII.
    fejezetben ugyanez jelölővel is szerepel). Ha csak a jelölt helyet
    oldanánk fel, a dokumentum két helyen mást állítana.

    A bekezdés első futása kapja a szöveget és a szokásos
    kitöltendő-jelölést, a többi futás kiürül."""
    runs = paragraph.runs
    if not runs:
        return False
    runs[0].text = text_value(text)
    mark_substituted(runs[0])
    for r in runs[1:]:
        r.text = ""
    clear_choice_marks(paragraph, except_runs={0})
    return True


_PENZNEM_EGYEB_PLACEHOLDERS = {"egyeb", "egyeb deviza", "egyeb devizanem", "mas", "mas deviza"}


def classify_penznem(raw):
    """A szabadon gépelhető pénznem-mező (Szv#12) besorolása. Ékezet- és
    kis/nagybetű-független illesztés a három ismert devizára, minden más
    gépelt szöveg "egyeb".

    FIGYELEM: a docx.js `classifyPenznem()` PONTOS párja kell legyen."""
    t = (raw or "").strip()
    tl = _deaccent(t.lower())
    if tl in ("", "forint", "huf"):
        return "forint"
    if tl in ("euro", "eur"):
        return "euro"
    if tl in ("usd", "dollar", "us dollar", "amerikai dollar"):
        return "usd"
    return "egyeb"


def penznem_neve(a):
    """A beszámoló pénznemének KIÍRHATÓ neve a szabadon gépelt válaszból."""
    raw = text_value(answer(a, "penznem_tipus", "forint")).strip()
    tipus = classify_penznem(raw)
    if tipus == "forint":
        return "forint"
    if tipus == "euro":
        return "euró"
    if tipus == "usd":
        return "USA dollár"
    nev = txt(a, "penznem_egyeb_neve")
    if nev:
        return nev
    raw_norm = _deaccent(raw.lower()).replace("_", " ")
    if raw and raw_norm not in _PENZNEM_EGYEB_PLACEHOLDERS:
        return raw
    return ""


def classify_beszamolo_forma(raw):
    """A szabadon gépelhető beszámoló-forma mező (Szv#13) besorolása a
    kérdőív (Munka3 munkalap) négy ismert válaszára. Ékezet- és
    kis/nagybetű-független, kulcsszó alapú illesztés; amit nem tud
    egyértelműen besorolni, azt változatlanul visszaadja - a hívó ilyenkor
    jelöl, nem tippel.

    FIGYELEM: a docx.js `classifyBeszamoloForma()` PONTOS párja kell
    legyen."""
    t = _deaccent((raw or "").strip().lower()).replace("_", " ")
    if "ifrs" in t:
        return "eves_beszamolo_ifrs"
    if "mikro" in t:
        return "mikrogazdalkodoi_egyszerusitett_eves_beszamolo"
    if "egyszerus" in t:
        return "egyszerusitett_eves_beszamolo"
    if "eves" in t:
        return "eves_beszamolo"
    return raw


# A beszámoló formájának TÁRGYESETŰ alakja - a sablon mondata
# ("Mindezek alapján a vállalkozás▶ készít.") tárgyesetet kíván.
BESZAMOLO_FORMA_TARGYESET = {
    "eves_beszamolo": "éves beszámolót",
    "egyszerusitett_eves_beszamolo": "egyszerűsített éves beszámolót",
    "mikrogazdalkodoi_egyszerusitett_eves_beszamolo":
        "mikrogazdálkodói egyszerűsített éves beszámolót",
    "eves_beszamolo_ifrs": "éves beszámolót az IFRS-ek szerint",
}

# Az eredménykimutatás eljárása - a sablon két mondatában is szerepel
# (IV.3. és XVI. fejezet), mindkettőt ugyanabból a válaszból oldjuk fel.
EREDMENY_ELJARAS = {
    "osszkoltseg_eljarassal": "összköltség eljárással",
    "forgalmi_koltseg_eljarassal": "forgalmi költség eljárással",
}

MERLEG_VALTOZAT = {"a_valtozat": "„A” változata", "b_valtozat": "„B” változata"}


def merlegkeszites_szoveg(a):
    """A mérlegkészítés napjának kiírható alakja a fordulónapból és a
    kérdőívben megadott napok számából (pl. "a tárgyévet követő év február
    hónap 28.").

    A záró " napja"/" napját" ragot a hívó teszi hozzá, mert a sablon két
    helyen két különböző esetben használja."""
    honap, nap, ev = illustrative_date(
        answer(a, "fordulonap_honap", ""),
        answer(a, "fordulonap_nap", ""),
        answer(a, "merlegkeszites_nap", 0),
    )
    nev = month_name(honap)
    if not nev:
        return ""
    elotag = "a tárgyévet követő év " if ev and ev >= 1 else "a tárgyév "
    return f"{elotag}{nev} hónap {nap}."


def fordulonap_szoveg(a):
    """A mérlegfordulónap kiírható alakja (pl. "december hónap 31. napja")."""
    nev = month_name(answer(a, "fordulonap_honap", ""))
    nap = strict_int(answer(a, "fordulonap_nap", ""))
    if not nev or nap is None:
        return ""
    return f"{nev} hónap {nap}. napja"


def letetbehelyezes_szoveg(a):
    """A letétbe helyezés határidejének kiírható alakja a választott ág
    szerint (Szv#19). A Számviteli törvény szerint az általános határidő a
    mérlegfordulónapot követő ötödik hónap utolsó napja; EGT-állam
    szabályozott piacán forgalmazott értékpapír esetén a negyedik."""
    if answer(a, "letetbehelyezes_hatarido_tipus", "altalanos") == "szabalyozott_piac":
        return "negyedik"
    return "ötödik"


def igenlo(value):
    """Szabadon gépelt igen/nem válasz értelmezése: bármi, ami "igen"-nel
    kezdődik, igenlő. Ékezet- és kis/nagybetű-független."""
    return _deaccent(text_value(value).strip().lower()).startswith("igen")


def nemleges(value):
    """Szabadon gépelt igen/nem válasz: "nem"-mel kezdődik-e."""
    return _deaccent(text_value(value).strip().lower()).startswith("nem")


# --------------------------------------------------------------------------
# I. fejezet - címlap és cégadatok
# --------------------------------------------------------------------------

def resolve_cegadatok(p, a):
    """A címlap (3-15.) és az "I.2. A vállalkozás bemutatása" (57-71.)
    adatsorai. A cégnév (Szv#1) KÉT helyen szerepel, ugyanabból a válaszból
    oldjuk fel mindkettőt."""
    for idx, key, cimke in (
        (3, "szv1", "a vállalkozás neve"),
        (6, "szv2", "a székhely címe"),
        (9, "cegjegyzekszam", "a cégjegyzékszám"),
        (12, "szv6", "az adószám"),
        (15, "szv5", "a képviseletre jogosult személy"),
        (57, "szv1", "a cég neve"),
        (59, "gazdalkodasi_forma", "a működési/gazdálkodási forma"),
        (61, "alapitas_idopontja", "az alapítás időpontja"),
        (63, "szv3", "a telephely"),
        (65, "fioktelepek", "a fióktelep(ek)"),
        (67, "szv4", "a tulajdonos(ok)"),
        (69, "szv7", "a statisztikai számjel"),
        (71, "szv8", "a főtevékenység"),
    ):
        fill_or_flag(p[idx], txt(a, key), f"hiányzik {cimke} a kérdőívből.")


def resolve_ecs_ertekhatar(p, a):
    """Szv#70 - a kis értékű tárgyi eszközök egyösszegű leírásának
    értékhatára, a számviteli politika kötelező tartalmi elemeit felsoroló
    listában (101.).

    A jelölő SORMINTÁBAN áll, közvetlenül utána a sablon "forint" szava -
    ha a válasz maga is hoz mértékegységet (pl. "200 eFt"), a sablon
    "forint" szavát ki kell venni, különben "200 eFt forint" lenne."""
    ertek = txt(a, "szv70")
    if not ertek:
        # A bekezdés a törvény kötelező tartalmi elemeit felsoroló lista egy
        # pontja - nem töröljük, csak jelöljük.
        fill_blanks(p[101], [inline_value("…")])
        flag_before(
            p[101],
            "nincs válasz a kis értékű tárgyi eszközök egyösszegű "
            "értékcsökkenési leírásának értékhatárára (Szv#70) - a "
            "felsorolás pontját kérjük kézzel kiegészíteni.",
        )
        return
    norm = _deaccent(ertek.lower())
    if re.search(r"(ft|forint)\.?$", norm):
        drop_leading_word(p[101], 3, "forint")
    fill_blanks(p[101], [inline_value(ertek)])


# --------------------------------------------------------------------------
# II. fejezet - a számviteli feladatok felelősei
# --------------------------------------------------------------------------

def resolve_szamviteli_felelosok(p, a):
    """II.1. és II.3. - a számviteli politikáért, illetve a
    számvitelpolitikai döntésekért felelős személy (Szv#30, Szv#31)."""
    fill_or_flag(
        p[188], txt(a, "szv30"),
        "nincs válasz arra, kinek a feladata a számviteli szabályzatok "
        "elkészítése, módosítása (Szv#30).",
    )

    if answer(a, "ertekeles_felelos_tipus") == "egyedi_delegalas":
        szoveg = txt(a, "ertekelesi_delegalas")
        fill_or_flag(
            p[191], szoveg,
            "egyedi delegálás lett megjelölve a számvitelpolitikai "
            "döntésekért való felelősségre (Szv#31), de a delegálás "
            "részletei nincsenek megadva.",
        )
    else:
        fill_blanks(
            p[191],
            ["a vezetés, illetve a vezetés felhatalmazása alapján a "
             "számviteli rendért felelős személy."],
        )
        if not txt(a, "ertekeles_felelos_tipus"):
            flag_before(
                p[191],
                "nincs válasz arra, hogy a számvitelpolitikai döntésekért a "
                "vezetés/számviteli rendért felelős személy felel-e, vagy "
                "egyedi delegálás van (Szv#31) - a sablon alapértelmezett "
                "ága lett feltételezve, kérjük ellenőrizni.",
            )


def resolve_leltar(p, a):
    """II.4. Leltárkészítés és leltározás felelősei (Szv#36).

    A sablon mondata a ZalaElektro-vázlatban konkrét beosztásokat nevez meg
    ("a vezetés (logisztikai vezető és gazdasági vezető) hatásköre és
    felelőssége"), ami más ügyfélnél hibás. A felvezető tagmondatot ezért
    általánosra írjuk át, a konkrét felelős pedig a kérdőív szabadszöveges
    válaszából (leltar_felelos_leiras) kerül a jelölő helyére; a bekezdés
    második, jogszabályi fele változatlan marad."""
    felelos = txt(a, "leltar_felelos_leiras")
    # A bekezdés a jelölő UTÁN egy ponttal folytatódik, ezért a beírt érték
    # nem végződhet mondatzáró írásjelre.
    set_plain_run_text(p[203], 0, "A leltárkészítés és leltározás felelősségének rendje: ")
    if felelos:
        fill_blanks(p[203], [felelos.rstrip(". ")])
    else:
        fill_blanks(p[203], ["a vezetés hatásköre és felelőssége"])
        flag_before(
            p[203],
            "nincs válasz arra, ki a leltárkészítés és leltározás felelőse "
            "(Szv#36) - a semleges 'a vezetés hatásköre és felelőssége' "
            "megfogalmazás lett beírva, kérjük pontosítani.",
        )


def resolve_penzkezeles(p, a):
    """II.5. Pénzkezelés felelősei (Szv#37).

    A sablon mondata konkrétan "a házi pénztár vezetésére kijelölt könyvelő"-t
    nevezi meg (és egy elgépelt ".:" zárja) - ezt a válaszból írjuk újra."""
    # A felelős lehet egy személy és több személy is, ezért kettősponttal
    # vezetjük be - így nem kell a magyar alany-állítmány egyeztetést
    # (végzi / végzik) egy szabadszöveges válaszból eltalálni.
    felelos = txt(a, "szv37")
    rewrite_fact(
        p[208], p[209],
        "A pénzkezelés a vezetés hatásköre és felelőssége, az ezzel "
        "kapcsolatos feladatokat a vonatkozó jogszabályi előírások, valamint "
        "a Vállalkozás pénzkezelési szabályzatának előírásai szerint kell "
        f"ellátni. A pénzkezelés felelőse: {as_sentence(felelos)}"
        if felelos else "",
        "nincs válasz arra, ki a pénzkezelés felelőse (Szv#37) - az alábbi "
        "mondat a sablon konkrét, MÁS cégre írt megfogalmazását tartalmazza, "
        "kérjük kézzel átírni.",
    )


# --------------------------------------------------------------------------
# III. fejezet - könyvvizsgálat, közzététel
# --------------------------------------------------------------------------

KONYVVIZSGALAT_MONDAT = {
    "kotelezo_altalanos":
        "A számviteli beszámoló szabályszerűségének, megbízhatóságának és "
        "valódiságának felülvizsgálatára és tanúsítására - a gazdálkodás "
        "volumenére és a kapcsolódó jogszabályi előírásokra tekintettel - a "
        "legfőbb szerv könyvvizsgálót választott.",
    "kotelezo_koztartozas_miatt":
        "A számviteli beszámoló szabályszerűségének, megbízhatóságának és "
        "valódiságának felülvizsgálatára és tanúsítására a legfőbb szerv "
        "könyvvizsgálót választott. A választás oka az, hogy az árbevételhez "
        "és létszámhoz kapcsolódó nagyságrendi mentesítések nem "
        "alkalmazhatók, mert a Vállalkozásnak a fordulónapon - "
        "összegszerűségében és lejáratában jogszabályban megszabott mértéket "
        "meghaladó jelentős - köztartozása van.",
    "onkentes":
        "A számviteli beszámoló szabályszerűségének, megbízhatóságának és "
        "valódiságának felülvizsgálatára és tanúsítására a legfőbb szerv - "
        "önkéntesen, függetlenül attól, hogy arra jogszabály nem ír elő "
        "kötelezettséget - könyvvizsgálót választott.",
    "nincs_kotelezettseg":
        "A számviteli beszámoló szabályszerűségének, megbízhatóságának és "
        "valódiságának felülvizsgálatára és tanúsítására a Vállalkozásnak "
        "jogszabályi kötelezettsége nem áll fenn, ezért a legfőbb szerv "
        "könyvvizsgálót nem választott.",
}


def resolve_konyvvizsgalat(p, a):
    """III.2. Könyvvizsgálat (Szv#25/26/28/115/116, Szv#27).

    A 223. bekezdés a sablon szerzőjének feltételes utasítását hordozza
    ("ezt csak akkor írja ha az alábbi 25.ös kérdésre igen a válasz, különben
    meg NE ÍRJA") - a kötelező könyvvizsgálatot kimondó mondat csak akkor
    maradhat, ha van könyvvizsgálati kötelezettség; az utasítás szövege
    viszont mindenképp kikerül (ld. keep_if)."""
    tipus = answer(a, "konyvvizsgalat_tipus", "")
    van_valasz = bool(text_value(tipus).strip())
    kotelezett = tipus in ("kotelezo_altalanos", "kotelezo_koztartozas_miatt")
    van_konyvvizsgalo = tipus in (
        "kotelezo_altalanos", "kotelezo_koztartozas_miatt", "onkentes"
    )

    keep_if(
        p[223], kotelezett,
        None if van_valasz else
        "nincs válasz a könyvvizsgálati kötelezettségre (Szv#25) - a "
        "kötelezettséget kimondó mondat biztonsági okból törölve lett, "
        "kérjük ellenőrizni/pontosítani.",
    )

    fill_or_flag(
        p[224], KONYVVIZSGALAT_MONDAT.get(tipus, ""),
        "nincs válasz a könyvvizsgálati kötelezettségre / a könyvvizsgáló "
        "választására (Szv#25, Szv#26) - kérjük kézzel pontosítani.",
    )

    if van_konyvvizsgalo:
        fill_or_flag(
            p[228], txt(a, "konyvvizsgalo_adatai"),
            "van könyvvizsgáló, de a könyvvizsgáló (cég) adatai hiányoznak "
            "a kérdőívből (Szv#28).",
        )
        fill_or_flag(
            p[230], txt(a, "konyvvizsgalo_nyilvantartasi_szam"),
            "van könyvvizsgáló, de a nyilvántartási száma hiányzik "
            "(Szv#115).",
        )
        fill_or_flag(
            p[233], txt(a, "konyvvizsgalatert_felelos_konyvvizsgalo"),
            "van könyvvizsgáló, de a könyvvizsgálat végrehajtásáért "
            "személyében felelős könyvvizsgáló neve hiányzik (Szv#116).",
        )
    else:
        # Nincs könyvvizsgáló: a könyvvizsgáló adatait kérő bekezdések
        # tárgytalanok, a hozzájuk tartozó feliratokkal együtt törlendők.
        for idx in (226, 228, 229, 230, 232, 233):
            delete(p[idx])

    # Az adatok könyvvizsgálatra történő átadásának határideje (Szv#27) a
    # sablonban egy "csak sárga" kiemelésű futásban áll, konkrét dátummal.
    atadas = txt(a, "szv27")
    if atadas:
        fill_run(p[238], 1, atadas.rstrip(". "))
    else:
        flag_before(
            p[238],
            "nincs válasz arra, mikor kell a beszámoló adatait "
            "könyvvizsgálatra átadni (Szv#27) - az alábbi mondat a sablon "
            "konkrét, MÁS cégre írt dátumát tartalmazza.",
        )


def resolve_kozzetetel(p, a):
    """III.4. A beszámoló letétbe helyezése és közzététele (Szv#24, Szv#19).

    A 259. bekezdésben KÉT ügyfélspecifikus adat van egy mondatban: a
    közzétételről gondoskodó személy (jelölő) és a letétbe helyezés
    határideje ("5. hónap utolsó napjáig", "csak sárga" futásban)."""
    felelos = txt(a, "szv24")
    fill_blanks(p[259], [inline_value(felelos)])
    if not felelos:
        flag_before(
            p[259],
            "nincs válasz arra, kinek a feladata gondoskodni a beszámoló "
            "közzétételéről és letétbe helyezéséről (Szv#24).",
        )
    fill_run(
        p[259], 4,
        f"{letetbehelyezes_szoveg(a)} hónap utolsó napjáig kell letétbe",
    )


# --------------------------------------------------------------------------
# IV. fejezet - a beszámoló formája, mérleg, eredménykimutatás
# --------------------------------------------------------------------------

def resolve_beszamolo_forma(p, a):
    """Szv#13 - "Mindezek alapján a vállalkozás▶ készít."

    A jelölő közvetlenül a "vállalkozás" szóhoz tapad (nincs köztük szóköz),
    utána pedig egy áthúzott szóköz-futás áll, amit a fill_blanks() kiürít -
    ezért a behelyettesített értéknek MINDKÉT oldali szóközt hoznia kell, és
    tárgyesetben kell állnia."""
    tipus = classify_beszamolo_forma(answer(a, "szv13"))
    forma = BESZAMOLO_FORMA_TARGYESET.get(tipus, "")
    if forma:
        fill_blanks(p[271], [inline_value(forma, prefix=" ")])
    else:
        fill_blanks(p[271], [inline_value("beszámolót", prefix=" ")])
        flag_before(
            p[271],
            "a beszámoló választott formája hiányzik vagy nem egyértelmű "
            "(Szv#13) - a semleges 'beszámolót' szó lett beírva, kérjük a "
            "pontos formát kézzel megadni.",
        )


def resolve_merleg_forma(p, a):
    """Szv#14 - "A mérleget a számviteli törvény 1. sz. melléklet ▶ szerint
    készítjük el". A jelölő után áthúzott szóköz-futás áll, ezért az érték
    hozza a záró szóközt."""
    valtozat = MERLEG_VALTOZAT.get(answer(a, "szv14"), "")
    if valtozat:
        fill_blanks(p[298], [inline_value(valtozat)])
    else:
        # Üres válasznál látható hiányjelet hagyunk, nem néma félmondatot.
        fill_blanks(p[298], [inline_value("…")])
        flag_before(p[298], "a mérleg választott formája (A/B változat) hiányzik (Szv#14).")


def resolve_eredmenykimutatas_forma(p, a):
    """Szv#15 - UGYANAZ a döntés a sablon KÉT pontján szerepel (IV.3. és a
    XVI. fejezet), mindkettőt ugyanabból a válaszból oldjuk fel, különben a
    dokumentum önmagának mondana ellent.

    A 305. bekezdésben a jelölő előtt a sablon névelője ("az ") áll, ami csak
    az összköltség eljáráshoz illik - a forgalmi költség eljárásnál "a"-ra
    kell javítani."""
    eljaras = EREDMENY_ELJARAS.get(answer(a, "szv15"), "")

    if eljaras:
        set_plain_run_text(p[305], 2, hu_article(eljaras) + " ")
        fill_blanks(p[305], [inline_value(eljaras)])
    else:
        # Üres válasznál a névelőt meghagyjuk, és látható hiányjelet írunk -
        # így a mondat nem csonkul némán értelmetlenre.
        set_plain_run_text(p[305], 2, "az ")
        fill_blanks(p[305], [inline_value("…")])
        flag_before(
            p[305], "az eredménykimutatás választott formája hiányzik (Szv#15)."
        )

    rewrite_fact(
        p[906], p[907],
        "Vállalkozásunk az üzleti év adózott eredményét a számviteli törvény "
        f"1. számú mellékletében meghatározott, {eljaras} készülő "
        "eredménykimutatás formában állapítja meg, mutatja be." if eljaras else "",
        "az eredménykimutatás választott formája hiányzik (Szv#15) - az "
        "alábbi mondat a sablon konkrét, MÁS cégre írt választását "
        "tartalmazza.",
    )


def resolve_tetelek_modositasa(p, a):
    """Szv#34 - a mérleg/eredménykimutatás tételeinek módosítása.

    A 313. bekezdés a sablon szerzőjének feltételes utasítását hordozza:
    "ezt akkor írja csak, ha az alábbi kérdésre a válasz az hogy nem,
    különben csak pirossal legyen itt hogy ide kell vmit írni". Az utasítás
    tehát KIFEJEZETTEN jelölést (és nem új mondatot) kér a másik ágban -
    ezért ott nem fogalmazunk meg saját szöveget."""
    valasz = txt(a, "szv34")
    nem_modosit = (not valasz) or nemleges(valasz)
    keep_if(
        p[313], nem_modosit,
        "a Társaság a mérleg/eredménykimutatás tételeit módosítja (Szv#34: "
        f"„{valasz}”) - a tételek összevonásáról/elhagyásáról szóló "
        "bekezdést a tényleges gyakorlat szerint kézzel kell megfogalmazni.",
    )
    if nem_modosit and not valasz:
        flag_before(
            p[313],
            "nincs válasz arra, módosítja-e a Társaság a mérleg és az "
            "eredménykimutatás tételeit (Szv#34) - a 'nem módosítja' ág lett "
            "feltételezve, kérjük ellenőrizni/pontosítani.",
        )
    delete(p[314])


def resolve_beszamolo_nyelve(p, a):
    """Szv#16 - "A Vállalkozás beszámolója kizárólag ▶ nyelven kerül
    összeállításra." """
    nyelv = txt(a, "beszamolo_nyelve")
    fill_blanks(p[320], [inline_value(nyelv or "magyar")])
    if not nyelv:
        flag_before(
            p[320],
            "nincs válasz a beszámoló nyelvére (Szv#16) - a 'magyar' lett "
            "feltételezve, kérjük ellenőrizni.",
        )


# --------------------------------------------------------------------------
# IV.6-IV.8. - fordulónap, mérlegkészítés, a beszámolóért való felelősség
#
# Ez a rész tartalmazza a sablon 7 db, FORMÁZÁS NÉLKÜL hagyott "innen
# hiányzik" helykitöltőjét (345-350., 355.), amiket szöveg szerint kell
# megkeresni, valamint azok "csak sárga" párjait (328., 334.), ahol a
# sablon UGYANAZT a tényt még egyszer, konkrét dátummal is kimondja - ezeket
# ugyanabból a válaszból kell feloldani, különben a dokumentum két helyen
# mást állítana.
# --------------------------------------------------------------------------

def resolve_beszamoloert_felelosseg(p, a):
    fordulonap = fordulonap_szoveg(a)
    merlegkeszites = merlegkeszites_szoveg(a)

    # IV.6. - a fordulónap a 328. bekezdésben "csak sárga" futásban áll.
    if fordulonap:
        fill_run(p[328], 1, f"12 hónap, fordulónapja {fordulonap}.")
    else:
        flag_before(
            p[328],
            "nincs (értelmezhető) válasz a mérleg fordulónapjára (Szv#18) - "
            "az alábbi mondat a sablon konkrét dátumát tartalmazza.",
        )

    # IV.7. - a mérlegkészítés időpontja szintén "csak sárga" futásban.
    if merlegkeszites:
        fill_run(p[334], 1, merlegkeszites)
    else:
        flag_before(
            p[334],
            "nincs (értelmezhető) válasz a mérlegkészítés időpontjára "
            "(Szv#17/Szv#18) - az alábbi mondat a sablon konkrét dátumát "
            "tartalmazza.",
        )

    # IV.8. - a beszámolóért felelős személy (Szv#20).
    fill_or_flag(
        p[341], txt(a, "szv20"),
        "nincs válasz arra, ki felelős a beszámoló elkészítéséért (Szv#20).",
    )

    # A beszámoló készítésének ütemterve - 5 db "innen hiányzik" helykitöltő.
    utemterv = (
        (345, fordulonap,
         "nincs (értelmezhető) válasz a mérleg fordulónapjára (Szv#18)."),
        (346, f"{merlegkeszites} napja" if merlegkeszites else "",
         "nincs (értelmezhető) válasz a mérlegkészítés időpontjára (Szv#17)."),
        (347, txt(a, "szv27"),
         "nincs válasz arra, mikor kell a beszámoló adatait "
         "könyvvizsgálatra átadni (Szv#27)."),
        (348,
         "a mérlegfordulónapot követő "
         f"{letetbehelyezes_szoveg(a)} hónap utolsó napja",
         ""),
        (349, txt(a, "beszamolo_kozzetel_datum"),
         "nincs válasz a beszámoló közzétételének határidejére."),
    )
    for idx, ertek, note in utemterv:
        if ertek:
            fill_text_placeholder(p[idx], GAP_PLACEHOLDER, ertek)
        else:
            fill_text_placeholder(p[idx], GAP_PLACEHOLDER, "…")
            flag_before(p[idx], note)

    # "... a munkafolyamatok irányítása, ellenőrzése a innen hiányzik
    # feladata." - a helykitöltő ELŐTT álló "a" névelőt is a válaszhoz kell
    # igazítani ("a gazdasági vezető" / "az ügyvezető"), ezért a névelővel
    # együtt cseréljük.
    iranyito = txt(a, "munkafolyamat_iranyitas_felelos")
    if iranyito:
        fill_text_placeholder(
            p[350], "a " + GAP_PLACEHOLDER,
            f"{hu_article(iranyito)} {iranyito.rstrip('.')}",
        )
    else:
        fill_text_placeholder(p[350], "a " + GAP_PLACEHOLDER, "a vezetés")
        flag_before(
            p[350],
            "nincs válasz arra, ki felelős a munkafolyamatok irányításáért, "
            "ellenőrzéséért - a semleges 'a vezetés' lett beírva.",
        )

    # A beszámolót aláíró személy (Szv#5/Szv#21).
    fill_or_flag(
        p[353], txt(a, "beszamolo_alairoja") or txt(a, "szv5"),
        "nincs válasz arra, ki írja alá a beszámolót (Szv#21).",
    )

    # A közzétételért felelős személy - UGYANAZ a válasz, mint a 259.
    # bekezdésben (Szv#24), különben a dokumentum két helyen mást mondana.
    kozzetevo = txt(a, "szv24")
    if kozzetevo:
        fill_text_placeholder(p[355], GAP_PLACEHOLDER, kozzetevo)
    else:
        fill_text_placeholder(p[355], GAP_PLACEHOLDER, "…")
        flag_before(
            p[355],
            "nincs válasz arra, ki felelős a beszámoló közzétételéért "
            "(Szv#24).",
        )

    # Szv#33 - a beszámoló elérhetősége.
    if answer(a, "beszamolo_honlapon") == "igen":
        cim = txt(a, "beszamolo_honlap_cim")
        rewrite_fact(
            p[357], p[358],
            "A beszámoló a vállalkozás székhelyén megtekinthető, továbbá "
            f"a(z) {cim} internetes honlapon is elérhető." if cim else "",
            "a beszámoló a honlapon elérhető, de a honlap címe hiányzik "
            "(Szv#33).",
        )
    else:
        keep(p[357])
        delete(p[358])
        if not txt(a, "beszamolo_honlapon"):
            flag_before(
                p[357],
                "nincs válasz arra, elérhető-e a beszámoló honlapon "
                "(Szv#33) - a 'csak a székhelyen tekinthető meg' ág lett "
                "feltételezve, kérjük ellenőrizni/pontosítani.",
            )


# --------------------------------------------------------------------------
# IV.9-VI. - kiegészítő melléklet, jelentős összeg, konszolidáció
# --------------------------------------------------------------------------

def resolve_kiegeszito_melleklet(p, a):
    fill_or_flag(
        p[448], txt(a, "szv105"),
        "nincs válasz a fedezeti ügyletekre, a fedezeti hatékonyság "
        "mérésére, illetve a kockázatkezelési politikára (Szv#105).",
    )
    fill_or_flag(
        p[463],
        "A Vállalkozás a kiegészítő mellékletben az alábbi, a vagyoni, "
        "pénzügyi és jövedelmi helyzetet bemutató mutatókat közli: "
        f"{as_sentence(txt(a, 'kiegeszito_melleklet_sajatossagok'))}"
        if txt(a, "kiegeszito_melleklet_sajatossagok") else "",
        "nincs válasz arra, milyen vagyoni, pénzügyi és jövedelmi helyzetet "
        "bemutató mutatókat közöl a kiegészítő melléklet (Szv#101) - az "
        "alábbi mutatólista a sablon alapértelmezése.",
    )

    # Szv#39 - "... jelentős összegnek tekinti a ▶ mFt-ot meghalódó
    # összeget." A sablon elgépelte a "meghaladó" szót, ezt is javítjuk.
    set_plain_run_text(p[480], 2, " mFt-ot meghaladó összeget.")
    fill_or_flag(
        p[480], txt(a, "jelentos_osszeg_bemutatas"),
        "nincs válasz arra, mit tekint a vállalkozás jelentős összegnek a "
        "bemutatási kötelezettség szempontjából (Szv#39).",
    )


KONSZOLIDACIO_MONDAT = {
    "nem_erintett":
        "Jelenleg a Vállalkozás nem minősül a Számviteli törvény 3. § (2) "
        "bek. 1-4. pontja szerinti definíció alapján sem anyavállalatnak, sem "
        "közös vezetésű vállalkozásnak, sem leányvállalatnak, sem társult "
        "vállalkozásnak, ezért nem tagja olyan vállalkozáscsoportnak, "
        "amelyről összevont (konszolidált) éves beszámolót kellene "
        "készíteni. Konszolidálásba történő bevonás esetén a számviteli "
        "politika e fejezetének felülvizsgálata és módosítása szükséges.",
    "leany_mentesul_bevonas_alol":
        "A Vállalkozás a Számviteli törvény szerint leányvállalatnak "
        "minősül, de a Számviteli törvény 119. §-a alapján mentesül az "
        "anyavállalat által készítendő összevont (konszolidált) éves "
        "beszámolóba való bevonás alól, így az összevont (konszolidált) éves "
        "beszámoló elkészítése a Vállalkozást nem érinti, csak az "
        "anyavállalat felé történő adatszolgáltatást kell teljesíteni.",
    "leany_bevonva":
        "A Vállalkozás a Számviteli törvény szerint leányvállalatnak "
        "minősül, amely bevonásra kerül az anyavállalat által készítendő "
        "összevont (konszolidált) éves beszámolóba. A konszolidációs "
        "adatszolgáltatás teljesítése a Vállalkozás számviteli beszámolóját "
        "nem érinti.",
    "anya_mentesul_folerendelt_anya_miatt":
        "A Vállalkozás a Számviteli törvény szerint anyavállalatnak minősül, "
        "de a számviteli törvény 116. §-a alapján mentesül a konszolidált "
        "éves beszámoló elkészítése alól, mert fölérendelt anyavállalata "
        "összevont (konszolidált) éves beszámolót készít, amelybe a "
        "Vállalkozást is bevonja.",
    "anya_mentesul_nagysagrend_alapjan":
        "A Vállalkozás a Számviteli törvény szerint anyavállalatnak minősül, "
        "de a számviteli törvény 117. §-a alapján mentesül az összevont "
        "(konszolidált) éves beszámoló készítése alól. A Vállalkozás a "
        "jövőben folyamatosan figyeli, hogy a mentesítés feltételei "
        "továbbra is fennállnak-e.",
    "anya_keszit_konszolidaltat":
        "A Vállalkozás a Számviteli törvény szerint anyavállalatnak minősül "
        "és összevont (konszolidált) éves beszámolót készít.",
}


def resolve_konszolidacio(p, a):
    """V. KONSZOLIDÁCIÓ (Szv#35).

    A 484. bekezdés a sablon szerzőjének feltételes utasítását hordozza
    ("ez csak akkor legyen itt ha az alábbi kérdésre a válasz igen"), és a
    ZalaElektro-vázlat konkrét helyzetét (leányvállalat, bevonásra kerül)
    írja le - ez más ügyfélnél hibás.

    A fejezet CÍMÉT (482.) akkor sem töröljük, ha a Vállalkozás nincs
    konszolidációval érintve: a sablonnak van erre az esetre is korrekt
    megfogalmazása (a régi sablonból átvéve), ezért a fejezet tartalommal
    marad, nem üresedik ki."""
    tipus = answer(a, "konszolidacio_tipus", "")
    mondat = KONSZOLIDACIO_MONDAT.get(tipus, "")
    reszlet = txt(a, "konszolidacio_reszletszabalyok")
    if mondat and tipus != "nem_erintett" and reszlet:
        mondat += (
            " Az anyavállalat felé történő adatszolgáltatás és az azt "
            "alátámasztó számviteli nyilvántartások vezetése: "
            f"{as_sentence(reszlet)}"
        )

    # Az eredeti (ZalaElektro-specifikus) mondat mindenképp elhagyandó; a
    # szerzői utasítás szövege is (keep_if gondoskodik róla).
    keep_if(p[484], False, None)
    fill_or_flag(
        p[485], mondat,
        "nincs válasz arra, tagja-e a vállalkozás konszolidálásba bevont "
        "vállalkozáscsoportnak (Szv#35) - a fejezet szövegét kézzel kell "
        "megadni.",
    )


def resolve_minositesek(p, a):
    """VI. Minősítések, értelmezések (Szv#40-43)."""
    # Szv#40 - a jelentős összegű hiba határa. A 489. bekezdés a számviteli
    # törvény 3. § (3) 3. pontjának SZÓ SZERINTI szövege (mérlegfőösszeg 2%-a,
    # de legalább 1 millió forint); a szabadszöveges választ ebbe nem lehet
    # biztonságosan behelyettesíteni, ezért nem tippelünk (ld. README).
    flag_unwired(
        p[490],
        "a jelentős összegű hiba értékhatárát (Szv#40) a fenti bekezdés a "
        "számviteli törvény szerinti általános szabállyal adja meg - ha a "
        "Vállalkozás ettől eltérő (a törvényi kereten belüli) értékhatárt "
        "határozott meg, azt kézzel kell átvezetni.",
    )

    # Szv#41 - a jelentős mérték.
    hanyad = txt(a, "jelentos_mertek_hanyad")
    rewrite_fact(
        p[503], p[504],
        "Amennyiben jogszabály, vagy a számviteli politika más része "
        "eltérően nem rendelkezik, úgy a konkrét eszközt, forrást, mérlegen "
        "kívüli tételt érintő számviteli elszámolás, illetve azok értékelése "
        "szempontjából jelentős mértékűnek minősül a vizsgált összeg, ha a "
        "mérleg vagy eredménykimutatás nagybetűvel jelölt tételét legalább "
        f"{hanyad.rstrip('%').strip()}%-kal megváltoztatja." if hanyad else "",
        "nincs válasz arra, milyen hányadot tekint a vállalkozás jelentős "
        "mértékűnek (Szv#41) - az alábbi bekezdés a sablon 2%-os "
        "alapértelmezését tartalmazza.",
    )

    # Szv#42 - a tartós tendencia. A 518. bekezdés törvényi fogalommeghatározás
    # ("egy éven (365 naptári napon) túli időintervallum"); szabadszöveges
    # válaszból ezt nem lehet biztonságosan átírni.
    flag_unwired(
        p[519],
        "a tartós tendencia fogalmát (Szv#42) a fenti bekezdés a törvényi "
        "általános szabállyal (365 naptári nap) adja meg - ha a Vállalkozás "
        "ettől eltérően értelmezi, azt kézzel kell átvezetni.",
    )

    # Szv#43 - kivételes nagyság. NÉGY külön válasz fut össze egy mondatba.
    osszeg = txt(a, "kiveteles_nagysag_osszeghatar")
    reszek = []
    for kulcs, cimke in (
        ("kiveteles_nagysag_sajattoke_szazalek", "az előző üzleti évi saját tőke"),
        ("kiveteles_nagysag_bevetel_szazalek", "az előző üzleti évi összes bevétel"),
        ("kiveteles_nagysag_koltseg_szazalek",
         "az előző üzleti évi összes költség és ráfordítás"),
    ):
        ertek = txt(a, kulcs).rstrip("%").strip()
        if ertek:
            reszek.append(f"{cimke} {ertek}%-át")
    felsorolas = join_clauses(
        [f"a(z) {osszeg} összeget" if osszeg else ""] + reszek, ", illetve "
    )
    rewrite_fact(
        p[525], p[526],
        "Kivételes nagyságú vagy előfordulású tételnek azokat a tételeket "
        f"tekintjük, amelyek értéke meghaladja {felsorolas}. Ennek "
        "megfelelően az elszámolt bevételeket, illetve ráfordításokat ezt "
        "meghaladó összeg esetén a Kiegészítő mellékletben tételesen "
        "bemutatjuk." if felsorolas else "",
        "nincs válasz arra, milyen értékű bevételt, költséget, ráfordítást "
        "tekintenek kivételes nagyságúnak (Szv#43) - az alábbi bekezdés a "
        "sablon konkrét, MÁS cégre írt értékhatárait tartalmazza.",
    )


# --------------------------------------------------------------------------
# VII. fejezet - könyvvezetés, bizonylatolás, könyvviteli zárlat
# --------------------------------------------------------------------------

KOLTSEGELSZAMOLAS_MONDAT = {
    "csak_5":
        "Vállalkozásunk a könyvvezetésében a költségeket csak az 5. "
        "számlaosztályban, költségnemenként számolja el.",
    "5_es_67":
        "Vállalkozásunk a könyvvezetésében a költségeket elsődlegesen az 5. "
        "számlaosztályban költségnemenként, másodlagosan a 6-7. "
        "számlaosztályokban költséghely-költségviselőnként számolja el.",
}


def resolve_konyvvezetes(p, a):
    """VII.1-VII.2. - könyvvezetés pénzneme, könyvelőprogram, a
    könyvvezetésért felelős személy, költségelszámolás."""
    # Szv#12 - a könyvvezetés és a beszámoló pénzneme.
    penznem = penznem_neve(a)
    rewrite_fact(
        p[533], p[534],
        "A létesítő okiratban a jegyzett tőke meghatározására rögzített "
        f"pénznem {hu_article(penznem)} {penznem}, ez a beszámoló pénzneme. "
        "Ennek megfelelően a könyvek vezetése és a beszámoló elkészítése a "
        f"beszámoló pénznemében, azaz {penznem} pénznemben történik."
        if penznem else "",
        "nincs (értelmezhető) válasz a beszámoló és a könyvvezetés "
        "pénznemére (Szv#12) - az alábbi bekezdés a forintot feltételezi.",
    )
    if penznem and classify_penznem(answer(a, "penznem_tipus", "forint")) != "forint":
        flag_before(
            p[534],
            "a beszámoló pénzneme nem forint - a dokumentum több, "
            "jelöletlen bekezdése (pl. a mértékegységről és a devizás "
            "tételek forintra átszámításáról szóló részek) forintot említ, "
            "kérjük ezeket is átnézni.",
        )

    # Szv#44 - a könyvelőprogram. A sablon konkrétan a ZalaElektro
    # rendszerét ("Business Central Microsoft") nevezi meg.
    program = txt(a, "konyvelo_program")
    rewrite_fact(
        p[538], p[539],
        "Vállalkozásunk esetében a könyvvezetés "
        f"{hu_article(program)} {program} segítségével történik."
        if program else "",
        "nincs válasz arra, milyen könyvelőprogramot használ a Társaság "
        "(Szv#44) - a sablon MÁS cég rendszerét megnevező mondata törölve "
        "lett, kérjük a ténylegesen használt programot beírni.",
        fallback="A könyvvezetéshez használt könyvelőprogram megnevezését a "
                 "számviteli politikában rögzíteni szükséges.",
    )

    # Szv#32 - a könyvvezetésért felelős személy. A sablon konkrét nevet és
    # regisztrációs számot tartalmaz ("Tábor Ildikó ... 174234").
    nev = txt(a, "konyvvezetesert_felelos_szemely") or txt(a, "konyveles_felelos_neve")
    regszam = txt(a, "konyveles_felelos_regszam")
    tipus = answer(a, "konyveles_felelos_tipus", "")
    kepesites = {
        "merlegkepes_konyvelo": "mérlegképes könyvelő",
        "oklev_konyvvizsgalo": "okleveles könyvvizsgáló",
    }.get(tipus, "")
    if tipus == "mentesseg_20mft_arbevetel_alatt":
        felelos_mondat = (
            "Tekintettel arra, hogy a Vállalkozás éves (éves szintre "
            "átszámított) nettó árbevétele két év átlagában a 20 MFt-ot nem "
            "haladja meg, a számviteli feladatok végzését, irányítását, a "
            "beszámoló összeállítását nem mérlegképes könyvelő vagy nem "
            "okleveles könyvvizsgáló végzi."
        )
        if nev:
            felelos_mondat += f" A könyvvezetésért felelős: {nev}."
    elif nev:
        felelos_mondat = "A könyvvezetésért felelős: " + join_clauses(
            [nev, f"({kepesites}" + (f", regisztrációs száma: {regszam})" if regszam
                                     else ")") if kepesites
             else (f"(regisztrációs száma: {regszam})" if regszam else "")]
        ).rstrip() + "."
    else:
        felelos_mondat = ""
    rewrite_fact(
        p[540], p[541], felelos_mondat,
        "nincs válasz arra, ki a könyvvezetésért felelős személy (Szv#32) - "
        "a sablon MÁS cég munkatársát (nevét és regisztrációs számát) "
        "megnevező mondata törölve lett, kérjük a tényleges felelőst "
        "beírni.",
        fallback="A könyvvezetésért felelős személy nevét és regisztrációs "
                 "számát a számviteli politikában rögzíteni szükséges.",
    )

    # Szv#45/#46 - a költségelszámolás módja és a könyvelőprogramban
    # használt kódrendszer. A sablon mindkettőt konkrétan kimondja.
    koltseg = KOLTSEGELSZAMOLAS_MONDAT.get(
        answer(a, "koltsegelszamolas_tipus", ""), ""
    )
    kodrendszer = txt(a, "szv46")
    if koltseg and kodrendszer and not nemleges(kodrendszer):
        koltseg += (
            " A könyvelőprogramban a vállalkozás az alábbiak szerinti "
            f"megbontást alkalmazza: {as_sentence(kodrendszer)}"
        )
    rewrite_fact(
        p[545], p[546], koltseg,
        "nincs válasz a költségelszámolás választott módjára (Szv#45) - az "
        "alábbi bekezdés a sablon konkrét, MÁS cégre írt gyakorlatát írja "
        "le.",
    )


def resolve_bizonylatok(p, a):
    """VII.3-VII.4. - bizonylati elv, rögzítési határidők, számlarend."""
    # Szv#54 - a bizonylatok könyvekben történő rögzítésének határideje.
    rewrite_fact(
        p[585], p[586],
        "A számviteli bizonylatok könyvekben történő rögzítésének "
        f"határideje: {as_sentence(txt(a, 'szv54'))}" if txt(a, "szv54") else "",
        "nincs válasz a számviteli bizonylatok rögzítésének határidejére "
        "(Szv#54) - az alábbi bekezdés a sablon általános, MÁS cégre írt "
        "határidőit tartalmazza.",
    )

    # Szv#71 - az értékcsökkenés elszámolásának gyakorisága.
    rewrite_fact(
        p[588], p[589],
        "Az értékcsökkenés elszámolása "
        f"{txt(a, 'szv71').rstrip('.')} történik." if txt(a, "szv71") else "",
        "nincs válasz arra, milyen gyakorisággal számolják el az "
        "értékcsökkenést (Szv#71) - az alábbi bekezdés a sablon konkrét, MÁS "
        "cégre írt gyakorlatát írja le.",
    )

    # Szv#49/#50 - a bizonylatok és a könyvelés közötti összhang, illetve a
    # bizonylatok jellege.
    osszhang = txt(a, "bizonylat_konyveles_osszhang")
    jelleg = txt(a, "bizonylatkezeles_szabalyai")
    rewrite_fact(
        p[598], p[599],
        join_sentences([
            f"A számviteli bizonylatok és a könyvelés közötti összhangot a "
            f"Vállalkozás a következők szerint teremti meg: {osszhang}"
            if osszhang else "",
            f"A Vállalkozás által alkalmazott bizonylatok jellege: {jelleg}"
            if jelleg else "",
        ]),
        "nincs válasz a bizonylatok és a könyvelés közötti összhang "
        "megteremtésére, illetve a bizonylatok jellegére (Szv#49, Szv#50) - "
        "az alábbi bekezdés a sablon konkrét, MÁS cégre írt gyakorlatát "
        "írja le.",
    )

    # Szv#51/#52/#53 - a bizonylatok hitelessége. A régi sablonban ez egy
    # táblázat volt, itt egyetlen prózai mondat.
    if answer(a, "bizonylat_hitelesites_tipus", "") == "feljogositott_szemelyek":
        alairas = join_sentences([
            "A bizonylatok alaki és tartalmi hitelességét az egyes "
            "bizonylattípusok esetében az arra feljogosított személyek "
            "aláírása igazolja.",
            f"Kimenő számla aláírására jogosult: {txt(a, 'szv52')}."
            if txt(a, "szv52") else "",
            "Egyéb kimenő számviteli bizonylat aláírására jogosult: "
            f"{txt(a, 'szv53')}." if txt(a, "szv53") else "",
        ])
    elif answer(a, "bizonylat_hitelesites_tipus", "") == "kepviselo_alairasa":
        alairas = (
            "A bizonylatot a vállalkozás képviseletére jogosult személy "
            "(ideértve a Ptk. szerinti képviseletet is) írja alá, a "
            "bizonylatok alaki és tartalmi hitelességét az ő aláírása "
            "igazolja."
        )
    else:
        alairas = ""
    rewrite_fact(
        p[617], p[618], alairas,
        "nincs válasz arra, hogy a bizonylatok hitelességét a képviseletre "
        "jogosult aláírása igazolja-e, vagy bizonylattípusonként külön "
        "feljogosított aláírók vannak (Szv#51).",
    )

    # Szv#48 - a számlarend kialakításáért felelős személy.
    rewrite_fact(
        p[625], p[626],
        "A számlatükör és a számlarend összeállításáért, annak folyamatos "
        "karbantartásáért, a naprakész könyvvezetés helyességéért "
        f"{txt(a, 'szv48').rstrip('.')} a felelős." if txt(a, "szv48") else "",
        "nincs válasz arra, kinek a felelőssége a számlatükör és a "
        "számlarend kialakítása, karbantartása (Szv#48) - az alábbi mondat a "
        "sablon alapértelmezését tartalmazza.",
    )

    # Szv#47 - a számlatükör felépítése. A sablon jelölője (628.) egy
    # kettősponttal záruló felvezető mondat ÉS a hozzá tartozó felsorolás
    # KÖZÉ került, ahol egy önálló mondat megtörné a felsorolást - ezért a
    # választ a felvezető mondat ELÉ szúrjuk be, a jelölőt pedig töröljük.
    struktura = txt(a, "szv47")
    if struktura:
        prepend_sentence(
            p[627], f"A programban a számlatükör felépítése: {as_sentence(struktura)} "
        )
        delete(p[628])
    else:
        flag_unwired(
            p[628],
            "nincs válasz arra, hogyan épül fel a programban a számlatükör "
            "struktúrája (Szv#47).",
        )


def resolve_zarlat(p, a):
    """VII.6. Kiemelt számviteli teendők, könyvviteli zárlat
    (Szv#55/56/57/58)."""
    fill_or_flag(
        p[658],
        "Az évközi zárások során a Vállalkozás az alábbi feladatokat végzi "
        f"el: {as_sentence(txt(a, 'evkozi_zarasok_feladatai'))}"
        if txt(a, "evkozi_zarasok_feladatai") else "",
        "nincs válasz arra, milyen feladatokat végez a Társaság az évközi "
        "zárások során (Szv#57).",
    )

    # Szv#55/#58 - a főkönyvi zárás gyakorisága. A sablon konkrét, MÁS cégre
    # írt ütemezést tartalmaz.
    zaras = txt(a, "szv55")
    vegyes = txt(a, "szv58")
    rewrite_fact(
        p[659], p[660],
        join_sentences([
            f"A főkönyvi zárás rendje: {zaras}" if zaras else "",
            "Évközi zárás hiányában a vegyes tételek könyvelésének és a "
            f"főkönyv ellenőrzésének rendje: {vegyes}" if vegyes else "",
        ]),
        "nincs válasz arra, történik-e havi vagy negyedéves főkönyvi zárás "
        "(Szv#55) - az alábbi mondat a sablon konkrét, MÁS cégre írt "
        "ütemezését tartalmazza.",
    )

    # Szv#56 - a zárások határideje. Az éves zárás határidejét a
    # mérlegkészítés választott időpontjából vezetjük le, hogy a dokumentum
    # két helyen se mondjon mást.
    merlegkeszites = merlegkeszites_szoveg(a)
    rewrite_fact(
        p[661], p[662],
        join_sentences([
            f"Az éves zárás határideje {merlegkeszites} napja."
            if merlegkeszites else "",
            f"Az évközi zárások határideje: {txt(a, 'szv56')}"
            if txt(a, "szv56") else "",
        ]),
        "nincs válasz a zárások határidejére (Szv#56), és a mérlegkészítés "
        "időpontja sem állapítható meg - az alábbi mondat a sablon konkrét, "
        "MÁS cégre írt dátumát tartalmazza.",
    )


# --------------------------------------------------------------------------
# VIII-XII. fejezet - mérlegtételek, értékelés, értékcsökkenés
# --------------------------------------------------------------------------

def resolve_befektetett_eszkozok(p, a):
    """VIII. A mérlegtételek tartalma (Szv#59-62, 81-82, 86)."""
    # Szv#59 - alapítás-átszervezés aktivált értéke.
    if answer(a, "alapitas_atszervezes_aktivalja") == "igen":
        ev = txt(a, "alapitas_atszervezes_leiras_ev")
        mondat = (
            "Az alapítás, átszervezés aktivált értékét a Vállalkozás az "
            "immateriális javak között mutatja ki, és "
            f"{ev} év alatt írja le." if is_number(ev) else ""
        )
        note = (
            "a Vállalkozás aktiválja az alapítás-átszervezés költségeit "
            "(Szv#59), de a leírási idő hiányzik vagy nem szám - az alábbi "
            "mondat ennek az ELLENKEZŐJÉT állítja, kérjük kézzel javítani."
        )
    else:
        mondat = (
            "Alapítás, átszervezés aktivált értékét nem mutatunk ki, az "
            "ilyen címen felmerült költségeket a tárgyévi költségek között "
            "számoljuk el."
        )
        note = ""
    rewrite_fact(p[681], p[682], mondat, note)
    if mondat and answer(a, "alapitas_atszervezes_aktivalja") != "igen" and not txt(
        a, "alapitas_atszervezes_aktivalja"
    ):
        flag_before(
            p[682],
            "nincs válasz arra, aktiválják-e az alapítás-átszervezés "
            "költségeit (Szv#59) - a 'nem aktiváljuk' ág lett feltételezve, "
            "kérjük ellenőrizni/pontosítani.",
        )

    # Szv#60 - kísérleti fejlesztés.
    rewrite_fact(
        p[692], p[693],
        f"Vállalkozásunk kísérleti fejlesztés tekintetében: {as_sentence(txt(a, 'szv60'))}"
        if txt(a, "szv60") else "",
        "nincs válasz a kísérleti fejlesztés költségeinek aktiválására "
        "(Szv#60) - az alábbi mondat a sablon konkrét, MÁS cégre írt "
        "állítását tartalmazza.",
    )

    # Szv#62 - vagyoni értékű jogok.
    rewrite_fact(
        p[695], p[696],
        "A vállalkozásnál vagyoni értékű jogként jellemzően a következők "
        f"fordulnak elő: {as_sentence(txt(a, 'szv62'))}" if txt(a, "szv62") else "",
        "nincs válasz arra, milyen jogokat tartanak nyilván a vagyoni "
        "értékű jogok mérlegsoron (Szv#62) - az alábbi mondat a sablon "
        "konkrét, MÁS cégre írt állítását tartalmazza.",
    )

    # Szv#81/#82 - vásárolt és saját termelésű készletek.
    rewrite_fact(
        p[727], p[728],
        join_sentences([
            "A vásárolt anyag- és árukészletek nyilvántartása, értékelése: "
            f"{txt(a, 'szv81')}" if txt(a, "szv81") else "",
            f"Saját termelésű készletek: {txt(a, 'szv82')}"
            if txt(a, "szv82") else "",
        ]),
        "nincs válasz a vásárolt, illetve a saját termelésű készletek "
        "nyilvántartására és értékelésére (Szv#81, Szv#82) - az alábbi "
        "bekezdés a sablon konkrét, MÁS cégre írt állítását tartalmazza.",
    )

    # Szv#86 - a vevőkövetelések csoportosítása (a megelőző bekezdés
    # általános törvényi szöveg, azt nem írjuk át).
    fill_or_flag(
        p[733],
        "A vevőkövetelések könyvelési csoportosítása: "
        f"{as_sentence(txt(a, 'szv86'))}" if txt(a, "szv86") else "",
        "nincs válasz arra, hogyan csoportosítják a vevőköveteléseket a "
        "könyvelésben (Szv#86).",
    )

    # Szv#95 / Ért#5 - a devizás hiteltartozások árfolyamvesztesége.
    if answer(a, "arfolyamveszteseg_halasztott_raforditas",
              "nem_szamoljuk_el") == "elszamoljuk":
        arfolyam = (
            "A devizaszámlán meglévő devizakészlettel nem fedezett, külföldi "
            "pénzértékre szóló hiteltartozások, devizakötvény-kibocsátásból "
            "származó tartozások esetén a tárgyévben, a mérlegfordulónapi "
            "értékelésből keletkezett, árfolyamnyereséggel nem ellentételezett "
            "árfolyamveszteséget halasztott ráfordításként elszámoljuk, és a "
            "kapcsolódó céltartalékot megképezzük."
        )
    else:
        arfolyam = (
            "A devizaszámlán meglévő devizakészlettel nem fedezett, külföldi "
            "pénzértékre szóló hiteltartozások, devizakötvény-kibocsátásból "
            "származó tartozások mérlegfordulónapi értékeléséből keletkezett "
            "árfolyamveszteséget nem számoljuk el halasztott ráfordításként, "
            "ezért céltartalék képzés sem kapcsolódik hozzá."
        )
    rewrite_fact(p[759], p[760], arfolyam, "")
    if not txt(a, "arfolyamveszteseg_halasztott_raforditas"):
        flag_before(
            p[760],
            "nincs válasz arra, halasztott ráfordításként számolják-e el a "
            "devizás hiteltartozások árfolyamveszteségét (Szv#95/Ért#5) - a "
            "'nem számoljuk el' ág lett feltételezve, kérjük "
            "ellenőrizni/pontosítani.",
        )


def resolve_ertekeles(p, a):
    """IX-XII. fejezet - értékelés, értékhelyesbítés, értékcsökkenés,
    értékvesztés."""
    fill_or_flag(
        p[771], txt(a, "ert1"),
        "nincs válasz arra, van-e a számviteli törvénytől eltérő értékelési "
        "módszerük (Ért#1).",
    )

    fill_or_flag(
        p[773],
        join_sentences([
            "A terven felüli értékcsökkenés elszámolásának gyakorisága: "
            f"{txt(a, 'szv77')}" if txt(a, "szv77") else "",
            f"Az értékvesztés elszámolásának gyakorisága: {txt(a, 'szv85')}"
            if txt(a, "szv85") else "",
        ]),
        "nincs válasz a terven felüli értékcsökkenés, illetve az "
        "értékvesztés elszámolásának gyakoriságára (Szv#77, Szv#85).",
    )

    # Szv#63/#64/#65 - értékhelyesbítés eszközcsoportonként. A sablon
    # egyetlen mondattal kijelenti, hogy a Vállalkozás nem él a
    # lehetőséggel - ezt mind a három eszközcsoport válaszából újraírjuk.
    csoportok = []
    hianyzo_kapu = False
    for gate, text_key, cimke in (
        ("ertekhelyesbites_immat_van", "ertekhelyesbites_immat_javak",
         "az immateriális javak"),
        ("ertekhelyesbites_targyi_van", "ertekhelyesbites_targyi_eszkoz",
         "a tárgyi eszközök"),
        ("ertekhelyesbites_penzugyi_van", "ertekhelyesbites_penzugyi_eszkoz",
         "a befektetett pénzügyi eszközök"),
    ):
        if not txt(a, gate):
            hianyzo_kapu = True
        if answer(a, gate) == "igen":
            reszlet = txt(a, text_key)
            csoportok.append(f"{cimke} körében{f' ({reszlet})' if reszlet else ''}")
    if csoportok:
        ertekhelyesbites = (
            "Az értékhelyesbítés alkalmazására vonatkozó lehetőséggel a "
            f"Vállalkozás {join_clauses(csoportok, ', ')} él."
        )
        szabalyok = txt(a, "szv64")
        hatar = txt(a, "szv65")
        if szabalyok:
            ertekhelyesbites += (
                " Az értékhelyesbítés alkalmazásának szabályai: "
                f"{as_sentence(szabalyok)}"
            )
        if hatar:
            ertekhelyesbites += (
                " Az értékhelyesbítés elszámolásához jelentősnek tekintett "
                f"értékhatár: {as_sentence(hatar)}"
            )
    else:
        ertekhelyesbites = (
            "Az értékhelyesbítés alkalmazására vonatkozó lehetőséggel a "
            "Vállalkozás nem kíván élni."
        )
    # A sablon a VIII. fejezetben eszközcsoportonként MÉG EGYSZER kimondja
    # ugyanezt (705., 721., 723.), jelölő nélkül - ezeket ugyanabból a
    # válaszból írjuk újra (a 705. bekezdés sablonbeli elgépelése -
    # "értékhelyesbítési nem alkalmazzuk" - így is javul).
    for idx, gate, cimke in (
        (705, "ertekhelyesbites_immat_van",
         "Vállalkozásunknál az immateriális javakkal kapcsolatos "
         "értékhelyesbítési lehetőséget"),
        (721, "ertekhelyesbites_targyi_van",
         "Vállalkozásunknál a tárgyi eszközeinkkel kapcsolatos "
         "értékhelyesbítési lehetőséget"),
        (723, "ertekhelyesbites_penzugyi_van",
         "A befektetett pénzügyi eszközök értékhelyesbítési lehetőségét"),
    ):
        alkalmaz = "alkalmazzuk." if answer(a, gate) == "igen" else "nem alkalmazzuk."
        restate(p[idx], f"{cimke} {alkalmaz}")

    rewrite_fact(p[774], p[775], ertekhelyesbites, "")
    if hianyzo_kapu and not csoportok:
        flag_before(
            p[775],
            "nincs válasz arra, alkalmazzák-e az értékhelyesbítést a "
            "befektetett eszközök valamely körére (Szv#63) - a 'nem "
            "alkalmazzuk' ág lett feltételezve, kérjük "
            "ellenőrizni/pontosítani.",
        )

    # Szv#78/#100 - valós értéken történő értékelés.
    piaci = txt(a, "szv78")
    valos = txt(a, "szv100")
    if (piaci and not nemleges(piaci)) or (valos and not nemleges(valos)):
        valos_mondat = join_sentences([
            f"A valós értéken történő értékelés alkalmazása: {valos}"
            if valos else "",
            "A befektetett eszközök piaci értéken történő értékelése: "
            f"{piaci}" if piaci else "",
        ])
    else:
        valos_mondat = (
            "A Vállalkozás a valós értéken történő értékelést, illetve a "
            "befektetett eszközök piaci értéken történő értékelését további "
            "külön döntésig nem alkalmazza."
        )
    rewrite_fact(p[779], p[780], valos_mondat, "")

    # Szv#69/#72/#73/#74 - az értékcsökkenés elszámolásának módszere és a
    # tervezéshez szükséges adatok (a megelőző bekezdés törvényi szöveg).
    fill_or_flag(
        p[792],
        join_sentences([
            f"Az alkalmazott leírási módszer: {txt(a, 'szv69')}"
            if txt(a, "szv69") else "",
            "A terv szerinti értékcsökkenés megállapításához szükséges "
            "adatokat (hasznos élettartam, maradványérték, üzembe helyezés "
            f"időpontja) a következők szerint határozzuk meg: {txt(a, 'szv72')}"
            if txt(a, "szv72") else "",
            f"A maradványérték meghatározásának módja: {txt(a, 'szv73')}"
            if txt(a, "szv73") else "",
            f"A maradványértéket nullának tekintjük: {txt(a, 'szv74')}"
            if txt(a, "szv74") else "",
        ]),
        "nincs válasz az értékcsökkenés leírási módszerére, a tervezéshez "
        "szükséges adatokra és a maradványérték meghatározására "
        "(Szv#69, #72, #73, #74).",
    )

    # Szv#61 - üzleti vagy cégérték. UGYANEZT a tényt a sablon a VIII.
    # fejezetben (702.) is kimondja, jelölő NÉLKÜL - ha csak a jelölt
    # helyet oldanánk fel, a dokumentum két helyen mást állítana.
    if answer(a, "van_cegertek") == "igen":
        restate(p[702], "A vállalkozás üzleti vagy cégértéket mutat ki.")
        restate(
            p[799],
            "A befejezett kísérleti fejlesztés, a befejezett "
            "alapítás-átszervezés, illetve az üzleti vagy cégérték aktivált "
            "értékének leírását a fentiek szerint végezzük.",
        )
    elif answer(a, "alapitas_atszervezes_aktivalja") == "igen":
        restate(
            p[799],
            "A befejezett alapítás-átszervezés aktivált értékét társaságunk "
            "kimutatja, annak leírását a fentiek szerint végezzük.",
        )
    if answer(a, "van_cegertek") == "igen":
        ev = txt(a, "cegertek_leiras_ev")
        cegertek = (
            "A Vállalkozás üzleti vagy cégértéket mutat ki, amelyet "
            f"{ev} év alatt ír le." if is_number(ev) else ""
        )
        if cegertek and answer(a, "van_negativ_cegertek") == "igen":
            nev = txt(a, "negativ_cegertek_leiras_ev")
            if is_number(nev):
                cegertek += (
                    " A negatív üzleti vagy cégértéket a Vállalkozás "
                    f"{nev} év alatt írja le."
                )
        note = (
            "a Vállalkozásnak van üzleti/cégértéke (Szv#61), de a leírási "
            "idő hiányzik vagy nem szám - az alábbi mondat ennek az "
            "ELLENKEZŐJÉT állítja, kérjük kézzel javítani."
        )
    else:
        cegertek = "Az üzleti vagy cégértékkel a Vállalkozás nem rendelkezik."
        note = ""
    rewrite_fact(p[801], p[802], cegertek, note)

    # Szv#80 / Szv#79 - értékvesztés hitelviszonyt megtestesítő
    # értékpapíroknál, illetve részesedéseknél (a megelőző bekezdések
    # törvényi szövegek).
    fill_or_flag(
        p[816],
        "A hosszú lejáratú, hitelviszonyt megtestesítő értékpapírok "
        f"értékvesztésének elszámolási szabályai: {as_sentence(txt(a, 'szv80'))}"
        if txt(a, "szv80") else "",
        "nincs válasz a hosszú lejáratú hitelviszonyt megtestesítő "
        "értékpapírok értékvesztési szabályaira (Szv#80).",
    )
    fill_or_flag(
        p[830],
        "A gazdasági társaságban lévő befektetések értékvesztésének "
        f"elszámolási szabályai: {as_sentence(txt(a, 'szv79'))}" if txt(a, "szv79") else "",
        "nincs válasz a gazdasági társaságban lévő befektetések "
        "értékvesztési szabályaira (Szv#79).",
    )

    # Szv#87 - a követelésekre elszámolt értékvesztés mértékei. A sablon
    # konkrét, MÁS cégre írt mértéket tartalmaz ("365 napon túl ... 50%").
    rewrite_fact(
        p[839], p[840],
        "A követelések minősítésének feltételrendszere, az elszámolandó "
        f"értékvesztés mértéke: {as_sentence(txt(a, 'szv87'))}" if txt(a, "szv87") else "",
        "nincs válasz arra, hogyan, milyen mértékek alkalmazásával "
        "határozzák meg a követelésekre elszámolandó értékvesztést "
        "(Szv#87) - az alábbi bekezdés a sablon konkrét, MÁS cégre írt "
        "mértékeit tartalmazza.",
    )


# --------------------------------------------------------------------------
# XIII-XVII. fejezet - árfolyamok, céltartalék, elhatárolások, eredmény
# --------------------------------------------------------------------------

def resolve_arfolyamok(p, a):
    # Szv#91 - a devizás tételek forintra átszámításához használt árfolyam.
    rewrite_fact(
        p[846], p[847],
        "A forintérték meghatározásakor a vállalkozás a következő árfolyamot "
        f"alkalmazza: {as_sentence(txt(a, 'szv91'))}" if txt(a, "szv91") else "",
        "nincs válasz arra, melyik bank melyik árfolyamát használják a "
        "devizás tételek forintértékének meghatározásához (Szv#91) - az "
        "alábbi mondat a sablon konkrét, MÁS cégre írt választását "
        "tartalmazza.",
    )

    # Ért#3 - az áfa alapjának megállapításához használt árfolyam
    # alkalmazása. A bekezdés második mondata ez a döntés.
    rewrite_fact(
        p[849], p[850],
        "Ha a Magyar Nemzeti Bank által nem jegyzett és nem konvertibilis "
        "valutát, ilyen valutára szóló eszközöket és kötelezettségeket kell "
        "forintra átszámítani, a valuta szabadpiaci árfolyamát használja a "
        "vállalkozás. A számviteli törvény 60. § (5a) bekezdése szerinti "
        "lehetőséggel (az általános forgalmi adó alapjának megállapításához "
        f"használt árfolyam alkalmazása) kapcsolatban: {as_sentence(txt(a, 'ert3'))}"
        if txt(a, "ert3") else "",
        "nincs válasz arra, élnek-e a számviteli törvény 60. § (5a) "
        "bekezdése szerinti árfolyam-választási lehetőséggel (Ért#3) - az "
        "alábbi bekezdés a sablon konkrét, MÁS cégre írt választását "
        "tartalmazza.",
    )

    # Szv#93 - a devizás pénzeszközök csökkenésének árfolyama.
    rewrite_fact(
        p[854], p[855],
        "A devizás pénzeszközök csökkenésekor a forintértéket a következők "
        f"szerint határozzuk meg: {as_sentence(txt(a, 'szv93'))}" if txt(a, "szv93") else "",
        "nincs válasz arra, átlag- vagy FIFO árfolyamot használnak a devizás "
        "pénzeszközök csökkenésénél (Szv#93) - az alábbi mondat a sablon "
        "konkrét, MÁS cégre írt választását tartalmazza.",
    )


def resolve_celtartalek(p, a):
    """XIV. Céltartalékképzés (Szv#96). A sablon kereken kijelenti, hogy
    "A társaság nem képez céltartalékot." - ezt a válaszból írjuk újra."""
    if answer(a, "celtartalek_kepez") == "igen":
        hatar = txt(a, "celtartalek_jelentos_hatar")
        elteres = txt(a, "celtartalek_elteres_szazalek")
        mondat = (
            "A jövőbeni költségekre a Vállalkozás céltartalékot képez, "
            "amennyiben azok összege jelentős."
        )
        if hatar:
            mondat += (
                " A képezhető céltartalék szempontjából a költségeket "
                f"jelentősnek minősítjük, ha azok összege eléri a "
                f"{hatar} értéket."
            )
        if elteres:
            mondat += (
                " Az egyes jogcímek szerint képzett céltartalék összegét az "
                "előző évi összegtől lényegesen eltérőnek minősítjük, "
                f"amennyiben a változás eléri a {elteres.rstrip('%').strip()}%-ot."
            )
    else:
        mondat = (
            "A jövőbeni költségekre a Vállalkozás nem képez céltartalékot "
            "akkor sem, ha azok összege jelentős."
        )
    # A sablon a VIII. fejezetben (758.) jelölő nélkül is kimondja ugyanezt.
    if answer(a, "celtartalek_kepez") == "igen":
        restate(
            p[758],
            "A jövőbeni költségekre céltartalékot képezünk, amennyiben azok "
            "összege jelentős.",
        )
    rewrite_fact(p[866], p[867], mondat, "")
    if not txt(a, "celtartalek_kepez"):
        flag_before(
            p[867],
            "nincs válasz arra, képez-e a vállalkozás lehetőség alapján "
            "céltartalékot (Szv#96) - a 'nem képez' ág lett feltételezve, "
            "kérjük ellenőrizni/pontosítani.",
        )


def resolve_elhatarolasok(p, a):
    """XV-XVI. fejezet - időbeli elhatárolások, projektelszámolás."""
    # Szv#99 - az időbeli elhatárolások dokumentálása.
    rewrite_fact(
        p[872], p[873],
        "Az aktív és passzív időbeli elhatárolások dokumentálása és "
        f"elszámolása a következők szerint történik: {as_sentence(txt(a, 'szv99'))}"
        if txt(a, "szv99") else "",
        "nincs válasz arra, hogyan dokumentálják az időbeli elhatárolás "
        "tételeit (Szv#99) - az alábbi mondat a sablon konkrét, MÁS cégre "
        "írt gyakorlatát írja le.",
    )

    # Szv#83/#103/#104 - a szerződés elszámolási egysége (projektelszámolás).
    # A sablon 896-899. bekezdései EGYÜTT írják le a ZalaElektro gyakorlatát,
    # ezért a teljes blokkot a válaszból építjük újra.
    if answer(a, "szerzodes_elszamolasi_egyseg",
              "nem_alkalmazzuk") == "alkalmazzuk":
        projekt = join_sentences([
            "A számviteli törvény 2020. évi módosításának megfelelően "
            "Társaságunk alkalmazza a szerződés elszámolási egységére "
            "(projektelszámolásra) vonatkozó szabályokat.",
            "A teljesítési fokot a következő módszerrel határozzuk meg: "
            f"{txt(a, 'szv103')}" if txt(a, "szv103") else "",
            "A készültségi fokot a következő módszerrel határozzuk meg: "
            f"{txt(a, 'szv104')}" if txt(a, "szv104") else "",
            "A teljes szerződéses ellenérték mérlegfordulónapi teljesítési "
            "fok arányában számított összege, valamint a már elszámolt "
            "árbevétel közötti különbözetet aktív vagy passzív időbeli "
            "elhatárolásként mutatjuk ki, az árbevétel növelésével vagy "
            "csökkentésével egyidejűleg.",
        ])
    else:
        projekt = (
            "Társaságunk a nagy tömegben, sorozatban előállított termékek "
            "esetében a szerződés elszámolási egységére vonatkozó "
            "szabályrendszert nem alkalmazza."
        )
    rewrite_fact(p[896], p[897], projekt, "")
    delete(p[898])
    delete(p[899])
    if not txt(a, "szerzodes_elszamolasi_egyseg"):
        flag_before(
            p[897],
            "nincs válasz arra, alkalmazzák-e a szerződés elszámolási "
            "egységére vonatkozó szabályrendszert (Szv#83) - a 'nem "
            "alkalmazzuk' ág lett feltételezve, kérjük "
            "ellenőrizni/pontosítani.",
        )


# --------------------------------------------------------------------------
# XVIII. fejezet - a sablonból átvett további pontok (nyilatkozatok)
#
# Ezek a bekezdések a törvényi hátteret ismertetik, majd az OLVASÓNAK szóló
# felszólítással zárulnak ("A Vállalkozásnak nyilatkoznia kell arról,
# hogy ..."). A felszólítás nem mehet ki az ügyfélhez: levágjuk
# (truncate_at), és a helyére a tényleges nyilatkozat kerül a jelölőbe.
# --------------------------------------------------------------------------

def resolve_nyilatkozatok(p, a):
    # Ért#4 - a hitelfelvételhez kapcsolódó költségek elhatárolása.
    truncate_at(p[919], "A Vállalkozásnak nyilatkoznia kell")
    if igenlo(answer(a, "ert4_alkalmazza_elhatarolas")):
        hatar = txt(a, "bizomanyi_dij_hatar")
        ert4 = (
            "A Vállalkozás él az elhatárolás lehetőségével: a hitel, kölcsön "
            "felvételéhez közvetlenül kapcsolódóan felmerült költségeket "
            "aktív időbeli elhatárolásként mutatja ki."
        )
        if hatar:
            ert4 += (
                " A bizományi díjat, vásárolt vételi opció díját jelentős "
                f"összegűnek tekintjük, ha eléri a {hatar} E Ft értéket."
            )
    else:
        ert4 = (
            "A Vállalkozás - a díjak összegétől függetlenül - nem él az "
            "elhatárolás lehetőségével, a hitel, kölcsön felvételéhez "
            "kapcsolódó költségeket a felmerüléskor költségként, "
            "ráfordításként számolja el."
        )
    fill_blanks(p[920], [ert4])
    if not txt(a, "ert4_alkalmazza_elhatarolas"):
        flag_before(
            p[920],
            "nincs válasz arra, élnek-e a hitelfelvételhez kapcsolódó "
            "költségek elhatárolásának lehetőségével (Ért#4) - a 'nem élünk "
            "vele' ág lett feltételezve, kérjük ellenőrizni/pontosítani.",
        )

    # Szv#29 - fenntarthatósági jelentés.
    truncate_at(p[923], "A Vállalkozásnak ennek fennállását")
    fenn = answer(a, "fenntarthatosagi_jelentes", "")
    fenn_mondat = {
        "nem_koteles":
            "A Vállalkozás a fenti feltételek alapján fenntarthatósági "
            "jelentés készítésére nem kötelezett.",
        "igen_ugyanaz_a_konyvvizsgalo":
            "Vállalkozásunk fenntarthatósági jelentés készítésére "
            "kötelezett, amelynek vizsgálatával fenntarthatósági "
            "minősítéssel rendelkező olyan kamarai tag könyvvizsgálót, "
            "könyvvizsgáló céget bízunk meg, aki az éves beszámoló "
            "könyvvizsgálatát is végzi.",
        "igen_mas_konyvvizsgalo":
            "Vállalkozásunk fenntarthatósági jelentés készítésére "
            "kötelezett, amelynek vizsgálatával az éves beszámoló "
            "jogszabályi kötelezettségen alapuló könyvvizsgálatával "
            "megbízott kamarai tag könyvvizsgálótól, könyvvizsgáló cégtől "
            "eltérő kamarai tag könyvvizsgálót, könyvvizsgáló céget bízunk "
            "meg.",
    }.get(fenn, "")
    fill_or_flag(
        p[924], fenn_mondat,
        "nincs válasz arra, van-e fenntarthatósági jelentés készítési "
        "kötelezettsége a társaságnak (Szv#29).",
    )

    # Szv#66/#67/#68 - környezetvédelmi közzétételek.
    truncate_at(p[927], "A Vállalkozásnak nyilatkoznia kell")
    fill_or_flag(
        p[928],
        join_sentences([
            "A környezet védelmét közvetlenül szolgáló tárgyi eszközök "
            f"tekintetében: {txt(a, 'szv66')}" if txt(a, "szv66") else "",
            f"Veszélyes hulladék tekintetében: {txt(a, 'szv67')}"
            if txt(a, "szv67") else "",
            f"Környezetvédelmi kötelezettség tekintetében: {txt(a, 'szv68')}"
            if txt(a, "szv68") else "",
        ]),
        "nincs válasz a környezetvédelemmel összefüggő tárgyi eszközökre, a "
        "veszélyes hulladékra és a környezetvédelmi kötelezettségre "
        "(Szv#66, #67, #68).",
    )

    # Szv#75/#76/#84 - a terven felüli értékcsökkenés / értékvesztés
    # jelentőségi határai. A 932. bekezdés kategóriánként 20%-os
    # sablon-alapértéket sorol fel; csak a kérdőívvel lefedett három
    # kategóriát írjuk felül, a többinél marad a 20%, mert az érvényes
    # alapértelmezés, nem kitöltetlen hely.
    set_plain_run_text(
        p[931], 0,
        "A könyv szerinti érték és a piaci érték közötti különbséget az "
        "alábbi mértékek felett tekintjük jelentősnek:",
    )
    mertekek = []
    for cimke, kulcs in (
        ("Immateriális javak", "szv75"),
        ("Tárgyi eszközök", None),
        ("Befektetett pénzügyi eszközök", "szv76"),
        ("Készletek", "szv84"),
        ("Követelések", None),
        ("Értékpapírok", None),
        ("Pénzeszközök", None),
        ("Részesedések", None),
    ):
        ertek = txt(a, kulcs).rstrip("%").strip() if kulcs else ""
        mertekek.append(f"{cimke}: {ertek or '20'}%")
    fill_run(p[932], 0, "  |  ".join(mertekek))
    delete(p[934])

    # Szv#94 - a kamatozó értékpapír névérték-különbözetének elhatárolása.
    # A sablon szövege itt konkrétan a "Zala-Elektro Kft."-t nevezi meg.
    if answer(a, "ertekpapir_kamatkulonbozet_elhatarolas",
              "nem_hataroljuk_el") == "elhataroljuk":
        ep = (
            "A kamatozó értékpapír névértéke és beszerzési értéke közötti - "
            "nyereségjellegű - különbözetből a beszerzéstől az üzleti év "
            "mérlegfordulónapjáig terjedő időszakra időarányosan jutó "
            "összeget a Vállalkozás aktív időbeli elhatárolásként mutatja ki."
        )
    else:
        ep = (
            "A kamatozó értékpapír névértéke és beszerzési értéke közötti - "
            "nyereségjellegű - különbözetből a beszerzéstől az üzleti év "
            "mérlegfordulónapjáig terjedő időszakra időarányosan jutó "
            "összeget a Vállalkozás nem határolja el."
        )
    rewrite_fact(p[937], p[938], ep, "")
    if not txt(a, "ertekpapir_kamatkulonbozet_elhatarolas"):
        flag_before(
            p[938],
            "nincs válasz arra, elhatárolják-e a kamatozó értékpapír "
            "névértéke és beszerzési értéke közötti különbözetet (Szv#94) - "
            "a 'nem határoljuk el' ág lett feltételezve, kérjük "
            "ellenőrizni/pontosítani.",
        )

    # Szv#108 - társasági adó / KIVA.
    truncate_at(p[941], "A Vállalkozásnak rögzítenie kell")
    fill_or_flag(
        p[942],
        f"A Vállalkozás adóalanyisága: {as_sentence(txt(a, 'szv108'))}"
        if txt(a, "szv108") else "",
        "nincs válasz arra, társasági adó vagy KIVA alanyok-e (Szv#108).",
    )

    # Szv#110 - globális minimumadó (GloBE).
    truncate_at(p[945], "A sablon alapértelmezése")
    fill_or_flag(
        p[946],
        "A Vállalkozás GloBE törvény szerinti kötelezettsége: "
        f"{as_sentence(txt(a, 'szv110'))}" if txt(a, "szv110") else "",
        "nincs válasz arra, van-e a vállalkozásnak globális minimumadó "
        "(GloBE) kötelezettsége (Szv#110).",
    )


def resolve_zaradek(p, a):
    """ZÁRADÉK - a számviteli politika elválaszthatatlan részét képező
    szabályzatok felsorolása (Szv#106).

    Ha a Vállalkozás mentesül az önköltségszámítási szabályzat készítése
    alól, akkor annak a felsorolásból is ki kell esnie - a 954. bekezdés a
    felsorolás EGY tétele (önálló bekezdés), így egészében törlendő.

    UGYANEZ a szabályzat az I.3. fejezet felsorolásában is szerepel (87.),
    jelölő nélkül - onnan is ki kell esnie, különben a dokumentum két helyen
    mást mondana."""
    van_szabalyzat = (
        answer(a, "onkoltsegszamitasi_szabalyzat", "mentesul") != "mentesul"
    )
    keep_if(p[87], van_szabalyzat, None)
    keep_if(p[954], van_szabalyzat, None)
    delete(p[955])


# --------------------------------------------------------------------------
# Fő generálási logika
# --------------------------------------------------------------------------

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

    resolve_cegadatok(p, answers)
    resolve_ecs_ertekhatar(p, answers)
    resolve_szamviteli_felelosok(p, answers)
    resolve_leltar(p, answers)
    resolve_penzkezeles(p, answers)
    resolve_konyvvizsgalat(p, answers)
    resolve_kozzetetel(p, answers)
    resolve_beszamolo_forma(p, answers)
    resolve_merleg_forma(p, answers)
    resolve_eredmenykimutatas_forma(p, answers)
    resolve_tetelek_modositasa(p, answers)
    resolve_beszamolo_nyelve(p, answers)
    resolve_beszamoloert_felelosseg(p, answers)
    resolve_kiegeszito_melleklet(p, answers)
    resolve_konszolidacio(p, answers)
    resolve_minositesek(p, answers)
    resolve_konyvvezetes(p, answers)
    resolve_bizonylatok(p, answers)
    resolve_zarlat(p, answers)
    resolve_befektetett_eszkozok(p, answers)
    resolve_ertekeles(p, answers)
    resolve_arfolyamok(p, answers)
    resolve_celtartalek(p, answers)
    resolve_elhatarolasok(p, answers)
    resolve_nyilatkozatok(p, answers)
    resolve_zaradek(p, answers)

    # Biztonsági háló: a VÉGLEGES bekezdéslistán (a törlések/beszúrások
    # után) nézzük meg, maradt-e kitöltetlen sablonjelölő.
    sweep_leftover_markers(doc.paragraphs)

    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(output_path))
    print(f"Elkészült: {output_path}")


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Használat: generate_policy.py <answers.json> <output.docx>")
        sys.exit(1)
    generate(sys.argv[1], sys.argv[2])
