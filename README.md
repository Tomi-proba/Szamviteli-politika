# Számviteli politika – automatikus vázlatgenerátor

Ez a projekt a 2026-os számviteli politika sablon és az ügyfél-kérdőív
alapján **automatikusan legenerál egy Word-vázlatot**, amit a kollégák
átnéznek/pontosítanak, mielőtt kiküldik az ügyfélnek.

> **2026: teljes sablonváltás.** A `sablon/merged.docx` egy TELJESEN ÚJ,
> átszerkesztett dokumentum (982 bekezdés a korábbi 2130 helyett, kb. 26%
> szöveg-egyezéssel), ezért a motor `resolve_*` logikája a nulláról íródott
> újra. A régi sablon `sablon/merged_2025_v1.docx` néven megmaradt
> hivatkozásnak. Az új sablonban **nincs egyetlen VAGY-alternatíva sem** és
> nincs kitöltendő táblázat sem; helyettük 82 db `▶ Adatbékérő: Sz.pol. N.
> kerdes` jelölő, 7 db formázatlan „innen hiányzik" helykitöltő, és a
> jelölők közel felénél egy MÁR MEGÍRT, konkrét (a ZalaElektro-vázlatból
> származó) mondat áll, amit a válaszból újra kell írni. Ld. az
> „Az új sablon három mechanizmusa" fejezetet.

## Két külön oldal - az ügyfél és az iroda soha nem lát ugyanazt

A `webapp-client/` mappa egy teljesen önálló, kliens-oldali (böngészőben
futó) rendszert épít, KÉT külön HTML-fájlból - nincs szükség Python-ra vagy
szerverre a végfelhasználók gépén:

```bash
python3 webapp-client/build.py
# -> webapp-client/dist_ugyfel.html
# -> webapp-client/dist_iroda.html
```

**`dist_ugyfel.html` - ezt küldjétek ki az ügyfélnek.** Az összes kérdést
tartalmazza - nincs külön, csak-az-irodán-belül-eldönthető kérdéscsoport,
mert a gyakorlatban minden választ az ügyféltől kérnek meg. (Korábban 15
kérdés `forras: "iroda"` volt - ezek a sablon egy-egy döntési pontját
professzionális minősítésként modellezték -, de ez nem illett a tényleges
munkafolyamathoz, ezért mind át lettek sorolva `forras: "ugyfel"`-re; a
`questions.json`-ban technikailag továbbra is létezik a `forras` mező, csak
jelenleg minden kérdésnél `"ugyfel"` az értéke.) Az oldal **nem generál semmilyen
dokumentumot** - a kitöltés végén egy kódot ad, amit az ügyfél e-mailben
visszaküld (vagy fájlba menthető, ld. lentebb). Az ügyfél ebből az oldalból
semmit nem tud kikövetkeztetni az automatizálásról vagy a végleges
dokumentumról.

**`dist_iroda.html` - ez a belső, jelszóval védett eszköz.** Itt illeszthető
be az ügyféltől kapott kód/fájl (ez előtölti az ügyfél válaszait), itt
ellenőrizhető/pontosítható bármelyik mező kiküldés előtt, és csak itt lehet
legenerálni a Word-vázlatot. **A jelszó alapértelmezetten `Politika2026!`**
- ez KIZÁRÓLAG visszatartás, nem valódi hozzáférés-vezérlés (statikus oldal,
nincs szerver, egy technikailag értő személy a forráskódból kiolvashatná a
hash-elés módját). Cseréhez:

```bash
python3 webapp-client/hash_password.py "ÚjJelszó"
# a kiírt hash-t írd be az iroda_source.html PASSWORD_HASH konstansába,
# majd futtasd újra a build.py-t
```

**Helyi Python-os webform** (ha valaki tud terminált használni, és nem kell
a kétoldali ügyfél/iroda szétválasztás):

```bash
python3 scripts/webapp.py
```

Ezután a böngészőben: **http://localhost:8000** - ugyanaz az (egybeépített,
jelszó nélküli) kérdőív, Python-szerveren keresztül futtatva
(`scripts/generate_policy.py`). Ez a régebbi, egyoldalas változat - inkább
fejlesztés/tesztelés közben hasznos, éles használatra a fenti két HTML-oldal
ajánlott.

**Fontos:** mindegyik változat UGYANAZT a motorlogikát futtatja (a Python és
a JS változat tesztekkel ellenőrzötten pontosan azonos kimenetet ad ugyanazon
válaszokra). Ha a sablon szövege változik, MINDKÉT oldali (Python + JS)
`resolve_*` logikát frissíteni kell, mert a bekezdés-indexek mindkettőben
hard-code-olva vannak.

## Hogyan működik belül

```
data/questions.json          <- a kérdőív definíciója (mit kérdezünk az ügyféltől / mit tud a könyvelő)
data/answers.example.json    <- egy kitöltött minta-válaszkészlet (teszteléshez)
data/answers.example2.json   <- második minta, az ELLENTÉTES ágakat gyakorolja
sablon/szamviteli_politika_2026_uj_sablon.docx <- a 2026-os, ÚJ sablon eredeti (nyers) feltöltött fájlja
sablon/merged.docx           <- ennek "megtisztított" (összevont futású) másolata - EZT olvassa a motor
sablon/merged_2025_v1.docx   <- a KORÁBBI (2025-ös) sablon run-merged másolata, hivatkozásnak
sablon/szamviteli_politika_JELOLT*.docx   <- a korábbi JELÖLT sablonok (hivatkozás, már nem használt)
scripts/generate_policy.py   <- a fő motor (Python): válaszok + sablon -> kitöltött Word-vázlat
scripts/webapp.py            <- helyi Python webform a motor fölé (egyoldalas, jelszó nélküli)
scripts/dev/parity_check.py  <- Python <-> JS motor-paritás teszt (Playwright/Chromium)
scripts/dev/content_check.py <- tartalmi ellenőrzés a legenerált vázlaton
webapp-client/zip.js         <- minimál ZIP olvasó/író böngészőben (CompressionStream alapú)
webapp-client/docx.js        <- docx bekezdés/futás-szintű segédfüggvények böngészőben
webapp-client/resolve.js     <- a generate_policy.py resolve_* függvényeinek JS portja
webapp-client/ugyfel_source.html <- az ügyfél-oldal sablonja (build.py egészíti ki adattal)
webapp-client/iroda_source.html  <- az iroda-oldal sablonja (jelszó-kapu + kódimport)
webapp-client/hash_password.py   <- jelszó-hash generátor az iroda-oldalhoz
webapp-client/build.py       <- összeállítja a publikálható dist_ugyfel.html / dist_iroda.html fájlokat
output/                      <- ide kerülnek a Python scripttel legenerált vázlatok
```

### A válaszok átadásának működése

Az ügyfél-oldal a válaszokat vagy egy **olvasható fájlba** menti, vagy - ha
valaki mégis a kódot választja - egy base64-be kódolt JSON-ná ("kód")
alakítja. Az irodához háromféleképp lehet eljuttatni:

- **Mentés fájlba** (ajánlott) - egy `.txt` fájlt tölt le (a cégnév +
  dátum alapján elnevezve), amit pl. a közös hálózati meghajtóra lehet
  menteni. Megnyitva a fájl elején rögtön olvasható, ember számára
  értelmezhető formában látszik minden kérdés és a rá adott válasz - a
  fájl végén egy géppel olvasható JSON-blokk (a `<<<ADATOK...>>>` jelző
  után) teszi lehetővé a visszatöltést, ezt nem kell/nem szabad
  szerkeszteni. Az iroda-oldalon a "Fájl megnyitása..." gombbal tallózható
  be - nincs kód-másolgatás/beillesztés.
- E-mailben (a "Küldés e-mailben" gomb egy előre kitöltött e-mailt nyit meg
  a kóddal).
- A kód kimásolásával/beillesztésével (mindkét oldalon egy lenyitható
  "...vagy a kód" részben érhető el, ha valaki mégis ezt az utat választja).

Mindhárom út ugyanahhoz az előtöltéshez vezet az iroda-oldalon - a
betöltő függvény (`parseAnswersFileText`) az olvasható fájlformátumot, a
régebbi sima JSON-mentést és a base64 "kódot" is felismeri, tehát egy
korábban mentett fájl is visszatölthető marad. Ha az irodának saját, fix
e-mail címe van, ahova a kódot kéritek, írd be az `ugyfel_source.html`
`OFFICE_EMAIL` konstansába, majd futtasd újra a `build.py`-t - így a
"Küldés e-mailben" gomb már eleve ki lesz címezve.

Mindhárom mentési pont (ügyfél félbehagyott mentése, ügyfél véglegesített
mentése, iroda mentése) **azonos fájlformátumot ír** - ezért bármelyik
ilyen fájl bármelyik oldalon visszatölthető:

- **Ügyfél-oldal**: a "Mentés és folytatás később" gomb bármikor (akár
  félbehagyott kitöltésnél is) fájlba menti az aktuális válaszokat; a lap
  tetején a "Korábbi mentés megnyitása..." ezt tölti vissza, hogy onnan
  folytatható legyen a kitöltés.
- **Iroda-oldal**: a "Válaszok mentése fájlba" gomb az aktuális állapotot
  menti (az ügyféltől importált ÉS az iroda által azóta kitöltött mezőket
  is) - így megszakítható/átadható a munka kollégák között, és mivel a
  fájlnév cégnév+dátum alapján generálódik, egy közös mappában több cég
  mentése is jól megkülönböztethető.

Parancssorból (JSON válaszfájlból), ha nem a webes űrlapot használod:

```bash
python3 scripts/generate_policy.py data/answers.example.json output/valami.docx
```

A script:
1. **Kitölti** a sablon 82 db `▶ Adatbékérő: Sz.pol. N. kerdes` jelölőjét és
   a 7 db formázatlan „innen hiányzik" helykitöltőt a kérdőív válaszaiból.
2. Ahol a jelölő MELLETT a sablon már kimond egy konkrét, egy MÁSIK cégre
   írt tényt (a jelölők közel felénél), ott **az egész mondatot újraírja** a
   válaszból, és a sablon eredeti mondatát törli - különben a dokumentum
   önmagának mondana ellent (ld. „D1" fejezet).
3. A sablon szerzőjének a szövegbe írt, feltételes szerkesztői utasításait
   (`ezt csak akkor írja ha …`) **mindig kiveszi**, és a hozzájuk tartozó
   mondatot a válasz alapján megtartja vagy törli (ld. „D2" fejezet).
4. Eltávolítja a `[Szv N. kérdés]` / `[Ért N. kérdés]` hivatkozásokat, és a
   XVIII. fejezetben levágja az OLVASÓNAK szóló felszólításokat
   („A Vállalkozásnak nyilatkoznia kell arról, hogy …"), amelyek helyére a
   tényleges nyilatkozat kerül.
5. A **behelyettesített válaszszöveget** a sablon eredeti "kitöltendő hely"
   jelölésével (sárga kiemelés, áthúzás, piros betűszín) hagyja - ez jelzi a
   felülvizsgáló kollégának, mely szövegrészek kerültek be automatikusan, és
   melyeket kell átnéznie kiküldés előtt. A bekezdés egyéb részein
   megszünteti a "még döntendő" betűszínt, hogy csak a tényleges
   behelyettesítés maradjon kiemelve.
6. Ahol a válasz nem egyértelmű (pl. "nincs cégérték" vagy "nem aktiválunk",
   miközben a sablon egy leírási időt kérne), **nem tippel** - egy jól
   látható, félkövér piros "⚠ ELLENŐRIZENDŐ" jelölést szúr be a bekezdés elé.
7. **A KITÖLTETLEN mező sosem egyenlő a "nem" válasszal.** Néhány igen/nem
   jellegű kérdésnél (honlapon közzététel, értékhelyesbítés
   eszközcsoportonként, céltartalék-képzés, alapítás-átszervezés aktiválása,
   szerződés elszámolási egysége, értékpapír-elhatárolás, árfolyamveszteség)
   a sablonban nincs olyan harmadik ág, ami "nincs válasz"-t jelentene - ha a
   mező üresen marad, a motor a "nem" ágat használja (mert ez a
   dokumentum-szerkezet szempontjából muszáj), **de ilyenkor mindig
   `⚠ ELLENŐRIZENDŐ` jelölést is beszúr**, hogy a felülvizsgáló lássa: ez nem
   az ügyfél tényleges válasza volt, csak feltételezés. Egy kifejezetten
   begépelt "nem" ugyanígy a "nem" ágra kerül, de flag NÉLKÜL, mert az
   valódi válasz.
8. A végén **biztonsági söprést** futtat: ha bármelyik jelölő vagy
   helykitöltő mégis bennmaradt volna, kiveszi a szövegét és jelölést hagy a
   helyén - a sablon belső jelölései sosem mehetnek ki az ügyfélhez.

## Lefedettség – mit tölt ki és mit NEM

A `data/questions.json` **a teljes, hivatalos adatbekérőt tartalmazza** (a
Számviteli politika + Értékelési szabályzat kérdéseit, pontosan a forrás
xlsx sorrendjében), kiegészítve néhány, a sablon szerkezetéhez igazodó
bontó/választó kérdéssel.

| | darab |
|---|---|
| kérdés összesen a `questions.json`-ban | **162** |
| ebből a motor automatikusan behelyettesíti / ág-választásra használja | **126** |
| ebből jelenleg nincs kitöltendő helye a sablonban (`_nincs_sablonhely`) | **36** |
| lefedett *forrás*-kérdés (Szv#/Ért# szám szerint) | **100** |
| `resolve_*` függvény (Python és JS oldalon egyaránt) | **26** |
| a sablon `▶ Adatbékérő` jelölője | **82** |
| ebből válaszból kitöltve / mondattá újraírva | **80** |
| ebből szándékosan `⚠ ELLENŐRIZENDŐ` jelöléssel hagyva | **2** (Szv#40, Szv#42) |
| formázatlan „innen hiányzik" helykitöltő | **7** (mind kitöltve) |

**A 82 jelölőből egy sem marad a kimenetben** - amit nem tudunk kitölteni,
azt is töröljük és jelöléssel helyettesítjük, mert a jelölő szövege
("▶ Adatbékérő: Sz.pol. 44. kerdes") sosem mehet ki az ügyfélhez. Ugyanez áll
a 7 „innen hiányzik" helykitöltőre és a sablon szerzőjének a szövegbe írt
utasításaira. Ezt a `scripts/dev/content_check.py` MINDEN válaszkészletnél
ellenőrzi (`TILTOTT` lista), és a generálás végén egy biztonsági söprés
(`sweep_leftover_markers()`) is fut.

Az új sablonhoz felvett kérdések: `gazdalkodasi_forma` (Szv#113),
`alapitas_idopontja` (Szv#112), `fioktelepek` (Szv#114),
`konyvvizsgalo_nyilvantartasi_szam` (Szv#115),
`konyvvizsgalatert_felelos_konyvvizsgalo` (Szv#116),
`beszamolo_kozzetel_datum` és `munkafolyamat_iranyitas_felelos` (az utóbbi
kettőnek nincs saját sorszáma az adatbekérőben, a kapcsolódó Szv#19 / Szv#20
mellé kerültek). A `cegjegyzekszam` `eredeti` mezője `Szv#6`-ról `Szv#111`-re
változott, mert az új adatbekérő már külön kérdésként kérdezi.

## Az új sablon három mechanizmusa

Az új sablon három olyan helyzetet hoz, amit a korábbi
`blank_groups()`/`fill_blanks()` páros nem tudott kezelni. Mindhárom
mechanizmus **mindkét motorban, azonos néven** létezik
(`scripts/generate_policy.py` ↔ `webapp-client/docx.js`).

### D1 – a jelölő MELLETT álló konkrét tény újraírása

A 82 jelölő közel felénél a sablon már kimond egy konkrét, **egy MÁSIK cégre
(a ZalaElektro-vázlatra) írt** tényt közvetlenül a jelölő előtti
bekezdésben - például *„Vállalkozásunk esetében Business Central Microsoft
rendszer segítségével történik a könyvvezetés."*, *„A könyvvezetésért
felelős: Tábor Ildikó gazdasági vezető (regisztráció száma: 174234)."* vagy
*„A társaság nem képez céltartalékot."*. Ha csak a jelölőt töltenénk ki, ez a
mondat ott maradna, és **ellentmondana a válasznak**.

A `rewrite_fact()` / `rewriteFact()` ezért a TELJES mondatot a válaszból
fogalmazza újra, és a sablon eredeti mondatát törli:

```python
def rewrite_fact(fact_paragraph, marker_paragraph, text, note, fallback=None):
    text = text_value(text).strip()
    if text:
        fill_blanks(marker_paragraph, [text])   # a jelölő helyére az ÚJ mondat
        delete(fact_paragraph)                  # a régi, MÁS cégre írt mondat törlődik
        return True
    if fallback:                                # nincs válasz, de a régi mondat
        fill_blanks(marker_paragraph, [fallback])  # azonosítható idegen adatot tartalmaz
        delete(fact_paragraph)                     # -> azt is töröljük
        flag_before(marker_paragraph, note)
        return False
    delete(marker_paragraph)                    # nincs válasz: a semleges, általános
    flag_before(fact_paragraph, note)           # sablonszöveg marad, jelöléssel
    keep(fact_paragraph)
    return False
```

A `fallback` ág a lényeges biztonsági különbség: ahol a sablon mondata egy
**másik cég azonosítható adatát** (nevét, regisztrációs számát, rendszerének
nevét) tartalmazza, ott üres válasz esetén sem maradhat bent - az
adatszivárgás lenne. Ezért az Szv#32 (könyvvezetésért felelős) és az Szv#44
(könyvelőprogram) mindig cserélődik, üres válasznál egy semleges,
cégfüggetlen mondatra.

Ahol a jelölő SORMINTÁBAN áll (nem külön bekezdésben), ott futás-szintű
segédek dolgoznak: `set_plain_run_text()` a nyelvtani ragasztóelemekhez
(pl. a névelő `az`→`a` cseréje az eredménykimutatás-mondatban), `fill_run()`
a „csak sárga" kiemelésű, konkrét adatot hordozó futásokhoz (fordulónap,
mérlegkészítés napja, letétbe helyezés hónapja).

### D2 – a sablonba írt, feltételes szerkesztői utasítások

A sablon NÉGY helyén a szerző magába a dokumentum szövegébe írta bele, hogy
az adott mondat csak egy bizonyos válasz esetén maradhat benn - pl.
*„ezt csak akkor írja ha az alábbi 25.ös kérdésre igen a válasz, különben meg
NE ÍRJA"*. Ez az utasítás **sosem kerülhet ki az ügyfélhez**, a mondatot
pedig a válasz alapján meg kell tartani vagy törölni:

```python
_AUTHOR_INSTRUCTION_RE = re.compile(r"\bez(?:t)?\s+(?:csak\s+)?akkor\b", re.IGNORECASE)

def keep_if(paragraph, condition, note=None):
    strip_author_instruction(paragraph)   # az utasítás MINDIG kikerül
    if condition:
        keep(paragraph)                   # a mondat marad
        return True
    if note:
        flag_before(paragraph, note)
    delete(paragraph)                     # a mondat törlődik
    return False
```

A regex a `merged.docx`-on ellenőrzötten **pontosan 3 futásra illeszkedik**,
és mind a három a szerző utasítása (a jogi szövegben ez a fordulat sehol
máshol nem szerepel). A negyedik D2-hely (ZÁRADÉK, Szv#106) nem hordoz
utasítás-szöveget, csak feltételes felsorolás-tételt - ott ugyanez a
`keep_if()` törli a bekezdést.

A négy hely és a feltételük:

| bekezdés | mit szabályoz | feltétel |
|---|---|---|
| 223. | „…155. § …könyvvizsgálatra kötelezett." | `konyvvizsgalat_tipus` kötelező ág |
| 313. | „…tételeket …nem vonjuk össze…" | `szv34` „nem"/üres |
| 484. | „…leányvállalatnak minősül. Bevonásra kerül…" | mindig cserélődik (Szv#35 ágai) |
| 87./954. | „Önköltségszámítási szabályzat" a felsorolásokban | `onkoltsegszamitasi_szabalyzat` |

A 313.-nál a szerző utasítása **kifejezetten jelölést kér** a másik ágban
(„különben csak pirossal legyen itt hogy ide kell vmit írni"), ezért ott
szándékosan NEM fogalmazunk meg saját mondatot - a `keep_if()` `note`
paraméterével csak jelölünk.

### A „innen hiányzik" helykitöltők – szöveg szerinti keresés

A IV.8. fejezetben (345-350., 355. bekezdés) a sablon 7 helyen egyszerűen
odaírta, hogy `innen hiányzik`, **mindenféle formázás nélkül** - nincs sem
áthúzott kitöltendő hely, sem `▶ Adatbékérő` jelölő, amire a formázás alapú
keresés épülhetne. Ezeket ezért SZÖVEG szerint keressük meg, és a
behelyettesített értéket utólag látjuk el a szokásos jelöléssel:

```python
def fill_text_placeholder(paragraph, placeholder, value):
    for i, r in enumerate(paragraph.runs):
        pos = r.text.find(placeholder)
        if pos >= 0:
            substitute_in_run(paragraph, i, r.text[:pos], value,
                              r.text[pos + len(placeholder):])
            return True
    return False
```

A `substitute_in_run()` a futást HÁROM futásra bontja (before / value /
after), és csak a középsőre teszi a `mark_substituted()` jelölést - így a
körülötte lévő mondat érintetlen marad. **Mindig pontosan három futás
keletkezik** (üres before/after esetén is), hogy a Python és a JS motor
futás-indexei bitre azonosak maradjanak.

Az egyik helykitöltő előtt a sablon névelője áll (*„…ellenőrzése a innen
hiányzik feladata."*), ezért ott a névelővel EGYÜTT cseréljük, és a
`hu_article()` segéd választja a helyeset: *„…ellenőrzése **a** gazdasági
vezető feladata."* / *„…ellenőrzése **az** ügyvezető feladata."*

### Ugyanaz a tény két helyen

A sablon több ténykérdést KÉTSZER is kimond: egyszer jelölővel, egyszer
jelölő nélkül (vagy egy „csak sárga" kiemelésű futásban). Ha csak a jelölt
helyet oldanánk fel, a dokumentum önmagának mondana ellent. Ezeket a
`restate()` / `fill_run()` írja újra UGYANABBÓL a válaszból:

| tény | jelölt hely | jelöletlen ismétlés(ek) |
|---|---|---|
| üzleti/cégérték (Szv#61) | 802. | 702., 799. |
| értékhelyesbítés (Szv#63) | 775. | 705., 721., 723. |
| céltartalék-képzés (Szv#96) | 867. | 758. |
| fordulónap (Szv#18) | 345. | 328. |
| mérlegkészítés napja (Szv#17) | 346. | 334., 661. |
| könyvvizsgálatra átadás (Szv#27) | 347. | 238. |
| letétbe helyezés (Szv#19) | 348. | 259. |
| közzétételért felelős (Szv#24) | 355. | 259. |
| eredménykimutatás formája (Szv#15) | 305. | 907. |
| cégnév (Szv#1) | 3. | 57. |
| önköltségszámítási szabályzat (Szv#106) | 955. | 87., 954. |

## Miért nincs bekötve a többi kérdés

A `_nincs_sablonhely: true` kérdéseket az űrlap továbbra is felteszi és a
válasz be is kerül a kódba/mentésbe, de a motor **nem helyettesíti be
automatikusan semmilyen bekezdésbe**. Az okok:

- **Beégetett törvényi szöveg, nincs biztonságos beillesztési pont.**
  Az Szv#40 (jelentős összegű hibahatár = a mérlegfőösszeg 2%-a, de legalább
  1 millió forint) és az Szv#42 (tartós tendencia = 365 naptári nap) a
  számviteli törvény szó szerinti szövege. Egy szabadszöveges választ ebbe
  behelyettesíteni a mondat átírását igényelné, amire nem vállalkozunk -
  ezért ezen a két helyen a jelölőt töröljük, és **mindig
  `⚠ ELLENŐRIZENDŐ` jelölést** hagyunk, hogy a felülvizsgáló eldönthesse,
  kell-e módosítani. Ez a két jelölés minden legenerált dokumentumban benne
  van; ez tudatos, nem hiba.
- **A választó-kérdés adja a választ, a szabadszöveges ikerpárja nem.**
  Több forráskérdéshez tartozik egy zárt választó ÉS egy szabadszöveges
  kérdés is (`szv45`↔`koltsegelszamolas_tipus`, `szv83`↔
  `szerzodes_elszamolasi_egyseg`, `szv94`↔
  `ertekpapir_kamatkulonbozet_elhatarolas`, `szv95`/`ert5`↔
  `arfolyamveszteseg_halasztott_raforditas`, `szv106`↔
  `onkoltsegszamitasi_szabalyzat`). A sablon mondatát a **választó** oldja
  fel; a szabadszöveges iker csak nyilvántartásra kerül, mert a kettő
  egymásnak ellentmondhatna.
- **Az Szv#19 (letétbe helyezés határideje) szabadszövegként nem
  illeszthető be.** A sablon két helyen is kötött szerkezetű tagmondatot
  vár („a mérlegfordulónapot követő *N.* hónap utolsó napjáig"), amibe egy
  „május 31." típusú válasz nyelvtanilag hibás mondatot adna. Mindkét helyet
  a `letetbehelyezes_hatarido_tipus` választó oldja fel, hogy a dokumentum
  két helyen se mondjon mást; az `szv19` szabad szövege nyilvántartásra
  kerül.
- **Az új sablonban egyszerűen nincs hozzájuk hely.** Az idegen nyelvű
  beszámoló (`van_idegen_nyelvu_beszamolo`, `idegen_nyelv`), a beszámoló
  mértékegysége, a kivételes mérlegkészítési időpont
  (`merlegkeszites_van_kivetel`, `merlegkeszites_kivetelek`), az euróra/USD-re
  áttérés dátuma (`penznem_atteres_datum`), a leltárkészítés sablon szerinti
  ágválasztója (`leltar_felelos_tipus`, `leltar_delegalas_reszletek` - itt az
  ügyfél szabad szövege, a `leltar_felelos_leiras` írja a mondatot), a
  pénzkezelés ágválasztója (`penzkezeles_felelos_tipus`) és a
  `jelentos_osszeg_ertelmezes` a régi sablon szerkezetéhez tartozott, az
  újban nincs megfelelő bekezdés. Ezeket **szándékosan nem kényszerítjük be**
  egy oda nem illő helyre.
- **Az adatbekérő tájékozódó kérdései.** A cégismertető (Szv#9-11), az SZMSZ
  (Szv#38), a behajthatatlan követelések (Szv#88-90, 92), a szállítói
  kötelezettségek (Szv#97-98), az önköltségszámítás módszere (Szv#107), a
  halasztott adó (Szv#109), az elfogadás rendje (Szv#22-23) és az Ért#2
  a sablonnak egyetlen bekezdésébe sem kerül be - ezek a kolléga
  háttér-információi.

Vagyis a legenerált dokumentum továbbra is **vázlat**: minden `⚠
ELLENŐRIZENDŐ` jelölést és a behelyettesített (sárga/áthúzott/piros)
szövegrészeket kiküldés előtt át kell nézni.

## Táblázatok kezelése

Az **ÚJ sablonban egyetlen táblázat** van (az üzleti jelentés célját
ismertető, tisztán leíró szövegdoboz), és **nincs benne kitöltendő hely** -
ezért a generálás egyetlen táblázat-primitívet sem hív. A régi sablon 6
táblázatának tartalma az újban prózává alakult:

| régi táblázat | hova került az új sablonban |
|---|---|
| fejléc cégadatok (Szv#1/2/5/6 + cégjegyzékszám) | önálló bekezdések (3-15., 57-71.) |
| „Bizonylat típus \| Feljogosított aláíró" (Szv#51/52/53) | egyetlen prózai mondat (617-618.) |
| értékvesztési %-os mértékek (Szv#75/76/84/87) | egyetlen szövegsor (932.) |
| feldolgozási határidők (Szv#54), zárlati ütemezés (Szv#55/56) | prózai bekezdések (585-586., 659-662.) |

A táblázat-primitívek (`table_rows()`/`row_cells()`/`cell_paragraph()`/
`fill_cell()`/`delete_table()`/`clone_row()` és JS párjaik) **szándékosan a
kódban maradtak**: a `parity_check.py` a `table_rows()`/`row_cells()` párost
használja a táblázatok run-lenyomatához (így egy véletlen táblázat-módosulás
is eltérésként jelentkezne), és egy jövőbeli, táblázatos sablonhoz újra
kellenének.

## A sablon és a bekezdés-indexek

A `sablon/` mappa tartalma:

| fájl | szerep |
|---|---|
| `szamviteli_politika_2026_uj_sablon.docx` | a 2026-os, ÚJ sablon eredeti feltöltött fájlja |
| `merged.docx` | **ezt olvassa a motor** – az új sablon összevont futású (run-merged) másolata (982 bekezdés, 1 táblázat) |
| `merged_2025_v1.docx` | a KORÁBBI sablon run-merged másolata (2130 bekezdés) – hivatkozás, a motor nem olvassa |
| `szamviteli_politika_JELOLT*.docx`, `merged_2026_megjelolve.docx` | a korábbi JELÖLT változatok – kizárólag hivatkozás |

**A sablont mindig run-merge-elni kell**, mielőtt a motor alá kerül:

```bash
python3 /mnt/skills/public/docx/scripts/merge_runs.py <új_sablon>.docx -o sablon/merged.docx
```

A Word ugyanis a helyesírás-ellenőrzési határok mentén sok apró futásra
vágja szét ugyanazt a formázott szövegrészt, ami megbízhatatlanná tenné a
kitöltendő helyek azonosítását. (A jelenlegi `merged.docx` ezzel a
szkripttel készült; a `word/document.xml` bájtra reprodukálható belőle.)

### Miért volt teljes újraírás a sablonváltás

A 2026-os sablon **nem a korábbi enyhe módosítása**: 982 bekezdés a korábbi
2130 helyett, kb. 26% szöveg-egyezéssel. Ezért az összes `resolve_*`
függvény a nulláról íródott újra, mindkét motorban. A szerkezeti
különbségek:

| | régi sablon | új sablon |
|---|---|---|
| bekezdés | 2130 | 982 |
| VAGY-alternatíva | sok tucat | **0** |
| kitöltendő táblázat | 6 (ebből 3 kitöltve) | **0** (1 leíró táblázat) |
| kitöltendő hely (blank-group) | 54 | **82** (mind `▶ Adatbékérő` jelölő) |
| formázatlan helykitöltő | 0 | **7** („innen hiányzik") |
| szerzői utasítás a szövegben | 0 | **3** + 1 feltételes felsorolás-tétel |

Ami **nem** változott: mind a 82 jelölő egyetlen futásból áll, és pontosan
ugyanazt a formázást (sárga kiemelés + áthúzás + piros `FF0000` betűszín)
viseli, mint a régi sablon „kitöltendő hely" jelölése - ezért a
`blank_groups()`/`fill_blanks()`/`mark_substituted()` primitívek
változtatás nélkül működnek rajtuk. Egy bekezdésben mindig legfeljebb egy
blank-group van.

**A bekezdés-indexek (pl. `p[655]`) a `sablon/merged.docx` jelenlegi
szerkezetéhez vannak rögzítve, MINDKÉT motorban.** Ha a sablon szövege
megváltozik, az indexek érvényüket vesztik, és a mappinget újra kell
generálni - a `parity_check.py` ilyenkor is „RENDBEN"-t adna (a két motor
egyformán téved), ezért ilyenkor a `content_check.py` és a kimenet kézi
átolvasása az igazi védelem.

## Tesztek

```bash
python3 scripts/dev/parity_check.py    # Python <-> JS motor-paritás (kötelező!)
python3 scripts/dev/content_check.py   # a kimenet tartalmi helyessége
```

Mindkettő tetszőleges további válaszfájlt is elfogad argumentumként:

```bash
python3 scripts/dev/parity_check.py data/answers.example.json sajat_teszt.json
python3 scripts/dev/content_check.py sajat_teszt.json
```

A `parity_check.py` a `generate_policy.py` kimenetét és a böngészőben
(Playwright/Chromium) futtatott `zip.js + docx.js + resolve.js` kimenetét
hasonlítja össze **run-szinten**: minden bekezdés minden futásának szövege,
kiemelése, áthúzása és betűszíne meg kell egyezzen. Ez a rendszer alapvető
helyességi garanciája – **minden `resolve_*` módosítás után le kell
futtatni**, és a két motort mindig lépésben kell tartani.

A `content_check.py` a közös kimenet HELYESSÉGÉT nézi. A minta-válaszkészletek
ágankénti ellenőrzésén túl **minden** kapott válaszfájlra lefuttatja a
`TILTOTT` listát: a sablon jelölői (`▶ Adatbékérő`), a formázatlan
helykitöltők (`innen hiányzik`), a szerzői utasítások (`ezt csak akkor
írja`, `NE ÍRJA`, …), az olvasónak szóló felszólítások (`A Vállalkozásnak
nyilatkoznia kell`, …) és a ZalaElektro-vázlatból örökölt konkrét adatok
(`Zala-Elektro`, `Business Central`, `Tábor Ildikó`, `174234`) közül
**egyik sem szerepelhet a kimenetben, egyetlen válaszkészletnél sem**. Ez
bizonyítja, hogy a D1 mondat-újraírás tényleg megtörtént, és nem maradtak
véletlen konstansként a dokumentumban.

A két minta mellett a `data/adversarial/` mappában **17 adverzariális
válaszkészlet** van, amit mindkét tesztre érdemes ráengedni:

```bash
python3 scripts/dev/parity_check.py data/answers.example*.json data/adversarial/*.json
python3 scripts/dev/content_check.py data/adversarial/*.json
```

Ezek a következőket gyakorolják: üres `{}` (`adv_empty`), csak-cégnév
(`adv_minimal`), minden select üres sztring (`adv_blank_selects` - az
iroda-oldal kitöltetlen mezője így érkezik), a négy D2-feltétel mindkét ága
(`adv_d2_true`/`adv_d2_false`), a D1-mondatok ellenágai (`adv_d1_alt`),
IFRS + euró + szabályozott piac (`adv_ifrs_euro`), mikrogazdálkodói + szabadon
gépelt deviza (`adv_mikro_egyeb`), szökőnapos / évfordulón átnyúló / tárgyéven
belüli / hibás hónapú / hibás napú / nem egész napszámú dátumok
(`adv_date_*`), JSON szám- és logikai típusok sztring helyett (`adv_types`),
magánhangzóval kezdődő és ponttal záruló szabad szövegek, mértékegységet is
hozó értékhatár (`adv_grammar`), valamint XML-ben különleges karakterek
`&`, `<`, `"`, `„”` (`adv_specialchars`).

A jelenlegi állapot mind a **19 válaszkészleten eltérésmentes paritást** ad, a
`content_check.py` mind a 19-en RENDBEN, és mind a 38 kimenet (19 Python + 19
JS) átmegy az XSD-ellenőrzésen.

**XSD- és OPC-ellenőrzés** a legenerált fájlokon:

```bash
python3 /mnt/skills/public/docx/scripts/office/validate.py output/valami.docx
```

## Kérdőív a kérdésbank alapján

A `data/questions.json` fájl `forras` mezője technikailag megkülönbözteti
az **ügyfélnek** kiküldendő és az **irodán belül** eldöntendő kérdéseket,
de jelenleg minden kérdés `forras: "ugyfel"` - a tényleges munkafolyamatban
nincs külön, csak-irodai lépés, minden választ az ügyféltől kérnek meg. Ha
a jövőben mégis szükség lenne rá (pl. egy kérdés jogi/szakmai minősítést
igényelne, amit egy átlagos ügyfél nem tud pontosan megválaszolni), a mező
`"iroda"`-ra állítható egy adott kérdésnél - ekkor az a kérdés eltűnik a
`dist_ugyfel.html`-ből, és csak a jelszóval védett `dist_iroda.html`-en
marad kitölthető.

A kérdőív technikai megvalósítására (Google Form / Typeform / saját űrlap)
ez a projekt nem tesz javaslatot - a `questions.json` bármelyikhez
könnyen átalakítható, a lényeg, hogy a beérkező válaszokból végül egy, a
`data/answers.example.json`-hoz hasonló szerkezetű JSON készüljön, amit a
`generate_policy.py` be tud olvasni.

## Munkafolyamat-javaslat

1. A kérdőívet (a `questions.json` alapján összeállítva) kiküldik az
   ügyfélnek - minden kérdést az ügyfél válaszol meg.
2. A válaszokból összeáll egy `answers.json`.
3. Lefut a `generate_policy.py` - elkészül a vázlat.
4. A vázlatot **egy kolléga átnézi**, különös figyelmet fordítva a
   `⚠ ELLENŐRIZENDŐ` jelölésekre és a még sárga/kék (döntetlen) részekre. Az
   iroda-oldal a generálás gomb után nem csak letölti a vázlatot, hanem meg
   is számolja és kiírja, hány `⚠ ELLENŐRIZENDŐ` jelölés került bele - így
   már a letöltés pillanatában látszik, mennyi kézi munka vár, nem csak
   Wordben lapozgatva derül ki.
5. Csak jóváhagyás után megy ki az ügyfélnek. **Az ügyfél a nyers, gép által
   generált vázlatot soha nem látja.**

## Fejlesztői megjegyzések

- A `scripts/generate_policy.py` a `sablon/merged.docx`-ot olvassa - ez a
  docx skill `merge_runs.py` szkriptjével összevont futású változat, mert a
  Word az eredeti fájlban helyesírás-ellenőrzési határok mentén sok apró
  futásra vágja szét ugyanazt a formázott szövegrészt, ami megbízhatatlanná
  tenné a kitöltendő helyek azonosítását.
- A bekezdés-indexek (pl. `p[655]`) a `sablon/merged.docx` jelenlegi
  szerkezetéhez vannak rögzítve. **Ha a sablon szövege megváltozik, az
  indexek érvényüket vesztik**, és a mappinget újra kell generálni.
- **A sormintában (inline) álló jelölőknél az értéknek magának kell hoznia
  az elválasztó szóközöket.** A sablon a jelölő UTÁN gyakran külön, áthúzott
  szóköz-futást tesz, amit a `fill_blanks()` a csoport többi tagjaként
  kiürít; néhol pedig a jelölő ELŐTT sincs szóköz (`a vállalkozás▶ …
  készít.`). Erre való az `inline_value()` / `inlineValue()` segéd - üres
  értéknél nem hagy lógó szóközt.
- **Szabad szövegből épített mondat végére mindig `as_sentence()` kell.**
  Az ügyfél ritkán ír mondatvégi pontot ("havonta", "átlagárfolyamot
  használunk"), a sablonba viszont MONDATKÉNT kerül be a válasza - e nélkül
  két egymás után fűzött válasz írásjel nélkül folyna egybe. Több válaszból
  álló bekezdéshez a `join_sentences()` a helyes segéd (a `join_clauses()`
  csak felsorolás-tagokat fűz, írásjel-pótlás nélkül).
- **Magyar névelő és alany-állítmány egyeztetés.** A `hu_article()` a
  válasz kezdőbetűje alapján ad `a`/`az` névelőt. Ahol a válasz lehet egy
  személy ÉS több személy is (pl. a pénzkezelés felelőse), ott szándékosan
  kettőspontos szerkezetet írunk (`A pénzkezelés felelőse: X`), hogy ne
  kelljen a `végzi`/`végzik` egyeztetést szabad szövegből eltalálni.
- **A `substitute_in_run()` MINDIG három futást hoz létre** (üres
  before/after esetén is) - e nélkül a Python és a JS futás-indexei
  elcsúsznának egymástól, és a paritás-teszt eltérést jelezne.
- **`strict_int()` / `strictInt()` mindenütt.** A Python `int(31.9)` némán
  31-et ad, a JS `strictInt()` viszont `null`-t; ezért a dátumszámítás
  (`illustrative_date()`) mindkét oldalon a szigorú egész-értelmezést és a
  kézi tartomány-ellenőrzést használja. Az `illustrative_date()` HÁRMAS
  értéket ad vissza (hónap, nap, **év-eltolódás**) - az év-eltolódás dönti
  el, hogy a mérlegkészítés napja "a tárgyév" vagy "a tárgyévet követő év"
  előtaggal íródik-e ki.
- Minden `resolve_*` függvény UGYANAZT a bekezdés-listát (`p`) kapja meg,
  amit a `generate()` egyszer, a legelején olvas ki a dokumentumból. Ez
  szándékos: törlés után a `doc.paragraphs` újraolvasása elcsúsztatná a
  többi indexet. A táblázatokkal dolgozó `resolve_*` függvények ugyanígy
  megkapják az egyszer kiolvasott táblázatlistát (`t`) is.
- **Ág-választáshoz mindig az `answer(a, kulcs, alapertelmezes)` segédet
  használd**, ne a nyers `a.get(...)` / `a.kulcs || ...` alakot. Az
  iroda-oldalon a kitöltetlen select ÜRES SZTRINGKÉNT érkezik, nem hiányzó
  kulcsként, és a JS `||` erre az alapértelmezést adná, a Python `.get()`
  viszont az üres sztringet - ez korábban a két motor eltérő kimenetét
  okozta. Az `answer()` a hiányzó és az üres mezőt egyformán kezeli, és
  mindkét motorban bitre azonos szemantikájú.
- **A behelyettesítendő értéket mindig a `text_value()` / `textValue()`
  segéden át kell szöveggé alakítani**, ne nyers `str()` / `String()`
  hívással. A kettő nem ekvivalens: `str(20.0)` → `"20.0"`, de
  `String(20.0)` → `"20"` (a JS-ben a `20.0` és a `20` ugyanaz az érték,
  megkülönböztetni nem lehet őket), és `str(True)` → `"True"`, míg
  `String(true)` → `"true"`. Az űrlapokról mindig sztring érkezik, ez a
  segéd a kézzel írt vagy programmal generált `answers.json`-ok szám- és
  logikai értékeit teszi biztonságossá.
- Ugyanez a csapda a dátum-segédben: a Python `datetime.date` hibás
  hónapra/napra kivételt dob, a JS `Date` viszont csendben túlcsordul
  (2024. 13. hó -> 2025. január). Az `illustrativeDate()` ezért kézzel
  ellenőrzi ugyanazokat a feltételeket - ne egyszerűsítsd vissza.
- **Egy tény a sablonban többször is előfordulhat** - és nem csak jelölővel!
  Az új sablon több helyen JELÖLŐ NÉLKÜL is kimondja ugyanazt (pl. "A
  vállalkozás üzleti vagy cégértékkel nem rendelkezik." a VIII. fejezetben,
  miközben a XII.-ben jelölővel is szerepel). Ezeket a `restate()` írja újra
  ugyanabból a válaszból - ld. az "Ugyanaz a tény két helyen" táblázatot. Új
  `resolve_*` írásakor **grepelj rá a sablonra** ("nem képez", "nem
  rendelkez", "nem alkalmaz", "Vállalkozásunk", "Társaságunk"), hogy nincs-e
  jelöletlen ismétlődés.
- **A sablon szerzőjének utasításai és a jelölők SOSEM mehetnek ki az
  ügyfélhez.** Ezt három réteg védi: a `keep_if()` mindig kiveszi az
  utasítást, a `fill_or_flag()`/`flag_unwired()` a kitöltetlen jelölőt is
  törli, és a generálás végén a `sweep_leftover_markers()` biztonsági háló
  fut a VÉGLEGES bekezdéslistán. A `content_check.py` `TILTOTT` listája
  minden válaszkészletnél ellenőrzi az eredményt.
