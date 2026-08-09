# Számviteli politika – automatikus vázlatgenerátor

Ez a projekt a "JELÖLT" számviteli politika sablon és az ügyfél-kérdőív
alapján **automatikusan legenerál egy Word-vázlatot**, amit a kollégák
átnéznek/pontosítanak, mielőtt kiküldik az ügyfélnek.

## Két külön oldal - az ügyfél és az iroda soha nem lát ugyanazt

A `webapp-client/` mappa egy teljesen önálló, kliens-oldali (böngészőben
futó) rendszert épít, KÉT külön HTML-fájlból - nincs szükség Python-ra vagy
szerverre a végfelhasználók gépén:

```bash
python3 webapp-client/build.py
# -> webapp-client/dist_ugyfel.html
# -> webapp-client/dist_iroda.html
```

**`dist_ugyfel.html` - ezt küldjétek ki az ügyfélnek.** Csak az ügyféltől
kérdezendő mezőket tartalmazza (a szakmai minősítést igénylő "iroda" kérdések
nincsenek benne), és **nem generál semmilyen dokumentumot** - a kitöltés
végén egy kódot ad, amit az ügyfél e-mailben visszaküld. Az ügyfél ebből az
oldalból semmit nem tud kikövetkeztetni az automatizálásról vagy a végleges
dokumentumról.

**`dist_iroda.html` - ez a belső, jelszóval védett eszköz.** Itt illeszthető
be az ügyféltől kapott kód (ez előtölti az ügyfél válaszait), itt tölthetők
ki a fennmaradó, szakmai minősítést igénylő mezők, és csak itt lehet
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
data/answers.example2.json   <- második minta, az ELLENTÉTES VAGY-ágakat gyakorolja
sablon/szamviteli_politika_JELOLT.docx   <- az eredeti, feltöltött sablon
sablon/merged.docx           <- a sablon "megtisztított" (összevont futású) másolata - ezt olvassa a script
sablon/*_2026_megjelolve.docx <- a 2026-os, előre megjelölt változat (csak hivatkozási térkép, ld. lentebb)
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

### A kódos átadás működése

Az ügyfél-oldal a válaszokat egy base64-be kódolt JSON-ná ("kód") alakítja,
amit az ügyfél e-mailben küld vissza (a "Küldés e-mailben" gomb egy előre
kitöltött e-mailt nyit meg). Az iroda-oldal ezt a kódot dekódolja és
előtölti vele a saját, azonos mezőit. Ha az irodának saját, fix e-mail
címe van, ahova ezt kéritek, írd be az `ugyfel_source.html`
`OFFICE_EMAIL` konstansába, majd futtasd újra a `build.py`-t - így a
"Küldés e-mailben" gomb már eleve ki lesz címezve.

Parancssorból (JSON válaszfájlból), ha nem a webes űrlapot használod:

```bash
python3 scripts/generate_policy.py data/answers.example.json output/valami.docx
```

A script:
1. **Kitölti** a kérdőívben megválaszolt ~45 helyet a sablonban.
2. A dokumentumban sok helyen két vagy több, "VAGY" szóval elválasztott
   **alternatív bekezdés** van (pl. "kötelező a könyvvizsgálat" / "nem
   kötelező" / "önkéntes"). A válaszok alapján kiválasztja a megfelelőt, és
   **törli a többit** (a "VAGY" elválasztókkal együtt).
3. Eltávolítja a `[Szv N. kérdés]` / `[Ért N. kérdés]` hivatkozásokat és a
   JELÖLT munkapéldány szerkesztői útmutatóját (az elejéről).
4. A **behelyettesített válaszszöveget** a sablon eredeti "kitöltendő hely"
   jelölésével (sárga kiemelés, áthúzás, piros betűszín) hagyja - ez jelzi a
   felülvizsgáló kollégának, mely szövegrészek kerültek be automatikusan, és
   melyeket kell átnéznie kiküldés előtt. A bekezdés egyéb részein (a
   VAGY-alternatívák közül kiválasztott, változatlanul hagyott szövegen)
   megszünteti a kék "még döntendő" betűszínt, hogy csak a tényleges
   behelyettesítés maradjon kiemelve.
5. Ahol a válasz nem egyértelmű (pl. "nincs cégérték" vagy "nem aktiválunk",
   miközben a sablon egy leírási időt kérne), **nem tippel** - egy jól
   látható, félkövér piros "⚠ ELLENŐRIZENDŐ" jelölést szúr be a bekezdés elé,
   és a szövegrészt eredeti, kitöltetlen állapotban hagyja.

## Lefedettség – mit tölt ki és mit NEM

A `data/questions.json` **a teljes, hivatalos adatbekérőt tartalmazza** (a
Számviteli politika + Értékelési szabályzat kérdéseit, pontosan a forrás
xlsx sorrendjében), kiegészítve néhány, csak az irodán belül eldönthető
választó-kérdéssel.

| | darab |
|---|---|
| kérdés összesen a `questions.json`-ban | **151** |
| ebből a motor automatikusan behelyettesíti / ág-választásra használja | **64** |
| ebből jelenleg nincs kitöltendő helye a sablonban (`_nincs_sablonhely`) | **87** |
| lefedett *forrás*-kérdés (Szv#/Ért# szám szerint) | **35** |
| `resolve_*` függvény (Python és JS oldalon egyaránt) | **27** |

A bekötött forráskérdések: Szv#12, 13, 14, 15, 16, 17, 18, 19, 21, 25/28,
29, 31, 32, 33, 35, 36, 37, 39, 41, 43, 44, 45, 49, 50, 57, 59, 61, 63,
83, 94, 95, 96, 101, 106 és Ért#4. (A Szv#95-höz felvett választó-kérdés
egyben az Ért#5 által is jelölt döntési pontot oldja fel – ugyanaz a
bekezdés-csoport.)

A fennmaradó ~87 kérdés (cégadatok, könyvelési gyakorlat részletei,
eszközértékelési szabályok stb.) az űrlapon megjelenik és a válasz
összegyűjtésre kerül (az ügyfél kódjában, majd az iroda oldalon), de a
motor **nem helyettesíti be automatikusan semmilyen bekezdésbe** – ezeket a
`questions.json`-ban az adott kérdés `"_nincs_sablonhely": true` jelzi, és
az iroda-oldalon minden ilyen mező alatt egy magyarázó megjegyzés is
látható.

### Miért nincs bekötve a többi jelölt kérdés

A 2026-os, előre megjelölt sablonváltozat (`sablon/*_2026_megjelolve.docx`,
ld. lentebb) 60 helyet jelöl meg. Ezek egy részéhez **szándékosan nem
készült automatizmus**, mert a sablon szerkezete nem teszi biztonságossá:

- **Nincs sem kitöltendő hely, sem VAGY-alternatíva** a jelölt bekezdésnél,
  csak beégetett törvényi szöveg (pl. Szv#40 jelentős hibahatár = a
  mérlegfőösszeg 2%-a, Szv#42 tartós tendencia = 365 nap, Szv#70 a
  200 eFt-os értékhatár, Szv#108 TAO/KIVA, Szv#110 GloBE). Ide egy válasz
  behelyettesítéséhez a mondat szövegét kellene átírni, amire a motor
  bekezdés- és futás-szintű primitívjei nem alkalmasak.
- **A válasz egy TÁBLÁZATBA tartozna** (Szv#51/52/53 aláírásra jogosultak,
  Szv#54 feldolgozási határidők, Szv#55/56 zárlati ütemezés, Szv#75/76/84/87
  értékvesztési %-os mértékek). A motor kizárólag a törzs bekezdéseivel
  dolgozik, táblázatot nem módosít és nem is töröl – így pl. a Szv#51/52
  VAGY-ágának törlése egy árván maradt táblázatot hagyna a dokumentumban.
- **A jelölés egy kötelező törvényi felsorolás egy sorára mutat**
  (Szv#66/67/68 környezetvédelmi közzétételek a kiegészítő melléklet
  tartalomjegyzékében). Egy "nem" válasz alapján statutórius szövegsort
  törölni szakmai döntés; ráadásul a jelölések itt bizonyíthatóan
  elcsúsztak (a Szv#66 jelölése nem a tárgyi eszközös, hanem a
  céltartalékos sor elé került).
- **A VAGY-ág szerkezete kétértelmű**: a Szv#60 (kísérleti fejlesztés
  aktiválása) csoportban egy oda nem illő, "felesleges" VAGY-elválasztó
  van a felvezető mondat és az első alternatíva között, így nem
  egyértelmű, hogy a köztes 5 magyarázó bekezdés melyik ághoz tartozik.
- **A "nem" ág csak mondatvég, nem önálló mondat** (Szv#19 szabályozott
  piaci letétbe helyezési határidő, Szv#95 "nem számoljuk el halasztott
  ráfordításként"). Ilyenkor a bekezdés törlése nyelvtanilag hibás mondatot
  hagyna, ezért a motor ezekben az esetekben **nem töröl, hanem
  `⚠ ELLENŐRIZENDŐ` jelölést szúr be** a bekezdés elé.

Emellett maga a sablon **jóval több döntési pontot (VAGY-alternatívát)
tartalmaz**, mint amennyit a bekötött kérdések lefednek – különösen a VIII.
fejezet (eszközök és források értékelése). **Ezeket a script szándékosan
nem nyúlja hozzá** – pontosan úgy maradnak a kimeneti fájlban (sárga
kiemeléssel, VAGY-gal), mint az eredeti JELÖLT sablonban.

Vagyis a legenerált dokumentum továbbra is egy **részleges vázlat**, de
2026-tól már a beszámoló formája, a mérleg és az eredménykimutatás választott
típusa, a teljes kiegészítő melléklet-blokk (éves vs. egyszerűsített), a
költségelszámolás módja, a pénzkezelés felelőse és több értékelési
VAGY-döntés is automatikusan feloldódik.

## A 2026-os megjelölt sablon és a bekezdés-indexek

A `sablon/` mappában három forrásfájl van:

| fájl | szerep |
|---|---|
| `szamviteli_politika_JELOLT.docx` | az eredeti, feltöltött JELÖLT sablon |
| `merged.docx` | **ezt olvassa a motor** – a JELÖLT sablon összevont futású (run-merged) másolata |
| `szamviteli_politika_JELOLT_2026_megjelolve.docx` + `merged_2026_megjelolve.docx` | a 2026-os, előre megjelölt változat (60 db `▶ Adatbékérő: …` jelölés) – **kizárólag hivatkozási térkép** |

**A 2026-os fájl NEM lett a generálás alapja, és ez szándékos.** A két
dokumentum jogi szövege lényegében azonos (`difflib` arány 0,96; az eltérés
gyakorlatilag csak a `[Szv N. kérdés]` hivatkozások és a JELÖLT fejléc
eltávolítása, plusz a 60 beszúrt jelölő bekezdés), **a 2026-os változatból
azonban eltűnt a kitöltendő helyek formázása**: az eredeti sablonban minden
kitöltendő hely sárga kiemelésű, áthúzott, piros futás (`……………`), a
2026-osban viszont ezek sima kék szöveggé olvadtak össze a környezetükkel.
Számokban: a régi fájlban 54 kitöltendő hely (blank-group) van 47
bekezdésben, az újban 60 – de az a 60 **maga a 60 jelölő bekezdés**, a
tényleges kitöltendő helyek száma **nulla**.

Mivel a `blank_groups()` (és rá épülve az egész `fill_blanks()`) pontosan
ezt az áthúzott formázást használja a kitöltendő helyek megtalálására, a
sablon lecserélése azt jelentette volna, hogy **mind a 45 meglévő
behelyettesítés némán, hibaüzenet nélkül elmarad**. Ezért a `merged.docx`
maradt a generálás alapja, a 2026-os fájl pedig térképként szolgált: a
jelölések pozícióit `difflib.SequenceMatcher`-rel képeztük vissza a
`merged.docx` bekezdés-indexeire (2130-ból 2123 bekezdés egyértelműen
megfeleltethető). Ennek járulékos haszna, hogy **a meglévő 24 kérdés egyetlen
bekezdés-indexe sem mozdult el**, tehát nem kellett kockázatos tömeges
átindexelést végezni.

## Tesztek

```bash
python3 scripts/dev/parity_check.py    # Python <-> JS motor-paritás (kötelező!)
python3 scripts/dev/content_check.py   # a kimenet tartalmi helyessége
```

A `parity_check.py` a `generate_policy.py` kimenetét és a böngészőben
(Playwright/Chromium) futtatott `zip.js + docx.js + resolve.js` kimenetét
hasonlítja össze **run-szinten**: minden bekezdés minden futásának szövege,
kiemelése, áthúzása és betűszíne meg kell egyezzen. Ez a rendszer alapvető
helyességi garanciája – **minden `resolve_*` módosítás után le kell
futtatni**, és a két motort mindig lépésben kell tartani.

## Kérdőív a kérdésbank alapján

A `data/questions.json` fájl `forras` mezője jelzi, hogy az adott kérdést
**az ügyfélnek** kell-e kiküldeni, vagy **az irodán belül** (könyvelő) kell
eldönteni - néhány kérdés (pl. "kötelező-e a könyvvizsgálat", "fenntartha-
tósági jelentésre kötelesek-e") jogi/szakmai minősítés, amit egy átlagos
ügyfél nem feltétlen tud pontosan megválaszolni, ezért ezeket célszerű a
kollégának előre kitöltenie, mielőtt lefut a generálás.

A kérdőív technikai megvalósítására (Google Form / Typeform / saját űrlap)
ez a projekt nem tesz javaslatot - a `questions.json` bármelyikhez
könnyen átalakítható, a lényeg, hogy a beérkező válaszokból végül egy, a
`data/answers.example.json`-hoz hasonló szerkezetű JSON készüljön, amit a
`generate_policy.py` be tud olvasni.

## Munkafolyamat-javaslat

1. A kérdőívet (a `questions.json` alapján összeállítva) kiküldik az
   ügyfélnek, illetve a belső ("iroda") kérdéseket egy kolléga tölti ki.
2. A válaszokból összeáll egy `answers.json`.
3. Lefut a `generate_policy.py` - elkészül a vázlat.
4. A vázlatot **egy kolléga átnézi**, különös figyelmet fordítva a
   `⚠ ELLENŐRIZENDŐ` jelölésekre és a még sárga/kék (döntetlen) részekre.
5. Csak jóváhagyás után megy ki az ügyfélnek. **Az ügyfél a nyers, gép által
   generált vázlatot soha nem látja.**

## Fejlesztői megjegyzések

- A `scripts/generate_policy.py` a `sablon/merged.docx`-ot olvassa (nem a
  JELÖLT eredetit) - ez a docx skill `merge_runs.py` szkriptjével
  összevont futású változat, mert a Word az eredeti fájlban helyesírás-
  ellenőrzési határok mentén sok apró futásra vágja szét ugyanazt a
  formázott szövegrészt, ami megbízhatatlanná tenné a kitöltendő helyek
  azonosítását.
- A bekezdés-indexek (pl. `p[655]`) a `sablon/merged.docx` jelenlegi
  szerkezetéhez vannak rögzítve. **Ha a sablon szövege megváltozik, az
  indexek érvényüket vesztik**, és a mappinget újra kell generálni.
- Minden `resolve_*` függvény UGYANAZT a bekezdés-listát (`p`) kapja meg,
  amit a `generate()` egyszer, a legelején olvas ki a dokumentumból. Ez
  szándékos: törlés után a `doc.paragraphs` újraolvasása elcsúsztatná a
  többi indexet.
- **Ág-választáshoz mindig az `answer(a, kulcs, alapertelmezes)` segédet
  használd**, ne a nyers `a.get(...)` / `a.kulcs || ...` alakot. Az
  iroda-oldalon a kitöltetlen select ÜRES SZTRINGKÉNT érkezik, nem hiányzó
  kulcsként, és a JS `||` erre az alapértelmezést adná, a Python `.get()`
  viszont az üres sztringet - ez korábban a két motor eltérő kimenetét
  okozta. Az `answer()` a hiányzó és az üres mezőt egyformán kezeli, és
  mindkét motorban bitre azonos szemantikájú.
- Ugyanez a csapda a dátum-segédben: a Python `datetime.date` hibás
  hónapra/napra kivételt dob, a JS `Date` viszont csendben túlcsordul
  (2024. 13. hó -> 2025. január). Az `illustrativeDate()` ezért kézzel
  ellenőrzi ugyanazokat a feltételeket - ne egyszerűsítsd vissza.
- Egy döntés a sablonban többször is előfordulhat: a "szerződés elszámolási
  egysége" VAGY-csoport pl. NÉGY helyen szerepel (1374, 1479, 1676, 1885) -
  mindet ugyanabból a válaszból kell feloldani, különben a dokumentum
  önmagának mond ellent. Új `resolve_*` írásakor érdemes rágrepelni a
  sablonra, hogy nincs-e ismétlődés.
