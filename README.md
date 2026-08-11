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
1. **Kitölti** a kérdőívben megválaszolt ~45 helyet a sablonban, valamint a
   fejléc cégadat-, a bizonylat-aláíró és az értékvesztési mértékek
   **táblázatának** celláit (ld. „Táblázatok kezelése”).
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
| kérdés összesen a `questions.json`-ban | **155** |
| ebből a motor automatikusan behelyettesíti / ág-választásra használja | **77** |
| ebből jelenleg nincs kitöltendő helye a sablonban (`_nincs_sablonhely`) | **78** |
| lefedett *forrás*-kérdés (Szv#/Ért# szám szerint) | **46** |
| `resolve_*` függvény (Python és JS oldalon egyaránt) | **30** |

A bekötött forráskérdések: Szv#1, 2, 5, 6, 12, 13, 14, 15, 16, 17, 18, 19,
21, 25/28, 29, 31, 32, 33, 35, 36, 37, 39, 41, 43, 44, 45, 49, 50, 51, 52,
53, 57, 59, 61, 63, 75, 76, 83, 84, 87, 94, 95, 96, 101, 106 és Ért#4.
(A Szv#95-höz felvett választó-kérdés egyben az Ért#5 által is jelölt
döntési pontot oldja fel – ugyanaz a bekezdés-csoport.)

Ebből Szv#1, 2, 5, 6, 51, 52, 53, 75, 76, 84 és 87 **táblázatba** kerül –
ld. a „Táblázatok kezelése” fejezetet lentebb.

Néhány kérdésnél az ügyfél csak leíró jelleggel válaszol (szabad szöveg),
és egy KÜLÖN, iroda-oldali választó-kérdés dönti el a sablon pontos
VAGY-ágát - ez a mintázat már korábban is megvolt (pl. könyvvezetésért/
könyvvizsgálatért felelős típusa), most a leltárkészítés felelősénél is
ezt követi: `leltar_felelos_leiras` (ügyfél, szabad szöveg, csak
nyilvántartásra) + `leltar_felelos_tipus` (iroda, választó a 3 sablon-ág
közül). Hasonlóan, a Szv#45 (költségelszámolás módja) ügyfél-oldali
kérdése szabad szöveges leírás marad, a tényleges ág-választást a
`koltsegelszamolas_tipus` iroda-mező végzi.

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
- **A válasz olyan TÁBLÁZATBA tartozna, amit nem biztonságos gépiesíteni**
  (Szv#54 feldolgozási határidők, Szv#55/56 zárlati ütemezés). Ezek a
  táblázatok már ki vannak töltve helyes, általános szakmai
  alapértelmezésekkel – ld. a „Táblázatok kezelése” fejezetet. (A többi
  táblázatos kérdés – Szv#51/52/53 és Szv#75/76/84/87 – 2026-tól **be van
  kötve**.)
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

## Táblázatok kezelése

A motor sokáig **kizárólag a törzs közvetlen bekezdéseivel** dolgozott: a
Python oldalon a `doc.paragraphs`, a JS oldalon a `getBodyParagraphs()` is
átlépi a `w:tbl` elemeket, így a táblázatok tartalma egyik motornak sem
volt látható. A `merged.docx`-ban **6 táblázat** van; az alábbiak szerint
kezeljük őket.

### Amit a motor kitölt

| # | táblázat | forrás-kérdés | mit csinál |
|---|---|---|---|
| 0 | fejléc cégadatok (CÉGNÉV / SZÉKHELY / ADÓSZÁM / CÉGJEGYZÉKSZÁM / képviselő) | Szv#1, 2, 5, 6 + `cegjegyzekszam` | mind az 5 értékcellát kitölti |
| 2 | „Bizonylat típus \| Feljogosított aláíró” | Szv#51, 52, 53 | a VAGY-ág szerint vagy kitölti (2 sorra bővítve), vagy **az egész táblázatot törli** |
| 5 | értékvesztési %-os mértékek kategóriánként | Szv#75, 76, 84, 87 | csak a 4 lefedett kategóriasor „Százalék” celláját írja felül, és csak ha van válasz |

A **0. táblázat** első sorának mindkét cellája három bekezdésből áll (üres /
szöveg / üres), ezért ott a középső bekezdésbe írunk, hogy az érték a
„CÉGNÉV” felirattal egy sorba kerüljön. A `cegjegyzekszam` az egyetlen
**újonnan felvett** kérdés: a sablon fejlécében van rá sor, az adatbekérő
viszont nem kérdezi külön, ezért az adószám (Szv#6) mellé került.

A **2. táblázat** a VII.3.6. („A bizonylatok hitelessége”) VAGY-csoport
második ágához tartozik. A README korábban éppen ezt hozta fel példaként
arra, miért nem lehet bekötni: a „képviselő aláírása” ág választásakor a
másik bekezdés törlése egy **árván maradt, üres táblázatot** hagyott volna a
dokumentumban. A `delete_table()` / `deleteTable()` primitívvel ez már
megoldható, ezért az Szv#51 helyére egy iroda-oldali választó-kérdés
(`bizonylat_hitelesites_tipus`) került, az Szv#52/53 pedig `csak_ha`-val
erre épül. A sablon egyetlen üres adatsorát `clone_row()` /
`cloneRow()` duplikálja, mert a kérdőívnek két, egymástól elkülönülő
válasza van rá (kimenő számla, illetve egyéb kimenő számviteli bizonylat) –
a bizonylattípus-feliratok magának az adatbekérőnek a szóhasználatát
követik, nem találunk ki új kategóriákat.

Az **5. táblázat** minden kategóriasora tartalmaz egy 20%-os
sablon-alapértéket. Csak azt a négy sort írjuk felül, amelyre az
adatbekérőnek célzott kérdése van (I. Immateriális javak, III. Befektetett
pénzügyi eszközök, IV. Készletek, V. Követelések), és **csak akkor, ha az
ügyfél tényleg megadott értéket** – üres válasznál marad a 20%, mert az egy
érvényes alapértelmezés, nem egy kitöltetlen hely. A II. Tárgyi eszközök,
VI. Értékpapírok, VII. Pénzeszközök és VIII. Részesedések sorokhoz nincs
kérdés, azokat sosem módosítjuk, ahogy a végig üres „Érték (E Ft)” oszlopot
sem.

### Amihez szándékosan NEM nyúlunk

| # | táblázat | miért marad érintetlen |
|---|---|---|
| 1 | „Hatáskör \| Felelős” (II.1., egyetlen üres adatsor) | A 2026-os jelölés szerint az Szv#31-hez tartozik, de az Szv#31 két meglévő kérdése (`ertekeles_felelos_tipus` / `ertekelesi_delegalas`) a II.3. fejezet *értékelési* feladataira vonatkozik, és oda is van bekötve (338. bekezdés). A táblázat két oszlopa (hatáskör + felelős) **párokat** várna, amire az adatbekérőben nincs kérdés; a „Hatáskör” oszlop szövegét nekünk kellene kitalálni. Ez fabrikálás lenne, ezért a sort a felülvizsgáló kolléga tölti ki. |
| 3 | „Bizonylat típus \| Analitikában \| Főkönyvben” (Szv#54) | Mind a 4 adatsor **helyes, általános szakmai alapértelmezéssel** érkezik (pl. „késedelem nélkül”, „legkésőbb a negyedévet követő hó végéig”). Az Szv#54 EGYETLEN szabadszöveges válasz, amiből 8 cellát kitölteni nem lehet; a meglévő, jogilag korrekt szöveg felülírása vagy törlése kárt okozna. |
| 4 | „Számviteli teendő \| Határidő” (Szv#55/56, 15 adatsor) | Ugyanaz az eset, még hangsúlyosabban: mind a 15 sor kitöltött, érvényes zárlati ütemezés. Az Szv#55 („történik-e havi/negyedéves zárás?”) és az Szv#56 („mi a határidő?”) két szabadszöveges válasz – ebből 15 sort gépiesen levezetni nem lehet. A sablon maga is „VÁLLAKOZÁSSPECIFIKUSRA KELL ALAKÍTANI” megjegyzést tesz elé: ez tudatosan a felülvizsgáló kolléga feladata. |

Ez ugyanaz az elv, mint a nem táblázatos, szándékosan bekötetlenül hagyott
kérdéseknél: **inkább maradjon `_nincs_sablonhely` egy dokumentált
indokkal, mint hogy a motor egy jogilag helyes alapértelmezést elrontson.**

### A táblázat-primitívek

A `blank_groups()` / `fill_blanks()` **nem használható** táblázatban: a
sablon egyetlen táblázatcellájában sincs áthúzott „kitöltendő hely” futás,
amire a kitöltendő helyek keresése épül. Ezért külön primitívek készültek,
mindkét motorban szigorúan párban:

| Python (`generate_policy.py`) | JS (`docx.js`) | szerep |
|---|---|---|
| `table_rows()` / `row_cells()` | `getRows()` / `getCells()` | a **nyers** `w:tr` / `w:tc` gyerekek |
| `cell_paragraph()` | `cellParagraph()` | cella adott sorszámú `w:p` gyereke |
| `fill_cell()` | `fillCell()` | cella kitöltése + a szokásos sárga/áthúzott/piros jelölés |
| `delete_table()` | `deleteTable()` | teljes táblázat törlése |
| `clone_row()` | `cloneRow()` | sor duplikálása közvetlenül az eredeti után |

**Fontos: a cellákat szándékosan a nyers `w:tc` gyerekekből indexeljük**, és
nem a python-docx `table.rows[i].cells[j]` rácsából. A python-docx a
vízszintesen összevont (`gridSpan`) cellákat a rácsban többször is
felsorolja – az 5. táblázat fejlécsora fizikailag 3 db `w:tc`, a rácsban
viszont 4 cella –, a JS oldal viszont csak a nyers XML-gyerekeket látja. Ha
a két motor máshogy indexelne, ugyanaz a `resolve_*` MÁS cellába írna.

A `generate()` a bekezdéslistához hasonlóan a **táblázatlistát is egyszer**
olvassa ki (`t = doc.tables`, illetve `const t = getTables(xmlDoc)`), mert a
`resolve_bizonylat_alairas()` egy egész táblázatot törölhet, ami utána
elcsúsztatná a többi táblázat indexét.

A `parity_check.py` run-lenyomata **kiterjed a táblázatokra is**
(táblázat/sor/cella/bekezdés indexekkel együtt) – e nélkül a
táblázat-kitöltés bármilyen Python↔JS eltérése némán átcsúszna a
paritás-ellenőrzésen, hiszen a `doc.paragraphs` a táblázatokat nem
tartalmazza. Így egy törölt táblázat, egy duplikált sor vagy egy elcsúszott
cella is azonnal eltérésként jelentkezik.

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
- Egy döntés a sablonban többször is előfordulhat: a "szerződés elszámolási
  egysége" VAGY-csoport pl. NÉGY helyen szerepel (1374, 1479, 1676, 1885) -
  mindet ugyanabból a válaszból kell feloldani, különben a dokumentum
  önmagának mond ellent. Új `resolve_*` írásakor érdemes rágrepelni a
  sablonra, hogy nincs-e ismétlődés.
