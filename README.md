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
sablon/szamviteli_politika_JELOLT.docx   <- az eredeti, feltöltött sablon
sablon/merged.docx           <- a sablon "megtisztított" (összevont futású) másolata - ezt olvassa a script
scripts/generate_policy.py   <- a fő motor (Python): válaszok + sablon -> kitöltött Word-vázlat
scripts/webapp.py            <- helyi Python webform a motor fölé (egyoldalas, jelszó nélküli)
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

## FONTOS – amit ez a verzió NEM fed le

A `data/questions.json` **a teljes, hivatalos adatbekérőt tartalmazza** (a
Számviteli politika + Értékelési szabályzat adatbekérő 115 kérdését,
pontosan a forrás xlsx sorrendjében) - ebből azonban **csak 25 kérdésnek
van kitöltendő helye ebben a sablonban**. A többi 90 kérdés (pl. cégadatok,
könyvelési gyakorlat részletei, eszközértékelési szabályok stb.) az
űrlapon megjelenik és a válasz összegyűjtésre kerül (az ügyfél kódjában,
majd az iroda oldalon), de a `generate_policy.py` / `resolve.js` motor
**nem helyettesíti be automatikusan semmilyen bekezdésbe** - ezeket a
`questions.json`-ban az adott kérdés `"_nincs_sablonhely": true` jelzi, és
az iroda-oldalon minden ilyen mező alatt egy magyarázó megjegyzés is
látható. (A hiányzó 90 kérdés jó része a mellékletekhez - leltár-,
pénzkezelési szabályzat - tartozna, azokból viszont nincs feltöltött
sablon; másik része olyan VAGY-alternatívát fed le a sablonban, amit még
nem kötöttünk össze a kérdőívvel.)

Emellett maga a sablon **jóval több döntési pontot (VAGY-alternatívát)
tartalmaz**, mint amennyit a 25 bekötött kérdés lefed - különösen a VIII.
fejezet (eszközök és források értékelése), ami önmagában több száz
bekezdésnyi választási lehetőséget sorol fel (pl. eszközcsoportonkénti
értékelési módszerek, écs-kulcsok, minden egyes elszámolási VAGY-döntés).
**Ezeket a script szándékosan nem nyúlja hozzá** - pontosan úgy maradnak a
kimeneti fájlban (sárga kiemeléssel, VAGY-gal), mint az eredeti JELÖLT
sablonban, mert nincs hozzájuk kötött kérdés.

Vagyis a legenerált dokumentum egy **részleges vázlat**: a "törzsadat"
jellegű részeket (pénznem, fordulónap, aláíró, könyvvizsgálat,
könyvelő/könyvvizsgáló felelős, leltár felelőse, konszolidáció,
jelentőségi küszöbök stb.) kitölti, de a döntés nagy részét (elsősorban a
VIII. fejezet eszközértékelési politikáját) továbbra is a kollégának kell
kézzel elvégeznie - pontosan úgy, ahogy most is teszi. **Ez így is valós
időmegtakarítás**, mert a leggyakrabban változó, ügyfélspecifikus
"törzsadatokat" nem kell manuálisan átírni és a helyükön keresgélni - és
mert a kérdőív a teljes adatbekérőt lefedi, az iroda egyetlen helyen látja
az összes választ, akkor is, ha egy részüket még kézzel kell felhasználnia.

### Ha a teljes dokumentumot szeretnék automatizálni

Ehhez bővíteni kellene a kérdőívet úgy, hogy minden fennmaradó VAGY-döntést
is lefedjen (jellemzően eszközértékelési politikai döntések), és a
`generate_policy.py`-ban minden ilyen döntéshez fel kellene venni egy új
`resolve_*` függvényt - ugyanazzal a mintával, mint a jelenlegi 17 függvény.
Ez egy jelentősen nagyobb, több körös munka (a VIII. fejezet önmagában
valószínűleg 60-80 további döntési pontot jelent).

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
