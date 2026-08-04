# Számviteli politika – automatikus vázlatgenerátor

Ez a projekt a "JELÖLT" számviteli politika sablon és az ügyfél-kérdőív
alapján **automatikusan legenerál egy Word-vázlatot**, amit a kollégák
átnéznek/pontosítanak, mielőtt kiküldik az ügyfélnek.

## Kipróbálás - két féle "rendszer" közül lehet választani

**A) Böngészőben futó, publikált oldal (nincs telepítés, csak egy link)**
A `webapp-client/` mappa egy teljesen önálló, kliens-oldali (böngészőben
futó) változat - nincs szükség Python-ra a felhasználó gépén, mert a teljes
motor (ZIP-írás/olvasás, docx-szerkesztés) JavaScriptben újraírva fut. Ebből
építhető egy publikálható, egyetlen HTML-fájlból álló oldal:

```bash
python3 webapp-client/build.py   # -> webapp-client/dist.html
```

A `dist.html` egy statikus, önálló fájl (a sablon és a kérdőív be van égetve
base64-ként) - egy erre alkalmas felületen (pl. Claude Artifact) közzétéve
egy sima linkként megnyitható, kattintható eszközzé válik: kitöltés ->
"Vázlat legenerálása" -> letöltés, mind a böngészőben, szerver nélkül.

**B) Helyi Python-os webform (ha van, aki tud terminált használni)**

```bash
python3 scripts/webapp.py
```

Ezután a böngészőben: **http://localhost:8000** - ugyanaz a kérdőív, csak
Python-szerveren keresztül futtatva (`scripts/generate_policy.py`).
(A `python3 scripts/webapp.py 8123` formával más portot is megadhatsz, ha
a 8000-es foglalt.)

**A/B eltérés:** mindkettő A/UGYANAZT a motorlogikát futtatja (a Python és a
JS változat tesztekkel ellenőrzötten pontosan azonos kimenetet ad ugyanazon
válaszokra) - csak a futtatás módja más. Ha a sablon szövege változik, MINDKÉT
oldali (Python + JS) `resolve_*` logikát frissíteni kell, mert a bekezdés-
indexek mindkettőben hard-code-olva vannak.

## Hogyan működik belül

```
data/questions.json          <- a kérdőív definíciója (mit kérdezünk az ügyféltől / mit tud a könyvelő)
data/answers.example.json    <- egy kitöltött minta-válaszkészlet (teszteléshez)
sablon/szamviteli_politika_JELOLT.docx   <- az eredeti, feltöltött sablon
sablon/merged.docx           <- a sablon "megtisztított" (összevont futású) másolata - ezt olvassa a script
scripts/generate_policy.py   <- a fő motor (Python): válaszok + sablon -> kitöltött Word-vázlat
scripts/webapp.py            <- helyi Python webform a motor fölé
webapp-client/zip.js         <- minimál ZIP olvasó/író böngészőben (CompressionStream alapú)
webapp-client/docx.js        <- docx bekezdés/futás-szintű segédfüggvények böngészőben
webapp-client/resolve.js     <- a generate_policy.py resolve_* függvényeinek JS portja
webapp-client/page_source.html <- az oldal HTML/CSS/JS sablonja (build.py egészíti ki adattal)
webapp-client/build.py       <- összeállítja a publikálható dist.html-t
output/                      <- ide kerülnek a legenerált vázlatok
```

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
4. A már **kész** (kitöltött vagy eldöntött) bekezdéseknél megszünteti a
   sárga kiemelést / áthúzást / kék "még döntendő" betűszínt, hogy a
   kollégák első pillantásra lássák: ez már megvan.
5. Ahol a válasz nem egyértelmű (pl. "nincs cégérték" vagy "nem aktiválunk",
   miközben a sablon egy leírási időt kérne), **nem tippel** - egy jól
   látható, félkövér piros "⚠ ELLENŐRIZENDŐ" jelölést szúr be a bekezdés elé,
   és a szövegrészt eredeti, kitöltetlen állapotban hagyja.

## FONTOS – amit ez a verzió NEM fed le

A feltöltött kérdőív (`kerdesek_szerint_v2.xlsx`) 136 kérdést tartalmaz, de
ebből **csak 25-nek van kitöltendő helye ebben a sablonban** (a többi a
mellékletekhez - leltár-, pénzkezelési, értékelési szabályzat - tartozna,
azokból viszont nincs feltöltött sablon).

Emellett maga a sablon **jóval több döntési pontot (VAGY-alternatívát)
tartalmaz**, mint amennyit ez a 25 kérdés lefed - különösen a VIII. fejezet
(eszközök és források értékelése), ami önmagában több száz bekezdésnyi
választási lehetőséget sorol fel (pl. eszközcsoportonkénti értékelési
módszerek, écs-kulcsok, minden egyes elszámolási VAGY-döntés). **Ezeket a
script szándékosan nem nyúlja hozzá** - pontosan úgy maradnak a kimeneti
fájlban (sárga kiemeléssel, VAGY-gal), mint az eredeti JELÖLT sablonban,
mert nincs hozzájuk ügyféladat.

Vagyis a legenerált dokumentum egy **részleges vázlat**: a "törzsadat"
jellegű részeket (pénznem, fordulónap, aláíró, könyvvizsgálat,
könyvelő/könyvvizsgáló felelős, leltár felelőse, konszolidáció,
jelentőségi küszöbök stb.) kitölti, de a döntés nagy részét (elsősorban a
VIII. fejezet eszközértékelési politikáját) továbbra is a kollégának kell
kézzel elvégeznie - pontosan úgy, ahogy most is teszi. **Ez így is valós
időmegtakarítás**, mert a leggyakrabban változó, ügyfélspecifikus
"törzsadatokat" nem kell manuálisan átírni és a helyükön keresgélni.

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
