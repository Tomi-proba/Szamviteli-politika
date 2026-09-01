// 1:1 port of generate_policy.py's resolve_* functions.
//
// A bekezdés-indexek a sablon/merged.docx JELENLEGI szerkezetéhez vannak
// rögzítve, PONTOSAN ugyanazok, mint a Python oldalon. Ha a sablon szövege
// változik, MINDKÉT motort frissíteni kell - a paritás-teszt
// (scripts/dev/parity_check.py) minden eltérést azonnal kimutat.

// --- konstans mondatkészletek (a generate_policy.py azonos nevű
// szótárainak PONTOS párjai) ---

// A beszámoló formájának TÁRGYESETŰ alakja - a sablon mondata
// ("Mindezek alapján a vállalkozás▶ készít.") tárgyesetet kíván.
const BESZAMOLO_FORMA_TARGYESET = {
  eves_beszamolo: 'éves beszámolót',
  egyszerusitett_eves_beszamolo: 'egyszerűsített éves beszámolót',
  mikrogazdalkodoi_egyszerusitett_eves_beszamolo:
    'mikrogazdálkodói egyszerűsített éves beszámolót',
  eves_beszamolo_ifrs: 'éves beszámolót az IFRS-ek szerint',
};

const EREDMENY_ELJARAS = {
  osszkoltseg_eljarassal: 'összköltség eljárással',
  forgalmi_koltseg_eljarassal: 'forgalmi költség eljárással',
};

const MERLEG_VALTOZAT = { a_valtozat: '„A” változata', b_valtozat: '„B” változata' };

const KONYVVIZSGALAT_MONDAT = {
  kotelezo_altalanos:
    'A számviteli beszámoló szabályszerűségének, megbízhatóságának és ' +
    'valódiságának felülvizsgálatára és tanúsítására - a gazdálkodás ' +
    'volumenére és a kapcsolódó jogszabályi előírásokra tekintettel - a ' +
    'legfőbb szerv könyvvizsgálót választott.',
  kotelezo_koztartozas_miatt:
    'A számviteli beszámoló szabályszerűségének, megbízhatóságának és ' +
    'valódiságának felülvizsgálatára és tanúsítására a legfőbb szerv ' +
    'könyvvizsgálót választott. A választás oka az, hogy az árbevételhez ' +
    'és létszámhoz kapcsolódó nagyságrendi mentesítések nem ' +
    'alkalmazhatók, mert a Vállalkozásnak a fordulónapon - ' +
    'összegszerűségében és lejáratában jogszabályban megszabott mértéket ' +
    'meghaladó jelentős - köztartozása van.',
  onkentes:
    'A számviteli beszámoló szabályszerűségének, megbízhatóságának és ' +
    'valódiságának felülvizsgálatára és tanúsítására a legfőbb szerv - ' +
    'önkéntesen, függetlenül attól, hogy arra jogszabály nem ír elő ' +
    'kötelezettséget - könyvvizsgálót választott.',
  nincs_kotelezettseg:
    'A számviteli beszámoló szabályszerűségének, megbízhatóságának és ' +
    'valódiságának felülvizsgálatára és tanúsítására a Vállalkozásnak ' +
    'jogszabályi kötelezettsége nem áll fenn, ezért a legfőbb szerv ' +
    'könyvvizsgálót nem választott.',
};

const KONSZOLIDACIO_MONDAT = {
  nem_erintett:
    'Jelenleg a Vállalkozás nem minősül a Számviteli törvény 3. § (2) ' +
    'bek. 1-4. pontja szerinti definíció alapján sem anyavállalatnak, sem ' +
    'közös vezetésű vállalkozásnak, sem leányvállalatnak, sem társult ' +
    'vállalkozásnak, ezért nem tagja olyan vállalkozáscsoportnak, ' +
    'amelyről összevont (konszolidált) éves beszámolót kellene ' +
    'készíteni. Konszolidálásba történő bevonás esetén a számviteli ' +
    'politika e fejezetének felülvizsgálata és módosítása szükséges.',
  leany_mentesul_bevonas_alol:
    'A Vállalkozás a Számviteli törvény szerint leányvállalatnak ' +
    'minősül, de a Számviteli törvény 119. §-a alapján mentesül az ' +
    'anyavállalat által készítendő összevont (konszolidált) éves ' +
    'beszámolóba való bevonás alól, így az összevont (konszolidált) éves ' +
    'beszámoló elkészítése a Vállalkozást nem érinti, csak az ' +
    'anyavállalat felé történő adatszolgáltatást kell teljesíteni.',
  leany_bevonva:
    'A Vállalkozás a Számviteli törvény szerint leányvállalatnak ' +
    'minősül, amely bevonásra kerül az anyavállalat által készítendő ' +
    'összevont (konszolidált) éves beszámolóba. A konszolidációs ' +
    'adatszolgáltatás teljesítése a Vállalkozás számviteli beszámolóját ' +
    'nem érinti.',
  anya_mentesul_folerendelt_anya_miatt:
    'A Vállalkozás a Számviteli törvény szerint anyavállalatnak minősül, ' +
    'de a számviteli törvény 116. §-a alapján mentesül a konszolidált ' +
    'éves beszámoló elkészítése alól, mert fölérendelt anyavállalata ' +
    'összevont (konszolidált) éves beszámolót készít, amelybe a ' +
    'Vállalkozást is bevonja.',
  anya_mentesul_nagysagrend_alapjan:
    'A Vállalkozás a Számviteli törvény szerint anyavállalatnak minősül, ' +
    'de a számviteli törvény 117. §-a alapján mentesül az összevont ' +
    '(konszolidált) éves beszámoló készítése alól. A Vállalkozás a ' +
    'jövőben folyamatosan figyeli, hogy a mentesítés feltételei ' +
    'továbbra is fennállnak-e.',
  anya_keszit_konszolidaltat:
    'A Vállalkozás a Számviteli törvény szerint anyavállalatnak minősül ' +
    'és összevont (konszolidált) éves beszámolót készít.',
};

const KOLTSEGELSZAMOLAS_MONDAT = {
  csak_5:
    'Vállalkozásunk a könyvvezetésében a költségeket csak az 5. ' +
    'számlaosztályban, költségnemenként számolja el.',
  '5_es_67':
    'Vállalkozásunk a könyvvezetésében a költségeket elsődlegesen az 5. ' +
    'számlaosztályban költségnemenként, másodlagosan a 6-7. ' +
    'számlaosztályokban költséghely-költségviselőnként számolja el.',
};

// --- a generate_policy.py levezetett-érték segédeinek párjai ---

function merlegkeszitesSzoveg(a) {
  const [honap, nap, ev] = illustrativeDate(
    answer(a, 'fordulonap_honap', ''),
    answer(a, 'fordulonap_nap', ''),
    answer(a, 'merlegkeszites_nap', 0));
  const nev = monthName(honap);
  if (!nev) return '';
  const elotag = (ev && ev >= 1) ? 'a tárgyévet követő év ' : 'a tárgyév ';
  return `${elotag}${nev} hónap ${nap}.`;
}

function fordulonapSzoveg(a) {
  const nev = monthName(answer(a, 'fordulonap_honap', ''));
  const nap = strictInt(answer(a, 'fordulonap_nap', ''));
  if (!nev || nap === null) return '';
  return `${nev} hónap ${nap}. napja`;
}

function letetbehelyezesSzoveg(a) {
  return answer(a, 'letetbehelyezes_hatarido_tipus', 'altalanos') === 'szabalyozott_piac'
    ? 'negyedik' : 'ötödik';
}

// --- I. fejezet - címlap és cégadatok ---

function resolveCegadatok(p, a, xmlDoc) {
  // A cégnév (Szv#1) KÉT helyen szerepel (3. és 57.), ugyanabból a válaszból.
  for (const [idx, key, cimke] of [
    [3, 'szv1', 'a vállalkozás neve'],
    [6, 'szv2', 'a székhely címe'],
    [9, 'cegjegyzekszam', 'a cégjegyzékszám'],
    [12, 'szv6', 'az adószám'],
    [15, 'szv5', 'a képviseletre jogosult személy'],
    [57, 'szv1', 'a cég neve'],
    [59, 'gazdalkodasi_forma', 'a működési/gazdálkodási forma'],
    [61, 'alapitas_idopontja', 'az alapítás időpontja'],
    [63, 'szv3', 'a telephely'],
    [65, 'fioktelepek', 'a fióktelep(ek)'],
    [67, 'szv4', 'a tulajdonos(ok)'],
    [69, 'szv7', 'a statisztikai számjel'],
    [71, 'szv8', 'a főtevékenység'],
  ]) {
    fillOrFlag(p[idx], txt(a, key), `hiányzik ${cimke} a kérdőívből.`, xmlDoc);
  }
}

function resolveEcsErtekhatar(p, a, xmlDoc) {
  // Szv#70 - a jelölő SORMINTÁBAN áll, közvetlenül utána a sablon "forint"
  // szava; ha a válasz maga is hoz mértékegységet, azt ki kell venni.
  const ertek = txt(a, 'szv70');
  if (!ertek) {
    fillBlanks(p[101], [inlineValue('…')], xmlDoc);
    flagBefore(p[101],
      'nincs válasz a kis értékű tárgyi eszközök egyösszegű ' +
      'értékcsökkenési leírásának értékhatárára (Szv#70) - a ' +
      'felsorolás pontját kérjük kézzel kiegészíteni.', xmlDoc);
    return;
  }
  const norm = deaccent(ertek.toLowerCase());
  if (/(ft|forint)\.?$/.test(norm)) dropLeadingWord(p[101], 3, 'forint', xmlDoc);
  fillBlanks(p[101], [inlineValue(ertek)], xmlDoc);
}

// --- II. fejezet - a számviteli feladatok felelősei ---

function resolveSzamviteliFelelosok(p, a, xmlDoc) {
  fillOrFlag(p[188], txt(a, 'szv30'),
    'nincs válasz arra, kinek a feladata a számviteli szabályzatok ' +
    'elkészítése, módosítása (Szv#30).', xmlDoc);

  if (answer(a, 'ertekeles_felelos_tipus') === 'egyedi_delegalas') {
    fillOrFlag(p[191], txt(a, 'ertekelesi_delegalas'),
      'egyedi delegálás lett megjelölve a számvitelpolitikai ' +
      'döntésekért való felelősségre (Szv#31), de a delegálás ' +
      'részletei nincsenek megadva.', xmlDoc);
  } else {
    fillBlanks(p[191],
      ['a vezetés, illetve a vezetés felhatalmazása alapján a ' +
       'számviteli rendért felelős személy.'], xmlDoc);
    if (!txt(a, 'ertekeles_felelos_tipus')) {
      flagBefore(p[191],
        'nincs válasz arra, hogy a számvitelpolitikai döntésekért a ' +
        'vezetés/számviteli rendért felelős személy felel-e, vagy ' +
        'egyedi delegálás van (Szv#31) - a sablon alapértelmezett ' +
        'ága lett feltételezve, kérjük ellenőrizni.', xmlDoc);
    }
  }
}

function resolveLeltar(p, a, xmlDoc) {
  // Szv#36 - a sablon konkrét beosztásokat nevez meg, ezért a felvezető
  // tagmondatot általánosra írjuk át. A bekezdés a jelölő UTÁN ponttal
  // folytatódik, ezért a beírt érték nem végződhet mondatzáró írásjelre.
  const felelos = txt(a, 'leltar_felelos_leiras');
  setPlainRunText(p[203], 0,
    'A leltárkészítés és leltározás felelősségének rendje: ', xmlDoc);
  if (felelos) {
    fillBlanks(p[203], [felelos.replace(/[. ]+$/, '')], xmlDoc);
  } else {
    fillBlanks(p[203], ['a vezetés hatásköre és felelőssége'], xmlDoc);
    flagBefore(p[203],
      'nincs válasz arra, ki a leltárkészítés és leltározás felelőse ' +
      "(Szv#36) - a semleges 'a vezetés hatásköre és felelőssége' " +
      'megfogalmazás lett beírva, kérjük pontosítani.', xmlDoc);
  }
}

function resolvePenzkezeles(p, a, xmlDoc) {
  // Szv#37 - kettőspontos szerkezet, hogy ne kelljen a magyar
  // alany-állítmány egyeztetést (végzi / végzik) szabadszövegből eltalálni.
  const felelos = txt(a, 'szv37');
  rewriteFact(p[208], p[209],
    felelos
      ? 'A pénzkezelés a vezetés hatásköre és felelőssége, az ezzel ' +
        'kapcsolatos feladatokat a vonatkozó jogszabályi előírások, valamint ' +
        'a Vállalkozás pénzkezelési szabályzatának előírásai szerint kell ' +
        `ellátni. A pénzkezelés felelőse: ${asSentence(felelos)}`
      : '',
    'nincs válasz arra, ki a pénzkezelés felelőse (Szv#37) - az alábbi ' +
    'mondat a sablon konkrét, MÁS cégre írt megfogalmazását tartalmazza, ' +
    'kérjük kézzel átírni.', xmlDoc);
}

// --- III. fejezet - könyvvizsgálat, közzététel ---

function resolveKonyvvizsgalat(p, a, xmlDoc) {
  // A 223. bekezdés a sablon szerzőjének feltételes utasítását hordozza
  // (ld. generate_policy.py) - a keepIf() az utasítást mindenképp kiveszi.
  const tipus = answer(a, 'konyvvizsgalat_tipus', '');
  const vanValasz = !!textValue(tipus).trim();
  const kotelezett = tipus === 'kotelezo_altalanos' || tipus === 'kotelezo_koztartozas_miatt';
  const vanKonyvvizsgalo = kotelezett || tipus === 'onkentes';

  keepIf(p[223], kotelezett,
    vanValasz ? null
      : 'nincs válasz a könyvvizsgálati kötelezettségre (Szv#25) - a ' +
        'kötelezettséget kimondó mondat biztonsági okból törölve lett, ' +
        'kérjük ellenőrizni/pontosítani.', xmlDoc);

  fillOrFlag(p[224],
    Object.prototype.hasOwnProperty.call(KONYVVIZSGALAT_MONDAT, tipus)
      ? KONYVVIZSGALAT_MONDAT[tipus] : '',
    'nincs válasz a könyvvizsgálati kötelezettségre / a könyvvizsgáló ' +
    'választására (Szv#25, Szv#26) - kérjük kézzel pontosítani.', xmlDoc);

  if (vanKonyvvizsgalo) {
    fillOrFlag(p[228], txt(a, 'konyvvizsgalo_adatai'),
      'van könyvvizsgáló, de a könyvvizsgáló (cég) adatai hiányoznak ' +
      'a kérdőívből (Szv#28).', xmlDoc);
    fillOrFlag(p[230], txt(a, 'konyvvizsgalo_nyilvantartasi_szam'),
      'van könyvvizsgáló, de a nyilvántartási száma hiányzik ' +
      '(Szv#115).', xmlDoc);
    fillOrFlag(p[233], txt(a, 'konyvvizsgalatert_felelos_konyvvizsgalo'),
      'van könyvvizsgáló, de a könyvvizsgálat végrehajtásáért ' +
      'személyében felelős könyvvizsgáló neve hiányzik (Szv#116).', xmlDoc);
  } else {
    // Nincs könyvvizsgáló: a könyvvizsgáló adatait kérő bekezdések
    // tárgytalanok, a hozzájuk tartozó feliratokkal együtt törlendők.
    for (const idx of [226, 228, 229, 230, 232, 233]) deleteP(p[idx]);
  }

  // Szv#27 - "csak sárga" kiemelésű futás, konkrét dátummal.
  const atadas = txt(a, 'szv27');
  if (atadas) {
    fillRun(p[238], 1, atadas.replace(/[. ]+$/, ''), xmlDoc);
  } else {
    flagBefore(p[238],
      'nincs válasz arra, mikor kell a beszámoló adatait ' +
      'könyvvizsgálatra átadni (Szv#27) - az alábbi mondat a sablon ' +
      'konkrét, MÁS cégre írt dátumát tartalmazza.', xmlDoc);
  }
}

function resolveKozzetetel(p, a, xmlDoc) {
  // A 259. bekezdésben KÉT ügyfélspecifikus adat van egy mondatban: a
  // közzétételről gondoskodó személy (jelölő) és a letétbe helyezés
  // határideje ("5. hónap utolsó napjáig", "csak sárga" futásban).
  const felelos = txt(a, 'szv24');
  fillBlanks(p[259], [inlineValue(felelos)], xmlDoc);
  if (!felelos) {
    flagBefore(p[259],
      'nincs válasz arra, kinek a feladata gondoskodni a beszámoló ' +
      'közzétételéről és letétbe helyezéséről (Szv#24).', xmlDoc);
  }
  fillRun(p[259], 4,
    `${letetbehelyezesSzoveg(a)} hónap utolsó napjáig kell letétbe`, xmlDoc);
}

// --- IV. fejezet - a beszámoló formája, mérleg, eredménykimutatás ---

function resolveBeszamoloForma(p, a, xmlDoc) {
  // A jelölő a "vállalkozás" szóhoz tapad (nincs köztük szóköz), utána
  // áthúzott szóköz-futás áll - az értéknek MINDKÉT oldali szóközt hoznia
  // kell, és tárgyesetben kell állnia.
  const tipus = classifyBeszamoloForma(answer(a, 'szv13'));
  const forma = Object.prototype.hasOwnProperty.call(BESZAMOLO_FORMA_TARGYESET, tipus)
    ? BESZAMOLO_FORMA_TARGYESET[tipus] : '';
  if (forma) {
    fillBlanks(p[271], [inlineValue(forma, ' ')], xmlDoc);
  } else {
    fillBlanks(p[271], [inlineValue('beszámolót', ' ')], xmlDoc);
    flagBefore(p[271],
      'a beszámoló választott formája hiányzik vagy nem egyértelmű ' +
      "(Szv#13) - a semleges 'beszámolót' szó lett beírva, kérjük a " +
      'pontos formát kézzel megadni.', xmlDoc);
  }
}

function resolveMerlegForma(p, a, xmlDoc) {
  const tipus = answer(a, 'szv14');
  const valtozat = Object.prototype.hasOwnProperty.call(MERLEG_VALTOZAT, tipus)
    ? MERLEG_VALTOZAT[tipus] : '';
  if (valtozat) {
    fillBlanks(p[298], [inlineValue(valtozat)], xmlDoc);
  } else {
    // Üres válasznál látható hiányjelet hagyunk, nem néma félmondatot.
    fillBlanks(p[298], [inlineValue('…')], xmlDoc);
    flagBefore(p[298], 'a mérleg választott formája (A/B változat) hiányzik (Szv#14).', xmlDoc);
  }
}

function resolveEredmenykimutatasForma(p, a, xmlDoc) {
  // UGYANAZ a döntés a sablon KÉT pontján (IV.3. és XVI. fejezet). A 305.
  // bekezdésben a jelölő előtti névelőt is a válaszhoz kell igazítani.
  const tipusKey = answer(a, 'szv15');
  const eljaras = Object.prototype.hasOwnProperty.call(EREDMENY_ELJARAS, tipusKey)
    ? EREDMENY_ELJARAS[tipusKey] : '';

  if (eljaras) {
    setPlainRunText(p[305], 2, huArticle(eljaras) + ' ', xmlDoc);
    fillBlanks(p[305], [inlineValue(eljaras)], xmlDoc);
  } else {
    // Üres válasznál a névelőt meghagyjuk, és látható hiányjelet írunk - így
    // a mondat nem csonkul némán értelmetlenre.
    setPlainRunText(p[305], 2, 'az ', xmlDoc);
    fillBlanks(p[305], [inlineValue('…')], xmlDoc);
    flagBefore(p[305], 'az eredménykimutatás választott formája hiányzik (Szv#15).', xmlDoc);
  }

  rewriteFact(p[906], p[907],
    eljaras
      ? 'Vállalkozásunk az üzleti év adózott eredményét a számviteli törvény ' +
        `1. számú mellékletében meghatározott, ${eljaras} készülő ` +
        'eredménykimutatás formában állapítja meg, mutatja be.'
      : '',
    'az eredménykimutatás választott formája hiányzik (Szv#15) - az ' +
    'alábbi mondat a sablon konkrét, MÁS cégre írt választását ' +
    'tartalmazza.', xmlDoc);
}

function resolveTetelekModositasa(p, a, xmlDoc) {
  // Szv#34 - a szerzői utasítás KIFEJEZETTEN jelölést (és nem új mondatot)
  // kér a másik ágban, ezért ott nem fogalmazunk meg saját szöveget.
  const valasz = txt(a, 'szv34');
  const nemModosit = !valasz || nemleges(valasz);
  keepIf(p[313], nemModosit,
    'a Társaság a mérleg/eredménykimutatás tételeit módosítja (Szv#34: ' +
    `„${valasz}”) - a tételek összevonásáról/elhagyásáról szóló ` +
    'bekezdést a tényleges gyakorlat szerint kézzel kell megfogalmazni.', xmlDoc);
  if (nemModosit && !valasz) {
    flagBefore(p[313],
      'nincs válasz arra, módosítja-e a Társaság a mérleg és az ' +
      "eredménykimutatás tételeit (Szv#34) - a 'nem módosítja' ág lett " +
      'feltételezve, kérjük ellenőrizni/pontosítani.', xmlDoc);
  }
  deleteP(p[314]);
}

function resolveBeszamoloNyelve(p, a, xmlDoc) {
  const nyelv = txt(a, 'beszamolo_nyelve');
  fillBlanks(p[320], [inlineValue(nyelv || 'magyar')], xmlDoc);
  if (!nyelv) {
    flagBefore(p[320],
      "nincs válasz a beszámoló nyelvére (Szv#16) - a 'magyar' lett " +
      'feltételezve, kérjük ellenőrizni.', xmlDoc);
  }
}

// --- IV.6-IV.8. - fordulónap, mérlegkészítés, a beszámolóért való felelősség
//
// Itt van a sablon 7 db, FORMÁZÁS NÉLKÜL hagyott "innen hiányzik"
// helykitöltője (345-350., 355.), amiket SZÖVEG szerint kell megkeresni,
// valamint azok "csak sárga" párjai (328., 334.), ahol a sablon UGYANAZT a
// tényt még egyszer, konkrét dátummal is kimondja.

function resolveBeszamoloertFelelosseg(p, a, xmlDoc) {
  const fordulonap = fordulonapSzoveg(a);
  const merlegkeszites = merlegkeszitesSzoveg(a);

  if (fordulonap) {
    fillRun(p[328], 1, `12 hónap, fordulónapja ${fordulonap}.`, xmlDoc);
  } else {
    flagBefore(p[328],
      'nincs (értelmezhető) válasz a mérleg fordulónapjára (Szv#18) - ' +
      'az alábbi mondat a sablon konkrét dátumát tartalmazza.', xmlDoc);
  }

  if (merlegkeszites) {
    fillRun(p[334], 1, merlegkeszites, xmlDoc);
  } else {
    flagBefore(p[334],
      'nincs (értelmezhető) válasz a mérlegkészítés időpontjára ' +
      '(Szv#17/Szv#18) - az alábbi mondat a sablon konkrét dátumát ' +
      'tartalmazza.', xmlDoc);
  }

  fillOrFlag(p[341], txt(a, 'szv20'),
    'nincs válasz arra, ki felelős a beszámoló elkészítéséért (Szv#20).', xmlDoc);

  const utemterv = [
    [345, fordulonap,
      'nincs (értelmezhető) válasz a mérleg fordulónapjára (Szv#18).'],
    [346, merlegkeszites ? `${merlegkeszites} napja` : '',
      'nincs (értelmezhető) válasz a mérlegkészítés időpontjára (Szv#17).'],
    [347, txt(a, 'szv27'),
      'nincs válasz arra, mikor kell a beszámoló adatait ' +
      'könyvvizsgálatra átadni (Szv#27).'],
    [348, 'a mérlegfordulónapot követő ' +
      `${letetbehelyezesSzoveg(a)} hónap utolsó napja`, ''],
    [349, txt(a, 'beszamolo_kozzetel_datum'),
      'nincs válasz a beszámoló közzétételének határidejére.'],
  ];
  for (const [idx, ertek, note] of utemterv) {
    if (ertek) {
      fillTextPlaceholder(p[idx], GAP_PLACEHOLDER, ertek, xmlDoc);
    } else {
      fillTextPlaceholder(p[idx], GAP_PLACEHOLDER, '…', xmlDoc);
      flagBefore(p[idx], note, xmlDoc);
    }
  }

  // "... a munkafolyamatok irányítása, ellenőrzése a innen hiányzik
  // feladata." - a helykitöltő ELŐTT álló "a" névelőt is a válaszhoz kell
  // igazítani, ezért a névelővel EGYÜTT cseréljük.
  const iranyito = txt(a, 'munkafolyamat_iranyitas_felelos');
  if (iranyito) {
    fillTextPlaceholder(p[350], 'a ' + GAP_PLACEHOLDER,
      `${huArticle(iranyito)} ${iranyito.replace(/\.+$/, '')}`, xmlDoc);
  } else {
    fillTextPlaceholder(p[350], 'a ' + GAP_PLACEHOLDER, 'a vezetés', xmlDoc);
    flagBefore(p[350],
      'nincs válasz arra, ki felelős a munkafolyamatok irányításáért, ' +
      "ellenőrzéséért - a semleges 'a vezetés' lett beírva.", xmlDoc);
  }

  fillOrFlag(p[353], txt(a, 'beszamolo_alairoja') || txt(a, 'szv5'),
    'nincs válasz arra, ki írja alá a beszámolót (Szv#21).', xmlDoc);

  // UGYANAZ a válasz, mint a 259. bekezdésben (Szv#24), különben a
  // dokumentum két helyen mást mondana.
  const kozzetevo = txt(a, 'szv24');
  if (kozzetevo) {
    fillTextPlaceholder(p[355], GAP_PLACEHOLDER, kozzetevo, xmlDoc);
  } else {
    fillTextPlaceholder(p[355], GAP_PLACEHOLDER, '…', xmlDoc);
    flagBefore(p[355],
      'nincs válasz arra, ki felelős a beszámoló közzétételéért ' +
      '(Szv#24).', xmlDoc);
  }

  // Szv#33 - a beszámoló elérhetősége.
  if (answer(a, 'beszamolo_honlapon') === 'igen') {
    const cim = txt(a, 'beszamolo_honlap_cim');
    rewriteFact(p[357], p[358],
      cim ? 'A beszámoló a vállalkozás székhelyén megtekinthető, továbbá ' +
            `a(z) ${cim} internetes honlapon is elérhető.` : '',
      'a beszámoló a honlapon elérhető, de a honlap címe hiányzik ' +
      '(Szv#33).', xmlDoc);
  } else {
    keepP(p[357], xmlDoc);
    deleteP(p[358]);
    if (!txt(a, 'beszamolo_honlapon')) {
      flagBefore(p[357],
        'nincs válasz arra, elérhető-e a beszámoló honlapon ' +
        "(Szv#33) - a 'csak a székhelyen tekinthető meg' ág lett " +
        'feltételezve, kérjük ellenőrizni/pontosítani.', xmlDoc);
    }
  }
}

// --- IV.9-VI. - kiegészítő melléklet, jelentős összeg, konszolidáció ---

function resolveKiegeszitoMelleklet(p, a, xmlDoc) {
  fillOrFlag(p[448], txt(a, 'szv105'),
    'nincs válasz a fedezeti ügyletekre, a fedezeti hatékonyság ' +
    'mérésére, illetve a kockázatkezelési politikára (Szv#105).', xmlDoc);
  fillOrFlag(p[463],
    txt(a, 'kiegeszito_melleklet_sajatossagok')
      ? 'A Vállalkozás a kiegészítő mellékletben az alábbi, a vagyoni, ' +
        'pénzügyi és jövedelmi helyzetet bemutató mutatókat közli: ' +
        `${asSentence(txt(a, 'kiegeszito_melleklet_sajatossagok'))}`
      : '',
    'nincs válasz arra, milyen vagyoni, pénzügyi és jövedelmi helyzetet ' +
    'bemutató mutatókat közöl a kiegészítő melléklet (Szv#101) - az ' +
    'alábbi mutatólista a sablon alapértelmezése.', xmlDoc);

  // Szv#39 - a sablon elgépelte a "meghaladó" szót, ezt is javítjuk.
  setPlainRunText(p[480], 2, ' mFt-ot meghaladó összeget.', xmlDoc);
  fillOrFlag(p[480], txt(a, 'jelentos_osszeg_bemutatas'),
    'nincs válasz arra, mit tekint a vállalkozás jelentős összegnek a ' +
    'bemutatási kötelezettség szempontjából (Szv#39).', xmlDoc);
}

function resolveKonszolidacio(p, a, xmlDoc) {
  // A 484. bekezdés szerzői utasítást hordoz ÉS a ZalaElektro-vázlat konkrét
  // helyzetét írja le - mindkettő elhagyandó. A fejezet CÍMÉT nem töröljük:
  // a "nem érintett" esetre is van korrekt megfogalmazás, így a fejezet
  // tartalommal marad, nem üresedik ki.
  const tipus = answer(a, 'konszolidacio_tipus', '');
  let mondat = Object.prototype.hasOwnProperty.call(KONSZOLIDACIO_MONDAT, tipus)
    ? KONSZOLIDACIO_MONDAT[tipus] : '';
  const reszlet = txt(a, 'konszolidacio_reszletszabalyok');
  if (mondat && tipus !== 'nem_erintett' && reszlet) {
    mondat += ' Az anyavállalat felé történő adatszolgáltatás és az azt ' +
      'alátámasztó számviteli nyilvántartások vezetése: ' +
      `${asSentence(reszlet)}`;
  }

  keepIf(p[484], false, null, xmlDoc);
  fillOrFlag(p[485], mondat,
    'nincs válasz arra, tagja-e a vállalkozás konszolidálásba bevont ' +
    'vállalkozáscsoportnak (Szv#35) - a fejezet szövegét kézzel kell ' +
    'megadni.', xmlDoc);
}

function resolveMinositesek(p, a, xmlDoc) {
  // Szv#40 - a 489. bekezdés a számviteli törvény 3. § (3) 3. pontjának SZÓ
  // SZERINTI szövege; a szabadszöveges választ ebbe nem lehet biztonságosan
  // behelyettesíteni, ezért nem tippelünk (ld. README).
  flagUnwired(p[490],
    'a jelentős összegű hiba értékhatárát (Szv#40) a fenti bekezdés a ' +
    'számviteli törvény szerinti általános szabállyal adja meg - ha a ' +
    'Vállalkozás ettől eltérő (a törvényi kereten belüli) értékhatárt ' +
    'határozott meg, azt kézzel kell átvezetni.', xmlDoc);

  const hanyad = txt(a, 'jelentos_mertek_hanyad');
  rewriteFact(p[503], p[504],
    hanyad
      ? 'Amennyiben jogszabály, vagy a számviteli politika más része ' +
        'eltérően nem rendelkezik, úgy a konkrét eszközt, forrást, mérlegen ' +
        'kívüli tételt érintő számviteli elszámolás, illetve azok értékelése ' +
        'szempontjából jelentős mértékűnek minősül a vizsgált összeg, ha a ' +
        'mérleg vagy eredménykimutatás nagybetűvel jelölt tételét legalább ' +
        `${hanyad.replace(/%+$/, '').trim()}%-kal megváltoztatja.`
      : '',
    'nincs válasz arra, milyen hányadot tekint a vállalkozás jelentős ' +
    'mértékűnek (Szv#41) - az alábbi bekezdés a sablon 2%-os ' +
    'alapértelmezését tartalmazza.', xmlDoc);

  // Szv#42 - törvényi fogalommeghatározás, szabadszövegből nem írható át.
  flagUnwired(p[519],
    'a tartós tendencia fogalmát (Szv#42) a fenti bekezdés a törvényi ' +
    'általános szabállyal (365 naptári nap) adja meg - ha a Vállalkozás ' +
    'ettől eltérően értelmezi, azt kézzel kell átvezetni.', xmlDoc);

  // Szv#43 - kivételes nagyság: NÉGY külön válasz fut össze egy mondatba.
  const osszeg = txt(a, 'kiveteles_nagysag_osszeghatar');
  const reszek = [];
  for (const [kulcs, cimke] of [
    ['kiveteles_nagysag_sajattoke_szazalek', 'az előző üzleti évi saját tőke'],
    ['kiveteles_nagysag_bevetel_szazalek', 'az előző üzleti évi összes bevétel'],
    ['kiveteles_nagysag_koltseg_szazalek',
      'az előző üzleti évi összes költség és ráfordítás'],
  ]) {
    const ertek = txt(a, kulcs).replace(/%+$/, '').trim();
    if (ertek) reszek.push(`${cimke} ${ertek}%-át`);
  }
  const felsorolas = joinClauses(
    [osszeg ? `a(z) ${osszeg} összeget` : ''].concat(reszek), ', illetve ');
  rewriteFact(p[525], p[526],
    felsorolas
      ? 'Kivételes nagyságú vagy előfordulású tételnek azokat a tételeket ' +
        `tekintjük, amelyek értéke meghaladja ${felsorolas}. Ennek ` +
        'megfelelően az elszámolt bevételeket, illetve ráfordításokat ezt ' +
        'meghaladó összeg esetén a Kiegészítő mellékletben tételesen ' +
        'bemutatjuk.'
      : '',
    'nincs válasz arra, milyen értékű bevételt, költséget, ráfordítást ' +
    'tekintenek kivételes nagyságúnak (Szv#43) - az alábbi bekezdés a ' +
    'sablon konkrét, MÁS cégre írt értékhatárait tartalmazza.', xmlDoc);
}

// --- VII. fejezet - könyvvezetés, bizonylatolás, könyvviteli zárlat ---

function resolveKonyvvezetes(p, a, xmlDoc) {
  // Szv#12 - a könyvvezetés és a beszámoló pénzneme.
  const penznem = penznemNeve(a);
  rewriteFact(p[533], p[534],
    penznem
      ? 'A létesítő okiratban a jegyzett tőke meghatározására rögzített ' +
        `pénznem ${huArticle(penznem)} ${penznem}, ez a beszámoló pénzneme. ` +
        'Ennek megfelelően a könyvek vezetése és a beszámoló elkészítése a ' +
        `beszámoló pénznemében, azaz ${penznem} pénznemben történik.`
      : '',
    'nincs (értelmezhető) válasz a beszámoló és a könyvvezetés ' +
    'pénznemére (Szv#12) - az alábbi bekezdés a forintot feltételezi.', xmlDoc);
  if (penznem && classifyPenznem(answer(a, 'penznem_tipus', 'forint')) !== 'forint') {
    flagBefore(p[534],
      'a beszámoló pénzneme nem forint - a dokumentum több, ' +
      'jelöletlen bekezdése (pl. a mértékegységről és a devizás ' +
      'tételek forintra átszámításáról szóló részek) forintot említ, ' +
      'kérjük ezeket is átnézni.', xmlDoc);
  }

  // Szv#44 - a sablon konkrétan a ZalaElektro rendszerét nevezi meg.
  const program = txt(a, 'konyvelo_program');
  rewriteFact(p[538], p[539],
    program
      ? 'Vállalkozásunk esetében a könyvvezetés ' +
        `${huArticle(program)} ${program} segítségével történik.`
      : '',
    'nincs válasz arra, milyen könyvelőprogramot használ a Társaság ' +
    '(Szv#44) - a sablon MÁS cég rendszerét megnevező mondata törölve ' +
    'lett, kérjük a ténylegesen használt programot beírni.', xmlDoc,
    'A könyvvezetéshez használt könyvelőprogram megnevezését a ' +
    'számviteli politikában rögzíteni szükséges.');

  // Szv#32 - a sablon konkrét nevet és regisztrációs számot tartalmaz.
  const nev = txt(a, 'konyvvezetesert_felelos_szemely') || txt(a, 'konyveles_felelos_neve');
  const regszam = txt(a, 'konyveles_felelos_regszam');
  const tipus = answer(a, 'konyveles_felelos_tipus', '');
  const kepesites = { merlegkepes_konyvelo: 'mérlegképes könyvelő',
    oklev_konyvvizsgalo: 'okleveles könyvvizsgáló' }[tipus] || '';
  let felelosMondat;
  if (tipus === 'mentesseg_20mft_arbevetel_alatt') {
    felelosMondat =
      'Tekintettel arra, hogy a Vállalkozás éves (éves szintre ' +
      'átszámított) nettó árbevétele két év átlagában a 20 MFt-ot nem ' +
      'haladja meg, a számviteli feladatok végzését, irányítását, a ' +
      'beszámoló összeállítását nem mérlegképes könyvelő vagy nem ' +
      'okleveles könyvvizsgáló végzi.';
    if (nev) felelosMondat += ` A könyvvezetésért felelős: ${nev}.`;
  } else if (nev) {
    felelosMondat = 'A könyvvezetésért felelős: ' + joinClauses([
      nev,
      kepesites
        ? `(${kepesites}` + (regszam ? `, regisztrációs száma: ${regszam})` : ')')
        : (regszam ? `(regisztrációs száma: ${regszam})` : ''),
    ]).replace(/\s+$/, '') + '.';
  } else {
    felelosMondat = '';
  }
  rewriteFact(p[540], p[541], felelosMondat,
    'nincs válasz arra, ki a könyvvezetésért felelős személy (Szv#32) - ' +
    'a sablon MÁS cég munkatársát (nevét és regisztrációs számát) ' +
    'megnevező mondata törölve lett, kérjük a tényleges felelőst ' +
    'beírni.', xmlDoc,
    'A könyvvezetésért felelős személy nevét és regisztrációs ' +
    'számát a számviteli politikában rögzíteni szükséges.');

  // Szv#45/#46 - a költségelszámolás módja és a program kódrendszere.
  const koltsegKey = answer(a, 'koltsegelszamolas_tipus', '');
  let koltseg = Object.prototype.hasOwnProperty.call(KOLTSEGELSZAMOLAS_MONDAT, koltsegKey)
    ? KOLTSEGELSZAMOLAS_MONDAT[koltsegKey] : '';
  const kodrendszer = txt(a, 'szv46');
  if (koltseg && kodrendszer && !nemleges(kodrendszer)) {
    koltseg += ' A könyvelőprogramban a vállalkozás az alábbiak szerinti ' +
      `megbontást alkalmazza: ${asSentence(kodrendszer)}`;
  }
  rewriteFact(p[545], p[546], koltseg,
    'nincs válasz a költségelszámolás választott módjára (Szv#45) - az ' +
    'alábbi bekezdés a sablon konkrét, MÁS cégre írt gyakorlatát írja ' +
    'le.', xmlDoc);
}

function resolveBizonylatok(p, a, xmlDoc) {
  // Szv#54 - a bizonylatok könyvekben történő rögzítésének határideje.
  rewriteFact(p[585], p[586],
    txt(a, 'szv54')
      ? 'A számviteli bizonylatok könyvekben történő rögzítésének ' +
        `határideje: ${asSentence(txt(a, 'szv54'))}` : '',
    'nincs válasz a számviteli bizonylatok rögzítésének határidejére ' +
    '(Szv#54) - az alábbi bekezdés a sablon általános, MÁS cégre írt ' +
    'határidőit tartalmazza.', xmlDoc);

  // Szv#71 - az értékcsökkenés elszámolásának gyakorisága.
  rewriteFact(p[588], p[589],
    txt(a, 'szv71')
      ? 'Az értékcsökkenés elszámolása ' +
        `${txt(a, 'szv71').replace(/\.+$/, '')} történik.` : '',
    'nincs válasz arra, milyen gyakorisággal számolják el az ' +
    'értékcsökkenést (Szv#71) - az alábbi bekezdés a sablon konkrét, MÁS ' +
    'cégre írt gyakorlatát írja le.', xmlDoc);

  // Szv#49/#50 - a bizonylatok és a könyvelés közötti összhang, illetve a
  // bizonylatok jellege.
  const osszhang = txt(a, 'bizonylat_konyveles_osszhang');
  const jelleg = txt(a, 'bizonylatkezeles_szabalyai');
  rewriteFact(p[598], p[599],
    joinSentences([
      osszhang ? 'A számviteli bizonylatok és a könyvelés közötti összhangot a ' +
        `Vállalkozás a következők szerint teremti meg: ${osszhang}` : '',
      jelleg ? `A Vállalkozás által alkalmazott bizonylatok jellege: ${jelleg}` : '',
    ]),
    'nincs válasz a bizonylatok és a könyvelés közötti összhang ' +
    'megteremtésére, illetve a bizonylatok jellegére (Szv#49, Szv#50) - ' +
    'az alábbi bekezdés a sablon konkrét, MÁS cégre írt gyakorlatát ' +
    'írja le.', xmlDoc);

  // Szv#51/#52/#53 - a bizonylatok hitelessége. A régi sablonban ez egy
  // táblázat volt, itt egyetlen prózai mondat.
  let alairas;
  const hitelesites = answer(a, 'bizonylat_hitelesites_tipus', '');
  if (hitelesites === 'feljogositott_szemelyek') {
    alairas = joinSentences([
      'A bizonylatok alaki és tartalmi hitelességét az egyes ' +
      'bizonylattípusok esetében az arra feljogosított személyek ' +
      'aláírása igazolja.',
      txt(a, 'szv52') ? `Kimenő számla aláírására jogosult: ${txt(a, 'szv52')}.` : '',
      txt(a, 'szv53') ? 'Egyéb kimenő számviteli bizonylat aláírására jogosult: ' +
        `${txt(a, 'szv53')}.` : '',
    ]);
  } else if (hitelesites === 'kepviselo_alairasa') {
    alairas =
      'A bizonylatot a vállalkozás képviseletére jogosult személy ' +
      '(ideértve a Ptk. szerinti képviseletet is) írja alá, a ' +
      'bizonylatok alaki és tartalmi hitelességét az ő aláírása ' +
      'igazolja.';
  } else {
    alairas = '';
  }
  rewriteFact(p[617], p[618], alairas,
    'nincs válasz arra, hogy a bizonylatok hitelességét a képviseletre ' +
    'jogosult aláírása igazolja-e, vagy bizonylattípusonként külön ' +
    'feljogosított aláírók vannak (Szv#51).', xmlDoc);

  // Szv#48 - a számlarend kialakításáért felelős személy.
  rewriteFact(p[625], p[626],
    txt(a, 'szv48')
      ? 'A számlatükör és a számlarend összeállításáért, annak folyamatos ' +
        'karbantartásáért, a naprakész könyvvezetés helyességéért ' +
        `${txt(a, 'szv48').replace(/\.+$/, '')} a felelős.` : '',
    'nincs válasz arra, kinek a felelőssége a számlatükör és a ' +
    'számlarend kialakítása, karbantartása (Szv#48) - az alábbi mondat a ' +
    'sablon alapértelmezését tartalmazza.', xmlDoc);

  // Szv#47 - a jelölő (628.) egy kettősponttal záruló felvezető mondat ÉS a
  // hozzá tartozó felsorolás KÖZÉ került, ahol egy önálló mondat megtörné a
  // felsorolást - ezért a választ a felvezető mondat ELÉ szúrjuk be.
  const struktura = txt(a, 'szv47');
  if (struktura) {
    prependSentence(p[627],
      `A programban a számlatükör felépítése: ${asSentence(struktura)} `, xmlDoc);
    deleteP(p[628]);
  } else {
    flagUnwired(p[628],
      'nincs válasz arra, hogyan épül fel a programban a számlatükör ' +
      'struktúrája (Szv#47).', xmlDoc);
  }
}

function resolveZarlat(p, a, xmlDoc) {
  fillOrFlag(p[658],
    txt(a, 'evkozi_zarasok_feladatai')
      ? 'Az évközi zárások során a Vállalkozás az alábbi feladatokat végzi ' +
        `el: ${asSentence(txt(a, 'evkozi_zarasok_feladatai'))}` : '',
    'nincs válasz arra, milyen feladatokat végez a Társaság az évközi ' +
    'zárások során (Szv#57).', xmlDoc);

  // Szv#55/#58 - a főkönyvi zárás gyakorisága.
  const zaras = txt(a, 'szv55');
  const vegyes = txt(a, 'szv58');
  rewriteFact(p[659], p[660],
    joinSentences([
      zaras ? `A főkönyvi zárás rendje: ${zaras}` : '',
      vegyes ? 'Évközi zárás hiányában a vegyes tételek könyvelésének és a ' +
        `főkönyv ellenőrzésének rendje: ${vegyes}` : '',
    ]),
    'nincs válasz arra, történik-e havi vagy negyedéves főkönyvi zárás ' +
    '(Szv#55) - az alábbi mondat a sablon konkrét, MÁS cégre írt ' +
    'ütemezését tartalmazza.', xmlDoc);

  // Szv#56 - az éves zárás határidejét a mérlegkészítés választott
  // időpontjából vezetjük le, hogy a dokumentum két helyen se mondjon mást.
  const merlegkeszites = merlegkeszitesSzoveg(a);
  rewriteFact(p[661], p[662],
    joinSentences([
      merlegkeszites ? `Az éves zárás határideje ${merlegkeszites} napja.` : '',
      txt(a, 'szv56') ? `Az évközi zárások határideje: ${txt(a, 'szv56')}` : '',
    ]),
    'nincs válasz a zárások határidejére (Szv#56), és a mérlegkészítés ' +
    'időpontja sem állapítható meg - az alábbi mondat a sablon konkrét, ' +
    'MÁS cégre írt dátumát tartalmazza.', xmlDoc);
}

// --- VIII-XII. fejezet - mérlegtételek, értékelés, értékcsökkenés ---

function resolveBefektetettEszkozok(p, a, xmlDoc) {
  // Szv#59 - alapítás-átszervezés aktivált értéke.
  let mondat;
  let note;
  if (answer(a, 'alapitas_atszervezes_aktivalja') === 'igen') {
    const ev = txt(a, 'alapitas_atszervezes_leiras_ev');
    mondat = isNumber(ev)
      ? 'Az alapítás, átszervezés aktivált értékét a Vállalkozás az ' +
        `immateriális javak között mutatja ki, és ${ev} év alatt írja le.`
      : '';
    note =
      'a Vállalkozás aktiválja az alapítás-átszervezés költségeit ' +
      '(Szv#59), de a leírási idő hiányzik vagy nem szám - az alábbi ' +
      'mondat ennek az ELLENKEZŐJÉT állítja, kérjük kézzel javítani.';
  } else {
    mondat =
      'Alapítás, átszervezés aktivált értékét nem mutatunk ki, az ' +
      'ilyen címen felmerült költségeket a tárgyévi költségek között ' +
      'számoljuk el.';
    note = '';
  }
  rewriteFact(p[681], p[682], mondat, note, xmlDoc);
  if (mondat && answer(a, 'alapitas_atszervezes_aktivalja') !== 'igen' &&
      !txt(a, 'alapitas_atszervezes_aktivalja')) {
    flagBefore(p[682],
      'nincs válasz arra, aktiválják-e az alapítás-átszervezés ' +
      "költségeit (Szv#59) - a 'nem aktiváljuk' ág lett feltételezve, " +
      'kérjük ellenőrizni/pontosítani.', xmlDoc);
  }

  // Szv#60 - kísérleti fejlesztés.
  rewriteFact(p[692], p[693],
    txt(a, 'szv60')
      ? `Vállalkozásunk kísérleti fejlesztés tekintetében: ${asSentence(txt(a, 'szv60'))}`
      : '',
    'nincs válasz a kísérleti fejlesztés költségeinek aktiválására ' +
    '(Szv#60) - az alábbi mondat a sablon konkrét, MÁS cégre írt ' +
    'állítását tartalmazza.', xmlDoc);

  // Szv#62 - vagyoni értékű jogok.
  rewriteFact(p[695], p[696],
    txt(a, 'szv62')
      ? 'A vállalkozásnál vagyoni értékű jogként jellemzően a következők ' +
        `fordulnak elő: ${asSentence(txt(a, 'szv62'))}` : '',
    'nincs válasz arra, milyen jogokat tartanak nyilván a vagyoni ' +
    'értékű jogok mérlegsoron (Szv#62) - az alábbi mondat a sablon ' +
    'konkrét, MÁS cégre írt állítását tartalmazza.', xmlDoc);

  // Szv#81/#82 - vásárolt és saját termelésű készletek.
  rewriteFact(p[727], p[728],
    joinSentences([
      txt(a, 'szv81') ? 'A vásárolt anyag- és árukészletek nyilvántartása, értékelése: ' +
        `${txt(a, 'szv81')}` : '',
      txt(a, 'szv82') ? `Saját termelésű készletek: ${txt(a, 'szv82')}` : '',
    ]),
    'nincs válasz a vásárolt, illetve a saját termelésű készletek ' +
    'nyilvántartására és értékelésére (Szv#81, Szv#82) - az alábbi ' +
    'bekezdés a sablon konkrét, MÁS cégre írt állítását tartalmazza.', xmlDoc);

  // Szv#86 - a vevőkövetelések csoportosítása (a megelőző bekezdés
  // általános törvényi szöveg, azt nem írjuk át).
  fillOrFlag(p[733],
    txt(a, 'szv86')
      ? 'A vevőkövetelések könyvelési csoportosítása: ' +
        `${asSentence(txt(a, 'szv86'))}` : '',
    'nincs válasz arra, hogyan csoportosítják a vevőköveteléseket a ' +
    'könyvelésben (Szv#86).', xmlDoc);

  // Szv#95 / Ért#5 - a devizás hiteltartozások árfolyamvesztesége.
  const arfolyam = answer(a, 'arfolyamveszteseg_halasztott_raforditas',
    'nem_szamoljuk_el') === 'elszamoljuk'
    ? 'A devizaszámlán meglévő devizakészlettel nem fedezett, külföldi ' +
      'pénzértékre szóló hiteltartozások, devizakötvény-kibocsátásból ' +
      'származó tartozások esetén a tárgyévben, a mérlegfordulónapi ' +
      'értékelésből keletkezett, árfolyamnyereséggel nem ellentételezett ' +
      'árfolyamveszteséget halasztott ráfordításként elszámoljuk, és a ' +
      'kapcsolódó céltartalékot megképezzük.'
    : 'A devizaszámlán meglévő devizakészlettel nem fedezett, külföldi ' +
      'pénzértékre szóló hiteltartozások, devizakötvény-kibocsátásból ' +
      'származó tartozások mérlegfordulónapi értékeléséből keletkezett ' +
      'árfolyamveszteséget nem számoljuk el halasztott ráfordításként, ' +
      'ezért céltartalék képzés sem kapcsolódik hozzá.';
  rewriteFact(p[759], p[760], arfolyam, '', xmlDoc);
  if (!txt(a, 'arfolyamveszteseg_halasztott_raforditas')) {
    flagBefore(p[760],
      'nincs válasz arra, halasztott ráfordításként számolják-e el a ' +
      'devizás hiteltartozások árfolyamveszteségét (Szv#95/Ért#5) - a ' +
      "'nem számoljuk el' ág lett feltételezve, kérjük " +
      'ellenőrizni/pontosítani.', xmlDoc);
  }
}

function resolveErtekeles(p, a, xmlDoc) {
  fillOrFlag(p[771], txt(a, 'ert1'),
    'nincs válasz arra, van-e a számviteli törvénytől eltérő értékelési ' +
    'módszerük (Ért#1).', xmlDoc);

  fillOrFlag(p[773],
    joinSentences([
      txt(a, 'szv77') ? 'A terven felüli értékcsökkenés elszámolásának gyakorisága: ' +
        `${txt(a, 'szv77')}` : '',
      txt(a, 'szv85') ? `Az értékvesztés elszámolásának gyakorisága: ${txt(a, 'szv85')}` : '',
    ]),
    'nincs válasz a terven felüli értékcsökkenés, illetve az ' +
    'értékvesztés elszámolásának gyakoriságára (Szv#77, Szv#85).', xmlDoc);

  // Szv#63/#64/#65 - értékhelyesbítés eszközcsoportonként. A sablon egyetlen
  // mondattal kijelenti, hogy a Vállalkozás nem él a lehetőséggel - ezt mind
  // a három eszközcsoport válaszából újraírjuk.
  const csoportok = [];
  let hianyzoKapu = false;
  for (const [gate, textKey, cimke] of [
    ['ertekhelyesbites_immat_van', 'ertekhelyesbites_immat_javak',
      'az immateriális javak'],
    ['ertekhelyesbites_targyi_van', 'ertekhelyesbites_targyi_eszkoz',
      'a tárgyi eszközök'],
    ['ertekhelyesbites_penzugyi_van', 'ertekhelyesbites_penzugyi_eszkoz',
      'a befektetett pénzügyi eszközök'],
  ]) {
    if (!txt(a, gate)) hianyzoKapu = true;
    if (answer(a, gate) === 'igen') {
      const reszlet = txt(a, textKey);
      csoportok.push(`${cimke} körében${reszlet ? ` (${reszlet})` : ''}`);
    }
  }
  let ertekhelyesbites;
  if (csoportok.length) {
    ertekhelyesbites =
      'Az értékhelyesbítés alkalmazására vonatkozó lehetőséggel a ' +
      `Vállalkozás ${joinClauses(csoportok, ', ')} él.`;
    const szabalyok = txt(a, 'szv64');
    const hatar = txt(a, 'szv65');
    if (szabalyok) {
      ertekhelyesbites += ' Az értékhelyesbítés alkalmazásának szabályai: ' +
        `${asSentence(szabalyok)}`;
    }
    if (hatar) {
      ertekhelyesbites += ' Az értékhelyesbítés elszámolásához jelentősnek tekintett ' +
        `értékhatár: ${asSentence(hatar)}`;
    }
  } else {
    ertekhelyesbites =
      'Az értékhelyesbítés alkalmazására vonatkozó lehetőséggel a ' +
      'Vállalkozás nem kíván élni.';
  }
  // A sablon a VIII. fejezetben eszközcsoportonként MÉG EGYSZER kimondja
  // ugyanezt (705., 721., 723.), jelölő nélkül - ezeket ugyanabból a
  // válaszból írjuk újra (a 705. bekezdés sablonbeli elgépelése is javul).
  for (const [idx, gate, cimke] of [
    [705, 'ertekhelyesbites_immat_van',
      'Vállalkozásunknál az immateriális javakkal kapcsolatos ' +
      'értékhelyesbítési lehetőséget'],
    [721, 'ertekhelyesbites_targyi_van',
      'Vállalkozásunknál a tárgyi eszközeinkkel kapcsolatos ' +
      'értékhelyesbítési lehetőséget'],
    [723, 'ertekhelyesbites_penzugyi_van',
      'A befektetett pénzügyi eszközök értékhelyesbítési lehetőségét'],
  ]) {
    const alkalmaz = answer(a, gate) === 'igen' ? 'alkalmazzuk.' : 'nem alkalmazzuk.';
    restate(p[idx], `${cimke} ${alkalmaz}`, xmlDoc);
  }

  rewriteFact(p[774], p[775], ertekhelyesbites, '', xmlDoc);
  if (hianyzoKapu && !csoportok.length) {
    flagBefore(p[775],
      'nincs válasz arra, alkalmazzák-e az értékhelyesbítést a ' +
      "befektetett eszközök valamely körére (Szv#63) - a 'nem " +
      "alkalmazzuk' ág lett feltételezve, kérjük " +
      'ellenőrizni/pontosítani.', xmlDoc);
  }

  // Szv#78/#100 - valós értéken történő értékelés.
  const piaci = txt(a, 'szv78');
  const valos = txt(a, 'szv100');
  let valosMondat;
  if ((piaci && !nemleges(piaci)) || (valos && !nemleges(valos))) {
    valosMondat = joinSentences([
      valos ? `A valós értéken történő értékelés alkalmazása: ${valos}` : '',
      piaci ? 'A befektetett eszközök piaci értéken történő értékelése: ' +
        `${piaci}` : '',
    ]);
  } else {
    valosMondat =
      'A Vállalkozás a valós értéken történő értékelést, illetve a ' +
      'befektetett eszközök piaci értéken történő értékelését további ' +
      'külön döntésig nem alkalmazza.';
  }
  rewriteFact(p[779], p[780], valosMondat, '', xmlDoc);

  // Szv#69/#72/#73/#74 - az értékcsökkenés elszámolásának módszere és a
  // tervezéshez szükséges adatok (a megelőző bekezdés törvényi szöveg).
  fillOrFlag(p[792],
    joinSentences([
      txt(a, 'szv69') ? `Az alkalmazott leírási módszer: ${txt(a, 'szv69')}` : '',
      txt(a, 'szv72') ? 'A terv szerinti értékcsökkenés megállapításához szükséges ' +
        'adatokat (hasznos élettartam, maradványérték, üzembe helyezés ' +
        `időpontja) a következők szerint határozzuk meg: ${txt(a, 'szv72')}` : '',
      txt(a, 'szv73') ? `A maradványérték meghatározásának módja: ${txt(a, 'szv73')}` : '',
      txt(a, 'szv74') ? `A maradványértéket nullának tekintjük: ${txt(a, 'szv74')}` : '',
    ]),
    'nincs válasz az értékcsökkenés leírási módszerére, a tervezéshez ' +
    'szükséges adatokra és a maradványérték meghatározására ' +
    '(Szv#69, #72, #73, #74).', xmlDoc);

  // Szv#61 - üzleti vagy cégérték. UGYANEZT a tényt a sablon a VIII.
  // fejezetben (702.) is kimondja, jelölő NÉLKÜL.
  if (answer(a, 'van_cegertek') === 'igen') {
    restate(p[702], 'A vállalkozás üzleti vagy cégértéket mutat ki.', xmlDoc);
    restate(p[799],
      'A befejezett kísérleti fejlesztés, a befejezett ' +
      'alapítás-átszervezés, illetve az üzleti vagy cégérték aktivált ' +
      'értékének leírását a fentiek szerint végezzük.', xmlDoc);
  } else if (answer(a, 'alapitas_atszervezes_aktivalja') === 'igen') {
    restate(p[799],
      'A befejezett alapítás-átszervezés aktivált értékét társaságunk ' +
      'kimutatja, annak leírását a fentiek szerint végezzük.', xmlDoc);
  }
  let cegertek;
  let cegNote;
  if (answer(a, 'van_cegertek') === 'igen') {
    const ev = txt(a, 'cegertek_leiras_ev');
    cegertek = isNumber(ev)
      ? `A Vállalkozás üzleti vagy cégértéket mutat ki, amelyet ${ev} év alatt ír le.`
      : '';
    if (cegertek && answer(a, 'van_negativ_cegertek') === 'igen') {
      const negEv = txt(a, 'negativ_cegertek_leiras_ev');
      if (isNumber(negEv)) {
        cegertek += ' A negatív üzleti vagy cégértéket a Vállalkozás ' +
          `${negEv} év alatt írja le.`;
      }
    }
    cegNote =
      'a Vállalkozásnak van üzleti/cégértéke (Szv#61), de a leírási ' +
      'idő hiányzik vagy nem szám - az alábbi mondat ennek az ' +
      'ELLENKEZŐJÉT állítja, kérjük kézzel javítani.';
  } else {
    cegertek = 'Az üzleti vagy cégértékkel a Vállalkozás nem rendelkezik.';
    cegNote = '';
  }
  rewriteFact(p[801], p[802], cegertek, cegNote, xmlDoc);

  // Szv#80 / Szv#79 - értékvesztés hitelviszonyt megtestesítő
  // értékpapíroknál, illetve részesedéseknél.
  fillOrFlag(p[816],
    txt(a, 'szv80')
      ? 'A hosszú lejáratú, hitelviszonyt megtestesítő értékpapírok ' +
        `értékvesztésének elszámolási szabályai: ${asSentence(txt(a, 'szv80'))}` : '',
    'nincs válasz a hosszú lejáratú hitelviszonyt megtestesítő ' +
    'értékpapírok értékvesztési szabályaira (Szv#80).', xmlDoc);
  fillOrFlag(p[830],
    txt(a, 'szv79')
      ? 'A gazdasági társaságban lévő befektetések értékvesztésének ' +
        `elszámolási szabályai: ${asSentence(txt(a, 'szv79'))}` : '',
    'nincs válasz a gazdasági társaságban lévő befektetések ' +
    'értékvesztési szabályaira (Szv#79).', xmlDoc);

  // Szv#87 - a sablon konkrét, MÁS cégre írt mértéket tartalmaz.
  rewriteFact(p[839], p[840],
    txt(a, 'szv87')
      ? 'A követelések minősítésének feltételrendszere, az elszámolandó ' +
        `értékvesztés mértéke: ${asSentence(txt(a, 'szv87'))}` : '',
    'nincs válasz arra, hogyan, milyen mértékek alkalmazásával ' +
    'határozzák meg a követelésekre elszámolandó értékvesztést ' +
    '(Szv#87) - az alábbi bekezdés a sablon konkrét, MÁS cégre írt ' +
    'mértékeit tartalmazza.', xmlDoc);
}

// --- XIII-XVII. fejezet - árfolyamok, céltartalék, elhatárolások, eredmény ---

function resolveArfolyamok(p, a, xmlDoc) {
  // Szv#91 - a devizás tételek forintra átszámításához használt árfolyam.
  rewriteFact(p[846], p[847],
    txt(a, 'szv91')
      ? 'A forintérték meghatározásakor a vállalkozás a következő árfolyamot ' +
        `alkalmazza: ${asSentence(txt(a, 'szv91'))}` : '',
    'nincs válasz arra, melyik bank melyik árfolyamát használják a ' +
    'devizás tételek forintértékének meghatározásához (Szv#91) - az ' +
    'alábbi mondat a sablon konkrét, MÁS cégre írt választását ' +
    'tartalmazza.', xmlDoc);

  // Ért#3 - a bekezdés MÁSODIK mondata ez a döntés.
  rewriteFact(p[849], p[850],
    txt(a, 'ert3')
      ? 'Ha a Magyar Nemzeti Bank által nem jegyzett és nem konvertibilis ' +
        'valutát, ilyen valutára szóló eszközöket és kötelezettségeket kell ' +
        'forintra átszámítani, a valuta szabadpiaci árfolyamát használja a ' +
        'vállalkozás. A számviteli törvény 60. § (5a) bekezdése szerinti ' +
        'lehetőséggel (az általános forgalmi adó alapjának megállapításához ' +
        `használt árfolyam alkalmazása) kapcsolatban: ${asSentence(txt(a, 'ert3'))}`
      : '',
    'nincs válasz arra, élnek-e a számviteli törvény 60. § (5a) ' +
    'bekezdése szerinti árfolyam-választási lehetőséggel (Ért#3) - az ' +
    'alábbi bekezdés a sablon konkrét, MÁS cégre írt választását ' +
    'tartalmazza.', xmlDoc);

  // Szv#93 - a devizás pénzeszközök csökkenésének árfolyama.
  rewriteFact(p[854], p[855],
    txt(a, 'szv93')
      ? 'A devizás pénzeszközök csökkenésekor a forintértéket a következők ' +
        `szerint határozzuk meg: ${asSentence(txt(a, 'szv93'))}` : '',
    'nincs válasz arra, átlag- vagy FIFO árfolyamot használnak a devizás ' +
    'pénzeszközök csökkenésénél (Szv#93) - az alábbi mondat a sablon ' +
    'konkrét, MÁS cégre írt választását tartalmazza.', xmlDoc);
}

function resolveCeltartalek(p, a, xmlDoc) {
  // Szv#96 - a sablon kereken kijelenti, hogy nem képez céltartalékot.
  let mondat;
  if (answer(a, 'celtartalek_kepez') === 'igen') {
    const hatar = txt(a, 'celtartalek_jelentos_hatar');
    const elteres = txt(a, 'celtartalek_elteres_szazalek');
    mondat =
      'A jövőbeni költségekre a Vállalkozás céltartalékot képez, ' +
      'amennyiben azok összege jelentős.';
    if (hatar) {
      mondat += ' A képezhető céltartalék szempontjából a költségeket ' +
        `jelentősnek minősítjük, ha azok összege eléri a ${hatar} értéket.`;
    }
    if (elteres) {
      mondat += ' Az egyes jogcímek szerint képzett céltartalék összegét az ' +
        'előző évi összegtől lényegesen eltérőnek minősítjük, ' +
        `amennyiben a változás eléri a ${elteres.replace(/%+$/, '').trim()}%-ot.`;
    }
  } else {
    mondat =
      'A jövőbeni költségekre a Vállalkozás nem képez céltartalékot ' +
      'akkor sem, ha azok összege jelentős.';
  }
  // A sablon a VIII. fejezetben (758.) jelölő nélkül is kimondja ugyanezt.
  if (answer(a, 'celtartalek_kepez') === 'igen') {
    restate(p[758],
      'A jövőbeni költségekre céltartalékot képezünk, amennyiben azok ' +
      'összege jelentős.', xmlDoc);
  }
  rewriteFact(p[866], p[867], mondat, '', xmlDoc);
  if (!txt(a, 'celtartalek_kepez')) {
    flagBefore(p[867],
      'nincs válasz arra, képez-e a vállalkozás lehetőség alapján ' +
      "céltartalékot (Szv#96) - a 'nem képez' ág lett feltételezve, " +
      'kérjük ellenőrizni/pontosítani.', xmlDoc);
  }
}

function resolveElhatarolasok(p, a, xmlDoc) {
  // Szv#99 - az időbeli elhatárolások dokumentálása.
  rewriteFact(p[872], p[873],
    txt(a, 'szv99')
      ? 'Az aktív és passzív időbeli elhatárolások dokumentálása és ' +
        `elszámolása a következők szerint történik: ${asSentence(txt(a, 'szv99'))}`
      : '',
    'nincs válasz arra, hogyan dokumentálják az időbeli elhatárolás ' +
    'tételeit (Szv#99) - az alábbi mondat a sablon konkrét, MÁS cégre ' +
    'írt gyakorlatát írja le.', xmlDoc);

  // Szv#83/#103/#104 - a sablon 896-899. bekezdései EGYÜTT írják le a
  // ZalaElektro gyakorlatát, ezért a teljes blokkot újraépítjük.
  let projekt;
  if (answer(a, 'szerzodes_elszamolasi_egyseg', 'nem_alkalmazzuk') === 'alkalmazzuk') {
    projekt = joinSentences([
      'A számviteli törvény 2020. évi módosításának megfelelően ' +
      'Társaságunk alkalmazza a szerződés elszámolási egységére ' +
      '(projektelszámolásra) vonatkozó szabályokat.',
      txt(a, 'szv103') ? 'A teljesítési fokot a következő módszerrel határozzuk meg: ' +
        `${txt(a, 'szv103')}` : '',
      txt(a, 'szv104') ? 'A készültségi fokot a következő módszerrel határozzuk meg: ' +
        `${txt(a, 'szv104')}` : '',
      'A teljes szerződéses ellenérték mérlegfordulónapi teljesítési ' +
      'fok arányában számított összege, valamint a már elszámolt ' +
      'árbevétel közötti különbözetet aktív vagy passzív időbeli ' +
      'elhatárolásként mutatjuk ki, az árbevétel növelésével vagy ' +
      'csökkentésével egyidejűleg.',
    ]);
  } else {
    projekt =
      'Társaságunk a nagy tömegben, sorozatban előállított termékek ' +
      'esetében a szerződés elszámolási egységére vonatkozó ' +
      'szabályrendszert nem alkalmazza.';
  }
  rewriteFact(p[896], p[897], projekt, '', xmlDoc);
  deleteP(p[898]);
  deleteP(p[899]);
  if (!txt(a, 'szerzodes_elszamolasi_egyseg')) {
    flagBefore(p[897],
      'nincs válasz arra, alkalmazzák-e a szerződés elszámolási ' +
      "egységére vonatkozó szabályrendszert (Szv#83) - a 'nem " +
      "alkalmazzuk' ág lett feltételezve, kérjük " +
      'ellenőrizni/pontosítani.', xmlDoc);
  }
}

// --- XVIII. fejezet - a sablonból átvett további pontok (nyilatkozatok) ---
//
// Ezek a bekezdések a törvényi hátteret ismertetik, majd az OLVASÓNAK szóló
// felszólítással zárulnak ("A Vállalkozásnak nyilatkoznia kell arról,
// hogy ..."). A felszólítás nem mehet ki az ügyfélhez: levágjuk
// (truncateAt), és a helyére a tényleges nyilatkozat kerül a jelölőbe.

function resolveNyilatkozatok(p, a, xmlDoc) {
  // Ért#4 - a hitelfelvételhez kapcsolódó költségek elhatárolása.
  truncateAt(p[919], 'A Vállalkozásnak nyilatkoznia kell', xmlDoc);
  let ert4;
  if (igenlo(answer(a, 'ert4_alkalmazza_elhatarolas'))) {
    const hatar = txt(a, 'bizomanyi_dij_hatar');
    ert4 =
      'A Vállalkozás él az elhatárolás lehetőségével: a hitel, kölcsön ' +
      'felvételéhez közvetlenül kapcsolódóan felmerült költségeket ' +
      'aktív időbeli elhatárolásként mutatja ki.';
    if (hatar) {
      ert4 += ' A bizományi díjat, vásárolt vételi opció díját jelentős ' +
        `összegűnek tekintjük, ha eléri a ${hatar} E Ft értéket.`;
    }
  } else {
    ert4 =
      'A Vállalkozás - a díjak összegétől függetlenül - nem él az ' +
      'elhatárolás lehetőségével, a hitel, kölcsön felvételéhez ' +
      'kapcsolódó költségeket a felmerüléskor költségként, ' +
      'ráfordításként számolja el.';
  }
  fillBlanks(p[920], [ert4], xmlDoc);
  if (!txt(a, 'ert4_alkalmazza_elhatarolas')) {
    flagBefore(p[920],
      'nincs válasz arra, élnek-e a hitelfelvételhez kapcsolódó ' +
      "költségek elhatárolásának lehetőségével (Ért#4) - a 'nem élünk " +
      "vele' ág lett feltételezve, kérjük ellenőrizni/pontosítani.", xmlDoc);
  }

  // Szv#29 - fenntarthatósági jelentés.
  truncateAt(p[923], 'A Vállalkozásnak ennek fennállását', xmlDoc);
  const fenn = answer(a, 'fenntarthatosagi_jelentes', '');
  const fennTabla = {
    nem_koteles:
      'A Vállalkozás a fenti feltételek alapján fenntarthatósági ' +
      'jelentés készítésére nem kötelezett.',
    igen_ugyanaz_a_konyvvizsgalo:
      'Vállalkozásunk fenntarthatósági jelentés készítésére ' +
      'kötelezett, amelynek vizsgálatával fenntarthatósági ' +
      'minősítéssel rendelkező olyan kamarai tag könyvvizsgálót, ' +
      'könyvvizsgáló céget bízunk meg, aki az éves beszámoló ' +
      'könyvvizsgálatát is végzi.',
    igen_mas_konyvvizsgalo:
      'Vállalkozásunk fenntarthatósági jelentés készítésére ' +
      'kötelezett, amelynek vizsgálatával az éves beszámoló ' +
      'jogszabályi kötelezettségen alapuló könyvvizsgálatával ' +
      'megbízott kamarai tag könyvvizsgálótól, könyvvizsgáló cégtől ' +
      'eltérő kamarai tag könyvvizsgálót, könyvvizsgáló céget bízunk ' +
      'meg.',
  };
  fillOrFlag(p[924],
    Object.prototype.hasOwnProperty.call(fennTabla, fenn) ? fennTabla[fenn] : '',
    'nincs válasz arra, van-e fenntarthatósági jelentés készítési ' +
    'kötelezettsége a társaságnak (Szv#29).', xmlDoc);

  // Szv#66/#67/#68 - környezetvédelmi közzétételek.
  truncateAt(p[927], 'A Vállalkozásnak nyilatkoznia kell', xmlDoc);
  fillOrFlag(p[928],
    joinSentences([
      txt(a, 'szv66') ? 'A környezet védelmét közvetlenül szolgáló tárgyi eszközök ' +
        `tekintetében: ${txt(a, 'szv66')}` : '',
      txt(a, 'szv67') ? `Veszélyes hulladék tekintetében: ${txt(a, 'szv67')}` : '',
      txt(a, 'szv68') ? `Környezetvédelmi kötelezettség tekintetében: ${txt(a, 'szv68')}` : '',
    ]),
    'nincs válasz a környezetvédelemmel összefüggő tárgyi eszközökre, a ' +
    'veszélyes hulladékra és a környezetvédelmi kötelezettségre ' +
    '(Szv#66, #67, #68).', xmlDoc);

  // Szv#75/#76/#84 - csak a kérdőívvel lefedett három kategóriát írjuk
  // felül, a többinél marad a sablon 20%-os alapértéke, mert az érvényes
  // alapértelmezés, nem kitöltetlen hely.
  setPlainRunText(p[931], 0,
    'A könyv szerinti érték és a piaci érték közötti különbséget az ' +
    'alábbi mértékek felett tekintjük jelentősnek:', xmlDoc);
  const mertekek = [];
  for (const [cimke, kulcs] of [
    ['Immateriális javak', 'szv75'],
    ['Tárgyi eszközök', null],
    ['Befektetett pénzügyi eszközök', 'szv76'],
    ['Készletek', 'szv84'],
    ['Követelések', null],
    ['Értékpapírok', null],
    ['Pénzeszközök', null],
    ['Részesedések', null],
  ]) {
    const ertek = kulcs ? txt(a, kulcs).replace(/%+$/, '').trim() : '';
    mertekek.push(`${cimke}: ${ertek || '20'}%`);
  }
  fillRun(p[932], 0, mertekek.join('  |  '), xmlDoc);
  deleteP(p[934]);

  // Szv#94 - a sablon szövege itt konkrétan a "Zala-Elektro Kft."-t nevezi meg.
  const ep = answer(a, 'ertekpapir_kamatkulonbozet_elhatarolas',
    'nem_hataroljuk_el') === 'elhataroljuk'
    ? 'A kamatozó értékpapír névértéke és beszerzési értéke közötti - ' +
      'nyereségjellegű - különbözetből a beszerzéstől az üzleti év ' +
      'mérlegfordulónapjáig terjedő időszakra időarányosan jutó ' +
      'összeget a Vállalkozás aktív időbeli elhatárolásként mutatja ki.'
    : 'A kamatozó értékpapír névértéke és beszerzési értéke közötti - ' +
      'nyereségjellegű - különbözetből a beszerzéstől az üzleti év ' +
      'mérlegfordulónapjáig terjedő időszakra időarányosan jutó ' +
      'összeget a Vállalkozás nem határolja el.';
  rewriteFact(p[937], p[938], ep, '', xmlDoc);
  if (!txt(a, 'ertekpapir_kamatkulonbozet_elhatarolas')) {
    flagBefore(p[938],
      'nincs válasz arra, elhatárolják-e a kamatozó értékpapír ' +
      'névértéke és beszerzési értéke közötti különbözetet (Szv#94) - ' +
      "a 'nem határoljuk el' ág lett feltételezve, kérjük " +
      'ellenőrizni/pontosítani.', xmlDoc);
  }

  // Szv#108 - társasági adó / KIVA.
  truncateAt(p[941], 'A Vállalkozásnak rögzítenie kell', xmlDoc);
  fillOrFlag(p[942],
    txt(a, 'szv108') ? `A Vállalkozás adóalanyisága: ${asSentence(txt(a, 'szv108'))}` : '',
    'nincs válasz arra, társasági adó vagy KIVA alanyok-e (Szv#108).', xmlDoc);

  // Szv#110 - globális minimumadó (GloBE).
  truncateAt(p[945], 'A sablon alapértelmezése', xmlDoc);
  fillOrFlag(p[946],
    txt(a, 'szv110') ? 'A Vállalkozás GloBE törvény szerinti kötelezettsége: ' +
      `${asSentence(txt(a, 'szv110'))}` : '',
    'nincs válasz arra, van-e a vállalkozásnak globális minimumadó ' +
    '(GloBE) kötelezettsége (Szv#110).', xmlDoc);
}

function resolveZaradek(p, a, xmlDoc) {
  // Szv#106 - ha a Vállalkozás mentesül az önköltségszámítási szabályzat
  // készítése alól, annak MINDKÉT felsorolásból ki kell esnie (87. és 954.),
  // különben a dokumentum két helyen mást mondana.
  const vanSzabalyzat =
    answer(a, 'onkoltsegszamitasi_szabalyzat', 'mentesul') !== 'mentesul';
  keepIf(p[87], vanSzabalyzat, null, xmlDoc);
  keepIf(p[954], vanSzabalyzat, null, xmlDoc);
  deleteP(p[955]);
}

async function generateFromAnswers(templateB64, answers) {
  const bin = atob(templateB64);
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  const zip = parseZip(bytes.buffer);
  const xmlText = await readEntryText(zip, 'word/document.xml');
  const xmlDoc = new DOMParser().parseFromString(xmlText, 'application/xml');
  // EGYETLEN alkalommal olvassuk ki a bekezdéslistát - minden resolve* és a
  // törlés is EBBE a listába indexel (ld. generate_policy.py megjegyzését).
  const p = getBodyParagraphs(xmlDoc);

  resolveCegadatok(p, answers, xmlDoc);
  resolveEcsErtekhatar(p, answers, xmlDoc);
  resolveSzamviteliFelelosok(p, answers, xmlDoc);
  resolveLeltar(p, answers, xmlDoc);
  resolvePenzkezeles(p, answers, xmlDoc);
  resolveKonyvvizsgalat(p, answers, xmlDoc);
  resolveKozzetetel(p, answers, xmlDoc);
  resolveBeszamoloForma(p, answers, xmlDoc);
  resolveMerlegForma(p, answers, xmlDoc);
  resolveEredmenykimutatasForma(p, answers, xmlDoc);
  resolveTetelekModositasa(p, answers, xmlDoc);
  resolveBeszamoloNyelve(p, answers, xmlDoc);
  resolveBeszamoloertFelelosseg(p, answers, xmlDoc);
  resolveKiegeszitoMelleklet(p, answers, xmlDoc);
  resolveKonszolidacio(p, answers, xmlDoc);
  resolveMinositesek(p, answers, xmlDoc);
  resolveKonyvvezetes(p, answers, xmlDoc);
  resolveBizonylatok(p, answers, xmlDoc);
  resolveZarlat(p, answers, xmlDoc);
  resolveBefektetettEszkozok(p, answers, xmlDoc);
  resolveErtekeles(p, answers, xmlDoc);
  resolveArfolyamok(p, answers, xmlDoc);
  resolveCeltartalek(p, answers, xmlDoc);
  resolveElhatarolasok(p, answers, xmlDoc);
  resolveNyilatkozatok(p, answers, xmlDoc);
  resolveZaradek(p, answers, xmlDoc);

  // Biztonsági háló a VÉGLEGES bekezdéslistán (ld. generate_policy.py).
  sweepLeftoverMarkers(getBodyParagraphs(xmlDoc), xmlDoc);

  // A flagBefore() által beszúrt "⚠ ELLENŐRIZENDŐ" bekezdések megszámolása -
  // a végleges (a flag-eket is tartalmazó) bekezdéslistán, nem a generálás
  // elején rögzített `p`-n, mert az utóbbi még a beszúrások előtti állapot.
  const flagCount = getBodyParagraphs(xmlDoc).filter(fp =>
    getRuns(fp).some(r => getRunText(r).includes('⚠ ELLENŐRIZENDŐ'))
  ).length;

  const newXml = new XMLSerializer().serializeToString(xmlDoc);
  const out = await buildZip(zip, new Map([['word/document.xml', new TextEncoder().encode(newXml)]]));
  return { bytes: out, flagCount };
}
