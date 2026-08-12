// 1:1 port of generate_policy.py's resolve_* functions.

function removeJeloltInstructions(p) {
  for (const i of [0, 1, 2, 3]) deleteP(p[i]);
}

function resolveNyelv(p, a, xmlDoc) {
  if (a.van_idegen_nyelvu_beszamolo === 'igen') {
    fillBlanks(p[600], [a.idegen_nyelv || ''], xmlDoc);
    deleteP(p[597]);
    deleteP(p[599]);
  } else {
    keepP(p[597], xmlDoc);
    deleteP(p[599]);
    deleteP(p[600]);
  }
}

function resolvePenznem(p, a, xmlDoc) {
  const raw = answer(a, 'penznem_tipus', 'forint');
  const tipus = classifyPenznem(raw);
  if (tipus === 'forint') {
    keepP(p[616], xmlDoc);
    for (const i of [618, 620, 622, 624, 625, 627, 629, 631]) deleteP(p[i]);
  } else if (tipus === 'euro' || tipus === 'usd') {
    keepP(p[620], xmlDoc);
    if (tipus === 'euro') {
      keepP(p[622], xmlDoc);
      deleteP(p[624]);
      deleteP(p[625]);
    } else {
      deleteP(p[622]);
      deleteP(p[624]);
      keepP(p[625], xmlDoc);
    }
    fillBlanks(p[627], [a.penznem_atteres_datum || ''], xmlDoc);
    deleteP(p[616]);
    deleteP(p[618]);
    deleteP(p[629]);
    deleteP(p[631]);
  } else {
    let nev = (a.penznem_egyeb_neve || '').trim();
    if (!nev) {
      const rawNorm = deaccent(raw.trim().toLowerCase()).replace(/_/g, ' ');
      if (raw.trim() && !PENZNEM_EGYEB_PLACEHOLDERS.has(rawNorm)) nev = raw.trim();
    }
    fillBlanks(p[631], [nev], xmlDoc);
    for (const i of [616, 618, 620, 622, 624, 625, 627, 629]) deleteP(p[i]);
  }
  keepP(p[633], xmlDoc);
}

function resolveMerlegkeszites(p, a, xmlDoc) {
  const nap = a.merlegkeszites_nap || '';
  const honap = a.fordulonap_honap || '';
  const fnap = a.fordulonap_nap || '';
  const [illHonap, illNap] = illustrativeDate(honap, fnap, isNumber(nap) ? nap : 0);
  if (a.merlegkeszites_van_kivetel !== 'igen') {
    fillBlanks(p[662], [nap, illHonap, illNap], xmlDoc);
    deleteP(p[664]);
    deleteP(p[666]);
    deleteP(p[667]);
    deleteP(p[668]);
  } else {
    fillBlanks(p[666], [nap, illHonap, illNap], xmlDoc);
    fillBlanks(p[667], [a.merlegkeszites_kivetelek || ''], xmlDoc);
    deleteP(p[662]);
    deleteP(p[664]);
    deleteP(p[668]);
  }
}

function resolveFordulonap(p, a, xmlDoc) {
  fillBlanks(p[655], [a.fordulonap_honap || '', a.fordulonap_nap || ''], xmlDoc);
  deleteP(p[651]);
  deleteP(p[653]);
}

function resolveAlairo(p, a, xmlDoc) {
  fillBlanks(p[369], [a.beszamolo_alairoja || ''], xmlDoc);
}

function resolveKonyvvizsgalat(p, a, xmlDoc) {
  const tipus = answer(a, 'konyvvizsgalat_tipus', 'nincs_kotelezettseg');
  const adatai = a.konyvvizsgalo_adatai || '';
  const branches = { kotelezo_altalanos: 405, kotelezo_koztartozas_miatt: 409, onkentes: 413 };
  if (tipus === 'nincs_kotelezettseg') {
    keepP(p[401], xmlDoc);
    for (const i of [405, 409, 413]) deleteP(p[i]);
  } else {
    const chosen = branches[tipus];
    fillBlanks(p[chosen], [adatai], xmlDoc);
    for (const i of [401, 405, 409, 413]) if (i !== chosen) deleteP(p[i]);
  }
  for (const i of [403, 407, 411]) deleteP(p[i]);

  const fenn = answer(a, 'fenntarthatosagi_jelentes', 'nem_koteles');
  deleteP(p[415]);
  if (fenn === 'nem_koteles') {
    deleteP(p[417]); deleteP(p[419]); deleteP(p[421]);
  } else if (fenn === 'igen_ugyanaz_a_konyvvizsgalo') {
    keepP(p[417], xmlDoc); deleteP(p[419]); deleteP(p[421]);
  } else {
    deleteP(p[417]); deleteP(p[419]); keepP(p[421], xmlDoc);
  }
}

function resolveErtekelesiFelelos(p, a, xmlDoc) {
  if (answer(a, 'ertekeles_felelos_tipus') === 'egyedi_delegalas') {
    fillBlanks(p[338], [a.ertekelesi_delegalas || ''], xmlDoc);
    deleteP(p[334]);
    deleteP(p[336]);
  } else {
    keepP(p[334], xmlDoc);
    deleteP(p[336]);
    deleteP(p[338]);
  }
}

// A generate_policy.py combine_nev_regszam() PONTOS párja.
function combineNevRegszam(nev, regszam) {
  nev = (nev || '').trim();
  regszam = (regszam || '').trim();
  if (nev && regszam) return `${nev}, nyilvántartási szám: ${regszam}`;
  return nev || regszam;
}

function resolveKonyvelesFelelos(p, a, xmlDoc) {
  const tipus = answer(a, 'konyveles_felelos_tipus', 'mentesseg_20mft_arbevetel_alatt');
  const nev = combineNevRegszam(a.konyveles_felelos_neve, a.konyveles_felelos_regszam);
  if (tipus === 'merlegkepes_konyvelo') {
    fillBlanks(p[313], [nev], xmlDoc);
    keepP(p[312], xmlDoc);
    deleteP(p[317]); deleteP(p[318]); deleteP(p[322]);
  } else if (tipus === 'oklev_konyvvizsgalo') {
    fillBlanks(p[318], [nev], xmlDoc);
    keepP(p[317], xmlDoc);
    deleteP(p[312]); deleteP(p[313]); deleteP(p[322]);
  } else {
    keepP(p[322], xmlDoc);
    deleteP(p[312]); deleteP(p[313]); deleteP(p[317]); deleteP(p[318]);
  }
  deleteP(p[315]);
  deleteP(p[320]);
  keepP(p[324], xmlDoc);
  fillBlanks(p[964], [a.konyvvezetesert_felelos_szemely || ''], xmlDoc);
  fillBlanks(p[962], [a.konyvelo_program || ''], xmlDoc);
}

function resolveNyilvanossagraHozatal(p, a, xmlDoc) {
  keepP(p[441], xmlDoc);
  keepP(p[443], xmlDoc);
  if (a.beszamolo_honlapon === 'igen') {
    fillBlanks(p[445], [a.beszamolo_honlap_cim || ''], xmlDoc);
    deleteP(p[449]);
  } else {
    keepP(p[449], xmlDoc);
    deleteP(p[445]);
  }
  deleteP(p[447]);
}

function resolveKonszolidacio(p, a, xmlDoc) {
  deleteP(p[844]);
  const tipus = answer(a, 'konszolidacio_tipus', 'nem_erintett');
  const reszlet = a.konszolidacio_reszletszabalyok || '';
  const branchMap = {
    leany_mentesul_bevonas_alol: [[850], 851],
    leany_bevonva: [[855], 856],
    anya_mentesul_folerendelt_anya_miatt: [[860], 861],
    anya_mentesul_nagysagrend_alapjan: [[865, 868], 869],
    anya_keszit_konszolidaltat: [[873, 874], 875],
  };
  const allIdx = [846, 850, 851, 855, 856, 860, 861, 865, 868, 869, 873, 874, 875];
  if (tipus === 'nem_erintett') {
    keepP(p[846], xmlDoc);
    for (const i of allIdx) if (i !== 846) deleteP(p[i]);
  } else {
    const [keepStatic, fillIdx] = branchMap[tipus];
    for (const i of keepStatic) keepP(p[i], xmlDoc);
    fillBlanks(p[fillIdx], [reszlet], xmlDoc);
    for (const i of allIdx) if (i !== fillIdx && !keepStatic.includes(i)) deleteP(p[i]);
  }
  for (const i of [848, 853, 858, 863, 871]) deleteP(p[i]);
}

function resolveLeltar(p, a, xmlDoc) {
  const tipus = answer(a, 'leltar_felelos_tipus', 'vezetes_vagy_szv_rendert_felelos');
  if (tipus === 'vezetes_vagy_szv_rendert_felelos') {
    keepP(p[342], xmlDoc); deleteP(p[346]); deleteP(p[350]);
  } else if (tipus === 'eszkoz_forras_felelosok') {
    keepP(p[346], xmlDoc); deleteP(p[342]); deleteP(p[350]);
  } else {
    fillBlanks(p[350], [a.leltar_delegalas_reszletek || ''], xmlDoc);
    deleteP(p[342]); deleteP(p[346]);
  }
  deleteP(p[344]);
  deleteP(p[348]);
}

function resolveSimpleFills(p, a, xmlDoc) {
  fillBlanks(p[833], [a.jelentos_osszeg_bemutatas || ''], xmlDoc);
  deleteP(p[831]); deleteP(p[835]); deleteP(p[837]);

  fillBlanks(p[916], [a.jelentos_osszeg_ertelmezes || ''], xmlDoc);
  deleteP(p[918]); deleteP(p[920]);

  fillBlanks(p[908], [a.jelentos_mertek_hanyad || ''], xmlDoc);
  fillBlanks(p[937], [a.kiveteles_nagysag_osszeghatar || ''], xmlDoc);
  fillBlanks(p[939], [a.kiveteles_nagysag_sajattoke_szazalek || ''], xmlDoc);
  fillBlanks(p[941], [a.kiveteles_nagysag_bevetel_szazalek || '', a.kiveteles_nagysag_koltseg_szazalek || ''], xmlDoc);

  fillBlanks(p[1062], [a.bizonylat_konyveles_osszhang || ''], xmlDoc);
  deleteP(p[1054]); deleteP(p[1056]); deleteP(p[1058]); deleteP(p[1060]);

  fillBlanks(p[1015], [a.bizonylatkezeles_szabalyai || ''], xmlDoc);
  // 1153-nak KÉT kitöltendő helye van (lásd generate_policy.py megjegyzését).
  fillBlanks(p[1153], ['stb.', a.evkozi_zarasok_feladatai || ''], xmlDoc);
  fillBlanks(p[823], [a.kiegeszito_melleklet_sajatossagok || ''], xmlDoc);
}

function resolveAlapitasAtszervezes(p, a, xmlDoc) {
  if (a.alapitas_atszervezes_aktivalja === 'igen' && isNumber(a.alapitas_atszervezes_leiras_ev)) {
    fillBlanks(p[1810], [a.alapitas_atszervezes_leiras_ev], xmlDoc);
    deleteP(p[1805]); deleteP(p[1807]); deleteP(p[1809]);
  } else {
    flagBefore(p[1805],
      "az alapítás-átszervezés aktiválásáról/leírási idejéről szóló válasz " +
      "nem egyértelmű vagy 'nem aktiválunk' - kérjük kézzel pontosítani, " +
      "ill. ellenőrizni, hogy a kapcsolódó bekerülésiérték-szabály (VIII. " +
      "fejezet) is összhangban van-e.", xmlDoc);
  }
}

function resolveCegertek(p, a, xmlDoc) {
  if (a.van_cegertek === 'igen' && isNumber(a.cegertek_leiras_ev)) {
    fillBlanks(p[1815], [a.cegertek_leiras_ev], xmlDoc);
    deleteP(p[1812]); deleteP(p[1814]); deleteP(p[1817]); deleteP(p[1818]);
  } else {
    flagBefore(p[1812],
      "nincs üzleti/cégérték, vagy a leírási idő nem egyértelmű - " +
      "kérjük kézzel törölni/pontosítani a nem releváns bekezdéseket.", xmlDoc);
  }

  if (a.van_cegertek === 'igen' && a.van_negativ_cegertek === 'igen' && isNumber(a.negativ_cegertek_leiras_ev)) {
    fillBlanks(p[1699], [a.negativ_cegertek_leiras_ev], xmlDoc);
  } else {
    flagBefore(p[1699],
      "nincs negatív üzleti/cégérték, vagy a leírási idő nem egyértelmű - " +
      "kérjük kézzel törölni/pontosítani.", xmlDoc);
  }
}

function resolveErtekhelyesbites(p, a, xmlDoc) {
  const specs = [
    ['ertekhelyesbites_immat_van', 'ertekhelyesbites_immat_javak', 1274, 1276, 1278],
    ['ertekhelyesbites_targyi_van', 'ertekhelyesbites_targyi_eszkoz', 1299, 1301, 1303],
    ['ertekhelyesbites_penzugyi_van', 'ertekhelyesbites_penzugyi_eszkoz', 1344, 1346, 1348],
  ];
  for (const [gateKey, textKey, fillIdx, sepIdx, staticIdx] of specs) {
    if (a[gateKey] === 'igen') {
      fillBlanks(p[fillIdx], [a[textKey] || ''], xmlDoc);
      deleteP(p[staticIdx]);
    } else {
      keepP(p[staticIdx], xmlDoc);
      deleteP(p[fillIdx]);
    }
    deleteP(p[sepIdx]);
  }
}

function resolveCeltartalek(p, a, xmlDoc) {
  if (a.celtartalek_kepez === 'igen') {
    fillBlanks(p[1608], [''], xmlDoc);
    deleteP(p[1612]);
    fillBlanks(p[1614], [a.celtartalek_jelentos_hatar || ''], xmlDoc);
    fillBlanks(p[1630], [a.celtartalek_elteres_szazalek || ''], xmlDoc);
  } else {
    keepP(p[1612], xmlDoc);
    deleteP(p[1608]); deleteP(p[1614]);
    flagBefore(p[1630],
      "a vállalkozás nem képez céltartalékot lehetőség alapján - " +
      "kérjük ellenőrizni, hogy ez a bekezdés (lényeges eltérés %-a) " +
      "továbbra is szükséges-e, vagy törlendő.", xmlDoc);
  }
  deleteP(p[1610]);
}

function resolveBeszamoloForma(p, a, xmlDoc) {
  // Szv#13 - IV.2. beszámoló-forma ág + IV.8. kiegészítő melléklet blokk
  // (lásd a generate_policy.py részletes megjegyzését).
  const tipus = classifyBeszamoloForma(answer(a, 'szv13'));
  const branches = {
    eves_beszamolo: 487,
    egyszerusitett_eves_beszamolo: 490,
    mikrogazdalkodoi_egyszerusitett_eves_beszamolo: 493,
  };
  if (Object.prototype.hasOwnProperty.call(branches, tipus)) {
    const chosen = branches[tipus];
    keepP(p[chosen], xmlDoc);
    for (const i of [487, 490, 493]) if (i !== chosen) deleteP(p[i]);
    deleteP(p[489]);
    deleteP(p[492]);
  } else {
    flagBefore(p[485],
      "a beszámoló választott formája nem egyértelmű (pl. IFRS szerinti " +
      "beszámoló, amihez ez a sablon nem tartalmaz alternatívát) - kérjük " +
      "kézzel kiválasztani a megfelelő bekezdést, a többit törölni.", xmlDoc);
  }

  if (tipus === 'eves_beszamolo') {
    for (let i = 678; i < 774; i++) keepP(p[i], xmlDoc);
    for (let i = 774; i < 815; i++) deleteP(p[i]);
  } else if (tipus === 'egyszerusitett_eves_beszamolo') {
    for (let i = 678; i < 775; i++) deleteP(p[i]);
    for (let i = 775; i < 815; i++) keepP(p[i], xmlDoc);
  } else {
    flagBefore(p[676],
      "a kiegészítő melléklet tartalmához (éves / egyszerűsített éves " +
      "beszámoló változat) nincs egyértelmű kérdőív-adat - mikrogazdálkodói " +
      "beszámolónak egyáltalán nincs kiegészítő melléklete. Kérjük kézzel " +
      "kiválasztani a megfelelő blokkot, a másikat törölni.", xmlDoc);
    return;
  }
  deleteP(p[676]);
}

function resolveMerlegForma(p, a, xmlDoc) {
  const tipus = answer(a, 'szv14');
  if (tipus === 'a_valtozat') {
    keepP(p[525], xmlDoc); deleteP(p[527]);
  } else if (tipus === 'b_valtozat') {
    keepP(p[527], xmlDoc); deleteP(p[525]);
  } else {
    flagBefore(p[524], 'a mérleg választott formája (A/B változat) hiányzik.', xmlDoc);
    return;
  }
  deleteP(p[526]);
}

function resolveEredmenykimutatasForma(p, a, xmlDoc) {
  // 531-534 a törvény MINDKÉT eljárást ismertető leírása - nem nyúlunk hozzá.
  const tipus = answer(a, 'szv15');
  if (tipus === 'osszkoltseg_eljarassal') {
    keepP(p[537], xmlDoc); deleteP(p[539]);
  } else if (tipus === 'forgalmi_koltseg_eljarassal') {
    keepP(p[539], xmlDoc); deleteP(p[537]);
  } else {
    flagBefore(p[536], 'az eredménykimutatás választott formája hiányzik.', xmlDoc);
    return;
  }
  keepP(p[536], xmlDoc);
  keepP(p[540], xmlDoc);
  deleteP(p[538]);
}

function resolveLetetbehelyezes(p, a, xmlDoc) {
  if (answer(a, 'letetbehelyezes_hatarido_tipus', 'altalanos') === 'altalanos') {
    keepP(p[427], xmlDoc);
    deleteP(p[429]);
    deleteP(p[431]);
  } else {
    flagBefore(p[427],
      "a vállalkozás értékpapírjait EGT-állam szabályozott piacán " +
      "forgalmazzák, ezért a letétbe helyezés határideje a negyedik hónap " +
      "utolsó napja - kérjük a bekezdést kézzel összevonni/pontosítani.", xmlDoc);
  }
}

function resolvePenzkezelesFelelos(p, a, xmlDoc) {
  if (answer(a, 'penzkezeles_felelos_tipus', 'vezetes_altal_kijelolt') === 'penzkezelesi_szabalyzat_szerint') {
    keepP(p[358], xmlDoc); deleteP(p[354]);
  } else {
    keepP(p[354], xmlDoc); deleteP(p[358]);
  }
  deleteP(p[356]);
}

function resolveKoltsegelszamolas(p, a, xmlDoc) {
  const branches = { csak_5: 968, '5_elsodleges': 972, '67_elsodleges': 976 };
  const tipus = answer(a, 'koltsegelszamolas_tipus', 'csak_5');
  const chosen = Object.prototype.hasOwnProperty.call(branches, tipus) ? branches[tipus] : 968;
  keepP(p[chosen], xmlDoc);
  for (const i of [968, 972, 976]) if (i !== chosen) deleteP(p[i]);
  deleteP(p[970]);
  deleteP(p[974]);
}

function resolveSzerzodesElszamolasiEgyseg(p, a, xmlDoc) {
  // UGYANAZ a döntés NÉGY helyen szerepel a sablonban (ld. generate_policy.py).
  const alkalmaz = answer(a, 'szerzodes_elszamolasi_egyseg', 'nem_alkalmazzuk') === 'alkalmazzuk';
  for (const [lead, igen, sep, nem] of [
    [1374, 1375, 1376, 1377],
    [1479, 1480, 1482, 1484],
    [1676, 1677, 1679, 1681],
    [1885, 1886, 1887, 1888],
  ]) {
    keepP(p[lead], xmlDoc);
    if (alkalmaz) {
      keepP(p[igen], xmlDoc); deleteP(p[nem]);
    } else {
      keepP(p[nem], xmlDoc); deleteP(p[igen]);
    }
    deleteP(p[sep]);
  }
}

function resolveErtekpapirElhatarolas(p, a, xmlDoc) {
  if (answer(a, 'ertekpapir_kamatkulonbozet_elhatarolas', 'nem_hataroljuk_el') === 'elhataroljuk') {
    keepP(p[1431], xmlDoc); deleteP(p[1435]);
  } else {
    keepP(p[1435], xmlDoc); deleteP(p[1431]);
  }
  deleteP(p[1433]);
}

function resolveArfolyamveszteseg(p, a, xmlDoc) {
  if (answer(a, 'arfolyamveszteseg_halasztott_raforditas', 'nem_szamoljuk_el') === 'elszamoljuk') {
    keepP(p[1453], xmlDoc);
    deleteP(p[1455]);
    deleteP(p[1457]);
  } else {
    flagBefore(p[1453],
      "a vállalkozás NEM számolja el halasztott ráfordításként a devizás " +
      "hiteltartozások árfolyamveszteségét - a két bekezdést (1453/1457) " +
      "kérjük kézzel egy mondattá összevonni.", xmlDoc);
  }
}

function resolveOnkoltsegszamitas(p, a, xmlDoc) {
  if (answer(a, 'onkoltsegszamitasi_szabalyzat', 'mentesul') === 'mentesul') {
    keepP(p[188], xmlDoc);
  } else {
    flagBefore(p[188],
      "a Vállalkozás NEM mentesül az önköltségszámítási szabályzat " +
      "készítése alól (van/kell önköltségszámítási szabályzata) - az alábbi " +
      "mentességi bekezdés kézi átírást igényel.", xmlDoc);
  }
}

function resolveErt4(p, a, xmlDoc) {
  if (a.ert4_alkalmazza_elhatarolas === 'igen') {
    fillBlanks(p[1494], [a.bizomanyi_dij_hatar || ''], xmlDoc);
    keepP(p[1496], xmlDoc);
    deleteP(p[1500]);
  } else {
    keepP(p[1500], xmlDoc);
    deleteP(p[1494]); deleteP(p[1496]);
  }
  deleteP(p[1498]);
}

function resolveCegadatok(p, a, t, xmlDoc) {
  // Szv#1/2/5/6 - a fejléc cégadat-táblázata (ld. generate_policy.py).
  // Az első sor értékcellája HÁROM bekezdésből áll, ott a középsőbe írunk.
  const tbl = t[0];
  fillCell(cellParagraph(tbl, 0, 1, 1), a.szv1 || '', xmlDoc);
  fillCell(cellParagraph(tbl, 1, 1), a.szv2 || '', xmlDoc);
  fillCell(cellParagraph(tbl, 2, 1), a.szv6 || '', xmlDoc);
  fillCell(cellParagraph(tbl, 3, 1), a.cegjegyzekszam || '', xmlDoc);
  fillCell(cellParagraph(tbl, 4, 1), a.szv5 || '', xmlDoc);
}

function resolveBizonylatAlairas(p, a, t, xmlDoc) {
  // Szv#51/52/53 - VII.3.6. A bizonylatok hitelessége. A 1039 ághoz tartozik
  // a t[2] táblázat, ezért a másik ágban a táblázatot is törölni kell
  // (ld. generate_policy.py részletes megjegyzését).
  const tbl = t[2];
  const tipus = answer(a, 'bizonylat_hitelesites_tipus', 'kepviselo_alairasa');
  if (tipus === 'feljogositott_szemelyek') {
    keepP(p[1039], xmlDoc);
    deleteP(p[1035]);
    keepP(cellParagraph(tbl, 0, 0), xmlDoc);
    keepP(cellParagraph(tbl, 0, 1), xmlDoc);
    fillCell(cellParagraph(tbl, 1, 0), 'Kimenő számla', xmlDoc);
    fillCell(cellParagraph(tbl, 1, 1), a.szv52 || '', xmlDoc);
    cloneRow(tbl, 1);
    fillCell(cellParagraph(tbl, 2, 0), 'Egyéb kimenő számviteli bizonylat', xmlDoc);
    fillCell(cellParagraph(tbl, 2, 1), a.szv53 || '', xmlDoc);
  } else {
    keepP(p[1035], xmlDoc);
    deleteP(p[1039]);
    deleteTable(tbl);
  }
  deleteP(p[1037]);
}

function resolveErtekvesztesMertekek(p, a, t, xmlDoc) {
  // Szv#75/76/84/87 - kategóriánkénti "jelentős eltérés" %-os mértékek.
  // Csak a négy, kérdéssel lefedett sort írjuk felül, és csak ha van válasz -
  // különben marad a sablon 20%-os alapértéke (ld. generate_policy.py).
  const tbl = t[5];
  for (const [rowIdx, key] of [[2, 'szv75'], [4, 'szv76'], [5, 'szv84'], [6, 'szv87']]) {
    const value = answer(a, key, '');
    if (value !== '') fillCell(cellParagraph(tbl, rowIdx, 2), value, xmlDoc);
  }
}

async function generateFromAnswers(templateB64, answers) {
  const bin = atob(templateB64);
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  const zip = parseZip(bytes.buffer);
  const xmlText = await readEntryText(zip, 'word/document.xml');
  const xmlDoc = new DOMParser().parseFromString(xmlText, 'application/xml');
  const p = getBodyParagraphs(xmlDoc);
  // A bekezdéslistához hasonlóan a táblázatlistát is EGYSZER olvassuk ki -
  // a resolveBizonylatAlairas() egy egész táblázatot törölhet.
  const t = getTables(xmlDoc);

  resolveNyelv(p, answers, xmlDoc);
  resolvePenznem(p, answers, xmlDoc);
  resolveMerlegkeszites(p, answers, xmlDoc);
  resolveFordulonap(p, answers, xmlDoc);
  resolveAlairo(p, answers, xmlDoc);
  resolveKonyvvizsgalat(p, answers, xmlDoc);
  resolveErtekelesiFelelos(p, answers, xmlDoc);
  resolveKonyvelesFelelos(p, answers, xmlDoc);
  resolveNyilvanossagraHozatal(p, answers, xmlDoc);
  resolveKonszolidacio(p, answers, xmlDoc);
  resolveLeltar(p, answers, xmlDoc);
  resolveSimpleFills(p, answers, xmlDoc);
  resolveAlapitasAtszervezes(p, answers, xmlDoc);
  resolveCegertek(p, answers, xmlDoc);
  resolveErtekhelyesbites(p, answers, xmlDoc);
  resolveCeltartalek(p, answers, xmlDoc);
  resolveErt4(p, answers, xmlDoc);
  resolveBeszamoloForma(p, answers, xmlDoc);
  resolveMerlegForma(p, answers, xmlDoc);
  resolveEredmenykimutatasForma(p, answers, xmlDoc);
  resolveLetetbehelyezes(p, answers, xmlDoc);
  resolvePenzkezelesFelelos(p, answers, xmlDoc);
  resolveKoltsegelszamolas(p, answers, xmlDoc);
  resolveSzerzodesElszamolasiEgyseg(p, answers, xmlDoc);
  resolveErtekpapirElhatarolas(p, answers, xmlDoc);
  resolveArfolyamveszteseg(p, answers, xmlDoc);
  resolveOnkoltsegszamitas(p, answers, xmlDoc);
  resolveCegadatok(p, answers, t, xmlDoc);
  resolveBizonylatAlairas(p, answers, t, xmlDoc);
  resolveErtekvesztesMertekek(p, answers, t, xmlDoc);
  removeJeloltInstructions(p);

  const newXml = new XMLSerializer().serializeToString(xmlDoc);
  const out = await buildZip(zip, new Map([['word/document.xml', new TextEncoder().encode(newXml)]]));
  return out;
}
