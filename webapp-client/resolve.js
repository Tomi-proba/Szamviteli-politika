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
  const tipus = a.penznem_tipus || 'forint';
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
    fillBlanks(p[631], [a.penznem_egyeb_neve || ''], xmlDoc);
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
  const tipus = a.konyvvizsgalat_tipus || 'nincs_kotelezettseg';
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

  const fenn = a.fenntarthatosagi_jelentes || 'nem_koteles';
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
  if (a.ertekeles_felelos_tipus === 'egyedi_delegalas') {
    fillBlanks(p[338], [a.ertekelesi_delegalas || ''], xmlDoc);
    deleteP(p[334]);
    deleteP(p[336]);
  } else {
    keepP(p[334], xmlDoc);
    deleteP(p[336]);
    deleteP(p[338]);
  }
}

function resolveKonyvelesFelelos(p, a, xmlDoc) {
  const tipus = a.konyveles_felelos_tipus || 'mentesseg_20mft_arbevetel_alatt';
  const nev = a.konyveles_felelos_nev_regszam || '';
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
  const tipus = a.konszolidacio_tipus || 'nem_erintett';
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
  const tipus = a.leltar_felelos_tipus || 'vezetes_vagy_szv_rendert_felelos';
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

async function generateFromAnswers(templateB64, answers) {
  const bin = atob(templateB64);
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  const zip = parseZip(bytes.buffer);
  const xmlText = await readEntryText(zip, 'word/document.xml');
  const xmlDoc = new DOMParser().parseFromString(xmlText, 'application/xml');
  const p = getBodyParagraphs(xmlDoc);

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
  removeJeloltInstructions(p);

  const newXml = new XMLSerializer().serializeToString(xmlDoc);
  const out = await buildZip(zip, new Map([['word/document.xml', new TextEncoder().encode(newXml)]]));
  return out;
}
