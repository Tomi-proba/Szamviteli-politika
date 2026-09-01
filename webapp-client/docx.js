// DOCX paragraph/run manipulation helpers, mirroring the semantics of the
// python-docx based generate_policy.py (paragraph = direct w:p child of
// w:body only, run = direct w:r child of a paragraph).

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

function getBodyParagraphs(xmlDoc) {
  const body = xmlDoc.getElementsByTagNameNS(W_NS, 'body')[0];
  return Array.from(body.children).filter(el => el.localName === 'p');
}

function getRuns(p) {
  return Array.from(p.children).filter(el => el.localName === 'r');
}

// --- táblázat-hozzáférés (a generate_policy.py table_rows()/row_cells()/
// cell_paragraph() függvényeinek PONTOS párja) ---
//
// A python-docx `doc.tables` a törzs közvetlen w:tbl gyerekeit adja, ezért
// itt is csak azokat vesszük. A sorokat/cellákat szándékosan a NYERS
// w:tr / w:tc gyerekekből indexeljük, nem a python-docx rács-nézetéből: az
// a vízszintesen összevont (gridSpan) cellákat többször is felsorolná, és a
// két motor cellaindexei elcsúsznának egymástól.

function getTables(xmlDoc) {
  const body = xmlDoc.getElementsByTagNameNS(W_NS, 'body')[0];
  return Array.from(body.children).filter(el => el.localName === 'tbl');
}

function getRows(tbl) {
  return Array.from(tbl.children).filter(el => el.localName === 'tr');
}

function getCells(tr) {
  return Array.from(tr.children).filter(el => el.localName === 'tc');
}

function cellParagraph(tbl, rowIdx, cellIdx, paraIdx) {
  const tc = getCells(getRows(tbl)[rowIdx])[cellIdx];
  const ps = Array.from(tc.children).filter(el => el.localName === 'p');
  return ps[paraIdx || 0];
}

function getRPr(r) {
  return Array.from(r.children).find(el => el.localName === 'rPr') || null;
}

function ensureRPr(r, xmlDoc) {
  let rPr = getRPr(r);
  if (!rPr) {
    rPr = xmlDoc.createElementNS(W_NS, 'w:rPr');
    r.insertBefore(rPr, r.firstChild);
  }
  return rPr;
}

function getRunText(r) {
  const ts = Array.from(r.getElementsByTagNameNS(W_NS, 't'));
  return ts.map(t => t.textContent).join('');
}

// A generate_policy.py text_value() függvényének PONTOS párja - ld. ott a
// részletes magyarázatot. Röviden: a Python str() és a JS String() nem
// ekvivalens minden típusra (str(20.0)="20.0" vs String(20.0)="20",
// str(True)="True" vs String(true)="true"), ezért a behelyettesítendő
// értékeket MINDIG ezen keresztül kell szöveggé alakítani, hogy a két motor
// ugyanazt a dokumentumot adja.
function textValue(value) {
  if (value === undefined || value === null) return '';
  if (typeof value === 'boolean') return value ? 'true' : 'false';
  return String(value);
}

function setRunText(r, text, xmlDoc) {
  let ts = Array.from(r.getElementsByTagNameNS(W_NS, 't'));
  if (ts.length === 0) {
    const t = xmlDoc.createElementNS(W_NS, 'w:t');
    r.appendChild(t);
    ts = [t];
  }
  ts[0].textContent = text == null ? '' : String(text);
  ts[0].setAttribute('xml:space', 'preserve');
  for (let i = 1; i < ts.length; i++) ts[i].textContent = '';
}

function hasHighlight(r) {
  const rPr = getRPr(r);
  if (!rPr) return false;
  return Array.from(rPr.children).some(el => el.localName === 'highlight');
}

function hasStrike(r) {
  const rPr = getRPr(r);
  if (!rPr) return false;
  const el = Array.from(rPr.children).find(c => c.localName === 'strike');
  if (!el) return false;
  const val = el.getAttributeNS(W_NS, 'val');
  return !(val === '0' || val === 'false' || val === 'none');
}

function hasColor(r) {
  const rPr = getRPr(r);
  if (!rPr) return false;
  return Array.from(rPr.children).some(el => el.localName === 'color');
}

function markSubstituted(r, xmlDoc) {
  // A sablon eredeti "kitöltendő hely" jelölését (sárga kiemelés, áthúzás,
  // piros betűszín) alkalmazza egy behelyettesített válaszra - a
  // felülvizsgáló munkatárs ebből lássa, mi került be automatikusan.
  clearRunMarks(r);
  const rPr = ensureRPr(r, xmlDoc);
  const highlight = xmlDoc.createElementNS(W_NS, 'w:highlight');
  highlight.setAttributeNS(W_NS, 'w:val', 'yellow');
  rPr.appendChild(highlight);
  const strike = xmlDoc.createElementNS(W_NS, 'w:strike');
  rPr.appendChild(strike);
  const color = xmlDoc.createElementNS(W_NS, 'w:color');
  color.setAttributeNS(W_NS, 'w:val', 'C00000');
  rPr.appendChild(color);
}

function clearRunMarks(r) {
  const rPr = getRPr(r);
  if (!rPr) return;
  for (const name of ['highlight', 'strike', 'color']) {
    const el = Array.from(rPr.children).find(c => c.localName === name);
    if (el) rPr.removeChild(el);
  }
}

// --- bekezdés/run szintű segédfüggvények (a generate_policy.py tükörképe) ---

function blankGroups(p) {
  const runs = getRuns(p);
  const groups = [];
  let current = [];
  const isStrikeArr = runs.map(hasStrike);
  for (let i = 0; i < runs.length; i++) {
    const strike = isStrikeArr[i];
    const emptyNonStrike = !strike && getRunText(runs[i]).trim() === '';
    if (strike || emptyNonStrike) {
      current.push(i);
    } else {
      if (current.length && current.some(j => isStrikeArr[j])) groups.push(current);
      current = [];
    }
  }
  if (current.length && current.some(j => isStrikeArr[j])) groups.push(current);
  return groups;
}

function stripBracketRef(p, xmlDoc) {
  for (const r of getRuns(p)) {
    if (hasHighlight(r) && !hasStrike(r) && /\[\s*(Szv|Ért)/.test(getRunText(r))) {
      setRunText(r, '', xmlDoc);
    }
  }
}

function clearChoiceMarks(p, exceptRuns) {
  const skip = exceptRuns || new Set();
  getRuns(p).forEach((r, i) => {
    if (!skip.has(i)) clearRunMarks(r);
  });
}

function fillBlanks(p, values, xmlDoc) {
  const runs = getRuns(p);
  const groups = blankGroups(p);
  const markedIndices = new Set();
  groups.forEach((grp, gi) => {
    const val = values[gi];
    setRunText(runs[grp[0]], textValue(val), xmlDoc);
    markSubstituted(runs[grp[0]], xmlDoc);
    markedIndices.add(grp[0]);
    for (let k = 1; k < grp.length; k++) setRunText(runs[grp[k]], '', xmlDoc);
  });
  stripBracketRef(p, xmlDoc);
  clearChoiceMarks(p, markedIndices);
}

function keepP(p, xmlDoc) {
  stripBracketRef(p, xmlDoc);
  clearChoiceMarks(p);
}

function deleteP(p) {
  if (p.parentNode) p.parentNode.removeChild(p);
}

// A generate_policy.py fill_cell() PONTOS párja: a sablon táblázataiban
// egyetlen áthúzott "kitöltendő hely" futás sincs, ezért itt nem a
// blankGroups()-ra építünk, hanem a cella első futását írjuk felül (üres
// cellába új futást hozunk létre), és ugyanazzal a jelöléssel hagyjuk, mint
// a törzsben.
function fillCell(p, value, xmlDoc) {
  const text = textValue(value);
  const runs = getRuns(p);
  if (runs.length) {
    setRunText(runs[0], text, xmlDoc);
    markSubstituted(runs[0], xmlDoc);
    for (let i = 1; i < runs.length; i++) setRunText(runs[i], '', xmlDoc);
    clearChoiceMarks(p, new Set([0]));
  } else {
    const r = xmlDoc.createElementNS(W_NS, 'w:r');
    const t = xmlDoc.createElementNS(W_NS, 'w:t');
    t.setAttribute('xml:space', 'preserve');
    t.textContent = text;
    r.appendChild(t);
    p.appendChild(r);
    markSubstituted(r, xmlDoc);
  }
}

// A generate_policy.py delete_table() párja.
function deleteTable(tbl) {
  if (tbl.parentNode) tbl.parentNode.removeChild(tbl);
}

// A generate_policy.py clone_row() párja: a sor másolatát közvetlenül az
// eredeti után szúrja be.
function cloneRow(tbl, rowIdx) {
  const tr = getRows(tbl)[rowIdx];
  const newTr = tr.cloneNode(true);
  tr.parentNode.insertBefore(newTr, tr.nextSibling);
  return newTr;
}

function flagBefore(p, note, xmlDoc) {
  const newP = xmlDoc.createElementNS(W_NS, 'w:p');
  const r = xmlDoc.createElementNS(W_NS, 'w:r');
  const rPr = xmlDoc.createElementNS(W_NS, 'w:rPr');
  const b = xmlDoc.createElementNS(W_NS, 'w:b');
  const color = xmlDoc.createElementNS(W_NS, 'w:color');
  color.setAttributeNS(W_NS, 'w:val', 'C00000');
  rPr.appendChild(b);
  rPr.appendChild(color);
  r.appendChild(rPr);
  const t = xmlDoc.createElementNS(W_NS, 'w:t');
  t.setAttribute('xml:space', 'preserve');
  t.textContent = `⚠ ELLENŐRIZENDŐ (nincs kérdőív-adat): ${note}`;
  r.appendChild(t);
  newP.appendChild(r);
  p.parentNode.insertBefore(newP, p);
}

// --- futás-szintű szerkesztés: a "D1" mondat-újraíráshoz és a formázatlan
// helykitöltőkhöz (a generate_policy.py azonos nevű primitívjeinek PONTOS
// párjai - ld. ott a részletes magyarázatot) ---

// A sablon jelölőinek szövegkezdete és a formázatlan helykitöltő - a
// biztonsági söprés (sweepLeftoverMarkers) ezekre keres.
const MARKER_PREFIX = '▶ Adatbékérő';
const GAP_PLACEHOLDER = 'innen hiányzik';

// A generate_policy.py set_plain_run_text() párja: futás szövegének
// felülírása JELÖLÉS NÉLKÜL (nyelvtani ragasztóelemekhez).
function setPlainRunText(p, runIdx, text, xmlDoc) {
  setRunText(getRuns(p)[runIdx], text, xmlDoc);
}

// A generate_policy.py fill_run() párja: MEGLÉVŐ futás lecserélése a
// válaszra, a szokásos kitöltendő-jelöléssel.
function fillRun(p, runIdx, value, xmlDoc) {
  const r = getRuns(p)[runIdx];
  setRunText(r, textValue(value), xmlDoc);
  markSubstituted(r, xmlDoc);
}

// A generate_policy.py substitute_in_run() párja: a futást HÁROM futásra
// bontja (before / value / after), a középsőt megjelöli. Mindig pontosan
// három futás keletkezik, hogy a két motor futás-indexei bitre egyezzenek.
function substituteInRun(p, runIdx, before, value, after, xmlDoc) {
  const original = getRuns(p)[runIdx];
  const mid = original.cloneNode(true);
  const tail = original.cloneNode(true);
  original.parentNode.insertBefore(tail, original.nextSibling);
  original.parentNode.insertBefore(mid, original.nextSibling);
  const runs = getRuns(p);
  setRunText(runs[runIdx], before, xmlDoc);
  setRunText(runs[runIdx + 1], textValue(value), xmlDoc);
  markSubstituted(runs[runIdx + 1], xmlDoc);
  setRunText(runs[runIdx + 2], after, xmlDoc);
  clearChoiceMarks(p, new Set([runIdx + 1]));
  return runIdx + 1;
}

// A generate_policy.py fill_text_placeholder() párja: a sablonban FORMÁZÁS
// NÉLKÜL hagyott helykitöltőt (pl. "innen hiányzik") szöveg szerint keresi
// meg és cseréli le.
function fillTextPlaceholder(p, placeholder, value, xmlDoc) {
  const runs = getRuns(p);
  for (let i = 0; i < runs.length; i++) {
    const t = getRunText(runs[i]);
    const pos = t.indexOf(placeholder);
    if (pos >= 0) {
      substituteInRun(p, i, t.slice(0, pos), value,
        t.slice(pos + placeholder.length), xmlDoc);
      return true;
    }
  }
  return false;
}

// A generate_policy.py prepend_sentence() párja.
function prependSentence(p, text, xmlDoc) {
  const runs = getRuns(p);
  if (!runs.length) return false;
  substituteInRun(p, 0, '', text, getRunText(runs[0]), xmlDoc);
  return true;
}

// A generate_policy.py truncate_at() párja: a bekezdés szövegéből a megadott
// résztől kezdve mindent levág (az olvasónak szóló sablon-felszólításokat).
// A Python str.rstrip() párja a /\s+$/ csere.
function truncateAt(p, marker, xmlDoc) {
  let found = false;
  for (const r of getRuns(p)) {
    if (found) { setRunText(r, '', xmlDoc); continue; }
    const t = getRunText(r);
    const pos = t.indexOf(marker);
    if (pos >= 0) {
      setRunText(r, t.slice(0, pos).replace(/\s+$/, ''), xmlDoc);
      found = true;
    }
  }
  return found;
}

// A generate_policy.py drop_leading_word() párja.
function dropLeadingWord(p, runIdx, word, xmlDoc) {
  const r = getRuns(p)[runIdx];
  const t = getRunText(r);
  const prefix = word + ' ';
  if (t.startsWith(prefix)) {
    setRunText(r, t.slice(prefix.length), xmlDoc);
    return true;
  }
  return false;
}

// A generate_policy.py restate() párja: JELÖLETLEN sablonmondat teljes
// újraírása a válaszból (ugyanaz a tény a sablon több pontján is szerepel).
function restate(p, text, xmlDoc) {
  const runs = getRuns(p);
  if (!runs.length) return false;
  setRunText(runs[0], textValue(text), xmlDoc);
  markSubstituted(runs[0], xmlDoc);
  for (let i = 1; i < runs.length; i++) setRunText(runs[i], '', xmlDoc);
  clearChoiceMarks(p, new Set([0]));
  return true;
}

// --- a sablon szerzőjének feltételes utasításai ("D2") ---
//
// A generate_policy.py _AUTHOR_INSTRUCTION_RE / strip_author_instruction() /
// keep_if() PONTOS párjai. A szerzői utasítás ("ezt csak akkor írja ha ...")
// SOSEM kerülhet ki az ügyfélhez, akkor sem, ha a hozzá tartozó mondat marad.
const AUTHOR_INSTRUCTION_RE = /\bez(?:t)?\s+(?:csak\s+)?akkor\b/i;

function stripAuthorInstruction(p, xmlDoc) {
  let removed = false;
  for (const r of getRuns(p)) {
    if (AUTHOR_INSTRUCTION_RE.test(getRunText(r))) {
      setRunText(r, '', xmlDoc);
      removed = true;
    }
  }
  return removed;
}

// FELTÉTELES BEKEZDÉS-MEGTARTÁS: a szerzői utasítást mindig kiveszi; ha a
// feltétel igaz, a bekezdés marad, egyébként törlődik (opcionális
// ⚠ ELLENŐRIZENDŐ jelöléssel).
function keepIf(p, condition, note, xmlDoc) {
  stripAuthorInstruction(p, xmlDoc);
  if (condition) { keepP(p, xmlDoc); return true; }
  if (note) flagBefore(p, note, xmlDoc);
  deleteP(p);
  return false;
}

// --- mondat-újraírás ("D1") és jelölő-kitöltés ---

// A generate_policy.py rewrite_fact() párja.
//
// `fallback` megadva: üres válasznál a sablon eredeti mondatát ilyenkor is
// TÖRÖLJÜK, és a semleges, cégfüggetlen `fallback` kerül be helyette - ezt
// kell használni ott, ahol a sablon mondata egy MÁSIK cég azonosítható
// adatát (nevét, regisztrációs számát, rendszernevét) tartalmazza.
// `fallback` nélkül a sablon általános alapértelmezése marad, jelöléssel.
function rewriteFact(factP, markerP, text, note, xmlDoc, fallback) {
  const value = textValue(text).trim();
  if (value) {
    fillBlanks(markerP, [value], xmlDoc);
    deleteP(factP);
    return true;
  }
  if (fallback) {
    fillBlanks(markerP, [fallback], xmlDoc);
    deleteP(factP);
    flagBefore(markerP, note, xmlDoc);
    return false;
  }
  deleteP(markerP);
  flagBefore(factP, note, xmlDoc);
  keepP(factP, xmlDoc);
  return false;
}

// A generate_policy.py fill_or_flag() párja.
function fillOrFlag(markerP, text, note, xmlDoc) {
  const value = textValue(text).trim();
  if (value) { fillBlanks(markerP, [value], xmlDoc); return true; }
  flagBefore(markerP, note, xmlDoc);
  deleteP(markerP);
  return false;
}

// A generate_policy.py flag_unwired() párja.
function flagUnwired(markerP, note, xmlDoc) {
  flagBefore(markerP, note, xmlDoc);
  deleteP(markerP);
}

// A generate_policy.py sweep_leftover_markers() párja - BIZTONSÁGI HÁLÓ a
// generálás végén, a VÉGLEGES bekezdéslistán.
function sweepLeftoverMarkers(paragraphs, xmlDoc) {
  let count = 0;
  for (const para of paragraphs) {
    const runs = getRuns(para);
    const text = runs.map(getRunText).join('');
    if (!text.includes(MARKER_PREFIX) && !text.includes(GAP_PLACEHOLDER)) continue;
    for (const r of runs) {
      setRunText(r, getRunText(r).split(MARKER_PREFIX).join('')
        .split(GAP_PLACEHOLDER).join(''), xmlDoc);
    }
    flagBefore(para,
      'a sablon egy kitöltetlen jelölője/helykitöltője maradt itt - a ' +
      'motor kivette a szövegét, de a bekezdést kérjük kézzel átnézni.', xmlDoc);
    count += 1;
  }
  return count;
}

// --- dátum segéd ---

const MONTHS = {
  'január': 1, 'február': 2, 'március': 3, 'április': 4, 'május': 5,
  'június': 6, 'július': 7, 'augusztus': 8, 'szeptember': 9,
  'október': 10, 'november': 11, 'december': 12,
};

const MONTH_NAMES = [
  'január', 'február', 'március', 'április', 'május', 'június',
  'július', 'augusztus', 'szeptember', 'október', 'november', 'december',
];

function monthNum(value) {
  const s = String(value).trim().toLowerCase();
  if (Object.prototype.hasOwnProperty.call(MONTHS, s)) return MONTHS[s];
  const m = s.match(/\d+/);
  return m ? parseInt(m[0], 10) : 12;
}

// A generate_policy.py month_name() párja: hónapszám -> magyar hónapnév,
// érvénytelen értékre üres sztring (a hívó ilyenkor jelöl).
function monthName(num) {
  const n = strictInt(num);
  if (n === null || n < 1 || n > 12) return '';
  return MONTH_NAMES[n - 1];
}

// A Python int(str) pontos tükörképe: csak egész szám fogadható el (nem
// "13.5", nem "abc"), opcionális előjellel és külső szóközökkel. Hibás
// bemenetre null - a Python ilyenkor ValueError-t dob.
function strictInt(value) {
  if (typeof value === 'boolean') return null;
  if (typeof value === 'number') return Number.isInteger(value) ? value : null;
  const s = String(value).trim();
  if (!/^[+-]?\d+$/.test(s)) return null;
  return parseInt(s, 10);
}

function isNumber(value) {
  if (value == null || value === '') return false;
  const n = Number(String(value).trim().replace(',', '.'));
  return !Number.isNaN(n);
}

// A generate_policy.py answer() függvényének PONTOS párja: a hiányzó és az
// üresen hagyott mezőt egyaránt "nincs válasz"-ként kezeli. NE cseréljük
// vissza `a.kulcs || alapertelmezes` alakra - az üres sztringre az másképp
// viselkedne, mint a Python `a.get(kulcs, alapertelmezes)`, és a két motor
// kimenete eltérne (az iroda-oldalon minden kitöltetlen select ÜRES sztring).
function answer(a, key, def_) {
  const fallback = def_ === undefined ? '' : def_;
  const value = a[key];
  if (value === undefined || value === null) return fallback;
  if (typeof value === 'string' && value.trim() === '') return fallback;
  return value;
}

const PENZNEM_EGYEB_PLACEHOLDERS = new Set(['egyeb', 'egyeb deviza', 'egyeb devizanem', 'mas', 'mas deviza']);

function deaccent(s) {
  return s.normalize('NFD').replace(/[̀-ͯ]/g, '');
}

// A generate_policy.py classify_penznem() PONTOS párja - lásd ott a
// magyarázatot (szabadon gépelt Szv#12 pénznem-mező besorolása).
function classifyPenznem(raw) {
  const t = (raw || '').trim();
  const tl = deaccent(t.toLowerCase());
  if (tl === '' || tl === 'forint' || tl === 'huf') return 'forint';
  if (tl === 'euro' || tl === 'eur') return 'euro';
  if (tl === 'usd' || tl === 'dollar' || tl === 'us dollar' || tl === 'amerikai dollar') return 'usd';
  return 'egyeb';
}

// A generate_policy.py classify_beszamolo_forma() PONTOS párja - a kérdőív
// (Munka3 munkalap) NÉGY ismert válaszára sorol be, az IFRS-es változatot is
// beleértve; amit nem tud besorolni, azt változatlanul visszaadja.
function classifyBeszamoloForma(raw) {
  const t = deaccent((raw || '').trim().toLowerCase()).replace(/_/g, ' ');
  if (t.includes('ifrs')) return 'eves_beszamolo_ifrs';
  if (t.includes('mikro')) return 'mikrogazdalkodoi_egyszerusitett_eves_beszamolo';
  if (t.includes('egyszerus')) return 'egyszerusitett_eves_beszamolo';
  if (t.includes('eves')) return 'eves_beszamolo';
  return raw;
}

// A generate_policy.py illustrative_date() PONTOS tükörképe: a fordulónap +
// offset (hónap, nap, ÉV-eltolódás) hármasa, hibás bemenetre [null,null,null].
//
// Fontos: a Python datetime.date szigorúan validál (hónap 1-12, a hónapban
// létező nap), és hibás értékre kivételt dob; a JS Date ezzel szemben csendben
// túlcsordul (2024. 13. hó -> 2025. január), ezért ugyanezeket a feltételeket
// itt kézzel kell ellenőrizni - NE egyszerűsítsd vissza. Az egész-értelmezés
// mindkét oldalon a strictInt()/strict_int() szigorúságát követi.
function illustrativeDate(fordulonapHonap, fordulonapNap, offsetDays) {
  const m = monthNum(fordulonapHonap);
  const d = strictInt(fordulonapNap);
  const off = strictInt(offsetDays);
  if (d === null || off === null || m < 1 || m > 12 || d < 1 || d > 31) {
    return [null, null, null];
  }
  const base = new Date(2024, m - 1, d);
  if (base.getFullYear() !== 2024 || base.getMonth() !== m - 1 || base.getDate() !== d) {
    return [null, null, null];  // pl. február 30. - a Python is ValueError-t dobna
  }
  base.setDate(base.getDate() + off);
  if (!Number.isFinite(base.getTime()) || base.getFullYear() < 1 || base.getFullYear() > 9999) {
    return [null, null, null];
  }
  return [base.getMonth() + 1, base.getDate(), base.getFullYear() - 2024];
}

// --- szöveg-segédek (a generate_policy.py azonos nevű függvényeinek párjai) ---

// txt(): szabadszöveges válasz kiolvasása, körülvágva.
function txt(a, key) {
  return textValue(answer(a, key, '')).trim();
}

// A generate_policy.py hu_article() párja: magyar határozott névelő.
function huArticle(word) {
  const s = textValue(word).trim();
  if (!s) return 'a';
  const first = deaccent(s[0].toLowerCase());
  return 'aeiou'.includes(first) ? 'az' : 'a';
}

// A generate_policy.py inline_value() párja: a sormintában álló jelölő értéke
// hozza magával az elválasztó szóközöket (üres értéknél nem marad lógó szóköz).
function inlineValue(text, prefix, suffix) {
  const s = textValue(text).trim();
  if (!s) return '';
  return (prefix === undefined ? '' : prefix) + s + (suffix === undefined ? ' ' : suffix);
}

// A generate_policy.py join_clauses() párja.
function joinClauses(parts, sep) {
  const sp = sep === undefined ? ' ' : sep;
  return parts.map(x => textValue(x).trim()).filter(x => x).join(sp);
}

// A generate_policy.py as_sentence() párja: mondatzáró írásjel pótlása.
function asSentence(text) {
  const s = textValue(text).trim();
  if (!s) return '';
  return '.!?:;'.includes(s[s.length - 1]) ? s : s + '.';
}

// A generate_policy.py join_sentences() párja.
function joinSentences(parts) {
  return parts.map(asSentence).filter(x => x).join(' ');
}

// A generate_policy.py igenlo() / nemleges() párjai: szabadon gépelt igen/nem.
function igenlo(value) {
  return deaccent(textValue(value).trim().toLowerCase()).startsWith('igen');
}

function nemleges(value) {
  return deaccent(textValue(value).trim().toLowerCase()).startsWith('nem');
}

// A generate_policy.py penznem_neve() párja: a pénznem KIÍRHATÓ neve.
function penznemNeve(a) {
  const raw = textValue(answer(a, 'penznem_tipus', 'forint')).trim();
  const tipus = classifyPenznem(raw);
  if (tipus === 'forint') return 'forint';
  if (tipus === 'euro') return 'euró';
  if (tipus === 'usd') return 'USA dollár';
  const nev = txt(a, 'penznem_egyeb_neve');
  if (nev) return nev;
  const rawNorm = deaccent(raw.toLowerCase()).replace(/_/g, ' ');
  if (raw && !PENZNEM_EGYEB_PLACEHOLDERS.has(rawNorm)) return raw;
  return '';
}
