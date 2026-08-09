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
    setRunText(runs[grp[0]], val == null ? '' : val, xmlDoc);
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

// --- dátum segéd ---

const MONTHS = {
  'január': 1, 'február': 2, 'március': 3, 'április': 4, 'május': 5,
  'június': 6, 'július': 7, 'augusztus': 8, 'szeptember': 9,
  'október': 10, 'november': 11, 'december': 12,
};

function monthNum(value) {
  const s = String(value).trim().toLowerCase();
  if (Object.prototype.hasOwnProperty.call(MONTHS, s)) return MONTHS[s];
  const m = s.match(/\d+/);
  return m ? parseInt(m[0], 10) : 12;
}

// A Python int(str) pontos tükörképe: csak egész szám fogadható el (nem
// "13.5", nem "abc"), opcionális előjellel és külső szóközökkel. Hibás
// bemenetre null - a Python ilyenkor ValueError-t dob.
function strictInt(value) {
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

function illustrativeDate(fordulonapHonap, fordulonapNap, offsetDays) {
  // A generate_policy.py illustrative_date() PONTOS tükörképe. Fontos: a
  // Python datetime.date szigorúan validál (hónap 1-12, a hónapban létező
  // nap), és hibás értékre kivételt dob -> "…". A JS Date ezzel szemben
  // csendben túlcsordul (2024. 13. hó -> 2025. január), ezért ugyanezeket a
  // feltételeket itt kézzel kell ellenőrizni, különben a két motor hibásan
  // kitöltött kérdőívre ELTÉRŐ dokumentumot adna.
  const m = monthNum(fordulonapHonap);
  const d = strictInt(fordulonapNap);
  const off = strictInt(offsetDays === '' || offsetDays == null ? 0 : offsetDays);
  if (d === null || off === null) return ['…', '…'];
  if (m < 1 || m > 12 || d < 1 || d > 31) return ['…', '…'];
  const base = new Date(2024, m - 1, d);
  if (base.getFullYear() !== 2024 || base.getMonth() !== m - 1 || base.getDate() !== d) {
    return ['…', '…'];  // pl. február 30. - a Python is ValueError-t dobna
  }
  base.setDate(base.getDate() + off);
  // A Python date-tartománya 1..9999 év; azon kívül OverflowError -> "…".
  if (!Number.isFinite(base.getTime()) || base.getFullYear() < 1 || base.getFullYear() > 9999) {
    return ['…', '…'];
  }
  return [base.getMonth() + 1, base.getDate()];
}
