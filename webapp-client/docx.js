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

function clearChoiceMarks(p) {
  for (const r of getRuns(p)) clearRunMarks(r);
}

function fillBlanks(p, values, xmlDoc) {
  const runs = getRuns(p);
  const groups = blankGroups(p);
  groups.forEach((grp, gi) => {
    const val = values[gi];
    setRunText(runs[grp[0]], val == null ? '' : val, xmlDoc);
    for (let k = 1; k < grp.length; k++) setRunText(runs[grp[k]], '', xmlDoc);
  });
  stripBracketRef(p, xmlDoc);
  clearChoiceMarks(p);
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
  if (MONTHS[s]) return MONTHS[s];
  const m = s.match(/\d+/);
  return m ? parseInt(m[0], 10) : 12;
}

function isNumber(value) {
  if (value == null || value === '') return false;
  const n = Number(String(value).trim().replace(',', '.'));
  return !Number.isNaN(n);
}

function illustrativeDate(fordulonapHonap, fordulonapNap, offsetDays) {
  try {
    const base = new Date(2024, monthNum(fordulonapHonap) - 1, parseInt(fordulonapNap, 10));
    base.setDate(base.getDate() + parseInt(offsetDays || 0, 10));
    return [base.getMonth() + 1, base.getDate()];
  } catch (e) {
    return ['…', '…'];
  }
}
