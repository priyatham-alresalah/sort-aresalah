export function normalizeName(value) {
  return String(value || "")
    .toUpperCase()
    .replace(/[,.`']/g, "")
    .replace(/[^A-Z0-9\s]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function tokens(value) {
  return normalizeName(value)
    .split(" ")
    .filter((t) => t.length > 1);
}

function similarToken(a, b) {
  if (a === b) return true;
  const longest = Math.max(a.length, b.length);
  if (longest < 5) return false;
  return levenshtein(a, b) <= 1;
}

function levenshtein(a, b) {
  if (a === b) return 0;
  if (!a.length) return b.length;
  if (!b.length) return a.length;
  const prev = new Array(b.length + 1);
  const curr = new Array(b.length + 1);
  for (let j = 0; j <= b.length; j += 1) prev[j] = j;
  for (let i = 1; i <= a.length; i += 1) {
    curr[0] = i;
    for (let j = 1; j <= b.length; j += 1) {
      const cost = a[i - 1] === b[j - 1] ? 0 : 1;
      curr[j] = Math.min(curr[j - 1] + 1, prev[j] + 1, prev[j - 1] + cost);
    }
    for (let j = 0; j <= b.length; j += 1) prev[j] = curr[j];
  }
  return prev[b.length];
}

export function scoreNames(ocrName, excelName) {
  const a = normalizeName(ocrName);
  const b = normalizeName(excelName);
  if (!a || !b) return 0;
  if (a === b) return 100;
  if (a.includes(b) || b.includes(a)) {
    const shorter = Math.min(a.length, b.length);
    const longer = Math.max(a.length, b.length);
    return Math.round(88 + (shorter / longer) * 10);
  }

  const ta = tokens(a);
  const tb = tokens(b);
  if (!ta.length || !tb.length) return 0;
  const used = new Set();
  let overlap = 0;
  for (const token of ta) {
    const idx = tb.findIndex((other, i) => !used.has(i) && similarToken(token, other));
    if (idx !== -1) {
      used.add(idx);
      overlap += 1;
    }
  }
  const jaccard = overlap / new Set([...ta, ...tb]).size;
  const prefix = ta.slice(0, 2).join(" ");
  const prefixB = tb.slice(0, 2).join(" ");
  const prefixBonus = prefix && prefix === prefixB ? 12 : 0;
  const maxLen = Math.max(a.length, b.length);
  const lev = 1 - levenshtein(a, b) / maxLen;
  return Math.round(jaccard * 70 + lev * 20 + prefixBonus);
}

export function matchCertificate(ocr, records) {
  if (!ocr.name || !records.length) {
    return { record: null, score: 0 };
  }

  let best = null;
  let bestScore = 0;
  for (const record of records) {
    let score = scoreNames(ocr.name, record.name);
    if (ocr.dateKey && record.dateKey && ocr.dateKey === record.dateKey) {
      score += 8;
    }
    if (ocr.course && record.title) {
      const course = normalizeName(ocr.course);
      const title = normalizeName(record.title);
      if (course.includes("INSPECTOR") && title.includes("INSPECTOR")) score += 4;
      if (course.includes("ERECTOR") && title.includes("ERECTOR")) score += 4;
      if (course.includes("SUPERVISOR") && title.includes("SUPERVISOR")) score += 4;
    }
    if (score > bestScore) {
      bestScore = score;
      best = record;
    }
  }

  if (bestScore < 68) return { record: null, score: bestScore };
  return { record: best, score: Math.min(bestScore, 100) };
}
