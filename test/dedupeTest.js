const { strictEqual } = require('assert');

function levenshtein(a, b) {
  const m = a.length, n = b.length;
  const d = Array.from({ length: m + 1 }, () => Array(n + 1).fill(0));
  for (let i = 0; i <= m; i++) d[i][0] = i;
  for (let j = 0; j <= n; j++) d[0][j] = j;
  for (let i = 1; i <= m; i++) for (let j = 1; j <= n; j++) {
    const cost = a[i - 1] === b[j - 1] ? 0 : 1;
    d[i][j] = Math.min(d[i - 1][j] + 1, d[i][j - 1] + 1, d[i - 1][j - 1] + cost);
  }
  return d[m][n];
}
function normalizeText(str) {
  return (str || '').toLowerCase().replace(/\s+/g, ' ').trim();
}
function similarity(a, b) {
  if (!a || !b) return 0;
  a = normalizeText(a);
  b = normalizeText(b);
  const dist = levenshtein(a, b);
  return 1 - dist / Math.max(a.length, b.length);
}
const SIMILARITY_THRESHOLD = 0.7;
function dedupeEvents(events) {
  const seen = {};
  const result = [];
  events.forEach(ev => {
    const dateKey = ev.deadline || '_none_';
    seen[dateKey] = seen[dateKey] || [];
    let merged = false;
    for (const existing of seen[dateKey]) {
      if (similarity(existing.title, ev.title) > SIMILARITY_THRESHOLD) {
        merged = true;
        break;
      }
    }
    if (!merged) {
      seen[dateKey].push(ev);
      result.push(ev);
    }
  });
  return result;
}

const events = [
  { title: 'レポート提出締切', deadline: '2024-07-01' },
  { title: '  レポート  提出  締切  ', deadline: '2024-07-01' },
  { title: 'ゲーム大会', deadline: '2024-07-01' }
];
const deduped = dedupeEvents(events);
console.log('結果件数:', deduped.length);
console.log('タイトル一覧:', deduped.map(e => e.title));
strictEqual(deduped.length, 2);
