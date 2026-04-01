function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Pole Moves')
    .addItem('排序選取的格子', 'sortSelectedCell')
    .addToUi();
}

function parseCompact(str) {
  // 在每個 c+數字 的 c 前面切分，保留 c 在後段
  const raw = str.split(/(?=c\d)/);
  const entries = [];
  let pending = '';
  for (const part of raw) {
    if (/\d/.test(part)) {
      entries.push(pending + part);
      pending = '';
    } else {
      pending += part;
    }
  }
  if (pending) entries.push(pending);
  return entries.filter(e => e.length > 0);
}

function sortCell(cell) {
  const raw = cell.getValue().toString().trim();
  if (!raw) return null;

  // 判斷格式：有逗號 → 格式A，否則 → 格式B（緊湊）
  let entries;
  if (raw.includes(',')) {
    entries = raw.split(',').map(e => e.trim()).filter(e => e.length > 0);
  } else {
    entries = parseCompact(raw);
  }

  const originalCount = entries.length;

  function extractNumber(entry) {
    const match = entry.match(/\d+/);
    return match ? parseInt(match[0], 10) : 0;
  }

  const sorted = entries
    .map((entry, i) => ({ entry, num: extractNumber(entry), i }))
    .sort((a, b) => a.num - b.num || a.i - b.i)
    .map(obj => obj.entry);

  cell.setValue(sorted.join(', '));

  return { originalCount, sortedCount: sorted.length };
}

function sortSelectedCell() {
  const range = SpreadsheetApp.getActiveRange();
  const numRows = range.getNumRows();
  const numCols = range.getNumColumns();

  let processed = 0;
  const warnings = [];

  for (let row = 1; row <= numRows; row++) {
    for (let col = 1; col <= numCols; col++) {
      const cell = range.getCell(row, col);
      const result = sortCell(cell);
      if (!result) continue;

      processed++;
      if (result.originalCount !== result.sortedCount) {
        const addr = cell.getA1Notation();
        warnings.push(`${addr}：輸入 ${result.originalCount} 個，輸出 ${result.sortedCount} 個`);
      }
    }
  }

  if (warnings.length > 0) {
    SpreadsheetApp.getUi().alert(`⚠️ 數量不一致：\n${warnings.join('\n')}`);
  } else {
    SpreadsheetApp.getUi().alert(`✓ 排序完成，共處理 ${processed} 格，數量皆一致。`);
  }
}
