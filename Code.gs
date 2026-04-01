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

function sortSelectedCell() {
  const cell = SpreadsheetApp.getActiveRange();
  const raw = cell.getValue().toString().trim();

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

  const sortedCount = sorted.length;
  if (originalCount !== sortedCount) {
    SpreadsheetApp.getUi().alert(
      `⚠️ 數量不一致！\n輸入：${originalCount} 個\n輸出：${sortedCount} 個\n請檢查。`
    );
  } else {
    SpreadsheetApp.getUi().alert(
      `✓ 排序完成\n共 ${sortedCount} 個，數量一致。`
    );
  }
}
