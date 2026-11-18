/**
 * 「下一位：Triage」回饋收集
 * 貼到 Google 試算表的「擴充功能 → Apps Script」，部署成網頁應用程式
 * （執行身分：我；誰可以存取：所有人）。
 * 每筆回饋會新增一列到「回饋」工作表（沒有的話會自動建立）。
 */
const SHEET_NAME = '回饋';
const MAX_PER_MINUTE = 30;   // 防止有人短時間大量灌資料

function save_(p) {
  if (!p || !String(p.text || '').trim()) return 'empty';
  // 每分鐘最多收 MAX_PER_MINUTE 筆
  const cache = CacheService.getScriptCache();
  const slot = 'n_' + Math.floor(Date.now() / 60000);
  const n = Number(cache.get(slot) || 0);
  if (n >= MAX_PER_MINUTE) return 'busy';
  cache.put(slot, String(n + 1), 120);

  const lock = LockService.getScriptLock();   // 多人同時送出時排隊寫入
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
    if (sh.getLastRow() === 0) sh.appendRow(['時間', '回饋', '暱稱', '評等', '分數', '稱號', '來源']);
    // 限制長度，並避免開頭是 = + - @ 的內容被當成公式
    const clean = s => {
      s = String(s || '').slice(0, 1000);
      return /^[=+\-@]/.test(s) ? "'" + s : s;
    };
    sh.appendRow([new Date(), clean(p.text), clean(p.nick), clean(p.grade), clean(p.score), clean(p.title), clean(p.src)]);
  } finally {
    lock.releaseLock();
  }
  return 'ok';
}

// 用 HTML 檔或其他網站打開遊戲時：背景送出，回傳 ok
function doPost(e) {
  return ContentService.createTextOutput(save_(e.parameter));
}

// 在 claude.ai 裡打開遊戲時：開新分頁到這個網址，打開就存好
function doGet(e) {
  const r = save_(e.parameter);
  const msg = r === 'ok' ? '已收到你的回饋，謝謝！' : r === 'busy' ? '現在送出的人太多，請過一分鐘再試。' : '沒有收到內容。';
  return HtmlService.createHtmlOutput(
    '<meta name="viewport" content="width=device-width,initial-scale=1">' +
    '<div style="font-family:system-ui,sans-serif;padding:48px 20px;text-align:center">' +
    '<h2>' + msg + '</h2><p>可以關閉這個分頁，回到遊戲了。</p></div>'
  ).setTitle('下一位：Triage 回饋');
}
