/**
 * 「下一位：Triage」回饋收集
 * 貼到 Google 試算表的「擴充功能 → Apps Script」，部署成網頁應用程式。
 * 每筆回饋會新增一列到「回饋」工作表（沒有的話會自動建立）。
 */
const SHEET_NAME = '回饋';

function save_(p) {
  if (!p || !String(p.text || '').trim()) return false;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) sh.appendRow(['時間', '回饋', '暱稱', '評等', '分數', '稱號', '來源']);
  // 限制長度，並避免開頭是 = + - @ 的內容被當成公式
  const clean = s => {
    s = String(s || '').slice(0, 1000);
    return /^[=+\-@]/.test(s) ? "'" + s : s;
  };
  sh.appendRow([new Date(), clean(p.text), clean(p.nick), clean(p.grade), clean(p.score), clean(p.title), clean(p.src)]);
  return true;
}

// 用 HTML 檔或其他網站打開遊戲時：背景直接送出
function doPost(e) {
  save_(e.parameter);
  return ContentService.createTextOutput('ok');
}

// 在 claude.ai 裡打開遊戲時：開新分頁到這個網址，打開就存好
function doGet(e) {
  const ok = save_(e.parameter);
  const msg = ok ? '已收到你的回饋，謝謝！' : '沒有收到內容。';
  return HtmlService.createHtmlOutput(
    '<meta name="viewport" content="width=device-width,initial-scale=1">' +
    '<div style="font-family:system-ui,sans-serif;padding:48px 20px;text-align:center">' +
    '<h2>' + msg + '</h2><p>可以關閉這個分頁，回到遊戲了。</p></div>'
  ).setTitle('下一位：Triage 回饋');
}
