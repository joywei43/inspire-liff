/**
 * INSPIRE 小魏JOY｜LINE 訂房・訂車系統
 * ------------------------------------------------------------
 * 兩個 LINE 官方帳號共用這一份程式（同一個 Webhook 網址）：
 *   ・客人帳號 @inspirekorea：會員登記、訂房、訂車、「我的預約」都在 LINE 彈窗（LIFF）裡完成
 *     彈窗頁面放在 GitHub Pages，透過這個網址的 API 讀寫資料
 *   ・小幫手帳號（例如 TEST）：通知你有新需求；你回覆訂房代號或「已經訂好了」，系統自動通知客人
 * 另有填單網頁，給你和助理直接建立訂房、查看所有紀錄。
 */

const TZ = 'Asia/Taipei';

// 試算表 ID（試算表網址 /d/ 和 /edit 中間那一段）。從試算表「擴充功能 → Apps Script」開的專案可以留空。
const SHEET_ID = '';

const CFG = {
  CHECKIN_TIME: '15:00',
  CHECKOUT_TIME: '11:00',
  TOWERS: ['輝映樓', '海洋樓'],          // 官網名稱：輝映樓 SUN TOWER、海洋樓 OCEAN TOWER（要加翠林樓就加進來）
  // 各樓可訂的房型（名稱照官網）
  ROOMS: {
    '輝映樓': ['豪華大床房', '豪華雙床房', '單間套房特大床', '輝映套房特大床', '輝映套房雙床', '迎仕柏套房特大床', '迎仕柏套房雙床'],
    '海洋樓': ['豪華大床房', '豪華雙床房'],
  },
  MAX_ROOM_PEOPLE: 2,
  MAX_CAR_PAX: 10,
  NOTE_DEPART: '上車地點一律為 VIP LOBBY',
  TRASH_PASSPORT_AFTER_BIND: true,      // 護照照片只用來代辦會員卡，綁定完成後移到雲端硬碟垃圾桶
  NOTE_ARRIVE: '接機工作人員會舉著 INSPIRE 的牌子等候您',
  ACCOUNT_NAME: '韓國INSPIRE 小魏JOY',
};

const SHEETS = { CUST: '客人', ROOM: '訂房', CAR: '訂車', STAFF: '使用者', LOG: '紀錄' };
const HEADERS = {
  '客人': ['LINE userId', 'LINE名稱', '會員ID', '中文姓名', '英文姓名', '狀態', '加入時間', '最後訊息', '備註',
          '登記狀態', '證件類型', '證件檔案', '登記時間', '綁定人', '綁定時間', '證件檔案ID', '通知訊息ID'],
  '訂房': ['編號', '狀態', '訂房代號', '會員ID', '中文姓名', '英文姓名', '樓', '房型', '入住日期', '退房日期', '晚數', '人數',
           '來源', '處理人', '建立時間', '確認時間', 'LINE userId', '通知訊息ID', '備註'],
  '訂車': ['編號', '狀態', '類型', '會員ID', '中文姓名', '英文姓名', '航班', '航班時間', '上車時間', '人數',
           '來源', '處理人', '建立時間', '安排時間', 'LINE userId', '通知訊息ID', '備註'],
  '使用者': ['名稱', '金鑰', '啟用', '建立時間', '小幫手 userId'],
  '紀錄': ['時間', '來源', '內容'],
};

const ST = {
  ROOM_PENDING: '待處理', ROOM_OK: '已確認',
  CAR_PENDING: '待安排', CAR_OK: '已安排',
  CANCELLED: '已取消', FAILED: '通知失敗',
};

const KW = {
  room: /^(我要訂房|訂房|預約訂房)$/,
  car: /^(我要訂車|訂車|接送|預約接送)$/,
  carArrive: /^(接機|我要接機)$/,
  carDepart: /^(送機|我要送機)$/,
  my: /^(我的預約|我的訂房|查詢訂房|訂房查詢|查詢預約|我的接送)$/,
  reg: /^(會員登記|修改資料|會員資料|修改會員資料)$/,
  quit: /^(取消|算了|不用了|離開)$/,
};
const OK_WORDS = /^(已經?訂好了?|訂好了?|好了?|好的|ok|OK|Ok|已安排|安排好了?|done|搞定|完成)[!！。.~～]*$/;
const CANCEL_WORDS = /^(取消|無法安排|沒有了?|沒房|訂不到|不行)$/;

const C = {
  plum: '#2B1B3D', gold: '#C9A45C', goldSoft: '#F6F1E7', goldText: '#8A6F3A',
  ink: '#1F1A24', mute: '#8A8494', grey: '#5B5560', ok: '#2F7D4F', warnBg: '#FFF4D6', warn: '#7A5600',
};

/* ============================================================
 * 試算表選單
 * ============================================================ */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('LINE 訂房後台')
    .addItem('① 初始化工作表', 'setupSheets')
    .addItem('② 設定 LINE 金鑰與網址', 'setupKeys')
    .addItem('③ 新增使用者（填單連結＋小幫手綁定碼）', 'addStaff')
    .addSeparator()
    .addItem('查看所有填單連結', 'listStaffLinks')
    .addItem('重新產生小幫手綁定碼', 'newBindCode')
    .addItem('測試 LINE 連線', 'testLine')
    .addToUi();
}

/* ----- 不管是從試算表選單或從 Apps Script 編輯器執行都能用 ----- */
function ui_() { try { return SpreadsheetApp.getUi(); } catch (_) { return null; } }
function notify_(msg) { const ui = ui_(); if (ui) ui.alert(msg); console.log(msg); return msg; }

/**
 * 從 Apps Script 編輯器直接執行（上方選函式 → 執行）。結果看下方「執行記錄」。
 *   初始化            建立工作表（第一次）
 *   測試連線          檢查兩個 LINE 帳號的金鑰
 *   新增使用者_小魏    產生小魏的填單網址與小幫手綁定碼
 *   新增使用者_助理    產生助理的填單網址與小幫手綁定碼
 *   查看填單連結       列出所有人的填單網址
 * LINE 金鑰、網址等設定放在「專案設定 → 指令碼屬性」：
 *   LINE_TOKEN（客人帳號）、HELPER_TOKEN（小幫手）、WEBAPP_URL、LIFF_ID
 */
function 初始化() { setupSheets(); }
function 測試連線() { testLine(); }
function 新增使用者_小魏() { addStaffNamed_('小魏'); }
function 新增使用者_助理() { addStaffNamed_('助理'); }
function 查看填單連結() { listStaffLinks(); }

function setupSheets() {
  const props = PropertiesService.getScriptProperties();
  let ss = null;
  try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (_) {}
  if (!ss) {
    const id = props.getProperty('SS_ID') || SHEET_ID;
    if (!id) throw new Error('找不到試算表：請在程式最上面的 SHEET_ID 填入試算表網址中的 ID。');
    ss = SpreadsheetApp.openById(id);
  }
  props.setProperty('SS_ID', ss.getId());
  Object.keys(HEADERS).forEach(function (name) {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    const head = HEADERS[name];
    if (sh.getMaxColumns() < head.length) sh.insertColumnsAfter(sh.getMaxColumns(), head.length - sh.getMaxColumns());
    sh.getRange(1, 1, 1, head.length).setValues([head])
      .setFontWeight('bold').setBackground(C.plum).setFontColor('#FFFFFF');
    sh.setFrozenRows(1);
    sh.getRange(1, 1, sh.getMaxRows(), head.length).setNumberFormat('@');
  });
  const blank = ss.getSheetByName('工作表1') || ss.getSheetByName('Sheet1');
  if (blank && ss.getSheets().length > 1 && blank.getLastRow() === 0) ss.deleteSheet(blank);
  notify_('✅ 工作表建立完成：客人、訂房、訂車、使用者、紀錄。\n\n下一步：部署網頁應用程式，然後設定 LINE 金鑰與網址。');
}

function setupKeys() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  const ask = function (title, msg) {
    const r = ui.prompt(title, msg, ui.ButtonSet.OK_CANCEL);
    return r.getSelectedButton() === ui.Button.OK ? r.getResponseText().trim() : null;
  };

  const t1 = ask('② 設定（1/4）客人帳號金鑰',
    '貼上 @inspirekorea 的 Channel access token（長期）。\n留空＝保留原本的設定。');
  if (t1 === null) return;
  if (t1) props.setProperty('LINE_TOKEN', t1);

  const t2 = ask('② 設定（2/4）小幫手帳號金鑰',
    '貼上小幫手帳號（例如 TEST）的 Channel access token（長期）。\n留空＝保留原本的設定。');
  if (t2 === null) return;
  if (t2) props.setProperty('HELPER_TOKEN', t2);

  const url = ask('② 設定（3/4）網頁網址', '貼上「部署 → 網頁應用程式」的網址（結尾是 /exec）。\n留空＝保留原本的設定。');
  if (url === null) return;
  if (url) {
    if (!/^https:\/\/script\.google\.com\/.+\/exec$/.test(url)) {
      ui.alert('網址看起來不對，應該是 https://script.google.com/…/exec，請重新設定一次。');
      return;
    }
    props.setProperty('WEBAPP_URL', url);
  }

  const liff = ask('② 設定（4/4）彈窗 LIFF ID', '貼上 LINE Developers 裡 LIFF 的 LIFF ID（像 2009829088-AbCdEfGh）。\n留空＝保留原本的設定。');
  if (liff === null) return;
  if (liff) {
    if (!/^\d{6,}-[A-Za-z0-9]+$/.test(liff)) { ui.alert('LIFF ID 格式不對，應該像 2009829088-AbCdEfGh。'); return; }
    props.setProperty('LIFF_ID', liff);
    props.setProperty('LOGIN_CHANNEL_ID', liff.split('-')[0]);
  }
  testLine();
}

function addStaff() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('③ 新增使用者', '請輸入名稱（例如：小魏、助理 Amy）', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  addStaffNamed_(r.getResponseText().trim());
}

function addStaffNamed_(name) {
  if (!name) return;
  const exist = readRows_(SHEETS.STAFF).filter(function (s) { return s['名稱'] === name && isOn_(s['啟用']); })[0];
  if (exist) { showStaffDialog_(name, exist['金鑰'], name + '（已存在，重新產生綁定碼）'); return; }
  const key = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '').slice(0, 8);
  appendText_(sheet_(SHEETS.STAFF), [name, key, '是', now_(), '']);
  showStaffDialog_(name, key, '已新增使用者');
}

function newBindCode() {
  const ui = SpreadsheetApp.getUi();
  const rows = readRows_(SHEETS.STAFF).filter(function (s) { return isOn_(s['啟用']); });
  if (!rows.length) { ui.alert('還沒有使用者，請先執行「③ 新增使用者」。'); return; }
  const r = ui.prompt('重新產生小幫手綁定碼', '要產生給誰？請輸入名稱：\n' + rows.map(function (s) { return '・' + s['名稱']; }).join('\n'), ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const s = rows.filter(function (x) { return x['名稱'] === r.getResponseText().trim(); })[0];
  if (!s) { ui.alert('找不到這個名稱。'); return; }
  showStaffDialog_(s['名稱'], s['金鑰'], '小幫手綁定碼');
}

function showStaffDialog_(name, key, title) {
  const code = makeBindCode_(key);
  const link = staffLink_(key);
  const html = HtmlService.createHtmlOutput(
    '<div style="font-family:sans-serif;font-size:14px;line-height:1.6">' +
    '<p style="margin:0 0 6px"><b>' + esc_(name) + '</b> 的填單網頁連結：</p>' +
    (link
      ? '<textarea style="width:100%;height:64px;font-size:12px" onclick="this.select()" readonly>' + esc_(link) + '</textarea>'
      : '<p style="color:#b00">還沒設定網頁網址，設定後到「查看所有填單連結」取得。</p>') +
    '<p style="margin:14px 0 6px"><b>小幫手綁定</b>：用 ' + esc_(name) + ' 的 LINE 加小幫手帳號好友，然後傳：</p>' +
    '<div style="font-size:22px;font-weight:bold;letter-spacing:2px;background:#F6F1E7;padding:10px;border-radius:8px;text-align:center">綁定 ' + code + '</div>' +
    '<p style="color:#666;font-size:12px">綁定碼 24 小時內有效、只能用一次。連結等同鑰匙，請私下傳給對方。<br>要停用某人時，把「使用者」工作表裡的「啟用」改成「否」。</p>' +
    '</div>').setWidth(460).setHeight(380);
  const ui = ui_();
  if (ui) ui.showModalDialog(html, title);
  console.log(title + '｜' + name + '\n填單網址：' + (link || '（還沒設定 WEBAPP_URL）') + '\n小幫手綁定：用 ' + name + ' 的 LINE 傳「綁定 ' + code + '」給小幫手帳號（24 小時內有效）');
}

function makeBindCode_(key) {
  const cache = CacheService.getScriptCache();
  let code;
  for (let i = 0; i < 10; i++) {
    code = String(Math.floor(100000 + Math.random() * 900000));
    if (!cache.get('bind:' + code)) break;
  }
  cache.put('bind:' + code, key, 21600); // Apps Script 快取上限 6 小時
  PropertiesService.getScriptProperties().setProperty('BIND_' + code, JSON.stringify({ key: key, exp: Date.now() + 86400000 }));
  return code;
}

function listStaffLinks() {
  const rows = readRows_(SHEETS.STAFF);
  const items = rows.map(function (r) {
    const on = isOn_(r['啟用']);
    return '<li style="margin-bottom:10px"><b>' + esc_(r['名稱']) + '</b>' + (on ? '' : '（已停用）') +
      (r['小幫手 userId'] ? '　<span style="color:#2F7D4F">✓ 已綁定小幫手</span>' : '　<span style="color:#999">未綁定小幫手</span>') +
      '<br><input style="width:100%;font-size:12px" readonly onclick="this.select()" value="' + esc_(staffLink_(r['金鑰']) || '（尚未設定網址）') + '"></li>';
  }).join('');
  const ui = ui_();
  if (ui) ui.showModalDialog(
    HtmlService.createHtmlOutput('<ul style="font-family:sans-serif;font-size:14px;padding-left:18px">' + (items || '<li>還沒有使用者</li>') + '</ul>')
      .setWidth(500).setHeight(340),
    '所有填單連結');
  rows.forEach(function (r) { console.log(r['名稱'] + (isOn_(r['啟用']) ? '' : '（已停用）') + (r['小幫手 userId'] ? '｜已綁定小幫手' : '｜未綁定小幫手') + '\n' + (staffLink_(r['金鑰']) || '（尚未設定網址）')); });
}

function testLine() {
  const props = PropertiesService.getScriptProperties();
  const lines = [];
  ['c', 'h'].forEach(function (who) {
    const label = who === 'c' ? '客人帳號' : '小幫手帳號';
    try {
      const info = lineFetch_(who, 'get', '/v2/bot/info');
      props.setProperty(who === 'c' ? 'CUSTOMER_BOT_ID' : 'HELPER_BOT_ID', info.userId);
      const q = getQuota_(who, true);
      lines.push('✅ ' + label + '：' + info.displayName + '（' + (info.premiumId ? '@' + info.premiumId : info.basicId) + '）' +
        '　本月已用 ' + q.used + (q.limit ? '/' + q.limit : '') + ' 則');
    } catch (err) {
      lines.push('❌ ' + label + '：' + err.message);
    }
  });
  lines.push('網頁網址：' + (props.getProperty('WEBAPP_URL') || '尚未設定'));
  lines.push('彈窗 LIFF：' + (props.getProperty('LIFF_ID') ? 'https://liff.line.me/' + props.getProperty('LIFF_ID') : '尚未設定'));
  const liff = props.getProperty('LIFF_ID');
  if (liff && !props.getProperty('LOGIN_CHANNEL_ID')) props.setProperty('LOGIN_CHANNEL_ID', liff.split('-')[0]);
  notify_('LINE 連線測試\n\n' + lines.join('\n'));
}

/* ============================================================
 * 網頁入口與 Webhook
 * ============================================================ */

function doGet(e) {
  const key = (e && e.parameter && e.parameter.k) || '';
  const staff = findStaff_(key);
  if (!staff) {
    return HtmlService.createHtmlOutput(
      '<meta name="viewport" content="width=device-width,initial-scale=1">' +
      '<p style="font-family:sans-serif;padding:32px;color:#555">請使用你的專屬填單連結開啟。</p>')
      .setTitle('訂房後台');
  }
  const t = HtmlService.createTemplateFromFile('Admin');
  t.keyJson = safeJson_(key);
  t.staffJson = safeJson_(staff.name);
  return t.evaluate()
    .setTitle('訂房後台｜小魏JOY')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1');
}

function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (body.api) {
      return ContentService.createTextOutput(JSON.stringify(handleApi_(body))).setMimeType(ContentService.MimeType.JSON);
    }
    const role = roleOf_(body.destination);
    (body.events || []).forEach(function (ev) {
      try {
        if (role === 'helper') handleHelper_(ev); else handleCustomer_(ev);
      } catch (err) {
        log_('webhook ' + role, err.message + ' ' + JSON.stringify(ev).slice(0, 300));
      }
    });
  } catch (err) {
    log_('webhook', err.message);
  }
  return ContentService.createTextOutput('{"ok":true}').setMimeType(ContentService.MimeType.JSON);
}

/** 依 Webhook 的 destination 判斷是哪個帳號收到的 */
function roleOf_(dest) {
  const props = PropertiesService.getScriptProperties();
  let helperId = props.getProperty('HELPER_BOT_ID');
  if (!helperId && props.getProperty('HELPER_TOKEN')) {
    try { helperId = lineFetch_('h', 'get', '/v2/bot/info').userId; props.setProperty('HELPER_BOT_ID', helperId); } catch (_) {}
  }
  return dest && helperId && dest === helperId ? 'helper' : 'customer';
}

/* ============================================================
 * 客人帳號：聊天室只回簡短文字＋「打開彈窗」按鈕
 * ============================================================ */

function handleCustomer_(ev) {
  const uid = ev.source && ev.source.type === 'user' ? ev.source.userId : '';
  if (!uid) return;

  if (ev.type === 'unfollow') { setCustomerStatus_(uid, '已封鎖'); return; }
  if (ev.type !== 'follow' && ev.type !== 'message') return;

  const cust = upsertCustomer_(uid);
  if (!cust) return;
  const reply = function (msgs) { replyC_(ev.replyToken, msgs); };
  const reg = isRegistered_(cust);
  const waiting = cust['登記狀態'] === REG.WAIT_BIND || cust['登記狀態'] === REG.APPLYING;
  const needReg = function () {
    return linkText_(waiting ? '小魏正在幫您' + (cust['登記狀態'] === REG.APPLYING ? '辦會員卡' : '綁定會員卡') + '，好了會馬上通知您 🙏'
      : '預約前請先拍會員卡完成登記（只需要一次）👇', [[waiting ? '📋 查看進度' : '📷 拍會員卡登記', waiting ? 'my' : 'reg']]);
  };

  if (ev.type === 'follow') {
    reply([reg
      ? menuText_('歡迎回來，' + cust['中文姓名'] + '！需要訂房、訂車，或查看預約，點下面的按鈕就可以 👇')
      : linkText_('歡迎加入 ' + CFG.ACCOUNT_NAME + '！\n我是 INSPIRE 度假村的台灣貴賓專屬公關小魏 😊\n\n第一步：請拍一張您的 INSPIRE 會員卡，小魏幫您綁定到 LINE。\n還沒有會員卡？上傳護照，小魏幫您辦！',
          [['📷 拍會員卡登記', 'reg']])]);
    return;
  }

  const txt = ev.message && ev.message.type === 'text' ? String(ev.message.text || '').trim() : '';
  if (!txt) return;
  if (KW.reg.test(txt)) { reply([reg ? linkText_('點下方按鈕查看會員資料 👇', [['🪪 會員資料', 'reg']]) : needReg()]); return; }
  if (KW.my.test(txt)) { reply([linkText_('點下方按鈕查看您的訂房和接送 👇', [['📋 我的預約', 'my']])]); return; }
  if (KW.room.test(txt)) { reply([reg ? linkText_('點下方按鈕填寫訂房需求 👇', [['🏨 我要訂房', 'room']]) : needReg()]); return; }
  if (KW.car.test(txt) || KW.carArrive.test(txt) || KW.carDepart.test(txt)) {
    reply([reg ? linkText_('點下方按鈕填寫接送需求 👇', [['🚗 我要訂車', 'car']]) : needReg()]);
    return;
  }
  // 其他訊息：一般聊天，由小魏在 LINE 官方帳號後台手動回覆
}

/** 彈窗網址；page = reg | room | car | my */
function liffUrl_(page, id) {
  const liffId = PropertiesService.getScriptProperties().getProperty('LIFF_ID');
  if (!liffId) return '';
  return 'https://liff.line.me/' + liffId + '?p=' + page + (id ? '&id=' + encodeURIComponent(id) : '');
}

/** 文字訊息＋快速按鈕（打開彈窗） */
function linkText_(text, buttons) {
  const items = buttons.map(function (b) {
    const url = liffUrl_(b[1], b[2]);
    return url ? { type: 'action', action: { type: 'uri', label: b[0].slice(0, 20), uri: url } } : null;
  }).filter(Boolean);
  const msg = { type: 'text', text: text };
  if (items.length) msg.quickReply = { items: items };
  else if (!PropertiesService.getScriptProperties().getProperty('LIFF_ID')) log_('彈窗', '還沒設定 LIFF ID，按鈕沒有顯示');
  return msg;
}

function menuText_(text) {
  return linkText_(text, [['🏨 我要訂房', 'room'], ['🚗 我要訂車', 'car'], ['📋 我的預約', 'my']]);
}

/* ----- 推給客人的簡短通知（取代原本的大卡片） ----- */

function roomNotice_(b, mode) {
  let t;
  if (mode === 'confirmed') {
    t = '✅ 訂房確認\n\n' + b.nameZh + ' ' + b.nameEn + '\n' + b.tower + '・' + b.room +
      '\n入住日期：' + fmtYmd_(b.ci) + ' ' + CFG.CHECKIN_TIME + '\n退房日期：' + fmtYmd_(b.co) + ' ' + CFG.CHECKOUT_TIME +
      '\n訂房代號：' + b.code;
  } else {
    t = '很抱歉，您 ' + fmtYmd_(b.ci) + ' 入住的訂房需求（' + b.id + '）目前無法安排，小魏會再與您聯繫。';
  }
  return linkText_(t, [['📋 查看我的預約', 'my', b.id]]);
}

function carNotice_(c, mode) {
  let t;
  if (mode === 'arranged') {
    t = '✅ ' + c.type + '已安排好\n\n航班：' + c.flight + '\n' + (c.type === '接機' ? '抵達：' : '起飛：') + fmtDT_(c.ftime) +
      (c.type === '送機' ? '\n上車：' + fmtDT_(c.pickup) + '\n📍 ' + CFG.NOTE_DEPART : '\n🪧 ' + CFG.NOTE_ARRIVE) +
      '\n上車人數：' + c.pax + ' 位';
  } else {
    t = '很抱歉，您 ' + c.flight + ' 的' + c.type + '需求（' + c.id + '）目前無法安排，小魏會再與您聯繫。';
  }
  return linkText_(t, [['📋 查看我的預約', 'my', c.id]]);
}

/* ============================================================
 * 會員登記：客人拍會員卡（或上傳護照代辦），小魏核對後綁定會員卡號
 * 照片存在你的 Google 雲端硬碟（不公開），護照照片在綁定後自動移到垃圾桶
 * ============================================================ */

const REG = { NONE: '', WAIT_BIND: '待綁定', APPLYING: '辦卡中', BOUND: '已綁定' };
const DOC = { card: '會員卡', passport: '護照' };

function validateRegistration_(d) {
  const nameZh = String(d.nameZh || '').trim();
  const nameEn = String(d.nameEn || '').trim().replace(/\s+/g, ' ').toUpperCase();
  const docType = d.docType === 'passport' ? 'passport' : d.docType === 'card' ? 'card' : '';
  if (!nameZh || nameZh.length > 20) throw new Error('請輸入中文姓名（20 字以內）。');
  if (!/^[A-Z][A-Z ,.'-]{1,59}$/.test(nameEn)) throw new Error('英文姓名請用英文字母，例如 CHEN CHIH MING。');
  if (!docType) throw new Error('請選擇有沒有會員卡。');
  const m = String(d.image || '').match(/^data:image\/(jpeg|png|webp);base64,([A-Za-z0-9+/=]+)$/);
  if (!m) throw new Error('請上傳' + DOC[docType] + '照片。');
  const bytes = Utilities.base64Decode(m[2]);
  if (bytes.length < 10 * 1024) throw new Error('照片太小或不清楚，請重新拍一張。');
  if (bytes.length > 8 * 1024 * 1024) throw new Error('照片太大了，請重新拍一張。');
  return { nameZh: nameZh, nameEn: nameEn, docType: docType, mime: 'image/' + m[1], bytes: bytes };
}

function submitRegistration_(uid, r) {
  let cust, row, oldFile;
  withLock_(function () {
    const found = findRow_(SHEETS.CUST, 'LINE userId', uid);
    if (!found) throw new Error('找不到您的資料，請重新開啟。');
    if (found.data['登記狀態'] === REG.BOUND) throw new Error('您的會員卡已經綁定了，要修改請直接傳訊息給小魏。');
    row = found.row; cust = found.data; oldFile = found.data['證件檔案ID'];
  });

  const folder = regFolder_();
  const ext = r.mime === 'image/png' ? 'png' : r.mime === 'image/webp' ? 'webp' : 'jpg';
  const name = Utilities.formatDate(new Date(), TZ, 'yyyyMMdd-HHmm') + '_' + r.nameEn.replace(/[^A-Z]+/g, '-') + '_' + DOC[r.docType] + '.' + ext;
  const file = folder.createFile(Utilities.newBlob(r.bytes, r.mime, name));
  file.setDescription('LINE ' + (cust['LINE名稱'] || '') + '｜' + r.nameZh + ' ' + r.nameEn);
  if (oldFile) trashFile_(oldFile); // 重新上傳時，舊照片移到垃圾桶

  const status = r.docType === 'card' ? REG.WAIT_BIND : REG.APPLYING;
  setCells_(SHEETS.CUST, row, {
    '中文姓名': r.nameZh, '英文姓名': r.nameEn, '登記狀態': status, '證件類型': DOC[r.docType],
    '證件檔案ID': file.getId(), '證件檔案': file.getUrl(), '登記時間': now_(), '綁定人': '', '綁定時間': '', '通知訊息ID': '',
  });
  const fresh = findRow_(SHEETS.CUST, 'LINE userId', uid).data;
  notifyStaffReg_(fresh);
  return status;
}

function regFolder_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('REG_FOLDER_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (_) {} }
  const folder = DriveApp.createFolder('INSPIRE 小魏JOY｜會員卡與護照（勿分享）');
  props.setProperty('REG_FOLDER_ID', folder.getId());
  return folder;
}

function trashFile_(fileId) {
  try { DriveApp.getFileById(fileId).setTrashed(true); } catch (err) { log_('刪除照片', fileId + '：' + err.message); }
}

function notifyStaffReg_(c) {
  const staff = readRows_(SHEETS.STAFF).filter(function (s) { return isOn_(s['啟用']) && s['小幫手 userId']; });
  if (!staff.length) { log_('小幫手', '沒有綁定任何使用者，' + c['中文姓名'] + ' 的會員登記沒有通知到人'); return; }
  const msg = staffRegFlex_(c);
  const ids = [];
  staff.forEach(function (s) {
    const res = push_('h', s['小幫手 userId'], [msg]);
    if (res.ok && res.ids) ids.push.apply(ids, res.ids);
    if (!res.ok) log_('小幫手通知失敗', s['名稱'] + '：' + res.error);
  });
  if (ids.length) {
    const found = findRow_(SHEETS.CUST, 'LINE userId', c['LINE userId']);
    if (found) setCells_(SHEETS.CUST, found.row, { '通知訊息ID': ids.join(',') });
  }
}

/** 綁定會員卡號（小幫手、網頁共用） */
function bindMember_(uid, memberId, staffName) {
  memberId = String(memberId || '').replace(/\s/g, '').toUpperCase();
  if (!/^[A-Z0-9-]{3,30}$/.test(memberId)) return { ok: false, msg: '會員卡號只能有英文字母和數字，請再確認一次。' };
  let c, row;
  const err = withLock_(function () {
    const found = findRow_(SHEETS.CUST, 'LINE userId', uid);
    if (!found) return '找不到這位客人。';
    const other = readRows_(SHEETS.CUST).filter(function (x) {
      return x['會員ID'] === memberId && x['登記狀態'] === REG.BOUND && x['LINE userId'] !== uid;
    })[0];
    if (other) return '會員卡號 ' + memberId + ' 已經綁定在 ' + (other['中文姓名'] || other['LINE名稱']) + ' 的 LINE 上了，請確認。';
    if (found.data['登記狀態'] === REG.BOUND && found.data['會員ID'] === memberId) return found.data['中文姓名'] + ' 已經綁定 ' + memberId + ' 了。';
    row = found.row; c = found.data;
    setCells_(SHEETS.CUST, row, { '會員ID': memberId, '登記狀態': REG.BOUND, '綁定人': staffName, '綁定時間': now_() });
    return '';
  });
  if (err) return { ok: false, msg: err };

  // 護照只用來代辦會員卡，綁定後移到垃圾桶
  if (c['證件類型'] === DOC.passport && CFG.TRASH_PASSPORT_AFTER_BIND && c['證件檔案ID']) {
    trashFile_(c['證件檔案ID']);
    setCells_(SHEETS.CUST, row, { '證件檔案ID': '', '證件檔案': '（護照照片已於綁定後刪除）' });
  }
  const applied = c['證件類型'] === DOC.passport;
  const text = (applied ? '🎉 您的 INSPIRE 會員卡辦好了！' : '✅ 會員卡綁定完成！') +
    '\n\n' + c['中文姓名'] + ' ' + c['英文姓名'] + '\n會員卡號：' + memberId +
    '\n\n現在可以直接在 LINE 裡訂房、訂車了 👇';
  const res = push_('c', uid, [menuText_(text)]);
  const who = c['中文姓名'] + '（' + memberId + '）';
  if (!res.ok) return { ok: true, msg: '✅ 已綁定 ' + who + '，但通知客人失敗：' + res.error };
  return { ok: true, msg: '✅ 已綁定 ' + who + '，也通知客人了。' + (applied && CFG.TRASH_PASSPORT_AFTER_BIND ? '\n護照照片已刪除。' : '') };
}

/** 照片不清楚等，請客人重新上傳 */
function rejectReg_(uid, reason, staffName) {
  const found = findRow_(SHEETS.CUST, 'LINE userId', uid);
  if (!found) return { ok: false, msg: '找不到這位客人。' };
  const c = found.data;
  if (c['登記狀態'] === REG.BOUND) return { ok: false, msg: c['中文姓名'] + ' 已經綁定了。' };
  if (c['證件檔案ID']) trashFile_(c['證件檔案ID']);
  setCells_(SHEETS.CUST, found.row, { '登記狀態': REG.NONE, '證件檔案ID': '', '證件檔案': '', '備註': '退件：' + reason + '（' + staffName + '）' });
  const res = push_('c', uid, [linkText_('不好意思，' + (reason || '照片看不太清楚') + '，麻煩再上傳一次 🙏', [['📷 重新上傳', 'reg']])]);
  return { ok: true, msg: '已請 ' + (c['中文姓名'] || c['LINE名稱']) + ' 重新上傳' + (res.ok ? '。' : '，但通知失敗：' + res.error) };
}

function pendingRegs_() {
  return readRows_(SHEETS.CUST).filter(function (c) {
    return c['登記狀態'] === REG.WAIT_BIND || c['登記狀態'] === REG.APPLYING;
  });
}

function staffRegBubble_(c) {
  const applying = c['登記狀態'] === REG.APPLYING;
  const uid = c['LINE userId'];
  const footer = [];
  if (c['證件檔案']) footer.push({ type: 'button', style: 'secondary', height: 'sm',
    action: { type: 'uri', label: '查看' + c['證件類型'] + '照片', uri: c['證件檔案'] } });
  footer.push({ type: 'button', style: 'primary', color: C.plum, height: 'sm',
    action: { type: 'postback', label: applying ? '會員卡辦好了，輸入卡號' : '輸入會員卡號綁定', data: 'a=bind&u=' + uid, displayText: '綁定 ' + c['中文姓名'] + ' 的會員卡號' } });
  footer.push({ type: 'button', style: 'secondary', height: 'sm',
    action: { type: 'postback', label: '照片不清楚，請重傳', data: 'a=reject&u=' + uid } });
  footer.push({ type: 'text', text: '也可以長按這則訊息「回覆」會員卡號', size: 'xxs', color: C.mute, align: 'center', margin: 'sm' });
  return {
    type: 'bubble', size: 'mega',
    header: { type: 'box', layout: 'vertical', paddingAll: '16px', backgroundColor: C.plum, contents: [
      { type: 'text', text: applying ? '📘 代辦會員卡（護照）' : '🪪 新會員登記', color: '#FFFFFF', weight: 'bold', size: 'md' },
    ] },
    body: { type: 'box', layout: 'vertical', paddingAll: '16px', spacing: 'sm', contents: [
      row_('客人', [c['中文姓名'], c['英文姓名']], { bold: true }),
      row_('LINE', c['LINE名稱'] || '-'),
      row_('上傳', c['證件類型'] + '・' + String(c['登記時間']).slice(5, 16)),
      row_('要做的事', applying ? '用護照幫客人辦 INSPIRE 會員卡，辦好後輸入卡號' : '看會員卡照片，輸入卡號綁定', { bold: true }),
    ] },
    footer: { type: 'box', layout: 'vertical', spacing: 'sm', paddingAll: '12px', contents: footer },
  };
}

function staffRegFlex_(c) {
  return { type: 'flex', altText: (c['登記狀態'] === REG.APPLYING ? '📘 代辦會員卡：' : '🪪 新會員登記：') + c['中文姓名'] + ' ' + c['英文姓名'],
    contents: staffRegBubble_(c) };
}

/* ----- 網頁用 ----- */

function apiBind(key, uid, memberId) {
  const staff = auth_(key);
  const r = bindMember_(uid, memberId, staff.name);
  return { ok: r.ok, msg: r.msg, customers: listCustomers_() };
}

function apiReject(key, uid, reason) {
  const staff = auth_(key);
  const r = rejectReg_(uid, String(reason || '照片看不太清楚').slice(0, 60), staff.name);
  return { ok: r.ok, msg: r.msg, customers: listCustomers_() };
}

/* ============================================================
 * 彈窗（LIFF）API：用 LINE 登入的身分驗證，只能看、改自己的資料
 * ============================================================ */

function handleApi_(body) {
  let uid;
  try { uid = verifyIdToken_(body.idToken); }
  catch (err) { return { ok: false, auth: true, error: err.message }; }

  const action = String(body.api || '');
  const d = body.data || {};
  try {
    const cust = upsertCustomer_(uid);
    if (!cust) return { ok: false, error: '請先把「' + CFG.ACCOUNT_NAME + '」加為好友，再打開這個頁面。', notFriend: true };

    if (action === 'me') return meResponse_(uid);
    if (action === 'register') {
      const st = submitRegistration_(uid, validateRegistration_(d));
      return meResponse_(uid, { msg: st === REG.APPLYING ? '已收到護照照片' : '已收到會員卡照片' });
    }
    const fresh = findRow_(SHEETS.CUST, 'LINE userId', uid).data;
    if (!isRegistered_(fresh)) {
      const waiting = fresh['登記狀態'] === REG.WAIT_BIND || fresh['登記狀態'] === REG.APPLYING;
      return { ok: false, needReg: true, error: waiting ? '小魏還在幫您綁定會員卡，綁定好就可以預約了。' : '請先完成會員登記。' };
    }

    if (action === 'room') {
      const b = createRoomRequest_(uid, fresh, validateCustomerRoom_(d), '客人');
      return meResponse_(uid, { created: b.id, msg: '已收到您的訂房需求' });
    }
    if (action === 'car') {
      const c = createCarRequest_(uid, fresh, validateCustomerCar_(d), '客人');
      return meResponse_(uid, { created: c.id, msg: '已收到您的' + c.type + '需求' });
    }
    if (action === 'cancel') {
      const r = customerCancel_(uid, String(d.id || ''));
      return meResponse_(uid, { msg: r });
    }
    return { ok: false, error: '不支援的操作。' };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

function meResponse_(uid, extra) {
  const c = findRow_(SHEETS.CUST, 'LINE userId', uid).data;
  const today = todayYmd_();
  const recent = addDays_(today, -7);
  const rooms = readRows_(SHEETS.ROOM).filter(function (r) {
    return r['LINE userId'] === uid && r['退房日期'] >= (r['狀態'] === ST.CANCELLED ? today : recent);
  }).map(publicRoom_);
  const cars = readRows_(SHEETS.CAR).filter(function (r) {
    return r['LINE userId'] === uid && r['航班時間'].slice(0, 10) >= (r['狀態'] === ST.CANCELLED ? today : recent);
  }).map(publicCar_);
  return Object.assign({
    ok: true,
    profile: {
      memberId: isRegistered_(c) ? c['會員ID'] : '', nameZh: c['中文姓名'], nameEn: c['英文姓名'], registered: isRegistered_(c),
      regStatus: c['登記狀態'] || '', docType: c['證件類型'] || '', regAt: c['登記時間'] || '',
      rejected: !c['登記狀態'] && /^退件：/.test(c['備註'] || '') ? c['備註'].replace(/^退件：/, '').replace(/（.*）$/, '') : '',
    },
    rooms: rooms, cars: cars,
    cfg: {
      towers: CFG.TOWERS, rooms: CFG.ROOMS, maxRoomPeople: CFG.MAX_ROOM_PEOPLE, maxCarPax: CFG.MAX_CAR_PAX,
      checkin: CFG.CHECKIN_TIME, checkout: CFG.CHECKOUT_TIME, noteDepart: CFG.NOTE_DEPART, noteArrive: CFG.NOTE_ARRIVE,
      today: today,
    },
  }, extra || {});
}

/** 給客人看的欄位（不含內部備註、處理人、LINE userId） */
function publicRoom_(r) {
  const st = r['狀態'];
  return {
    kind: 'room', id: r['編號'], status: st === ST.FAILED ? ST.ROOM_OK : st, code: r['訂房代號'],
    memberId: r['會員ID'], nameZh: r['中文姓名'], nameEn: r['英文姓名'], tower: r['樓'], room: r['房型'],
    ci: r['入住日期'], co: r['退房日期'], nights: parseInt(r['晚數'], 10) || 0, people: parseInt(r['人數'], 10) || 1,
    created: r['建立時間'],
  };
}

function publicCar_(r) {
  const st = r['狀態'];
  return {
    kind: 'car', id: r['編號'], status: st === ST.FAILED ? ST.CAR_OK : st, type: r['類型'],
    nameZh: r['中文姓名'], nameEn: r['英文姓名'], flight: r['航班'], ftime: r['航班時間'], pickup: r['上車時間'],
    pax: parseInt(r['人數'], 10) || 1, created: r['建立時間'],
  };
}

function validateCustomerRoom_(d) {
  const r = {
    tower: String(d.tower || ''), room: String(d.room || ''),
    ci: String(d.ci || ''), co: String(d.co || ''), people: parseInt(d.people, 10),
  };
  if (CFG.TOWERS.indexOf(r.tower) < 0) throw new Error('請選擇樓。');
  if ((CFG.ROOMS[r.tower] || []).indexOf(r.room) < 0) throw new Error('請選擇房型。');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(r.ci) || !/^\d{4}-\d{2}-\d{2}$/.test(r.co)) throw new Error('請選擇入住和退房日期。');
  if (r.ci < todayYmd_()) throw new Error('入住日期不能是過去的日期。');
  r.nights = daysBetween_(r.ci, r.co);
  if (r.nights < 1) throw new Error('退房日期要晚於入住日期。');
  if (r.nights > 60) throw new Error('入住超過 60 晚，請確認日期。');
  if (!(r.people >= 1 && r.people <= CFG.MAX_ROOM_PEOPLE)) throw new Error('每間房最多 ' + CFG.MAX_ROOM_PEOPLE + ' 位。');
  return r;
}

function validateCustomerCar_(d) {
  const dt = function (v) { return String(v || '').replace('T', ' ').slice(0, 16); };
  const c = {
    type: String(d.type || ''), flight: String(d.flight || '').replace(/\s/g, '').toUpperCase(),
    ftime: dt(d.ftime), pickup: dt(d.pickup), pax: parseInt(d.pax, 10),
  };
  const isDT = function (v) { return /^\d{4}-\d{2}-\d{2} \d{2}:\d{2}$/.test(v); };
  if (c.type !== '接機' && c.type !== '送機') throw new Error('請選擇接機或送機。');
  if (!/^[A-Z0-9]{2,3}\d{1,4}[A-Z]?$/.test(c.flight)) throw new Error('請輸入正確的航班號碼，例如 BR160。');
  if (!isDT(c.ftime)) throw new Error('請選擇班機' + (c.type === '接機' ? '抵達' : '起飛') + '時間。');
  if (c.ftime.slice(0, 10) < todayYmd_()) throw new Error('班機日期不能是過去的日期。');
  if (c.type === '送機') {
    if (!isDT(c.pickup)) throw new Error('請選擇預計上車時間。');
    if (c.pickup >= c.ftime) throw new Error('上車時間要早於班機起飛時間。');
    if (minutesBetween_(c.pickup, c.ftime) > 24 * 60) throw new Error('上車時間離起飛超過 24 小時，請確認日期。');
  } else {
    c.pickup = '';
  }
  if (!(c.pax >= 1 && c.pax <= CFG.MAX_CAR_PAX)) throw new Error('上車人數請填 1～' + CFG.MAX_CAR_PAX + ' 位。');
  return c;
}

/** 客人自己取消「還在處理中」的需求 */
function customerCancel_(uid, id) {
  const kind = id.charAt(0) === 'R' ? 'room' : id.charAt(0) === 'C' ? 'car' : '';
  if (!kind) throw new Error('找不到這筆預約。');
  const sheetName = kind === 'room' ? SHEETS.ROOM : SHEETS.CAR;
  let item;
  withLock_(function () {
    const found = findRow_(sheetName, '編號', id);
    if (!found || found.data['LINE userId'] !== uid) throw new Error('找不到這筆預約。');
    const pending = found.data['狀態'] === ST.ROOM_PENDING || found.data['狀態'] === ST.CAR_PENDING;
    if (!pending) throw new Error('這筆已經安排好了，要取消請直接傳訊息給小魏。');
    setCells_(sheetName, found.row, { '狀態': ST.CANCELLED, '處理人': '客人', '備註': '客人於 ' + now_() + ' 自行取消' });
    item = kind === 'room' ? rowToRoom_(found.data) : rowToCar_(found.data);
  });
  const desc = kind === 'room'
    ? item.nameZh + '（' + fmtMD_(item.ci) + ' 入住，' + item.tower + item.room + '）'
    : item.nameZh + '（' + item.type + ' ' + item.flight + '）';
  notifyStaffText_('↩️ 客人取消了 ' + id + '\n' + desc + '\n不用再處理了。');
  return '已取消 ' + id;
}

function notifyStaffText_(text) {
  readRows_(SHEETS.STAFF).filter(function (s) { return isOn_(s['啟用']) && s['小幫手 userId']; }).forEach(function (s) {
    const res = push_('h', s['小幫手 userId'], [text_(text)]);
    if (!res.ok) log_('小幫手通知失敗', s['名稱'] + '：' + res.error);
  });
}

/** 驗證彈窗傳來的 LINE ID Token，回傳 userId */
function verifyIdToken_(idToken) {
  if (!idToken || String(idToken).length < 20) throw new Error('請重新登入。');
  const cache = CacheService.getScriptCache();
  const ck = 'idt_' + Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, idToken));
  const hit = cache.get(ck);
  if (hit) return hit;
  const pp = PropertiesService.getScriptProperties();
  const channelId = pp.getProperty('LOGIN_CHANNEL_ID') || String(pp.getProperty('LIFF_ID') || '').split('-')[0];
  if (!channelId) throw new Error('系統還沒設定 LIFF，請聯繫小魏。');
  const res = UrlFetchApp.fetch('https://api.line.me/oauth2/v2.1/verify', {
    method: 'post', payload: { id_token: idToken, client_id: channelId }, muteHttpExceptions: true,
  });
  if (res.getResponseCode() !== 200) throw new Error('登入已過期，請重新開啟。');
  const info = JSON.parse(res.getContentText());
  if (!info.sub) throw new Error('登入資訊不正確，請重新開啟。');
  const ttl = Math.max(60, Math.min(1800, (info.exp || 0) - Math.floor(Date.now() / 1000) - 30));
  cache.put(ck, info.sub, ttl);
  return info.sub;
}

/* ============================================================
 * 建立需求、通知小幫手
 * ============================================================ */

function createRoomRequest_(uid, cust, d, source) {
  const b = {
    uid: uid, memberId: cust['會員ID'], nameZh: cust['中文姓名'], nameEn: cust['英文姓名'],
    tower: d.tower, room: d.room, ci: d.ci, co: d.co, nights: d.nights, people: d.people, code: '',
  };
  withLock_(function () {
    b.id = nextId_('R');
    appendText_(sheet_(SHEETS.ROOM), [b.id, ST.ROOM_PENDING, '', b.memberId, b.nameZh, b.nameEn, b.tower, b.room,
      b.ci, b.co, String(b.nights), String(b.people), source, '', now_(), '', uid, '', '']);
  });
  b.status = ST.ROOM_PENDING;
  notifyStaff_('room', b, cust);
  return b;
}

function createCarRequest_(uid, cust, d, source) {
  const c = {
    uid: uid, memberId: cust['會員ID'], nameZh: cust['中文姓名'], nameEn: cust['英文姓名'],
    type: d.type, flight: d.flight, ftime: d.ftime, pickup: d.type === '送機' ? d.pickup : '', pax: d.pax,
  };
  withLock_(function () {
    c.id = nextId_('C');
    appendText_(sheet_(SHEETS.CAR), [c.id, ST.CAR_PENDING, c.type, c.memberId, c.nameZh, c.nameEn, c.flight, c.ftime,
      c.pickup, String(c.pax), source, '', now_(), '', uid, '', '']);
  });
  c.status = ST.CAR_PENDING;
  notifyStaff_('car', c, cust);
  return c;
}

function notifyStaff_(kind, item, cust) {
  const staff = readRows_(SHEETS.STAFF).filter(function (s) { return isOn_(s['啟用']) && s['小幫手 userId']; });
  if (!staff.length) { log_('小幫手', '沒有綁定任何使用者，' + item.id + ' 沒有通知到人'); return; }
  const msg = kind === 'room' ? staffRoomFlex_(item, cust) : staffCarFlex_(item, cust);
  const ids = [];
  staff.forEach(function (s) {
    const res = push_('h', s['小幫手 userId'], [msg]);
    if (res.ok && res.ids) ids.push.apply(ids, res.ids);
    if (!res.ok) log_('小幫手通知失敗', s['名稱'] + '：' + res.error);
  });
  if (ids.length) {
    const sheetName = kind === 'room' ? SHEETS.ROOM : SHEETS.CAR;
    const found = findRow_(sheetName, '編號', item.id);
    if (found) setCells_(sheetName, found.row, { '通知訊息ID': ids.join(',') });
  }
}

/* ============================================================
 * 小幫手帳號
 * ============================================================ */

function handleHelper_(ev) {
  const uid = ev.source && ev.source.type === 'user' ? ev.source.userId : '';
  if (!uid) return;
  const reply = function (msgs) { replyH_(ev.replyToken, msgs); };
  const staff = staffByHelperUid_(uid);

  if (ev.type === 'follow') {
    reply([text_(staff
      ? '歡迎回來，' + staff.name + '！有新的訂房、訂車需求會在這裡通知你。'
      : '你好！這是小魏的訂房小幫手。\n請傳「綁定 六位數綁定碼」完成綁定（綁定碼在試算表的「LINE 訂房後台」選單產生）。')]);
    return;
  }
  if (ev.type !== 'message' && ev.type !== 'postback') return;

  const txt = ev.type === 'message' && ev.message.type === 'text' ? String(ev.message.text).trim() : '';
  const bind = txt.match(/^綁定\s*(\d{6})$/);
  if (bind) { reply([text_(bindHelper_(uid, bind[1]))]); return; }
  if (!staff) { reply([text_('你還沒綁定。請傳「綁定 六位數綁定碼」。')]); return; }

  // 按鈕
  if (ev.type === 'postback') {
    const pb = parseQS_(ev.postback.data);
    if (pb.a === 'code') {
      const r = findRow_(SHEETS.ROOM, '編號', pb.id);
      if (!r) { reply([text_('找不到 ' + pb.id)]); return; }
      if (r.data['狀態'] !== ST.ROOM_PENDING) { reply([text_(pb.id + ' 已經是「' + r.data['狀態'] + '」' + (r.data['處理人'] ? '（' + r.data['處理人'] + '）' : '') + '，不用再處理。')]); return; }
      saveState_('h', uid, { await: 'code', id: pb.id });
      reply([text_('請輸入 ' + pb.id + '（' + r.data['中文姓名'] + '）的訂房代號：')]);
      return;
    }
    if (pb.a === 'codeok') { reply([text_(confirmRoom_(pb.id, pb.code, staff.name).msg)]); return; }
    if (pb.a === 'carok') { reply([text_(arrangeCar_(pb.id, staff.name).msg)]); return; }
    if (pb.a === 'cancel') {
      reply([qrText_('確定 ' + pb.id + ' 無法安排嗎？會通知客人「小魏會再與您聯繫」。', [
        ['確定，通知客人', 'a=cancelok&id=' + pb.id], ['不要取消', 'a=noop'],
      ])]);
      return;
    }
    if (pb.a === 'cancelok') { reply([text_(cancelItem_(pb.id.charAt(0) === 'R' ? 'room' : 'car', pb.id, staff.name, true).msg)]); return; }
    if (pb.a === 'list') { reply(pendingList_()); return; }
    if (pb.a === 'bind') {
      const f = findRow_(SHEETS.CUST, 'LINE userId', pb.u);
      if (!f) { reply([text_('找不到這位客人。')]); return; }
      if (f.data['登記狀態'] === REG.BOUND) { reply([text_(f.data['中文姓名'] + ' 已經綁定 ' + f.data['會員ID'] + ' 了（' + f.data['綁定人'] + '）。')]); return; }
      saveState_('h', uid, { await: 'bind', u: pb.u });
      reply([text_('請輸入 ' + f.data['中文姓名'] + '（' + f.data['英文姓名'] + '）的會員卡號：')]);
      return;
    }
    if (pb.a === 'reject') {
      reply([qrText_('要請客人重新上傳的原因？', [
        ['照片模糊', 'a=rejectok&u=' + pb.u + '&r=' + encodeURIComponent('照片有點模糊')],
        ['看不到卡號', 'a=rejectok&u=' + pb.u + '&r=' + encodeURIComponent('照片看不到會員卡號')],
        ['不是會員卡/護照', 'a=rejectok&u=' + pb.u + '&r=' + encodeURIComponent('上傳的不是會員卡或護照資料頁')],
        ['不要退回', 'a=noop'],
      ])]);
      return;
    }
    if (pb.a === 'rejectok') { reply([text_(rejectReg_(pb.u, pb.r, staff.name).msg)]); return; }
    if (pb.a === 'noop') { reply([text_('好的，沒有變更。')]); return; }
    return;
  }

  if (!txt) return;

  // 1) 引用（長按回覆）通知訊息
  const quoted = ev.message.quotedMessageId;
  if (quoted) {
    const hit = findByNoticeId_(quoted);
    if (hit && hit.kind === 'reg') { clearState_('h', uid); reply([text_(bindMember_(hit.id, txt, staff.name).msg)]); return; }
    if (hit) { clearState_('h', uid); reply([text_(handleStaffAnswer_(hit.kind, hit.id, txt, staff.name))]); return; }
  }

  // 2) 「R0012 代號」或「C0007 已經訂好了」
  const m = txt.match(/^([RCrc]\d{4,})[\s:：,，]+(.+)$/);
  if (m) { clearState_('h', uid); reply([text_(handleStaffAnswer_(m[1].charAt(0).toUpperCase() === 'R' ? 'room' : 'car', m[1].toUpperCase(), m[2].trim(), staff.name))]); return; }

  // 3) 剛按了「輸入訂房代號」
  const hst = loadState_('h', uid);
  if (hst && hst.await === 'bind') {
    clearState_('h', uid);
    if (CANCEL_WORDS.test(txt)) { reply([text_('好的，沒有變更。')]); return; }
    reply([text_(bindMember_(hst.u, txt, staff.name).msg)]);
    return;
  }
  if (hst && hst.await === 'code') {
    clearState_('h', uid);
    if (CANCEL_WORDS.test(txt)) { reply([text_('好的，沒有變更。')]); return; }
    reply([text_(confirmRoom_(hst.id, txt, staff.name).msg)]);
    return;
  }

  if (/^(待處理|清單|列表|list)$/i.test(txt)) { reply(pendingList_()); return; }

  // 4) 只有一筆待處理時，直接回覆也可以
  const pend = pendingItems_();
  if (OK_WORDS.test(txt) && pend.cars.length === 1) {
    reply([text_(arrangeCar_(pend.cars[0].id, staff.name).msg)]);
    return;
  }
  if (looksLikeCode_(txt) && pend.rooms.length === 1) {
    const b = pend.rooms[0];
    reply([qrText_('要把訂房代號「' + txt + '」傳給 ' + b.id + ' ' + b.nameZh + '（' + fmtMD_(b.ci) + ' 入住）嗎？', [
      ['是，傳給客人', 'a=codeok&id=' + b.id + '&code=' + encodeURIComponent(txt)], ['不是', 'a=noop'],
    ])]);
    return;
  }
  reply(pendingList_(true));
}

function handleStaffAnswer_(kind, id, txt, staffName) {
  if (CANCEL_WORDS.test(txt)) {
    return '如果 ' + id + ' 確定無法安排，請按通知卡上的「無法安排」按鈕，避免誤按。';
  }
  if (kind === 'room') {
    if (OK_WORDS.test(txt)) return '請回覆 ' + id + ' 的訂房代號（例如：' + id + ' INS2611-0502）。';
    return confirmRoom_(id, txt, staffName).msg;
  }
  return arrangeCar_(id, staffName).msg;
}

function pendingList_(withHelp) {
  const p = pendingItems_();
  const regs = pendingRegs_();
  const msgs = [];
  if (withHelp) {
    msgs.push(text_('使用方式：\n・會員登記：按「輸入會員卡號綁定」，或長按通知回覆卡號\n・訂房：按通知上的「輸入訂房代號」，或回覆「R0012 訂房代號」\n・訂車：按「已訂好車」，或長按通知回覆「已經訂好了」\n・傳「待處理」查看所有待處理'));
  }
  if (!p.rooms.length && !p.cars.length && !regs.length) { msgs.push(text_('目前沒有待處理的事項 👍')); return msgs; }
  const bubbles = regs.map(staffRegBubble_)
    .concat(p.rooms.map(function (b) { return staffRoomBubble_(b); }))
    .concat(p.cars.map(function (c) { return staffCarBubble_(c); })).slice(0, 12);
  msgs.push({ type: 'flex', altText: '待處理：會員 ' + regs.length + '、訂房 ' + p.rooms.length + '、訂車 ' + p.cars.length,
    contents: bubbles.length === 1 ? bubbles[0] : { type: 'carousel', contents: bubbles } });
  return msgs;
}

function pendingItems_() {
  return {
    rooms: readRows_(SHEETS.ROOM).filter(function (r) { return r['狀態'] === ST.ROOM_PENDING; }).map(rowToRoom_),
    cars: readRows_(SHEETS.CAR).filter(function (r) { return r['狀態'] === ST.CAR_PENDING; }).map(rowToCar_),
  };
}

function bindHelper_(uid, code) {
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty('BIND_' + code);
  if (!raw) return '綁定碼不正確或已使用過，請重新產生一組。';
  const info = JSON.parse(raw);
  props.deleteProperty('BIND_' + code);
  if (Date.now() > info.exp) return '綁定碼已過期，請重新產生一組。';
  const found = findRow_(SHEETS.STAFF, '金鑰', info.key);
  if (!found || !isOn_(found.data['啟用'])) return '這個使用者已停用。';
  setCells_(SHEETS.STAFF, found.row, { '小幫手 userId': uid });
  return '✅ 綁定完成！' + found.data['名稱'] + '，之後有新的訂房、訂車需求都會在這裡通知你。\n傳「待處理」可以隨時查看。';
}

/* ============================================================
 * 處理需求（小幫手與網頁共用）
 * ============================================================ */

function confirmRoom_(id, code, staffName) {
  code = String(code || '').trim();
  if (!code) return { ok: false, msg: '訂房代號不能是空白。' };
  if (code.length > 40 || /\n/.test(code)) return { ok: false, msg: '訂房代號太長了，請確認。' };
  let b, row;
  const err = withLock_(function () {
    const found = findRow_(SHEETS.ROOM, '編號', id);
    if (!found) return '找不到 ' + id + '。';
    if (found.data['狀態'] === ST.CANCELLED) return id + ' 已經取消了。';
    if (found.data['狀態'] === ST.ROOM_OK) return id + ' 已經由 ' + found.data['處理人'] + ' 確認過了（代號 ' + found.data['訂房代號'] + '）。';
    const dup = readRows_(SHEETS.ROOM).some(function (r) { return r['訂房代號'] === code && r['狀態'] !== ST.CANCELLED && r['編號'] !== id; });
    if (dup) return '訂房代號「' + code + '」已經用在別筆訂房了，請確認。';
    setCells_(SHEETS.ROOM, found.row, { '訂房代號': code, '狀態': ST.ROOM_OK, '處理人': staffName, '確認時間': now_(), '備註': '' });
    row = found.row;
    b = rowToRoom_(Object.assign({}, found.data, { '訂房代號': code }));
    return '';
  });
  if (err) return { ok: false, msg: err };
  const res = push_('c', b.uid, [roomNotice_(b, 'confirmed')]);
  if (!res.ok) {
    setCells_(SHEETS.ROOM, row, { '狀態': ST.FAILED, '備註': res.error });
    return { ok: false, msg: '⚠️ 已存下代號，但傳給客人失敗：' + res.error + '\n可到填單網頁按「重試」。' };
  }
  return { ok: true, msg: '✅ 已把訂房代號 ' + code + ' 傳給 ' + b.nameZh + '（' + id + '，' + fmtMD_(b.ci) + ' 入住）。' };
}

function arrangeCar_(id, staffName) {
  let c, row;
  const err = withLock_(function () {
    const found = findRow_(SHEETS.CAR, '編號', id);
    if (!found) return '找不到 ' + id + '。';
    if (found.data['狀態'] === ST.CANCELLED) return id + ' 已經取消了。';
    if (found.data['狀態'] === ST.CAR_OK) return id + ' 已經由 ' + found.data['處理人'] + ' 安排好了。';
    setCells_(SHEETS.CAR, found.row, { '狀態': ST.CAR_OK, '處理人': staffName, '安排時間': now_(), '備註': '' });
    row = found.row;
    c = rowToCar_(found.data);
    return '';
  });
  if (err) return { ok: false, msg: err };
  const res = push_('c', c.uid, [carNotice_(c, 'arranged')]);
  if (!res.ok) {
    setCells_(SHEETS.CAR, row, { '狀態': ST.FAILED, '備註': res.error });
    return { ok: false, msg: '⚠️ 已標記安排好，但通知客人失敗：' + res.error };
  }
  return { ok: true, msg: '✅ 已通知 ' + c.nameZh + '：' + c.type + '車子安排好了（' + id + '，' + c.flight + '）。' };
}

function cancelItem_(kind, id, staffName, notify) {
  const sheetName = kind === 'room' ? SHEETS.ROOM : SHEETS.CAR;
  const found = findRow_(sheetName, '編號', id);
  if (!found) return { ok: false, msg: '找不到 ' + id + '。' };
  if (found.data['狀態'] === ST.CANCELLED) return { ok: false, msg: id + ' 已經取消了。' };
  setCells_(sheetName, found.row, { '狀態': ST.CANCELLED, '處理人': staffName, '備註': '由 ' + staffName + ' 於 ' + now_() + ' 取消' });
  if (!notify) return { ok: true, msg: '已取消 ' + id + '（沒有通知客人）。' };
  const item = kind === 'room' ? rowToRoom_(found.data) : rowToCar_(found.data);
  const msg = kind === 'room' ? roomNotice_(item, 'cancelled') : carNotice_(item, 'cancelled');
  const res = push_('c', item.uid, [msg]);
  if (!res.ok) return { ok: false, msg: '已取消 ' + id + '，但通知客人失敗：' + res.error };
  return { ok: true, msg: '已取消 ' + id + '，並通知了 ' + item.nameZh + '。' };
}

function resendItem_(kind, id) {
  const sheetName = kind === 'room' ? SHEETS.ROOM : SHEETS.CAR;
  const found = findRow_(sheetName, '編號', id);
  if (!found) return { ok: false, msg: '找不到 ' + id + '。' };
  const r = found.data;
  let msg, newStatus;
  if (kind === 'room') {
    if (!r['訂房代號']) return { ok: false, msg: '這筆還沒有訂房代號。' };
    msg = roomNotice_(rowToRoom_(r), 'confirmed'); newStatus = ST.ROOM_OK;
  } else {
    if (r['狀態'] === ST.CAR_PENDING) return { ok: false, msg: '這筆還沒安排好車。' };
    msg = carNotice_(rowToCar_(r), 'arranged'); newStatus = ST.CAR_OK;
  }
  if (r['狀態'] === ST.CANCELLED) return { ok: false, msg: '這筆已取消。' };
  const res = push_('c', r['LINE userId'], [msg]);
  setCells_(sheetName, found.row, res.ok ? { '狀態': newStatus, '備註': '' } : { '狀態': ST.FAILED, '備註': res.error });
  return { ok: res.ok, msg: res.ok ? '已重新發送 ' + id : res.error };
}

/* ============================================================
 * 填單網頁 API（每次都驗證金鑰）
 * ============================================================ */

function apiInit(key) {
  const staff = auth_(key);
  return {
    staff: staff.name,
    customers: listCustomers_(),
    rooms: listRows_(SHEETS.ROOM, rowToRoom_, 80),
    cars: listRows_(SHEETS.CAR, rowToCar_, 80),
    quota: safeQuota_(),
    cfg: CFG,
  };
}
function apiRefresh(key) { return apiInit(key); }

/** 網頁直接建立訂房（已有訂房代號，送出即通知客人） */
function apiSubmitRoom(key, d) {
  const staff = auth_(key);
  const b = validateStaffRoom_(d);
  const found = findRow_(SHEETS.CUST, 'LINE userId', b.uid);
  if (!found) throw new Error('找不到這位客人，請重新整理名單。');
  if (found.data['狀態'] === '已封鎖') throw new Error('這位客人已封鎖官方帳號，訊息無法送達。');

  withLock_(function () {
    const dup = readRows_(SHEETS.ROOM).some(function (r) { return r['訂房代號'] === b.code && r['狀態'] !== ST.CANCELLED; });
    if (dup) throw new Error('訂房代號「' + b.code + '」已經存在。');
    b.id = nextId_('R');
    appendText_(sheet_(SHEETS.ROOM), [b.id, ST.ROOM_PENDING, '', b.memberId, b.nameZh, b.nameEn, b.tower, b.room,
      b.ci, b.co, String(b.nights), String(b.people), '後台：' + staff.name, '', now_(), '', b.uid, '', '']);
    // 客人還沒登記的資料，順便補上
    const fill = {};
    if (!found.data['會員ID']) fill['會員ID'] = b.memberId;
    if (!found.data['中文姓名']) fill['中文姓名'] = b.nameZh;
    if (!found.data['英文姓名']) fill['英文姓名'] = b.nameEn;
    if (Object.keys(fill).length) setCells_(SHEETS.CUST, found.row, fill);
  });
  const r = confirmRoom_(b.id, b.code, staff.name);
  return { ok: r.ok, msg: r.msg, id: b.id, rooms: listRows_(SHEETS.ROOM, rowToRoom_, 80), quota: safeQuota_() };
}

function apiConfirmRoom(key, id, code) {
  const staff = auth_(key);
  const r = confirmRoom_(id, code, staff.name);
  return { ok: r.ok, msg: r.msg, rooms: listRows_(SHEETS.ROOM, rowToRoom_, 80), quota: safeQuota_() };
}

function apiArrangeCar(key, id) {
  const staff = auth_(key);
  const r = arrangeCar_(id, staff.name);
  return { ok: r.ok, msg: r.msg, cars: listRows_(SHEETS.CAR, rowToCar_, 80), quota: safeQuota_() };
}

function apiCancel(key, kind, id, notify) {
  const staff = auth_(key);
  const r = cancelItem_(kind, id, staff.name, !!notify);
  return { ok: r.ok, msg: r.msg, rooms: listRows_(SHEETS.ROOM, rowToRoom_, 80), cars: listRows_(SHEETS.CAR, rowToCar_, 80), quota: safeQuota_() };
}

function apiResend(key, kind, id) {
  auth_(key);
  const r = resendItem_(kind, id);
  return { ok: r.ok, msg: r.msg, rooms: listRows_(SHEETS.ROOM, rowToRoom_, 80), cars: listRows_(SHEETS.CAR, rowToCar_, 80), quota: safeQuota_() };
}

function validateStaffRoom_(d) {
  d = d || {};
  const s = function (v) { return String(v == null ? '' : v).trim(); };
  const b = {
    uid: s(d.uid), code: s(d.code), memberId: s(d.memberId).toUpperCase(), nameZh: s(d.nameZh),
    nameEn: s(d.nameEn).replace(/\s+/g, ' ').toUpperCase(), tower: s(d.tower), room: s(d.room),
    ci: s(d.ci), co: s(d.co), people: parseInt(d.people, 10),
  };
  if (!b.uid) throw new Error('請選擇客人。');
  if (!b.memberId) throw new Error('請填寫會員ID。');
  if (!b.nameZh) throw new Error('請填寫中文姓名。');
  if (!/^[A-Z][A-Z ,.'-]{1,59}$/.test(b.nameEn)) throw new Error('請填寫英文姓名（英文字母）。');
  if (CFG.TOWERS.indexOf(b.tower) < 0) throw new Error('請選擇樓。');
  if ((CFG.ROOMS[b.tower] || []).indexOf(b.room) < 0) throw new Error('請選擇房型。');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(b.ci) || !/^\d{4}-\d{2}-\d{2}$/.test(b.co)) throw new Error('請填寫入住與退房日期。');
  b.nights = daysBetween_(b.ci, b.co);
  if (b.nights < 1) throw new Error('退房日期必須晚於入住日期。');
  if (b.nights > 60) throw new Error('入住超過 60 晚，請確認日期。');
  if (!(b.people >= 1 && b.people <= CFG.MAX_ROOM_PEOPLE)) throw new Error('人數最多 ' + CFG.MAX_ROOM_PEOPLE + ' 位。');
  if (!b.code) throw new Error('請填寫訂房代號。');
  if (b.code.length > 40) throw new Error('訂房代號太長了。');
  return b;
}

/* ============================================================
 * 客人資料
 * ============================================================ */

function upsertCustomer_(uid) {
  return withLock_(function () {
    const found = findRow_(SHEETS.CUST, 'LINE userId', uid);
    if (found) {
      const upd = { '最後訊息': now_() };
      if (found.data['狀態'] === '已封鎖') upd['狀態'] = '好友';
      setCells_(SHEETS.CUST, found.row, upd);
      return found.data;
    }
    // 向 LINE 確認這是真的好友（也擋掉偽造的 Webhook）
    let profile;
    try { profile = lineFetch_('c', 'get', '/v2/bot/profile/' + encodeURIComponent(uid)); }
    catch (err) { log_('客人資料', '無法取得 ' + uid + '：' + err.message); return null; }
    const row = [uid, profile.displayName || '', '', '', '', '好友', now_(), now_(), ''];
    appendText_(sheet_(SHEETS.CUST), row);
    const o = {};
    HEADERS['客人'].forEach(function (h, i) { o[h] = row[i]; });
    return o;
  });
}

function setCustomerStatus_(uid, status) {
  const found = findRow_(SHEETS.CUST, 'LINE userId', uid);
  if (found) setCells_(SHEETS.CUST, found.row, { '狀態': status });
}

function isRegistered_(c) {
  return !!(c && c['登記狀態'] === REG.BOUND && c['會員ID']);
}

function listCustomers_() {
  return readRows_(SHEETS.CUST).map(function (r) {
    return {
      uid: r['LINE userId'], lineName: r['LINE名稱'], memberId: r['會員ID'], nameZh: r['中文姓名'], nameEn: r['英文姓名'],
      status: r['狀態'] || '好友', last: r['最後訊息'] || r['加入時間'] || '',
      regStatus: r['登記狀態'] || '', docType: r['證件類型'] || '', docUrl: r['證件檔案'] || '', regAt: r['登記時間'] || '',
    };
  }).filter(function (c) { return c.uid; })
    .sort(function (a, b) { return a.last < b.last ? 1 : -1; });
}

function rowToRoom_(r) {
  return {
    id: r['編號'], status: r['狀態'], code: r['訂房代號'], memberId: r['會員ID'], nameZh: r['中文姓名'], nameEn: r['英文姓名'],
    tower: r['樓'], room: r['房型'], ci: r['入住日期'], co: r['退房日期'],
    nights: parseInt(r['晚數'], 10) || daysBetween_(r['入住日期'], r['退房日期']), people: parseInt(r['人數'], 10) || 1,
    source: r['來源'], staff: r['處理人'], created: r['建立時間'], done: r['確認時間'], uid: r['LINE userId'], note: r['備註'],
  };
}

function rowToCar_(r) {
  return {
    id: r['編號'], status: r['狀態'], type: r['類型'], memberId: r['會員ID'], nameZh: r['中文姓名'], nameEn: r['英文姓名'],
    flight: r['航班'], ftime: r['航班時間'], pickup: r['上車時間'], pax: parseInt(r['人數'], 10) || 1,
    source: r['來源'], staff: r['處理人'], created: r['建立時間'], done: r['安排時間'], uid: r['LINE userId'], note: r['備註'],
  };
}

function listRows_(name, mapper, limit) {
  const pendingFirst = { '待處理': 0, '待安排': 0, '通知失敗': 1 };
  return readRows_(name).map(mapper).reverse().slice(0, limit)
    .sort(function (a, b) { return (pendingFirst[a.status] !== undefined ? pendingFirst[a.status] : 2) - (pendingFirst[b.status] !== undefined ? pendingFirst[b.status] : 2); });
}

function findByNoticeId_(msgId) {
  const regs = readRows_(SHEETS.CUST);
  for (let k = regs.length - 1; k >= 0; k--) {
    if (regs[k]['通知訊息ID'] && regs[k]['通知訊息ID'].split(',').indexOf(String(msgId)) >= 0) {
      return { kind: 'reg', id: regs[k]['LINE userId'] };
    }
  }
  const sheets = [[SHEETS.ROOM, 'room'], [SHEETS.CAR, 'car']];
  for (let i = 0; i < sheets.length; i++) {
    const rows = readRows_(sheets[i][0]);
    for (let j = rows.length - 1; j >= 0; j--) {
      if (rows[j]['通知訊息ID'] && rows[j]['通知訊息ID'].split(',').indexOf(String(msgId)) >= 0) {
        return { kind: sheets[i][1], id: rows[j]['編號'] };
      }
    }
  }
  return null;
}

function staffByHelperUid_(uid) {
  const r = readRows_(SHEETS.STAFF).filter(function (s) { return s['小幫手 userId'] === uid && isOn_(s['啟用']); })[0];
  return r ? { name: r['名稱'] } : null;
}

function findStaff_(key) {
  if (!key || String(key).length < 20) return null;
  const r = readRows_(SHEETS.STAFF).filter(function (x) { return x['金鑰'] === key && isOn_(x['啟用']); })[0];
  return r ? { name: r['名稱'] } : null;
}

function auth_(key) {
  const s = findStaff_(key);
  if (!s) throw new Error('連結已失效或已停用，請向管理者索取新連結。');
  return s;
}

/* ============================================================
 * 小幫手通知卡（Flex Message）
 * ============================================================ */

function row_(label, value, opt) {
  opt = opt || {};
  const val = Array.isArray(value)
    ? { type: 'box', layout: 'vertical', flex: 7, contents: value.filter(String).map(function (v, i) {
        return { type: 'text', text: String(v), size: i ? 'xs' : 'sm', color: i ? C.mute : C.ink, wrap: true, weight: !i && opt.bold ? 'bold' : 'regular' };
      }) }
    : { type: 'text', text: String(value || '-'), size: 'sm', color: C.ink, flex: 7, wrap: true, weight: opt.bold ? 'bold' : 'regular' };
  return { type: 'box', layout: 'horizontal', spacing: 'md', contents: [
    { type: 'text', text: label, size: 'sm', color: C.mute, flex: 3 }, val,
  ] };
}

/* ----- 小幫手通知卡 ----- */

function staffRoomBubble_(b, cust) {
  return {
    type: 'bubble', size: 'mega',
    header: { type: 'box', layout: 'vertical', paddingAll: '16px', backgroundColor: C.plum, contents: [
      { type: 'text', text: '🏨 新訂房需求　' + b.id, color: '#FFFFFF', weight: 'bold', size: 'md' },
    ] },
    body: { type: 'box', layout: 'vertical', paddingAll: '16px', spacing: 'sm', contents: [
      row_('客人', [b.nameZh, b.nameEn], { bold: true }),
      row_('會員ID', b.memberId),
      cust && cust['LINE名稱'] ? row_('LINE', cust['LINE名稱']) : null,
      row_('樓／房型', b.tower + '・' + b.room),
      row_('入住', fmtYmd_(b.ci) + ' ' + CFG.CHECKIN_TIME, { bold: true }),
      row_('退房', fmtYmd_(b.co) + ' ' + CFG.CHECKOUT_TIME + '（' + b.nights + ' 晚）'),
      row_('人數', b.people + ' 位'),
    ].filter(Boolean) },
    footer: { type: 'box', layout: 'vertical', spacing: 'sm', paddingAll: '12px', contents: [
      { type: 'button', style: 'primary', color: C.plum, height: 'sm',
        action: { type: 'postback', label: '輸入訂房代號', data: 'a=code&id=' + b.id, displayText: b.id + ' 輸入訂房代號' } },
      { type: 'button', style: 'secondary', height: 'sm',
        action: { type: 'postback', label: '無法安排', data: 'a=cancel&id=' + b.id } },
      { type: 'text', text: '也可以長按這則訊息「回覆」訂房代號', size: 'xxs', color: C.mute, align: 'center', margin: 'sm' },
    ] },
  };
}

function staffCarBubble_(c, cust) {
  const body = [
    row_('客人', [c.nameZh, c.nameEn], { bold: true }),
    row_('會員ID', c.memberId),
    cust && cust['LINE名稱'] ? row_('LINE', cust['LINE名稱']) : null,
    row_('航班', c.flight, { bold: true }),
    row_(c.type === '接機' ? '抵達' : '起飛', fmtDT_(c.ftime)),
    c.type === '送機' ? row_('上車', fmtDT_(c.pickup) + '・VIP LOBBY', { bold: true }) : null,
    row_('人數', c.pax + ' 位'),
  ].filter(Boolean);
  return {
    type: 'bubble', size: 'mega',
    header: { type: 'box', layout: 'vertical', paddingAll: '16px', backgroundColor: C.plum, contents: [
      { type: 'text', text: '🚗 新訂車・' + c.type + '　' + c.id, color: '#FFFFFF', weight: 'bold', size: 'md' },
    ] },
    body: { type: 'box', layout: 'vertical', paddingAll: '16px', spacing: 'sm', contents: body },
    footer: { type: 'box', layout: 'vertical', spacing: 'sm', paddingAll: '12px', contents: [
      { type: 'button', style: 'primary', color: C.plum, height: 'sm',
        action: { type: 'postback', label: '已訂好車，通知客人', data: 'a=carok&id=' + c.id, displayText: c.id + ' 已經訂好了' } },
      { type: 'button', style: 'secondary', height: 'sm',
        action: { type: 'postback', label: '無法安排', data: 'a=cancel&id=' + c.id } },
      { type: 'text', text: '也可以長按這則訊息回覆「已經訂好了」', size: 'xxs', color: C.mute, align: 'center', margin: 'sm' },
    ] },
  };
}

function staffRoomFlex_(b, cust) {
  return { type: 'flex', altText: '🏨 新訂房需求 ' + b.id + '｜' + b.nameZh + ' ' + fmtMD_(b.ci) + ' 入住', contents: staffRoomBubble_(b, cust) };
}
function staffCarFlex_(c, cust) {
  return { type: 'flex', altText: '🚗 新訂車 ' + c.id + '｜' + c.type + ' ' + c.flight + ' ' + c.nameZh, contents: staffCarBubble_(c, cust) };
}

function text_(t) { return { type: 'text', text: t }; }

function qrText_(t, buttons) {
  return { type: 'text', text: t, quickReply: { items: buttons.map(function (b) {
    return { type: 'action', action: { type: 'postback', label: b[0], data: b[1], displayText: b[0] } };
  }) } };
}

function stripUndef_(o) { return JSON.parse(JSON.stringify(o)); }

/* ============================================================
 * LINE API
 * ============================================================ */

function token_(who) {
  const name = who === 'h' ? 'HELPER_TOKEN' : 'LINE_TOKEN';
  const t = PropertiesService.getScriptProperties().getProperty(name);
  if (!t) throw new Error('還沒設定' + (who === 'h' ? '小幫手' : '客人') + '帳號的 LINE 金鑰，請執行「② 設定 LINE 金鑰與網址」。');
  return t;
}

function lineFetch_(who, method, path, payload, extraHeaders) {
  const opt = {
    method: method,
    headers: Object.assign({ Authorization: 'Bearer ' + token_(who) }, extraHeaders || {}),
    muteHttpExceptions: true,
  };
  if (payload) { opt.contentType = 'application/json'; opt.payload = JSON.stringify(payload); }
  const res = UrlFetchApp.fetch('https://api.line.me' + path, opt);
  const code = res.getResponseCode();
  const text = res.getContentText();
  if (code >= 200 && code < 300) return text ? JSON.parse(text) : {};
  const err = new Error(friendlyLineError_(code, text));
  err.httpCode = code;
  throw err;
}

function replyC_(token, messages) { reply_('c', token, messages); }
function replyH_(token, messages) { reply_('h', token, messages); }

function reply_(who, token, messages) {
  if (!token || !messages || !messages.length) return;
  try {
    lineFetch_(who, 'post', '/v2/bot/message/reply', { replyToken: token, messages: messages.slice(0, 5) });
  } catch (err) {
    log_('回覆失敗', err.message + ' ' + JSON.stringify(messages).slice(0, 300));
  }
}

function push_(who, uid, messages) {
  try {
    const r = lineFetch_(who, 'post', '/v2/bot/message/push', { to: uid, messages: messages },
      { 'X-Line-Retry-Key': Utilities.getUuid() });
    CacheService.getScriptCache().remove('quota_' + who);
    return { ok: true, ids: (r.sentMessages || []).map(function (m) { return String(m.id); }) };
  } catch (err) {
    if (err.httpCode === 409) return { ok: true, ids: [] };
    log_('推播失敗', err.message);
    return { ok: false, error: err.message };
  }
}

function friendlyLineError_(code, text) {
  let detail = '';
  try {
    const j = JSON.parse(text);
    detail = j.message || '';
    if (j.details && j.details.length) detail += '：' + j.details.map(function (d) { return d.property + ' ' + d.message; }).join('；');
  } catch (_) { detail = String(text).slice(0, 200); }
  if (code === 401) return 'LINE 金鑰錯誤或已失效（401），請重新設定 Channel access token。';
  if (code === 403) return 'LINE 拒絕這個操作（403）：' + detail;
  if (code === 429) return '本月訊息額度已用完，或傳送太頻繁（429）。';
  if (code === 404) return '找不到這位客人（404），可能已封鎖或不是好友。';
  if (code === 400 && /not found|Failed to send/i.test(detail)) return '訊息無法送達（400），客人可能已封鎖官方帳號。';
  return 'LINE 回應錯誤 ' + code + '：' + detail;
}

function getQuota_(who, fresh) {
  const cache = CacheService.getScriptCache();
  const ck = 'quota_' + who;
  if (!fresh) { const hit = cache.get(ck); if (hit) return JSON.parse(hit); }
  const q = lineFetch_(who, 'get', '/v2/bot/message/quota');
  const used = lineFetch_(who, 'get', '/v2/bot/message/quota/consumption');
  const out = { limit: q.type === 'limited' ? q.value : null, used: used.totalUsage || 0 };
  cache.put(ck, JSON.stringify(out), 300);
  return out;
}

function safeQuota_() {
  try { return getQuota_('c', false); } catch (err) { return { limit: null, used: null, error: err.message }; }
}

/* ============================================================
 * 對話狀態（暫存 30 分鐘）
 * ============================================================ */

function loadState_(who, uid) {
  const raw = CacheService.getScriptCache().get('st_' + who + '_' + uid);
  return raw ? JSON.parse(raw) : null;
}
function saveState_(who, uid, st) {
  CacheService.getScriptCache().put('st_' + who + '_' + uid, JSON.stringify(st), 1800);
}
function clearState_(who, uid) {
  CacheService.getScriptCache().remove('st_' + who + '_' + uid);
}

/* ============================================================
 * 小工具
 * ============================================================ */

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function nextId_(prefix) {
  const props = PropertiesService.getScriptProperties();
  const n = (parseInt(props.getProperty('SEQ_' + prefix), 10) || 0) + 1;
  props.setProperty('SEQ_' + prefix, String(n));
  return prefix + ('000' + n).slice(-Math.max(4, String(n).length));
}

function ss_() {
  const id = PropertiesService.getScriptProperties().getProperty('SS_ID');
  if (id || SHEET_ID) return SpreadsheetApp.openById(id || SHEET_ID);
  return SpreadsheetApp.getActiveSpreadsheet();
}

function sheet_(name) {
  const sh = ss_().getSheetByName(name);
  if (!sh) throw new Error('找不到工作表「' + name + '」，請先執行「① 初始化工作表」。');
  return sh;
}

function readRows_(name) {
  const sh = sheet_(name);
  const head = HEADERS[name];
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, head.length).getDisplayValues().map(function (vals) {
    const o = {};
    head.forEach(function (h, i) { o[h] = String(vals[i]).trim(); });
    return o;
  });
}

function findRow_(name, col, value) {
  if (!value) return null;
  const rows = readRows_(name);
  for (let i = rows.length - 1; i >= 0; i--) {
    if (rows[i][col] === String(value)) return { row: i + 2, data: rows[i] };
  }
  return null;
}

function setCells_(name, row, obj) {
  const sh = sheet_(name);
  const head = HEADERS[name];
  Object.keys(obj).forEach(function (k) {
    const c = head.indexOf(k);
    if (c >= 0) sh.getRange(row, c + 1).setNumberFormat('@').setValue(String(obj[k]));
  });
}

function appendText_(sh, values) {
  const row = sh.getLastRow() + 1;
  sh.getRange(row, 1, 1, values.length).setNumberFormat('@').setValues([values]);
  return row;
}

function staffLink_(key) {
  const url = PropertiesService.getScriptProperties().getProperty('WEBAPP_URL');
  return url ? url + '?k=' + key : '';
}

function isOn_(v) {
  const s = String(v).trim().toUpperCase();
  return s !== '否' && s !== 'FALSE' && s !== 'NO' && s !== '0' && s !== '停用';
}

function parseQS_(s) {
  const o = {};
  String(s || '').split('&').forEach(function (p) {
    if (!p) return;
    const i = p.indexOf('=');
    const k = i < 0 ? p : p.slice(0, i);
    o[decodeURIComponent(k)] = i < 0 ? '' : decodeURIComponent(p.slice(i + 1));
  });
  return o;
}

function looksLikeCode_(t) {
  return /^[A-Za-z0-9][A-Za-z0-9\-_/#.]{3,39}$/.test(t);
}

function now_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'); }
function todayYmd_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }

function ymdParts_(ymd) { return String(ymd).slice(0, 10).split('-').map(Number); }

function daysBetween_(a, b) {
  const pa = ymdParts_(a), pb = ymdParts_(b);
  if (pa.length !== 3 || pb.length !== 3 || !pa[0] || !pb[0]) return 0;
  return Math.round((Date.UTC(pb[0], pb[1] - 1, pb[2]) - Date.UTC(pa[0], pa[1] - 1, pa[2])) / 86400000);
}

function addDays_(ymd, n) {
  const p = ymdParts_(ymd);
  return new Date(Date.UTC(p[0], p[1] - 1, p[2] + n)).toISOString().slice(0, 10);
}

function dtToUtcMs_(dt) {
  const p = ymdParts_(dt);
  const t = String(dt).slice(11, 16).split(':').map(Number);
  return Date.UTC(p[0], p[1] - 1, p[2], t[0] || 0, t[1] || 0);
}

function addMinutes_(dt, mins) {
  return new Date(dtToUtcMs_(dt) + mins * 60000).toISOString().slice(0, 16).replace('T', ' ');
}

function minutesBetween_(a, b) { return Math.round((dtToUtcMs_(b) - dtToUtcMs_(a)) / 60000); }

function pad2_(n) { return ('0' + n).slice(-2); }

/** 2026/11/20 */
function fmtYmd_(ymd) {
  const p = ymdParts_(ymd);
  if (p.length !== 3 || !p[0]) return String(ymd || '');
  return p[0] + '/' + pad2_(p[1]) + '/' + pad2_(p[2]);
}

/** 11/20 */
function fmtMD_(ymd) {
  const p = ymdParts_(ymd);
  return p.length === 3 && p[0] ? pad2_(p[1]) + '/' + pad2_(p[2]) : String(ymd || '');
}

/** 2026/11/20 15:00 */
function fmtDT_(dt) {
  if (!dt) return '';
  return fmtYmd_(dt) + ' ' + String(dt).slice(11, 16);
}

/** 接受 2026-11-20、2026/11/20、11/20、11月20日 */
function parseDate_(t) {
  if (!t) return '';
  const s = String(t).trim().replace(/[年月.]/g, '/').replace(/[日號号]/g, '').replace(/-/g, '/').replace(/\s+/g, '');
  let y, m, d;
  let mm = s.match(/^(\d{4})\/(\d{1,2})\/(\d{1,2})$/);
  if (mm) { y = +mm[1]; m = +mm[2]; d = +mm[3]; }
  else {
    mm = s.match(/^(\d{1,2})\/(\d{1,2})$/);
    if (!mm) return '';
    m = +mm[1]; d = +mm[2];
    const today = todayYmd_();
    y = +today.slice(0, 4);
    if (y + '-' + pad2_(m) + '-' + pad2_(d) < today) y += 1;
  }
  const dt = new Date(Date.UTC(y, m - 1, d));
  if (dt.getUTCMonth() !== m - 1 || dt.getUTCDate() !== d) return '';
  return y + '-' + pad2_(m) + '-' + pad2_(d);
}

/** 接受 11/20 14:30、2026/11/20 1430、11/20 下午2:30 */
function parseDateTime_(t) {
  if (!t) return '';
  let s = String(t).trim().replace('T', ' ').replace(/：/g, ':');
  const tm = s.match(/(上午|下午|晚上|早上)?\s*(\d{1,2}):?(\d{2})\s*$/);
  if (!tm) return '';
  const datePart = s.slice(0, tm.index).trim();
  const date = parseDate_(datePart);
  if (!date) return '';
  let h = +tm[2];
  const mi = +tm[3];
  if ((tm[1] === '下午' || tm[1] === '晚上') && h < 12) h += 12;
  if (h > 23 || mi > 59) return '';
  return date + ' ' + pad2_(h) + ':' + pad2_(mi);
}

function esc_(s) {
  return String(s).replace(/[&<>"']/g, function (c) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
  });
}

function safeJson_(v) {
  return JSON.stringify(v).replace(/</g, '\\u003c').replace(/>/g, '\\u003e');
}

function log_(source, msg) {
  try {
    const sh = sheet_(SHEETS.LOG);
    sh.appendRow([now_(), source, String(msg).slice(0, 1000)]);
    if (sh.getLastRow() > 2000) sh.deleteRows(2, 500);
  } catch (_) { /* 紀錄失敗不影響主流程 */ }
}
