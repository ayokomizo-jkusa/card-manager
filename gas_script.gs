// ============================================================
// Card Manager - Google Apps Script
// ============================================================

const SHEET_NAME_CARDS   = 'Cards';
const SHEET_NAME_HISTORY = 'History';
const SHEET_NAME_SETTINGS = 'Settings';
const FIREBASE_PROJECT_ID = 'card-manager-b9a2f';

// ── ルーティング ──────────────────────────────────────────
function doGet(e)  { return handleRequest(e); }
function doPost(e) { return handleRequest(e); }

function handleRequest(e) {
  let result;
  try {
    const body   = e.postData ? JSON.parse(e.postData.contents) : e.parameter;
    const action = body.action || e.parameter?.action;

    switch (action) {
      case 'getData':        result = getData();                   break;
      case 'saveData':       result = saveData(body.data);        break;
      case 'saveHistory':    result = saveHistory(body.record);   break;
      case 'getHistory':     result = getHistory();               break;
      case 'saveFcmToken':   result = saveFcmToken(body.token);   break;
      default:               result = { success: false, error: 'Unknown action' };
    }
  } catch (err) {
    result = { success: false, error: err.toString() };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── データ保存・取得 ──────────────────────────────────────
function getData() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME_CARDS);
  if (!sheet) return { success: true, data: null };
  const raw = sheet.getRange(1, 1).getValue();
  try { return { success: true, data: raw ? JSON.parse(raw) : null }; }
  catch(e) { return { success: true, data: null }; }
}

function saveData(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME_CARDS) || ss.insertSheet(SHEET_NAME_CARDS);
  sheet.getRange(1, 1).setValue(JSON.stringify(data));
  return { success: true };
}

function saveHistory(record) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME_HISTORY);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME_HISTORY);
    sheet.appendRow(['Date', 'Card', 'Benefit', 'Period', 'Action']);
  }
  sheet.appendRow([record.date, record.cardName, record.benefitName, record.period, record.action]);
  return { success: true };
}

function getHistory() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME_HISTORY);
  if (!sheet) return { success: true, history: [] };
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { success: true, history: [] };
  const headers = data[0];
  const history = data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });
  return { success: true, history };
}

// ── FCMトークン ───────────────────────────────────────────
function saveFcmToken(token) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME_SETTINGS) || ss.insertSheet(SHEET_NAME_SETTINGS);
  sheet.getRange(1, 1).setValue(token);
  return { success: true };
}

function getFcmToken() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME_SETTINGS);
  if (!sheet) return null;
  const val = sheet.getRange(1, 1).getValue();
  return val || null;
}

// ── 毎朝の通知送信（時間トリガーで実行）─────────────────
function sendDailyNotifications() {
  const token = getFcmToken();
  if (!token) return;

  const result = getData();
  if (!result.data || !result.data.cards) return;

  const urgent = getUrgentBenefits(result.data);
  if (urgent.length === 0) return;

  // まとめて1件の通知にする
  const lines = urgent.map(b => `${b.cardName}: ${b.name} 残${b.daysLeft}日`);
  const title = `CardMgr — ${urgent.length}件の特典が期限間近`;
  const body  = lines.join('\n');

  sendFcmNotification(token, title, body);
}

// ── 期限間近の特典を取得 ──────────────────────────────────
const CARD_PRESETS_GAS = {
  amex_gold:    [
    { id: 'uber_cash',       period: 'monthly',     name: 'Uber Cash' },
    { id: 'dining_credit',   period: 'monthly',     name: 'Dining Credit' },
    { id: 'dunkin_credit',   period: 'monthly',     name: "Dunkin' Credit" },
    { id: 'resy_credit',     period: 'semi-annual', name: 'Resy Credit' },
    { id: 'hotel_collection',period: 'annual',      name: 'Hotel Collection Credit' },
    { id: 'uber_one',        period: 'limited',     name: 'Uber One Credit', expiry: '2026-10-30' },
  ],
  chase_csp:    [
    { id: 'doordash_credit', period: 'monthly',     name: 'DoorDash Credit' },
    { id: 'chase_hotel',     period: 'annual',      name: 'Chase Travel Hotel Credit' },
  ],
  hilton_aspire:[
    { id: 'flight_credit',   period: 'quarterly',   name: 'Flight Credit' },
    { id: 'resort_credit',   period: 'semi-annual', name: 'Resort Credit' },
    { id: 'clear_credit',    period: 'annual',      name: 'CLEAR+ Credit' },
    { id: 'free_night',      period: 'annual',      name: 'Free Night Reward' },
  ],
  bilt_blue:    [
    { id: 'bilt_cash',       period: 'annual',      name: 'Bilt Cash' },
  ],
  delta_gold:   [
    { id: 'delta_stays',     period: 'annual',      name: 'Delta Stays Credit' },
    { id: 'delta_flight',    period: 'annual',      name: 'Flight Credit ($10K条件)' },
  ],
  atmos_ascent: [
    { id: 'companion_fare',  period: 'annual',      name: 'Companion Fare ($6K条件)' },
  ],
};

function getUrgentBenefits(data) {
  const now      = new Date();
  const y        = now.getFullYear();
  const m        = now.getMonth() + 1;
  const urgent   = [];

  data.cards.forEach(card => {
    const benefits = CARD_PRESETS_GAS[card.id] || [];
    const notifyDays = card.notifyDays || data.settings?.defaultNotifyDays || 3;

    benefits.forEach(benefit => {
      const periodKey = getCurrentPeriodKeyGas(benefit.period, y, m);
      const checkKey  = `${card.id}_${benefit.id}_${periodKey}`;
      const isChecked = data.checks?.[checkKey]?.checked || false;
      if (isChecked) return;

      const daysLeft = getDaysLeftGas(benefit.period, benefit.expiry);
      if (daysLeft !== null && daysLeft <= notifyDays) {
        urgent.push({ cardName: card.name, name: benefit.name, daysLeft });
      }
    });
  });

  return urgent;
}

function getCurrentPeriodKeyGas(period, y, m) {
  const q = Math.ceil(m / 3);
  const h = m <= 6 ? 1 : 2;
  switch(period) {
    case 'monthly':     return `${y}-${String(m).padStart(2,'0')}`;
    case 'quarterly':   return `${y}-Q${q}`;
    case 'semi-annual': return `${y}-H${h}`;
    case 'annual':      return `${y}`;
    case 'limited':     return 'limited';
    default:            return `${y}`;
  }
}

function getDaysLeftGas(period, expiry) {
  const now = new Date();
  const y   = now.getFullYear();
  const m   = now.getMonth();
  let end;
  switch(period) {
    case 'monthly':     end = new Date(y, m + 1, 0); break;
    case 'quarterly':   end = new Date(y, (Math.floor(m / 3) + 1) * 3, 0); break;
    case 'semi-annual': end = m < 6 ? new Date(y, 5, 30) : new Date(y, 11, 31); break;
    case 'annual':      end = new Date(y, 11, 31); break;
    case 'limited':     end = expiry ? new Date(expiry) : null; break;
    default:            return null;
  }
  return end ? Math.ceil((end - now) / 86400000) : null;
}

// ── FCM送信（v1 API）─────────────────────────────────────
function sendFcmNotification(token, title, body) {
  const accessToken = ScriptApp.getOAuthToken();
  const url = `https://fcm.googleapis.com/v1/projects/${FIREBASE_PROJECT_ID}/messages:send`;

  UrlFetchApp.fetch(url, {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${accessToken}`,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify({
      message: {
        token: token,
        notification: { title, body },
        webpush: {
          notification: {
            icon: 'https://ayokomizo-jkusa.github.io/card-manager/icon-192.png',
            requireInteraction: true
          }
        }
      }
    }),
    muteHttpExceptions: true
  });
}
