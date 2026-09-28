/**
 * 習慣追蹤 App 後端（Google Apps Script，綁定在 Google 試算表上）
 * 2026-09-28 重建：原試算表遭誤刪且已無法復原，改寫此版並存入 git（教訓 6、10：
 * 後端程式碼原本不在版本控制裡，是最容易被忘記、也最怕遺失的部分）。
 *
 * 試算表結構（第一列為表頭，順序不可變，A~E 欄）：
 *   date | habit | done | backfilled | type
 * - date：一律存純文字 'YYYY-MM-DD'。寫入時在字串前加單引號強制純文字（見 setDateCellAsText_）
 *   再 setValue，不依賴欄位事先設定的格式（實測發現只設欄位格式仍會被自動轉成 Date，
 *   即最初版本的 H-8）。
 * - habit：習慣或每月事項的 key，對應前端 HABIT_CONFIG / MONTHLY_CONFIG。
 * - done、backfilled：布林值 TRUE/FALSE。
 * - type：'daily' 或 'monthly'。
 *
 * 寫入邏輯（upsert）：
 * - 以 date+habit+type 為鍵。找到既有列就覆寫 done、backfilled。
 * - body 帶 allowDuplicate:true 時，略過比對，直接新增一列（每月「更新」用這個，
 *   讓同一天可以有多筆完成紀錄）。
 * - 用 LockService 序列化寫入，避免多裝置同時打卡互相覆蓋。
 *
 * 寫入防護（2026-09-28，H-1）：doPost 只接受 habit 在下方白名單、date 為
 * yyyy-MM-dd、type 為 daily/monthly 的請求，其餘一律回 400 不寫入。
 * Web App 設「任何人」可存取、網址又公開在 GitHub 上，這層檢查擋不掉知道網址的人
 * 亂送請求，只能防止「打錯字的假資料」寫進試算表；要真正擋人，得改成需要登入或限流。
 *
 * ⚠️ 維護提醒：下面 VALID_HABITS 要跟前端 index.html 的 HABIT_CONFIG／
 * MONTHLY_CONFIG 手動保持一致。前端新增或刪除習慣時，這裡要同步改，
 * 否則新習慣會被這裡擋掉、寫不進去。
 */

const SHEET_NAME = 'records';

const VALID_HABITS = [
  // 每日（HABIT_CONFIG）
  'english-reading', 'english-listening', 'english-writing', 'healthy-diet',
  'exercise', 'dental-floss', 'skincare', 'water-intake', 'supplement',
  'blue-light', 'relax', 'less-phone', 'reading', 'web-dev',
  'wealth-mindset', 'maritime-law',
  // 每月（MONTHLY_CONFIG）
  'towel', 'razor', 'bathroom', 'jacket', 'bedsheet', 'bath-mat',
  'supplement-check', 'accounting', 'diary', 'backup', 'weight',
  'habit-review', 'rent', 'hair-removal', 'family-finance',
];
const DATE_RE = /^\d{4}-\d{2}-\d{2}$/;

function validateBody_(body) {
  if (!body || typeof body !== 'object') return '缺少 body';
  if (VALID_HABITS.indexOf(body.habit) === -1) return '不合法的 habit: ' + body.habit;
  if (!DATE_RE.test(body.date)) return '不合法的日期格式: ' + body.date;
  if (body.type !== 'daily' && body.type !== 'monthly') return '不合法的 type: ' + body.type;
  return null;
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
}

// Date 物件、或看起來像日期字串但被試算表轉換過的值，一律轉回本地時區的 'YYYY-MM-DD'
function formatDate_(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, 'Asia/Taipei', 'yyyy-MM-dd');
  }
  const s = String(v);
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s; // 已經是純文字的正確格式
  // 防呆：萬一儲存格還是被轉成日期、讀回來變成奇怪字串，嘗試重新解析
  const parsed = new Date(s);
  if (!isNaN(parsed.getTime())) {
    return Utilities.formatDate(parsed, 'Asia/Taipei', 'yyyy-MM-dd');
  }
  return s;
}

function doGet(e) {
  const sheet = getSheet_();
  const values = sheet.getDataRange().getValues();
  const rows = values.slice(1); // 去表頭

  const records = rows
    .filter(r => r[0] !== '' && r[0] !== null)
    .map(r => ({
      date:       formatDate_(r[0]),
      habit:      r[1],
      done:       r[2] === true || r[2] === 'TRUE',
      backfilled: r[3] === true || r[3] === 'TRUE',
      type:       r[4] || 'daily',
    }));

  return ContentService.createTextOutput(JSON.stringify(records))
    .setMimeType(ContentService.MimeType.JSON);
}

// 用前導單引號強制存成純文字：這是 Google 試算表官方認可的強制文字寫法，
// 比單靠 setNumberFormat('@') 更可靠（實測發現後者在部分情況仍會被自動轉成 Date）。
// 讀回來時單引號會自動消失，不影響資料本身。
function setDateCellAsText_(sheet, row, dateStr) {
  sheet.getRange(row, 1).setValue("'" + dateStr);
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const body = JSON.parse(e.postData.contents);
    const invalidReason = validateBody_(body);
    if (invalidReason) {
      return ContentService.createTextOutput(JSON.stringify({ ok: false, error: invalidReason }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    const date = String(body.date);
    const habit = String(body.habit);
    const type = body.type || 'daily';
    const done = !!body.done;
    const backfilled = !!body.backfilled;

    const sheet = getSheet_();
    const values = sheet.getDataRange().getValues();

    if (!body.allowDuplicate) {
      for (let i = 1; i < values.length; i++) {
        const r = values[i];
        if (formatDate_(r[0]) === date && r[1] === habit && (r[4] || 'daily') === type) {
          sheet.getRange(i + 1, 3).setValue(done);       // C 欄 done
          sheet.getRange(i + 1, 4).setValue(backfilled);  // D 欄 backfilled
          SpreadsheetApp.flush();
          return jsonOk_({ action: 'update', row: i + 1 });
        }
      }
    }

    const newRow = sheet.getLastRow() + 1;
    sheet.getRange(newRow, 2, 1, 4).setValues([[habit, done, backfilled, type]]);
    setDateCellAsText_(sheet, newRow, date); // date 欄最後寫，且強制純文字
    SpreadsheetApp.flush();
    return jsonOk_({ action: 'insert', row: newRow });
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function jsonOk_(extra) {
  return ContentService.createTextOutput(JSON.stringify(Object.assign({ ok: true }, extra)))
    .setMimeType(ContentService.MimeType.JSON);
}
