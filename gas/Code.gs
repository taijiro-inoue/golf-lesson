/**
 * ゴルフレッスン サブスク管理 - Google Apps Script
 * 
 * 【シート構成】
 *   シート1: "members"   → 会員マスタ
 *   シート2: "lessons"   → レッスン受講ログ
 *   シート3: "calendars" → カレンダー設定
 *
 * 【セットアップ手順】
 *   1. スプレッドシートを新規作成
 *   2. 拡張機能 > Apps Script にこのコードを貼り付け
 *   3. setupSheets() を一度だけ手動実行してシートを初期化
 *   4. setTrigger() を一度だけ手動実行してトリガーを登録
 *   5. カレンダーIDをcalendarsシートに入力
 *   6. membersシートに会員情報を入力
 */

// ============================================================
// 定数
// ============================================================
const SHEET_MEMBERS   = 'members';
const SHEET_LESSONS   = 'lessons';
const SHEET_CALENDARS = 'calendars';

// ============================================================
// 1. 初期セットアップ（最初に一度だけ手動実行）
// ============================================================
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- members シート ---
  let ms = ss.getSheetByName(SHEET_MEMBERS) || ss.insertSheet(SHEET_MEMBERS);
  ms.clearContents();
  ms.getRange(1,1,1,7).setValues([[
    'id', '名前', 'メールアドレス', 'プラン回数/月', '契約開始日', '持越し回数', '備考'
  ]]);
  ms.setFrozenRows(1);
  // サンプルデータ
  ms.getRange(2,1,2,7).setValues([
    [1, '田中 恵子', 'tanaka@example.com', 4, '2025-12-10', 0, ''],
    [2, '鈴木 大介', 'suzuki@example.com', 2, '2026-03-20', 1, ''],
  ]);

  // --- lessons シート ---
  let ls = ss.getSheetByName(SHEET_LESSONS) || ss.insertSheet(SHEET_LESSONS);
  ls.clearContents();
  ls.getRange(1,1,1,6).setValues([[
    '日付', '会員id', '名前', 'メールアドレス', 'カレンダー名', 'メモ'
  ]]);
  ls.setFrozenRows(1);

  // --- calendars シート ---
  let cs = ss.getSheetByName(SHEET_CALENDARS) || ss.insertSheet(SHEET_CALENDARS);
  cs.clearContents();
  cs.getRange(1,1,1,3).setValues([['カレンダー名', 'カレンダーID', 'プラン回数']]);
  cs.setFrozenRows(1);
  cs.getRange(2,1,3,3).setValues([
    ['月2回プラン', 'ここにGoogleカレンダーのIDを貼り付け@group.calendar.google.com', 2],
    ['月4回プラン', 'ここにGoogleカレンダーのIDを貼り付け@group.calendar.google.com', 4],
    ['月6回プラン', 'ここにGoogleカレンダーのIDを貼り付け@group.calendar.google.com', 6],
  ]);

  SpreadsheetApp.getUi().alert('✅ セットアップ完了！\ncalendarsシートにカレンダーIDを入力してください。');
}

// ============================================================
// 1b. 会員初期データ投入（最初に一度だけ手動実行）
//     ※ setupSheets() 実行後に実行してください
//     ※ メールアドレスは @placeholder.com になっているので
//       実際のアドレスに差し替えてください
// ============================================================
function seedMembers() {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const mSheet = ss.getSheetByName(SHEET_MEMBERS);

  // ヘッダー行以降をクリア
  if (mSheet.getLastRow() > 1) {
    mSheet.getRange(2, 1, mSheet.getLastRow() - 1, 7).clearContent();
  }

  // [id, 名前, メール, プラン回数/月, 契約開始日, 持越し回数, 備考]
  // plan=0 はレッスンパック会員（月次付与なし・残回数はcarryOverで管理）
  const data = [
    // ── 月2回（隔週）──────────────────────────────────
    [1,  'コジマ',         'kojima@placeholder.com',        2, '2026-06-20',  0, '公式LINE｜月2時間サブスク｜¥59,000/レッスン'],
    [2,  '大塚',           'otsuka@placeholder.com',        2, '2026-05-28',  0, '公式LINE｜月2時間サブスク｜¥59,000/レッスン'],
    [3,  'Nao.H（縦也）',  'naoh@placeholder.com',          3, '2026-06-20',  0, '公式LINE｜月3時間サブスク｜¥59,000/レッスン'],
    [4,  '田中宏和',        'tanaka.hirokazu@placeholder.com', 0, '2026-06-11', 1, '公式LINE｜レッスンパック残1/12｜¥57,000/レッスン（Square手数料込）'],
    [5,  'イタバ',          'itaba@placeholder.com',         4, '2026-06-11',  0, '公式LINE｜月4時間サブスク｜月額¥235,000｜¥59,000/レッスン'],
    [6,  'Miyuki',          'miyuki@placeholder.com',        0, '2026-07-14', 12, 'レッスンパック残12/24｜¥24,000/レッスン'],
    [7,  '小林けん',        'kobayashi.ken@placeholder.com', 0, '2026-07-14', 14, 'レッスンパック残14/24＋ラウンドパック｜¥22,500/レッスン｜ラウンド:¥200,000'],
    [8,  '檜山',            'hiyama@placeholder.com',        0, '2026-07-14', 10, 'レッスンパック残10/24＋ラウンドパック｜¥22,800/レッスン｜ラウンド:¥200,000'],
    [9,  '大場',            'oba@placeholder.com',           4, '2026-03-01',  0, '公式LINE(J.S.OB)｜サブスク申込5/7｜月額¥117,000｜¥28,518/レッスン｜更新:毎月1日'],
    // ── 隔週（2週に1回）────────────────────────────────
    [10, 'Min（Enmin）',   'min.enmin@placeholder.com',     3, '2026-03-08',  0, 'Instagram｜月3時間サブスク｜土曜｜月額¥120,000｜¥40,000/レッスン'],
    [11, '岩井',            'iwai@placeholder.com',          0, '2026-07-14',  6, 'レッスンパック残6/24＋ラウンドパック｜合計¥524,000｜¥21,250/レッスン｜ラウンド:¥140,000'],
    [12, '水谷',            'mizutani@placeholder.com',      0, '2026-07-14',  4, '公式LINE(トモヒコ)｜レッスンパック残4/12｜¥50,000/レッスン'],
    [13, '清水',            'shimizu@placeholder.com',       0, '2026-07-14',  7, 'レッスンパック残7/12｜¥19,800/レッスン｜次回トレーニングプラン準備'],
    [14, '堀川',            'horikawa@placeholder.com',      0, '2026-07-14', 11, '公式LINE｜高田馬場｜レッスンパック残11/24｜¥15,500/レッスン'],
    [15, 'Shota',           'shota@placeholder.com',         0, '2026-07-14',  5, 'Ken・公式LINE｜マレーシアオンライン｜レッスンパック残5/24｜¥17,900/レッスン'],
    [16, 'Akiko Naruse',   'akiko.naruse@placeholder.com',  0, '2026-07-14', 12, '公式LINE｜レッスンパック残12/24｜¥15,500/レッスン｜最終¥150,000は2/14期日'],
    // ── 月1回──────────────────────────────────────────
    [17, '森',              'mori@placeholder.com',          4, '2026-02-14',  0, 'Instagram(Mori)｜月4時間サブスク｜月額¥115,000｜¥28,500/レッスン'],
    [18, 'Tak A',           'taka@placeholder.com',          2, '2026-07-14',  0, 'Instagram｜月2時間サブスク｜月額¥100,000｜※5月分以降サブスク確認'],
    [19, 'Takuya',          'takuya@placeholder.com',        0, '2026-07-14',  4, '個人LINE｜レッスンパック残4/12｜¥15,000/レッスン（パック¥13,500）'],
    [20, '荒川',            'arakawa@placeholder.com',       1, '2026-07-14',  0, '公式LINE｜単発レッスン｜¥22,000/レッスン（パック¥20,900）'],
  ];

  mSheet.getRange(2, 1, data.length, 7).setValues(data);
  SpreadsheetApp.getUi().alert(
    `✅ ${data.length}名の会員データを投入しました。\n\n` +
    '⚠ メールアドレスが @placeholder.com になっています。\n' +
    'Googleカレンダー予約と照合するために、実際のアドレスに更新してください。'
  );
}

// ============================================================
// 2. トリガー登録（最初に一度だけ手動実行）
// ============================================================
function setTrigger() {
  // 既存トリガーを削除
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));

  // 毎日 23:00 に自動実行（カレンダー同期＋月次更新を一括実行）
  ScriptApp.newTrigger('dailySyncWithRenewal')
    .timeBased()
    .everyDays(1)
    .atHour(23)
    .create();

  SpreadsheetApp.getUi().alert('✅ トリガー設定完了！\n毎日23:00に自動でGoogleカレンダー確認＋月次更新を実行します。');
}

// ============================================================
// 3. メイン処理（毎日自動実行 or 手動実行可）
// ============================================================
function dailySync() {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const mSheet  = ss.getSheetByName(SHEET_MEMBERS);
  const lSheet  = ss.getSheetByName(SHEET_LESSONS);
  const cSheet  = ss.getSheetByName(SHEET_CALENDARS);

  // --- カレンダー設定を読み込む ---
  const calData = cSheet.getRange(2, 1, cSheet.getLastRow()-1, 3).getValues()
    .filter(r => r[1] && !r[1].toString().includes('ここに'));
  if (!calData.length) {
    console.log('カレンダーIDが未設定です');
    return;
  }

  // --- 会員マスタをメール→行情報のMapにする ---
  const memberRows = mSheet.getRange(2, 1, Math.max(mSheet.getLastRow()-1,1), 7).getValues();
  const memberByEmail = {};
  memberRows.forEach((row, i) => {
    if (row[2]) memberByEmail[row[2].trim().toLowerCase()] = { rowIndex: i+2, data: row };
  });

  // --- 処理済みレッスンログをSetにする（重複防止）---
  const logRows = lSheet.getLastRow() > 1
    ? lSheet.getRange(2, 1, lSheet.getLastRow()-1, 4).getValues()
    : [];
  const processedKeys = new Set(logRows.map(r => `${r[0]}_${r[3]}`)); // 日付_メール

  // --- 今日の予約をカレンダーから取得 ---
  const today     = new Date();
  const startOfDay = new Date(today.getFullYear(), today.getMonth(), today.getDate(), 0, 0, 0);
  const endOfDay   = new Date(today.getFullYear(), today.getMonth(), today.getDate(), 23, 59, 59);
  const todayStr  = Utilities.formatDate(today, 'Asia/Tokyo', 'yyyy-MM-dd');

  let newLessons = [];

  calData.forEach(([calName, calId, planCount]) => {
    let cal;
    try {
      cal = CalendarApp.getCalendarById(calId);
    } catch(e) {
      console.log(`カレンダー取得エラー: ${calId}`);
      return;
    }
    if (!cal) return;

    const events = cal.getEvents(startOfDay, endOfDay);
    events.forEach(event => {
      // 予約スケジュールのゲストからメールを取得
      const guests = event.getGuestList();
      guests.forEach(guest => {
        const email = guest.getEmail().trim().toLowerCase();
        const key   = `${todayStr}_${email}`;

        if (processedKeys.has(key)) return; // 既処理スキップ

        const member = memberByEmail[email];
        if (!member) {
          console.log(`未登録の予約者: ${email}`);
          return;
        }

        // lessonsシートに記録
        const guestName = guest.getName() || member.data[1];
        newLessons.push([todayStr, member.data[0], guestName, email, calName, 'カレンダー自動取得']);
        processedKeys.add(key);
        console.log(`✅ レッスン記録: ${guestName} (${email}) - ${calName}`);
      });
    });
  });

  // まとめてシートに書き込み
  if (newLessons.length > 0) {
    const nextRow = lSheet.getLastRow() + 1;
    lSheet.getRange(nextRow, 1, newLessons.length, 6).setValues(newLessons);
    console.log(`${newLessons.length}件のレッスンを記録しました`);
  } else {
    console.log('本日の新規レッスン予約はありませんでした');
  }
}

// ============================================================
// 4. 月次更新処理（各会員の更新日に自動実行）
//    dailySync() の中から呼び出される
// ============================================================
function monthlyRenewal() {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const mSheet = ss.getSheetByName(SHEET_MEMBERS);
  const lSheet = ss.getSheetByName(SHEET_LESSONS);
  const today  = new Date();
  const todayStr = Utilities.formatDate(today, 'Asia/Tokyo', 'yyyy-MM-dd');

  const memberRows = mSheet.getRange(2, 1, Math.max(mSheet.getLastRow()-1,1), 7).getValues();

  memberRows.forEach((row, i) => {
    const [id, name, email, planCount, startDateRaw, carryOver] = row;
    if (!startDateRaw) return;

    const startDate = new Date(startDateRaw);
    // 更新日かどうか判定（契約開始日と同じ「日」で、月が異なる）
    if (today.getDate() !== startDate.getDate()) return;
    if (today <= startDate) return; // 開始月はスキップ

    // 今期（前の期間）のレッスン使用数を計算
    const periodStart = new Date(today.getFullYear(), today.getMonth()-1, startDate.getDate());
    const periodEnd   = new Date(today.getFullYear(), today.getMonth(),   startDate.getDate());
    const pStartStr   = Utilities.formatDate(periodStart, 'Asia/Tokyo', 'yyyy-MM-dd');
    const pEndStr     = Utilities.formatDate(periodEnd,   'Asia/Tokyo', 'yyyy-MM-dd');

    const logRows = lSheet.getLastRow() > 1
      ? lSheet.getRange(2, 1, lSheet.getLastRow()-1, 4).getValues()
      : [];

    const usedThisPeriod = logRows.filter(r => {
      const d = r[0] ? r[0].toString().slice(0,10) : '';
      return r[3]?.toString().toLowerCase() === email?.toString().toLowerCase()
          && d >= pStartStr && d < pEndStr;
    }).length;

    const totalAllotted = Number(planCount) + Number(carryOver);
    const newCarry      = Math.max(0, totalAllotted - usedThisPeriod);

    // 持越し回数を更新
    mSheet.getRange(i+2, 6).setValue(newCarry);
    console.log(`更新: ${name} 持越し ${newCarry}回 (使用${usedThisPeriod}/${totalAllotted})`);
  });
}

// dailySyncに月次更新を組み込む（関数の末尾に追記する形）
// ※ dailySync()の末尾でmonthlyRenewal()を呼ぶ
function dailySyncWithRenewal() {
  dailySync();
  monthlyRenewal();
}

// ============================================================
// 5. Web App として公開するためのエンドポイント
//    HTMLツールからfetchで叩いてデータを取得
// ============================================================
function doGet(e) {
  const action = e?.parameter?.action || 'getAll';

  if (action === 'getAll') {
    return ContentService.createTextOutput(JSON.stringify(getAllData()))
      .setMimeType(ContentService.MimeType.JSON);
  }

  return ContentService.createTextOutput(JSON.stringify({ error: 'unknown action' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const body = JSON.parse(e.postData.contents);
  const action = body.action;

  if (action === 'addLesson') {
    return ContentService.createTextOutput(JSON.stringify(addLesson(body)))
      .setMimeType(ContentService.MimeType.JSON);
  }
  if (action === 'addMember') {
    return ContentService.createTextOutput(JSON.stringify(addMember(body)))
      .setMimeType(ContentService.MimeType.JSON);
  }
  if (action === 'updateMember') {
    return ContentService.createTextOutput(JSON.stringify(updateMember(body)))
      .setMimeType(ContentService.MimeType.JSON);
  }

  return ContentService.createTextOutput(JSON.stringify({ error: 'unknown action' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// 6. データ取得・操作ヘルパー
// ============================================================
function getAllData() {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const mSheet = ss.getSheetByName(SHEET_MEMBERS);
  const lSheet = ss.getSheetByName(SHEET_LESSONS);

  const members = mSheet.getLastRow() > 1
    ? mSheet.getRange(2,1,mSheet.getLastRow()-1,7).getValues()
        .filter(r => r[0])
        .map(r => ({
          id:        r[0],
          name:      r[1],
          email:     r[2],
          plan:      r[3],
          startDate: r[4] ? Utilities.formatDate(new Date(r[4]),'Asia/Tokyo','yyyy-MM-dd') : '',
          carryOver: r[5] || 0,
          note:      r[6] || ''
        }))
    : [];

  const lessons = lSheet.getLastRow() > 1
    ? lSheet.getRange(2,1,lSheet.getLastRow()-1,6).getValues()
        .filter(r => r[0])
        .map(r => ({
          date:      r[0] ? Utilities.formatDate(new Date(r[0]),'Asia/Tokyo','yyyy-MM-dd') : '',
          memberId:  r[1],
          name:      r[2],
          email:     r[3],
          calendar:  r[4],
          note:      r[5]
        }))
    : [];

  return { members, lessons };
}

function addLesson(body) {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const lSheet = ss.getSheetByName(SHEET_LESSONS);
  const today  = Utilities.formatDate(new Date(),'Asia/Tokyo','yyyy-MM-dd');
  lSheet.appendRow([body.date || today, body.memberId, body.name, body.email, body.calendar || '手動', body.note || '手動追加']);
  return { ok: true };
}

function addMember(body) {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const mSheet = ss.getSheetByName(SHEET_MEMBERS);
  const lastId = mSheet.getLastRow() > 1
    ? mSheet.getRange(mSheet.getLastRow(),1).getValue()
    : 0;
  mSheet.appendRow([lastId+1, body.name, body.email, body.plan, body.startDate, 0, body.note||'']);
  return { ok: true, id: lastId+1 };
}

function updateMember(body) {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const mSheet = ss.getSheetByName(SHEET_MEMBERS);
  const rows   = mSheet.getRange(2, 1, Math.max(mSheet.getLastRow()-1, 1), 7).getValues();
  const rowIdx = rows.findIndex(r => r[0] == body.id);
  if (rowIdx === -1) return { ok: false, error: 'member not found' };
  const sheetRow = rowIdx + 2;
  if (body.name      !== undefined) mSheet.getRange(sheetRow, 2).setValue(body.name);
  if (body.email     !== undefined) mSheet.getRange(sheetRow, 3).setValue(body.email.toLowerCase());
  if (body.plan      !== undefined) mSheet.getRange(sheetRow, 4).setValue(body.plan);
  if (body.startDate !== undefined) mSheet.getRange(sheetRow, 5).setValue(body.startDate);
  if (body.note      !== undefined) mSheet.getRange(sheetRow, 7).setValue(body.note);
  return { ok: true };
}
