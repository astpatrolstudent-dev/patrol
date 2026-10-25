/**
 * platform clover パトロール シフト管理
 * Google スプレッドシートにバインドした Apps Script として使う
 *
 * シート構成（メニュー「パトロール > ① 初期設定」で自動作成）
 *   メンバー  : 名前 | 有効(TRUE/FALSE) | SlackメンバーID | 登録日時 | 導入前の累計回数
 *               | 氏名（報告書） | フリガナ | 日中連絡先 | 業務内容
 *   月設定    : 対象月(yyyy-MM) | 締切日 | シフト状態(未作成/下書き/確定) | 更新日時
 *   投票      : 対象月 | 名前 | 更新日時 | 1 | 2 | … | 31   （〇 △ ×）
 *   yyyy-MM_パトロール : 日付 | 曜日 | 担当（予定） | 実施者 | 報告日時   （月ごとに1枚・1日1行）
 *   設定      : 項目 | 値 | 説明
 *   業務報告  : ID | 対象月 | 名前 | 日付 | 作業時間(h) | 作業詳細 | 登録日時   （パトロール以外の業務）
 *   業務報告_交通費 : 対象月 | 名前 | 往復運賃(円) | 日数 | 更新日時
 *   業務報告書テンプレート : 業務報告書の書式（xlsx を取り込んだもの。出力時にこれを複製して記入する）
 *   実績_yyyy-MM : 活動実績（出力時に自動作成）
 */

const SH = {
  MEMBERS: 'メンバー',
  MONTHS: '月設定',
  VOTES: '投票',
  SETTINGS: '設定',
  WORK: '業務報告',
  FARE: '業務報告_交通費',
  PROFILE: '報告者情報',
  LINE: 'LINE登録',
};
const LINE_HEADER = ['名前', 'LINEユーザーID', '登録日時'];
const APP_VERSION = '1.21';
/* ▼ サイトの URL（通知の末尾に載せる。「設定」シートの APP_URL があればそちらを優先） */
const SITE_URL = 'https://astpatrolstudent-dev.github.io/patrol/';
/* ▼ 最初の画面のデータを、サーバー側に一時保存する時間（秒）。書き込みがあれば自動で捨てる */
const BOOT_CACHE_SEC = 600;

/*
 * ▼ ニーズ調査サイトのログイン情報（⚙画面でコピーできるように表示する）
 *   公開される index.html には書かず、ここ（Apps Script）にだけ置く。
 *   変更するときは、ここを書き換えるか、「設定」シートに NEEDS_ID / NEEDS_PW の行を足す（シートが優先）。
 */
const NEEDS_LOGIN = { id: 'teruya@sustainable-transition.or.jp', pw: 'Ast201201' };
const WORK_TEMPLATE = '業務報告書テンプレート';
const WORK_ROWS = 22;          // テンプレートの明細行数（14〜35行目）
const WORK_FIRST_ROW = 14;
const WORK_PRINT_RANGE = 'B1:AD46';
const WORK_HEADER = ['ID', '対象月', '名前', '日付', '作業時間(h)', '作業詳細', '登録日時'];
const FARE_HEADER = ['対象月', '名前', '往復運賃(円)', '日数', '更新日時'];
const PROFILE_HEADER = ['名前', '氏名（報告書）', 'フリガナ', '日中連絡先', '業務内容', '更新日時'];

/*
 * ▼ データを置くスプレッドシート（URL の /d/ と /edit の間の文字列）
 *   SPREADSHEET_ID     : 基本（設定・メンバー・月設定・投票）
 *   PATROL_SS_ID       : AST業務報告（月ごとのパトロール記録・活動実績）
 *   WORK_SS_ID         : 業務報告関連・パトロール以外（業務報告・交通費・報告者情報・テンプレート）
 *   MASTER_DATA_SS_ID  : パトロールマスタデータ（パトロールで見つけた新しい投稿・ユーザーを「入力」タブから追記）
 *   PDF_FOLDER_ID      : PDF・Excel の保存先フォルダ（URL の /folders/ の後ろ）
 *   PATROL_SS_ID / WORK_SS_ID を空にすると、基本のスプレッドシートに入ります。
 */
const SPREADSHEET_ID = '1WMhp1ptPfDuiCkg7nSzGRGI-hBLNrbzhwA8Awx9kVug';
const PATROL_SS_ID = '1znHFJb-ftirbCYkCNBUFCN2JTJm4qmHqp8Ss0tRvgSY';
const WORK_SS_ID = '1BT-KMa5jbGzW66be5IFA0BmNWkew2zyyPSinPVv_wGI';
const MASTER_DATA_SS_ID = '1SEAkV7cwFemPQ7YEuVwcGY-chcHBAOPLbKFLG_bqGRY';
const PDF_FOLDER_ID = '1i3ZEu6zPNJPULVLLi0nckY6T4cKqPhOt';

/*
 * ▼ パトロールマスタデータの各シート
 *   4行目が見出し、5行目からデータ。列は B:項番 C:No. D:名前 E:投稿者（ユーザーは B:項番 C:名前）
 *   dateCol / staffCol / reportCol : 日付・担当者・報告の有無を書き込む列（列記号）
 *   reportText : 報告の有無に入れる文字（赤字）
 */
const MASTER_KINDS = [
  { key: 'project',  sheet: 'プロジェクト',     label: 'プロジェクト',     nameLabel: 'プロジェクト名',   user: false, dateCol: 'FE', staffCol: 'FF', reportCol: 'FG', reportText: '新規投稿確認' },
  { key: 'activity', sheet: 'アクティビティ',   label: 'アクティビティ',   nameLabel: 'アクティビティ名', user: false, dateCol: 'EL', staffCol: 'EM', reportCol: 'EN', reportText: '新規投稿確認' },
  { key: 'needs',    sheet: 'ニーズ',           label: 'ニーズ',           nameLabel: 'ニーズ名',         user: false, dateCol: 'EK', staffCol: 'EL', reportCol: 'EM', reportText: '新規投稿確認' },
  { key: 'seeds',    sheet: 'シーズ',           label: 'シーズ',           nameLabel: 'シーズ名',         user: false, dateCol: 'EQ', staffCol: 'ER', reportCol: 'ES', reportText: '新規投稿確認' },
  { key: 'userP',    sheet: 'ユーザー（個人）', label: 'ユーザー（個人）', nameLabel: 'ユーザー名',       user: true,  dateCol: 'CS', staffCol: 'CT', reportCol: 'CU', reportText: '新規確認' },
  { key: 'userO',    sheet: 'ユーザー（組織）', label: 'ユーザー（組織）', nameLabel: 'ユーザー名',       user: true,  dateCol: 'CW', staffCol: 'CX', reportCol: 'CY', reportText: '新規確認' },
];
const MASTER_FIRST_ROW = 5;
const TZ = 'Asia/Tokyo';
const WORK_HOURS = 0.25;
const MARKS = ['〇', '△', '×'];
const STATUS = { NONE: '未作成', DRAFT: '下書き', FIXED: '確定' };
const PDF_FOLDER = 'パトロール活動実績';
const PATROL_SUFFIX = '_パトロール';
const PATROL_RE = /^(\d{4}-\d{2})_パトロール$/;
const PATROL_HEADER = ['日付', '曜日', '担当（予定）', '実施者', '報告日時'];

/* =========================================================
 * Web アプリ入口
 * ========================================================= */

/**
 * Apps Script の URL（…/exec）を直接開いたとき：今のサイト（GitHub Pages）へ案内する。
 * Apps Script の中の古い画面は使わない（API は doPost だけ）。
 */
function doGet() {
  const url = getSettings_().APP_URL || SITE_URL;
  const safe = String(url).replace(/[<>"']/g, '');
  const html = `<!doctype html><html lang="ja"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1"></head>
<body style="font-family:-apple-system,BlinkMacSystemFont,'Hiragino Sans',sans-serif;padding:32px 20px;text-align:center;line-height:1.7">
<p style="font-size:18px;font-weight:700">パトロールシフトのサイトは移動しました</p>
<p><a href="${safe}" target="_top" style="display:inline-block;background:#2f6fde;color:#fff;padding:12px 22px;border-radius:10px;text-decoration:none;font-weight:700">サイトを開く</a></p>
<p style="color:#666;font-size:14px">${safe}<br>ブックマークやホーム画面は、このアドレスで登録し直してください。</p>
<script>try { window.top.location.href = ${JSON.stringify(safe)}; } catch (e) {}</script>
</body></html>`;
  return HtmlService.createHtmlOutput(html)
    .setTitle('パトロールシフト')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/*
 * 外部サイト（GitHub Pages など）に置いた画面からの呼び出し口。
 * 画面は {fn: '関数名', args: [...]} を POST し、{ok, result | error} の JSON を受け取る。
 * ここに載っている関数だけが外から呼べる。
 */
const API_ = {
  getBootstrap, getMonthData, addMember, setMemberActive, saveVote, setDeadline,
  generateShift, saveShift, setShiftStatus, reportDone, cancelReport, setManager,
  exportReport, testNotify, setLineAdmin, lineQuota,
  saveProfile, addWork, deleteWork, saveFare, exportWorkReport,
  getMasterToday, getMasterLatest, getMasterInfo, addMaster, undoMaster,
};

function doPost(e) {
  let body = {};
  try { body = JSON.parse((e && e.postData && e.postData.contents) || '{}'); } catch (err) { body = {}; }

  // 画面からの呼び出し
  if (body.fn) {
    let out;
    try {
      const f = API_[body.fn];
      if (!f) throw new Error('不明な操作です：' + body.fn);
      out = { ok: true, result: f.apply(null, Array.isArray(body.args) ? body.args : []) };
    } catch (err) {
      out = { ok: false, error: (err && err.message) || String(err) };
    }
    return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
  }

  // LINE の Webhook（友だち追加した人が名前を送ると、その人の LINE を登録する）
  try {
    (body.events || []).forEach(ev => { try { lineEvent_(ev); } catch (err) { console.error(err); } });
  } catch (err) {
    console.error(err);
  }
  return ContentService.createTextOutput('OK');
}

function onOpen() {
  if (SPREADSHEET_ID) return; // アプリを別アカウントに分けている場合、メニューは出さない（エディタから実行する）
  SpreadsheetApp.getUi()
    .createMenu('パトロール')
    .addItem('① 初期設定（シート作成）', 'setup')
    .addItem('② 朝8時・夜22時のお知らせを有効にする', 'installTriggers')
    .addItem('リマインドを止める', 'removeTriggers')
    .addSeparator()
    .addItem('通知テスト', 'testNotifyFromMenu')
    .addItem('今すぐリマインドを実行（動作確認）', 'dailyReminder')
    .addToUi();
}

/* =========================================================
 * 初期設定・トリガー
 * ========================================================= */

function setup() {
  const ss = ss_();
  const defs = [
    [SH.MEMBERS, ['名前', '有効', 'SlackメンバーID', '登録日時', '導入前の累計回数']],
    [SH.MONTHS, ['対象月', '締切日', 'シフト状態', '更新日時']],
    [SH.VOTES, ['対象月', '名前', '更新日時'].concat(range_(31).map(i => String(i + 1)))],
    [SH.SETTINGS, ['項目', '値', '説明']],
  ];
  defs.forEach(([name, header]) => {
    let s = ss.getSheetByName(name);
    if (!s) s = ss.insertSheet(name);
    if (s.getMaxColumns() < header.length) {
      s.insertColumnsAfter(s.getMaxColumns(), header.length - s.getMaxColumns());
    }
    // 日付などが勝手に変換されないよう、すべて「書式なしテキスト」にする
    s.getRange(1, 1, s.getMaxRows(), header.length).setNumberFormat('@');
    // 見出しは毎回書き直す（列を追加したバージョンアップにも対応）
    s.getRange(1, 1, 1, header.length).setValues([header])
      .setFontWeight('bold').setBackground('#e8eef5');
    s.setFrozenRows(1);
  });

  const defaults = [
    ['SLACK_WEBHOOK_URL', '', 'Slack の Incoming Webhook URL（空なら Slack には送らない）'],
    ['LINE_CHANNEL_ACCESS_TOKEN', '', 'LINE公式アカウント（Messaging API）のチャネルアクセストークン（空なら LINE には送らない）'],
    ['LINE_ADMIN', '', 'LINE で「完了」「シフト下書き」の連絡を受け取る管理者の名前（Webの⚙からも変更可）'],
    ['REMIND_ENABLED', 'TRUE', 'FALSE にするとリマインドを一時停止'],
    ['VOTE_ANNOUNCE_DAY', '25', '次月のシフト希望（投票）を全員に告知する日（17時ごろ）'],
    ['VOTE_REMIND_DAY', '', 'まだ回答していない人にリマインドする日（17時ごろ）。空なら月末の前日'],
    ['APP_URL', '', 'この Web アプリの URL（通知の末尾に載せる）'],
    ['MANAGER_NAME', '', '管理者（名目）の名前。活動実績の「手当」欄に載る（Webの「メンバー」タブからも変更可）'],
    ['MANAGER_ALLOWANCE_H', '3', '管理者手当（時間）'],
  ];
  dirty_();
  const have = new Set(readRows_(SH.SETTINGS).map(r => r[0]));
  defaults.forEach(d => { if (!have.has(d[0])) appendRow_(sh_(SH.SETTINGS), d); });

  ['シート1', 'Sheet1'].forEach(n => {
    const s = ss.getSheetByName(n);
    if (s && s.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(s);
  });
  // 業務報告関連のシート（別ファイル）
  ensureSheet_(SH.WORK, WORK_HEADER);
  ensureSheet_(SH.FARE, FARE_HEADER);
  ensureSheet_(SH.PROFILE, PROFILE_HEADER);
  migrateLegacy_();
  toast_('初期設定が完了しました');
}

/**
 * 【1回だけ実行】基本のスプレッドシートにあるデータを、分割先のファイルへ移す（v1.8）
 *   ・○○_パトロール／実績_○○ → AST業務報告（PATROL_SS_ID）
 *   ・業務報告／業務報告_交通費／業務報告書テンプレート → 業務報告関連（WORK_SS_ID）
 *   ・メンバーシートF〜I列（氏名・フリガナ・連絡先・業務内容）→ 業務報告関連の「報告者情報」
 * 元のシートは消さずに「旧_」を付けて残します（確認後に手で削除してください）。
 * 何度実行しても、移行済みのものは飛ばします。
 */
function migrateToSplit() {
  const main = ss_();
  const log = [];
  const move = (sheet, target) => {
    const name = sheet.getName();
    if (target.getId() === main.getId()) return;
    const exist = target.getSheetByName(name);
    if (exist) {
      if (exist.getLastRow() <= 1) target.deleteSheet(exist);   // setup で作った空のシートは置き換える
      else { log.push(`飛ばしました（移行先にすでにデータあり）：${name}`); return; }
    }
    const copy = sheet.copyTo(target);
    copy.setName(name);
    sheet.setName('旧_' + name);
    log.push(`移しました：${name} → ${target.getName()}`);
  };
  const workNames = [SH.WORK, SH.FARE, WORK_TEMPLATE];
  main.getSheets().forEach(sheet => {
    const n = sheet.getName();
    if (PATROL_RE.test(n) || /^実績_\d{4}-\d{2}$/.test(n)) move(sheet, ssPatrol_());
    else if (workNames.indexOf(n) >= 0) move(sheet, ssWork_());
  });
  dirty_();

  // 報告者情報（メンバーシートF〜I列）
  const ms = sh_(SH.MEMBERS);
  const last = ms.getLastRow();
  if (last >= 2 && ms.getLastColumn() >= 6) {
    const rows = ms.getRange(2, 1, last - 1, 9).getDisplayValues();
    const ps = ensureSheet_(SH.PROFILE, PROFILE_HEADER);
    const have = new Set(readRowsOpt_(SH.PROFILE).map(r => r[0].trim()));
    let n = 0;
    rows.forEach(r => {
      const name = String(r[0]).trim();
      if (!name || have.has(name) || !r.slice(5, 9).some(v => String(v).trim())) return;
      appendRow_(ps, [name, r[5], r[6], r[7], r[8], now_()]);
      n++;
    });
    if (n) {
      ms.getRange(1, 6, last, 4).clearContent();   // 個人情報を基本のスプレッドシートから消す
      log.push(`報告者情報を移しました：${n}人（メンバーシートのF〜I列は空にしました）`);
    }
  }

  // 新しいファイルの空の「シート1」を片付ける
  [ssPatrol_(), ssWork_()].forEach(t => {
    ['シート1', 'Sheet1'].forEach(n => {
      const x = t.getSheetByName(n);
      if (x && x.getLastRow() === 0 && t.getSheets().length > 1) t.deleteSheet(x);
    });
  });
  dirty_();
  setup();
  const msg = log.length ? log.join('\n') : '移すものはありませんでした';
  console.log(msg);
  return msg;
}

/** 旧バージョンの「シフト」「完了報告」シートがあれば、月ごとのシートへ移す */
function migrateLegacy_() {
  const ss = ss_();
  const old = { shift: ss.getSheetByName('シフト'), report: ss.getSheetByName('完了報告') };
  if (!old.shift && !old.report) return;
  const byMonth = {};
  const take = (sheet, fn) => {
    if (!sheet || sheet.getLastRow() < 2) return;
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getDisplayValues().forEach(r => {
      if (!isDate_(r[0])) return;
      const ym = monthOf_(r[0]);
      byMonth[ym] = byMonth[ym] || { shifts: {}, reports: {} };
      fn(byMonth[ym], r);
    });
  };
  take(old.shift, (m, r) => { m.shifts[r[0]] = r[1]; });
  take(old.report, (m, r) => { if (r[1]) m.reports[r[0]] = [r[1], r[2]]; });
  Object.keys(byMonth).sort().forEach(ym => {
    const s = ensureMonthSheet_(ym);
    const dates = datesOf_(ym);
    const m = byMonth[ym];
    // すでに月のシートに入っている値は残し、空欄だけ旧データで埋める
    const cur = s.getRange(2, 3, dates.length, 3).getDisplayValues();
    s.getRange(2, 3, dates.length, 3).setNumberFormat('@').setValues(dates.map((d, i) => {
      const r = m.reports[d] || ['', ''];
      const c = cur[i] || ['', '', ''];
      return [c[0] || m.shifts[d] || '', c[1] || r[0], c[1] ? c[2] : r[1]];
    }));
    // 旧データでシフトがあった月は「下書き」扱いにして画面に出す
    if (Object.keys(m.shifts).some(d => m.shifts[d]) && monthSetting_(ym).status === STATUS.NONE) {
      updateMonthSetting_(ym, { status: STATUS.DRAFT });
    }
  });
  if (old.shift) old.shift.setName('シフト（旧・移行済み）');
  if (old.report) old.report.setName('完了報告（旧・移行済み）');
  dirty_();
}

function installTriggers() {
  removeTriggers();
  ScriptApp.newTrigger('dailyReminder')
    .timeBased().atHour(22).nearMinute(0).everyDays(1).inTimezone(TZ)
    .create();
  ScriptApp.newTrigger('morningReminder')
    .timeBased().atHour(8).nearMinute(0).everyDays(1).inTimezone(TZ)
    .create();
  ScriptApp.newTrigger('voteReminder')
    .timeBased().atHour(17).nearMinute(0).everyDays(1).inTimezone(TZ)
    .create();
  ScriptApp.newTrigger('warmBootCache').timeBased().everyMinutes(5).create();
  warmBootCache();
  toast_('毎朝8時ごろ・17時ごろ（シフト希望）・毎晩22時ごろのお知らせと、表示を速くする準備（5分おき）を設定しました');
}

function removeTriggers() {
  ScriptApp.getProjectTriggers()
    .filter(t => ['dailyReminder', 'morningReminder', 'voteReminder', 'warmBootCache'].indexOf(t.getHandlerFunction()) >= 0)
    .forEach(t => ScriptApp.deleteTrigger(t));
}

/* =========================================================
 * 画面から呼ぶ API（google.script.run）
 * ========================================================= */

function getAppData() {
  const t = today_();
  const st = getSettings_();
  return {
    today: t,
    thisMonth: monthOf_(t),
    nextMonth: addMonths_(monthOf_(t), 1),
    members: members_(true),
    notify: {
      slack: !!st.SLACK_WEBHOOK_URL,
      line: !!st.LINE_CHANNEL_ACCESS_TOKEN,
      lineUsers: Object.keys(lineUsers_()),
      lineAdmin: st.LINE_ADMIN || '',
    },
    workHours: WORK_HOURS,
    manager: st.MANAGER_NAME || '',
    allowance: Number(st.MANAGER_ALLOWANCE_H || 3),
    version: APP_VERSION,
    needs: { id: st.NEEDS_ID || NEEDS_LOGIN.id, pw: st.NEEDS_PW || NEEDS_LOGIN.pw },
    masterKinds: MASTER_KINDS.map(k => ({ key: k.key, label: k.label, nameLabel: k.nameLabel, user: k.user })),
  };
}

/** 起動時に1回だけ呼ぶ：前月・今月・来月をまとめて返す */
function getBootstrap() {
  // サーバー側の一時保存があればそれを返す（スプレッドシートを読まない分だけ速い）
  const key = bootKey_();
  try {
    const hit = CacheService.getScriptCache().get(key);
    if (hit) return JSON.parse(hit);
  } catch (e) { console.warn(e.message); }
  const app = buildBootstrap_();
  try { CacheService.getScriptCache().put(key, JSON.stringify(app), BOOT_CACHE_SEC); } catch (e) { console.warn(e.message); }
  return app;
}
function bootKey_() { return 'boot:' + APP_VERSION + ':' + today_(); }
/** 5分おきのトリガー：一時保存を作り直しておく（スプレッドシートを直接直した分も5分以内に反映） */
function warmBootCache() {
  let app;
  try {
    app = buildBootstrap_();
    CacheService.getScriptCache().put(bootKey_(), JSON.stringify(app), BOOT_CACHE_SEC);
  } catch (e) { console.warn(e.message); return; }
  try { writeSnapshot_(app); } catch (e) { console.warn('表示用スナップショット：' + e.message); }
}

/* ---------- 表示用スナップショット（初めて開くときに Apps Script を待たないため） ----------
 * 最初の画面のデータ（個人情報・ID/PW を除く）を、専用の小さなスプレッドシートに JSON で書いておく。
 * 画面は Google スプレッドシートの読み出し（gviz）から1秒以内に取り、まず表示する。
 * そのあと、いつもどおり Apps Script から最新を取って差し替える。
 * 作り方：エディタで setupSnapshot を1回実行 → ログに出る ID を index.html の SNAPSHOT_ID に入れる。
 */
const SNAPSHOT_SHEET = 'snapshot';
const SNAPSHOT_CHUNK = 20000;

function setupSnapshot() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty('SNAPSHOT_SS_ID');
  let ss = null;
  if (id) { try { ss = SpreadsheetApp.openById(id); } catch (e) { ss = null; } }
  if (!ss) {
    ss = SpreadsheetApp.create('パトロールシフト_表示用（自動作成・編集しないでください）');
    id = ss.getId();
    props.setProperty('SNAPSHOT_SS_ID', id);
    ss.getSheets()[0].setName(SNAPSHOT_SHEET);
    if (PDF_FOLDER_ID) { try { DriveApp.getFileById(id).moveTo(DriveApp.getFolderById(PDF_FOLDER_ID)); } catch (e) { console.warn(e.message); } }
  }
  // リンクを知っている人は閲覧のみ（中身は最初の画面のデータだけ。個人情報・ID/PW は入れない）
  DriveApp.getFileById(id).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  props.deleteProperty('SNAPSHOT_HASH');
  warmBootCache();
  const msg = 'index.html の SNAPSHOT_ID に、次の ID を入れてください：' + id;
  console.log(msg);
  toast_(msg);
  return id;
}

function writeSnapshot_(app) {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('SNAPSHOT_SS_ID');
  if (!id) return;
  const pub = JSON.parse(JSON.stringify(app));
  delete pub.needs;
  pub.members = (pub.members || []).map(m => ({ name: m.name, active: m.active, baseCount: m.baseCount }));
  pub.snapshotAt = now_();
  const body = JSON.stringify(Object.assign({}, pub, { snapshotAt: '' }));
  const hash = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, body, Utilities.Charset.UTF_8));
  if (props.getProperty('SNAPSHOT_HASH') === hash) return;   // 変わっていなければ書かない
  const json = JSON.stringify(pub);
  const rows = [];
  for (let i = 0; i < json.length; i += SNAPSHOT_CHUNK) rows.push([json.slice(i, i + SNAPSHOT_CHUNK)]);
  const sh = SpreadsheetApp.openById(id).getSheetByName(SNAPSHOT_SHEET) || SpreadsheetApp.openById(id).getSheets()[0];
  sh.clearContents();
  sh.getRange(1, 1, rows.length, 1).setNumberFormat('@').setValues(rows);
  props.setProperty('SNAPSHOT_HASH', hash);
}
function buildBootstrap_() {
  const app = getAppData();
  app.months = {};
  [-1, 0, 1].forEach(k => {
    const ym = addMonths_(app.thisMonth, k);
    app.months[ym] = monthData_(ym);
  });
  return app;
}

function getMonthData(ym) {
  assertYm_(ym);
  return monthData_(ym);
}

/* ---------- メンバー ---------- */

function addMember(name) {
  name = String(name || '').trim();
  if (!name) throw new Error('名前を入力してください');
  return withLock_(() => {
    const s = sh_(SH.MEMBERS);
    const rows = readRows_(SH.MEMBERS);
    const i = rows.findIndex(r => r[0].trim() === name);
    if (i >= 0) writeCell_(s, i + 2, 2, 'TRUE');
    else appendRow_(s, [name, 'TRUE', '', now_()]);
    return getAppData();
  });
}

function setMemberActive(name, active) {
  return withLock_(() => {
    const s = sh_(SH.MEMBERS);
    const i = readRows_(SH.MEMBERS).findIndex(r => r[0].trim() === name);
    if (i < 0) throw new Error('メンバーが見つかりません');
    writeCell_(s, i + 2, 2, active ? 'TRUE' : 'FALSE');
    return getAppData();
  });
}

/* ---------- 投票 ---------- */

function saveVote(ym, name, answers) {
  assertYm_(ym);
  if (!members_().some(m => m.name === name)) throw new Error('名前を選んでください');
  const n = daysInMonth_(ym);
  if (!Array.isArray(answers) || answers.length !== n) throw new Error('回答の日数が合いません');
  const clean = answers.map(v => (MARKS.indexOf(v) >= 0 ? v : ''));

  return withLock_(() => {
    if (monthSetting_(ym).status === STATUS.FIXED) {
      throw new Error('この月のシフトは確定済みのため、投票は変更できません');
    }
    const s = sh_(SH.VOTES);
    const i = readRows_(SH.VOTES).findIndex(r => r[0] === ym && r[1] === name);
    const vals = [ym, name, now_()].concat(range_(31).map(k => clean[k] || ''));
    if (i >= 0) writeRow_(s, i + 2, vals);
    else appendRow_(s, vals);

    // 全員の回答がそろったら、自動でシフト下書きを作る
    let autoGenerated = false;
    const md = monthData_(ym);
    if (md.members.length && md.unanswered.length === 0 && md.setting.status === STATUS.NONE) {
      generate_(ym);
      autoGenerated = true;
    }
    const out = monthData_(ym);
    out.autoGenerated = autoGenerated;
    return out;
  });
}

function setDeadline(ym, deadline) {
  assertYm_(ym);
  deadline = String(deadline || '');
  if (deadline && !isDate_(deadline)) throw new Error('締切日の形式が不正です');
  return withLock_(() => {
    updateMonthSetting_(ym, { deadline });
    return monthData_(ym);
  });
}

/* ---------- シフト ---------- */

function generateShift(ym) {
  assertYm_(ym);
  return withLock_(() => {
    generate_(ym);
    return monthData_(ym);
  });
}

/** 手動調整した結果を保存（確定前のみ） */
function saveShift(ym, map) {
  assertYm_(ym);
  return withLock_(() => {
    const cur = monthSetting_(ym);
    if (cur.status === STATUS.FIXED) throw new Error('確定済みです。編集するには先に「確定を解除」してください');
    const valid = new Set(members_(true).map(m => m.name));
    const clean = {};
    datesOf_(ym).forEach(d => {
      const v = String((map || {})[d] || '');
      if (v && !valid.has(v)) throw new Error('不明なメンバー：' + v);
      clean[d] = v;
    });
    writeShifts_(ym, clean);
    if (cur.status === STATUS.NONE) updateMonthSetting_(ym, { status: STATUS.DRAFT });
    return monthData_(ym);
  });
}

function setShiftStatus(ym, status) {
  assertYm_(ym);
  if (status !== STATUS.DRAFT && status !== STATUS.FIXED) throw new Error('状態の指定が不正です');
  const md = withLock_(() => {
    const cur = shifts_(ym);
    if (status === STATUS.FIXED && !Object.keys(cur).some(d => cur[d])) {
      throw new Error('シフトがまだありません');
    }
    updateMonthSetting_(ym, { status });
    return monthData_(ym);
  });
  if (status === STATUS.FIXED) {
    const counts = countBy_(md.dates.map(d => md.shifts[d]).filter(Boolean));
    const summary = Object.keys(counts).map(k => `${k} ${counts[k]}回`).join(' / ');
    safeNotify_(() => `【${ymLabel_(ym)}のシフトが確定しました】\n${summary}${appUrl_()}`);
    // LINE：一人ひとりに、自分の担当日を送る
    try { notifyFixed_(ym, md); } catch (e) { console.error(e); }
  }
  return md;
}

/** シフト確定のお知らせ（LINE）。参加中の全員に、その人の担当日を1通ずつ */
function notifyFixed_(ym, md) {
  const mine = {};
  md.dates.forEach(d => { const w = md.shifts[d]; if (w) (mine[w] = mine[w] || []).push(md_(d)); });
  const names = members_(false).map(m => m.name);
  names.forEach(n => {
    const ds = mine[n] || [];
    lineSendTo_([n],
      `【${ymLabel_(ym)}のシフトが確定しました】\n` +
      (ds.length ? `${n}さんの担当（${ds.length}回）：\n${ds.join('、')}` : `${n}さんの担当はありません。`) +
      `\n担当日の朝8時ごろにも、お知らせが届きます。${appUrl_()}`);
  });
}

/* ---------- 完了報告 ---------- */

function reportDone(date, name) {
  if (!isDate_(date)) throw new Error('日付が不正です');
  if (date > today_()) throw new Error('未来の日付は報告できません');
  if (!members_(true).some(m => m.name === name)) throw new Error('実施者を選んでください');
  const res = withLock_(() => {
    writePatrol_(date, 4, [name, now_()]);
    return monthData_(monthOf_(date));
  });
  // 完了の連絡（LINE）：担当者・実施者・管理者へ
  try {
    const who = shifts_(monthOf_(date))[date] || '';
    lineSendTo_([who, name, getSettings_().LINE_ADMIN],
      `【完了】${md_(date)} のパトロールが完了しました。\n実施：${name}さん` +
      (who && who !== name ? `（担当：${who}さん）` : ''));
  } catch (e) { console.error(e); }
  return res;
}

function cancelReport(date) {
  if (!isDate_(date)) throw new Error('日付が不正です');
  return withLock_(() => {
    writePatrol_(date, 4, ['', '']);
    return monthData_(monthOf_(date));
  });
}

/* ---------- 管理者（名目） ---------- */

function setManager(name) {
  name = String(name || '');
  if (name && !members_(true).some(m => m.name === name)) throw new Error('メンバーが見つかりません');
  return withLock_(() => {
    setSetting_('MANAGER_NAME', name);
    return getAppData();
  });
}

/* ---------- LINE 通知の設定 ---------- */

function setLineAdmin(name) {
  name = String(name || '');
  if (name && !members_(true).some(m => m.name === name)) throw new Error('メンバーが見つかりません');
  return withLock_(() => {
    setSetting_('LINE_ADMIN', name);
    return getAppData();
  });
}

/** 今月の LINE 送信数（無料プランは月200通まで） */
function lineQuota() {
  const st = getSettings_();
  if (!st.LINE_CHANNEL_ACCESS_TOKEN) throw new Error('LINE が設定されていません');
  const get = path => {
    const r = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/' + path, {
      headers: { Authorization: 'Bearer ' + st.LINE_CHANNEL_ACCESS_TOKEN }, muteHttpExceptions: true,
    });
    if (r.getResponseCode() >= 300) throw new Error('LINE エラー：' + r.getContentText());
    return JSON.parse(r.getContentText());
  };
  const q = get('quota'), c = get('quota/consumption');
  return { used: c.totalUsage || 0, limit: q.type === 'limited' ? q.value : null };
}

/* ---------- 活動実績（スプシ＋PDF） ---------- */

function reportSummary_(ym) {
  const dates = datesOf_(ym);
  const reps = reports_(ym);
  const st = getSettings_();
  const all = members_(true);
  const lastDate = dates[dates.length - 1];

  // 月内の回数
  const month = {};
  dates.forEach(d => { if (reps[d]) month[reps[d].name] = (month[reps[d].name] || 0) + 1; });

  // トータル = 導入前の累計 + この月の末日までの完了報告
  const totals = totals_(lastDate);

  // 表に載せる人：参加中のメンバー＋その月に実績がある人（メンバー一覧の順）
  const names = all.filter(m => m.active || month[m.name]).map(m => m.name);
  Object.keys(month).forEach(n => { if (names.indexOf(n) < 0) names.push(n); });

  const manager = st.MANAGER_NAME || '';
  const allowance = Number(st.MANAGER_ALLOWANCE_H || 3);
  return {
    ym,
    workHours: WORK_HOURS,
    manager,
    allowance,
    days: dates.map(d => ({ date: d, name: reps[d] ? reps[d].name : '' })),
    members: names.map(n => ({
      name: n,
      count: month[n] || 0,
      hours: (month[n] || 0) * WORK_HOURS,
      allowance: n === manager ? allowance : '',
      total: totals[n] || 0,
    })),
    missing: dates.filter(d => !reps[d] && d <= today_()),
  };
}

/**
 * 書式（いただいた例に準拠）
 *  左  : 月日 | 曜日 | 担当 | 作業時間（h）   … パトロールがない日は担当・作業時間を空欄＋黄色
 *  右上: ○月担当回数 | 回数 | 作業時間 | 手当（管理者のみ）
 *  右下: トータル担当回数 | 回数 | 数値
 */
function exportReport(ym) {
  assertYm_(ym);
  return withLock_(() => {
    const ss = ssPatrol_();
    const sum = reportSummary_(ym);
    const mNum = Number(ym.slice(5));

    const name = '実績_' + ym;
    let s = ss.getSheetByName(name);
    if (s) s.clear();
    else s = ss.insertSheet(name, ss.getSheets().length);
    s.setHiddenGridlines(true);

    // 左：日別
    const n = sum.days.length;
    s.getRange('C1:D1').setValues([['担当', '作業時間（h）']]);
    s.getRange(2, 1, n, 2).setNumberFormat('@');
    s.getRange(2, 1, n, 4).setValues(sum.days.map(d => [
      `${mNum}/${Number(d.date.slice(8))}`, wd_(d.date), d.name, d.name ? sum.workHours : '',
    ]));
    s.getRange(2, 1, n, 4).setBorder(true, true, true, true, true, true).setFontSize(12);
    s.getRange(2, 2, n, 1).setHorizontalAlignment('center');
    s.getRange(2, 3, n, 1).setFontWeight('bold');
    s.getRange(2, 4, n, 1).setNumberFormat('0.00');
    s.getRange(2, 3, n, 2).setBackgrounds(sum.days.map(d => (d.name ? [null, null] : ['#ffff00', '#ffff00'])));

    // 右上：その月の担当回数
    const k = sum.members.length;
    s.getRange('F2').setValue(`${mNum}月担当回数`);
    s.getRange('H2:I2').setValues([['作業時間', '手当']]);
    if (k) {
      s.getRange(3, 6, k, 4).setValues(sum.members.map(m => [m.name, m.count, m.hours, m.allowance]));
      s.getRange(3, 6, k, 4).setBorder(true, true, true, true, true, true).setFontSize(12);
      s.getRange(3, 6, k, 1).setBackground('#e2efda').setFontWeight('bold').setHorizontalAlignment('center');
      s.getRange(3, 7, k, 1).setNumberFormat('0"回"').setHorizontalAlignment('center');

      // 右下：トータル担当回数
      const t = 3 + k;
      s.getRange(t, 6).setValue('トータル担当回数');
      s.getRange(t + 1, 6, k, 3).setValues(sum.members.map(m => [m.name, m.total, m.total]));
      s.getRange(t + 1, 6, k, 3).setBorder(true, true, true, true, true, true).setFontSize(12);
      s.getRange(t + 1, 6, k, 1).setBackground('#e2efda').setFontWeight('bold').setHorizontalAlignment('center');
      s.getRange(t + 1, 7, k, 1).setNumberFormat('0"回"').setHorizontalAlignment('center');
    }
    [90, 45, 90, 105, 24, 90, 75, 75, 55].forEach((w, i) => s.setColumnWidth(i + 1, w));
    SpreadsheetApp.flush();

    const fileName = `活動実績_${ym}.pdf`;
    const blob = pdfOf_(ss, s, fileName);
    const folder = folder_();
    const old = folder.getFilesByName(fileName);
    while (old.hasNext()) old.next().setTrashed(true);
    const file = folder.createFile(blob);

    return {
      fileName,
      pdfBase64: Utilities.base64Encode(blob.getBytes()),
      pdfUrl: file.getUrl(),
      sheetUrl: `${ss.getUrl()}#gid=${s.getSheetId()}`,
      count: sum.days.filter(d => d.name).length,
      missing: sum.missing,
    };
  });
}

/* ---------- 業務報告書（パトロール以外の業務） ---------- */

function saveProfile(name, p) {
  p = p || {};
  if (!members_(true).some(m => m.name === name)) throw new Error('メンバーが見つかりません');
  return withLock_(() => {
    const s = ensureSheet_(SH.PROFILE, PROFILE_HEADER);
    const i = readRowsOpt_(SH.PROFILE).findIndex(r => r[0].trim() === name);
    const vals = [name, String(p.fullName || '').trim(), String(p.kana || '').trim(),
      String(p.tel || '').trim(), String(p.duty || '').trim(), now_()];
    if (i >= 0) writeRow_(s, i + 2, vals);
    else appendRow_(s, vals);
    return getAppData();
  });
}

function addWork(name, date, hours, detail) {
  if (!isDate_(date)) throw new Error('日付を選んでください');
  if (!members_(true).some(m => m.name === name)) throw new Error('報告者を選んでください');
  hours = Number(hours);
  if (!(hours > 0 && hours <= 24)) throw new Error('作業時間は 0〜24 の数字で入力してください');
  detail = String(detail || '').trim();
  if (!detail) throw new Error('作業詳細を入力してください');
  const ym = monthOf_(date);
  return withLock_(() => {
    if ((work_(ym)[name] || []).length >= WORK_ROWS) {
      throw new Error(`業務報告書1枚に書けるのは${WORK_ROWS}行までです`);
    }
    const s = ensureSheet_(SH.WORK, WORK_HEADER);
    appendRow_(s, [Utilities.getUuid().slice(0, 8), ym, name, date, String(hours), detail, now_()]);
    return monthData_(ym);
  });
}

function deleteWork(id) {
  return withLock_(() => {
    const rows = readRowsOpt_(SH.WORK);
    const i = rows.findIndex(r => r[0] === id);
    if (i < 0) throw new Error('その記録は見つかりません（すでに削除されています）');
    const ym = rows[i][1];
    sh_(SH.WORK).deleteRow(i + 2);
    dirty_();
    return monthData_(ym);
  });
}

function saveFare(ym, name, fare, days) {
  assertYm_(ym);
  fare = fare === '' || fare == null ? '' : Number(fare);
  days = days === '' || days == null ? '' : Number(days);
  if (fare !== '' && !(fare >= 0)) throw new Error('往復運賃は数字で入力してください');
  if (days !== '' && !(days >= 0 && days <= 31)) throw new Error('日数は 0〜31 の数字で入力してください');
  return withLock_(() => {
    const s = ensureSheet_(SH.FARE, FARE_HEADER);
    const i = readRowsOpt_(SH.FARE).findIndex(r => r[0] === ym && r[1] === name);
    const vals = [ym, name, String(fare), String(days), now_()];
    if (i >= 0) writeRow_(s, i + 2, vals);
    else appendRow_(s, vals);
    return monthData_(ym);
  });
}

/** テンプレートを複製して記入し、Excel と PDF を作る */
function exportWorkReport(ym, name) {
  assertYm_(ym);
  return withLock_(() => {
    const ss = ssWork_();
    const tpl = ss.getSheetByName(WORK_TEMPLATE);
    if (!tpl) {
      throw new Error(`業務報告関連のスプレッドシートに「${WORK_TEMPLATE}」シートがありません。` +
        'セットアップ手順の「業務報告書テンプレートの取り込み」を行ってください。');
    }
    const m = members_(true).find(x => x.name === name);
    if (!m) throw new Error('報告者を選んでください');
    const items = work_(ym)[name] || [];
    if (!items.length) throw new Error(`${ymLabel_(ym)}の作業記録がまだありません`);
    if (items.length > WORK_ROWS) throw new Error(`1枚に書けるのは${WORK_ROWS}行までです`);
    const fare = fares_(ym)[name] || {};
    const title = `業務報告書_${ym}_${m.fullName || name}`;

    const tmp = SpreadsheetApp.create(title);
    try {
      const sh = tpl.copyTo(tmp).setName(`${Number(ym.slice(5))}月`);
      tmp.getSheets().forEach(x => { if (x.getSheetId() !== sh.getSheetId()) tmp.deleteSheet(x); });
      fillWorkSheet_(sh, ym, m, items, fare);
      SpreadsheetApp.flush();

      const base = `https://docs.google.com/spreadsheets/d/${tmp.getId()}/export`;
      const xlsx = fetchExport_(`${base}?format=xlsx`).setName(title + '.xlsx');
      const pdf = fetchExport_(`${base}?format=pdf&gid=${sh.getSheetId()}` +
        `&range=${encodeURIComponent(WORK_PRINT_RANGE)}&size=A4&portrait=true&fitw=true` +
        '&gridlines=false&printtitle=false&sheetnames=false&pagenum=UNDEFINED&horizontal_alignment=CENTER' +
        '&top_margin=0.5&bottom_margin=0.5&left_margin=0.5&right_margin=0.5').setName(title + '.pdf');

      const folder = subFolder_('業務報告書');
      const save = blob => {
        const old = folder.getFilesByName(blob.getName());
        while (old.hasNext()) old.next().setTrashed(true);
        return folder.createFile(blob);
      };
      const fx = save(xlsx);
      const fp = save(pdf);
      return {
        title,
        xlsx: { name: xlsx.getName(), base64: Utilities.base64Encode(xlsx.getBytes()), url: fx.getUrl() },
        pdf: { name: pdf.getName(), base64: Utilities.base64Encode(pdf.getBytes()), url: fp.getUrl() },
        folderUrl: folder.getUrl(),
      };
    } finally {
      try { DriveApp.getFileById(tmp.getId()).setTrashed(true); } catch (e) { console.warn(e); }
    }
  });
}

/** テンプレートのセル位置に合わせて記入する */
function fillWorkSheet_(sh, ym, m, items, fare) {
  if (m.duty) sh.getRange('I4').setValue(m.duty);           // 業務内容（空ならテンプレートの文言のまま）
  sh.getRange('E8').setValue(m.kana || '');                // フリガナ
  sh.getRange('E9').setValue(m.fullName || m.name);         // 氏名
  sh.getRange('V8').setValue(m.tel || '');                 // 日中連絡先TEL
  sh.getRange('E11:H11').setValues([ym.slice(0, 4).split('').map(Number)]);  // 年（1桁ずつ）
  sh.getRange('J11:K11').setValues([ym.slice(5, 7).split('').map(Number)]);  // 月（1桁ずつ）

  const r0 = WORK_FIRST_ROW;
  sh.getRange(r0, 2, WORK_ROWS, 3).clearContent();          // B:D（日・曜・作業時間）
  sh.getRange(r0, 8, WORK_ROWS, 1).clearContent();          // H（作業詳細）
  if (items.length) {
    sh.getRange(r0, 3, items.length, 1).setNumberFormat('@');
    sh.getRange(r0, 2, items.length, 3).setValues(items.map(it => [Number(it.date.slice(8)), wd_(it.date), it.hours]));
    sh.getRange(r0, 8, items.length, 1).setValues(items.map(it => [it.detail]));
  }

  // 実費交通費：往復(L45 円) ＊ Q45 日 ＝ W45
  if (fare.fare !== '' && fare.fare != null) {
    const days = fare.days !== '' && fare.days != null ? fare.days : new Set(items.map(it => it.date)).size;
    sh.getRange('L45').setValue(fare.fare);
    sh.getRange('Q45').setValue(days);
    sh.getRange('W45').setValue(fare.fare * days).setNumberFormat('#,##0');
  } else {
    sh.getRangeList(['L45', 'Q45', 'W45']).clearContent();
  }
}

function work_(ym) {
  const o = {};
  readRowsOpt_(SH.WORK).forEach(r => {
    if (r[1] !== ym || !r[2] || !isDate_(r[3])) return;
    (o[r[2]] = o[r[2]] || []).push({ id: r[0], date: r[3], hours: Number(r[4]) || 0, detail: r[5], at: r[6] });
  });
  Object.keys(o).forEach(k => o[k].sort((a, b) =>
    (a.date < b.date ? -1 : a.date > b.date ? 1 : a.at < b.at ? -1 : a.at > b.at ? 1 : 0)));
  return o;
}

function fares_(ym) {
  const o = {};
  readRowsOpt_(SH.FARE).forEach(r => {
    if (r[0] !== ym || !r[1]) return;
    o[r[1]] = { fare: r[2] === '' ? '' : Number(r[2]), days: r[3] === '' ? '' : Number(r[3]) };
  });
  return o;
}

function fetchExport_(url) {
  const res = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true,
  });
  if (res.getResponseCode() !== 200) throw new Error('ファイルの作成に失敗しました（' + res.getResponseCode() + '）。少し待ってもう一度お試しください');
  return res.getBlob();
}

function subFolder_(name) {
  const parent = folder_();
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

/* ---------- パトロールマスタデータ（入力タブ） ---------- */

/** 列記号（A, Z, AA, FE …）→ 列番号 */
function colNum_(letters) {
  return String(letters).toUpperCase().split('').reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
}

function masterKind_(key) {
  const k = MASTER_KINDS.filter(x => x.key === key)[0];
  if (!k) throw new Error('種類を選んでください');
  return k;
}

function masterSheet_(k) {
  if (!MASTER_DATA_SS_ID) throw new Error('Code.gs の MASTER_DATA_SS_ID が設定されていません');
  const s = openSs_(MASTER_DATA_SS_ID).getSheetByName(k.sheet);
  if (!s) throw new Error(`パトロールマスタデータに「${k.sheet}」シートがありません`);
  return s;
}

/** 今日サイトから入力したもの（取り消し用。日付が変わると空になる） */
function masterLog_() {
  let o = null;
  try { o = JSON.parse(PropertiesService.getScriptProperties().getProperty('MASTER_TODAY') || 'null'); } catch (e) { o = null; }
  if (!o || o.date !== today_()) o = { date: today_(), items: [] };
  return o;
}
function saveMasterLog_(o) {
  PropertiesService.getScriptProperties().setProperty('MASTER_TODAY', JSON.stringify(o));
}

/** 種類ごとに、シートの一番下（最新）の記録を返す（二重入力の防止用） */
function getMasterLatest() {
  if (!MASTER_DATA_SS_ID) throw new Error('Code.gs の MASTER_DATA_SS_ID が設定されていません');
  const ss = openSs_(MASTER_DATA_SS_ID);
  const kinds = MASTER_KINDS.filter(k => ss.getSheetByName(k.sheet));
  const ranges = [];
  kinds.forEach(k => {
    ranges.push(`'${k.sheet}'!B${MASTER_FIRST_ROW}:${k.user ? 'C' : 'E'}`);
    ranges.push(`'${k.sheet}'!${k.dateCol}${MASTER_FIRST_ROW}:${k.staffCol}`);
  });
  let data = null;
  try {
    if (typeof Sheets !== 'undefined' && ranges.length) {
      const res = Sheets.Spreadsheets.Values.batchGet(ss.getId(), { ranges });
      data = res.valueRanges.map(v => v.values || []);
    }
  } catch (e) {
    console.warn('batchGet に失敗したので通常の読み込みに切り替えます: ' + e.message);
  }
  if (!data) {
    data = [];
    kinds.forEach(k => {
      const s = ss.getSheetByName(k.sheet);
      const n = Math.max(s.getLastRow() - MASTER_FIRST_ROW + 1, 0);
      data.push(n ? s.getRange(MASTER_FIRST_ROW, 2, n, k.user ? 2 : 4).getDisplayValues() : []);
      data.push(n ? s.getRange(MASTER_FIRST_ROW, colNum_(k.dateCol), n, 2).getDisplayValues() : []);
    });
  }
  const out = {};
  kinds.forEach((k, i) => {
    const main = data[i * 2], ds = data[i * 2 + 1];
    const nameIdx = k.user ? 1 : 2;
    let last = -1;
    main.forEach((r, j) => { if (String(r[nameIdx] || '').trim()) last = j; });
    if (last < 0) { out[k.key] = null; return; }
    const r = main[last], d = ds[last] || [];
    out[k.key] = {
      row: MASTER_FIRST_ROW + last,
      seq: String(r[0] || '').trim(),
      no: k.user ? '' : String(r[1] || '').trim(),
      name: String(r[nameIdx] || '').trim(),
      poster: k.user ? '' : String(r[3] || '').trim(),
      date: String(d[0] || '').trim(),
      staff: String(d[1] || '').trim(),
    };
  });
  return out;
}

/** 入力タブ用：今日の入力と最新の記録を1回で返す（通信回数を減らす） */
function getMasterInfo() {
  let latest = null;
  try { latest = getMasterLatest(); } catch (e) { console.warn(e.message); }
  return { today: getMasterToday(), latest };
}

function getMasterToday() {
  return masterLog_().items.map(it => Object.assign({ label: masterKind_(it.kind).label }, it));
}

/** 1件追加：最後に埋まっている行のすぐ下へ */
function addMaster(kindKey, no, name, poster, date, staff) {
  const k = masterKind_(kindKey);
  no = String(no == null ? '' : no).trim().replace(/[０-９]/g, c => String.fromCharCode(c.charCodeAt(0) - 0xFEE0));
  name = String(name || '').trim();
  poster = String(poster || '').trim();
  if (!name) throw new Error(k.nameLabel + 'を入力してください');
  date = String(date || today_());
  if (!isDate_(date)) throw new Error('日付を選んでください');
  staff = String(staff || '').trim();
  if (!staff || !members_(true).some(m => m.name === staff)) throw new Error('担当者を選んでください');
  if (!k.user) {
    if (!/^\d+$/.test(no)) throw new Error('No. は数字で入力してください');
    if (!poster) throw new Error('投稿者を入力してください');
  }
  return withLock_(() => {
    const s = masterSheet_(k);
    const width = k.user ? 2 : 4;                         // B:C または B:E
    const last = Math.max(s.getLastRow(), MASTER_FIRST_ROW - 1);
    const rows = last >= MASTER_FIRST_ROW
      ? s.getRange(MASTER_FIRST_ROW, 2, last - MASTER_FIRST_ROW + 1, width).getDisplayValues() : [];
    const nameIdx = k.user ? 1 : 2;

    if (!k.user && rows.some(r => String(r[1]).trim() === no)) {
      throw new Error(`No.${no} は【${k.label}】に登録済みです`);
    }
    let lastFilled = -1;
    rows.forEach((r, i) => { if (String(r[nameIdx]).trim()) lastFilled = i; });
    const idx = lastFilled + 1;                           // rows 内の位置
    const row = MASTER_FIRST_ROW + idx;

    // 項番：すでに振ってあればそのまま、なければ直前の項番＋1
    const pre = rows[idx] ? String(rows[idx][0]).trim() : '';
    let seq = pre;
    let wroteSeq = false;
    if (!seq) {
      const prev = lastFilled >= 0 ? Number(String(rows[lastFilled][0]).replace(/[^\d]/g, '')) : 0;
      seq = String((prev || 0) + 1);
      wroteSeq = true;
      if (row > MASTER_FIRST_ROW) {                        // 罫線などの書式を上の行からコピー
        s.getRange(row - 1, 2, 1, width).copyFormatToRange(s, 2, 1 + width, row, row);
        s.getRange(row, 2, 1, width).setBackground(null);
      }
    }
    const vals = k.user ? [[Number(seq), name]] : [[Number(seq), Number(no), name, poster]];
    s.getRange(row, 2, 1, width).setValues(vals);

    // 日付・担当者・報告の有無
    const dc = colNum_(k.dateCol), sc = colNum_(k.staffCol), rc = colNum_(k.reportCol);
    if (wroteSeq && row > MASTER_FIRST_ROW) {
      [dc, sc, rc].forEach(c => {
        s.getRange(row - 1, c).copyFormatToRange(s, c, c, row, row);
        s.getRange(row, c).setBackground(null);
      });
    }
    const [y, mo, d] = date.split('-').map(Number);
    // Date オブジェクトで書くと、スプレッドシート側のタイムゾーンが日本以外のとき
    // 1日前にずれるため、文字列で書いてシートに日付として解釈させる
    s.getRange(row, dc).setNumberFormat('yyyy/m/d').setValue(y + '/' + mo + '/' + d);
    s.getRange(row, sc).setValue(staff);
    s.getRange(row, rc).setValue(k.reportText).setFontColor('#ff0000');

    const log = masterLog_();
    log.items.push({ id: Utilities.getUuid().slice(0, 8), kind: k.key, row, seq, no: k.user ? '' : no, name, poster, date, staff, wroteSeq, at: now_() });
    saveMasterLog_(log);
    return { added: { label: k.label, seq, row }, today: getMasterToday() };
  });
}

/** 今日入力したものの取り消し（そのシートで一番最後の行のときだけ） */
function undoMaster(id) {
  return withLock_(() => {
    const log = masterLog_();
    const it = log.items.filter(x => x.id === id)[0];
    if (!it) throw new Error('取り消せるのは今日入力したものだけです');
    const k = masterKind_(it.kind);
    const s = masterSheet_(k);
    const width = k.user ? 2 : 4;
    const nameIdx = k.user ? 1 : 2;
    const cur = s.getRange(it.row, 2, 1, width).getDisplayValues()[0];
    if (String(cur[nameIdx]).trim() !== it.name) {
      throw new Error('スプレッドシートの内容が変わっているため取り消せません。スプレッドシートで直接直してください');
    }
    const below = s.getLastRow() > it.row
      ? s.getRange(it.row + 1, 2 + nameIdx, s.getLastRow() - it.row, 1).getDisplayValues().some(r => String(r[0]).trim()) : false;
    if (below) throw new Error('この後に別の追加があるため取り消せません。スプレッドシートで直接直してください');
    if (it.wroteSeq) s.getRange(it.row, 2, 1, width).clearContent();
    else s.getRange(it.row, 3, 1, width - 1).clearContent();   // 振ってあった項番は残す
    [k.dateCol, k.staffCol, k.reportCol].forEach(c => s.getRange(it.row, colNum_(c)).clearContent());
    log.items = log.items.filter(x => x.id !== id);
    saveMasterLog_(log);
    return getMasterToday();
  });
}

/* ---------- 通知テスト ---------- */

function testNotify(name) {
  const r = notify_(() => '【テスト】パトロールシフトからの通知テストです');
  if (name && getSettings_().LINE_CHANNEL_ACCESS_TOKEN) {
    const uid = lineUsers_()[name];
    if (!uid) r.errors.push(`${name}さんは LINE 未登録です（公式アカウントを友だち追加して名前を送ってください）`);
    else {
      try { linePush_(uid, `【テスト】${name}さん、パトロールシフトからの LINE 通知テストです`); r.sent.push('LINE（' + name + 'さん）'); }
      catch (e) { r.errors.push('LINE送信エラー：' + e.message); }
    }
  }
  if (!r.sent.length && !r.errors.length) throw new Error('通知先が設定されていません（「設定」シートを確認してください）');
  if (r.errors.length) throw new Error(r.errors.join('\n'));
  return r.sent.join(' / ') + ' に送信しました';
}

/* ---------- LINE（Messaging API）----------
 * 無料プランは月200通まで。グループに送ると人数分数えられるので、必要な人にだけ個別に送る。
 * 返信（Reply）は通数に数えられないので、登録の受け付けは返信で行う。
 */
function lineUsers_() {
  const o = {};
  readRowsOpt_(SH.LINE).forEach(r => { if (r[0] && r[1]) o[r[0].trim()] = r[1].trim(); });
  return o;
}
function lineApi_(path, payload) {
  const token = getSettings_().LINE_CHANNEL_ACCESS_TOKEN;
  if (!token) return null;
  const res = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/' + path, {
    method: 'post', contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + token },
    payload: JSON.stringify(payload), muteHttpExceptions: true,
  });
  if (res.getResponseCode() >= 300) throw new Error(res.getContentText());
  return true;
}
function linePush_(userId, text) {
  // LINE の中のブラウザではなく、ふだんのブラウザ（Safari・Chrome）で開くようにする（前回の内容が残って速く開ける）
  const u = getSettings_().APP_URL || SITE_URL;
  if (u) text = String(text).split(u).join(u + (u.indexOf('?') >= 0 ? '&' : '?') + 'openExternalBrowser=1');
  return lineApi_('push', { to: userId, messages: [{ type: 'text', text: String(text).slice(0, 4900) }] });
}
function lineReply_(token, text) {
  return lineApi_('reply', { replyToken: token, messages: [{ type: 'text', text: String(text).slice(0, 4900) }] });
}
/** 名前のリストに LINE を送る（未登録・空・重複は飛ばす。エラーでも止めない） */
function lineSendTo_(names, text) {
  if (!getSettings_().LINE_CHANNEL_ACCESS_TOKEN) return 0;
  const users = lineUsers_();
  const done = {};
  let n = 0;
  (names || []).forEach(nm => {
    nm = String(nm || '').trim();
    const uid = users[nm];
    if (!nm || !uid || done[uid]) return;
    done[uid] = true;
    try { linePush_(uid, text); n++; } catch (e) { console.error('LINE送信エラー（' + nm + '）：' + e.message); }
  });
  return n;
}
function lineEvent_(ev) {
  const src = ev.source || {};
  if (src.type !== 'user' || !src.userId) return;
  const uid = src.userId;
  if (ev.type === 'unfollow') { lineUnregister_(uid); return; }
  if (ev.type === 'follow') {
    lineReply_(ev.replyToken, '友だち追加ありがとうございます。\nパトロールのお知らせを受け取るには、サイトで使っている自分の名前（例：蔵本）を1回送ってください。');
    return;
  }
  if (ev.type !== 'message' || !ev.message || ev.message.type !== 'text') return;
  const text = String(ev.message.text || '').trim().replace(/^登録[\s　:：]*/, '').replace(/さん$/, '').replace(/\s+/g, '');
  const registered = Object.keys(lineUsers_()).filter(n => lineUsers_()[n] === uid)[0];
  if (text === '解除') {
    lineUnregister_(uid);
    lineReply_(ev.replyToken, '登録を解除しました。もう一度受け取るときは、名前を送ってください。');
    return;
  }
  const m = members_(true).filter(x => x.active).find(x => x.name.replace(/\s+/g, '') === text);
  if (m) {
    withLock_(() => {
      const s = ensureSheet_(SH.LINE, LINE_HEADER);
      const rows = readRowsOpt_(SH.LINE);
      // 同じ名前・同じ LINE の古い登録を消してから追加
      for (let i = rows.length - 1; i >= 0; i--) {
        if (rows[i][0].trim() === m.name || rows[i][1].trim() === uid) s.deleteRow(i + 2);
      }
      dirty_();
      appendRow_(s, [m.name, uid, now_()]);
    });
    lineReply_(ev.replyToken, `${m.name}さんで登録しました。\n担当日の朝8時ごろと、完了報告をしたときに、ここにお知らせが届きます。\n（やめるときは「解除」と送ってください）`);
    return;
  }
  if (!registered) {
    lineReply_(ev.replyToken, '名前が見つかりませんでした。サイトの「あなた」に出てくる名前と同じように送ってください（例：蔵本）。');
  }
}
function lineUnregister_(uid) {
  withLock_(() => {
    const s = ssFor_(SH.LINE).getSheetByName(SH.LINE);
    if (!s) return;
    const rows = readRowsOpt_(SH.LINE);
    for (let i = rows.length - 1; i >= 0; i--) if (rows[i][1].trim() === uid) s.deleteRow(i + 2);
    dirty_();
  });
}

function testNotifyFromMenu() {
  try { toast_(testNotify()); } catch (e) { toast_(e.message); }
}

/* =========================================================
 * 22時リマインド（時間主導トリガーから実行）
 * ========================================================= */

function dailyReminder() {
  const st = getSettings_();
  if (String(st.REMIND_ENABLED || 'TRUE').toUpperCase() === 'FALSE') return;
  const t = today_();
  const ym = monthOf_(t);

  // ① 今日のパトロール完了報告
  if (!reports_(ym)[t]) {
    const who = shifts_(ym)[t];
    safeNotify_(m =>
      `【パトロール未報告】${md_(t)} の完了報告がまだです。\n` +
      `担当：${who ? m(who) : '（未設定）'}\n` +
      `パトロールが終わったら完了報告をお願いします。${appUrl_()}`);
    // LINE は担当者だけに（担当がいない日は管理者へ）
    lineSendTo_([who || st.LINE_ADMIN],
      `${md_(t)} のパトロールの完了報告がまだです。\n` +
      (who ? `担当：${who}さん\n` : '担当者：（未設定）\n') +
      `終わっていたら、サイトで「完了報告」を押してください。${appUrl_()}`);
  }

  // ② 次月の担当可否投票
  const nx = addMonths_(ym, 1);
  const md = monthData_(nx);
  if (!md.members.length || md.setting.status !== STATUS.NONE) return;

  const closing = t >= md.deadline;   // 締切日の夜（22時ごろ）に締め切る
  if (md.unanswered.length === 0 || md.deadlinePassed || closing) {
    // 全員回答済み or 締切（未回答は△扱い）→ 下書きを作る
    try { withLock_(() => generate_(nx, closing)); } catch (e) { console.error(e); }
    return;
  }

  // 投票の告知・リマインドは朝のお知らせ（morningReminder）で、25日・27日に送る
}

/** 毎朝8時ごろ：今日の担当者へのお知らせ */
function morningReminder() {
  const st = getSettings_();
  if (String(st.REMIND_ENABLED || 'TRUE').toUpperCase() === 'FALSE') return;
  const t = today_();
  const ym = monthOf_(t);
  if (monthData_(ym).setting.status !== STATUS.FIXED) return; // シフト確定前は送らない
  const who = shifts_(ym)[t];
  if (!who || reports_(ym)[t]) return;
  lineSendTo_([who],
    `${who}さん、おはようございます。\n今日 ${md_(t)} のパトロール担当です。よろしくお願いします。\n` +
    `終わったら、サイトで「完了報告」を押してください。${appUrl_()}`);
}

/**
 * 毎日17時ごろのトリガー：次月のシフト希望（投票）
 *   告知（VOTE_ANNOUNCE_DAY、既定25日）：参加中で、まだ回答していない人へ（先に出した人には送らない）
 *   リマインド（VOTE_REMIND_DAY、既定は空＝月末の前日）：まだ回答していない人へ
 * 下書きができていたら送らない。
 */
function voteReminder() {
  const st = getSettings_();
  if (String(st.REMIND_ENABLED || 'TRUE').toUpperCase() === 'FALSE') return;
  voteNotice_(st, today_());
}
function voteNotice_(st, t) {
  const day = Number(t.slice(8));
  const announceDay = Number(st.VOTE_ANNOUNCE_DAY || 25);
  const lastDay = daysInMonth_(monthOf_(t));
  const remindDay = Number(st.VOTE_REMIND_DAY || (lastDay - 1));   // 月末の前日（10月なら30日、2月なら27日か28日）
  if (day !== announceDay && day !== remindDay) return;
  const nx = addMonths_(monthOf_(t), 1);
  const md = monthData_(nx);
  if (!md.members.length || md.setting.status !== STATUS.NONE) return;
  const dl = `\n締切：${md_(md.deadline)}（締切後の未回答は△扱い）`;
  if (day === announceDay) {
    safeNotify_(() => `【${ymLabel_(nx)}のシフト希望】\n担当できる日を〇△×で回答してください。${dl}\nサイトの「投票」タブから回答できます。${appUrl_()}`);
    md.unanswered.forEach(n => lineSendTo_([n],
      `${n}さん、${ymLabel_(nx)}のシフト希望の受付を始めました。\n担当できる日を〇△×で回答してください。${dl}\nサイトの「投票」タブから回答できます。${appUrl_()}`));
  } else if (md.unanswered.length) {
    safeNotify_(m => `【${ymLabel_(nx)}のシフト希望 リマインド】\n未回答：${md.unanswered.map(m).join('、')}${dl}\n回答をお願いします。${appUrl_()}`);
    md.unanswered.forEach(n => lineSendTo_([n],
      `${n}さん、${ymLabel_(nx)}のシフト希望がまだ回答されていません。${dl}\nサイトの「投票」タブから回答をお願いします。${appUrl_()}`));
  }
}

/* =========================================================
 * シフト生成
 * ========================================================= */

function generate_(ym, closing) {
  const md = monthData_(ym);
  if (md.setting.status === STATUS.FIXED) {
    throw new Error('確定済みです。作り直す場合は先に「確定を解除」してください');
  }
  if (!md.members.length) throw new Error('メンバーが登録されていません');
  if (md.unanswered.length && !md.deadlinePassed && !closing) {
    throw new Error(
      `まだ回答していない人がいます：${md.unanswered.join('、')}\n` +
      '全員が回答するか、締切日を過ぎると作成できます（締切後の未回答は△扱い）。');
  }
  const avail = availOf_(md, '△');
  const res = buildShift_(md.dates.length, md.members, avail, md.prevLast, 150);
  const map = {};
  md.dates.forEach((d, i) => { map[d] = res.assign[i] || ''; });
  writeShifts_(ym, map);
  updateMonthSetting_(ym, { status: STATUS.DRAFT });

  const blanks = res.assign.filter(x => !x).length;
  safeNotify_(() =>
    `【${ymLabel_(ym)}のシフト】下書きを自動作成しました。` +
    (blanks ? `\n⚠ 担当者がいない日が${blanks}日あります。` : '') +
    `\n確認・調整して「確定」をお願いします。${appUrl_()}`);
  lineSendTo_([getSettings_().LINE_ADMIN],
    `【${ymLabel_(ym)}のシフト】下書きができました。` +
    (blanks ? `\n⚠ 担当者がいない日が${blanks}日あります。` : '') +
    `\n確認して「確定」をお願いします。${appUrl_()}`);
}

/**
 * ルール
 *  1. × の人は入れない
 *  2. 同じ人を2日連続にしない（前月末日との連続も含む）
 *  3. 〇 の人で埋められる日は 〇 を優先、△ はどうしても足りない日だけ
 *  4. その月の中で回数を均等に
 * ランダム性のある貪欲法 + 局所改善を何度か回して、一番良い案を採用する。
 * 評価：空き日数 → △の使用数 → 回数の偏り（二乗和） の順に小さいほど良い
 */
function buildShift_(n, names, avail, prevLast, iterations) {
  const rank = v => (v === '〇' ? 0 : v === '△' ? 1 : 9);
  let best = null;

  for (let it = 0; it < iterations; it++) {
    const a = new Array(n).fill('');
    const c = {};
    names.forEach(nm => { c[nm] = 0; });
    const canPut = (nm, i) => {
      if (rank(avail[nm][i]) === 9) return false;
      const left = i === 0 ? (prevLast || '') : a[i - 1];
      const right = i === n - 1 ? '' : a[i + 1];
      return nm !== left && nm !== right;
    };

    // 候補が少ない日から埋める
    const order = range_(n).map(i => ({
      i,
      o: names.filter(nm => avail[nm][i] === '〇').length,
      t: names.filter(nm => avail[nm][i] === '△').length,
      r: Math.random(),
    })).sort((x, y) => x.o - y.o || x.t - y.t || x.r - y.r).map(x => x.i);

    order.forEach(i => {
      const ok = names.filter(nm => canPut(nm, i));
      if (!ok.length) return;
      const bestRank = Math.min.apply(null, ok.map(nm => rank(avail[nm][i])));
      const pool = ok.filter(nm => rank(avail[nm][i]) === bestRank);
      const mn = Math.min.apply(null, pool.map(nm => c[nm]));
      const top = pool.filter(nm => c[nm] === mn);
      const pick = top[Math.floor(Math.random() * top.length)];
      a[i] = pick;
      c[pick]++;
    });

    // 局所改善：空きを埋める / △→〇 / 多い人→少ない人
    let changed = true;
    let guard = 0;
    while (changed && guard++ < 1000) {
      changed = false;
      for (let i = 0; i < n; i++) {
        const cur = a[i];
        const rc = cur ? rank(avail[cur][i]) : 99;
        for (let k = 0; k < names.length; k++) {
          const nm = names[k];
          if (nm === cur || !canPut(nm, i)) continue;
          const rn = rank(avail[nm][i]);
          if (rn < rc || (rn === rc && c[nm] + 1 < c[cur])) {
            a[i] = nm;
            c[nm]++;
            if (cur) c[cur]--;
            changed = true;
            break;
          }
        }
      }
    }

    const blanks = a.filter(x => !x).length;
    const tri = a.filter((x, i) => x && avail[x][i] === '△').length;
    const sq = names.reduce((s, nm) => s + c[nm] * c[nm], 0);
    const score = blanks * 1e6 + tri * 1e3 + sq;
    if (!best || score < best.score) best = { assign: a.slice(), counts: Object.assign({}, c), score };
  }
  return best;
}

/* =========================================================
 * データ読み書き
 * ========================================================= */

function monthData_(ym) {
  const dates = datesOf_(ym);
  const members = members_().map(m => m.name);
  const setting = monthSetting_(ym);
  const votes = votes_(ym);
  const shifts = setting.status === STATUS.NONE ? {} : shifts_(ym);
  const reports = reports_(ym);
  const unanswered = members.filter(nm => !votes[nm] || votes[nm].answers.some(v => !v));
  // 締切：設定がなければ「前月の末日」（その日の夜22時ごろに締め切って下書きを作る）
  const deadlineAuto = !setting.deadline;
  const deadline = setting.deadline || autoDeadline_(ym);
  const deadlinePassed = today_() > deadline;
  const prevYm = addMonths_(ym, -1);
  const prevDates = datesOf_(prevYm);
  const prevLast = shifts_(prevYm)[prevDates[prevDates.length - 1]] || '';
  const totals = totals_(dates[dates.length - 1]);
  const md = {
    ym, dates, members, setting, votes, shifts, reports, unanswered, deadline, deadlineAuto, deadlinePassed, prevLast, totals,
    work: work_(ym), fares: fares_(ym), today: today_(),
  };
  md.warnings = shiftWarnings_(md);
  return md;
}

/** 締切を設定していない月の締切＝前月の末日（11月分なら10/31） */
function autoDeadline_(ym) {
  const p = addMonths_(ym, -1);
  return `${p}-${String(daysInMonth_(p)).padStart(2, '0')}`;
}

function availOf_(md, fillBlank) {
  const o = {};
  md.members.forEach(nm => {
    const a = (md.votes[nm] && md.votes[nm].answers) || [];
    o[nm] = md.dates.map((_, i) => a[i] || fillBlank);
  });
  return o;
}

function shiftWarnings_(md) {
  if (!md.dates.some(d => md.shifts[d] !== undefined)) return [];
  const av = availOf_(md, '△');
  const w = [];
  md.dates.forEach((d, i) => {
    const who = md.shifts[d];
    if (!who) {
      const anyone = md.members.some(nm => av[nm][i] !== '×');
      w.push({ date: d, level: 'error', text: anyone ? '担当者が決まっていません' : '全員×のため担当者なし' });
      return;
    }
    if (md.members.indexOf(who) < 0) w.push({ date: d, level: 'warn', text: `${who}さんはメンバー一覧で無効になっています` });
    else if (av[who][i] === '×') w.push({ date: d, level: 'error', text: `${who}さんは×の日です` });
    const prev = i === 0 ? md.prevLast : md.shifts[md.dates[i - 1]];
    if (prev && prev === who) w.push({ date: d, level: 'warn', text: `${who}さんが2日連続です` });
  });
  return w;
}

function members_(includeInactive) {
  return readRows_(SH.MEMBERS)
    .filter(r => r[0].trim())
    .map(r => ({
      name: r[0].trim(),
      active: String(r[1]).toUpperCase() !== 'FALSE',
      slackId: (r[2] || '').trim(),
      baseCount: Number(r[4]) || 0,
      fullName: (r[5] || '').trim(),   // 旧バージョン（メンバーシートF〜I列）からの引き継ぎ用
      kana: (r[6] || '').trim(),
      tel: (r[7] || '').trim(),
      duty: r[8] || '',
    }))
    .map(m => {
      const p = profiles_()[m.name];
      return p ? Object.assign(m, p) : m;
    })
    .filter(m => includeInactive || m.active);
}

function profiles_() {
  const o = {};
  readRowsOpt_(SH.PROFILE).forEach(r => {
    if (r[0].trim()) o[r[0].trim()] = { fullName: r[1].trim(), kana: r[2].trim(), tel: r[3].trim(), duty: r[4] };
  });
  return o;
}

function monthSetting_(ym) {
  const r = readRows_(SH.MONTHS).find(r => r[0] === ym);
  return { deadline: r ? r[1] : '', status: (r && r[2]) || STATUS.NONE };
}

function updateMonthSetting_(ym, patch) {
  const s = sh_(SH.MONTHS);
  const i = readRows_(SH.MONTHS).findIndex(r => r[0] === ym);
  const next = Object.assign(monthSetting_(ym), patch);
  const vals = [ym, next.deadline || '', next.status || STATUS.NONE, now_()];
  if (i >= 0) writeRow_(s, i + 2, vals);
  else appendRow_(s, vals);
}

function votes_(ym) {
  const n = daysInMonth_(ym);
  const o = {};
  readRows_(SH.VOTES).forEach(r => {
    if (r[0] !== ym || !r[1]) return;
    o[r[1]] = {
      updatedAt: r[2],
      answers: range_(n).map(i => (MARKS.indexOf(r[3 + i]) >= 0 ? r[3 + i] : '')),
    };
  });
  return o;
}

/* ---------- 月ごとのパトロールシート（シフト＋完了報告） ---------- */

function patrolName_(ym) { return ym + PATROL_SUFFIX; }
function patrolRows_(ym) { return db_().patrol[ym] || []; }

function ensureMonthSheet_(ym) {
  const ss = ssPatrol_();
  let s = ss.getSheetByName(patrolName_(ym));
  if (s) return s;
  const dates = datesOf_(ym);
  // 月の順に並ぶよう、自分より新しい月のシートの手前に入れる
  const sheets = ss.getSheets();
  let pos = sheets.length;
  for (let i = 0; i < sheets.length; i++) {
    const m = sheets[i].getName().match(PATROL_RE);
    if (m && m[1] > ym) { pos = i; break; }
  }
  s = ss.insertSheet(patrolName_(ym), pos);
  s.getRange(1, 1, dates.length + 1, 5).setNumberFormat('@');
  s.getRange(1, 1, 1, 5).setValues([PATROL_HEADER]).setFontWeight('bold').setBackground('#e8eef5');
  s.setFrozenRows(1);
  s.getRange(2, 1, dates.length, 2).setValues(dates.map(d => [d, wd_(d)]));
  s.getRange(2, 2, dates.length, 1).setHorizontalAlignment('center');
  [100, 45, 110, 110, 140].forEach((w, i) => s.setColumnWidth(i + 1, w));
  dirty_();
  return s;
}

/** その日の行の col 列目から values を書く（4=実施者, 3=担当） */
function writePatrol_(date, col, values) {
  const ym = monthOf_(date);
  const s = ensureMonthSheet_(ym);
  let i = patrolRows_(ym).findIndex(r => r[0] === date);
  if (i < 0) i = Number(date.slice(8)) - 1; // 行が見つからなければ日付の位置に書く
  dirty_();
  s.getRange(i + 2, col, 1, values.length).setNumberFormat('@').setValues([values]);
}

function shifts_(ym) {
  const o = {};
  patrolRows_(ym).forEach(r => { if (r[0] && monthOf_(r[0]) === ym) o[r[0]] = r[2]; });
  return o;
}

function writeShifts_(ym, map) {
  const s = ensureMonthSheet_(ym);
  const dates = datesOf_(ym);
  const rows = patrolRows_(ym);
  // 日付の並びが崩れていなければ1回でまとめて書く
  const aligned = dates.every((d, i) => !rows[i] || rows[i][0] === d || rows[i][0] === '');
  dirty_();
  if (aligned) {
    s.getRange(2, 3, dates.length, 1).setNumberFormat('@').setValues(dates.map(d => [map[d] || '']));
  } else {
    dates.forEach(d => writePatrol_(d, 3, [map[d] || '']));
  }
}

function reports_(ym) {
  const o = {};
  patrolRows_(ym).forEach(r => {
    if (r[0] && monthOf_(r[0]) === ym && r[3]) o[r[0]] = { name: r[3], at: r[4] };
  });
  return o;
}

/** トータル担当回数 = 導入前の累計 + lastDate までの完了報告 */
function totals_(lastDate) {
  const t = {};
  members_(true).forEach(m => { t[m.name] = m.baseCount; });
  const p = db_().patrol;
  Object.keys(p).forEach(ym => p[ym].forEach(r => {
    if (r[0] && r[3] && r[0] <= lastDate) t[r[3]] = (t[r[3]] || 0) + 1;
  }));
  return t;
}

function getSettings_() {
  const o = {};
  readRows_(SH.SETTINGS).forEach(r => { if (r[0]) o[r[0].trim()] = String(r[1] || '').trim(); });
  return o;
}

function setSetting_(key, val) {
  const s = sh_(SH.SETTINGS);
  const i = readRows_(SH.SETTINGS).findIndex(r => r[0] === key);
  if (i >= 0) writeCell_(s, i + 2, 2, val);
  else appendRow_(s, [key, val, '']);
}

/* =========================================================
 * 通知（Slack / LINE）
 * ========================================================= */

function appUrl_() {
  const u = getSettings_().APP_URL || SITE_URL;
  return u ? '\n' + u : '';
}

function mention_(channel) {
  const ids = {};
  members_(true).forEach(m => { ids[m.name] = m.slackId; });
  return nm => (channel === 'slack' && ids[nm] ? `<@${ids[nm]}>` : `${nm}さん`);
}

function notify_(build) {
  const st = getSettings_();
  const sent = [];
  const errors = [];

  if (st.SLACK_WEBHOOK_URL) {
    try {
      const res = UrlFetchApp.fetch(st.SLACK_WEBHOOK_URL, {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify({ text: build(mention_('slack')) }),
        muteHttpExceptions: true,
      });
      if (res.getResponseCode() >= 300) throw new Error(res.getContentText());
      sent.push('Slack');
    } catch (e) {
      errors.push('Slack送信エラー：' + e.message);
    }
  }

  return { sent, errors };
}

function safeNotify_(build) {
  try {
    const r = notify_(build);
    r.errors.forEach(e => console.error(e));
    return r.sent;
  } catch (e) {
    console.error(e);
    return [];
  }
}

/* =========================================================
 * PDF
 * ========================================================= */

function pdfOf_(ss, sheet, fileName) {
  const url = `https://docs.google.com/spreadsheets/d/${ss.getId()}/export` +
    `?format=pdf&gid=${sheet.getSheetId()}&size=A4&portrait=true&fitw=true` +
    '&gridlines=false&printtitle=false&sheetnames=false&pagenum=CENTER' +
    '&top_margin=0.6&bottom_margin=0.6&left_margin=0.6&right_margin=0.6';
  const res = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true,
  });
  if (res.getResponseCode() !== 200) throw new Error('PDFの作成に失敗しました（' + res.getResponseCode() + '）');
  return res.getBlob().setName(fileName);
}

function folder_() {
  if (PDF_FOLDER_ID) {
    try { return DriveApp.getFolderById(PDF_FOLDER_ID); }
    catch (e) { console.warn('PDF_FOLDER_ID のフォルダを開けないのでマイドライブに保存します: ' + e.message); }
  }
  const it = DriveApp.getFoldersByName(PDF_FOLDER);
  return it.hasNext() ? it.next() : DriveApp.createFolder(PDF_FOLDER);
}

/* =========================================================
 * 小さな道具
 * ========================================================= */

let SS_ = null;
function ss_() {
  if (SS_) return SS_;
  SS_ = SPREADSHEET_ID ? SpreadsheetApp.openById(SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  if (!SS_) throw new Error('スプレッドシートが見つかりません。Code.gs の SPREADSHEET_ID を設定してください。');
  return SS_;
}

const OPENED_ = {};
function openSs_(id) {
  if (!id) return ss_();
  if (!OPENED_[id]) OPENED_[id] = SpreadsheetApp.openById(id);
  return OPENED_[id];
}
/** AST業務報告（パトロール記録・活動実績） */
function ssPatrol_() { return openSs_(PATROL_SS_ID); }
/** 業務報告関連（パトロール以外） */
function ssWork_() { return openSs_(WORK_SS_ID); }

const WORK_SHEETS_ = ['WORK', 'FARE', 'PROFILE'];
/** シート名から、それが入っているスプレッドシートを返す */
function ssFor_(name) {
  if (WORK_SHEETS_.some(k => SH[k] === name)) return ssWork_();
  return ss_();
}

function sh_(name) {
  const s = ssFor_(name).getSheetByName(name);
  if (!s) throw new Error(`シート「${name}」がありません。Apps Script で setup を実行してください。`);
  return s;
}

/*
 * 読み込みの高速化：
 *  1回の呼び出しの中で、全シートを Sheets API でまとめて1回だけ読み、結果を使い回す。
 *  書き込みをしたら印をつけ、次に読むときだけ読み直す。
 */
let DB_ = null;
const WIDTH_ = 34;

function db_() {
  if (DB_) return DB_;
  DB_ = { patrol: {} };
  const mainNames = Object.keys(SH).filter(k => WORK_SHEETS_.indexOf(k) < 0).map(k => SH[k]);
  const workNames = WORK_SHEETS_.map(k => SH[k]);
  const groups = [
    { ss: ss_(), pick: n => mainNames.indexOf(n) >= 0 || (!PATROL_SS_ID && PATROL_RE.test(n)) || (!WORK_SS_ID && workNames.indexOf(n) >= 0) },
  ];
  if (PATROL_SS_ID) groups.push({ ss: ssPatrol_(), pick: n => PATROL_RE.test(n) });
  if (WORK_SS_ID) groups.push({ ss: ssWork_(), pick: n => workNames.indexOf(n) >= 0 });
  let flushed = false;
  groups.forEach(g => {
    const sheets = g.ss.getSheets().filter(x => g.pick(x.getName()));
    const names = sheets.map(x => x.getName());
    if (!names.length) return;
    let data = null;
    try {
      if (typeof Sheets !== 'undefined') {
        if (!flushed) { SpreadsheetApp.flush(); flushed = true; }
        const res = Sheets.Spreadsheets.Values.batchGet(g.ss.getId(), { ranges: names.map(n => `'${n}'!A2:AH`) });
        data = res.valueRanges.map(v => (v.values || []).map(pad_));
      }
    } catch (e) {
      console.warn('batchGet に失敗したので通常の読み込みに切り替えます: ' + e.message);
      data = null;
    }
    if (!data) data = sheets.map(readSheetDirect_);
    names.forEach((n, i) => {
      const m = n.match(PATROL_RE);
      if (m) DB_.patrol[m[1]] = data[i];
      else DB_[n] = data[i];
    });
  });
  return DB_;
}

function pad_(r) {
  const a = r.map(v => (v == null ? '' : String(v)));
  while (a.length < WIDTH_) a.push('');
  return a;
}

function readSheetDirect_(s) {
  const last = s.getLastRow();
  if (last < 2) return [];
  return s.getRange(2, 1, last - 1, s.getLastColumn()).getDisplayValues().map(pad_);
}

function readRows_(name) {
  const rows = db_()[name];
  if (!rows) sh_(name); // シートがなければ分かりやすいエラーを出す
  return rows || [];
}

let BOOT_DROPPED_ = false;
function dirty_() {
  DB_ = null;
  // 書き込みがあったら、最初の画面の一時保存を捨てる（1回の実行で1度だけ）
  if (!BOOT_DROPPED_) {
    BOOT_DROPPED_ = true;
    try { CacheService.getScriptCache().remove(bootKey_()); } catch (e) {}
  }
}

/** シートがなくても空として読む（後から追加したシート用） */
function readRowsOpt_(name) { return db_()[name] || []; }

/** シートがなければ見出し付きで作る */
function ensureSheet_(name, header) {
  const ss = ssFor_(name);
  let s = ss.getSheetByName(name);
  if (!s) {
    s = ss.insertSheet(name);
    s.getRange(1, 1, s.getMaxRows(), header.length).setNumberFormat('@');
    s.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold').setBackground('#e8eef5');
    s.setFrozenRows(1);
    dirty_();
  }
  return s;
}

function writeRow_(s, row, values) {
  dirty_();
  const r = s.getRange(row, 1, 1, values.length);
  r.setNumberFormat('@');
  r.setValues([values.map(v => (v == null ? '' : String(v)))]);
}

function appendRow_(s, values) { writeRow_(s, s.getLastRow() + 1, values); }

function writeCell_(s, row, col, value) {
  dirty_();
  s.getRange(row, col).setNumberFormat('@').setValue(String(value));
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function toast_(msg) {
  try { ss_().toast(msg, 'パトロール', 6); } catch (e) { console.log(msg); }
}

function range_(n) { return Array.from({ length: n }, (_, i) => i); }
function pad2_(n) { return String(n).padStart(2, '0'); }
function now_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm'); }
function today_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }
function monthOf_(d) { return String(d).slice(0, 7); }
function isDate_(s) { return /^\d{4}-\d{2}-\d{2}$/.test(String(s)); }
function assertYm_(ym) { if (!/^\d{4}-\d{2}$/.test(String(ym))) throw new Error('月の指定が不正です'); }

function addMonths_(ym, n) {
  let [y, m] = ym.split('-').map(Number);
  m += n;
  while (m > 12) { m -= 12; y++; }
  while (m < 1) { m += 12; y--; }
  return `${y}-${pad2_(m)}`;
}

function daysInMonth_(ym) {
  const [y, m] = ym.split('-').map(Number);
  return new Date(Date.UTC(y, m, 0)).getUTCDate();
}

function datesOf_(ym) { return range_(daysInMonth_(ym)).map(i => `${ym}-${pad2_(i + 1)}`); }

function utc_(d) { const [y, m, dd] = d.split('-').map(Number); return Date.UTC(y, m - 1, dd); }
function daysBetween_(a, b) { return Math.round((utc_(b) - utc_(a)) / 864e5); }
function wd_(d) { return '日月火水木金土'[new Date(utc_(d)).getUTCDay()]; }
function md_(d) { const p = d.split('-').map(Number); return `${p[1]}/${p[2]}(${wd_(d)})`; }
function ymLabel_(ym) { const [y, m] = ym.split('-').map(Number); return `${y}年${m}月`; }
function countBy_(arr) { const o = {}; arr.forEach(x => { o[x] = (o[x] || 0) + 1; }); return o; }
