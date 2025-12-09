/**
 * スプレッドシート UI を安全に取得するヘルパー。
 * Webアプリ実行や手動実行など UI の無いコンテキストでは null を返す。
 */
function getUiSafe_() {
  try {
    return SpreadsheetApp.getUi();
  } catch (e) {
    Logger.log('getUiSafe_: このコンテキストでは UI が利用できません: ' + e);
    return null;
  }
}

/**
 * ユーザーのスキル情報を管理する主要なシート名。
 */
const USER_SKILL_SHEET_NAME = 'HQ_UserSkill';

/**
 * ユーザースキルシートの基本ヘッダー名
 */
const USER_ID_HEADER   = 'userId';
const USER_NAME_HEADER = 'displayName';
const TOTAL_STARS_HEADER = 'totalStars';

/**
 * エントリーシート名と xpStatus の定数
 * （HatakeQuest_Entries は 2行目がヘッダー）
 */
const ENTRIES_SHEET_NAME = 'HatakeQuest_Entries';
const ENTRIES_HEADER_ROW = 2;
const XP_STATUS_TODO = 'todo';
const XP_STATUS_DONE = 'done';

/**
 * LINE アクセストークン（スクリプトプロパティから取得）
 * プロジェクト設定 → スクリプトプロパティ に
 *   名前: LINE_ACCESS_TOKEN
 *   値  : 実際のチャネルアクセストークン
 * を保存しておくこと。
 */
const LINE_ACCESS_TOKEN =
  PropertiesService.getScriptProperties().getProperty('LINE_ACCESS_TOKEN');

/**
 * 参加回数からレベル情報を返すヘルパー。
 * 戻り値: { key, jp, en, range }
 */
function getLevelInfoByCount_(count) {
  if (!count || count <= 0) {
    return {
      key: 'none',
      jp: 'レベル未設定',
      en: 'No level',
      range: '0回',
    };
  }

  // ★ 畑クエスト：レベル称号マスタ
  const table = [
    {
      key:   'beginner',
      min:   1,
      max:   2,
      jp:    '🌱 ビギナー（畑見習い）',
      en:    'Beginner / Farm Trainee',
      range: '1〜2回',
    },
    {
      key:   'novice',
      min:   3,
      max:   6,
      jp:    '🪴 見習い農士（ファームノービス）',
      en:    'Farm Novice',
      range: '3〜6回',
    },
    {
      key:   'warrior',
      min:   7,
      max:   12,
      jp:    '🧺 収穫戦士（ハーベストウォリアー）',
      en:    'Harvest Warrior',
      range: '7〜12回',
    },
    {
      key:   'knight',
      min:   13,
      max:   20,
      jp:    '⚔️ 畑騎士（ファームナイト）',
      en:    'Farm Knight',
      range: '13〜20回',
    },
    {
      key:   'mage',
      min:   21,
      max:   30,
      jp:    '🔮 畑魔導士（ファームメイジ）',
      en:    'Farm Mage',
      range: '21〜30回',
    },
    {
      key:   'dark_knight',
      min:   31,
      max:   40,
      jp:    '🌑 畑の黒騎士（ダークナイト・オブ・ザ・フィールド）',
      en:    'Dark Knight of the Field',
      range: '31〜40回',
    },
    {
      key:   'sage',
      min:   41,
      max:   50,
      jp:    '📜 畑の賢者（ファーム・セージ）',
      en:    'Farm Sage',
      range: '41〜50回',
    },
    {
      key:   'paladin',
      min:   51,
      max:   9999, // 実運用想定は 51〜52回だが、それ以上でも聖騎士扱い
      jp:    '🕊️ 畑の聖騎士（ファーム・パラディン）',
      en:    'Farm Paladin',
      range: '51回以上（年間コンプリート目安：52回）',
    },
    // 🛡 畑の守護者（承認制）は自動ロジックの外で扱う
  ];

  for (let i = 0; i < table.length; i++) {
    const lv = table[i];
    if (count >= lv.min && count <= lv.max) {
      return lv;
    }
  }

  // 万一マッチしない場合は一番上位を返す
  return table[table.length - 1];
}

/**
 * HQ_UserSkill から userId 一人分の情報を取得する（レベル通知用）
 * 見つからなければ null を返す
 */
function getUserSkillRecordByUserId_(userId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(USER_SKILL_SHEET_NAME);
  if (!sh) throw new Error('HQ_UserSkill シートがありません。');

  const header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const colMap = {};
  header.forEach(function (h, i) {
    const key = String(h || '').trim();
    if (key) colMap[key] = i + 1;
  });

  const userIdCol      = colMap[USER_ID_HEADER];
  const displayNameCol = colMap[USER_NAME_HEADER];
  const totalStarsCol  = colMap[TOTAL_STARS_HEADER];
  const levelNameCol   = colMap['levelName'];

  if (!userIdCol) throw new Error('HQ_UserSkill に userId 列がありません。');

  const lastRow = sh.getLastRow();
  if (lastRow <= 1) return null;

  const values = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).getValues();

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    if (String(row[userIdCol - 1] || '') === String(userId)) {
      return {
        rowIndex:    i + 2,
        userId:      userId,
        displayName: displayNameCol ? row[displayNameCol - 1] || '' : '',
        totalStars:  totalStarsCol ? row[totalStarsCol - 1] || 0 : 0,
        levelName:   levelNameCol ? row[levelNameCol - 1] || '' : '',
      };
    }
  }
  return null;
}

/**
 * スキル定義。ID（キー）は内部処理用、
 * labelはシートの列ヘッダー名およびサイドバーの表示名と完全に一致させます。
 */
const SKILLS = {
  uneta:         { label: '畝立て (★)',                   maxStars: 20 },
  tanemaki:      { label: '種まき / 植付 (★)',            maxStars: 50 },
  zassotori:     { label: '雑草とり (★)',                 maxStars: 50 },
  mizuyari:      { label: '水やり (★)',                   maxStars: 50 },
  shichu:        { label: '支柱組み (★)',                 maxStars: 30 },
  syukaku_basic: { label: '収穫 (★)',                     maxStars: 50 },
  rope:          { label: 'ロープ結び (★)',               maxStars: 10 },
  shitate:       { label: '仕立て / 誘引 / 剪定 (★)',     maxStars: 40 },
  hatakehelp:    { label: '畑ヘルプ (★)',                 maxStars: 20 },
  osowari:       { label: '教わり / 感謝 (★)',            maxStars: 20 },
  gomi:          { label: 'ゴミ拾い / 整備 (★)',           maxStars: 15 },
  nisshi:        { label: '日誌に記録 (★)',               maxStars: 25 },
  kikai:         { label: '機械操作 (★)',                 maxStars: 30 },
  herb:          { label: 'ハーブ栽培 / 活用 (★)',        maxStars: 30 },
  kusa:          { label: '草・植物観察 (★)',             maxStars: 20 },
  syukaku:       { label: '収穫 → 調理 → 発表 (★)',      maxStars: 80 },
  chokubaisho:   { label: '直売所手伝い (★)',             maxStars: 30 },
  shokai:        { label: '新規お客紹介 (★)',             maxStars: 30 },
  eigyo:         { label: '営業信頼構築 (★)',             maxStars: 10 },
  sns:           { label: 'SNS発信 (★)',                  maxStars: 100 },
  allskill:      { label: '全スキル達成 (★)',             maxStars: 500 },
};

/**
 * 「新スキル出現」LINE通知を出さないスキルID一覧
 */
const SKILL_UNLOCK_NOTIFY_EXCLUDE = [
  'uneta',          // 畝立て
  'tanemaki',       // 種まき / 植付
  'zassotori',      // 雑草とり
  'mizuyari',       // 水やり
  'shichu',         // 支柱組み
  'syukaku_basic',  // 収穫 (基本)
  'shitate',        // 仕立て / 誘引 / 剪定
  'nisshi',         // 日誌に記録
  // 必要になったらここに追加できる！
  // '〇〇', // コメントも書ける
];


/**
 * バッジマスタ定義
 */
const BADGES = {
  // Entry（エントリー）
  une_master: {
    badgeId: 'une_master',
    badgeName: '🌱 畝マスター（リッジ・メイカー / Ridge Maker）',
    skillKey: 'une',
    category: 'entry',
    description: '畝作りを5回以上経験し、任せられると判断された。',
  },
  tane_master: {
    badgeId: 'tane_master',
    badgeName: '🌾 種撒きマスター（ライフ・プランター / Life Planter）',
    skillKey: 'tane',
    category: 'entry',
    description: '種まき・植付を3回以上経験した。',
  },
  kusa_master: {
    badgeId: 'kusa_master',
    badgeName: '🍃 草取りマスター（ウィード・ハンドラー / Weed Handler）',
    skillKey: 'kusa',
    category: 'entry',
    description: '雑草取りを3回以上経験し、畑維持を安心して任せられる。',
  },
  mizu_master: {
    badgeId: 'mizu_master',
    badgeName: '💧 水やりの達人（ウォーター・テイマー / Water Tamer）',
    skillKey: 'mizu',
    category: 'entry',
    description: '水やりの量とタイミングを理解できた。',
  },

  // Basic
  shichu_master: {
    badgeId: 'shichu_master',
    badgeName: '🎋 支柱マスター（ステム・エンジニア / Stem Engineer）',
    skillKey: 'shichu',
    category: 'basic',
    description: '植物の成長に合わせて支柱を組める。',
  },
  rope_master: {
    badgeId: 'rope_master',
    badgeName: '🔥 ロープマスター（ノット・アーティスト / Knot Artist）',
    skillKey: 'rope',
    category: 'basic',
    description: '支柱を結ぶロープ結びを習得した。',
  },
  syukaku_basic_master: {
    badgeId: 'syukaku_basic_master',
    badgeName: '🍅 収穫マスター（ハーベスト・ハンドラー / Harvest Handler）',
    skillKey: 'syukaku',
    category: 'basic',
    description: '作物ごとの収穫方法を理解し実践できる。',
  },

  // Intermediate
  shitate_master: {
    badgeId: 'shitate_master',
    badgeName: '✂️ 剪定師（グリーン・スカルプター / Green Sculptor）',
    skillKey: 'sentei',
    category: 'intermediate',
    description: '仕立て・誘引・剪定の流れを理解している。',
  },
  hatakehelp_master: {
    badgeId: 'hatakehelp_master',
    badgeName: '💪 畑助人（フィールド・ヘルパー / Field Helper）',
    skillKey: 'help',
    category: 'intermediate',
    description: '他の畑作業を手伝えるフィールドヘルパー。',
  },
  osowari_master: {
    badgeId: 'osowari_master',
    badgeName: '🧓 伝承の徒（レガシー・ラーナー / Legacy Learner）',
    skillKey: 'denso',
    category: 'intermediate',
    description: '教わり・感謝の流れを実践できる。',
  },
  gomi_master: {
    badgeId: 'gomi_master',
    badgeName: '🗑️ 掃除の勇者（クリーン・キーパー / Clean Keeper）',
    skillKey: 'souji',
    category: 'intermediate',
    description: '畑や土手のゴミ拾いを継続できる。',
  },

  // Advanced
  nisshi_master: {
    badgeId: 'nisshi_master',
    badgeName: '💬 伝承ノート（メモリー・キーパー / Memory Keeper）',
    skillKey: 'note',
    category: 'advance',
    description: '教わったことを日誌に書き残せる。',
  },
  kikai_master: {
    badgeId: 'kikai_master',
    badgeName: '🛠️ 機械士（メカニック / Mechanic）',
    skillKey: 'kikai',
    category: 'advance',
    description: '重機を安全に扱えるメカニック。',
  },
  herb_master: {
    badgeId: 'herb_master',
    badgeName: '🌿 香草師（ハーバル・クラフター / Herbal Crafter）',
    skillKey: 'herb',
    category: 'advance',
    description: 'ハーブ栽培と活用ができる。',
  },
  kusa_observer: {
    badgeId: 'kusa_observer',
    badgeName: '🍀 草語り（グリーン・リスナー / Green Listener）',
    skillKey: 'kusamira',
    category: 'advance',
    description: '草・植物観察を通じて畑のバランスを読める。',
  },

  // Unique
  market_master: {
    badgeId: 'market_master',
    badgeName: '🧭 マーケッター（マーケット・ナビゲーター / Market Navigator）',
    skillKey: 'market',
    category: 'unique',
    description: '直売所運営を任せられる。',
  },
  customer_master: {
    badgeId: 'customer_master',
    badgeName: '🚶 行商人（トラベル・マーチャント / Travel Merchant）',
    skillKey: 'customer',
    category: 'unique',
    description: '新規お客を連れてくることができる。',
  },
  en_master: {
    badgeId: 'en_master',
    badgeName: '🤝 縁結び（コネクター / Connector）',
    skillKey: 'en',
    category: 'unique',
    description: '信頼構築ができるコネクター。',
  },
  chef_master: {
    badgeId: 'chef_master',
    badgeName: '👩‍🍳 シェフ（ハーベスト・シェフ / Harvest Chef）',
    skillKey: 'chef',
    category: 'unique',
    description: '収穫→調理→発表まで責任を持ってできる。',
  },
  koho_master: {
    badgeId: 'koho_master',
    badgeName: '📢 広報（コミュニティ・リポーター / Community Reporter）',
    skillKey: 'koho',
    category: 'unique',
    description: 'SNS発信で畑の魅力を広める。',
  },

  // Legendary
  legend_allskill: {
    badgeId: 'legend_allskill',
    badgeName: '💫 伝説の農士（レジェンダリー・ファーマー / Legendary Farmer）',
    skillKey: 'legend',
    category: 'legendary',
    description: '全スキルコンプリート。',
  },
};

/**
 * BADGES 定義一覧（シンプル版）
 */
function getBadgeDefinitions() {
  const order = Object.keys(BADGES);
  const badges = {};
  order.forEach((id) => {
    badges[id] = BADGES[id];
  });
  return { order, badges };
}

/**
 * Sidebar 用のバッジ定義
 */
function getBadgeDefinitionsForSidebar() {
  const order = Object.keys(BADGES);
  const list = order.map((id) => {
    const b = BADGES[id];
    return {
      badgeId: id,
      badgeName: b.badgeName,
      description: b.description || '',
      skillKey: b.skillKey || '',
      category: b.category || 'other',
    };
  });
  return { order, list };
}

/**
 * HQ_Badges シートを取得し、なければ作成＆ヘッダーを整える
 * @returns {{sheet: GoogleAppsScript.Spreadsheet.Sheet, colMap: Object}}
 */
function ensureBadgesSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('HQ_Badges');
  if (!sh) {
    sh = ss.insertSheet('HQ_Badges');
  }

  const requiredHeaders = [
    'userId',
    'displayName',
    'badgeName',
    'skillKey',
    'status',
    'updatedAt',
    'updatedBy',
    'timestamp',
    'eventType',
    'msgType',
    'messageId',
    'text',
    'photoUrl',
    'approvedAt',
    'approvedBy',
    'skillStars',
    'skillStarsHistory',
  ];

  const lastRow = sh.getLastRow();
  const lastCol = Math.max(sh.getLastColumn(), requiredHeaders.length);

  let header = [];
  if (lastRow >= 1) {
    header = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  }

  const colMap = {};
  header.forEach((v, i) => {
    const key = String(v || '').trim();
    if (key) colMap[key] = i + 1;
  });

  // 足りないヘッダーを右側に追加
  let col = header.length;
  requiredHeaders.forEach((name) => {
    if (!colMap[name]) {
      col += 1;
      sh.getRange(1, col).setValue(name);
      colMap[name] = col;
    }
  });

  return { sheet: sh, colMap };
}

/**
 * 12桁のランダムな英数字のユーザーIDを生成します。
 */
function generateUserId() {
  const length = 12;
  const chars =
    'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
  let result = '';
  for (let i = 0; i < length; i++) {
    result += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return result;
}

/**
 * 畑クエストから LINE にプッシュ通知を送る共通関数
 */
function sendHatakeQuestNotification_(userId, messages) {
  if (!userId) return;
  if (!LINE_ACCESS_TOKEN) {
    Logger.log('LINE_ACCESS_TOKEN が設定されていません');
    return;
  }

  const url = 'https://api.line.me/v2/bot/message/push';

  const payload = {
    to: userId,
    messages: messages,
  };

  const params = {
    method: 'post',
    headers: {
      'Content-Type': 'application/json; charset=UTF-8',
      Authorization: 'Bearer ' + LINE_ACCESS_TOKEN,
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  };

  const res = UrlFetchApp.fetch(url, params);
  Logger.log(
    'LINE push status: ' +
      res.getResponseCode() +
      ' ' +
      res.getContentText()
  );
}

/**
 * HQ_Profiles から、指定 userId の MYカードURL（token 列）を取得
 */
function getMyCardUrlFromProfiles_(userId) {
  if (!userId) return '';

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('HQ_Profiles');
  if (!sh) return '';

  const lastRow = sh.getLastRow();
  if (lastRow <= 1) return '';

  const header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const colMap = {};
  header.forEach((h, i) => {
    const key = String(h || '').trim();
    if (key) colMap[key] = i + 1;
  });

  const userIdCol = colMap['userId'];
  const tokenCol  = colMap['token'];
  if (!userIdCol || !tokenCol) return '';

  const values = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).getValues();
  for (let i = 0; i < values.length; i++) {
    const rowUserId = String(values[i][userIdCol - 1] || '');
    if (rowUserId === String(userId)) {
      const token = values[i][tokenCol - 1] || '';
      return String(token || '');
    }
  }

  return '';
}


/**
 * レベルアップ時のメッセージを組み立てて送信
 */
function notifyLevelUpToLine_(
  userId,
  displayName,
  beforeLevel,
  afterLevel,
  count
) {
  if (!userId) return;

  let beforeText = '';
  if (beforeLevel && String(beforeLevel).trim() !== '') {
    beforeText = '【前のレベル】' + beforeLevel + '\n';
  } else {
    beforeText = '【前のレベル】（初回レベル認定）\n';
  }

  const text =
    '🌱 畑クエスト レベルアップ！\n\n' +
    (displayName ? displayName + ' さん\n\n' : '') +
    beforeText +
    '【新しいレベル】' +
    afterLevel +
    '\n\n' +
    'これまでの参加回数：' +
    count +
    ' 回\n';

  const messages = [
    {
      type: 'text',
      text: text,
    },
  ];

  sendHatakeQuestNotification_(userId, messages);
}

/**
 * ★ 新スキル出現 通知
 *  - そのスキルの★が 0 → 1 以上になったタイミングで一度だけ飛ぶイメージ
 *
 *  メッセージの並び：
 *  🔍 新スキル出現！「直売所手伝い」
 *
 *  イマMさん
 *
 *  🆕 新しい力が芽吹いた
 *
 *  畑クエストのMYカードをてぇっくしてみてね。
 *  （あればその下に MYカードURL）
 */
function notifyNewSkillToLine_(userId, displayName, skillLabel) {
  if (!userId) return;

  // 名前行（あれば）
  const nameLine = displayName ? displayName + ' さんに' : '';

  // HQ_Profiles から MYカードURL（token 列）を取得
  const cardUrl = getMyCardUrlFromProfiles_(userId);

  // ★ メッセージ本文（君が指定した並び順）
  let text =
    '🔍 新スキル出現！「' + skillLabel + '」\n\n' +
    (nameLine ? nameLine + '\n\n' : '') +
    '🆕 新しい力が芽吹いた\n\n' +
    '畑クエストのMYカードをチェックしてみてね。';

  // カードURLがあれば、最後に追加
  if (cardUrl) {
    text += '\n' + cardUrl;
  }

  const messages = [
    {
      type: 'text',
      text: text,
    },
  ];

  sendHatakeQuestNotification_(userId, messages);
}

/**
 * 🏅 バッジ付与通知（LINE）
 */
function notifyBadgeToLine_(userId, displayName, badgeName, cardUrl) {
  if (!userId) return;

  const text =
    '🏅 新バッジ獲得！\n\n' +
    (displayName ? displayName + ' さん\n\n' : '') +
    '師匠からの認定の証\n' +
    '✨ 「' + badgeName + '」を獲得しました！ ✨\n\n' +
    'おめでとうございます！🎉\n' +
    '畑クエストのMYカードをチェックしてね🌱' +
    (cardUrl ? '\n\n' + cardUrl : '');

  const messages = [
    {
      type: 'text',
      text: text,
    },
  ];

  sendHatakeQuestNotification_(userId, messages);
}


/**
 * スプレッドシートが開かれたときにカスタムメニューを追加します。
 */
function onOpen(e) {
  const ui = getUiSafe_();
  if (!ui) {
    Logger.log(
      'onOpen: UI のないコンテキストのため、カスタムメニューの追加をスキップします。'
    );
    return;
  }

  ui.createMenu('★スキル管理')
    .addItem('★スキル管理サイドバーを開く (付与/登録)', 'showSidebarUnified')
    .addSeparator()
    .addItem('参加回数からレベル再計算', 'recalcParticipationAndLevelsFromEntries')
    .addItem('合計★を再計算する (管理)', 'recalculateTotalStars')
    .addItem('全ユーザーデータをクリアする (管理)', 'clearAllUserData')
    .addToUi();

  ui.createMenu('畑クエスト・ログ管理')
    .addItem('✅ 選択行を承認＋レベル判定', 'approveSelectedRowWithLevelNotify')
    .addItem('（旧）選択行を承認のみ', 'approveSelectedEntryRow')
    .addItem('⭐ 選択行を★付与済みにする（xpStatus=done）', 'markXpDoneSelectedRow')
    .addSeparator()
    .addItem('👥 ユーザー別＋日時で並べ替え', 'sortEntriesByUserAndTime')
    .addSeparator()
    .addItem('👣 参加回数＆レベル再計算', 'recalcParticipationAndLevelsFromEntries')
    .addSeparator()
    .addItem(
      '⚠️ 全データ完全リセット（テスト用）',
      'clearAllHatakeQuestData'
    )
    .addToUi();
}

/**
 * 統合された単一のサイドバーを表示します。
 */
function showSidebarUnified() {
  const html = HtmlService.createTemplateFromFile('SidebarStar')
    .evaluate()
    .setTitle('★スキル管理 (付与/登録)');
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * シートのヘッダー行を読み込み、列名とそのインデックスのマップを作成します。
 */
function getColumnHeaderMap(sheet) {
  const headers = sheet
    .getRange(1, 1, 1, sheet.getLastColumn())
    .getValues()[0];
  const headerMap = {};
  headers.forEach((header, index) => {
    const key = String(header || '').trim();
    if (key) headerMap[key] = index + 1;
  });
  return headerMap;
}

/**
 * HatakeQuest_Entries 用ヘッダーマップ
 */
function getEntriesHeaderMapForHatakeQuest() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(ENTRIES_SHEET_NAME);
  if (!sheet) {
    throw new Error('シート「' + ENTRIES_SHEET_NAME + '」が見つかりません。');
  }

  const headerRowIndex = ENTRIES_HEADER_ROW;
  const headers = sheet
    .getRange(headerRowIndex, 1, 1, sheet.getLastColumn())
    .getValues()[0];
  const map = {};
  headers.forEach((h, i) => {
    const key = String(h || '').trim();
    if (key) {
      map[key] = i + 1;
    }
  });

  return { sheet, headerMap: map, headerRowIndex };
}

/**
 * 指定した userId のユーザー行が HQ_UserSkill に存在するか確認し、
 * なければ新規行を作成する。
 */
function ensureUserRowForUserId(userId, displayName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(USER_SKILL_SHEET_NAME);
  if (!sheet) {
    throw new Error('シート「' + USER_SKILL_SHEET_NAME + '」が見つかりません。');
  }

  const headerMap = getColumnHeaderMap(sheet);
  const userIdCol   = headerMap[USER_ID_HEADER];
  const userNameCol = headerMap[USER_NAME_HEADER];

  if (!userIdCol || !userNameCol) {
    throw new Error('HQ_UserSkill に userId / displayName 列がありません。');
  }

  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();

  // 既存行を探す
  for (let i = 1; i < values.length; i++) {
    if (values[i][userIdCol - 1] === userId) {
      if (displayName && values[i][userNameCol - 1] !== displayName) {
        sheet.getRange(i + 1, userNameCol).setValue(displayName);
      }
      return i + 1;
    }
  }

  // 無ければ新規
  const allHeaders = sheet
    .getRange(1, 1, 1, sheet.getLastColumn())
    .getValues()[0];
  const skillLabels = Object.values(SKILLS).map((s) => s.label);

  const newRow = allHeaders.map((header) => {
    const h = String(header).trim();
    if (h === USER_ID_HEADER) return userId;
    if (h === USER_NAME_HEADER) return displayName || '';

    if (h === TOTAL_STARS_HEADER || skillLabels.includes(h)) {
      return 0;
    }
    return '';
  });

  sheet.appendRow(newRow);
  return sheet.getLastRow();
}

/**
 * サイドバーから送信されたスター付与リクエストを処理します。
 * @param {object} data - { userId, displayName, skillId, starCount }
 */
function processStarFromSidebar(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(USER_SKILL_SHEET_NAME);

    if (!sheet) {
      throw new Error('シート「' + USER_SKILL_SHEET_NAME + '」が見つかりません。');
    }

    const headerMap = getColumnHeaderMap(sheet);

    // ★ SKILLS から今回のスキル定義を取得
    const skillDef = SKILLS[data.skillId];
    if (!skillDef) {
      throw new Error('未知の skillId です: ' + data.skillId);
    }

    // 「収穫 (★)」みたいな列ヘッダー名
    const skillHeader      = skillDef.label;
    const totalStarsHeader = TOTAL_STARS_HEADER;

    const userIdCol   = headerMap[USER_ID_HEADER];
    const userNameCol = headerMap[USER_NAME_HEADER];
    const skillCol    = headerMap[skillHeader];
    const totalCol    = headerMap[totalStarsHeader];

    if (!userIdCol || !skillCol || !totalCol || !userNameCol) {
      throw new Error(
        '必要な列（' +
          USER_ID_HEADER +
          ', ' +
          USER_NAME_HEADER +
          ', ' +
          skillHeader +
          ', ' +
          totalStarsHeader +
          '）のいずれかが見つかりません。'
      );
    }

    // ★ ユーザー行を確実に用意
    const targetRow = ensureUserRowForUserId(data.userId, data.displayName);

    const rowValues = sheet
      .getRange(targetRow, 1, 1, sheet.getLastColumn())
      .getValues()[0];

    // いまのスキル値（付与前）
    const currentSkillValueRaw = rowValues[skillCol - 1];
    const currentSkillValue =
      typeof currentSkillValueRaw === 'number'
        ? currentSkillValueRaw
        : parseInt(currentSkillValueRaw, 10) || 0;

    // ★★ New：ここで「新スキル出現かどうか」を判定
    const wasZeroBefore = currentSkillValue <= 0;

    // 付与後の値
    const newSkillValue = currentSkillValue + data.starCount;
    sheet.getRange(targetRow, skillCol).setValue(newSkillValue);

    // 合計★を更新
    const currentTotalValueRaw = rowValues[totalCol - 1];
    const currentTotalValue =
      typeof currentTotalValueRaw === 'number'
        ? currentTotalValueRaw
        : parseInt(currentTotalValueRaw, 10) || 0;
    const newTotalValue = currentTotalValue + data.starCount;
    sheet.getRange(targetRow, totalCol).setValue(newTotalValue);

    // 名前を最新に
    sheet.getRange(targetRow, userNameCol).setValue(data.displayName);

    // xpStatus=todo → done
    if (typeof markXpStatusDoneForUser === 'function') {
      markXpStatusDoneForUser(data.userId);
    }

    // ★★ New：ここで「初めて★が付いたスキルなら LINE 通知」
    // ただし、除外リストに入っている skillId は通知しない
    const isExcluded =
      Array.isArray(SKILL_UNLOCK_NOTIFY_EXCLUDE) &&
      SKILL_UNLOCK_NOTIFY_EXCLUDE.indexOf(data.skillId) !== -1;

    if (
      wasZeroBefore &&
      newSkillValue > 0 &&
      !isExcluded &&                             // ← ここが「除外は false にする」ポイント
      typeof notifyNewSkillToLine_ === 'function'
    ) {
      // 「営業信頼構築 (★)」 → 「営業信頼構築」 にする
      const skillLabel = String(skillHeader).replace(' (★)', '');
      notifyNewSkillToLine_(data.userId, data.displayName, skillLabel);
    }


    return {
      success: true,
      message:
        data.displayName +
        ' (' +
        data.userId +
        ') に「' +
        skillDef.label +
        '」の★' +
        data.starCount +
        '個を付与しました。',
    };
  } catch (e) {
    Logger.log(
      'processStarFromSidebar 処理中にエラー: ' + e.message + ' Stack: ' + e.stack
    );
    return {
      success: false,
      message: '★付与処理中に致命的なエラーが発生しました: ' + e.message,
    };
  }
}


/**
 * サイドバーから送信されたバッジ付与リクエストを処理する
 * @param {{userId:string, displayName:string, badgeId:string}} data
 */
function processBadgeFromSidebar(data) {
  try {
    if (!data) {
      throw new Error(
        'data が渡されていません。（SidebarStar の google.script.run 呼び出しを確認してください）'
      );
    }

    const userId      = String(data.userId || '').trim();
    const displayName = String(data.displayName || '').trim();
    const badgeId     = String(data.badgeId || '').trim();

    if (!userId || !badgeId) {
      throw new Error('userId または badgeId が空です。');
    }

    const badgeDef = BADGES[badgeId];
    if (!badgeDef) {
      throw new Error('未知のバッジIDです: ' + badgeId);
    }

    const { sheet: sh, colMap } = ensureBadgesSheet_();
    const userIdCol      = colMap['userId'];
    const displayNameCol = colMap['displayName'];
    const badgeNameCol   = colMap['badgeName'];
    const skillKeyCol    = colMap['skillKey'];
    const statusCol      = colMap['status'];
    const updatedAtCol   = colMap['updatedAt'];
    const updatedByCol   = colMap['updatedBy'];

    const lastRow = sh.getLastRow();
    const lastCol = sh.getLastColumn();
    const now       = new Date();
    const updatedBy = Session.getActiveUser().getEmail() || 'manual';
    const tz = Session.getScriptTimeZone() || 'Asia/Tokyo';

    // yyyy-MM-dd で保存
    const nowStr = Utilities.formatDate(now, tz, 'yyyy-MM-dd');

    // 既存の同一 userId + badgeName 行を探す
    let targetRow = 0;
    if (lastRow > 1) {
      const dataValues = sh
        .getRange(2, 1, lastRow - 1, lastCol)
        .getValues();
      for (let i = 0; i < dataValues.length; i++) {
        const row = dataValues[i];
        const rowUserId    = String(row[userIdCol - 1] || '');
        const rowBadgeName = String(row[badgeNameCol - 1] || '');
        if (rowUserId === userId && rowBadgeName === badgeDef.badgeName) {
          targetRow = i + 2;
          break;
        }
      }
    }

    if (!targetRow) {
      targetRow = lastRow + 1;
    }

    if (userIdCol)      sh.getRange(targetRow, userIdCol).setValue(userId);
    if (displayNameCol) sh.getRange(targetRow, displayNameCol).setValue(displayName);
    if (badgeNameCol)   sh.getRange(targetRow, badgeNameCol).setValue(badgeDef.badgeName);
    if (skillKeyCol)    sh.getRange(targetRow, skillKeyCol).setValue(badgeDef.skillKey || '');
    if (statusCol)      sh.getRange(targetRow, statusCol).setValue('active');
    if (updatedAtCol)   sh.getRange(targetRow, updatedAtCol).setValue(nowStr);
    if (updatedByCol)   sh.getRange(targetRow, updatedByCol).setValue(updatedBy);
    
    // 🆕 バッジ付与が完了したので、LINE に通知を送る
    const cardUrl = getMyCardUrlFromProfiles_(userId);
    notifyBadgeToLine_(userId, displayName, badgeDef.badgeName, cardUrl);

    return {
      success: true,
      message:
        displayName +
        ' (' +
        userId +
        ') にバッジ「' +
        badgeDef.badgeName +
        '」を付与 / 更新しました。',
    };

  } catch (e) {
    Logger.log('processBadgeFromSidebar error: ' + e.message + ' / ' + e.stack);
    return {
      success: false,
      message: 'バッジ付与処理中にエラーが発生しました: ' + e.message,
    };
  }
}

/**
 * HatakeQuest_Entries で、指定ユーザーの xpStatus=todo 行を done に更新する。
 */
function markXpStatusDoneForUser(userId) {
  try {
    const info = getEntriesHeaderMapForHatakeQuest();
    const sheet = info.sheet;
    const headerMap = info.headerMap;
    const headerRowIndex = info.headerRowIndex;

    const lastRow = sheet.getLastRow();
    if (lastRow <= headerRowIndex) return;

    const userIdCol   = headerMap['userId'] || headerMap['UserID'] || headerMap['UserId'];
    const xpStatusCol = headerMap['xpStatus'];

    if (!userIdCol || !xpStatusCol) {
      Logger.log(
        'markXpStatusDoneForUser: userId / xpStatus 列が見つからない'
      );
      return;
    }

    const lastCol = sheet.getLastColumn();
    const dataRowCount = lastRow - headerRowIndex;

    const dataRange = sheet.getRange(headerRowIndex + 1, 1, dataRowCount, lastCol);
    const values    = dataRange.getValues();

    const xpRange  = sheet.getRange(headerRowIndex + 1, xpStatusCol, dataRowCount, 1);
    const xpValues = xpRange.getValues();

    let changed = false;

    for (let i = 0; i < values.length; i++) {
      const rowUserId = values[i][userIdCol - 1];
      const curStatus = xpValues[i][0];

      if (rowUserId === userId) {
        const s = (curStatus || '').toString().trim();
        if (s === '' || s === XP_STATUS_TODO) {
          xpValues[i][0] = XP_STATUS_DONE;
          changed = true;
        }
      }
    }

    if (changed) {
      xpRange.setValues(xpValues);
    }
  } catch (e) {
    Logger.log(
      'markXpStatusDoneForUser error: ' + e.message + ' / ' + e.stack
    );
  }
}

/**
 * 新規ユーザーを HQ_UserSkill に登録し、一意のトークン（UserID）を発行します。
 */
function registerNewUser(displayName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(USER_SKILL_SHEET_NAME);

    if (!sheet) {
      throw new Error('シート「' + USER_SKILL_SHEET_NAME + '」が見つかりません。');
    }

    const newUserId = generateUserId();

    const requiredHeaderLabels = [
      USER_ID_HEADER,
      USER_NAME_HEADER,
      ...Object.values(SKILLS).map((s) => s.label),
      TOTAL_STARS_HEADER,
    ];

    const headerMap = getColumnHeaderMap(sheet);

    if (!requiredHeaderLabels.every((label) => headerMap.hasOwnProperty(label))) {
      if (sheet.getLastRow() <= 1) {
        sheet
          .getRange(1, 1, 1, requiredHeaderLabels.length)
          .setValues([requiredHeaderLabels]);
      } else {
        throw new Error(
          '必要なヘッダー列が見つかりません。HQ_UserSkill の1行目を確認してください。'
        );
      }
    }

    const allHeaders = sheet
      .getRange(1, 1, 1, sheet.getLastColumn())
      .getValues()[0];
    const skillLabels = Object.values(SKILLS).map((s) => s.label);

    const newRowData = allHeaders.map((header) => {
      const h = String(header).trim();
      if (h === USER_ID_HEADER) return newUserId;
      if (h === USER_NAME_HEADER) return displayName;

      if (h === TOTAL_STARS_HEADER || skillLabels.includes(h)) {
        return 0;
      }
      return '';
    });

    sheet.appendRow(newRowData);

    return {
      success: true,
      message:
        'ユーザー「' +
        displayName +
        '」を登録しました。IDをコピーしてください。',


      userId: newUserId,
      displayName: displayName,
    };
  } catch (e) {
    Logger.log(
      'registerNewUser 処理中にエラー: ' + e.message + ' Stack: ' + e.stack
    );
    return {
      success: false,
      message: '新規ユーザー登録中にエラーが発生しました: ' + e.message,
    };
  }
}

/**
 * スキル定義と表示順序をサイドバーに返します。
 */
function getSkillDefinitions() {
  return {
    skills: SKILLS,
    order: Object.keys(SKILLS),
  };
}

// ----------------------------------------------------
// 管理メニュー機能
// ----------------------------------------------------

/**
 * 全ユーザーのスキル値からTotal★列を再計算します。
 */
function recalculateTotalStars() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(USER_SKILL_SHEET_NAME);

  if (!sheet) {
    ui.alert('エラー', 'シート「' + USER_SKILL_SHEET_NAME + '」が見つかりません。', ui.ButtonSet.OK);
    return;
  }

  try {
    const headerMap = getColumnHeaderMap(sheet);
    const skillLabels = Object.values(SKILLS).map((s) => s.label);
    const totalCol = headerMap[TOTAL_STARS_HEADER];

    if (!totalCol) {
      throw new Error('合計スター列「' + TOTAL_STARS_HEADER + '」が見つかりません。');
    }

    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    const numRows = values.length;

    if (numRows <= 1) {
      ui.alert('情報', 'ユーザーデータがありません。', ui.ButtonSet.OK);
      return;
    }

    const totalStarUpdates = [];

    for (let i = 1; i < numRows; i++) {
      let totalStars = 0;

      skillLabels.forEach((skillLabel) => {
        const colIndex = headerMap[skillLabel];
        if (colIndex) {
          const value = values[i][colIndex - 1];
          totalStars +=
            typeof value === 'number' && !isNaN(value)
              ? value
              : parseInt(value, 10) || 0;
        }
      });

      totalStarUpdates.push([totalStars]);
    }

    sheet
      .getRange(2, totalCol, totalStarUpdates.length, 1)
      .setValues(totalStarUpdates);

    ui.alert('成功', '合計★の再計算が完了しました。', ui.ButtonSet.OK);
  } catch (e) {
    Logger.log(
      'recalculateTotalStars 処理中にエラー: ' + e.message + ' Stack: ' + e.stack
    );
    ui.alert(
      'エラー',
      '合計★の再計算中にエラーが発生しました: ' + e.message,
      ui.ButtonSet.OK
    );
  }
}

/**
 * ヘッダー行を除く全ユーザーデータをクリアします。
 */
function clearAllUserData() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    '全データクリアの確認',
    '【警告】シート「' +
      USER_SKILL_SHEET_NAME +
      '」のヘッダー行を除く全ユーザーデータをクリアします。続行しますか？',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    ui.alert('キャンセル', 'データクリアはキャンセルされました。', ui.ButtonSet.OK);
    return;
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(USER_SKILL_SHEET_NAME);

    if (!sheet) {
      throw new Error('シート「' + USER_SKILL_SHEET_NAME + '」が見つかりません。');
    }

    const lastRow = sheet.getLastRow();

    if (lastRow > 1) {
      sheet
        .getRange(2, 1, lastRow - 1, sheet.getLastColumn())
        .clearContent();
      ui.alert('完了', '全ユーザーデータのクリアが完了しました。', ui.ButtonSet.OK);
    } else {
      ui.alert('情報', 'クリアするユーザーデータがありませんでした。', ui.ButtonSet.OK);
    }
  } catch (e) {
    Logger.log(
      'clearAllUserData 処理中にエラー: ' + e.message + ' Stack: ' + e.stack
    );
    ui.alert(
      'エラー',
      'データクリア中にエラーが発生しました: ' + e.message,
      ui.ButtonSet.OK
    );
  }
}

/**
 * 畑クエスト関連の全シートの「データ部分」をまるごとクリアする。
 */
function clearAllHatakeQuestData() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const response = ui.alert(
    '⚠️ 畑クエスト全データリセット',
    'HatakeQuest_Entries / HQ_FormEntries / HQ_Profiles / HQ_UserSkill / HQ_Badges / HQ_SkillMap / HatakeQuest_Summary の「ヘッダー以外のデータ」をすべて削除します。\n\nテスト用の初期化です。本当に実行しますか？',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    ui.alert('キャンセル', '全データリセットはキャンセルされました。', ui.ButtonSet.OK);
    return;
  }

  const targetSheets = [
    'HatakeQuest_Entries',
    'HQ_FormEntries',
    'HQ_Profiles',
    'HQ_UserSkill',
    'HQ_Badges',
    'HQ_SkillMap',
    'HatakeQuest_Summary',
  ];

  targetSheets.forEach(function (name) {
    const sh = ss.getSheetByName(name);
    if (!sh) {
      Logger.log(
        'clearAllHatakeQuestData: シート「' + name + '」が見つかりません。スキップします。'
      );
      return;
    }

    const lastRow = sh.getLastRow();
    const lastCol = sh.getLastColumn();
    if (lastRow <= 1) return;

    let dataStartRow = 2;

    if (name === 'HatakeQuest_Entries') {
      dataStartRow =
        (typeof ENTRIES_HEADER_ROW !== 'undefined'
          ? ENTRIES_HEADER_ROW
          : 2) + 1;
    }

    if (lastRow >= dataStartRow) {
      const numRows = lastRow - dataStartRow + 1;
      sh.getRange(dataStartRow, 1, numRows, lastCol).clearContent();
      Logger.log(
        'clearAllHatakeQuestData: シート「' +
          name +
          '」の ' +
          dataStartRow +
          '行目以降をクリアしました。'
      );
    }
  });

  ui.alert('完了', '畑クエスト関連の全データをリセットしました。', ui.ButtonSet.OK);
}

// ----------------------------------------------------
// ウェブアプリケーション用のデータ提供
// ----------------------------------------------------

/**
 * HQ_Badges から userId ごとのバッジ一覧を作る
 */
function getBadgesMapByUserId_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('HQ_Badges');
  if (!sh) return {};

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow <= 1) return {};

  const values  = sh.getRange(1, 1, lastRow, lastCol).getValues();
  const headers = values[0];

  const colMap = {};
  headers.forEach((h, i) => {
    colMap[String(h || '').trim()] = i + 1;
  });

  const userIdCol    = colMap['userId'];
  const badgeNameCol = colMap['badgeName'];
  const skillKeyCol  = colMap['skillKey'];
  const updatedAtCol = colMap['updatedAt'];

  if (!userIdCol || !badgeNameCol) {
    Logger.log('getBadgesMapByUserId_: 必要な列(userId, badgeName)が見つかりません。');
    return {};
  }

  const tz  = Session.getScriptTimeZone() || 'Asia/Tokyo';
  const map = {};

  // 2行目以降
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const uid = String(row[userIdCol - 1] || '');
    if (!uid) continue;

    const badgeName = row[badgeNameCol - 1] || '';
    if (!badgeName) continue;

    const skillKey = skillKeyCol ? (row[skillKeyCol - 1] || '') : '';

    const updatedRaw = updatedAtCol ? row[updatedAtCol - 1] : '';
    let updatedStr = '';

    if (updatedRaw instanceof Date) {
      // Date 型 → yyyy-MM-dd
      updatedStr = Utilities.formatDate(updatedRaw, tz, 'yyyy-MM-dd');
    } else if (updatedRaw) {
      // 文字列 → Date にパースできれば yyyy-MM-dd に揃える
      const parsed = new Date(updatedRaw);
      if (!isNaN(parsed.getTime())) {
        updatedStr = Utilities.formatDate(parsed, tz, 'yyyy-MM-dd');
      } else {
        updatedStr = String(updatedRaw);
      }
    }

    if (!map[uid]) map[uid] = [];

    map[uid].push({
      badgeName: badgeName,
      skillKey:  skillKey,
      updatedAt: updatedStr
    });
  }

  // 同じ badgeName は最新だけに圧縮
  Object.keys(map).forEach((uid) => {
    const list   = map[uid];
    const byName = {};
    list.forEach((b) => { byName[b.badgeName] = b; });
    map[uid] = Object.values(byName);
  });

  return map;
}

/**
 * ウェブアプリケーション用の全ユーザーのスキルデータを取得し、整形して返します。
 * （UserCardHtml から google.script.run で呼ぶ想定）
 * @returns {Array<object>} 整形されたユーザーデータの配列。
 */
/**
 * ウェブアプリケーション用の全ユーザーのスキルデータを取得し、整形して返します。
 * （UserCardHtml から google.script.run で呼ぶ想定）
 * @returns {Array<object>} 整形されたユーザーデータの配列。
 */
function getSkillDataForWebCard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(USER_SKILL_SHEET_NAME);

  if (!sheet || sheet.getLastRow() <= 1) {
    return [];
  }

  try {
    const dataRange = sheet.getDataRange();
    const values    = dataRange.getValues();
    const headerMap = getColumnHeaderMap(sheet);

    const usersData = [];
    const skillIds  = Object.keys(SKILLS);

    // ★ levelName 列（「🌱 ビギナー（畑見習い）」など）の列位置
    const levelNameCol = headerMap['levelName'];

    // ★ HQ_Badges から userId -> badges[] を取得
    const badgesMap = getBadgesMapByUserId_();

    const getRank = (starCount, maxStars) => {
      if (starCount >= maxStars * 0.9) return 'Mythic';
      if (starCount >= maxStars * 0.7) return 'Legendary';
      if (starCount >= maxStars * 0.5) return 'Epic';
      if (starCount >= maxStars * 0.2) return 'Rare';
      return 'Common';
    };

    for (let i = 1; i < values.length; i++) {
      const row = values[i];

      const userId     = row[headerMap[USER_ID_HEADER]   - 1] || ('missing-id-' + i);
      const userName   = row[headerMap[USER_NAME_HEADER] - 1] || ('名無しユーザー' + i);
      const totalStars = row[headerMap[TOTAL_STARS_HEADER] - 1] || 0;

      // ★ レベル名（例：🌱 ビギナー（畑見習い））
      const levelName = levelNameCol ? (row[levelNameCol - 1] || '') : '';

      // ★ このユーザーが持っているバッジ（RAW）
      const rawBadges = badgesMap[userId] || [];

      // skillKey（= SKILLS の id）でバッジを持っているか？
      const hasBadgeForSkill = function (skillKey) {
        return rawBadges.some(function (b) {
          return String(b.skillKey || '') === String(skillKey || '');
        });
      };

      const userSkills = [];

      // スキルごとのデータ
      skillIds.forEach(id => {
        const skillDef = SKILLS[id];
        const colIndex = headerMap[skillDef.label];
        if (colIndex) {
          const starRaw = row[colIndex - 1];
          const starCount = (typeof starRaw === 'number' ? starRaw : parseInt(starRaw) || 0);
          const maxStars = skillDef.maxStars;
          const level = starCount > 0
            ? Math.min(10, Math.ceil(starCount / (maxStars / 10)))
            : 0;
          
          let icon = 'star';
          if (id === 'kikai')   icon = 'zap'; 
          else if (id === 'nisshi')   icon = 'book-open';
          else if (id === 'syukaku')  icon = 'sun';
          else if (id === 'shichu')   icon = 'shield';
          
          if (starCount > 0) {
            userSkills.push({
              title: skillDef.label.replace(' (★)', ''),
              level: level,                 // 並び替え用レベル
              stars: starCount,             // 生の★数
              maxStars: maxStars,           // 将来用
              rank: getRank(starCount, maxStars),
              icon: icon                    // ← ここで icon を使う
            });
          }
        }
      });

      const maxTotalStarsForLevel = 5000;
      const maxLevel              = 100;
      const currentLevel = Math.min(
        maxLevel,
        Math.floor(totalStars / (maxTotalStarsForLevel / maxLevel))
      ) || 1;
      const nextLevelExp = (currentLevel + 1) * (maxTotalStarsForLevel / maxLevel);
      const overallRank  = getRank(totalStars, 1000);

      // ★ バッジ一覧（表示用に整形）
      const userBadges = rawBadges.map(function (b) {
        const master =
          Object.values(BADGES).find(function (m) {
            return String(m.skillKey || '') === String(b.skillKey || '');
          }) || {};

        return {
          badgeId:    master.badgeId    || '',
          badgeName:  master.badgeName  || b.badgeName || '',
          skillKey:   b.skillKey        || master.skillKey || '',
          category:   master.category   || '',
          description: master.description || '',
          updatedAt:  b.updatedAt       || ''
        };
      });

      usersData.push({
        userId:   userId,
        userName: userName,

        // ★ カード上部に出すレベル文言
        title: levelName || (overallRank + ' 土に降り立った者'),

        level:        currentLevel,
        badgeRank:    overallRank,
        currentExp:   totalStars,
        nextLevelExp: nextLevelExp,

        skills: userSkills.sort(function (a, b) {
          return (b.stars || 0) - (a.stars || 0);
        }),
        badges: userBadges,

        items: [],
        stats: {
          power:   Math.floor(totalStars / 5)  + 50,
          defense: Math.floor(totalStars / 8)  + 40,
          speed:   Math.floor(totalStars / 10) + 30,
          magic:   Math.floor(totalStars / 6)  + 50
        }
      });
    }

    return usersData;

  } catch (e) {
    Logger.log('getSkillDataForWebCard 処理中にエラー: ' + e.message + ' Stack: ' + e.stack);
    return { error: true, message: 'データ取得エラー: ' + e.message };
  }
}


/**
 * ウェブアプリケーションのエントリーポイント。
 * （UserCardHtml を表示：React/HTML 側から getSkillDataForWebCard を叩く）
 */
function doGet(e) {
  const htmlTemplate = HtmlService.createTemplateFromFile('UserCardHtml');
  const htmlOutput = htmlTemplate
    .evaluate()
    .setTitle('ユーザープロフィールカード')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  return htmlOutput;
}

/** 選択中の行の status を approved にする（旧・通知なし版） */
function approveSelectedEntryRow() {
  const sh = getEntriesSheet_(); // webhook.gs 側の関数想定
  const row = sh.getActiveRange().getRow();
  if (row <= 1) {
    SpreadsheetApp.getUi().alert('ヘッダー行以外のデータ行を選択してください。');
    return;
  }

  const map = ensureHeaders_(sh);
  const statusCol     = map['status'];
  const approvedAtCol = map['approvedAt'];
  const approvedByCol = map['approvedBy'];

  if (!statusCol) {
    SpreadsheetApp.getUi().alert(
      'status 列が見つかりません。ensureHeaders_ を確認してください。'
    );
    return;
  }

  const now  = now_();
  const user = Session.getActiveUser().getEmail() || 'manual';

  sh.getRange(row, statusCol).setValue('approved');
  if (approvedAtCol) sh.getRange(row, approvedAtCol).setValue(now);
  if (approvedByCol) sh.getRange(row, approvedByCol).setValue(user);

  try {
    recalcParticipationAndLevelsFromEntries();
  } catch (e) {
    Logger.log(
      'approveSelectedEntryRow → recalcParticipationAndLevelsFromEntries エラー: ' +
        e
    );
  }
}

/**
 * ✅ 選択行を承認 + 参加回数＆レベル再計算 + レベルアップ通知
 */
function approveSelectedRowWithLevelNotify() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const entriesSh = ss.getSheetByName(ENTRIES_SHEET_NAME);
  if (!entriesSh) {
    SpreadsheetApp.getUi().alert(
      'シート「' + ENTRIES_SHEET_NAME + '」が見つかりません。'
    );
    return;
  }

  const activeRange = entriesSh.getActiveRange();
  if (!activeRange) {
    SpreadsheetApp.getUi().alert('承認したい行を1つ選択してください。');
    return;
  }

  const row = activeRange.getRow();
  if (row <= ENTRIES_HEADER_ROW) {
    SpreadsheetApp.getUi().alert('ヘッダー行以外のデータ行を選択してください。');
    return;
  }

  const map = ensureHeaders_(entriesSh);
  const userIdCol = map['userId'];
  if (!userIdCol) {
    SpreadsheetApp.getUi().alert(
      'HatakeQuest_Entries に userId 列がありません。'
    );
    return;
  }

  const userId = String(entriesSh.getRange(row, userIdCol).getValue() || '');
  if (!userId) {
    SpreadsheetApp.getUi().alert('この行に userId が入っていません。');
    return;
  }

  const before = getUserSkillRecordByUserId_(userId);

  approveSelectedEntryRow();
  recalcParticipationAndLevelsFromEntries();

  const after = getUserSkillRecordByUserId_(userId);
  if (!after) return;

  const beforeLevel = before ? before.levelName || '' : '';
  const afterLevel  = after.levelName || '';
  const count       = after.totalStars || 0; // 参加回数

  if (beforeLevel !== afterLevel) {
    if (typeof notifyLevelUpToLine_ === 'function') {
      notifyLevelUpToLine_(
        after.userId,
        after.displayName,
        beforeLevel,
        afterLevel,
        count
      );
    }
  }
}

/** 選択中の行の xpStatus を done にする（★付与済みフラグ） */
function markXpDoneSelectedRow() {
  const sh = getEntriesSheet_();
  const row = sh.getActiveRange().getRow();
  if (row <= 1) {
    SpreadsheetApp.getUi().alert('ヘッダー行以外のデータ行を選択してください。');
    return;
  }

  const map = ensureHeaders_(sh);
  const xpCol = map['xpStatus'];
  if (!xpCol) {
    SpreadsheetApp.getUi().alert(
      'xpStatus 列が見つかりません。ensureHeaders_ を確認してください。'
    );
    return;
  }

  sh.getRange(row, xpCol).setValue('done');
}

/** ユーザー別＋日時順に並べ替える（管理しやすくする用） */
function sortEntriesByUserAndTime() {
  const sh = getEntriesSheet_();
  const lastRow = sh.getLastRow();
  if (lastRow <= 1) return;

  const map = ensureHeaders_(sh);

  const timestampCol = map['timestamp'];
  const userIdCol    = map['userId'] || map['displayName'];

  if (!timestampCol) {
    SpreadsheetApp.getUi().alert('timestamp 列が見つかりません。');
    return;
  }

  const range = sh.getRange(ENTRIES_HEADER_ROW + 1, 1, lastRow - ENTRIES_HEADER_ROW, sh.getLastColumn());
  range.sort([
    { column: userIdCol || 3, ascending: true },
    { column: timestampCol,  ascending: false },
  ]);
}

/**
 * HatakeQuest_Entries の "approved" 行からユーザーごとの参加回数を集計し、
 * HQ_UserSkill.totalStars / HQ_UserSkill.levelName を更新する。
 */
function recalcParticipationAndLevelsFromEntries() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const entriesSh = ss.getSheetByName('HatakeQuest_Entries');
  if (!entriesSh) {
    throw new Error('シート「HatakeQuest_Entries」が見つかりません。');
  }

  const headerMap = getHeaderMap_(entriesSh); // webhook.gs 側の関数
  const userIdCol = headerMap.userId;
  const statusCol = headerMap.status;

  if (!userIdCol || !statusCol) {
    throw new Error(
      'HatakeQuest_Entries に userId / status 列が見つかりません。'
    );
  }

  const lastRow = entriesSh.getLastRow();
  if (lastRow <= (typeof ENTRIES_HEADER_ROW !== 'undefined' ? ENTRIES_HEADER_ROW : 2)) {
    Logger.log('データ行がありません。');
    return;
  }

  const startRow = (typeof ENTRIES_HEADER_ROW !== 'undefined'
    ? ENTRIES_HEADER_ROW
    : 2) + 1;
  const numRows  = lastRow - startRow + 1;
  const range    = entriesSh.getRange(startRow, 1, numRows, entriesSh.getLastColumn());
  const values   = range.getValues();

  const userIdIdx = userIdCol - 1;
  const statusIdx = statusCol - 1;

  const counts = {};

  values.forEach((row) => {
    const uid = String(row[userIdIdx] || '');
    if (!uid) return;

    const statusVal = String(row[statusIdx] || '').toLowerCase();
    if (statusVal !== 'approved') return;

    counts[uid] = (counts[uid] || 0) + 1;
  });

  Logger.log('participation counts = ' + JSON.stringify(counts));

  const skillSh = ss.getSheetByName(USER_SKILL_SHEET_NAME);
  if (!skillSh) {
    throw new Error('シート「' + USER_SKILL_SHEET_NAME + '」が見つかりません。');
  }

  let lastColSkill = skillSh.getLastColumn();
  const headerSkill = skillSh
    .getRange(1, 1, 1, lastColSkill)
    .getValues()[0];

  const colMap = {};
  headerSkill.forEach((h, i) => {
    const key = String(h || '').trim();
    if (key) colMap[key] = i + 1;
  });

  const userIdColSkill     = colMap[USER_ID_HEADER]     || colMap['userId'];
  const totalStarsColSkill = colMap[TOTAL_STARS_HEADER] || colMap['totalStars'];

  if (!userIdColSkill || !totalStarsColSkill) {
    throw new Error(
      'HQ_UserSkill に userId / totalStars 列が見つかりません。'
    );
  }

  let levelNameCol = colMap['levelName'];
  if (!levelNameCol) {
    levelNameCol = lastColSkill + 1;
    skillSh.getRange(1, levelNameCol).setValue('levelName');
    lastColSkill = levelNameCol;
  }

  const lastRowSkill = skillSh.getLastRow();
  if (lastRowSkill <= 1) {
    Logger.log('HQ_UserSkill にデータ行がありません。');
    return;
  }

  for (let r = 2; r <= lastRowSkill; r++) {
    const uidCell = skillSh.getRange(r, userIdColSkill);
    const uid     = String(uidCell.getValue() || '');
    if (!uid) continue;

    const count = counts[uid] || 0;

    skillSh.getRange(r, totalStarsColSkill).setValue(count);

    const lv = getLevelInfoByCount_(count);
    skillSh.getRange(r, levelNameCol).setValue(lv.jp);
  }

  Logger.log('recalcParticipationAndLevelsFromEntries 完了');
}
