/***** 畑クエスト Webhook【自動ユーザー登録＋ユーザーカード doGet 版】 *****
 * 必須スクリプトプロパティ：
 *  - CHANNEL_ACCESS_TOKEN
 *  - SHEET_ID
 *  - DRIVE_FOLDER_ID
 *  - FORM_URL
 *  - MYCARD_BASE_URL   ← 追加（Webカードの /exec URL）
 ****************************************************************/

// ★ LINE アクセストークン（ここにまとめて管理）
const CHANNEL_ACCESS_TOKEN =
  "DcjdAbYPCDWRVVAOD0rtHjEgERYNX0rd9Mszk8YvlWblOTaJRaj8MpnxcW6TwnEONv7iS0hClUB1/ndnKJqExnaFQVGZ2S1Qnm13c772bHqWnULbjzTGWXj/wbCPIHkQ4+Wa/F1hYqLIJy4RnuRR+QdB04t89/1O/w1cDnyilFU=";

const PROP = PropertiesService.getScriptProperties();
const TOKEN = (PROP.getProperty("CHANNEL_ACCESS_TOKEN") || "").trim();
const SHEET_ID = (PROP.getProperty("SHEET_ID") || "").trim();
const DRIVE_FOLDER_ID = (PROP.getProperty("DRIVE_FOLDER_ID") || "").trim();
const FORM_URL = (PROP.getProperty("FORM_URL") || "").trim();
const MYCARD_BASE_URL = (PROP.getProperty("MYCARD_BASE_URL") || "").trim();

const SHEET_NAME = "HatakeQuest_Entries";
const PROFILE_SHEET_NAME = "HQ_Profiles";
const ATTACH_WINDOW_MIN = 15;
// HatakeQuest_Entries は 2 行目がヘッダー行
const ENTRIES_HEADER_ROW = 2;

// ★ MYカード用：スキルカテゴリ定義

// スキル名（HQ_UserSkill の列ヘッダー）→ カテゴリキー
const SKILL_CATEGORY = {
  "畝立て (★)": "entry", // une
  "種まき / 植付 (★)": "entry", // tane
  "雑草とり (★)": "entry", // kusa
  "水やり (★)": "entry", // mizu

  "収穫 (★)": "basic", // syukaku
  "支柱組み (★)": "basic", // shichu
  "ロープ結び (★)": "basic", // rope

  "仕立て / 誘引 / 剪定 (★)": "intermediate", // sentei
  "機械操作 (★)": "intermediate", // kikai
  "ハーブ栽培 / 活用 (★)": "intermediate", // herb

  "日誌に記録 (★)": "advance", // note
  "草・植物観察 (★)": "advance", // kusamira
  "直売所手伝い (★)": "advance", // market

  "畑ヘルプ (★)": "epic", // help（特別スキルに昇格）
  "ゴミ拾い / 整備 (★)": "unique", // souji
  "教わり / 感謝 (★)": "unique", // denso
  "収穫 → 調理 → 発表 (★)": "unique", // chef
  "新規お客紹介 (★)": "epic", // customer
  "営業信頼構築 (★)": "epic", // en
  "SNS発信 (★)": "unique", // koho

  "全スキル達成 (★)": "legendary", // legend
};

// カテゴリごとの表示名など
const CATEGORY_INFO = {
  entry: {
    label: "入門スキル（Entry）",
    desc: "畑に触れる最初のステップ。",
  },
  basic: {
    label: "基本スキル（Basic）",
    desc: "一人で安全に動ける基礎技術。",
  },
  intermediate: {
    label: "応用スキル（Intermediate）",
    desc: "状況判断・工夫が求められるスキル。",
  },
  advance: {
    label: "探究スキル（Advanced）",
    desc: "観察・記録・運営など、本質に近づく学び。",
  },
  unique: {
    label: "個性スキル（Unique）",
    desc: "その人らしさが光る行動。",
  },
  epic: {
    label: "特別スキル（Epic）",
    desc: "畑の外へ影響を広げる力。",
  },
  legendary: {
    label: "伝説級（Legendary）",
    desc: "全スキルを制した者だけが到達。",
  },
};

// 表示順
const CATEGORY_ORDER = [
  "entry",
  "basic",
  "intermediate",
  "advance",
  "unique",
  "epic",
  "legendary",
];

// 「常に枠だけは見せる」カテゴリ
// Unique / Epic / Legendary は ★付くまで非表示にしたいので含めない
const ALWAYS_SHOW_CATEGORIES = {
  entry: true,
  basic: true,
  intermediate: true,
  advance: true,
  unique: false,
  epic: false,
  legendary: false,
};

/* ============================= Webhook ============================= */

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const events = body.events || [];

    events.forEach((ev) => {
      if (ev.type !== "message") return;

      const userId = (ev.source && ev.source.userId) || "";
      const m = ev.message || {};

      // まず表示名を取得（text / image 共通で使う）
      const displayName = getLineDisplayName_(userId) || "";

      // ★ HQ_UserSkill / HQ_Profiles を自動登録
      const regResult = autoRegisterUserIfNeeded_(userId, displayName);
      const isNewUser = regResult.isNew;
      const myCardUrl = regResult.myCardUrl;

      /* ---------- テキストメッセージ ---------- */
      if (m.type === "text") {
        const text = String(m.text || "").trim();

        if (text === "ping") {
          replyText_(ev.replyToken, "ok");
          return;
        }

        if (text === "記録") {
          // 基本のフォーム案内メッセージ
          let msg =
            "📒 今日の記録フォームはこちら！\n" +
            (FORM_URL ? FORM_URL : "（FORM_URL未設定）") +
            "\n\n📸 フォーム入力後、写真をこのトークに直接送って下さい。（記録の信憑性を高める為）";

          // 初回ユーザーなら「メンバー認定＋MYカードURL」を頭につける
          if (isNewUser && myCardUrl) {
            msg =
              "🎉 畑クエストメンバー認定！\n" +
              (displayName ? displayName + " さん、ようこそ！\n" : "") +
              "あなたの育成カードはこちら👇\n" +
              myCardUrl +
              "\n\n" +
              msg;
          }

          replyText_(ev.replyToken, msg);
          return; // ログ行は作らない（重複防止）
        }

        replyText_(
          ev.replyToken,
          "了解しました。写真はこのトークに送ってください。"
        );
        return;
      }

      /* ---------- 画像メッセージ ---------- */
      if (m.type === "image") {
        const blob = downloadLineImage_(m.id);
        const photoUrl = saveToDriveAndGetLink_(blob, userId);

        // HatakeQuest_Entries に写真を紐付け
        upsertPhotoToRecentRow_({
          userId,
          displayName,
          messageId: m.id,
          photoUrl,
        });

        // 通常の返信
        let msg = "📸 写真受け取りました！承認後に反映されます。";

        // 初回アクションが画像だった場合も MYカードURL を付ける
        if (isNewUser && myCardUrl) {
          msg =
            "🎉 畑クエストメンバー認定！\n" +
            (displayName ? displayName + " さん、ようこそ！\n" : "") +
            "あなたの育成カードはこちら👇\n" +
            myCardUrl +
            "\n\n" +
            msg;
        }

        replyText_(ev.replyToken, msg);
        return;
      }

      /* ---------- その他種類 ---------- */
      replyText_(ev.replyToken, "このメッセージ種別には対応していません。");
    });

    return ContentService.createTextOutput("OK");
  } catch (err) {
    Logger.log("doPost error: " + err);
    return ContentService.createTextOutput("ERROR");
  }
}

/* ======================= LINE API ヘルパ ======================== */

function replyText_(replyToken, text) {
  const url = "https://api.line.me/v2/bot/message/reply";
  const payload = { replyToken, messages: [{ type: "text", text }] };
  UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json",
    headers: { Authorization: "Bearer " + TOKEN },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });
}

/**
 * 任意タイミングでユーザーにメッセージを送る（push）
 */
function pushTextToUser_(userId, text) {
  if (!userId || !text) return;

  const url = "https://api.line.me/v2/bot/message/push";
  const payload = {
    to: userId,
    messages: [
      {
        type: "text",
        text: text,
      },
    ],
  };

  UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json",
    headers: { Authorization: "Bearer " + TOKEN },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });
}

/**
 * ★ 新スキル出現 通知
 *
 * 例）notifyNewSkillUnlocked_(
 *        'Uxxxxxxxxxx',          // userId
 *        '直売所手伝い',         // 表示したいスキル名
 *        '個性スキル（Unique）'  // カテゴリ名（任意）
 *     );
 */
function notifyNewSkillUnlocked_(userId, skillLabel, categoryLabel) {
  if (!userId || !skillLabel) return;

  const title = "🆕 新しい力が芽吹きました！";
  const line1 = `新スキル出現：『${skillLabel}★』`;
  const line2 = categoryLabel ? `カテゴリ：${categoryLabel}` : "";
  const line3 = "畑での体験を重ねるほど、まだ見ぬスキルが開いていきます🌱";

  const msg = [title, line1, line2, line3].filter(Boolean).join("\n");

  pushTextToUser_(userId, msg);
}

function getLineDisplayName_(userId) {
  if (!userId) return "";
  const url =
    "https://api.line.me/v2/bot/profile/" + encodeURIComponent(userId);
  const res = UrlFetchApp.fetch(url, {
    method: "get",
    headers: { Authorization: "Bearer " + TOKEN },
    muteHttpExceptions: true,
  });
  if (res.getResponseCode() !== 200) return "";
  const obj = JSON.parse(res.getContentText() || "{}");
  return obj.displayName || "";
}

function downloadLineImage_(messageId) {
  const url =
    "https://api-data.line.me/v2/bot/message/" +
    encodeURIComponent(messageId) +
    "/content";
  const res = UrlFetchApp.fetch(url, {
    method: "get",
    headers: { Authorization: "Bearer " + TOKEN },
    muteHttpExceptions: true,
  });
  const blob = res.getBlob();
  blob.setName("photo_" + messageId + ".jpg");
  return blob;
}

/* ========================= Drive 保存 ========================== */

function saveToDriveAndGetLink_(blob, userId) {
  const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  file.setDescription("userId: " + userId);
  return (
    "https://drive.google.com/file/d/" + file.getId() + "/view?usp=drivesdk"
  );
}

/* ==================== スプレッドシート操作 ===================== */

function getEntriesSheet_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) throw new Error("シートが見つかりません: " + SHEET_NAME);
  return sh;
}

/**
 * HatakeQuest_Entries 用のヘッダーを保証し、ヘッダー名→列番号マップを返す。
 */
function ensureHeaders_(sheet) {
  const need = [
    "timestamp",
    "eventType",
    "userId",
    "displayName",
    "msgType",
    "messageId",
    "text",
    "photoUrl",
    // status/approved はどちらか1本あればOK。無い場合は status を作る
    "status",
    "xpStatus",
    "approvedAt",
    "approvedBy",
    "skillKey",
    "skillStars",
    "skillStarsHistory",
  ];

  const headerRow = ENTRIES_HEADER_ROW; // ★ ここが 2 行目
  const lastCol = Math.max(sheet.getLastColumn(), need.length);

  const header = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0];

  const map = {};
  header.forEach((v, i) => {
    if (v) map[String(v).trim()] = i + 1;
  });

  let col = header.length;
  const hasApproved = !!map["approved"];

  need.forEach((name) => {
    if (name === "status" && hasApproved) return;

    if (!map[name]) {
      col += 1;
      sheet.getRange(headerRow, col).setValue(name);
      map[name] = col;
    }
  });

  return map;
}

function getHeaderMap_(sheet) {
  const headerRow = ENTRIES_HEADER_ROW;
  const row = sheet
    .getRange(headerRow, 1, 1, sheet.getLastColumn())
    .getValues()[0];

  const tmp = {};
  row.forEach((v, i) => {
    tmp[String(v || "").trim()] = i + 1;
  });

  const statusCol = tmp["status"] || tmp["approved"];

  return {
    timestamp: tmp["timestamp"],
    eventType: tmp["eventType"],
    userId: tmp["userId"],
    displayName: tmp["displayName"],
    msgType: tmp["msgType"],
    messageId: tmp["messageId"],
    text: tmp["text"],
    photoUrl: tmp["photoUrl"],
    status: statusCol,
    xpStatus: tmp["xpStatus"] || null,
  };
}

function now_() {
  return Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone() || "Asia/Tokyo",
    "yyyy-MM-dd HH:mm:ss"
  );
}

function parseDateLoose_(v) {
  if (!v) return null;
  if (Object.prototype.toString.call(v) === "[object Date]") return v;
  const s = String(v)
    .replace(/[年月日]/g, "/")
    .replace("T", " ")
    .trim();
  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

function appendRow_(o) {
  const sh = getEntriesSheet_();
  const map = ensureHeaders_(sh);
  const lastCol = sh.getLastColumn();

  const row = new Array(lastCol).fill("");

  const nowStr = now_();

  if (map["timestamp"]) row[map["timestamp"] - 1] = nowStr;
  if (map["eventType"]) row[map["eventType"] - 1] = o.eventType || "";
  if (map["userId"]) row[map["userId"] - 1] = o.userId || "";
  if (map["displayName"]) row[map["displayName"] - 1] = o.displayName || "";
  if (map["msgType"]) row[map["msgType"] - 1] = o.msgType || "";
  if (map["messageId"]) row[map["messageId"] - 1] = o.messageId || "";
  if (map["text"]) row[map["text"] - 1] = o.text || "";
  if (map["photoUrl"]) row[map["photoUrl"] - 1] = o.photoUrl || "";
  if (map["status"]) row[map["status"] - 1] = o.status || "pending";
  if (map["source"]) row[map["source"] - 1] = o.source || "";
  if (map["approvedAt"]) row[map["approvedAt"] - 1] = "";
  if (map["approvedBy"]) row[map["approvedBy"] - 1] = "";
  if (map["approvedEmail"]) row[map["approvedEmail"] - 1] = "";
  if (map["skillKey"]) row[map["skillKey"] - 1] = o.skillKey || "";
  if (map["skillStars"]) row[map["skillStars"] - 1] = o.skillStars || "";
  if (map["skillStarsHistory"])
    row[map["skillStarsHistory"] - 1] = o.skillStarsHistory || "";
  if (map["xpStatus"]) row[map["xpStatus"] - 1] = o.xpStatus || "todo";

  sh.appendRow(row);
}

/* ======== 画像を“フォーム行1本”へ紐付け（強化版） ======== */

function _eqLoose_(a, b) {
  const n = (s) =>
    String(s || "")
      .replace(/\s+/g, "")
      .toLowerCase();
  return n(a) === n(b);
}

function upsertPhotoToRecentRow_({ userId, displayName, messageId, photoUrl }) {
  const sh = getEntriesSheet_();
  ensureHeaders_(sh);
  const h = getHeaderMap_(sh);

  const now = new Date();
  const limit = new Date(now.getTime() - ATTACH_WINDOW_MIN * 60 * 1000);

  const last = sh.getLastRow();
  let fallbackRow = 0;
  let fallbackTime = 0;

  for (let r = last; r >= 2; r--) {
    const statusVal = h.status
      ? String(sh.getRange(r, h.status).getValue() || "").toLowerCase()
      : "";
    const consideredPending = h.status ? statusVal !== "approved" : true;
    if (!consideredPending) continue;

    const photo = String(sh.getRange(r, h.photoUrl).getValue() || "");
    if (photo) continue;

    const tsVal = sh.getRange(r, h.timestamp).getValue();
    const t = parseDateLoose_(tsVal);
    if (!t || t < limit) continue;

    const rowUid = String(sh.getRange(r, h.userId).getValue() || "");
    const rowName = String(sh.getRange(r, h.eventType).getValue() || "");

    const uidMatch = rowUid && userId && rowUid === userId;
    const nameMatch = !rowUid && displayName && _eqLoose_(rowName, displayName);

    if (uidMatch || nameMatch) {
      sh.getRange(r, h.msgType).setValue("image");
      sh.getRange(r, h.messageId).setValue(messageId);
      sh.getRange(r, h.photoUrl).setValue(photoUrl);
      if (!rowUid && userId) sh.getRange(r, h.userId).setValue(userId);
      if (sh.getRange(r, h.displayName).getValue() === "" && displayName) {
        sh.getRange(r, h.displayName).setValue(displayName);
      }
      return;
    }

    const tsNum = t.getTime();
    if (tsNum > fallbackTime) {
      fallbackTime = tsNum;
      fallbackRow = r;
    }
  }

  if (fallbackRow > 0) {
    sh.getRange(fallbackRow, h.msgType).setValue("image");
    sh.getRange(fallbackRow, h.messageId).setValue(messageId);
    sh.getRange(fallbackRow, h.photoUrl).setValue(photoUrl);
    if (userId) sh.getRange(fallbackRow, h.userId).setValue(userId);
    if (displayName) {
      const cur = String(
        sh.getRange(fallbackRow, h.displayName).getValue() || ""
      );
      if (!cur) sh.getRange(fallbackRow, h.displayName).setValue(displayName);
    }
    return;
  }

  appendRow_({
    eventType: "",
    userId,
    displayName,
    msgType: "image",
    messageId,
    text: "",
    photoUrl,
    status: "pending",
  });
}

/* ========== 自動ユーザー登録（HQ_UserSkill / HQ_Profiles） ========== */

function autoRegisterUserIfNeeded_(userId, displayName) {
  if (!userId) {
    return { isNew: false, myCardUrl: null };
  }

  const ss = SpreadsheetApp.openById(SHEET_ID);

  /* ① HQ_UserSkill にユーザー行を自動作成・更新 */
  try {
    let skillSh = ss.getSheetByName("HQ_UserSkill");
    if (skillSh) {
      const lastCol = skillSh.getLastColumn();
      if (lastCol > 0) {
        const header = skillSh.getRange(1, 1, 1, lastCol).getValues()[0];
        const colMap = {};
        header.forEach((h, i) => {
          const key = String(h || "").trim();
          if (key) colMap[key] = i + 1;
        });

        const userIdCol = colMap["userId"];
        const nameCol = colMap["displayName"];
        const totalStarsCol = colMap["totalStars"];

        const lastRow = skillSh.getLastRow();
        let foundRow = 0;

        if (userIdCol && lastRow > 1) {
          const ids = skillSh
            .getRange(2, userIdCol, lastRow - 1, 1)
            .getValues();
          for (let i = 0; i < ids.length; i++) {
            if (ids[i][0] === userId) {
              foundRow = i + 2; // 行番号
              break;
            }
          }
        }

        if (foundRow) {
          // 既存 → 名前が変わっていたら更新
          if (nameCol && displayName) {
            const curName = skillSh.getRange(foundRow, nameCol).getValue();
            if (curName !== displayName) {
              skillSh.getRange(foundRow, nameCol).setValue(displayName);
            }
          }
        } else {
          // 新規ユーザー行
          const row = new Array(lastCol).fill("");

          if (userIdCol) row[userIdCol - 1] = userId;
          if (nameCol) row[nameCol - 1] = displayName || "";
          if (totalStarsCol) row[totalStarsCol - 1] = 0;

          // スキル列（★が入っている列）は 0 で初期化
          header.forEach((h, idx) => {
            const name = String(h || "");
            if (name.indexOf("★") >= 0 && !row[idx]) {
              row[idx] = 0;
            }
          });

          skillSh.appendRow(row);
        }
      }
    }
  } catch (e) {
    Logger.log("autoRegisterUserIfNeeded_ HQ_UserSkill error: " + e);
  }

  /* ② HQ_Profiles シート（displayName / userId / token / memberNo / createdAt） */

  let sh = ss.getSheetByName(PROFILE_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(PROFILE_SHEET_NAME);
  }

  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, 5).setValues([
      ["displayName", "userId", "token", "memberNo", "createdAt"],
    ]);
  }

  const header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const colMap = {};
  header.forEach((h, i) => {
    const key = String(h || "").trim();
    if (key) colMap[key] = i + 1;
  });

  const nameCol = colMap["displayName"] || 1;
  const userIdCol = colMap["userId"] || 2;
  const tokenCol = colMap["token"] || 3;
  const memberCol = colMap["memberNo"] || 4;
  const createdCol = colMap["createdAt"] || 5;

  const lastRow = sh.getLastRow();

  let existingRow = 0;
  if (lastRow > 1) {
    const idValues = sh.getRange(2, userIdCol, lastRow - 1, 1).getValues();
    for (let i = 0; i < idValues.length; i++) {
      if (idValues[i][0] === userId) {
        existingRow = i + 2;
        break;
      }
    }
  }

  // ★ ここでカードのベースURLを決める（プロパティが最優先）
  let baseUrl = "";
  if (MYCARD_BASE_URL) {
    baseUrl = MYCARD_BASE_URL.replace(/\/$/, ""); // 末尾の / を削る
  } else {
    const raw = ScriptApp.getService().getUrl();
    baseUrl = raw ? raw.replace(/\/$/, "") : "";
  }

  const myCardUrl = baseUrl
    ? baseUrl + "?uid=" + encodeURIComponent(userId)
    : "";

  const nowStr = now_();

  // 既存ユーザー
  if (existingRow) {
    const currentName = sh.getRange(existingRow, nameCol).getValue();
    let currentToken = String(
      sh.getRange(existingRow, tokenCol).getValue() || ""
    );

    // 名前更新
    if (displayName && currentName !== displayName) {
      sh.getRange(existingRow, nameCol).setValue(displayName);
    }

    // ★ URL が古かったら「必ず」新しいものに上書き
    if (myCardUrl && currentToken !== myCardUrl) {
      sh.getRange(existingRow, tokenCol).setValue(myCardUrl);
      currentToken = myCardUrl;
    }

    const finalUrl = currentToken || myCardUrl || null;

    return {
      isNew: false,
      myCardUrl: finalUrl,
    };
  }

  // 新規ユーザー → 1行追加
  const memberNo = lastRow; // ヘッダ1行を引いた数を利用
  const row = new Array(sh.getLastColumn()).fill("");

  row[nameCol - 1] = displayName || "";
  row[userIdCol - 1] = userId;
  row[tokenCol - 1] = myCardUrl;
  row[memberCol - 1] = memberNo;
  row[createdCol - 1] = nowStr;

  sh.appendRow(row);

  return {
    isNew: true,
    myCardUrl: myCardUrl || null,
  };
}

/* ========== ユーザーカード表示用 doGet ========== */
function doGet(e) {
  try {
    const uid =
      e && e.parameter && e.parameter.uid ? String(e.parameter.uid) : "";

    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sh = ss.getSheetByName("HQ_UserSkill");

    if (!uid || !sh) {
      return HtmlService.createHtmlOutput(
        "<h2>畑クエスト ユーザーカード</h2>" +
          "<p>ユーザーIDが指定されていないか、シートが見つかりません。</p>"
      )
        .setTitle("畑クエスト ユーザーカード")
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    const lastRow = sh.getLastRow();
    const lastCol = sh.getLastColumn();

    if (lastRow < 2) {
      return HtmlService.createHtmlOutput(
        "<h2>畑クエスト ユーザーカード</h2>" +
          "<p>表示できるユーザーがいません。</p>"
      )
        .setTitle("畑クエスト ユーザーカード")
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    const values = sh.getRange(1, 1, lastRow, lastCol).getValues();
    const header = values[0];
    const rows = values.slice(1);

    // ヘッダー → インデックス
    const colMap = {};
    header.forEach((h, i) => {
      const key = String(h || "").trim();
      if (key) colMap[key] = i;
    });

    const userIdIdx = colMap["userId"];
    const displayNameIdx = colMap["displayName"];
    const totalStarsIdx = colMap["totalStars"];
    const levelNameIdx = colMap["levelName"]; // レベル名

    if (userIdIdx == null || displayNameIdx == null) {
      return HtmlService.createHtmlOutput(
        "<h2>畑クエスト ユーザーカード</h2>" +
          "<p>シートのヘッダー設定が不完全です。（userId / displayName が必要）</p>"
      )
        .setTitle("畑クエスト ユーザーカード")
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    // 対象ユーザー行を探す
    let target = null;
    for (let i = 0; i < rows.length; i++) {
      if (rows[i][userIdIdx] === uid) {
        target = rows[i];
        break;
      }
    }

    if (!target) {
      return HtmlService.createHtmlOutput(
        "<h2>畑クエスト ユーザーカード</h2>" +
          "<p>このユーザーのカードはまだ登録されていません。</p>"
      )
        .setTitle("畑クエスト ユーザーカード")
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    const name = target[displayNameIdx] || "ななしさん";
    const total = totalStarsIdx != null ? target[totalStarsIdx] || 0 : 0;
    const levelName =
      levelNameIdx != null
        ? target[levelNameIdx] || "レベル未設定"
        : "レベル未設定";

    /* ---------- スキル一覧（カテゴリ別） ---------- */
    // カテゴリごとに配列を用意
    const grouped = {};
    CATEGORY_ORDER.forEach((k) => {
      grouped[k] = [];
    });

    // マスタ的に全スキルも持っておく（Entry/Basic 全表示用）
    const skillMasterList = [];

    for (let i = 0; i < header.length; i++) {
      const hName = String(header[i] || "");
      if (!hName) continue;
      if (hName.indexOf("★") < 0) continue; // ★がない列はスキル列ではない
      if (i === totalStarsIdx) continue; // totalStars 自身は除外

      const catKey = SKILL_CATEGORY[hName] || "basic"; // 未定義は basic 扱い
      const stars = Number(target[i] || 0);

      skillMasterList.push({
        name: hName,
        catKey: catKey,
        stars: stars,
      });

      if (!grouped[catKey]) grouped[catKey] = [];
      if (stars > 0) {
        grouped[catKey].push({
          name: hName,
          stars: stars,
        });
      }
    }

    // HTML生成（Entry / Basic は★0でも全スキル表示）
    let skillsHtml = "";

    CATEGORY_ORDER.forEach((catKey) => {
      const info = CATEGORY_INFO[catKey];
      if (!info) return;

      let list;
      if (catKey === "entry" || catKey === "basic") {
        // ★の有無に関わらず、そのカテゴリに属するスキルを全部出す
        list = skillMasterList.filter((s) => s.catKey === catKey);
      } else {
        // それ以外のカテゴリは、「★が付いているスキルだけ」表示対象
        list = grouped[catKey] || [];
      }

      const alwaysShow =
        typeof ALWAYS_SHOW_CATEGORIES !== "undefined"
          ? !!ALWAYS_SHOW_CATEGORIES[catKey]
          : false;

      // Unique / Epic / Legendary などは
      // 「alwaysShow=false かつ スキル0件」のときは丸ごと非表示
      if (!alwaysShow && list.length === 0) {
        return;
      }

      skillsHtml += '<div class="skill-cat skill-cat-' + catKey + '">';
      skillsHtml += '<div class="skill-cat-title">' + info.label + "</div>";

      if (info.desc) {
        skillsHtml += '<div class="skill-cat-desc">' + info.desc + "</div>";
      }

      if (list.length === 0) {
        skillsHtml +=
          '<p class="skill-cat-empty">まだこの分野のスキルは芽吹いていません。</p>';
      } else {
        skillsHtml += "<ul>";
        list.forEach((s) => {
          skillsHtml +=
            "<li>" +
            '<span class="skill-name">' +
            s.name +
            "</span>" +
            '：<span class="skill-stars">★' +
            s.stars +
            "</span>" +
            "</li>";
        });
        skillsHtml += "</ul>";
      }

      skillsHtml += "</div>";
    });

    /* ---------- バッジ表示エリア（HQ_Badges を読む） ---------- */
    let badgeHtml = "";
    const badgeSh = ss.getSheetByName("HQ_Badges");
    const userBadges = [];

    if (badgeSh && badgeSh.getLastRow() > 1) {
      const bValues = badgeSh.getDataRange().getValues();
      const bHeader = bValues[0];
      const bRows = bValues.slice(1);

      const bColMap = {};
      bHeader.forEach((h, i) => {
        const key = String(h || "").trim();
        if (key) bColMap[key] = i;
      });

      const bUserIdIdx = bColMap["userId"];
      const bBadgeNameIdx = bColMap["badgeName"];
      const bSkillKeyIdx = bColMap["skillKey"];
      const bStatusIdx = bColMap["status"];
      const bUpdatedAtIdx = bColMap["updatedAt"];

      if (bUserIdIdx != null && bBadgeNameIdx != null) {
        bRows.forEach((row) => {
          if (row[bUserIdIdx] !== uid) return;

          const badgeName = row[bBadgeNameIdx];
          if (!badgeName) return;

          // status があれば "revoked" だけ除外（それ以外は表示）
          let status = bStatusIdx != null ? String(row[bStatusIdx] || "") : "";
          if (status && status.toLowerCase() === "revoked") return;

          userBadges.push({
            badgeName: String(badgeName),
            skillKey:
              bSkillKeyIdx != null ? String(row[bSkillKeyIdx] || "") : "",
            updatedAt: bUpdatedAtIdx != null ? row[bUpdatedAtIdx] : "",
          });
        });
      }
    }

    if (userBadges.length === 0) {
      // バッジがまだ無い場合
      badgeHtml =
        '<div class="badges-title">習得済みスキル（バッジ）</div>' +
        "<p>バッジは、師匠が「任せられる」と判断したときに授与されます。</p>" +
        "<p>まだバッジはありません。</p>";
    } else {
      // バッジがある場合はリスト表示
      badgeHtml =
        '<div class="badges-title">習得済みスキル（バッジ）</div>' +
        "<p>バッジは、師匠が「任せられる」と判断したときに授与されます。</p>" +
        '<ul class="badge-list">';

      userBadges.forEach((b) => {
        const metaParts = [];
        if (b.skillKey) metaParts.push(b.skillKey);
        if (b.updatedAt) metaParts.push(String(b.updatedAt));

        const meta = metaParts.length
          ? "（" + metaParts.join(" / ") + "）"
          : "";

        badgeHtml +=
          '<li class="badge-item">' +
          '<span class="badge-name">' +
          b.badgeName +
          "</span>" +
          (meta ? '<span class="badge-meta">' + meta + "</span>" : "") +
          "</li>";
      });

      badgeHtml += "</ul>";
    }

    /* ---------- 全体HTML ---------- */
    const html =
      "<!DOCTYPE html>" +
      '<html lang="ja">' +
      "<head>" +
      '<meta charset="UTF-8" />' +
      '<meta name="viewport" content="width=device-width, initial-scale=1" />' +
      "<title>畑クエスト ユーザーカード</title>" +
      "<style>" +
      'body { font-family: system-ui, -apple-system, BlinkMacSystemFont, "Yu Gothic", sans-serif; background:#f6f4ec; padding:16px; }' +
      ".card { max-width:480px; margin:0 auto; background:#fff; border-radius:12px; padding:16px 20px; box-shadow:0 2px 8px rgba(0,0,0,0.08); }" +
      ".title { font-size:18px; font-weight:bold; margin-bottom:4px; }" +
      ".subtitle { font-size:13px; color:#666; margin-bottom:12px; }" +
      ".name { font-size:20px; font-weight:bold; margin-bottom:4px; }" +
      ".level { font-size:14px; color:#333; margin-bottom:4px; }" +
      ".total { font-size:14px; color:#444; margin-bottom:12px; }" +
      ".skills-title { font-size:14px; font-weight:bold; margin-bottom:4px; }" +
      ".badges-title { font-size:14px; font-weight:bold; margin-top:16px; margin-bottom:4px; }" +
      "ul { padding-left:20px; margin:4px 0 0; }" +
      "li { font-size:13px; margin:2px 0; }" +
      ".skill-name { font-weight:bold; }" +
      ".skill-stars { color:#e0a800; font-weight:bold; }" +
      ".skill-cat { margin-top:10px; padding-top:8px; border-top:1px solid #eee; }" +
      ".skill-cat-title { font-size:14px; font-weight:bold; margin-bottom:2px; }" +
      ".skill-cat-desc { font-size:12px; color:#666; margin-bottom:4px; }" +
      ".skill-cat-empty { font-size:12px; color:#999; }" +
      ".badge-list { list-style:none; padding-left:0; margin:4px 0 0; }" +
      ".badge-item { font-size:13px; margin:2px 0; }" +
      ".badge-name { font-weight:bold; }" +
      ".badge-meta { font-size:11px; color:#666; margin-left:4px; }" +
      "</style>" +
      "</head>" +
      "<body>" +
      '<div class="card">' +
      '<div class="title">畑クエスト ユーザーカード</div>' +
      '<div class="subtitle">畑の参加記録から、「学びの量」と「スキルの練度」を可視化しています。</div>' +
      '<div class="name">' +
      name +
      "</div>" +
      '<div class="level">現在のレベル：' +
      levelName +
      "</div>" +
      '<div class="total">投稿回数&スキル学習★：' +
      total +
      "</div>" +
      '<div class="skills-title">スキル練度</div>' +
      skillsHtml +
      badgeHtml +
      "</div>" +
      "</body>" +
      "</html>";

    return HtmlService.createHtmlOutput(html)
      .setTitle("畑クエスト ユーザーカード")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (err) {
    Logger.log(
      "doGet error: " + err + "\n" + (err && err.stack ? err.stack : "")
    );

    return HtmlService.createHtmlOutput(
      "<h2>畑クエスト ユーザーカード</h2><p>エラーが発生しました。</p>"
    )
      .setTitle("畑クエスト ユーザーカード")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
}
