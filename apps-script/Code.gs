/**
 * KatagrMa マネジメントスケジュール診断 — 共有台帳バックエンド
 * Google Apps Script（スプレッドシートに紐づくウェブアプリ）
 *
 * 役割：
 *   - action=save … 診断結果を1行追記する
 *   - action=list … 保存済みの全レコードを新しい順で返す
 *
 * フロント（index.html）とは JSONP でやり取りするため、
 * GitHub Pages など別ドメインから呼んでも CORS の問題が起きません。
 *
 * ※ TOKEN は index.html の CONFIG.TOKEN と必ず同じ文字列にすること。
 */

const SHEET_NAME = "records";
const MASTER_SHEET = "";                 // 法人名リストのシート名。空なら「records以外の最初のシート」のA列を使う
const MASTER_COL = 1;                    // 法人名が入っている列番号（A列=1、B列=2 …）
const TOKEN = "kata-ls-8995de7ce34c";   // ← index.html の CONFIG.TOKEN と一致させること（CS専用・閲覧含む全操作）
const SUBMIT_TOKEN = "kata-ls-submit-2f91ab7c4e"; // ← respond.html の CONFIG.SUBMIT_TOKEN と一致させること（顧客回答URL専用・新規追加のみ）
const APP_VERSION = "v3-public-submit";       // デプロイ確認用。pingで返る。

// このウェブアプリは「全員（匿名アクセスも含む）」でデプロイする想定。
// TOKEN（CS用）は list/hojins/save/update など全操作を許可する一方、
// SUBMIT_TOKEN（顧客回答URL用）は submitSurvey（1行追加のみ）しか許可しない。
// 顧客用リンクが漏れても他法人のデータ閲覧・改変はできない設計。
function doGet(e) {
  const p = (e && e.parameter) || {};
  const cb = p.callback || "callback";
  let out;
  try {
    if (p.action === "ping")         return textOut(cb, { ok: true, version: APP_VERSION });
    if (p.action === "submitSurvey") {
      if (p.token !== SUBMIT_TOKEN) throw new Error("認証トークンが一致しません");
      out = handleSubmitSurvey(p);
    } else {
      if (p.token !== TOKEN) throw new Error("認証トークンが一致しません");
      if (p.action === "save")         out = handleSave(p);
      else if (p.action === "update")  out = handleUpdate(p);
      else if (p.action === "list")    out = handleList();
      else if (p.action === "hojins")  out = handleHojins();
      else                             out = { ok: false, error: "不明なaction: " + p.action };
    }
  } catch (err) {
    out = { ok: false, error: String(err && err.message ? err.message : err) };
  }
  return textOut(cb, out);
}

function textOut(cb, out) {
  return ContentService
    .createTextOutput(cb + "(" + JSON.stringify(out) + ")")
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

// 列構成： A保存日時 B法人名 C実施日 D記入者 E観点 F レベル G サイクル H開始月 I payload J メモ K id
function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(["保存日時", "法人名", "実施日", "記入者", "観点", "レベル", "サイクル", "開始月", "payload", "メモ", "id"]);
  } else {
    if (!sh.getRange(1, 10).getValue()) sh.getRange(1, 10).setValue("メモ");  // 既存シートに列を追加
    if (!sh.getRange(1, 11).getValue()) sh.getRange(1, 11).setValue("id");
  }
  return sh;
}

function handleSave(p) {
  const sh = getSheet();
  const id = p.id || ("r" + Date.now() + Math.floor(Math.random() * 10000));
  sh.appendRow([
    new Date(),
    p.hojin || "",
    p.date || "",
    p.author || "",
    p.focusesText || "",
    p.levelsText || "",
    p.cycle || "",
    p.month || "",
    p.payload || "",
    p.memo || "",
    id
  ]);
  return { ok: true, id: id };
}

// 顧客回答URL専用：新規行を1件追加するだけ。CS用の handleSave と処理は近いが、
// 信頼できるフィールドだけを明示的に拾い、他フィールド（cycle/monthなど）は常に空にする。
// list/update などの操作はこのトークンでは一切許可しない（doGet 側で分岐済み）。
function handleSubmitSurvey(p) {
  if (!p.hojin) throw new Error("法人名が指定されていません");
  const sh = getSheet();
  const id = "r" + Date.now() + Math.floor(Math.random() * 10000);
  sh.appendRow([
    new Date(),
    p.hojin || "",
    p.date || "",
    p.author || "",
    p.focusesText || "",
    p.levelsText || "",
    "",           // サイクルはCS側で後から設定
    "",           // 開始月はCS側で後から設定
    p.payload || "",
    "",           // メモはCS側で追記
    id
  ]);
  return { ok: true, id: id };
}

// 既存レコード（id一致の行）を上書き更新。渡されたフィールドだけ更新する。
function handleUpdate(p) {
  const id = p.id;
  if (!id) throw new Error("idが指定されていません");
  const sh = getSheet();
  const values = sh.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][10]) === String(id)) {       // K列=id（index 10）
      var row = i + 1;
      if (p.hojin       !== undefined) sh.getRange(row, 2).setValue(p.hojin);
      if (p.date        !== undefined) sh.getRange(row, 3).setValue(p.date);
      if (p.author      !== undefined) sh.getRange(row, 4).setValue(p.author);
      if (p.focusesText !== undefined) sh.getRange(row, 5).setValue(p.focusesText);
      if (p.levelsText  !== undefined) sh.getRange(row, 6).setValue(p.levelsText);
      if (p.cycle       !== undefined) sh.getRange(row, 7).setValue(p.cycle);
      if (p.month       !== undefined) sh.getRange(row, 8).setValue(p.month);
      if (p.payload     !== undefined) sh.getRange(row, 9).setValue(p.payload);
      if (p.memo        !== undefined) sh.getRange(row, 10).setValue(p.memo);
      return { ok: true };
    }
  }
  return { ok: false, error: "該当レコードが見つかりません（id: " + id + "）" };
}

// 法人名リストのシートを探す：MASTER_SHEET 指定があればそれ、無ければ records 以外の最初のシート
function findMasterSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (MASTER_SHEET) {
    const s = ss.getSheetByName(MASTER_SHEET);
    if (s) return s;
  }
  const sheets = ss.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    if (sheets[i].getName() !== SHEET_NAME) return sheets[i];
  }
  return null;
}

function handleHojins() {
  const sh = findMasterSheet();
  if (!sh) return { ok: true, hojins: [] };
  const values = sh.getDataRange().getValues();
  const HEADERS = { "取引先名": 1, "法人名": 1, "": 1 };  // 見出し行や空セルは除外
  const seen = {};
  const names = [];
  for (let i = 0; i < values.length; i++) {
    const name = String(values[i][MASTER_COL - 1] || "").trim();
    if (!name || HEADERS[name] || seen[name]) continue;
    seen[name] = 1;
    names.push(name);
  }
  return { ok: true, hojins: names };
}

function handleList() {
  const sh = getSheet();
  const values = sh.getDataRange().getValues();
  const rows = [];
  for (let i = 1; i < values.length; i++) {
    const r = values[i];
    if (!r[1]) continue;  // 法人名が空の行はスキップ
    rows.push({
      ts:          r[0] ? new Date(r[0]).toISOString() : "",
      hojin:       r[1],
      date:        r[2],
      author:      r[3],
      focusesText: r[4],
      levelsText:  r[5],
      cycle:       r[6],
      month:       r[7],
      payload:     r[8],
      memo:        r[9] || "",
      id:          r[10] || ""
    });
  }
  rows.reverse();  // 新しい順
  return { ok: true, rows: rows };
}
