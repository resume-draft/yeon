/**
 * Pawprints v2 — 둘만의 기록장 저장소 (Google Apps Script)
 *
 * 스크립트 속성
 *   PIN      : Partner(도장 찍는 사람) 비밀번호 — 기존 값 그대로 사용
 *   PET_PIN  : Pet(기록하는 사람) 비밀번호 — 새로 추가
 * 두 비밀번호 중 하나가 있어야 읽기·쓰기가 모두 가능해요. 링크만으로는 아무것도 보이지 않아요.
 *
 * 배포: [배포] → [배포 관리] → 연필 → 버전: 새 버전 → 배포 (웹 앱 URL은 그대로)
 */

const SHEETS = {
  entries: ["id", "type", "author", "at", "title", "goal_id", "context", "reinforcers", "mood", "stage", "paused", "tags", "note", "special"],
  reactions: ["id", "entry_id", "kind", "value", "author", "at"],
  goals: ["id", "name", "stage", "active", "next_step", "created_at", "author"],
  settings: ["key", "value"]
};
const TYPES = ["practice", "moment", "stamp", "note", "order", "answer", "rule", "wish"];
const KINDS = ["stamp", "emoji", "comment", "heart", "done"];
// 설정마다 바꿀 수 있는 사람: any = 둘 다
const SETTING_KEYS = { start: "any", agreements: "any", signal: "pet", reward: "partner" };

function doGet() {
  return json_({ error: "auth" });
}

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json_({ error: "bad_request" }); }

  const role = role_(body.key);
  if (!role) return json_({ error: "pin" });
  if (body.action === "check") return json_({ role: role });

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    let out = {};
    switch (body.action) {
      case "read": break;
      case "addEntry": out.added = addEntry_(body.entry || {}, role); break;
      case "updateEntry": updateEntry_(body.id, body.entry || {}, role); break;
      case "removeEntry": removeEntry_(body.id, role); break;
      case "react": out.added = react_(body.reaction || {}, role); break;
      case "removeReaction": removeReaction_(body.id, role); break;
      case "star": if (role !== "partner") throw new Error("forbidden"); star_(body.id, !!body.on); break;
      case "setGoal": if (role !== "partner") throw new Error("forbidden"); out.added = setGoal_(body.goal || {}, role); break;
      case "setSetting": setSetting_(body.k, body.v, role); break;
      default: return json_({ error: "bad_request" });
    }
    const data = readAll_();
    data.role = role;
    Object.assign(data, out);
    return json_(data);
  } catch (err) {
    return json_({ error: err.message === "forbidden" ? "forbidden" : "server" });
  } finally {
    lock.releaseLock();
  }
}

function role_(key) {
  const p = PropertiesService.getScriptProperties();
  const k = String(key || "");
  if (!k) return null;
  const pet = p.getProperty("PET_PIN"), partner = p.getProperty("PIN");
  if (pet && k === pet) return "pet";
  if (partner && k === partner) return "partner";
  return null;
}

/* ---------- entries ---------- */
function clean_(v, n) { return String(v == null ? "" : v).trim().slice(0, n); }
function list_(v) { return (Array.isArray(v) ? v : []).map(x => clean_(x, 20)).filter(Boolean).slice(0, 8).join("|"); }
function num_(v, lo, hi) { const n = Number(v); return n >= lo && n <= hi ? Math.round(n) : ""; }

function addEntry_(e, role) {
  const type = TYPES.indexOf(e.type) >= 0 ? e.type : "note";
  let title = clean_(e.title, 300);
  let ref = clean_(e.goal_id, 12);
  if (type === "stamp" && role !== "partner") throw new Error("forbidden");
  if (type === "order" && role !== "partner") {
    // 연이는 지시 서랍에서 뽑을 때만 지시를 만들 수 있어요 (내용은 파트너가 쓴 것)
    const gs = sheet_("goals"), r = findRow_(gs, ref);
    if (r < 0 || gs.getRange(r, 4).getValue() !== "1") throw new Error("forbidden");
    title = String(gs.getRange(r, 2).getValue());
    gs.getRange(r, 4).setValue("");
  }
  if (!title) throw new Error("empty");
  const special = type === "rule" && role === "partner" ? "1" : "";
  const id = Utilities.getUuid().slice(0, 8);
  sheet_("entries").appendRow([id, type, role, new Date().toISOString(), title, ref,
    list_(e.context), list_(e.reinforcers), num_(e.mood, 1, 5), num_(e.stage, 1, 3), e.paused ? "1" : "",
    list_(e.tags), clean_(e.note, 500), special]);
  return id;
}

function findRow_(sh, id) {
  const ids = sh.getRange(1, 1, Math.max(sh.getLastRow(), 1), 1).getValues();
  for (let i = ids.length - 1; i >= 1; i--) if (String(ids[i][0]) === String(id)) return i + 1;
  return -1;
}

function updateEntry_(id, e, role) {
  const sh = sheet_("entries"), r = findRow_(sh, id);
  if (r < 0) return;
  const row = sh.getRange(r, 1, 1, SHEETS.entries.length).getValues()[0];
  if (row[2] !== role) throw new Error("forbidden");
  if (e.title != null) { const t = clean_(e.title, 120); if (t) row[4] = t; }
  if (e.note != null) row[12] = clean_(e.note, 500);
  if (e.context) row[6] = list_(e.context);
  if (e.reinforcers) row[7] = list_(e.reinforcers);
  if (e.mood != null) row[8] = num_(e.mood, 1, 5);
  if (e.stage != null) row[9] = num_(e.stage, 1, 3);
  if (e.tags) row[11] = list_(e.tags);
  sh.getRange(r, 1, 1, row.length).setValues([row]);
}

function removeEntry_(id, role) {
  const sh = sheet_("entries"), r = findRow_(sh, id);
  if (r < 0) return;
  const author = sh.getRange(r, 3).getValue(), type = sh.getRange(r, 2).getValue();
  const shared = (role === "partner" && ["rule", "wish", "order"].indexOf(type) >= 0) || type === "wish";
  if (author !== role && !shared) throw new Error("forbidden");
  sh.deleteRow(r);
  const rs = sheet_("reactions"), vals = rs.getDataRange().getValues();
  for (let i = vals.length - 1; i >= 1; i--) if (String(vals[i][1]) === String(id)) rs.deleteRow(i + 1);
}

function star_(id, on) {
  const sh = sheet_("entries"), r = findRow_(sh, id);
  if (r > 0) sh.getRange(r, 14).setValue(on ? "1" : "");
}

/* ---------- reactions ---------- */
function react_(x, role) {
  const kind = KINDS.indexOf(x.kind) >= 0 ? x.kind : null;
  if (!kind) throw new Error("bad");
  if (kind === "stamp" && role !== "partner") throw new Error("forbidden");
  const value = clean_(x.value, kind === "comment" || kind === "done" ? 120 : 8);
  if ((kind === "comment" || kind === "emoji") && !value) throw new Error("empty");
  const id = Utilities.getUuid().slice(0, 8);
  sheet_("reactions").appendRow([id, clean_(x.entry_id, 12), kind, value, role, new Date().toISOString()]);
  return id;
}

function removeReaction_(id, role) {
  const sh = sheet_("reactions"), r = findRow_(sh, id);
  if (r < 0) return;
  if (sh.getRange(r, 5).getValue() !== role) throw new Error("forbidden");
  sh.deleteRow(r);
}

/* ---------- goals & settings ---------- */
function setGoal_(g, role) {
  const sh = sheet_("goals");
  if (g.id) {
    const r = findRow_(sh, g.id);
    if (r < 0) return g.id;
    const row = sh.getRange(r, 1, 1, SHEETS.goals.length).getValues()[0];
    if (g.name != null) { const n = clean_(g.name, 40); if (n) row[1] = n; }
    if (g.stage != null) row[2] = num_(g.stage, 1, 3) || row[2];
    if (g.active != null) row[3] = g.active ? "1" : "";
    if (g.next_step != null) row[4] = clean_(g.next_step, 80);
    sh.getRange(r, 1, 1, row.length).setValues([row]);
    return g.id;
  }
  const name = clean_(g.name, 40);
  if (!name) throw new Error("empty");
  const id = "g" + Utilities.getUuid().slice(0, 6);
  sh.appendRow([id, name, num_(g.stage, 1, 3) || 1, "1", clean_(g.next_step, 80), new Date().toISOString(), role]);
  return id;
}

function bumpGoal_(id, stage) {
  const sh = sheet_("goals"), r = findRow_(sh, id);
  if (r < 0) return;
  const cur = Number(sh.getRange(r, 3).getValue()) || 1;
  if (stage > cur) sh.getRange(r, 3).setValue(stage);
}

function setSetting_(k, v, role) {
  const who = SETTING_KEYS[k];
  if (!who) throw new Error("bad");
  if (who !== "any" && who !== role) throw new Error("forbidden");
  let val = clean_(v, 2000);
  if (k === "start" && !/^\d{4}-\d{2}-\d{2}$/.test(val)) throw new Error("bad");
  if (k === "signal" && ["green", "yellow", "red"].indexOf(val) < 0) throw new Error("bad");
  putSetting_(k, val);
  putSetting_(k + "_at", new Date().toISOString());
}

function putSetting_(k, v) {
  const sh = sheet_("settings"), rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) if (rows[i][0] === k) { sh.getRange(i + 1, 2).setValue(v); return; }
  sh.appendRow([k, v]);
}

/* ---------- read ---------- */
function rows_(name) {
  const h = SHEETS[name];
  return sheet_(name).getDataRange().getValues().slice(1).filter(r => r[0] !== "").map(r => {
    const o = {};
    h.forEach((k, i) => { const v = r[i]; o[k] = v instanceof Date ? v.toISOString() : String(v); });
    return o;
  });
}

function readAll_() {
  const settings = {};
  sheet_("settings").getDataRange().getValues().slice(1).forEach(r => {
    settings[r[0]] = r[1] instanceof Date ? Utilities.formatDate(r[1], "Asia/Seoul", "yyyy-MM-dd") : String(r[1]);
  });
  if (!settings.ntfy_partner) { settings.ntfy_partner = "paw-" + Utilities.getUuid().replace(/-/g, "").slice(0, 16); putSetting_("ntfy_partner", settings.ntfy_partner); }
  if (!settings.ntfy_pet) { settings.ntfy_pet = "yeon-fd0cb02d0964fd89"; putSetting_("ntfy_pet", settings.ntfy_pet); }
  return { entries: rows_("entries"), reactions: rows_("reactions"), goals: rows_("goals"), settings: settings };
}

function sheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 2000, SHEETS[name].length).setNumberFormat("@");
    sh.appendRow(SHEETS[name]);
    if (name === "entries") migrateStamps_(sh);
  }
  return sh;
}

/** 예전 stamps 시트의 칭찬 도장을 entries로 옮겨요 (벌매·정산 테스트 기록은 제외) */
function migrateStamps_(sh) {
  const old = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("stamps");
  if (!old) return;
  const rows = old.getDataRange().getValues().slice(1)
    .filter(r => r[0] !== "" && !/^⟦/.test(String(r[1])))
    .map(r => [String(r[0]), "stamp", "partner", r[2] instanceof Date ? r[2].toISOString() : String(r[2]), String(r[1]), "", "", "", "", "", "", "", "", ""]);
  if (rows.length) sh.getRange(2, 1, rows.length, SHEETS.entries.length).setValues(rows);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
