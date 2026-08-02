/* 申込窓口の画面。サーバー(api.php)が正で、この側は「表を描く・集める」だけ。
   検証はサーバーで行い、返ってきたエラーをそのまま出す（規則の二重実装を避ける）。 */
"use strict";

const $ = (sel) => document.querySelector(sel);
const $$ = (sel, root) => Array.from((root || document).querySelectorAll(sel));

let S = null;        // state: {school, members, tournament, limits, year, entries, meta}
let currentTab = "advisors";
let entryDirty = false;

async function api(action, data) {
  const res = await fetch(data === undefined ? `api.php?action=${action}` : "api.php", {
    method: data === undefined ? "GET" : "POST",
    headers: data === undefined ? {} : { "Content-Type": "application/json", "X-Entry-Api": "1" },
    body: data === undefined ? undefined : JSON.stringify({ action, ...data }),
  });
  let body = null;
  try { body = await res.json(); } catch { /* 下で拾う */ }
  if (!body) throw new Error(`サーバーから読めない応答（HTTP ${res.status}）`);
  return body;
}

function esc(s) {
  return String(s ?? "").replace(/[&<>"']/g, (c) =>
    ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}

function showErrors(boxSel, body) {
  const box = $(boxSel);
  const msgs = body.errors || (body.error ? [body.error] : []);
  if (msgs.length) {
    box.textContent = msgs.map((m) => "・" + m).join("\n");
    box.classList.remove("hidden");
    box.scrollIntoView({ behavior: "smooth", block: "nearest" });
  } else {
    box.classList.add("hidden");
  }
  return msgs.length > 0;
}

function flash(sel, text) {
  const el = $(sel);
  el.textContent = text;
  setTimeout(() => { el.textContent = ""; }, 4000);
}

/* ---------- ログイン ---------- */

async function initLogin() {
  const body = await api("schools");
  $("#login-school").innerHTML = '<option value="">— 学校を選ぶ —</option>' +
    body.schools.map((n) => `<option>${esc(n)}</option>`).join("");
}

async function doLogin() {
  $("#login-error").classList.add("hidden");
  const body = await api("login", {
    school: $("#login-school").value,
    password: $("#login-password").value,
  });
  if (!body.ok) { showErrors("#login-error", body); return; }
  await loadState();
}

async function doLogout() {
  await api("logout", {});
  S = null;
  $("#view-app").classList.add("hidden");
  $("#view-login").classList.remove("hidden");
  $("#login-password").value = "";
  updateSavebar();
}

/* ---------- 状態の読み込み ---------- */

async function loadState() {
  const body = await api("state");
  if (!body.ok) {
    $("#view-login").classList.remove("hidden");
    $("#view-app").classList.add("hidden");
    return;
  }
  S = body;
  $("#view-login").classList.add("hidden");
  $("#view-app").classList.remove("hidden");
  $("#hd-school").textContent = S.school.name;
  $("#hd-title").textContent = S.tournament
    ? `令和${S.year}年度 ${S.tournament.name} — 参加申込` : "参加申込（受付中の大会なし）";
  renderDue();
  renderAdvisors();
  renderMembers();
  renderEntry();
  renderUploads();
  renderSteps();
  setDirty(false);
}

function renderDue() {
  const el = $("#hd-due");
  const d = S.tournament && S.tournament.deadline;
  if (!d) { el.classList.add("hidden"); return; }
  let text = `締切 ${d}`;
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(d);
  if (m) {
    const due = new Date(+m[1], +m[2] - 1, +m[3]);
    const days = Math.ceil((due - new Date()) / 86400000);
    text = `締切 ${+m[2]}月${+m[3]}日` + (days >= 0 ? `（あと${days}日）` : "（締切済み）");
  }
  el.textContent = text;
  el.classList.remove("hidden");
}

function renderSteps() {
  const adv = (S.school.advisors || []).length;
  const mem = (S.members || []).length;
  const ent = Object.values(S.entries || {}).filter((e) =>
    e.team_kata_chk || e.team_kumi_chk || e.kata_chk || e.kumi_chk).length;
  const set = (sel, n, unit) => {
    const el = $(sel);
    el.textContent = n > 0 ? `済 ${n}${unit}` : "未";
    el.classList.toggle("done", n > 0);
  };
  set("#st-adv", adv, "名");
  set("#st-mem", mem, "名");
  set("#st-ent", ent, "名");
  const up = (S.uploads || []).length;
  const el = $("#st-form");
  el.textContent = up > 0 ? "提出済" : "未提出";
  el.classList.toggle("done", up > 0);
}

/* ---------- 1. 顧問 ---------- */

const ROLE_OPTS = ["審判", "競技記録", "係員"];

function advisorRow(a = {}) {
  const tr = document.createElement("tr");
  tr.innerHTML = `
    <td><input type="text" class="a-name" value="${esc(a.name)}"></td>
    <td><select class="a-role">${ROLE_OPTS.map((r) =>
      `<option ${a.role === r ? "selected" : ""}>${r}</option>`).join("")}</select></td>
    <td style="text-align:center"><input type="checkbox" class="a-d1" ${a.d1 ? "checked" : ""}></td>
    <td style="text-align:center"><input type="checkbox" class="a-d2" ${a.d2 ? "checked" : ""}></td>
    <td class="del"><button class="x" title="この行を外す">×</button></td>`;
  tr.querySelector(".x").onclick = () => tr.remove();
  return tr;
}

function renderAdvisors() {
  $("#principal").value = S.school.principal || "";
  const tb = $("#adv-table tbody");
  tb.innerHTML = "";
  (S.school.advisors || []).forEach((a) => tb.appendChild(advisorRow(a)));
  if (!tb.children.length) tb.appendChild(advisorRow());
}

async function saveAdvisors() {
  const advisors = $$("#adv-table tbody tr").map((tr) => ({
    name: tr.querySelector(".a-name").value.trim(),
    role: tr.querySelector(".a-role").value,
    d1: tr.querySelector(".a-d1").checked,
    d2: tr.querySelector(".a-d2").checked,
  })).filter((a) => a.name !== "");
  const body = await api("save_advisors", { principal: $("#principal").value, advisors });
  if (showErrors("#adv-errors", body.ok ? {} : body)) return;
  flash("#adv-msg", "保存しました");
  await loadState();
}

/* ---------- 2. 名簿 ---------- */

function memberRow(m = {}) {
  const tr = document.createElement("tr");
  tr.innerHTML = `
    <td><input type="text" class="m-no num" value="${esc(m.display_order)}"></td>
    <td><input type="text" class="m-name" value="${esc(m.name)}"></td>
    <td><select class="m-sex">
      <option value="">—</option>
      <option ${m.sex === "男子" ? "selected" : ""}>男子</option>
      <option ${m.sex === "女子" ? "selected" : ""}>女子</option></select></td>
    <td><select class="m-grade">
      <option value="">—</option>
      ${[1, 2, 3].map((g) => `<option ${String(m.grade) === String(g) ? "selected" : ""}>${g}</option>`).join("")}
    </select></td>
    <td><input type="text" class="m-dob num" placeholder="2010/04/01" value="${esc(m.dob)}"></td>
    <td><input type="text" class="m-jkf num" value="${esc(m.jkf_no)}"></td>
    <td class="del"><button class="x" title="この行を外す">×</button></td>`;
  tr.querySelector(".x").onclick = () => tr.remove();
  return tr;
}

function renderMembers() {
  const tb = $("#mem-table tbody");
  tb.innerHTML = "";
  (S.members || []).forEach((m) => tb.appendChild(memberRow(m)));
  if (!tb.children.length) tb.appendChild(memberRow());
}

async function saveMembers() {
  const members = $$("#mem-table tbody tr").map((tr) => ({
    display_order: tr.querySelector(".m-no").value,
    name: tr.querySelector(".m-name").value,
    sex: tr.querySelector(".m-sex").value,
    grade: tr.querySelector(".m-grade").value,
    dob: tr.querySelector(".m-dob").value,
    jkf_no: tr.querySelector(".m-jkf").value,
  }));
  const body = await api("save_members", { members });
  if (showErrors("#mem-errors", body.ok ? {} : body)) return;
  flash("#mem-msg", "保存しました");
  await loadState();
}

/* ---------- 3. エントリー ---------- */

function eligibleNames(sex) {
  const grades = (S.tournament?.grades || [1, 2, 3]).map(Number);
  return (S.members || [])
    .filter((m) => m.sex === sex && grades.includes(Number(m.grade)))
    .map((m) => m.name);
}

function nameOptions(sex, current) {
  return '<option value=""></option>' + eligibleNames(sex).map((n) => {
    const g = (S.members.find((m) => m.name === n) || {}).grade;
    return `<option value="${esc(n)}" ${n === current ? "selected" : ""}>${esc(n)}（${esc(g)}年）</option>`;
  }).join("");
}

function entriesBySex(sex, pred) {
  const names = new Set(eligibleNames(sex));
  return Object.entries(S.entries || {})
    .filter(([n, e]) => names.has(n) && pred(e))
    .map(([n, e]) => ({ name: n, ...e }));
}

/* --- 団体（固定枠） --- */

function renderTeamSlots(tbodyId, sex, chk, roleKey, lim) {
  const tb = document.getElementById(tbodyId);
  tb.innerHTML = "";
  if (!lim) return;
  const cur = entriesBySex(sex, (e) => e[chk]);
  const regs = cur.filter((e) => e[roleKey] !== "補").map((e) => e.name);
  const subs = cur.filter((e) => e[roleKey] === "補").map((e) => e.name);
  for (let i = 0; i < lim.max; i++) {
    tb.insertAdjacentHTML("beforeend",
      `<tr><td class="slot">正</td><td><select class="t-name">${nameOptions(sex, regs[i] || "")}</select></td></tr>`);
  }
  for (let i = 0; i < lim.sub_max; i++) {
    tb.insertAdjacentHTML("beforeend",
      `<tr><td class="slot">補</td><td><select class="t-name">${nameOptions(sex, subs[i] || "")}</select></td></tr>`);
  }
}

function setPaneOff(paneId, off) {
  const pane = document.getElementById(paneId);
  pane.classList.toggle("off", off);
  pane.querySelector(".body").classList.toggle("hidden", off);
  pane.querySelector(".offnote").classList.toggle("hidden", !off);
}

function refreshTeamKata() {
  const lim = S.limits.team_kata;
  const mOn = $("#m-tk-part").value === "yes";
  const wOn = $("#w-tk-part").value === "yes";
  setPaneOff("pane-m_tk", !mOn);
  setPaneOff("pane-w_tk", !wOn);
  if (mOn) renderTeamSlots("tbl-m_tk", "男子", "team_kata_chk", "team_kata_role", lim);
  if (wOn) renderTeamSlots("tbl-w_tk", "女子", "team_kata_chk", "team_kata_role", lim);
  refreshCounts();
}

function refreshTeamKumite() {
  const mMode = $("#m-mode").value;
  const wMode = $("#w-mode").value;
  setPaneOff("pane-m_tku", mMode === "none");
  setPaneOff("pane-w_tku", wMode === "none");
  if (mMode !== "none") {
    renderTeamSlots("tbl-m_tku", "男子", "team_kumi_chk", "team_kumi_role", S.limits["team_kumite_" + mMode]);
  }
  if (wMode !== "none") {
    renderTeamSlots("tbl-w_tku", "女子", "team_kumi_chk", "team_kumi_role", S.limits["team_kumite_" + wMode]);
  }
  refreshCounts();
}

function collectTeamTable(tbodyId) {
  return $$(`#${tbodyId} tr`).map((tr) => ({
    role: tr.querySelector(".slot").textContent.trim(),
    name: tr.querySelector(".t-name").value,
  }));
}

/* --- 個人（選手1行ずつ） --- */

const KUBUN_OPTS = ["正", "シード", "補"];

function kubunSelect(current) {
  return `<select class="i-kubun">${KUBUN_OPTS.map((k) =>
    `<option ${current === k ? "selected" : ""}>${k}</option>`).join("")}</select>`;
}

function indRow(sex, row = { kubun: "正", rank: "", name: "" }) {
  const tr = document.createElement("tr");
  const hosu = row.kubun === "補";
  tr.innerHTML = `
    <td><select class="i-name">${nameOptions(sex, row.name)}</select></td>
    <td>${kubunSelect(row.kubun)}</td>
    <td class="raux"><input type="text" class="i-rank rk num" value="${esc(row.rank)}"
        ${hosu ? 'disabled placeholder="—"' : ""}></td>
    <td class="del"><button class="x" title="この行を外す">×</button></td>`;
  tr.querySelector(".x").onclick = () => { tr.remove(); setDirty(true); refreshCounts(); };
  return tr;
}

/* 階級のある大会（新人・選抜）かどうか。**サーバーの save_entry と同じ判定にすること**
   （食い違うと、画面が読む欄と保存する欄がずれてデータが壊れる） */
function isWeightTournament() {
  return ((S.tournament?.weights_m_names || []).length > 0)
      || ((S.tournament?.weights_w_names || []).length > 0);
}

function kuRow(sex, weights, row = { weight: "", kubun: "正", rank: "", name: "" }) {
  const tr = document.createElement("tr");
  const hosu = row.kubun === "補";
  const wcell = isWeightTournament()
    ? `<td><select class="i-weight">${weights.map((w) =>
        `<option ${row.weight === w ? "selected" : ""}>${esc(w)}</option>`).join("")}</select></td>`
    : "";
  tr.innerHTML = `
    <td><select class="i-name">${nameOptions(sex, row.name)}</select></td>
    ${wcell}
    <td>${kubunSelect(row.kubun)}</td>
    <td class="raux"><input type="text" class="i-rank rk num" value="${esc(row.rank)}"
        ${hosu ? 'disabled placeholder="—"' : ""}></td>
    <td class="del"><button class="x" title="この行を外す">×</button></td>`;
  tr.querySelector(".x").onclick = () => { tr.remove(); setDirty(true); refreshCounts(); };
  return tr;
}

function renderIndKata(tbodyId, sex) {
  const tb = document.getElementById(tbodyId);
  tb.innerHTML = "";
  const rows = entriesBySex(sex, (e) => e.kata_chk)
    .map((e) => ({ kubun: e.kata_val || "正", rank: e.kata_rank, name: e.name }));
  rows.forEach((r) => tb.appendChild(indRow(sex, r)));
  if (!rows.length) tb.appendChild(indRow(sex));
}

/* 個人組手の持ち方は大会種別で違う（申込システム app.py 由来の仕様）:
     階級のある大会 … kumi_val=階級、kumi_sub_val=区分（正/シード/補）
     階級のない大会 … kumi_val=区分。kumi_sub_val は空
   ここを取り違えると、画面ではシードが「正」に見え、そのまま保存すると
   実際にシードが消える（インハイ予選の実データで発覚・2026-07-31）。 */
function renderKumite(tbodyId, sex, weights) {
  const tb = document.getElementById(tbodyId);
  tb.innerHTML = "";
  const w = isWeightTournament();
  const rows = entriesBySex(sex, (e) => e.kumi_chk).map((e) => ({
    weight: w ? e.kumi_val : "",
    kubun: (w ? e.kumi_sub_val : e.kumi_val) || "正",
    rank: e.kumi_rank,
    name: e.name,
  }));
  rows.forEach((r) => tb.appendChild(kuRow(sex, weights, r)));
  if (!rows.length) tb.appendChild(kuRow(sex, weights));
}

function chipLabel(w) {
  return w.replace("kg級", "");
}

function renderChips(chipId, tbodyId, weights) {
  const root = document.getElementById(chipId);
  root.innerHTML = weights.map((w) =>
    `<span class="wchip" data-w="${esc(w)}"><b>${esc(chipLabel(w))}</b><span>0</span></span>`).join("");
  recountChips(chipId, tbodyId);
}

function recountChips(chipId, tbodyId) {
  const counts = {};
  $$(`#${tbodyId} .i-weight`).forEach((sel) => {
    const tr = sel.closest("tr");
    if (tr.querySelector(".i-name").value !== "") {
      counts[sel.value] = (counts[sel.value] || 0) + 1;
    }
  });
  $$(`#${chipId} .wchip`).forEach((ch) => {
    const n = counts[ch.dataset.w] || 0;
    ch.classList.toggle("on", n > 0);
    ch.querySelector("span:last-child").textContent = n ? `${n}名` : "0";
  });
}

function collectIndTable(tbodyId) {
  return $$(`#${tbodyId} tr`).map((tr) => ({
    kubun: tr.querySelector(".i-kubun").value,
    rank: tr.querySelector(".i-rank").value,
    name: tr.querySelector(".i-name").value,
  })).filter((r) => r.name !== "" || r.rank !== "");
}

function collectKumite(tbodyId) {
  return $$(`#${tbodyId} tr`).map((tr) => ({
    weight: tr.querySelector(".i-weight")?.value ?? "",
    kubun: tr.querySelector(".i-kubun").value,
    rank: tr.querySelector(".i-rank").value,
    name: tr.querySelector(".i-name").value,
  })).filter((r) => r.name !== "" || r.rank !== "");
}

function refreshCounts() {
  const filled = (tbodyId, cls) => $$(`#${tbodyId} ${cls}`).filter((s) => s.value !== "").length;
  $("#cnt-m_tk").textContent = `${filled("tbl-m_tk", ".t-name")}名`;
  $("#cnt-w_tk").textContent = `${filled("tbl-w_tk", ".t-name")}名`;
  $("#cnt-m_tku").textContent = `${filled("tbl-m_tku", ".t-name")}名`;
  $("#cnt-w_tku").textContent = `${filled("tbl-w_tku", ".t-name")}名`;
  $("#cnt-m_k").textContent = `${filled("tbl-m_k", ".i-name")}名`;
  $("#cnt-w_k").textContent = `${filled("tbl-w_k", ".i-name")}名`;
  $("#cnt-m_ku").textContent = `${filled("tbl-m_ku", ".i-name")}名`;
  $("#cnt-w_ku").textContent = `${filled("tbl-w_ku", ".i-name")}名`;
  recountChips("chips-m", "tbl-m_ku");
  recountChips("chips-w", "tbl-w_ku");
}

function renderEntry() {
  const has = !!S.tournament;
  $("#entry-none").classList.toggle("hidden", has);
  $("#entry-body").classList.toggle("hidden", !has);
  updateSavebar();
  if (!has) return;

  const lim = S.limits;
  $("#aux-tk").textContent =
    `正 ${lim.team_kata.max}名・補欠 ${lim.team_kata.sub_max}名まで`;
  $("#aux-tku").textContent =
    `5人制＝正${lim.team_kumite_5.min}〜${lim.team_kumite_5.max}名・補${lim.team_kumite_5.sub_max} ／ ` +
    `3人制＝正${lim.team_kumite_3.min}〜${lim.team_kumite_3.max}名・補${lim.team_kumite_3.sub_max}`;
  $("#aux-k").textContent =
    `正 ${lim.ind_kata_reg.max}名・補欠 ${lim.ind_kata_sub.max}名まで（シードは別枠）`;

  $("#m-mode").value = S.meta?.m_kumite_mode || "5";
  $("#w-mode").value = S.meta?.w_kumite_mode || "5";
  $("#m-tk-part").value = (S.meta?.part_m_tk === false) ? "no" : "yes";
  $("#w-tk-part").value = (S.meta?.part_w_tk === false) ? "no" : "yes";

  refreshTeamKata();
  refreshTeamKumite();
  renderIndKata("tbl-m_k", "男子");
  renderIndKata("tbl-w_k", "女子");
  // 階級のない大会（関東・インハイ）では階級の欄そのものを出さない
  const wm = S.tournament.weights_m_names || [];
  const ww = S.tournament.weights_w_names || [];
  const hasW = isWeightTournament();
  $$(".grid.ku th.wcol").forEach((th) => th.classList.toggle("hidden", !hasW));
  $("#chips-m").classList.toggle("hidden", !hasW);
  $("#chips-w").classList.toggle("hidden", !hasW);
  renderChips("chips-m", "tbl-m_ku", wm);
  renderChips("chips-w", "tbl-w_ku", ww);
  renderKumite("tbl-m_ku", "男子", wm);
  renderKumite("tbl-w_ku", "女子", ww);
  refreshCounts();
}

async function saveEntry() {
  const payload = {
    meta: {
      m_kumite_mode: $("#m-mode").value,
      w_kumite_mode: $("#w-mode").value,
      part_m_tk: $("#m-tk-part").value === "yes",
      part_w_tk: $("#w-tk-part").value === "yes",
    },
    tables: {
      m_tk: $("#m-tk-part").value === "yes" ? collectTeamTable("tbl-m_tk") : [],
      w_tk: $("#w-tk-part").value === "yes" ? collectTeamTable("tbl-w_tk") : [],
      m_tku: $("#m-mode").value === "none" ? [] : collectTeamTable("tbl-m_tku"),
      w_tku: $("#w-mode").value === "none" ? [] : collectTeamTable("tbl-w_tku"),
      m_k: collectIndTable("tbl-m_k"),
      w_k: collectIndTable("tbl-w_k"),
      m_ku: collectKumite("tbl-m_ku"),
      w_ku: collectKumite("tbl-w_ku"),
    },
  };
  const body = await api("save_entry", payload);
  if (showErrors("#entry-errors", body.ok ? {} : body)) return;
  await loadState();
  switchTab("entry");
  flashSavebar();
}

function flashSavebar() {
  $("#dirty-text").textContent = "保存しました";
  setTimeout(() => { if (!entryDirty) $("#dirty-text").textContent = "保存済み"; }, 4000);
}

/* ---------- 4. 申込書の提出 ---------- */

function renderUploads() {
  const ups = S.uploads || [];
  $("#up-limit").textContent = S.upload_limit ? `1ファイル ${S.upload_limit} まで` : "";
  $("#up-count").textContent = ups.length ? `${ups.length}件` : "";
  $("#up-none").classList.toggle("hidden", ups.length > 0);
  $("#up-list").innerHTML = ups.map((u, i) => `
    <tr>
      <td class="num">${esc(u.at)}</td>
      <td style="width:3.5em">${i === 0
        ? '<span class="badge new">最新</span>' : '<span class="badge">旧</span>'}</td>
      <td class="num" style="width:8em; color:var(--ink-3)">${esc(u.ext.toUpperCase())} ${esc(u.size_text)}</td>
      <td style="width:6em; text-align:right"><a
        href="api.php?action=download_upload&name=${encodeURIComponent(u.name)}">取り出す</a></td>
    </tr>`).join("");
}

/* 申込書は PDF だけ受け付ける（22名を超えると2ページになるため。写真だと2枚
   バラバラの提出になり、「いちばん新しいものが提出物」の決まりと噛み合わない）。
   サーバーでも同じ判定をするが、選んだ時点で知らせたほうが親切なのでここでも見る。 */
function isPdf(f) {
  return /\.pdf$/i.test(f.name);
}

function notPdfMessage(f) {
  const ext = (f.name.split(".").pop() || "").toLowerCase();
  return `申込書は PDF で提出してください（選ばれたもの: ${ext ? "." + ext : "拡張子なし"}）。`
    + "写真ではなく「書類をスキャン」でPDFにしてください"
    + "（上の「スマートフォンでPDFにするには」をご覧ください）。"
    + "どうしてもPDFにできない場合は、専門部までメールでお送りください";
}

async function uploadForm() {
  const input = $("#up-file");
  const f = input.files && input.files[0];
  if (!f) return;
  showErrors("#up-errors", {});
  if (!isPdf(f)) {
    showErrors("#up-errors", { error: notPdfMessage(f) });
    return;
  }
  if (S.upload_limit_bytes && f.size > S.upload_limit_bytes) {
    showErrors("#up-errors", { error: `ファイルが大きすぎます（上限 ${S.upload_limit}）` });
    return;
  }
  const btn = $("#up-send");
  btn.disabled = true;
  btn.textContent = "送信中…";
  try {
    const fd = new FormData();
    fd.append("file", f);
    // 送るのはファイルだけなので JSON の api() は通さない。CSRF よけの
    // ヘッダ（X-Entry-Api）は同じ決まりで付ける
    const res = await fetch("api.php?action=upload_form",
      { method: "POST", headers: { "X-Entry-Api": "1" }, body: fd });
    let body = null;
    try { body = await res.json(); } catch { /* 下で拾う */ }
    if (!body) throw new Error(`送信できませんでした（HTTP ${res.status}）`);
    if (showErrors("#up-errors", body.ok ? {} : body)) return;
    input.value = "";
    // ここで loadState() を呼ばない（エントリータブの未保存の入力が消えるため）
    S.uploads = body.uploads || [];
    renderUploads();
    renderSteps();
    flash("#up-msg", "提出しました");
  } finally {
    btn.textContent = "提出する";
    btn.disabled = !(input.files && input.files[0]);
  }
}

/* ---------- 保存バー・未保存の管理 ---------- */

function setDirty(v) {
  entryDirty = v;
  const el = $("#dirty");
  el.classList.toggle("clean", !v);
  $("#dirty-text").textContent = v ? "未保存の変更があります" : "保存済み";
}

function updateSavebar() {
  const show = currentTab === "entry" && S && !!S.tournament;
  $("#savebar").classList.toggle("hidden", !show);
}

/* ---------- タブと起動 ---------- */

function switchTab(name) {
  currentTab = name;
  $$(".step").forEach((b) => {
    if (b.dataset.tab === name) b.setAttribute("aria-current", "page");
    else b.removeAttribute("aria-current");
  });
  ["advisors", "members", "entry", "form"].forEach((t) =>
    $("#tab-" + t).classList.toggle("hidden", t !== name));
  updateSavebar();
}

window.addEventListener("DOMContentLoaded", async () => {
  $$(".step").forEach((b) => (b.onclick = () => switchTab(b.dataset.tab)));
  $("#btn-login").onclick = () => doLogin().catch((e) => showErrors("#login-error", { error: e.message }));
  $("#login-password").addEventListener("keydown", (e) => { if (e.key === "Enter") $("#btn-login").click(); });
  $("#btn-logout").onclick = doLogout;

  $("#adv-add").onclick = () => $("#adv-table tbody").appendChild(advisorRow());
  $("#adv-save").onclick = () => saveAdvisors().catch((e) => showErrors("#adv-errors", { error: e.message }));
  $("#mem-add").onclick = () => $("#mem-table tbody").appendChild(memberRow());
  $("#mem-save").onclick = () => saveMembers().catch((e) => showErrors("#mem-errors", { error: e.message }));

  $("#m-mode").onchange = () => { refreshTeamKumite(); setDirty(true); };
  $("#w-mode").onchange = () => { refreshTeamKumite(); setDirty(true); };
  $("#m-tk-part").onchange = () => { refreshTeamKata(); setDirty(true); };
  $("#w-tk-part").onchange = () => { refreshTeamKata(); setDirty(true); };

  $$(".add-ind").forEach((b) => (b.onclick = () => {
    const sex = b.dataset.target.startsWith("m") ? "男子" : "女子";
    document.getElementById("tbl-" + b.dataset.target).appendChild(indRow(sex));
    setDirty(true); refreshCounts();
  }));
  $$(".add-ku").forEach((b) => (b.onclick = () => {
    const sex = b.dataset.target.startsWith("m") ? "男子" : "女子";
    const weights = sex === "男子" ? (S.tournament.weights_m_names || []) : (S.tournament.weights_w_names || []);
    document.getElementById("tbl-" + b.dataset.target).appendChild(kuRow(sex, weights));
    setDirty(true); refreshCounts();
  }));

  // エントリータブ内の入力はどれでも「未保存」に。区分=補は順位を入力不可に
  $("#tab-entry").addEventListener("change", (e) => {
    setDirty(true);
    if (e.target.classList.contains("i-kubun")) {
      const rk = e.target.closest("tr").querySelector(".i-rank");
      const hosu = e.target.value === "補";
      rk.disabled = hosu;
      if (hosu) rk.value = "";
      rk.placeholder = hosu ? "—" : "";
    }
    refreshCounts();
  });
  $("#tab-entry").addEventListener("input", () => setDirty(true));

  $("#up-file").onchange = () => {
    showErrors("#up-errors", {});
    const f = $("#up-file").files[0];
    // 押す前に知らせる（提出してから断られるより分かりやすい）
    $("#up-send").disabled = !f || !isPdf(f);
    if (f && !isPdf(f)) showErrors("#up-errors", { error: notPdfMessage(f) });
  };
  $("#up-send").onclick = () => uploadForm().catch((e) => showErrors("#up-errors", { error: e.message }));

  $("#entry-save").onclick = () => saveEntry().catch((e) => showErrors("#entry-errors", { error: e.message }));
  $("#entry-discard").onclick = async () => { await loadState(); switchTab("entry"); };

  window.addEventListener("beforeunload", (e) => {
    if (entryDirty) { e.preventDefault(); e.returnValue = ""; }
  });

  await initLogin().catch(() => {});
  await loadState().catch(() => { $("#view-login").classList.remove("hidden"); });
});
