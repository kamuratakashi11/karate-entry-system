/* 申込窓口の画面。サーバー(api.php)が正で、この側は「表を描く・集める」だけ。
   検証はサーバーで行い、返ってきたエラーをそのまま出す（規則の二重実装を避ける）。 */
"use strict";

const $ = (sel) => document.querySelector(sel);
const $$ = (sel, root) => Array.from((root || document).querySelectorAll(sel));

let S = null;        // state: {school, members, tournament, limits, year, entries, meta, deadlines, shinsa}
let currentTab = "advisors";
let entryDirty = false;
let shinsaDirty = false;
/* いま開いている申込。hub＝ログインのあとの選ぶ画面、taikai＝大会申込、shinsa＝級位審査申込。
   読み込み直しても同じ画面に戻れるよう、URL の # に持つ */
let mode = "hub";
const MODE_FIRST_TAB = { taikai: "advisors", shinsa: "shinsa" };

/* 送るときは「この画面が読み込んだ学校」を添える。管理者が学校として入る機能で
   同じブラウザのまま別の学校に切り替わっていたら、サーバーが保存を断る
   （前の学校の画面から保存して、切り替えた先の学校に書き込むのを防ぐ） */
function postHeaders(extra) {
  const h = { "X-Entry-Api": "1", ...extra };
  if (S && S.school && S.school.id) h["X-Entry-School"] = S.school.id;
  return h;
}

async function api(action, data) {
  const res = await fetch(data === undefined ? `api.php?action=${action}` : "api.php", {
    method: data === undefined ? "GET" : "POST",
    headers: data === undefined ? {} : postHeaders({ "Content-Type": "application/json" }),
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

/* 保存は通ったが確認してほしいこと。errbox と違い**赤くせず・戻り値も見ない**
   （止めるものではないため）。サーバーが warnings を返さなければ黙って消える。 */
function showNotice(boxSel, msgs) {
  const box = $(boxSel);
  if (msgs && msgs.length) {
    box.textContent = msgs.map((m) => "・" + m).join("\n");
    box.classList.remove("hidden");
  } else {
    box.classList.add("hidden");
  }
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
  const acting = S && S.acting_admin;
  const body = await api("logout", {});
  // 管理者が学校として入っていたときは、学校の操作だけ終えて管理画面へ戻る
  if (acting && body.admin) { location.href = "admin.html"; return; }
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
  // 管理者が学校として入っているときは、取り違えないよう画面の上に帯を出す
  $("#acting-name").textContent = S.school.name;
  $("#acting").classList.toggle("hidden", !S.acting_admin);
  $("#btn-logout").textContent = S.acting_admin ? "管理画面に戻る" : "ログアウト";
  renderAdvisors();
  renderMembers();
  renderEntry();
  renderUploads();
  renderHub();
  renderShinsa();
  renderSform();
  renderSteps();
  setDirty(false);
  setShinsaDirty(false);
  applyMode();
}

/* 'YYYY-MM-DD' → 「10月3日（あと5日）」。今日は**サーバーの日付**で数える
   （端末の時計がずれていても、締め切りの判定と表示が食い違わないように） */
function dueText(d, prefix = "締切 ") {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(d || "");
  if (!m) return d ? prefix + d : "";
  const t = /^(\d{4})-(\d{2})-(\d{2})$/.exec(S.today || "");
  const today = t ? new Date(+t[1], +t[2] - 1, +t[3]) : new Date();
  const days = Math.round((new Date(+m[1], +m[2] - 1, +m[3]) - today) / 86400000);
  return `${prefix}${+m[2]}月${+m[3]}日` +
    (days > 0 ? `（あと${days}日）` : days === 0 ? "（今日まで）" : "（締切済み）");
}

function mdText(d) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(d || "");
  return m ? `${+m[2]}月${+m[3]}日` : (d || "");
}

function renderDue() {
  const el = $("#hd-due");
  const d = mode === "taikai" ? (S.tournament && S.tournament.deadline)
          : mode === "shinsa" ? (S.shinsa.round && S.shinsa.round.deadline) : "";
  el.textContent = dueText(d);
  el.classList.toggle("hidden", !d);
}

function renderTitle() {
  const sr = S.shinsa.round;
  $("#hd-title").textContent =
      mode === "taikai" ? (S.tournament
        ? `令和${S.year}年度 ${S.tournament.name} — 参加申込` : "参加申込（受付中の大会なし）")
    : mode === "shinsa" ? (sr ? `${sr.title} — 受審申込` : "級位審査申込（受付中の回なし）")
    : "申込メニュー";
}

/* ---------- 選ぶ画面と、大会／審査の切り替え ---------- */

function renderHub() {
  const t = S.tournament;
  $("#hub-taikai-name").textContent = t ? t.name : "受付中の大会はありません";
  $("#hub-taikai-due").textContent = t ? dueText(t.deadline) : "";
  const r = S.shinsa.round;
  $("#hub-shinsa-name").textContent = r
    ? `第${r.no}回 級位審査会（${mdText(r.date)}）` : "受付中の級位審査はありません";
  $("#hub-shinsa-due").textContent = r ? dueText(r.deadline) : "";
  const dl = S.deadlines || [];
  $("#dl-none").classList.toggle("hidden", dl.length > 0);
  // スマートフォンでも読めるよう 2 列（何の締め切りか／いつまでか）にまとめる
  $("#dl-list").innerHTML = dl.map((d) => `
    <tr>
      <td><span class="badge">${esc(d.kind)}</span> ${esc(d.name)}${d.date
        ? ` <span class="aux">（${esc(mdText(d.date))}）</span>` : ""}</td>
      <td class="dl-due">${esc(dueText(d.deadline, ""))}${d.open ? ' <span class="badge new">受付中</span>' : ""}</td>
    </tr>`).join("");
}

function modeFromHash() {
  const h = location.hash.replace("#", "");
  return h === "taikai" || h === "shinsa" ? h : "hub";
}

/* 画面を mode に合わせる（タブの出し入れ・見出し・締め切り） */
function applyMode() {
  mode = modeFromHash();
  const inHub = mode === "hub";
  $("#tab-hub").classList.toggle("hidden", !inHub);
  $("#tabs").classList.toggle("hidden", inHub);
  $("#btn-menu").classList.toggle("hidden", inHub);
  $$(".step").forEach((b) => b.classList.toggle("hidden", !b.dataset.modes.split(" ").includes(mode)));
  renderTitle();
  renderDue();
  if (inHub) {
    ["advisors", "members", "entry", "form", "shinsa", "sform"].forEach((t) =>
      $("#tab-" + t).classList.add("hidden"));
    currentTab = "hub";
    updateSavebar();
    return;
  }
  const btn = $(`.step[data-tab="${currentTab}"]`);
  const ok = btn && btn.dataset.modes.split(" ").includes(mode);
  switchTab(ok ? currentTab : MODE_FIRST_TAB[mode]);
}

/* 未保存の入力があれば、画面を移る前に確かめる */
function leaveOk() {
  if (!entryDirty && !shinsaDirty) return true;
  return confirm("保存していない変更があります。移ると消えます。よろしいですか？");
}

function goMode(m) {
  if (!leaveOk()) return;
  if (entryDirty || shinsaDirty) {
    // 移るなら、未保存の入力は捨てて保存済みの内容で描き直す
    renderEntry(); renderShinsa(); setDirty(false); setShinsaDirty(false);
  }
  location.hash = m === "hub" ? "" : m;   // hashchange で applyMode が走る
  if (m === "hub") applyMode();
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
  const sh = S.shinsa.round ? (S.shinsa.apps || []).length : 0;
  set("#st-shinsa", sh, "名");
  const sf = $("#st-sform");
  sf.textContent = sh > 0 ? "作成できます" : "—";
  sf.classList.toggle("done", false);
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
    <td><input type="text" class="m-no num" placeholder="No." value="${esc(m.display_order)}"></td>
    <td><input type="text" class="m-name" placeholder="氏名" value="${esc(m.name)}"></td>
    <td><select class="m-sex">
      <option value="">—</option>
      <option ${m.sex === "男子" ? "selected" : ""}>男子</option>
      <option ${m.sex === "女子" ? "selected" : ""}>女子</option></select></td>
    <td><select class="m-grade">
      <option value="">—</option>
      ${[1, 2, 3].map((g) => `<option ${String(m.grade) === String(g) ? "selected" : ""}>${g}</option>`).join("")}
    </select></td>
    <td><input type="text" class="m-dob num" placeholder="2010/04/01" value="${esc(m.dob)}"></td>
    <td><input type="text" class="m-jkf num" placeholder="JKF番号" value="${esc(m.jkf_no)}"></td>
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
  // エラーで戻るときも通る位置に置く（warnings が無ければ前回の警告が消える）。
  // #entry-warn は #entry-body の中だが loadState は class を切り替えるだけなので消えない
  showNotice("#entry-warn", body.warnings);
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
    + "どうしてもPDFにできない場合は、加村までメールをください";
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
      { method: "POST", headers: postHeaders(), body: fd });
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

/* ---------- 級位審査の申込 ----------
   選択肢（区分・級・形・年月）はサーバーの shinsa.php が持っていて、S.shinsa.options で
   届く。ここで選択肢を書き足さないこと（決まりを 2 か所に書かない） */

function shOpt(list, current, label = (v) => v) {
  return '<option value=""></option>' + list.map((v) =>
    `<option value="${esc(v)}" ${String(v) === String(current ?? "") ? "selected" : ""}>${esc(label(v))}</option>`).join("");
}

function shKubun(id) {
  return S.shinsa.options.kubun.find((k) => k.id === id);
}

function shinsaKubunList() {
  // 形講習会・レベルチェックは第1回・第3回だけ（回の設定 lc。どの区分が該当かはサーバーの lc_round）
  return S.shinsa.options.kubun.filter((k) => !k.lc_round || S.shinsa.round.lc);
}

function shinsaRow(a = {}) {
  const tr = document.createElement("tr");
  const o = S.shinsa.options;
  const names = (S.members || []).map((m) => m.name);
  const memberLabel = (n) => {
    const m = (S.members || []).find((x) => x.name === n);
    return m ? `${n}（${m.sex === "女子" ? "女" : "男"}・${m.grade}年）` : `${n}（名簿にありません）`;
  };
  // 名簿から消えた人も行は残す（保存するとエラーで知らせる）
  const nameList = a.name && !names.includes(a.name) ? [a.name, ...names] : names;
  const kubun = shinsaKubunList().map((k) => k.id);
  if (a.kubun && !kubun.includes(a.kubun)) kubun.push(a.kubun);
  const kyus = a.kubun && shKubun(a.kubun) ? shKubun(a.kubun).kyu : [];
  tr.innerHTML = `
    <td class="sh-no num"></td>
    <td><select class="s-name">${shOpt(nameList, a.name, memberLabel)}</select></td>
    <td><select class="s-kubun">${shOpt(kubun, a.kubun, (id) => (shKubun(id) || {}).label || id)}</select></td>
    <td><select class="s-kyu" ${kyus.length ? "" : "disabled"}>${shOpt(kyus, a.kyu)}</select></td>
    <td><select class="s-kata">${shOpt(o.kata, a.kata)}</select></td>
    <td class="ym"><select class="s-years">${shOpt(o.years, a.years)}</select>年
      <select class="s-months">${shOpt(o.months, a.months)}</select>か月</td>
    <td class="del"><button class="mv up" title="上へ">↑</button><button class="mv dn" title="下へ">↓</button><button class="x" title="この行を外す">×</button></td>`;
  tr.querySelector(".x").onclick = () => { tr.remove(); shinsaChanged(); };
  tr.querySelector(".up").onclick = () => { if (tr.previousElementSibling) tr.parentNode.insertBefore(tr, tr.previousElementSibling); shinsaChanged(); };
  tr.querySelector(".dn").onclick = () => { if (tr.nextElementSibling) tr.parentNode.insertBefore(tr.nextElementSibling, tr); shinsaChanged(); };
  // 区分を変えたら、選べる級を入れ替える（レベルチェックは級・段）
  tr.querySelector(".s-kubun").onchange = (e) => {
    const k = shKubun(e.target.value);
    const sel = tr.querySelector(".s-kyu");
    const cur = sel.value;
    sel.innerHTML = shOpt(k ? k.kyu : [], k && k.kyu.includes(cur) ? cur : "");
    sel.disabled = !k;
  };
  return tr;
}

function renumberShinsa() {
  $$("#sh-table tr").forEach((tr, i) => { tr.querySelector(".sh-no").textContent = i + 1; });
  const n = $$("#sh-table .s-name").filter((s) => s.value !== "").length;
  $("#sh-count").textContent = n ? `${n}名` : "";
}

function shinsaChanged() {
  renumberShinsa();
  setShinsaDirty(true);
}

function setShinsaDirty(v) {
  shinsaDirty = v;
  $("#sh-dirty").textContent = v ? "未保存の変更があります" : "";
}

function renderShinsa() {
  const sh = S.shinsa;
  const r = sh.round;
  $("#sh-none").classList.toggle("hidden", !!r);
  $("#sh-body").classList.toggle("hidden", !r);
  if (!r) return;
  $("#sh-aux").textContent = `審査期日 ${mdText(r.date)} ／ ${dueText(r.deadline)}`;
  $("#sh-lc-hint").innerHTML = r.lc
    ? "黒帯で<b>レベルチェックテスト</b>を受ける部員は「レベルチェックテスト（黒帯）」、審査は受けずに<b>形講習会だけ</b>出る部員は「形講習会のみ」を選びます（どちらも無料。形講習会のみの人は審査用紙を刷りません）。"
    : "この回は形講習会・レベルチェックテストはありません（第1回・第3回だけ）。";
  const tb = $("#sh-table");
  tb.innerHTML = "";
  (sh.apps || []).forEach((a) => tb.appendChild(shinsaRow(a)));
  if (!tb.children.length && sh.can_edit) tb.appendChild(shinsaRow());
  renumberShinsa();
  const locked = !sh.can_edit;
  $("#sh-locked").classList.toggle("hidden", !locked);
  $("#sh-locked").textContent = locked
    ? "締め切りを過ぎたため、申込は直せません。直す必要がある場合は審査部会へ連絡してください。" : "";
  $$("#sh-table select, #sh-table button").forEach((el) => { if (locked) el.disabled = true; });
  $("#sh-add").classList.toggle("hidden", locked);
  $("#sh-save").classList.toggle("hidden", locked);
}

async function saveShinsa() {
  const apps = $$("#sh-table tr").map((tr) => ({
    name: tr.querySelector(".s-name").value,
    kubun: tr.querySelector(".s-kubun").value,
    kyu: tr.querySelector(".s-kyu").value,
    kata: tr.querySelector(".s-kata").value,
    years: tr.querySelector(".s-years").value,
    months: tr.querySelector(".s-months").value,
  })).filter((a) => a.name !== "" || a.kubun !== "" || a.kata !== "");
  // 氏名だけ空の行は、黙って捨てずに知らせる
  if (apps.some((a) => a.name === "")) {
    showErrors("#sh-errors", { error: "氏名を選んでいない行があります（要らない行は × で外してください）" });
    return;
  }
  const body = await api("shinsa_save", { round_id: S.shinsa.round.id, apps });
  if (showErrors("#sh-errors", body.ok ? {} : body)) return;
  await loadState();
  flash("#sh-msg", "保存しました");
}

function renderSform() {
  const sh = S.shinsa;
  const r = sh.round;
  $("#sf-none").classList.toggle("hidden", !!r);
  $("#sf-body").classList.toggle("hidden", !r);
  if (!r) return;
  const apps = sh.apps || [];
  const pages = sh.form_pages || 1;   // 1 枚に何人入るかはサーバーが決める
  $("#sf-forms").innerHTML = !apps.length
    ? '<span class="aux">先に「3. 審査申込」で申込を保存してください。</span>'
    : Array.from({ length: pages }, (_, i) =>
        `<a class="btn primary" href="api.php?action=shinsa_form&page=${i + 1}">申込書をダウンロード${pages > 1 ? `（${i + 1}枚目）` : ""}</a>`).join(" ")
      + (pages > 1 ? '<span class="aux">　1 枚に 20 名まで入るので、2 枚に分かれます。</span>' : "");
  // 審査用紙を刷るのは用紙のある区分だけ（形講習会のみの人は刷らない）
  const withSheet = apps.filter((a) => (shKubun(a.kubun) || {}).sheet);
  const sheet = $("#sf-sheets");
  sheet.classList.toggle("hidden", !withSheet.length);
  $("#sf-sheets-count").textContent = withSheet.length
    ? `${withSheet.length}名 × 3枚 ＝ ${withSheet.length * 3}枚`
      + (withSheet.length < apps.length ? `（形講習会のみの ${apps.length - withSheet.length} 名は刷りません）` : "")
    : "";
  const numbered = withSheet.length > 0 && withSheet.every((a) => a.no);
  renderReceipts();
  $("#sf-no-hint").textContent = !apps.length ? ""
    : numbered ? "受審番号が振られています。番号も印字されます。"
    : "受審番号は締め切りのあとに審査部会が振ります。番号が振られる前に印刷すると、番号の欄は空いたままになります（手で書いてもかまいません）。";
}

/* 領収書（審査部会が管理画面で出せるようにしてから）。合否は顧問が選ぶ。
   受審料の金額と、どの区分に領収書を出すか（fee）はサーバーの shinsa.php が決める */
const RESULT_OPTS = [["", "—"], ["pass", "合格"], ["fail", "不合格"], ["absent", "欠席"]];

function renderReceipts() {
  const sh = S.shinsa;
  const r = sh.round;
  const on = !!(r && r.receipt_on);
  $("#rc-off").classList.toggle("hidden", on);
  $("#rc-on").classList.toggle("hidden", !on);
  $("#rc-aux").textContent = on ? `領収書の日付 ${mdText(r.receipt_date)}` : "";
  if (!on) return;
  const fee = sh.options.fee;
  $("#rc-pass").textContent = fee.pass.toLocaleString();
  $("#rc-fail").textContent = fee.fail.toLocaleString();
  const feeOf = Object.fromEntries(sh.options.kubun.map((k) => [k.id, k.fee]));
  const list = (sh.apps || []).filter((a) => feeOf[a.kubun]);
  $("#rc-table").innerHTML = list.map((a) => `
    <tr data-name="${esc(a.name)}">
      <td>${esc(a.name)}</td>
      <td>${esc(a.sex === "女子" ? "女子" : "男子")}${esc((shKubun(a.kubun) || {}).short || "")}</td>
      <td>${RESULT_OPTS.map(([v, l]) => `<label style="margin-right:10px; white-space:nowrap">
        <input type="radio" name="rc-${esc(a.name)}" value="${v}" ${(a.result || "") === v ? "checked" : ""}> ${l}</label>`).join("")}</td>
    </tr>`).join("");
  countReceipts();
}

function countReceipts() {
  const picked = $$("#rc-table input[type=radio]:checked").map((i) => i.value);
  const n = picked.filter((v) => v === "pass" || v === "fail").length;
  const none = picked.filter((v) => v === "").length;
  $("#rc-count").textContent = `領収書 ${n}枚` + (none ? `（まだ選んでいない人が ${none} 名）` : "");
}

async function printReceipts() {
  const results = {};
  $$("#rc-table tr").forEach((tr) => {
    const c = tr.querySelector("input[type=radio]:checked");
    results[tr.dataset.name] = c ? c.value : "";
  });
  // 押した瞬間にタブを開いておく（保存を待ってから開くと、ポップアップとして止められることがある）
  const w = window.open("", "_blank");
  const body = await api("shinsa_results", { round_id: S.shinsa.round.id, results });
  if (showErrors("#rc-errors", body.ok ? {} : body)) { if (w) w.close(); return; }
  if (w) w.location.href = "shinsa_receipt.html";
  else location.href = "shinsa_receipt.html";
  await loadState();
  flash("#rc-msg", "保存しました");
}

/* ---------- 保存バー・未保存の管理 ---------- */

function setDirty(v) {
  entryDirty = v;
  const el = $("#dirty");
  el.classList.toggle("clean", !v);
  $("#dirty-text").textContent = v ? "未保存の変更があります" : "保存済み";
}

function updateSavebar() {
  const show = mode === "taikai" && currentTab === "entry" && S && !!S.tournament;
  $("#savebar").classList.toggle("hidden", !show);
}

/* ---------- タブと起動 ---------- */

function switchTab(name) {
  currentTab = name;
  $$(".step").forEach((b) => {
    if (b.dataset.tab === name) b.setAttribute("aria-current", "page");
    else b.removeAttribute("aria-current");
  });
  ["advisors", "members", "entry", "form", "shinsa", "sform"].forEach((t) =>
    $("#tab-" + t).classList.toggle("hidden", t !== name));
  updateSavebar();
}

window.addEventListener("DOMContentLoaded", async () => {
  $$(".step").forEach((b) => (b.onclick = () => switchTab(b.dataset.tab)));
  $("#btn-login").onclick = () => doLogin().catch((e) => showErrors("#login-error", { error: e.message }));
  $("#login-password").addEventListener("keydown", (e) => { if (e.key === "Enter") $("#btn-login").click(); });
  $("#btn-logout").onclick = () => { if (leaveOk()) { setDirty(false); setShinsaDirty(false); doLogout(); } };
  $("#btn-menu").onclick = () => goMode("hub");
  $$(".hub-card").forEach((b) => (b.onclick = () => goMode(b.dataset.mode)));
  window.addEventListener("hashchange", () => { if (S) applyMode(); });

  $("#sh-add").onclick = () => { $("#sh-table").appendChild(shinsaRow()); shinsaChanged(); };
  $("#sh-save").onclick = () => saveShinsa().catch((e) => showErrors("#sh-errors", { error: e.message }));
  $("#tab-shinsa").addEventListener("change", () => shinsaChanged());
  $("#rc-table").addEventListener("change", countReceipts);
  $("#rc-print").onclick = () => printReceipts().catch((e) => showErrors("#rc-errors", { error: e.message }));

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
    if (entryDirty || shinsaDirty) { e.preventDefault(); e.returnValue = ""; }
  });

  await initLogin().catch(() => {});
  await loadState().catch(() => { $("#view-login").classList.remove("hidden"); });
});
