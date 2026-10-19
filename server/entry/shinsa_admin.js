/* 級位審査の管理画面。サーバー（shinsa_admin_api.php → shinsa.php）が正で、ここは描くだけ。
   受審番号の振り方・合格の基準・領収書の金額などの決まりは shinsa.php にある */
"use strict";

const $ = (s) => document.querySelector(s);
const $$ = (s, r) => Array.from((r || document).querySelectorAll(s));
let A = null;          // state: {rounds, today, as_admin, groups, kubun}
let cur = null;        // 開いている回の id
let D = null;          // 開いている回: {round, status, group_counts}
let roundsDirty = false;

async function api(action, data, query = "") {
  const res = await fetch(data === undefined ? `shinsa_admin_api.php?action=${action}${query}` : "shinsa_admin_api.php", {
    method: data === undefined ? "GET" : "POST",
    headers: data === undefined ? {} : { "Content-Type": "application/json", "X-Entry-Api": "1" },
    body: data === undefined ? undefined : JSON.stringify({ action, ...data }),
  });
  let b = null;
  try { b = await res.json(); } catch { /* 下で拾う */ }
  if (!b) throw new Error(`サーバーから読めない応答（HTTP ${res.status}）`);
  return b;
}

function esc(s) {
  return String(s ?? "").replace(/[&<>"']/g, (c) =>
    ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}

function showErrors(sel, body) {
  const box = $(sel);
  const msgs = body.errors || (body.error ? [body.error] : []);
  box.textContent = msgs.map((m) => "・" + m).join("\n");
  box.classList.toggle("hidden", !msgs.length);
  return msgs.length > 0;
}

function flash(sel, t) { const e = $(sel); e.textContent = t; setTimeout(() => { e.textContent = ""; }, 4000); }

function mdText(d) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(d || "");
  return m ? `${+m[1]}年${+m[2]}月${+m[3]}日` : (d || "");
}

/* ---------- 読み込み ---------- */

async function load() {
  const b = await api("state");
  if (!b.ok) {
    $("#view-login").classList.remove("hidden");
    $("#view-app").classList.add("hidden");
    return;
  }
  A = b;
  $("#view-login").classList.add("hidden");
  $("#view-app").classList.remove("hidden");
  $("#asadmin").classList.toggle("hidden", !A.as_admin);
  $("#btn-pw").classList.toggle("hidden", A.as_admin);
  $("#rc-fee").textContent = `合格 ${A.fee.pass.toLocaleString()} 円・不合格 ${A.fee.fail.toLocaleString()} 円`;
  renderRounds();
  // 開く回: 前に開いていた回 → 学校に見せている回 → これからの回 → 最後の回
  const ids = A.rounds.map((r) => r.id);
  if (!ids.includes(cur)) {
    const open = A.rounds.find((r) => r.open);
    const next = A.rounds.find((r) => r.date >= A.today);
    cur = (open || next || A.rounds[A.rounds.length - 1] || {}).id || null;
  }
  await openRound(cur);
}

/* ---------- 1. 回 ---------- */

function roundRow(r = {}) {
  const tr = document.createElement("tr");
  tr.dataset.id = r.id || "";
  const n = (r.counts || {}).total || 0;
  tr.innerHTML = `
    <td>令和<input type="number" class="r-year num" min="1" max="99" value="${esc(r.year)}">年度</td>
    <td>第<input type="number" class="r-no num" min="1" max="9" value="${esc(r.no)}">回</td>
    <td><input type="date" class="r-date" value="${esc(r.date)}"></td>
    <td><input type="date" class="r-dl" value="${esc(r.deadline)}"></td>
    <td class="c"><input type="checkbox" class="r-lc" ${r.lc ? "checked" : ""}></td>
    <td class="c"><input type="checkbox" class="r-open" ${r.open ? "checked" : ""}></td>
    <td class="c num">${r.id ? `${n}名` : '<span class="tag">新</span>'}</td>
    <td style="white-space:nowrap">${r.id ? '<button class="btn r-view" style="padding:3px 12px">開く</button>' : ""}
      ${n ? "" : '<button class="x" title="この回を外す">×</button>'}</td>`;
  const x = tr.querySelector(".x");
  if (x) x.onclick = () => { tr.remove(); roundsDirty = true; };
  const v = tr.querySelector(".r-view");
  if (v) v.onclick = () => openRound(r.id);
  return tr;
}

function renderRounds() {
  const tb = $("#r-body");
  tb.innerHTML = "";
  A.rounds.forEach((r) => tb.appendChild(roundRow(r)));
  if (!A.rounds.length) tb.appendChild(roundRow());
  markSelected();
  roundsDirty = false;
}

function markSelected() {
  $$("#r-body tr").forEach((tr) => tr.classList.toggle("sel", !!cur && tr.dataset.id === cur));
}

/* 次の年度の 3 回ぶんの行を足す（年度は今ある最後の年度の次。期日と締め切りは空けておく）。
   形講習・レベルチェックは第 1 回・第 3 回だけ（申込書の注意事項⑤） */
function addYear() {
  const years = $$("#r-body .r-year").map((i) => +i.value).filter((v) => v > 0);
  const t = /^(\d{4})/.exec(A.today);
  // 4 月からが新しい年度。いまの年度（令和）＝ 西暦 − 2018（1〜3 月は前の年）
  const m = +A.today.slice(5, 7);
  const nowNendo = t ? (+t[1] - 2018 - (m < 4 ? 1 : 0)) : 8;
  const y = years.length ? Math.max(...years) + 1 : nowNendo;
  [1, 2, 3].forEach((no) => $("#r-body").appendChild(roundRow({ year: y, no, lc: no !== 2 })));
  roundsDirty = true;
}

async function saveRounds() {
  const rounds = $$("#r-body tr").map((tr) => ({
    id: tr.dataset.id,
    year: tr.querySelector(".r-year").value,
    no: tr.querySelector(".r-no").value,
    date: tr.querySelector(".r-date").value,
    deadline: tr.querySelector(".r-dl").value,
    lc: tr.querySelector(".r-lc").checked,
    open: tr.querySelector(".r-open").checked,
  }));
  const b = await api("save_rounds", { rounds });
  if (showErrors("#r-errors", b.ok ? {} : b)) return;
  flash("#r-msg", "保存しました");
  await load();
}

/* ---------- 2. 選んだ回 ---------- */

async function openRound(id) {
  cur = id;
  markSelected();
  $("#d-none").classList.toggle("hidden", !!id);
  $("#d-body").classList.toggle("hidden", !id);
  if (!id) { $("#d-title").textContent = "—"; $("#d-aux").textContent = ""; return; }
  const b = await api("round", undefined, `&round_id=${encodeURIComponent(id)}`);
  if (showErrors("#d-errors", b.ok ? {} : b)) return;
  D = b;
  renderRound();
}

function renderRound() {
  const r = D.round;
  $("#d-title").textContent = r.title;
  $("#d-aux").textContent = `審査期日 ${mdText(r.date)}　締め切り ${mdText(r.deadline)}`
    + (r.deadline < A.today ? "（締切済み）" : "") + (r.open ? "　学校に見せています" : "　学校には見せていません");

  // 人数（組ごと）
  const gc = D.group_counts;
  const total = Object.values(gc).reduce((s, n) => s + n, 0);
  $("#d-sum").innerHTML = `<div class="card-n"><div class="lbl">申込</div><div class="v">${total}<small>名</small></div></div>`
    + A.groups.map((g) => `<div class="card-n"><div class="lbl">${esc(g.label)}</div>
        <div class="v">${gc[g.key] || 0}<small>名</small></div></div>`).join("");

  // 受審番号
  const numbered = D.status.reduce((s, x) => s + (x.numbered || 0), 0);
  const toNumber = D.status.reduce((s, x) => s + (x.to_number || 0), 0);   // 形講習会のみは振らない
  const st = $("#n-status");
  st.innerHTML = !total ? "まだ申込がありません。"
    : !r.numbered_at ? '<span class="tag warn">まだ振っていません</span>'
    : numbered < toNumber ? `<span class="tag warn">番号の無い人が ${toNumber - numbered} 名います</span>（振ったあとに申込が増えた。振り直してください）`
    : `<span class="tag on">振ってあります</span> ${esc(r.numbered_at)}`;
  $("#n-run").textContent = r.numbered_at ? "受審番号を振り直す" : "受審番号を振る";
  $("#n-run").disabled = !total;
  $("#x-day").href = `shinsa_admin_api.php?action=day&round_id=${encodeURIComponent(r.id)}`;

  // 領収書
  $("#rc-date").value = r.receipt_date || r.date || "";
  $("#rc-state").textContent = r.receipt_on ? "顧問の画面に出ています" : "止めています";
  $("#rc-state").classList.toggle("on", !!r.receipt_on);
  $("#rc-toggle").textContent = r.receipt_on ? "止める" : "出せるようにする";

  renderStatus();
}

function renderStatus() {
  const only = $("#only-applied").checked;
  const rows = D.status.filter((s) => !only || s.total);
  const cell = (n) => `<td class="c num ${n ? "" : "zero"}">${n || "—"}</td>`;
  $("#st-body").innerHTML = rows.map((s) => `
    <tr class="${s.total ? "" : "none"}">
      <td>${esc(s.name)}</td>
      ${cell(s.jokyu_F)}${cell(s.jokyu_M)}${cell(s.shokyu_F)}${cell(s.shokyu_M)}${cell(s.lc)}${cell(s.koshu)}
      <td class="c num"><b>${s.total || "—"}</b></td>
      <td class="c">${!s.to_number ? "" : (s.numbered || 0) >= s.to_number ? '<span class="tag on">済</span>' : '<span class="tag">未</span>'}</td>
      <td class="c num">${s.receipts ? `${s.receipts}枚` : ""}</td>
    </tr>`).join("");
  const sum = (k) => D.status.reduce((t, s) => t + (s[k] || 0), 0);
  const applied = D.status.filter((s) => s.total).length;
  $("#st-foot").innerHTML = `<tr><td>計（${applied}校）</td>${["jokyu_F", "jokyu_M", "shokyu_F", "shokyu_M", "lc", "koshu", "total"]
    .map((k) => `<td class="c num">${sum(k)}</td>`).join("")}<td></td><td class="c num">${sum("receipts") || ""}</td></tr>`;
  $("#st-count").textContent = `申込 ${applied}校 ／ 全 ${D.status.length}校`;
}

async function runNumber() {
  const r = D.round;
  if (r.deadline >= A.today && !confirm("まだ締め切り前です。このあと申込が増えたら振り直しが要ります。振りますか？")) return;
  if (r.numbered_at && !confirm("受審番号を振り直します。すでに審査用紙を印刷した学校があれば、番号が変わることがあります。よろしいですか？")) return;
  const b = await api("number", { round_id: r.id });
  if (showErrors("#d-errors", b.ok ? {} : b)) return;
  flash("#n-msg", "振りました");
  await load();
}

async function toggleReceipt() {
  const r = D.round;
  const on = !r.receipt_on;
  if (on && !confirm(`領収書の日付を ${mdText($("#rc-date").value)} にして、顧問の画面に領収書を出します。よろしいですか？`)) return;
  const b = await api("receipt", { round_id: r.id, on, date: $("#rc-date").value });
  if (showErrors("#d-errors", b.ok ? {} : b)) return;
  flash("#rc-msg", on ? "出せるようにしました" : "止めました");
  await load();
}

async function saveReceiptDate() {
  const r = D.round;
  const b = await api("receipt", { round_id: r.id, on: !!r.receipt_on, date: $("#rc-date").value });
  if (showErrors("#d-errors", b.ok ? {} : b)) return;
  flash("#rc-msg", "日付を保存しました");
  D.round.receipt_date = $("#rc-date").value;
}

/* ---------- 起動 ---------- */

window.addEventListener("DOMContentLoaded", async () => {
  $("#btn-login").onclick = async () => {
    const b = await api("login", { password: $("#pw").value });
    if (showErrors("#login-error", b.ok ? {} : b)) return;
    await load();
  };
  $("#pw").addEventListener("keydown", (e) => { if (e.key === "Enter") $("#btn-login").click(); });
  $("#btn-logout").onclick = async () => {
    if (roundsDirty && !confirm("回の設定に保存していない変更があります。ログアウトしますか？")) return;
    const b = await api("logout", {});
    // 専門部の管理者として入っていたときは、そのログインは残るので管理画面へ戻す
    if (b.still_admin) { location.href = "admin.html"; return; }
    A = null; $("#view-app").classList.add("hidden");
    $("#view-login").classList.remove("hidden"); $("#pw").value = "";
  };
  $("#btn-pw").onclick = () => $("#pw-pane").classList.toggle("hidden");
  $("#pw-save").onclick = async () => {
    const b = await api("set_password", { password: $("#new-pw").value });
    if (showErrors("#pw-errors", b.ok ? {} : b)) return;
    $("#new-pw").value = ""; flash("#pw-msg", "変更しました");
  };
  $("#r-add").onclick = () => { $("#r-body").appendChild(roundRow()); roundsDirty = true; };
  $("#r-add-year").onclick = addYear;
  $("#r-body").addEventListener("change", () => { roundsDirty = true; });
  $("#r-save").onclick = () => saveRounds().catch((e) => showErrors("#r-errors", { error: e.message }));
  $("#n-run").onclick = () => runNumber().catch((e) => showErrors("#d-errors", { error: e.message }));
  $("#rc-toggle").onclick = () => toggleReceipt().catch((e) => showErrors("#d-errors", { error: e.message }));
  $("#rc-date").onchange = () => saveReceiptDate().catch((e) => showErrors("#d-errors", { error: e.message }));
  $("#only-applied").onchange = renderStatus;
  window.addEventListener("beforeunload", (e) => { if (roundsDirty) { e.preventDefault(); e.returnValue = ""; } });
  await load().catch(() => $("#view-login").classList.remove("hidden"));
});
