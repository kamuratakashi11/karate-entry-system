/* 領収書の印刷。顧問が選んだ合否（合格 pass ／不合格 fail）で地紙を選び、宛名・年度・日付を重ねる。
   金額・どの区分に出すかはサーバー（shinsa.php）が決める。ここが持つのは「どこに書くか」（mm）だけ。
   位置は領収書の PDF の文字の座標から測った（2026-10-14）。下の 1 枚は上の 1 枚の 144.4mm 下 */
"use strict";

const HALF = 144.4;   // 上の領収書 → 下の領収書のずれ（mm）
const POS = {
  name:  { left: 40.4, top: 43.0, width: 65.4 },   // 「＿＿＿＿ 様」
  nendo: { left: 65.6, top: 76.1, width: 10 },     // 「但し、令和 ＿ 年度 …」
  y:     { left: 51.8, top: 101.0, width: 10 },    // 「令和 ＿ 年 ＿ 月 ＿ 日」
  m:     { left: 67.3, top: 101.0, width: 10 },
  d:     { left: 85.7, top: 101.0, width: 10 },
};

function esc(s) {
  return String(s ?? "").replace(/[&<>"']/g, (c) =>
    ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}

function field(p, dy, text, cls = "") {
  return `<div class="f ${cls}" style="left:${p.left}mm; top:${p.top + dy}mm; width:${p.width}mm">${esc(text)}</div>`;
}

function receipt(a, dy, nendo, ymd) {
  return field(POS.name, dy, a.name, "name") + field(POS.nendo, dy, nendo)
    + field(POS.y, dy, ymd[0]) + field(POS.m, dy, ymd[1]) + field(POS.d, dy, ymd[2]);
}

function showError(msg) {
  const el = document.getElementById("err");
  el.textContent = msg;
  el.hidden = false;
  document.getElementById("info").textContent = "";
}

async function main() {
  let body = null;
  try {
    body = await (await fetch("api.php?action=state")).json();
  } catch {
    showError("読み込めませんでした。申込の画面からもう一度開いてください。");
    return;
  }
  if (!body.ok) { showError("ログインが切れています。申込の画面でログインしてから開いてください。"); return; }
  const sh = body.shinsa;
  if (!sh.round || !sh.round.receipt_on || !sh.receipt_reiwa) {
    showError("領収書はまだ出せません（審査のあとに審査部会が出せるようにします）。");
    return;
  }
  const feeOf = Object.fromEntries(sh.options.kubun.map((k) => [k.id, k.fee]));
  const list = (sh.apps || []).filter((a) => feeOf[a.kubun] && (a.result === "pass" || a.result === "fail"));
  if (!list.length) { showError("合格・不合格を選んだ人がいません。申込の画面で選んでから開いてください。"); return; }

  // 金額ごとに 2 人ずつ 1 枚に（地紙に金額が刷ってあるので、合格と不合格は別の紙）
  const html = [];
  for (const kind of ["pass", "fail"]) {
    const people = list.filter((a) => a.result === kind);
    for (let i = 0; i < people.length; i += 2) {
      const lower = people[i + 1];
      html.push(`<div class="page"><img src="shinsa/receipt_${kind}.png" alt="">
        ${receipt(people[i], 0, sh.round.year, sh.receipt_reiwa)}
        ${lower ? receipt(lower, HALF, sh.round.year, sh.receipt_reiwa) : '<div class="blank"></div>'}
      </div>`);
    }
  }
  document.getElementById("pages").innerHTML = html.join("");
  const fee = sh.options.fee;
  const nPass = list.filter((a) => a.result === "pass").length;
  const nFail = list.length - nPass;
  document.getElementById("info").innerHTML =
    `<b>${esc(sh.round.title)}</b>　${esc(body.school.name)}　領収書 <b>${list.length}枚</b>` +
    `（${fee.pass.toLocaleString()}円 ${nPass}枚・${fee.fail.toLocaleString()}円 ${nFail}枚）＝ 用紙 ${html.length}枚`;
  const btn = document.getElementById("btn-print");
  btn.disabled = false;
  btn.onclick = () => window.print();
}

main();
