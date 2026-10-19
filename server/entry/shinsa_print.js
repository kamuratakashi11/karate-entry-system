/* 審査用紙の印刷。申込の内容（api.php の state）を、審査部会の様式の画像に重ねる。
   どの区分がどの用紙かはサーバー（shinsa.php の shinsa_kubun の sheet）が決める。
   ここが持つのは「用紙のどこに書くか」（位置・mm）だけ。位置は様式の PDF の
   文字と下線の座標から測った（2026-10-14。tools/make_shinsa_assets.py で地紙を作り直したら、
   ずれていないか印刷して確かめる） */
"use strict";

const SETS = 3;   // 1 人 3 枚（審査員 3 人ぶん）

/* 用紙ごとの書く位置（mm）。top は文字の箱（高さ 7mm）の上端。文字は箱の下端に乗る */
const LAYOUT = {
  jokyu: {                        // 級位審査用紙（受審級上級）
    no:     { left: 38.1, top: 49.9, width: 30, c: true },
    name:   { left: 88.2, top: 49.9, width: 60, c: true },
    male:   { left: 158.4, top: 50.4 }, female: { left: 174.1, top: 50.4 },
    kyu:    { left: 38.1, top: 67.2, width: 20, c: true },
    school: { left: 88.2, top: 67.2, width: 40, c: true },
    grade:  { left: 147.8, top: 67.2, width: 20, c: true },
    kata:   { left: 78.0, top: 84.6, width: 60, c: true },
  },
  shokyu: {                       // 級位審査用紙（受審級初級）。上級より上に詰めてある
    no:     { left: 38.1, top: 44.7, width: 30, c: true },
    name:   { left: 88.2, top: 44.7, width: 60, c: true },
    male:   { left: 158.4, top: 45.2 }, female: { left: 174.1, top: 45.2 },
    kyu:    { left: 38.1, top: 56.8, width: 20, c: true },
    school: { left: 88.2, top: 56.8, width: 40, c: true },
    grade:  { left: 147.8, top: 56.8, width: 20, c: true },
    kata:   { left: 82.9, top: 132.2, width: 60, c: false },   // 「２ 基本形 形名」の横（下線なし）
  },
  lc: {                           // レベルチェックテスト審査用紙
    no:     { left: 38.1, top: 49.9, width: 30, c: true },
    name:   { left: 88.2, top: 49.9, width: 60, c: true },
    male:   { left: 158.4, top: 50.4 }, female: { left: 174.1, top: 50.4 },
    kyu:    { left: 38.1, top: 67.2, width: 20, c: true },     // 「現級段位 __ 級・段」
    school: { left: 88.2, top: 67.2, width: 40, c: true },
    grade:  { left: 147.8, top: 67.2, width: 20, c: true },
    kata:   { left: 98.0, top: 106.1, width: 60, c: true },
  },
};

function esc(s) {
  return String(s ?? "").replace(/[&<>"']/g, (c) =>
    ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}

function field(pos, text) {
  if (!pos || text === "" || text == null) return "";
  return `<div class="f${pos.c ? " c" : ""}" style="left:${pos.left}mm; top:${pos.top}mm; width:${pos.width}mm">${esc(text)}</div>`;
}

function mark(pos) {
  return pos ? `<div class="mark" style="left:${pos.left}mm; top:${pos.top}mm"></div>` : "";
}

/* 上級・初級の用紙は「現在級 __ 級」と刷ってあるので数字だけ（無級は「無」）。
   レベルチェックは「__ 級・段」なので、級か段かが分かるよう そのまま書く */
function kyuText(sheet, kyu) {
  if (sheet === "lc") return kyu;
  return String(kyu || "").replace(/級$/, "");
}

function page(sheet, a, schoolLabel, setNo) {
  const L = LAYOUT[sheet];
  return `<div class="page"><img src="shinsa/sheet_${sheet}.png" alt="">
    ${field(L.no, a.no)}
    ${field(L.name, a.name)}
    ${a.sex === "男子" ? mark(L.male) : a.sex === "女子" ? mark(L.female) : ""}
    ${field(L.kyu, kyuText(sheet, a.kyu))}
    ${field(L.school, schoolLabel)}
    ${field(L.grade, a.grade)}
    ${field(L.kata, a.kata)}
    <div class="set">${setNo} / ${SETS}</div>
  </div>`;
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
    const res = await fetch("api.php?action=state");
    body = await res.json();
  } catch {
    showError("読み込めませんでした。申込の画面からもう一度開いてください。");
    return;
  }
  if (!body.ok) {
    showError("ログインが切れています。申込の画面でログインしてから開いてください。");
    return;
  }
  const sh = body.shinsa;
  if (!sh.round) { showError("受付中の級位審査はありません。"); return; }
  if (!(sh.apps || []).length) { showError("まだ申込が保存されていません。「3. 審査申込」で保存してから開いてください。"); return; }

  const sheetOf = {};
  sh.options.kubun.forEach((k) => { sheetOf[k.id] = k.sheet; });
  // 用紙の無い区分（形講習会のみ）は刷らない
  const apps = (sh.apps || []).filter((a) => sheetOf[a.kubun]);
  if (!apps.length) { showError("審査用紙を刷る人がいません（形講習会のみの人は審査用紙を使いません）。"); return; }
  // 受審番号が全員に振られていれば番号順。まだなら学校が並べた順
  const numbered = apps.every((a) => a.no);
  const list = numbered
    ? [...apps].sort((x, y) => String(x.no).localeCompare(String(y.no), "ja", { numeric: true }))
    : apps;
  const html = [];
  for (let s = 1; s <= SETS; s++) {
    list.forEach((a) => html.push(page(sheetOf[a.kubun], a, sh.school_label, s)));
  }
  document.getElementById("pages").innerHTML = html.join("");
  document.getElementById("info").innerHTML =
    `<b>${esc(sh.round.title)}</b>　${esc(body.school.name)}　${apps.length}名 × ${SETS}枚 ＝ <b>${apps.length * SETS}枚</b>` +
    (numbered ? "" : "　<span style=\"color:#8a5a12\">（受審番号はまだ振られていません）</span>");
  const btn = document.getElementById("btn-print");
  btn.disabled = false;
  btn.onclick = () => window.print();
}

main();
