"""SQLite → Sheets の書き戻し（切り戻し用の保険）。

    python tools/export_to_sheets.py            何が変わるかだけ見る（Sheets は読むだけ）
    python tools/export_to_sheets.py --write    実際に Sheets を書き換える

新人戦の受付中に新システムで致命傷が出たとき、**現行 Streamlit へ戻す**ための道具。
新システムで受け付けた内容（名簿・エントリー・顧問）を Sheets へ返し、Streamlit を
そのまま使える状態にする。二重受付はしない（案内は必ず片方だけにする）。

守ること:

- **パスワードは戻せない。** SQLite にはハッシュしか無い（設計どおり・逆算できない）。
  そのため v2_auth の `password` は **Sheets の現在値をそのまま残す**。新システムで
  管理者が再設定した学校は「Sheets 側の古いパスワード」に戻るので、**名指しで警告する**。
- **既定では書かない。** `--write` を付けたときだけ書く。
- **書く前に必ずスナップショットを取る**（取り込み側と同じ流儀。戻すときの戻し先になる）。
- **Sheets にしか無いものは触らない。** v1系・`*_backup`・SQLite に無い大会のタブ・
  学校ごとの見知らぬ項目は、そのまま残す（消すより残すほうが安全）。
- 書き方は app.py の `save_json` / `save_members_master` と同じ（clear → update）。
  形式もそのまま: v2_auth・v2_config・v2_entry_* は「キー＋JSON」の2列（見出し無し）、
  v2_members は見出し付きの8列。

値の組み立て（`build_*`）は Sheets に触らない純粋な関数にしてある。
往復の検算は tools/test_roundtrip.py が手元だけで行う。
"""

import argparse
import datetime
import hashlib
import hmac
import json
import os
import sqlite3
import sys

from migrate_from_sheets import (DATA_DIR, MEMBERS_COLS, SHEET_NAME, V2_PREFIX,
                                 open_spreadsheet, parse_kv, snapshot)

DB_PATH = os.path.join(DATA_DIR, "entry.sqlite3")


def password_matches(plain, stored):
    """Sheets の平文が、SQLite のハッシュと同じものか（＝新システムで変えていないか）"""
    try:
        algo, iters, salt, digest = str(stored).split("$")
        if algo != "pbkdf2_sha256":
            return False
        calc = hashlib.pbkdf2_hmac("sha256", str(plain).encode("utf-8"),
                                   bytes.fromhex(salt), int(iters)).hex()
        return hmac.compare_digest(calc, digest)
    except (ValueError, TypeError):
        return False


def load_db(db_path):
    con = sqlite3.connect(db_path)
    con.row_factory = sqlite3.Row
    out = {
        "schools": [dict(r) for r in con.execute("SELECT * FROM schools")],
        "members": [dict(r) for r in con.execute("SELECT * FROM members")],
        "entries": [dict(r) for r in con.execute("SELECT * FROM entries")],
        "config": {r["key"]: json.loads(r["value_json"]) for r in con.execute("SELECT * FROM config")},
    }
    con.close()
    return out


def build_auth(schools, live_auth):
    """v2_auth。Sheets の現在の中身に、新システムが持っている項目だけを重ねる。

    まるごと置き換えないのは、パスワードや、こちらが知らない項目（登録日時など）を
    落とさないため。戻り値は (キー→値, 警告の一覧)。
    """
    out = dict(live_auth)          # Sheets にしか無い学校もそのまま残る
    warn_no_pw, warn_changed = [], []
    for s in schools:
        sid = s["school_id"]
        base = dict(live_auth.get(sid) or {})
        if not isinstance(live_auth.get(sid), dict):
            base = {}
        raw = json.loads(s["raw_json"] or "{}")
        for k, v in raw.items():                 # 移行時の項目（password は入っていない）
            base.setdefault(k, v)
        # 新システムが書き換えている項目を上書きする
        base["base_name"] = s["base_name"]
        base["short_name"] = s["short_name"]
        base["school_no"] = s["school_no"]
        base["principal"] = s["principal"]
        base["advisors"] = json.loads(s["advisors_json"] or "[]")

        plain = (live_auth.get(sid) or {}).get("password") if isinstance(live_auth.get(sid), dict) else None
        if plain in (None, ""):
            base.setdefault("password", "")
            warn_no_pw.append(s["base_name"])
        else:
            base["password"] = plain             # ← 戻せないので Sheets の値を残す
            if not password_matches(plain, s["password_hash"]):
                warn_changed.append(s["base_name"])
        out[sid] = base
    return out, {"パスワードが無い": warn_no_pw, "新システムで変えた": warn_changed}


def build_members(members):
    """v2_members。見出し付きの表。新システムが名簿の正なので丸ごと置き換える。"""
    rows = [[str(m.get(c, "") or "") for c in MEMBERS_COLS] for m in members]
    return [list(MEMBERS_COLS)] + rows


def build_entries(entries):
    """v2_entry_<大会>。大会ごとに「キー→中身」。SQLite にある大会のタブだけ作る。"""
    out = {}
    for e in entries:
        out.setdefault(e["tournament_id"], {})[e["entry_key"]] = json.loads(e["data_json"])
    return out


def build_config(config, live_config):
    """v2_config。admin_password がハッシュ化されていたら Sheets の平文を残す
    （Streamlit は平文で照合するため、ハッシュを書くと管理画面に入れなくなる）。"""
    out = dict(live_config)
    warn = []
    for k, v in config.items():
        if k == "admin_password" and isinstance(v, str) and v.startswith("pbkdf2_sha256$"):
            warn.append("管理者パスワード（新システムで変更済み。Sheets 側の古いものに戻る）")
            continue
        out[k] = v
    return out, warn


def kv_values(d):
    """app.py の save_json と同じ 2列（キー＋JSON）に直す"""
    return [[str(k), json.dumps(v, ensure_ascii=False)] for k, v in d.items()]


def diff_kv(live, new):
    added = [k for k in new if k not in live]
    changed = [k for k in new if k in live and live[k] != new[k]]
    only_live = [k for k in live if k not in new]
    return added, changed, only_live


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--write", action="store_true",
                    help="実際に Sheets を書き換える（付けなければ読むだけ）")
    ap.add_argument("--db", default=DB_PATH)
    args = ap.parse_args()

    if not os.path.isfile(args.db):
        raise SystemExit(f"DB がありません: {args.db}")
    db = load_db(args.db)
    print(f"DB: {args.db}")
    print(f"  学校{len(db['schools'])} 名簿{len(db['members'])} "
          f"エントリー{len(db['entries'])} 設定{len(db['config'])}")

    sh = open_spreadsheet()
    print(f"Sheets: '{SHEET_NAME}' に接続")
    live = {ws.title: ws.get_all_values() for ws in sh.worksheets()}
    live_auth = parse_kv(live.get(f"{V2_PREFIX}auth", []), {})
    live_config = parse_kv(live.get(f"{V2_PREFIX}config", []), {})

    auth, auth_warn = build_auth(db["schools"], live_auth)
    members = build_members(db["members"])
    entries = build_entries(db["entries"])
    config, conf_warn = build_config(db["config"], live_config)

    # --- 何が変わるか ---
    print("\n== 書き戻す内容 ==")
    a, c, o = diff_kv(live_auth, auth)
    print(f"  {V2_PREFIX}auth      : 追加{len(a)} 変更{len(c)} 触らない{len(o)}")
    live_mem = live.get(f"{V2_PREFIX}members", [])
    print(f"  {V2_PREFIX}members   : {max(0, len(live_mem) - 1)}名 → {len(members) - 1}名")
    for tid, d in sorted(entries.items()):
        lv = parse_kv(live.get(f"{V2_PREFIX}entry_{tid}", []), {})
        a, c, o = diff_kv(lv, d)
        print(f"  {V2_PREFIX}entry_{tid:<9}: {len(lv)}件 → {len(d)}件（追加{len(a)} 変更{len(c)} 消える{len(o)}）")
    a, c, o = diff_kv(live_config, config)
    print(f"  {V2_PREFIX}config    : 追加{len(a)} 変更{len(c)} 触らない{len(o)}")

    skipped = sorted(t for t in live
                     if t.startswith(f"{V2_PREFIX}entry_")
                     and t[len(f"{V2_PREFIX}entry_"):] not in entries)
    if skipped:
        print(f"  触らないタブ: {' '.join(skipped)}")

    # --- 戻せないものの報告（ここが一番大事） ---
    print("\n== 戻せないもの・注意 ==")
    said = False
    if auth_warn["新システムで変えた"]:
        said = True
        print("  ★ 新システムでパスワードを再設定した学校は、Sheets 側の"
              "**古いパスワード**に戻ります。該当校へ連絡してください:")
        for n in auth_warn["新システムで変えた"]:
            print(f"      - {n}")
    if auth_warn["パスワードが無い"]:
        said = True
        print("  ★ Sheets 側にパスワードが無い学校（ログインできません。手で設定を）:")
        for n in auth_warn["パスワードが無い"]:
            print(f"      - {n}")
    for w in conf_warn:
        said = True
        print(f"  ★ {w}")
    print("  ★ 押印済み申込書の提出は Sheets に戻せません"
          "（新システムの ~/entry_data/uploads/ に残ります）")
    if not said:
        print("  （パスワード関係の食い違いはありません）")

    if not args.write:
        print("\n読むだけで終わりました。実際に書き戻すには --write を付けてください。")
        return

    # --- 書く前に、いまの Sheets を丸ごと控える ---
    stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    snapdir = os.path.join(DATA_DIR, f"snapshot_before_rollback_{stamp}")
    print(f"\n書き戻す前の控え → {snapdir}")
    snapshot(sh, snapdir)

    def put(title, values):
        ws = sh.worksheet(title)
        ws.clear()
        if values:
            ws.update(values)
        print(f"  書いた: {title}  {len(values)}行")

    print("書き戻し中…")
    put(f"{V2_PREFIX}auth", kv_values(auth))
    put(f"{V2_PREFIX}members", members)
    for tid, d in sorted(entries.items()):
        put(f"{V2_PREFIX}entry_{tid}", kv_values(d))
    put(f"{V2_PREFIX}config", kv_values(config))
    print("\n完了。Streamlit を開いて、学校ログインとエントリーの表示を確かめてください。")
    print(f"元に戻したいときは {snapdir} の中身を貼り直します。")


if __name__ == "__main__":
    main()
