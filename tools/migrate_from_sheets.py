"""Sheets → SQLite の移行（さくら移設・マイルストーン①）。

使い方:
    python tools/migrate_from_sheets.py --check     接続確認だけ（タブ一覧と行数を表示）
    python tools/migrate_from_sheets.py             スナップショット → SQLite 取り込み → 検算

- 鍵:     C:/Users/kamur/karate_keys/service_account.json（git の外）
- 出力先: C:/Users/kamur/karate_entry_data/（git の外。実名・生年月日・平文PWが入るため）
    snapshot_<日時>/<タブ名>.json   Sheets の生の値。改変前のバックアップ兼、検算の正
    entry.sqlite3                   新システムの DB（実行のたびに作り直す）
- Sheets は**読み取りのみ**。書き込みは一切しない（切り戻し用の逆方向は別スクリプト）。
- 何度でも実行してよい（受付切替の直前にもう一度流して最新化する想定）。

パスワードは PBKDF2-SHA256 でハッシュ化して入れる。PHP 側の照合はこう書く:
    [$algo, $iters, $salt_hex, $hash_hex] = explode('$', $stored);
    $calc = bin2hex(hash_pbkdf2('sha256', $input_pw, hex2bin($salt_hex), (int)$iters, 0, true));
    $ok = hash_equals($hash_hex, $calc);
"""

import argparse
import datetime
import hashlib
import json
import os
import re
import sqlite3
import sys

import gspread

KEY_PATH = r"C:/Users/kamur/karate_keys/service_account.json"
DATA_DIR = r"C:/Users/kamur/karate_entry_data"
SHEET_NAME = "tournament_db"   # app.py の SHEET_NAME と同じ
V2_PREFIX = "v2_"              # app.py の V2_PREFIX と同じ
MEMBERS_COLS = ["school_id", "name", "sex", "grade", "dob", "jkf_no", "display_order", "active"]
PBKDF2_ITERS = 100_000


def open_spreadsheet():
    gc = gspread.service_account(filename=KEY_PATH)
    return gc.open(SHEET_NAME)


def parse_kv(values, default):
    """app.py の load_json と同じ解釈: 1セルのJSON か、key+JSON の行形式。"""
    if not values:
        return default
    if len(values) == 1 and len(values[0]) >= 1:
        val = str(values[0][0])
        if val.startswith("{") or val.startswith("["):
            parsed = json.loads(val)
            return parsed if parsed is not None else default
    result = {}
    for row in values:
        if len(row) >= 2 and row[0]:
            try:
                result[row[0]] = json.loads(row[1])
            except (json.JSONDecodeError, TypeError):
                result[row[0]] = row[1]
    return result if result else default


def hash_password(plain):
    salt = os.urandom(16)
    digest = hashlib.pbkdf2_hmac("sha256", str(plain).encode("utf-8"), salt, PBKDF2_ITERS)
    return f"pbkdf2_sha256${PBKDF2_ITERS}${salt.hex()}${digest.hex()}"


def snapshot(sh, outdir):
    """全タブの生の値を保存する。移行対象外のタブ（v1等）も含めて全部。"""
    os.makedirs(outdir, exist_ok=True)
    tabs = {}
    for ws in sh.worksheets():
        values = ws.get_all_values()
        safe = re.sub(r'[\\/:*?"<>|]', "_", ws.title)
        with open(os.path.join(outdir, f"{safe}.json"), "w", encoding="utf-8") as f:
            json.dump(values, f, ensure_ascii=False, indent=1)
        tabs[ws.title] = values
        print(f"  取得: {ws.title}  {len(values)}行")
    return tabs


def import_sqlite(tabs, db_path):
    if os.path.exists(db_path):
        os.remove(db_path)
    con = sqlite3.connect(db_path)
    cur = con.cursor()
    cur.executescript("""
        CREATE TABLE schools (
            school_id TEXT PRIMARY KEY,
            base_name TEXT, short_name TEXT, school_no INTEGER,
            principal TEXT, advisors_json TEXT,
            password_hash TEXT,
            raw_json TEXT       -- v2_auth の元の値そのまま（平文PWは抜いてある）
        );
        CREATE TABLE members (
            school_id TEXT, name TEXT, sex TEXT, grade TEXT,
            dob TEXT, jkf_no TEXT, display_order TEXT, active TEXT
        );
        CREATE TABLE entries (
            tournament_id TEXT, entry_key TEXT, data_json TEXT,
            PRIMARY KEY (tournament_id, entry_key)
        );
        CREATE TABLE config (key TEXT PRIMARY KEY, value_json TEXT);
        CREATE INDEX idx_members_school ON members(school_id);
    """)

    counts = {}

    # --- 学校（v2_auth） ---
    auth = parse_kv(tabs.get(f"{V2_PREFIX}auth", []), {})
    for sid, d in auth.items():
        if not isinstance(d, dict):
            continue
        d = dict(d)
        plain = d.pop("password", "")
        cur.execute(
            "INSERT INTO schools VALUES (?,?,?,?,?,?,?,?)",
            (sid, d.get("base_name", ""), d.get("short_name", d.get("base_name", "")),
             int(d.get("school_no", 999) or 999), d.get("principal", ""),
             json.dumps(d.get("advisors", []), ensure_ascii=False),
             hash_password(plain), json.dumps(d, ensure_ascii=False)),
        )
    counts["schools"] = len(auth)

    # --- 名簿（v2_members・表形式） ---
    mv = tabs.get(f"{V2_PREFIX}members", [])
    if mv:
        header = mv[0]
        missing = [c for c in MEMBERS_COLS if c not in header]
        if missing:
            raise SystemExit(f"v2_members に想定列が無い: {missing}（列を推測して黙って取り込まない）")
        idx = {c: header.index(c) for c in MEMBERS_COLS}
        rows = [
            tuple(r[idx[c]] if idx[c] < len(r) else "" for c in MEMBERS_COLS)
            for r in mv[1:] if any(x.strip() for x in r)
        ]
        cur.executemany("INSERT INTO members VALUES (?,?,?,?,?,?,?,?)", rows)
        counts["members"] = len(rows)
    else:
        counts["members"] = 0

    # --- エントリー（v2_entry_*） ---
    n = 0
    for title, values in tabs.items():
        if not title.startswith(f"{V2_PREFIX}entry_"):
            continue
        tid = title[len(f"{V2_PREFIX}entry_"):]
        data = parse_kv(values, {})
        for key, v in data.items():
            cur.execute("INSERT OR REPLACE INTO entries VALUES (?,?,?)",
                        (tid, key, json.dumps(v, ensure_ascii=False)))
            n += 1
    counts["entries"] = n

    # --- 設定（v2_config） ---
    conf = parse_kv(tabs.get(f"{V2_PREFIX}config", []), {})
    for k, v in conf.items():
        cur.execute("INSERT INTO config VALUES (?,?)", (k, json.dumps(v, ensure_ascii=False)))
    counts["config"] = len(conf)

    con.commit()

    # --- 検算: DB の件数がいま取り込んだ元データと一致するか ---
    ok = True
    for table, expected in counts.items():
        got = cur.execute(f"SELECT COUNT(*) FROM {table}").fetchone()[0]
        mark = "OK" if got == expected else "★不一致★"
        if got != expected:
            ok = False
        print(f"  {table}: 元 {expected} 件 → DB {got} 件  {mark}")
    con.close()
    return ok


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--check", action="store_true", help="接続確認だけ（読み取りのみ）")
    args = ap.parse_args()

    sh = open_spreadsheet()
    if args.check:
        print(f"接続OK: '{SHEET_NAME}'")
        for ws in sh.worksheets():
            print(f"  {ws.title}  {ws.row_count}x{ws.col_count}")
        return

    stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    snapdir = os.path.join(DATA_DIR, f"snapshot_{stamp}")
    print(f"スナップショット → {snapdir}")
    tabs = snapshot(sh, snapdir)

    db_path = os.path.join(DATA_DIR, "entry.sqlite3")
    print(f"取り込み → {db_path}")
    ok = import_sqlite(tabs, db_path)
    print("完了" if ok else "検算に不一致あり。取り込み結果を確認すること")
    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()
