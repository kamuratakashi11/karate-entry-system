"""書き戻しの検算（手元だけで完結・Sheets には触らない）。

    python tools/test_roundtrip.py

やること: SQLite →（export_to_sheets が作る Sheets の値）→ SQLite と一周させて、
最初と最後の DB が一致するかを見る。一致すれば「書き戻しで何も落ちない」と言える。

Sheets の現在値の代わりには、取り込みのときの snapshot（生の値）を使う。
パスワードだけは塩が毎回変わるのでハッシュ同士は比べず、**平文が新しいハッシュでも
通ること**を確かめる（これが通れば先生方は同じパスワードで入れる）。
"""

import json
import os
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from export_to_sheets import (build_auth, build_config, build_entries, build_members,
                              kv_values, load_db, password_matches)
from migrate_from_sheets import DATA_DIR, V2_PREFIX, import_sqlite, parse_kv

DB = os.path.join(DATA_DIR, "entry.sqlite3")


def newest_snapshot():
    dirs = sorted(d for d in os.listdir(DATA_DIR)
                  if d.startswith("snapshot_") and os.path.isdir(os.path.join(DATA_DIR, d)))
    if not dirs:
        raise SystemExit("snapshot_* が見つかりません（先に migrate_from_sheets.py を流す）")
    return os.path.join(DATA_DIR, dirs[-1])


def load_tab(snapdir, title):
    path = os.path.join(snapdir, f"{title}.json")
    if not os.path.isfile(path):
        return []
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def fake_hash(plain):
    import hashlib, os as _os
    salt = _os.urandom(16)
    d = hashlib.pbkdf2_hmac("sha256", plain.encode("utf-8"), salt, 100_000)
    return f"pbkdf2_sha256$100000${salt.hex()}${d.hex()}"


def scenario_checks(check, live_auth, db):
    """警告の経路（＝この道具の一番大事なところ）を、作り物のデータで通す。"""
    sid = db["schools"][0]["school_id"]
    plain = (live_auth.get(sid) or {}).get("password")

    # 1. 新システムで再設定した学校 → 名指しで警告し、Sheets の古い平文を残す
    changed = dict(db["schools"][0], password_hash=fake_hash(plain + "_changed"))
    auth, warn = build_auth([changed], live_auth)
    check("再設定した学校を名指しできる", warn["新システムで変えた"] == [changed["base_name"]],
          f"{warn['新システムで変えた']}")
    check("そのとき Sheets の平文は消さない", auth[sid]["password"] == plain)

    # 2. Sheets に無い学校（新システムだけにある）→ 別の警告
    newbie = dict(db["schools"][0], school_id="sch_new_x", base_name="新設高校",
                  password_hash=fake_hash("whatever"))
    auth2, warn2 = build_auth([newbie], live_auth)
    check("Sheets に無い学校を知らせる", warn2["パスワードが無い"] == ["新設高校"])
    check("その学校は空のパスワードで書く", auth2["sch_new_x"]["password"] == "")
    check("既存の学校は消さない", sid in auth2 and len(auth2) == len(live_auth) + 1)

    # 3. 変えていない学校は警告に出ない（狼少年にしない）
    _, warn3 = build_auth([db["schools"][0]], live_auth)
    check("変えていない学校は警告しない", not warn3["新システムで変えた"])

    # 4. 管理者パスワード: ハッシュ化されていたら Sheets の平文を残す
    live_conf = {"admin_password": "oldadmin", "year": "8"}
    conf, cw = build_config({"admin_password": fake_hash("newadmin"), "year": "9"}, live_conf)
    check("ハッシュ化された管理者PWは書かない", conf["admin_password"] == "oldadmin" and bool(cw))
    check("他の設定は新システムの値で上書き", conf["year"] == "9")
    conf2, cw2 = build_config({"admin_password": "plain123"}, live_conf)
    check("平文のままなら素通し", conf2["admin_password"] == "plain123" and not cw2)

    # 5. 名簿・エントリーは新システムが正（増減がそのまま出る）
    mem = build_members([dict(zip(
        ("school_id", "name", "sex", "grade", "dob", "jkf_no", "display_order", "active"),
        ("sch_1", "山田 太郎", "男子", "2", "", "", "", "True")))])
    check("名簿は見出し＋人数ぶん", mem[0][1] == "name" and len(mem) == 2 and mem[1][1] == "山田 太郎")
    ent = build_entries([{"tournament_id": "shinjin", "entry_key": "sch_1_山田 太郎",
                          "data_json": '{"kata_chk": true}'}])
    check("エントリーは大会ごとにまとまる",
          ent == {"shinjin": {"sch_1_山田 太郎": {"kata_chk": True}}})


def main():
    fails = []

    def check(label, ok, detail=""):
        print(f"  {'OK  ' if ok else '★NG '} {label}{('  ' + detail) if detail else ''}")
        if not ok:
            fails.append(label)

    snapdir = newest_snapshot()
    print(f"元の DB   : {DB}")
    print(f"Sheetsの控え: {snapdir}\n")

    db = load_db(DB)
    live_auth = parse_kv(load_tab(snapdir, f"{V2_PREFIX}auth"), {})
    live_config = parse_kv(load_tab(snapdir, f"{V2_PREFIX}config"), {})

    auth, auth_warn = build_auth(db["schools"], live_auth)
    members = build_members(db["members"])
    entries = build_entries(db["entries"])
    config, _ = build_config(db["config"], live_config)

    # Sheets に書いたのと同じ形に組み直して、取り込み側にそのまま食わせる
    tabs = {f"{V2_PREFIX}auth": kv_values(auth),
            f"{V2_PREFIX}members": members,
            f"{V2_PREFIX}config": kv_values(config)}
    for tid, d in entries.items():
        tabs[f"{V2_PREFIX}entry_{tid}"] = kv_values(d)

    tmp = os.path.join(tempfile.gettempdir(), "roundtrip.sqlite3")
    print("一周させる（SQLite → Sheetsの値 → SQLite）")
    import_sqlite(tabs, tmp)
    back = load_db(tmp)
    print()

    print("== 学校 ==")
    check("件数", len(back["schools"]) == len(db["schools"]),
          f"{len(db['schools'])} → {len(back['schools'])}")
    a = {s["school_id"]: s for s in db["schools"]}
    b = {s["school_id"]: s for s in back["schools"]}
    diff = []
    for sid, s in a.items():
        t = b.get(sid)
        if not t:
            diff.append(f"{s['base_name']}（消えた）")
            continue
        for col in ("base_name", "short_name", "school_no", "principal"):
            if s[col] != t[col]:
                diff.append(f"{s['base_name']}.{col}: {s[col]!r} → {t[col]!r}")
        if json.loads(s["advisors_json"]) != json.loads(t["advisors_json"]):
            diff.append(f"{s['base_name']}.advisors")
        if json.loads(s["raw_json"] or "{}") != json.loads(t["raw_json"] or "{}"):
            diff.append(f"{s['base_name']}.raw_json")
    check("中身（パスワード以外）", not diff, "" if not diff else f"{diff[:4]}")

    # パスワード: 平文が「元のハッシュ」でも「一周後のハッシュ」でも通ること
    bad = []
    for sid, s in a.items():
        plain = (live_auth.get(sid) or {}).get("password")
        if plain in (None, ""):
            continue
        if s["base_name"] in auth_warn["新システムで変えた"]:
            continue                       # 戻せないと分かっているぶんは対象外
        if not password_matches(plain, b[sid]["password_hash"]):
            bad.append(s["base_name"])
    check("同じパスワードで入れる", not bad, "" if not bad else f"{bad[:4]}")

    print("== 名簿 ==")
    key = lambda m: (m["school_id"], m["name"])
    ma = sorted((tuple(m[c] for c in ("school_id", "name", "sex", "grade", "dob",
                                      "jkf_no", "display_order", "active"))
                 for m in db["members"]))
    mb = sorted((tuple(m[c] for c in ("school_id", "name", "sex", "grade", "dob",
                                      "jkf_no", "display_order", "active"))
                 for m in back["members"]))
    check("件数", len(ma) == len(mb), f"{len(ma)} → {len(mb)}")
    lost = [x for x in ma if x not in mb]
    check("中身", not lost, "" if not lost else f"{lost[:2]}")

    print("== エントリー ==")
    ea = {(e["tournament_id"], e["entry_key"]): json.loads(e["data_json"]) for e in db["entries"]}
    eb = {(e["tournament_id"], e["entry_key"]): json.loads(e["data_json"]) for e in back["entries"]}
    check("件数", len(ea) == len(eb), f"{len(ea)} → {len(eb)}")
    ediff = [k for k, v in ea.items() if eb.get(k) != v]
    check("中身", not ediff, "" if not ediff else f"{len(ediff)}件ちがう 例:{ediff[:2]}")

    print("== 設定 ==")
    cdiff = [k for k, v in db["config"].items()
             if k != "admin_password" and back["config"].get(k) != v]
    check("中身（管理者パスワードを除く）", not cdiff, "" if not cdiff else f"{cdiff}")

    os.remove(tmp)

    # 実データでは差分ゼロ（＝まだ新システムに入力が無い）なので、
    # 「戻せないものを見落とさないか」は作り物のデータで確かめる
    print("\n== 戻せないものの見つけ方（作り物のデータ） ==")
    scenario_checks(check, live_auth, db)

    print()
    if fails:
        print(f"★ 不一致 {len(fails)}件: {fails}")
        sys.exit(1)
    print("すべて一致。書き戻しても中身は落ちません。")


if __name__ == "__main__":
    main()
