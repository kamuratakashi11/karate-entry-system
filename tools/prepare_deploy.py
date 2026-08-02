"""さくら配備用の一式を組み立てる（②⑤・ワンコマンド）。

    python tools/prepare_deploy.py            初回の設置（DB込み・Sheets取り込みから）
    python tools/prepare_deploy.py --update   コードだけ上げ直す（DBに触らない）

--update は、設置後にプログラムを直したときの上げ直し用。**Sheets を読まず、
DBも作らない**ので、先生方が入力中でも安全に流せる（提出済みの申込書も消えない）。
出力は C:/Users/kamur/karate_entry_data/deploy_update/。

やること（--update なしのとき）:
  1. Sheets から最新を取り込む（tools/migrate_from_sheets.py をそのまま実行）
  2. DB に本番設定を当てる: 新人戦（shinjin）だけ受付中・締切 2026-09-25
  3. アップロード用フォルダを組み立てる:
       C:/Users/kamur/karate_entry_data/deploy_entry/
         www_entry/    → さくらの ~/www/entry/ へ丸ごと（lib/ ごと）
         手順.txt      → 孝さん向けの手順（CP932・メモ帳用）

**さくらのファイルマネージャーは www の中しか開けない**（2026-07-31 実機で判明）。
そのため DB もいったん www_entry に同梱し、サーバー上で `setup.php` を開いて
`~/entry_data/` を作って移す。`.htaccess` は移すまでの数分間の保険。

**受付切替の直前（8/19ごろ）にもう一度これを流して、entry.sqlite3 と setup.php を
上げ直す**（それまで現行 Streamlit 側で変わったデータを反映するため）。
"""

import json
import shutil
import sqlite3
import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
REPO = HERE.parent
DATA = Path(r"C:/Users/kamur/karate_entry_data")
DEPLOY = DATA / "deploy_entry"
UPDATE = DATA / "deploy_update"

ACTIVE_TID = "shinjin"
DEADLINE = "2026-09-25"

APP_FILES = ["index.html", "style.css", "app.js", "api.php",
             "auth.php", "db.php", "config.php", "uploads.php", "check.php", "setup.php",
             "admin.html", "admin_api.php"]
# lib/ は中身をまるごと運ぶ（大会ごとのテンプレート・座標が増えるため列挙しない）

# 上げ直し（--update）で運ぶもの。setup.php は初回だけの道具なので入れない
# （DBを動かすものを毎回上げると、要らない事故の芽になる）。lib/ は申込書の
# テンプレート・座標なので、そこを直したときだけ手で上げる
UPDATE_FILES = [f for f in APP_FILES if f != "setup.php"]

TEJUN_UPDATE = """申込窓口（/entry/）のプログラムを上げ直す手順
==============================================

データベース（名簿・エントリー）と、提出された申込書には触れません。
先生方が入力中でも安全です。

1. さくらのファイルマネージャーで www\\entry を開く

2. deploy_update\\www_entry の中の{n_files}ファイルを、そこへアップロードする
   （同じ名前のものを上書きする。「lib」フォルダは触らなくて構いません）

3. ブラウザで開いて確かめる:
   https://saitamahs-karate.sakura.ne.jp/entry/check.php
   → 緑の「配備OK」が出ればOK
   ※「アップロードの上限」の行を見てください。3MB未満で赤が出たら、
     押印済み申込書（写真）が提出できないので教えてください

4. 確かめたら check.php を削除する
   （サーバーの中身を映すページなので、置きっぱなしにしない）

5. https://saitamahs-karate.sakura.ne.jp/entry/ を開いて動作を確認
   ・「4. 申込書」タブに「4.2 押印して提出」が出ていればOK
   ・管理画面（admin.html）に「提出書類」タブが増えています

■ 今回の変更（2026-08-01〜02）

  押印済み申込書の提出をつくりました。
  ・先生方: 「4. 申込書」タブから、押印したものをPDFで提出できます
    間違えたら出し直せます（新しいほうが提出物。古いものも残ります）
  ・受け付けるのはPDFだけです（写真は不可）。エントリーが22名を超えると
    申込書は2ページになり、写真だと2枚バラバラの提出になってしまうため。
    画面に「スマートフォンでPDFにするには」の手順を載せてあります。
    PDFにできない先生からは、加村あてにメールをもらう運用です
  ・専門部: 管理画面の「提出書類」タブで、提出状況の確認・未提出の学校の
    絞り込み・ZIPでの一括ダウンロードができます
  ・提出物は www の外（entry_data\\uploads）に置かれ、この画面からだけ
    取り出せます。URLを知られても取られません
  ・いらないものは×で1件ずつ、たまったものは年度ごとにまとめて消せます
    （消す前に必ずZIPで手元に保存してください。戻せません）
"""

TEJUN = """申込窓口（/entry/）をさくらに置く手順
=========================================

さくらのファイルマネージャーは www の中しか開けません。
DBを置く「www の外」のフォルダは、setup.php がブラウザから作ります。
（ファイルマネージャーとブラウザだけで終わります）

■ いま（設置と動作確認）

1. ファイルマネージャーで、www の中に「entry」フォルダを作る

2. entry の中へ、deploy_entry\\www_entry の中身をアップロードする
   ・ファイル{n_files}個（.htaccess と entry.sqlite3 を含む）
   ・さらに entry の中に「lib」フォルダを作り、www_entry\\lib の{n_lib}ファイルを入れる
   ※ この時点では entry.sqlite3（データ本体）も entry の中にあります。
     次の3で www の外へ移します

3. ブラウザで開く:
   https://saitamahs-karate.sakura.ne.jp/entry/setup.php
   → 緑の「できました」が出る
     （www の外に entry_data フォルダを作り、entry.sqlite3 をそこへ移します）

4. ファイルマネージャーで entry の中の setup.php を削除する
   ※ entry.sqlite3 が消えていれば正常です（3で外へ移ったため）

5. ブラウザで開く:
   https://saitamahs-karate.sakura.ne.jp/entry/check.php
   → 緑の「配備OK」が出ればOK。赤が出たら、その画面をそのまま伝えてください

6. https://saitamahs-karate.sakura.ne.jp/entry/ を開き、
   自分の学校で（今までと同じパスワードで）ログインして触ってみる
   ※ 試しに保存しても大丈夫（8/20直前にデータを入れ直すので消えます）

7. 確認できたら check.php も削除する

■ 8/20 の直前（受付開始の前日ごろ・データの最終化）

1. パソコンで  python tools/prepare_deploy.py  をもう一度実行
   （それまでに現行システムで変わった名簿などを取り込み直すため）
2. できた www_entry の中から entry.sqlite3 と setup.php の2つを、
   サーバーの entry フォルダにアップロード
3. https://saitamahs-karate.sakura.ne.jp/entry/setup.php を開く
   （古いDBが新しいものに入れ替わります）
4. setup.php を削除する
5. 先生方に新しいURLを案内する: https://saitamahs-karate.sakura.ne.jp/entry/

■ 管理画面（専門部用）

  https://saitamahs-karate.sakura.ne.jp/entry/admin.html
  ・パスワードは今までの申込システムの管理者パスワードと同じ
  ・受付状況（どの学校が入力したか）／提出書類（押印済み申込書）／
    データ出力（CSV）／大会設定／学校のパスワード再設定 ができます
  ・押印済みの申込書は、先生方が「4. 申込書」の画面から提出します。
    集まったものは「提出書類」タブでZIPにまとめて取り出せます
    （置き場所は www の外なので、URLでは取り出せません）
  ・先生方の画面からはリンクしていません（このURLを直接開いてください）
  ・最初に開いたら「大会設定」タブで管理者パスワードを変更してください
    （いまは暗号化されずに保存されているため）

■ 注意

・www\\entry に BASIC認証は掛けない（学校ごとのログインがあるため。
  掛けると先生方が入れなくなります）
・.htaccess も必ず一緒に置く（設置中のあいだ、DBが外から取られないようにする保険）
・現行の Streamlit は当面そのまま残す（切り戻し用）。二重に受け付けない
  よう、案内は片方だけにする
"""


def build_update() -> None:
    """設置後の上げ直し用。Sheets もDBも触らず、プログラムだけを集める。"""
    if UPDATE.exists():
        shutil.rmtree(UPDATE)
    www = UPDATE / "www_entry"
    www.mkdir(parents=True)

    src = REPO / "server" / "entry"
    for f in UPDATE_FILES:
        shutil.copy(src / f, www / f)

    files = sorted(p.name for p in www.iterdir() if p.is_file())
    (UPDATE / "手順.txt").write_text(
        TEJUN_UPDATE.format(n_files=len(files)), encoding="cp932")

    print(f"できた: {UPDATE}")
    print(f"  www_entry: {len(files)}ファイル（DBは入れていません）")
    print(f"    {' '.join(files)}")
    print("  → www/entry へ上書きアップロード（手順.txt）")


def main() -> None:
    # 1. 最新の取り込み（読み取り専用・何度でも安全）
    print("== 1/3 Sheets から最新を取り込む ==")
    r = subprocess.run([sys.executable, str(HERE / "migrate_from_sheets.py")],
                       cwd=str(REPO))
    if r.returncode != 0:
        raise SystemExit("取り込みに失敗（上のメッセージを確認）")

    # 2. 配備フォルダを組み立て
    print("== 2/3 配備フォルダを組み立てる ==")
    if DEPLOY.exists():
        shutil.rmtree(DEPLOY)
    www = DEPLOY / "www_entry"
    (www / "lib").mkdir(parents=True)

    src = REPO / "server" / "entry"
    for f in APP_FILES:
        shutil.copy(src / f, www / f)
    lib_files = sorted(p for p in (src / "lib").iterdir() if p.is_file())
    for p in lib_files:
        shutil.copy(p, www / "lib" / p.name)
    # 配布時だけ .htaccess の名前で置く（リポジトリ側は見える名前のまま）
    shutil.copy(src / "htaccess.txt", www / ".htaccess")

    # DBもいったん www_entry に同梱する（ファイルマネージャーが www の外へ
    # 上げられないため。サーバー上で setup.php が ~/entry_data/ へ移す）
    db_path = www / "entry.sqlite3"
    shutil.copy(DATA / "entry.sqlite3", db_path)

    # 3. 本番設定（新人戦だけ受付中・締切）
    print("== 3/3 本番設定を当てる ==")
    con = sqlite3.connect(db_path)
    row = con.execute("SELECT value_json FROM config WHERE key='tournaments'").fetchone()
    ts = json.loads(row[0])
    if ACTIVE_TID not in ts:
        raise SystemExit(f"設定に {ACTIVE_TID} が無い: {list(ts)}")
    for tid in ts:
        ts[tid]["active"] = (tid == ACTIVE_TID)
    ts[ACTIVE_TID]["deadline"] = DEADLINE
    con.execute("UPDATE config SET value_json=? WHERE key='tournaments'",
                (json.dumps(ts, ensure_ascii=False),))
    con.commit()

    n = {t: con.execute(f"SELECT COUNT(*) FROM {t}").fetchone()[0]
         for t in ("schools", "members", "entries")}
    con.close()

    files = sorted(p.name for p in www.iterdir() if p.is_file())
    # 手順書のファイル数は実際に組み立てた数から書く（増減のたびに食い違わないように）
    (DEPLOY / "手順.txt").write_text(
        TEJUN.format(n_files=len(files), n_lib=len(lib_files)), encoding="cp932")

    print(f"できた: {DEPLOY}")
    print(f"  受付中: {ACTIVE_TID}（締切 {DEADLINE}）")
    print(f"  学校 {n['schools']} / 名簿 {n['members']} / エントリー {n['entries']}")
    print(f"  www_entry: {len(files)}ファイル + lib({len(lib_files)})")
    print(f"    {' '.join(files)}")
    print("  → www/entry へ丸ごと上げ、setup.php をブラウザで開く（手順.txt）")


if __name__ == "__main__":
    if "--update" in sys.argv[1:]:
        build_update()
    else:
        main()
