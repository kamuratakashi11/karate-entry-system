"""参加校一覧（新人大会）のテンプレートを作る。

実物 `[訂正]2025新人大会参加校一覧.xlsx` から**データだけ消して**書式・罫線・
合計行の数式を残したものを `server/entry/lib/template_schools_shinjin.xlsx` として保存する。
申込システムはこの空欄に値を入れて配る（罫線を引き直さないので実物と同じ見た目になる）。

    python tools/make_schoollist_template.py

※ 実物は git 管理外（OneDrive）。作ったテンプレートには学校名も人数も残らない。
"""

import shutil
from pathlib import Path

import openpyxl

SRC = Path(r"C:/Users/kamur/OneDrive - さいたま市教育委員会/20240412競技部会長資料"
           r"/競技部会長資料/2025年度/2025新人大会/[訂正]2025新人大会参加校一覧.xlsx")
DST = Path(__file__).resolve().parent.parent / "server" / "entry" / "lib" / "template_schools_shinjin.xlsx"

DATA_FIRST, DATA_LAST = 7, 74      # 学校の行（実物は7〜37に31校＋予備）
SUM_FIRST, SUM_LAST = 75, 78       # 学校数 / 参加人数 / 別枠シード / 合計人数
LABEL_COL = 2                      # B列のラベル（「学校数」など）だけは残す
TITLE_CELL = "A2"                  # 「令和7年度　埼玉県空手道新人大会」→ 申込システムが入れ直す

# ★ 合計行の数式は**消す**。申込システムが数えた結果を数値で書き込む。
#   実物の数式は COUNTA/SUM に加えて「=O75+1」のような手直しが混ざっており、
#   そのまま使うと数え落としが再発する（-61K と +66K に +1 されていた）。


def main() -> None:
    if not SRC.exists():
        raise SystemExit(f"実物が見つかりません: {SRC}")
    tmp = DST.with_suffix(".tmp.xlsx")
    DST.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy(SRC, tmp)

    wb = openpyxl.load_workbook(tmp)
    ws = wb["参加校一覧"]

    cleared = 0
    for r in range(DATA_FIRST, DATA_LAST + 1):
        for c in range(1, ws.max_column + 1):
            cell = ws.cell(r, c)
            if cell.value is not None:
                cell.value = None      # 書式（罫線・塗り・配置）はそのまま残る
                cleared += 1
    for r in range(SUM_FIRST, SUM_LAST + 1):
        for c in range(1, ws.max_column + 1):
            if c == LABEL_COL:
                continue               # 「学校数」「参加人数」…のラベルは残す
            cell = ws.cell(r, c)
            if cell.value is not None:
                cell.value = None
                cleared += 1
    ws[TITLE_CELL] = None              # 大会名は毎回入れ直す

    wb.save(DST)
    tmp.unlink()

    chk = openpyxl.load_workbook(DST)["参加校一覧"]
    formulas = sum(1 for row in chk.iter_rows()
                   for c in row if isinstance(c.value, str) and c.value.startswith("="))
    print(f"できた: {DST}")
    print(f"  消したセル: {cleared} / 残った結合セル: {len(chk.merged_cells.ranges)}")
    print(f"  残った数式: {formulas}個（0が正しい）")
    print(f"  見出し行3-6: {[chk.cell(3,1).value, chk.cell(5,4).value, chk.cell(6,10).value]}")
    print(f"  合計行のラベル: {[chk.cell(r,2).value for r in range(SUM_FIRST, SUM_LAST+1)]}")
    print(f"  学校の行 {DATA_FIRST}-{DATA_LAST} が空: "
          f"{all(chk.cell(r,2).value is None for r in range(DATA_FIRST, DATA_LAST+1))}")


if __name__ == "__main__":
    main()
