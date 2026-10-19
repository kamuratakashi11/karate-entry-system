"""級位審査の様式を、申込システムが使う形に写す（審査の様式を作り直したときに流す）。

    python tools/make_shinsa_assets.py <審査部会の資料フォルダ> <申込書の xlsx>

  例: python tools/make_shinsa_assets.py ^
        "C:/Users/kamur/OneDrive/デスクトップ/審査部会HP掲載資料" ^
        "C:/Users/kamur/Documents/審査部会/級位審査申込書-1.xlsx"

作るもの（どちらも server/entry/ の下。git に入れる。実名は入っていない）:
  lib/template_shinsa.xlsx   … 紙の申込書の型。原本から 3 か所だけ変える:
                               ① 外部リンク（他人の OneDrive の .xls を指す「大会名」）を外す。
                                  使っていないのに、開くたびに「リンクの更新」を聞かれるため
                               ② 開いたときに式（男女・上級初級の人数）を計算し直させる
                               ③ 「受審申込書 」を A4 1 枚に収める（原本は日時だけ 2 枚目に出た）
  shinsa/sheet_jokyu.png ほか … 審査用紙と領収書の地紙（PDF を画像にしたもの）。この上に文字を重ねて刷る

審査の様式は孝さん（審査部会）が作っているので、原紙のアプリに合わせなくてよい
（方針 2026-10-13）。画像を敷いてブラウザから刷る形にしてある。

PDF を画像にするのに pymupdf が要る（申込システム本体は要らない）。
"""

import re
import sys
import zipfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
ENTRY = HERE.parent / "server" / "entry"

# 審査用紙: (PDF ファイル名, ページ番号, 書き出す名前)
SHEETS = [
    ("01.審査用紙.pdf", 0, "sheet_jokyu.png"),                  # 上級（1〜4級を受ける）
    ("01.審査用紙.pdf", 1, "sheet_shokyu.png"),                 # 初級（5〜8級を受ける）
    ("02.レベルチェックテスト審査用紙.pdf", 0, "sheet_lc.png"),  # レベルチェックテスト
    # 領収書（A4 に 2 枚。金額は刷ってある）
    ("03.級位審査会領収書（3000円）.pdf", 0, "receipt_pass.png"),
    ("04.級位審査会領収書（返金あり用）.pdf", 0, "receipt_fail.png"),
]
DPI = 200


def make_template(src: Path, dst: Path) -> None:
    drop = {"xl/externalLinks/externalLink1.xml", "xl/externalLinks/_rels/externalLink1.xml.rels"}
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        names = zin.namelist()
        for need in drop:
            if need not in names:
                raise SystemExit(f"想定した外部リンクが無い（様式が変わった？）: {need}")
        for item in zin.infolist():
            if item.filename in drop:
                continue
            data = zin.read(item.filename)
            if item.filename == "xl/workbook.xml":
                s = data.decode("utf-8")
                s, n1 = re.subn(r"<externalReferences>.*?</externalReferences>", "", s)
                s, n2 = re.subn(r'<definedName name="大会名">[^<]*</definedName>', "", s)
                s, n3 = re.subn(r"<calcPr([^>]*)/>",
                                lambda m: "<calcPr" + m.group(1).replace(' fullCalcOnLoad="1"', "")
                                + ' fullCalcOnLoad="1"/>', s)
                if (n1, n2, n3) != (1, 1, 1):
                    raise SystemExit(f"workbook.xml が想定と違う: {n1, n2, n3}")
                data = s.encode("utf-8")
            elif item.filename == "xl/_rels/workbook.xml.rels":
                s = data.decode("utf-8")
                s, n = re.subn(r'<Relationship Id="rId3" Type="[^"]*/externalLink" Target="[^"]*"/>', "", s)
                if n != 1:
                    raise SystemExit("workbook.xml.rels に外部リンクの rel が無い")
                data = s.encode("utf-8")
            elif item.filename == "xl/worksheets/sheet2.xml":
                # ③ 「受審申込書 」を 1 枚に収める。原本は倍率 100% のままで、45 行目
                #    （=NOW() の日時）だけが 2 枚目に刷られていた（2026-10-14 に Excel で確認）
                s = data.decode("utf-8")
                s, n1 = re.subn(r'<sheetPr codeName="Sheet2"><tabColor rgb="FFFF0000"/></sheetPr>',
                                '<sheetPr codeName="Sheet2"><tabColor rgb="FFFF0000"/>'
                                '<pageSetUpPr fitToPage="1"/></sheetPr>', s)
                s, n2 = re.subn(r'<pageSetup paperSize="9" orientation="portrait"',
                                '<pageSetup paperSize="9" fitToWidth="1" fitToHeight="1" orientation="portrait"', s)
                if (n1, n2) != (1, 1):
                    raise SystemExit(f"受審申込書の印刷設定が想定と違う: {n1, n2}")
                data = s.encode("utf-8")
            elif item.filename == "[Content_Types].xml":
                s = data.decode("utf-8")
                s, n = re.subn(r'<Override PartName="/xl/externalLinks/externalLink1.xml"[^>]*/>', "", s)
                if n != 1:
                    raise SystemExit("[Content_Types].xml に外部リンクの登録が無い")
                data = s.encode("utf-8")
            zout.writestr(item, data)
    print("書いた:", dst)


def make_sheets(pdf_dir: Path, out_dir: Path) -> None:
    import pymupdf  # 地紙を作るときだけ要る

    out_dir.mkdir(exist_ok=True)
    for pdf, page, name in SHEETS:
        doc = pymupdf.open(pdf_dir / pdf)
        pix = doc[page].get_pixmap(dpi=DPI, colorspace=pymupdf.csGRAY)
        pix.save(out_dir / name)
        print("書いた:", out_dir / name, f"{pix.width}x{pix.height}")


def main() -> None:
    if len(sys.argv) != 3:
        print(__doc__)
        raise SystemExit(2)
    pdf_dir, xlsx = Path(sys.argv[1]), Path(sys.argv[2])
    make_template(xlsx, ENTRY / "lib" / "template_shinsa.xlsx")
    make_sheets(pdf_dir, ENTRY / "shinsa")


if __name__ == "__main__":
    main()
