<?php
/**
 * 表を新しく組み立てて xlsx にする小さな書き手。依存は ext-zip だけ。
 * 様式に書き込む XlsxFill と違い、**行の数が決まっていない表**（当日用の名簿・集計）に使う。
 *
 * 使い方:
 *   $b = new XlsxBook();
 *   $s = $b->sheet('受付名簿', [6, 14, 16]);          // シート名・列幅
 *   $s->put(1, 1, '表題', XlsxBook::TITLE);
 *   $s->put(3, 1, 12, XlsxBook::CELL);                 // 数は int / float で（式が数える）
 *   $s->put(3, 4, '', XlsxBook::CELL, 'SUM(B3:C3)');   // 式（先頭の = は付けない）
 *   $s->headerRows = '3:3';                            // 印刷の各ページに見出しを繰り返す
 *   $b->save($path);
 *
 * 開いたときに式を計算し直させる（fullCalcOnLoad）。顧問出欠（advisor_sheet.php）と同じ書き方。
 */
declare(strict_types=1);

final class XlsxBookSheet
{
    /** @var array<int, array<int, array{0:mixed,1:int,2:?string}>> */
    public array $cells = [];
    public string $headerRows = '';
    public bool $landscape = false;

    /** @param array<int,float> $widths 1 列目からの幅 */
    public function __construct(public string $name, public array $widths)
    {
    }

    public function put(int $row, int $col, mixed $value, int $style = 0, ?string $formula = null): void
    {
        $this->cells[$row][$col] = [$value, $style, $formula];
    }
}

final class XlsxBook
{
    // 書式（styles.xml の cellXfs の並び）
    const PLAIN = 0;
    const HEAD  = 1;   // 見出し（太字・灰色・中央・細線）
    const CELL  = 2;   // 本体（中央・細線・縮小）
    const LEFT  = 3;   // 本体（左・細線・縮小）
    const TITLE = 4;   // 表題（太字・大きめ・左）
    const NOTE  = 5;   // 注記（左・線なし）
    const INPUT = 6;   // 当日書き込む欄（中央・細線・薄い黄色）

    /** @var list<XlsxBookSheet> */
    private array $sheets = [];

    /** @param array<int,float> $widths */
    public function sheet(string $name, array $widths): XlsxBookSheet
    {
        $s = new XlsxBookSheet($name, $widths);
        $this->sheets[] = $s;
        return $s;
    }

    public static function col(int $n): string
    {
        $s = '';
        while ($n > 0) {
            $m = ($n - 1) % 26;
            $s = chr(65 + $m) . $s;
            $n = intdiv($n - 1, 26);
        }
        return $s;
    }

    private static function esc(string $s): string
    {
        return htmlspecialchars($s, ENT_XML1 | ENT_QUOTES, 'UTF-8');
    }

    private static function sheetXml(XlsxBookSheet $sh): string
    {
        $cells = $sh->cells;
        ksort($cells);
        $rows = '';
        $maxCol = 1;
        foreach ($cells as $r => $cols) {
            ksort($cols);
            $rows .= "<row r=\"{$r}\">";
            foreach ($cols as $c => [$v, $s, $f]) {
                $maxCol = max($maxCol, $c);
                $ref = self::col($c) . $r;
                if ($f !== null) {
                    $rows .= "<c r=\"{$ref}\" s=\"{$s}\"><f>" . self::esc($f) . '</f></c>';
                } elseif (is_int($v) || is_float($v)) {
                    $rows .= "<c r=\"{$ref}\" s=\"{$s}\"><v>{$v}</v></c>";
                } elseif ($v === '' || $v === null) {
                    $rows .= "<c r=\"{$ref}\" s=\"{$s}\"/>";
                } else {
                    $rows .= "<c r=\"{$ref}\" s=\"{$s}\" t=\"inlineStr\"><is><t>" . self::esc((string)$v) . '</t></is></c>';
                }
            }
            $rows .= '</row>';
        }
        $colsXml = '';
        foreach ($sh->widths as $i => $w) {
            $c = $i + 1;
            $colsXml .= "<col min=\"{$c}\" max=\"{$c}\" width=\"{$w}\" customWidth=\"1\"/>";
        }
        $last = $cells ? max(array_keys($cells)) : 1;
        $orient = $sh->landscape ? 'landscape' : 'portrait';
        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            . '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            . 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            . '<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>'
            . '<dimension ref="A1:' . self::col($maxCol) . $last . '"/>'
            . '<sheetViews><sheetView workbookViewId="0"/></sheetViews>'
            . '<sheetFormatPr defaultRowHeight="18"/>'
            . ($colsXml !== '' ? "<cols>{$colsXml}</cols>" : '')
            . "<sheetData>{$rows}</sheetData>"
            . '<printOptions horizontalCentered="1"/>'
            . '<pageMargins left="0.4" right="0.4" top="0.5" bottom="0.5" header="0.3" footer="0.3"/>'
            // 横は 1 ページに収め、縦は人数ぶんページを増やす（0 ＝ 指定なし）
            . "<pageSetup paperSize=\"9\" orientation=\"{$orient}\" fitToWidth=\"1\" fitToHeight=\"0\"/>"
            . '<headerFooter><oddFooter>&amp;P / &amp;N</oddFooter></headerFooter>'
            . '</worksheet>';
    }

    public function save(string $path): void
    {
        $font  = '<font><sz val="11"/><name val="ＭＳ Ｐゴシック"/><family val="3"/><charset val="128"/></font>';
        $bold  = '<font><b/><sz val="11"/><name val="ＭＳ Ｐゴシック"/><family val="3"/><charset val="128"/></font>';
        $big   = '<font><b/><sz val="14"/><name val="ＭＳ Ｐゴシック"/><family val="3"/><charset val="128"/></font>';
        $small = '<font><sz val="9"/><name val="ＭＳ Ｐゴシック"/><family val="3"/><charset val="128"/></font>';
        $thin = '<border><left style="thin"/><right style="thin"/><top style="thin"/><bottom style="thin"/><diagonal/></border>';
        $center = '<alignment horizontal="center" vertical="center" shrinkToFit="1"/>';
        $left = '<alignment horizontal="left" vertical="center" shrinkToFit="1"/>';
        $styles = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            . '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            . "<fonts count=\"4\">{$font}{$bold}{$big}{$small}</fonts>"
            . '<fills count="4"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>'
            . '<fill><patternFill patternType="solid"><fgColor rgb="FFE7E9EE"/><bgColor indexed="64"/></patternFill></fill>'
            . '<fill><patternFill patternType="solid"><fgColor rgb="FFFFF8D6"/><bgColor indexed="64"/></patternFill></fill></fills>'
            . "<borders count=\"2\"><border><left/><right/><top/><bottom/><diagonal/></border>{$thin}</borders>"
            . '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
            . '<cellXfs count="7">'
            . '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
            . "<xf numFmtId=\"0\" fontId=\"1\" fillId=\"2\" borderId=\"1\" xfId=\"0\" applyFont=\"1\" applyFill=\"1\" applyBorder=\"1\" applyAlignment=\"1\">{$center}</xf>"
            . "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"1\" xfId=\"0\" applyBorder=\"1\" applyAlignment=\"1\">{$center}</xf>"
            . "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"1\" xfId=\"0\" applyBorder=\"1\" applyAlignment=\"1\">{$left}</xf>"
            . '<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>'
            . '<xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>'
            . "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"3\" borderId=\"1\" xfId=\"0\" applyFill=\"1\" applyBorder=\"1\" applyAlignment=\"1\">{$center}</xf>"
            . '</cellXfs>'
            . '<cellStyles count="1"><cellStyle name="標準" xfId="0" builtinId="0"/></cellStyles>'
            . '</styleSheet>';

        $overrides = $sheetList = $rels = $names = '';
        foreach ($this->sheets as $i => $sh) {
            $n = $i + 1;
            $overrides .= "<Override PartName=\"/xl/worksheets/sheet{$n}.xml\" "
                . 'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>';
            $sheetList .= '<sheet name="' . self::esc($sh->name) . "\" sheetId=\"{$n}\" r:id=\"rId{$n}\"/>";
            $rels .= "<Relationship Id=\"rId{$n}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet{$n}.xml\"/>";
            if ($sh->headerRows !== '') {
                [$a, $b] = explode(':', $sh->headerRows);
                $names .= '<definedName name="_xlnm.Print_Titles" localSheetId="' . $i . '">'
                    . "'" . self::esc(str_replace("'", "''", $sh->name)) . "'!\${$a}:\${$b}</definedName>";
            }
        }
        $stylesId = count($this->sheets) + 1;

        $zip = new ZipArchive();
        if ($zip->open($path, ZipArchive::CREATE | ZipArchive::OVERWRITE) !== true) {
            throw new RuntimeException("書き出し先を開けない: {$path}");
        }
        $zip->addFromString('[Content_Types].xml',
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            . '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            . '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            . '<Default Extension="xml" ContentType="application/xml"/>'
            . '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
            . $overrides
            . '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
            . '</Types>');
        $zip->addFromString('_rels/.rels',
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            . '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            . '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
            . '</Relationships>');
        $zip->addFromString('xl/workbook.xml',
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            . '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            . 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            . "<sheets>{$sheetList}</sheets>"
            . ($names !== '' ? "<definedNames>{$names}</definedNames>" : '')
            . '<calcPr calcId="0" fullCalcOnLoad="1"/>'
            . '</workbook>');
        $zip->addFromString('xl/_rels/workbook.xml.rels',
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            . '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            . $rels
            . "<Relationship Id=\"rId{$stylesId}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>"
            . '</Relationships>');
        $zip->addFromString('xl/styles.xml', $styles);
        foreach ($this->sheets as $i => $sh) {
            $zip->addFromString('xl/worksheets/sheet' . ($i + 1) . '.xml', self::sheetXml($sh));
        }
        $zip->close();
    }
}
