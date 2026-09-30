<?php
/**
 * 顧問出欠（Excel）を組み立てる。依存は ext-zip だけ。
 *
 * 見本は `2025関東予選顧問出欠とコート本部記録.xls` の「顧問出欠」シート（孝さん・2026-09-29）:
 *   通番・学校名・顧問名・1日目・2日目 の表を左右2列に並べ、左の下に「計」、
 *   右の下に「計」「総計」。ＭＳ Ｐゴシック 11pt・細線・A4縦1枚。
 * 見本との違いは**役割の列を足した**こと（審判・競技記録・係員。孝さん「色分けは要らない、
 * 役割が入っているほうがありがたい」）。
 *
 * 顧問の人数は大会ごとに変わるので、様式に埋めるのではなく**人数ぶんの行を作る**。
 * 計は COUNTIF の式にしておく（あとで ○×を手で直しても数え直される。見本と同じ）。
 */
declare(strict_types=1);

final class AdvisorSheet
{
    // 1列目（B）から: 通番・学校名・顧問名・役割・1日目・2日目。右の表は H から
    const HEAD   = ['', '学校名', '顧問名', '役割', '1日目', '2日目'];
    const WIDTHS = [3.45, 11.91, 12.0, 8.0, 6.27, 6.27];
    const LEFT   = 2;    // B
    const RIGHT  = 8;    // H

    /** 列番号 → 列名（1 = A） */
    private static function col(int $n): string
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

    /**
     * @param list<array{school:string,name:string,role:string,d1:string,d2:string}> $rows
     *        学校番号順・同じ学校の中は登録順に並べて渡す
     */
    public static function write(string $path, array $rows): void
    {
        $n = count($rows);
        // 左に多め（右の下には「計」と「総計」の2行が来るので、そのぶん左を長くする。
        // 見本は 47名 → 左25・右22）
        $leftN = max(1, (int)ceil(($n + 3) / 2));
        if ($leftN > $n) {
            $leftN = $n;
        }
        $left = array_slice($rows, 0, $leftN);
        $right = array_slice($rows, $leftN);

        $cells = [];     // [row][col] => [value, style, formula?]
        $put = function (int $r, int $c, $v, int $s, ?string $f = null) use (&$cells) {
            $cells[$r][$c] = [$v, $s, $f];
        };
        // 見出し
        foreach ([self::LEFT, self::RIGHT] as $base) {
            if ($base === self::RIGHT && !$right) {
                break;
            }
            foreach (self::HEAD as $i => $h) {
                $put(1, $base + $i, $h, 1);
            }
        }
        // 本体
        $block = function (array $list, int $base, int $startNo) use ($put): array {
            $r = 2;
            foreach ($list as $k => $a) {
                $put($r, $base,     $startNo + $k, 2);
                $put($r, $base + 1, $a['school'], 2);
                $put($r, $base + 2, $a['name'], 2);
                $put($r, $base + 3, $a['role'], 2);
                $put($r, $base + 4, $a['d1'], 2);
                $put($r, $base + 5, $a['d2'], 2);
                $r++;
            }
            return [2, $r - 1];       // 本体の最初と最後の行
        };
        [$l0, $l1] = $block($left, self::LEFT, 1);
        $count = function (array $list, string $key): int {
            return count(array_filter($list, fn($a) => $a[$key] === '○'));
        };
        // 左の「計」
        $lr = $l1 + 1;
        $put($lr, self::LEFT + 3, '計', 3);
        foreach ([4 => 'd1', 5 => 'd2'] as $i => $key) {
            $c = self::col(self::LEFT + $i);
            $put($lr, self::LEFT + $i, $count($left, $key), 3, "COUNTIF({$c}{$l0}:{$c}{$l1},\"○\")");
        }
        if ($right) {
            [$r0, $r1] = $block($right, self::RIGHT, $leftN + 1);
            $rr = $r1 + 1;
            $put($rr, self::RIGHT + 3, '計', 3);
            $put($rr + 1, self::RIGHT + 3, '総計', 3);
            foreach ([4 => 'd1', 5 => 'd2'] as $i => $key) {
                $c = self::col(self::RIGHT + $i);
                $lc = self::col(self::LEFT + $i);
                $put($rr, self::RIGHT + $i, $count($right, $key), 3,
                     "COUNTIF({$c}{$r0}:{$c}{$r1},\"○\")");
                $put($rr + 1, self::RIGHT + $i, $count($left, $key) + $count($right, $key), 3,
                     "{$lc}{$lr}+{$c}{$rr}");
            }
        } else {
            // 右が無い（顧問が少ない）ときは左の下に「総計」も置く
            $put($lr + 1, self::LEFT + 3, '総計', 3);
            foreach ([4, 5] as $i) {
                $c = self::col(self::LEFT + $i);
                $put($lr + 1, self::LEFT + $i, $cells[$lr][self::LEFT + $i][0], 3, "{$c}{$lr}");
            }
        }

        // シートの XML
        ksort($cells);
        $xmlRows = '';
        $maxCol = 1;
        foreach ($cells as $r => $cols) {
            ksort($cols);
            $xmlRows .= "<row r=\"{$r}\">";
            foreach ($cols as $c => [$v, $s, $f]) {
                $maxCol = max($maxCol, $c);
                $ref = self::col($c) . $r;
                if ($f !== null) {
                    $xmlRows .= "<c r=\"{$ref}\" s=\"{$s}\"><f>" . self::esc($f) . "</f><v>{$v}</v></c>";
                } elseif (is_int($v)) {
                    $xmlRows .= "<c r=\"{$ref}\" s=\"{$s}\"><v>{$v}</v></c>";
                } elseif ($v === '') {
                    $xmlRows .= "<c r=\"{$ref}\" s=\"{$s}\"/>";
                } else {
                    $xmlRows .= "<c r=\"{$ref}\" s=\"{$s}\" t=\"inlineStr\"><is><t>"
                        . self::esc((string)$v) . "</t></is></c>";
                }
            }
            $xmlRows .= '</row>';
        }
        $colsXml = '<col min="1" max="1" width="2" customWidth="1"/>';
        foreach ([self::LEFT, self::RIGHT] as $base) {
            foreach (self::WIDTHS as $i => $w) {
                $c = $base + $i;
                $colsXml .= "<col min=\"{$c}\" max=\"{$c}\" width=\"{$w}\" customWidth=\"1\"/>";
            }
        }
        $lastRow = max(array_keys($cells));
        $dim = 'A1:' . self::col($maxCol) . $lastRow;
        $sheet = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            . '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            . 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            . '<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>'
            . "<dimension ref=\"{$dim}\"/>"
            . '<sheetViews><sheetView workbookViewId="0"/></sheetViews>'
            . '<sheetFormatPr defaultRowHeight="13.5"/>'
            . "<cols>{$colsXml}</cols>"
            . "<sheetData>{$xmlRows}</sheetData>"
            . '<printOptions horizontalCentered="1"/>'
            . '<pageMargins left="0.59" right="0.59" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'
            . '<pageSetup paperSize="9" orientation="portrait" fitToWidth="1" fitToHeight="1"/>'
            . '</worksheet>';

        // 書式: 0 既定 / 1 見出し（中央・細線・縮小） / 2 本体（中央・細線・縮小） / 3 計（中央・細線）
        $font = '<font><sz val="11"/><name val="ＭＳ Ｐゴシック"/><family val="3"/><charset val="128"/></font>';
        $thin = '<border><left style="thin"/><right style="thin"/><top style="thin"/>'
              . '<bottom style="thin"/><diagonal/></border>';
        $center = '<alignment horizontal="center" vertical="center" shrinkToFit="1"/>';
        $styles = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            . '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            . "<fonts count=\"1\">{$font}</fonts>"
            . '<fills count="2"><fill><patternFill patternType="none"/></fill>'
            . '<fill><patternFill patternType="gray125"/></fill></fills>'
            . "<borders count=\"2\"><border><left/><right/><top/><bottom/><diagonal/></border>{$thin}</borders>"
            . '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
            . '<cellXfs count="4">'
            . '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
            . "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"1\" xfId=\"0\" applyBorder=\"1\" applyAlignment=\"1\">{$center}</xf>"
            . "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"1\" xfId=\"0\" applyBorder=\"1\" applyAlignment=\"1\">{$center}</xf>"
            . '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1">'
            . '<alignment horizontal="center" vertical="center"/></xf>'
            . '</cellXfs>'
            . '<cellStyles count="1"><cellStyle name="標準" xfId="0" builtinId="0"/></cellStyles>'
            . '</styleSheet>';

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
            . '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
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
            . '<sheets><sheet name="顧問出欠" sheetId="1" r:id="rId1"/></sheets>'
            . '<calcPr calcId="0" fullCalcOnLoad="1"/>'
            . '</workbook>');
        $zip->addFromString('xl/_rels/workbook.xml.rels',
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            . '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            . '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
            . '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
            . '</Relationships>');
        $zip->addFromString('xl/styles.xml', $styles);
        $zip->addFromString('xl/worksheets/sheet1.xml', $sheet);
        $zip->close();
    }
}
