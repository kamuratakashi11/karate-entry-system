<?php
/**
 * 顧問出欠（Excel）を組み立てる。依存は ext-zip だけ。
 *
 * 見本は `2025関東予選顧問出欠とコート本部記録.xls`（孝さん・2026-09-29〜30）。2枚:
 *
 * 1. 「顧問出欠」… 通番・学校名・顧問名・1日目・2日目 の表を左右2列に並べ、左の下に「計」、
 *    右の下に「計」「総計」。ＭＳ Ｐゴシック 11pt・細線・A4縦1枚。見本との違いは**役割の列を
 *    足した**こと（審判・競技記録・係員。孝さん「色分けは要らない、役割が入っているほうがありがたい」）。
 * 2. 「コート・本部記録」… とりまとめの下書き:
 *    - コート: T1〜T4 のオフィシャルの学校（空欄。孝さんが書き入れる）と、その**候補**＝
 *      エントリー人数の多い6校（孝さん「人数が多い学校を選んで入れている。候補と人数を表示して
 *      くれればこちらで加工する。6校ほど見繕ってほしい」）
 *    - 本部記録: 役割が競技記録の先生を「両日／1日目のみ／2日目のみ」に分けて
 *    - コート運営係: 役割が係員の先生と出欠、割り振りを書き込む空欄
 *
 * 顧問の人数は大会ごとに変わるので、様式に埋めるのではなく**人数ぶんの行を作る**。
 * 計は COUNTIF の式にしておく（あとで ○×を手で直しても数え直される。見本と同じ）。
 */
declare(strict_types=1);

final class AdvisorSheet
{
    // 書式の番号（styles.xml の cellXfs の並び）
    const S_PLAIN = 0;   // 既定
    const S_HEAD  = 1;   // 見出し（中央・細線・縮小）
    const S_CELL  = 2;   // 本体（中央・細線・縮小）
    const S_SUM   = 3;   // 計（中央・細線）
    const S_TITLE = 4;   // 節の見出し（太字・左・線なし）
    const S_NOTE  = 5;   // 注記（左・線なし）

    // 1枚目: 1列目（B）から 通番・学校名・顧問名・役割・1日目・2日目。右の表は H から
    const HEAD   = ['', '学校名', '顧問名', '役割', '1日目', '2日目'];
    const WIDTHS = [3.45, 11.91, 12.0, 8.0, 6.27, 6.27];
    const LEFT   = 2;    // B
    const RIGHT  = 8;    // H

    const CANDIDATES = 6;   // コートの候補の数

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
     *        顧問（学校番号順・同じ学校の中は登録順）
     * @param list<array{school:string,entered:int}> $schools
     *        学校（学校番号順）とエントリー人数
     */
    public static function write(string $path, array $rows, array $schools = []): void
    {
        $sheets = [
            ['顧問出欠', self::attendance($rows)],
            ['コート・本部記録', self::courts($rows, $schools)],
        ];
        self::save($path, $sheets);
    }

    /** 1枚目「顧問出欠」のセル */
    private static function attendance(array $rows): array
    {
        $n = count($rows);
        // 左に多め（右の下には「計」と「総計」の2行が来るので、そのぶん左を長くする。
        // 見本は 47名 → 左25・右22）
        $leftN = min($n, max(1, (int)ceil(($n + 3) / 2)));
        $left = array_slice($rows, 0, $leftN);
        $right = array_slice($rows, $leftN);

        $cells = [];
        $put = function (int $r, int $c, $v, int $s, ?string $f = null) use (&$cells) {
            $cells[$r][$c] = [$v, $s, $f];
        };
        foreach ([self::LEFT, self::RIGHT] as $base) {
            if ($base === self::RIGHT && !$right) {
                break;
            }
            foreach (self::HEAD as $i => $h) {
                $put(1, $base + $i, $h, self::S_HEAD);
            }
        }
        $block = function (array $list, int $base, int $startNo) use ($put): array {
            $r = 2;
            foreach ($list as $k => $a) {
                $put($r, $base,     $startNo + $k, self::S_CELL);
                $put($r, $base + 1, $a['school'], self::S_CELL);
                $put($r, $base + 2, $a['name'], self::S_CELL);
                $put($r, $base + 3, $a['role'], self::S_CELL);
                $put($r, $base + 4, $a['d1'], self::S_CELL);
                $put($r, $base + 5, $a['d2'], self::S_CELL);
                $r++;
            }
            return [2, $r - 1];
        };
        $count = fn(array $list, string $key): int
            => count(array_filter($list, fn($a) => $a[$key] === '○'));

        [$l0, $l1] = $block($left, self::LEFT, 1);
        $lr = $l1 + 1;
        $put($lr, self::LEFT + 3, '計', self::S_SUM);
        foreach ([4 => 'd1', 5 => 'd2'] as $i => $key) {
            $c = self::col(self::LEFT + $i);
            $put($lr, self::LEFT + $i, $count($left, $key), self::S_SUM,
                 "COUNTIF({$c}{$l0}:{$c}{$l1},\"○\")");
        }
        if ($right) {
            [$r0, $r1] = $block($right, self::RIGHT, $leftN + 1);
            $rr = $r1 + 1;
            $put($rr, self::RIGHT + 3, '計', self::S_SUM);
            $put($rr + 1, self::RIGHT + 3, '総計', self::S_SUM);
            foreach ([4 => 'd1', 5 => 'd2'] as $i => $key) {
                $c = self::col(self::RIGHT + $i);
                $lc = self::col(self::LEFT + $i);
                $put($rr, self::RIGHT + $i, $count($right, $key), self::S_SUM,
                     "COUNTIF({$c}{$r0}:{$c}{$r1},\"○\")");
                $put($rr + 1, self::RIGHT + $i, $count($rows, $key), self::S_SUM,
                     "{$lc}{$lr}+{$c}{$rr}");
            }
        } else {
            // 右が無い（顧問が少ない）ときは左の下に「総計」も置く
            $put($lr + 1, self::LEFT + 3, '総計', self::S_SUM);
            foreach ([4 => 'd1', 5 => 'd2'] as $i => $key) {
                $c = self::col(self::LEFT + $i);
                $put($lr + 1, self::LEFT + $i, $count($rows, $key), self::S_SUM, "{$c}{$lr}");
            }
        }
        $widths = [1 => 2];
        foreach ([self::LEFT, self::RIGHT] as $base) {
            foreach (self::WIDTHS as $i => $w) {
                $widths[$base + $i] = $w;
            }
        }
        return ['cells' => $cells, 'widths' => $widths, 'fitH' => 1];   // 見本どおり1枚
    }

    /** 2枚目「コート・本部記録」のセル */
    private static function courts(array $rows, array $schools): array
    {
        $cells = [];
        $put = function (int $r, int $c, $v, int $s) use (&$cells) {
            $cells[$r][$c] = [$v, $s, null];
        };

        // ---- コート（左: T1〜T4 は空欄 ／ 右: 候補＝エントリー人数の多い6校）
        $put(1, 2, 'コート（タタミのオフィシャル）', self::S_TITLE);
        foreach (['コート', '学校名'] as $i => $h) {
            $put(2, 2 + $i, $h, self::S_HEAD);
        }
        foreach (['T1', 'T2', 'T3', 'T4'] as $i => $t) {
            $put(3 + $i, 2, $t, self::S_CELL);
            $put(3 + $i, 3, '', self::S_CELL);
        }
        $att = [];      // 学校 → [1日目の顧問数, 2日目の顧問数]
        foreach ($rows as $a) {
            $att[$a['school']][0] = ($att[$a['school']][0] ?? 0) + ($a['d1'] === '○' ? 1 : 0);
            $att[$a['school']][1] = ($att[$a['school']][1] ?? 0) + ($a['d2'] === '○' ? 1 : 0);
        }
        $cand = array_values(array_filter($schools, fn($s) => (int)$s['entered'] > 0));
        // 人数の多い順（同じなら学校番号順＝渡された順のまま）
        $order = array_keys($cand);
        usort($order, fn($x, $y) => [-(int)$cand[$x]['entered'], $x] <=> [-(int)$cand[$y]['entered'], $y]);
        $put(1, 5, '候補（エントリー人数の多い順）', self::S_TITLE);
        foreach (['', '学校名', 'エントリー', '顧問 1日目', '顧問 2日目'] as $i => $h) {
            $put(2, 5 + $i, $h, self::S_HEAD);
        }
        foreach (array_slice($order, 0, self::CANDIDATES) as $k => $idx) {
            $s = $cand[$idx];
            $r = 3 + $k;
            $put($r, 5, $k + 1, self::S_CELL);
            $put($r, 6, $s['school'], self::S_CELL);
            $put($r, 7, (int)$s['entered'], self::S_CELL);
            $put($r, 8, $att[$s['school']][0] ?? 0, self::S_CELL);
            $put($r, 9, $att[$s['school']][1] ?? 0, self::S_CELL);
        }
        $r = 3 + max(4, min(self::CANDIDATES, count($order))) + 1;

        // ---- 本部記録（役割＝競技記録）
        $r++;
        $put($r++, 2, '本部記録（役割が競技記録の先生）', self::S_TITLE);
        foreach (['区分', '学校名', '顧問名'] as $i => $h) {
            $put($r, 2 + $i, $h, self::S_HEAD);
        }
        $r++;
        $rec = array_values(array_filter($rows, fn($a) => $a['role'] === '競技記録'));
        $groups = [
            '両日'      => fn($a) => $a['d1'] === '○' && $a['d2'] === '○',
            '1日目のみ' => fn($a) => $a['d1'] === '○' && $a['d2'] !== '○',
            '2日目のみ' => fn($a) => $a['d1'] !== '○' && $a['d2'] === '○',
        ];
        $any = false;
        foreach ($groups as $label => $test) {
            foreach (array_values(array_filter($rec, $test)) as $a) {
                $put($r, 2, $label, self::S_CELL);
                $put($r, 3, $a['school'], self::S_CELL);
                $put($r, 4, $a['name'], self::S_CELL);
                $r++;
                $any = true;
            }
        }
        if (!$any) {
            $put($r++, 2, '（役割が競技記録の先生はいません）', self::S_NOTE);
        }

        // ---- コート運営係（役割＝係員）。割り振りは手で書き込む
        $r++;
        $put($r++, 2, 'コート運営係（役割が係員の先生）— 割り振りを書き込んでください', self::S_TITLE);
        foreach (['', '学校名', '顧問名', '1日目', '2日目', '割り振り'] as $i => $h) {
            $put($r, 2 + $i, $h, self::S_HEAD);
        }
        $r++;
        $staff = array_values(array_filter($rows, fn($a) => $a['role'] === '係員'));
        foreach ($staff as $k => $a) {
            $put($r, 2, $k + 1, self::S_CELL);
            $put($r, 3, $a['school'], self::S_CELL);
            $put($r, 4, $a['name'], self::S_CELL);
            $put($r, 5, $a['d1'], self::S_CELL);
            $put($r, 6, $a['d2'], self::S_CELL);
            $put($r, 7, '', self::S_CELL);
            $r++;
        }
        if (!$staff) {
            $put($r, 2, '（役割が係員の先生はいません）', self::S_NOTE);
        }
        $widths = [1 => 2, 2 => 9.5, 3 => 11.91, 4 => 12.0, 5 => 6.27, 6 => 11.91,
                   7 => 9.0, 8 => 9.0, 9 => 9.0];
        // 縦は伸びてよい（係員が多ければ2枚目に続く）
        return ['cells' => $cells, 'widths' => $widths, 'fitH' => 0];
    }

    /** セルの表 → シートの XML */
    private static function sheetXml(array $cells, array $widths, int $fitH): string
    {
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
        ksort($widths);
        $colsXml = '';
        foreach ($widths as $c => $w) {
            $colsXml .= "<col min=\"{$c}\" max=\"{$c}\" width=\"{$w}\" customWidth=\"1\"/>";
        }
        $lastRow = $cells ? max(array_keys($cells)) : 1;
        $dim = 'A1:' . self::col($maxCol) . $lastRow;
        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
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
            . "<pageSetup paperSize=\"9\" orientation=\"portrait\" fitToWidth=\"1\" fitToHeight=\"{$fitH}\"/>"
            . '</worksheet>';
    }

    /** @param list<array{0:string,1:array}> $sheets [シート名, [cells, widths]] */
    private static function save(string $path, array $sheets): void
    {
        $font = '<font><sz val="11"/><name val="ＭＳ Ｐゴシック"/><family val="3"/><charset val="128"/></font>';
        $bold = '<font><b/><sz val="11"/><name val="ＭＳ Ｐゴシック"/><family val="3"/><charset val="128"/></font>';
        $thin = '<border><left style="thin"/><right style="thin"/><top style="thin"/>'
              . '<bottom style="thin"/><diagonal/></border>';
        $center = '<alignment horizontal="center" vertical="center" shrinkToFit="1"/>';
        $styles = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            . '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            . "<fonts count=\"2\">{$font}{$bold}</fonts>"
            . '<fills count="2"><fill><patternFill patternType="none"/></fill>'
            . '<fill><patternFill patternType="gray125"/></fill></fills>'
            . "<borders count=\"2\"><border><left/><right/><top/><bottom/><diagonal/></border>{$thin}</borders>"
            . '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
            . '<cellXfs count="6">'
            . '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
            . "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"1\" xfId=\"0\" applyBorder=\"1\" applyAlignment=\"1\">{$center}</xf>"
            . "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"1\" xfId=\"0\" applyBorder=\"1\" applyAlignment=\"1\">{$center}</xf>"
            . '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1">'
            . '<alignment horizontal="center" vertical="center"/></xf>'
            . '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1">'
            . '<alignment horizontal="left" vertical="center"/></xf>'
            . '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1">'
            . '<alignment horizontal="left" vertical="center"/></xf>'
            . '</cellXfs>'
            . '<cellStyles count="1"><cellStyle name="標準" xfId="0" builtinId="0"/></cellStyles>'
            . '</styleSheet>';

        $overrides = '';
        $sheetList = '';
        $rels = '';
        foreach ($sheets as $i => [$name, $_]) {
            $n = $i + 1;
            $overrides .= "<Override PartName=\"/xl/worksheets/sheet{$n}.xml\" "
                . 'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>';
            $sheetList .= '<sheet name="' . self::esc($name) . "\" sheetId=\"{$n}\" r:id=\"rId{$n}\"/>";
            $rels .= "<Relationship Id=\"rId{$n}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet{$n}.xml\"/>";
        }
        $stylesId = count($sheets) + 1;

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
            . '<calcPr calcId="0" fullCalcOnLoad="1"/>'
            . '</workbook>');
        $zip->addFromString('xl/_rels/workbook.xml.rels',
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            . '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            . $rels
            . "<Relationship Id=\"rId{$stylesId}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>"
            . '</Relationships>');
        $zip->addFromString('xl/styles.xml', $styles);
        foreach ($sheets as $i => [$_, $sheet]) {
            $zip->addFromString('xl/worksheets/sheet' . ($i + 1) . '.xml',
                self::sheetXml($sheet['cells'], $sheet['widths'], $sheet['fitH']));
        }
        $zip->close();
    }
}
