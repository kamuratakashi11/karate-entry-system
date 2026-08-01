<?php
/**
 * xlsx の「セルの値だけ」を書き換える薄い型埋めエンジン。依存は ext-zip / ext-dom のみ。
 *
 * 罫線・結合セル・列幅・印刷設定には一切触れない: 書き換えるのはシートXMLのセル値
 * だけで、zip 内の他のエントリはバイト単位でそのまま残る（PhpSpreadsheet のように
 * ブック全体を書き直さないので、テンプレートの体裁が崩れようがない）。
 *
 * openpyxl 側（申込システム app.py の safe_write）と合わせている仕様:
 * - 結合セルの中への書き込みは、結合範囲の左上（アンカー）に付け替える
 * - 値はすべて文字列（inlineStr）として書く（現行も str() してから書いている）
 */
final class XlsxFill
{
    const NS     = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
    const NS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

    /**
     * @param array<string,string> $cells 'E3' => '値' の連想配列
     * @return array<string,string> 実際に書いたセル（結合の付け替え後の参照 => 値）。検証用
     */
    public static function fill(string $templatePath, string $outPath, array $cells): array
    {
        if (!is_file($templatePath)) {
            throw new RuntimeException("テンプレートが見つからない: {$templatePath}（アップロード漏れ・名前違いを確認）");
        }
        if (!is_readable($templatePath)) {
            throw new RuntimeException("テンプレートを読めない（権限）: {$templatePath}");
        }
        if (!@copy($templatePath, $outPath)) {
            $err = error_get_last();
            throw new RuntimeException("出力先に書けない: {$outPath}"
                . ($err ? "（OSの報告: {$err['message']}）" : ''));
        }
        $zip = new ZipArchive();
        if ($zip->open($outPath) !== true) {
            throw new RuntimeException("zipとして開けない: {$outPath}");
        }
        try {
            $sheetPath = self::firstSheetPath($zip);
            $xml = $zip->getFromName($sheetPath);
            if ($xml === false) {
                throw new RuntimeException("シートが見つからない: {$sheetPath}");
            }
            $doc = new DOMDocument();
            if (!$doc->loadXML($xml)) {
                throw new RuntimeException('シートXMLを解釈できない');
            }
            $written = self::applyCells($doc, $cells);
            if ($zip->addFromString($sheetPath, $doc->saveXML()) !== true) {
                throw new RuntimeException('シートXMLの書き戻しに失敗');
            }
        } finally {
            $zip->close();
        }
        return $written;
    }

    /** 列番号(1始まり) → 列名。1 → A, 27 → AA */
    public static function colName(int $n): string
    {
        $s = '';
        while ($n > 0) {
            $n--;
            $s = chr(65 + $n % 26) . $s;
            $n = intdiv($n, 26);
        }
        return $s;
    }

    /** 列名 → 列番号(1始まり) */
    public static function colNum(string $col): int
    {
        $n = 0;
        foreach (str_split($col) as $ch) {
            $n = $n * 26 + (ord($ch) - 64);
        }
        return $n;
    }

    /** 'C8' → ['C', 8] */
    private static function splitRef(string $ref): array
    {
        if (!preg_match('/^([A-Z]+)([0-9]+)$/', $ref, $m)) {
            throw new RuntimeException("セル参照が不正: {$ref}");
        }
        return [$m[1], (int)$m[2]];
    }

    /** ブックの先頭シートの zip 内パス（例 xl/worksheets/sheet1.xml） */
    private static function firstSheetPath(ZipArchive $zip): string
    {
        $wbXml = $zip->getFromName('xl/workbook.xml');
        $relXml = $zip->getFromName('xl/_rels/workbook.xml.rels');
        if ($wbXml === false || $relXml === false) {
            throw new RuntimeException('workbook.xml が見つからない（xlsxではない？）');
        }
        $wb = new DOMDocument();
        $wb->loadXML($wbXml);
        $sheet = $wb->getElementsByTagNameNS(self::NS, 'sheet')->item(0);
        if (!$sheet) {
            throw new RuntimeException('シート定義が無い');
        }
        $rid = $sheet->getAttributeNS(self::NS_REL, 'id');
        $rels = new DOMDocument();
        $rels->loadXML($relXml);
        foreach ($rels->getElementsByTagName('Relationship') as $rel) {
            if ($rel->getAttribute('Id') === $rid) {
                $t = $rel->getAttribute('Target');
                return str_starts_with($t, '/') ? ltrim($t, '/') : 'xl/' . $t;
            }
        }
        throw new RuntimeException("シートのrelが見つからない: {$rid}");
    }

    /** @return array<string,string> */
    private static function applyCells(DOMDocument $doc, array $cells): array
    {
        // 結合範囲（書き込み先の付け替えに使う）
        $merges = [];
        foreach ($doc->getElementsByTagNameNS(self::NS, 'mergeCell') as $m) {
            $parts = explode(':', $m->getAttribute('ref'));
            if (count($parts) !== 2) {
                continue;
            }
            [$c1, $r1] = self::splitRef($parts[0]);
            [$c2, $r2] = self::splitRef($parts[1]);
            $merges[] = [self::colNum($c1), $r1, self::colNum($c2), $r2, $parts[0]];
        }
        $sheetData = $doc->getElementsByTagNameNS(self::NS, 'sheetData')->item(0);
        if (!$sheetData) {
            throw new RuntimeException('sheetData が無い');
        }
        $rowsByNum = [];
        foreach ($sheetData->childNodes as $n) {
            if ($n instanceof DOMElement && $n->localName === 'row') {
                $rowsByNum[(int)$n->getAttribute('r')] = $n;
            }
        }

        $written = [];
        foreach ($cells as $ref => $value) {
            $ref = strtoupper(trim((string)$ref));
            [$col, $rowNum] = self::splitRef($ref);
            $cn = self::colNum($col);
            foreach ($merges as [$mc1, $mr1, $mc2, $mr2, $anchor]) {
                if ($cn >= $mc1 && $cn <= $mc2 && $rowNum >= $mr1 && $rowNum <= $mr2) {
                    $ref = $anchor;
                    [$col, $rowNum] = self::splitRef($ref);
                    $cn = self::colNum($col);
                    break;
                }
            }
            $row = $rowsByNum[$rowNum] ?? self::insertRow($doc, $sheetData, $rowNum);
            $rowsByNum[$rowNum] = $row;
            self::setCellValue($doc, $row, $ref, $cn, (string)$value);
            $written[$ref] = (string)$value;
        }
        return $written;
    }

    private static function insertRow(DOMDocument $doc, DOMElement $sheetData, int $rowNum): DOMElement
    {
        $row = $doc->createElementNS(self::NS, 'row');
        $row->setAttribute('r', (string)$rowNum);
        foreach ($sheetData->childNodes as $n) {
            if ($n instanceof DOMElement && $n->localName === 'row'
                && (int)$n->getAttribute('r') > $rowNum) {
                $sheetData->insertBefore($row, $n);
                return $row;
            }
        }
        $sheetData->appendChild($row);
        return $row;
    }

    private static function setCellValue(DOMDocument $doc, DOMElement $row, string $ref, int $cn, string $value): void
    {
        $target = null;
        $before = null;
        foreach ($row->childNodes as $n) {
            if (!($n instanceof DOMElement) || $n->localName !== 'c') {
                continue;
            }
            $cr = $n->getAttribute('r');
            if ($cr === $ref) {
                $target = $n;
                break;
            }
            if ($before === null && preg_match('/^([A-Z]+)/', $cr, $m)
                && self::colNum($m[1]) > $cn) {
                $before = $n;
            }
        }
        if (!$target) {
            // テンプレートに無いセルは新規（スタイル無し＝既定の見た目になる。
            // 座標が枠の中を指している限りここには来ない）
            $target = $doc->createElementNS(self::NS, 'c');
            $target->setAttribute('r', $ref);
            $before ? $row->insertBefore($target, $before) : $row->appendChild($target);
        }
        // 値だけ入れ替える。スタイル属性 s（罫線・フォント・配置）は残す
        $target->removeAttribute('t');
        while ($target->firstChild) {
            $target->removeChild($target->firstChild);
        }
        $target->setAttribute('t', 'inlineStr');
        $is = $doc->createElementNS(self::NS, 'is');
        $t = $doc->createElementNS(self::NS, 't');
        $t->setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve');
        $t->appendChild($doc->createTextNode($value));
        $is->appendChild($t);
        $target->appendChild($is);
    }
}
