<?php
/**
 * 計量チェック表（Excel）。管理画面の「データ出力」から書き出す。
 *
 * 新人大会の計量で、階級ごとに選手を確かめて体重を書き込む表（孝さん 2026-10-19
 * 「申し込みの段階で何kg級かがわかっているから、それぞれの階級ごとにチェックシートを」）。
 * 階級は申込の個人組手の階級（entries の kumi_val）から取る。
 *
 * シート:
 *   一覧 … 階級ごとの人数と、計量済みの数（各シートの「計量」欄の ○ を数える式）
 *   男子 -55kg級 など … 階級ごとに 1 枚。学校番号順・学校の中は名簿の順（補欠は学校の最後）。
 *                        計量（○）・体重・備考は当日に書き込む欄
 *
 * 合否（階級に収まっているか）の判定は入れない。計量の決まり（許容の幅など）は
 * RULES.md に無いので、体重を見て人が決める。
 */
declare(strict_types=1);

require_once __DIR__ . '/xlsx_book.php';

final class WeighSheet
{
    const HEAD_ROW = 4;    // 見出しの行（印刷の各ページに繰り返す）
    const CHECK_COL = 6;   // 「計量」の列（F）。一覧の式が数える
    const HEAD = ['No', '学校名', '氏名', '学年', '区分', '計量', '体重(kg)', '備考'];
    const WIDTHS = [5, 14, 16, 5, 7, 7, 10, 22];

    /**
     * 大会の個人組手の選手を、階級ごとに集める。
     *
     * @param array{男子:list<string>,女子:list<string>} $weights 階級（「-55kg級」の形。weight_list() の戻り）
     * @param list<array{school_id:string,name:string}> $schools 学校（学校番号順）
     * @param callable(string):list<array> $membersOf 学校 → 部員（名簿の順）
     * @param callable(string):array<string,array> $entriesOf 学校 → 氏名 → エントリー
     * @return list<array{sex:string,class:string,players:list<array>}>
     *         階級（男子の軽い順 → 女子の軽い順）。エントリーが無い階級も空で入る
     */
    public static function collect(array $weights, array $schools, callable $membersOf, callable $entriesOf): array
    {
        $out = [];
        foreach (['男子', '女子'] as $sex) {
            foreach ($weights[$sex] ?? [] as $class) {
                $out["{$sex}|{$class}"] = ['sex' => $sex, 'class' => $class, 'players' => []];
            }
        }
        foreach ($schools as $s) {
            $sid = (string)$s['school_id'];
            $ents = $entriesOf($sid);
            $subs = [];
            foreach ($membersOf($sid) as $m) {
                $e = $ents[(string)$m['name']] ?? null;
                if (!$e || empty($e['kumi_chk'])) {
                    continue;
                }
                $key = "{$m['sex']}|" . (string)($e['kumi_val'] ?? '');
                if (!isset($out[$key])) {
                    continue;      // 階級の無い大会・階級を選んでいない
                }
                $kubun = (string)(($e['kumi_sub_val'] ?? '') ?: '正');
                $row = ['key' => $key, 'school' => (string)$s['name'], 'name' => (string)$m['name'],
                        'grade' => (string)$m['grade'], 'kubun' => $kubun];
                if ($kubun === '補') {
                    $subs[] = $row;
                } else {
                    $out[$key]['players'][] = $row;
                }
            }
            foreach ($subs as $row) {
                $out[$row['key']]['players'][] = $row;
            }
        }
        return array_values($out);
    }

    /** 階級 → シート名（Excel の 31 文字の制限に収まる） */
    public static function sheetName(array $c): string
    {
        return mb_substr("{$c['sex']} {$c['class']}", 0, 31);
    }

    /** @param list<array{sex:string,class:string,players:list<array>}> $classes */
    public static function book(string $title, array $classes): XlsxBook
    {
        $b = new XlsxBook();
        $sum = $b->sheet('一覧', [12, 14, 8, 10, 8]);
        $sum->put(1, 1, "{$title}　計量チェック表", XlsxBook::TITLE);
        $sum->put(2, 1, '各階級のシートの「計量」欄に ○ を入れると、計量済みが数え直されます。', XlsxBook::NOTE);
        $sum->put(3, 1, '補欠も載せてあります（区分が「補」）。', XlsxBook::NOTE);
        foreach (['種目', '階級', '人数', '計量済み', '残り'] as $i => $h) {
            $sum->put(self::HEAD_ROW, $i + 1, $h, XlsxBook::HEAD);
        }
        $sum->headerRows = self::HEAD_ROW . ':' . self::HEAD_ROW;
        $r = self::HEAD_ROW + 1;
        $first = $r;
        foreach ($classes as $c) {
            $n = count($c['players']);
            $sum->put($r, 1, "{$c['sex']}個人組手", XlsxBook::CELL);
            $sum->put($r, 2, $c['class'], XlsxBook::CELL);
            $sum->put($r, 3, $n, XlsxBook::CELL);
            if ($n > 0) {
                $name = self::sheetName($c);
                $col = XlsxBook::col(self::CHECK_COL);
                $top = self::HEAD_ROW + 1;
                $end = self::HEAD_ROW + $n;
                $ref = "'" . str_replace("'", "''", $name) . "'!{$col}{$top}:{$col}{$end}";
                $sum->put($r, 4, 0, XlsxBook::CELL, "COUNTIF({$ref},\"○\")");
                $sum->put($r, 5, $n, XlsxBook::CELL, "C{$r}-D{$r}");
            } else {
                $sum->put($r, 4, '—', XlsxBook::CELL);
                $sum->put($r, 5, '—', XlsxBook::CELL);
            }
            $r++;
        }
        if ($r === $first) {
            $sum->put($r, 1, '（この大会には階級がありません）', XlsxBook::NOTE);
        } else {
            $last = $r - 1;
            $sum->put($r, 2, '計', XlsxBook::HEAD);
            foreach ([3 => 'C', 4 => 'D', 5 => 'E'] as $ci => $cl) {
                $sum->put($r, $ci, 0, XlsxBook::CELL, "SUM({$cl}{$first}:{$cl}{$last})");
            }
        }

        foreach ($classes as $c) {
            if (!$c['players']) {
                continue;
            }
            $s = $b->sheet(self::sheetName($c), self::WIDTHS);
            $n = count($c['players']);
            $s->put(1, 1, "{$title}　{$c['sex']}個人組手 {$c['class']}　計量チェック表", XlsxBook::TITLE);
            $col = XlsxBook::col(self::CHECK_COL);
            $top = self::HEAD_ROW + 1;
            $end = self::HEAD_ROW + $n;
            $s->put(2, 1, '', XlsxBook::NOTE,
                    "\"{$n}名　計量済み \"&COUNTIF({$col}{$top}:{$col}{$end},\"○\")&\"名\"");
            foreach (self::HEAD as $i => $h) {
                $s->put(self::HEAD_ROW, $i + 1, $h, XlsxBook::HEAD);
            }
            $s->headerRows = self::HEAD_ROW . ':' . self::HEAD_ROW;
            foreach ($c['players'] as $k => $p) {
                $r = $top + $k;
                $s->put($r, 1, $k + 1, XlsxBook::CELL);
                $s->put($r, 2, $p['school'], XlsxBook::LEFT);
                $s->put($r, 3, $p['name'], XlsxBook::LEFT);
                $s->put($r, 4, is_numeric($p['grade']) ? (int)$p['grade'] : $p['grade'], XlsxBook::CELL);
                $s->put($r, 5, $p['kubun'], XlsxBook::CELL);
                foreach ([6, 7, 8] as $ci) {
                    $s->put($r, $ci, '', XlsxBook::INPUT);
                }
            }
        }
        return $b;
    }
}
