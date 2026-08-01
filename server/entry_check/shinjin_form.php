<?php
/**
 * 新人戦の申込書: 座標(coords_shinjin.json) ＋ 学校・部員データ → セル連想配列。
 *
 * 申込システム app.py の generate_excel（階級制=shinjin の分岐）の移植。
 * 文字の組み立て・並べ替え・行の割り付けの規則を同じにしてある:
 * - 年度は「8年」→「8」（safe_write の末尾「年」除去と同じ）
 * - 顧問は4名まで。引率は ○/×
 * - 部員は 学年で絞り込み（新人戦は1・2年）→ display_order → 男子先 → 学年降順 → 名前順
 * - 行 = start_row + (i ÷ cap)×offset + (i mod cap)   … cap人で次のブロックへ
 * - 団体 = ○/補、個人形 = ○順位/シ順位/補、個人組手 = 階級の列に ○順位/シ順位/補
 */
final class ShinjinForm
{
    /**
     * @param array $coords coords_shinjin.json の中身
     * @param array $school year / tournament_name / date? / school_name / principal /
     *                      advisors[{name,d1,d2}] / m_kumite_mode / w_kumite_mode
     * @param array $members 部員の配列。name/grade/dob/jkf_no/sex/display_order と
     *                       エントリー: team_kata_chk,team_kata_role,team_kumi_chk,
     *                       team_kumi_role,kata_chk,kata_val,kata_rank,
     *                       kumi_chk,kumi_val(=階級),kumi_sub_val,kumi_rank
     * @param int[] $grades 出場できる学年（新人戦は [1,2]）
     * @return array<string,string> セル参照 => 値
     */
    public static function cells(array $coords, array $school, array $members, array $grades = [1, 2]): array
    {
        $out = [];
        $put = function ($target, $value) use (&$out) {
            if (!$target) {
                return;
            }
            $ref = is_array($target) ? XlsxFill::colName((int)$target[1]) . (int)$target[0] : (string)$target;
            $out[$ref] = (string)$value;
        };

        $year = (string)($school['year'] ?? '');
        if (preg_match('/^([0-9]+)年$/u', $year, $m)) {
            $year = $m[1];
        }
        $put($coords['year'] ?? null, $year);
        $put($coords['tournament_name'] ?? null, (string)($school['tournament_name'] ?? ''));
        $date = (string)($school['date'] ?? ('令和' . (idate('Y') - 2018) . '年' . idate('n') . '月' . idate('j') . '日'));
        $put($coords['date'] ?? null, $date);
        $put($coords['school_name'] ?? null, (string)($school['school_name'] ?? ''));
        $put($coords['principal'] ?? null, (string)($school['principal'] ?? ''));

        $advs = array_values($school['advisors'] ?? []);
        $put($coords['head_advisor'] ?? null, (string)($advs[0]['name'] ?? ''));
        foreach (array_slice($advs, 0, 4) as $i => $a) {
            $c = $coords['advisors'][$i] ?? null;
            if (!$c) {
                continue;
            }
            $put($c['name'] ?? null, (string)($a['name'] ?? ''));
            $put($c['d1'] ?? null, !empty($a['d1']) ? '○' : '×');
            $put($c['d2'] ?? null, !empty($a['d2']) ? '○' : '×');
        }

        // --- 部員の絞り込みと並べ替え（app.py と同じ基準） ---
        $list = array_values(array_filter(
            $members,
            fn($mm) => in_array((int)($mm['grade'] ?? 0), $grades, true)
        ));
        usort($list, function ($a, $b) {
            $cmp = self::orderKey($a) <=> self::orderKey($b);
            if ($cmp !== 0) {
                return $cmp;
            }
            $cmp = self::sexRank($a) <=> self::sexRank($b);
            if ($cmp !== 0) {
                return $cmp;
            }
            $cmp = self::gradeRank($a) <=> self::gradeRank($b);
            if ($cmp !== 0) {
                return $cmp;
            }
            return strcmp((string)($a['name'] ?? ''), (string)($b['name'] ?? ''));
        });

        $cols = $coords['cols'];
        $sr = (int)$coords['start_row'];
        $cap = (int)$coords['cap'];
        $off = (int)$coords['offset'];

        foreach ($list as $i => $mm) {
            $r = $sr + intdiv($i, $cap) * $off + ($i % $cap);
            $sexP = (($mm['sex'] ?? '') === '女子') ? 'w' : 'm';

            $put([$r, $cols['name']], $mm['name'] ?? '');
            $put([$r, $cols['grade']], $mm['grade'] ?? '');
            $put([$r, $cols['dob']], $mm['dob'] ?? '');
            $put([$r, $cols['jkf_no']], $mm['jkf_no'] ?? '');

            if (!empty($mm['team_kata_chk'])) {
                $c = $cols["{$sexP}_team_kata"] ?? null;
                if ($c) {
                    $put([$r, $c], ($mm['team_kata_role'] ?? '') === '補' ? '補' : '○');
                }
            }
            if (!empty($mm['team_kumi_chk'])) {
                $mode = (string)($school["{$sexP}_kumite_mode"] ?? '5');
                $c = $cols["{$sexP}_team_kumite_{$mode}"] ?? null;
                if ($c) {
                    $put([$r, $c], ($mm['team_kumi_role'] ?? '') === '補' ? '補' : '○');
                }
            }
            if (!empty($mm['kata_chk'])) {
                $c = $cols["{$sexP}_kata"] ?? null;
                if ($c) {
                    $put([$r, $c], self::markWithRank((string)($mm['kata_val'] ?? ''), (string)($mm['kata_rank'] ?? '')));
                }
            }
            if (!empty($mm['kumi_chk'])) {
                $weight = (string)($mm['kumi_val'] ?? '');
                $c = $cols["{$sexP}_kumite_{$weight}"] ?? null;
                if ($c) {
                    $sub = (string)($mm['kumi_sub_val'] ?? '正');
                    $put([$r, $c], self::markWithRank($sub, (string)($mm['kumi_rank'] ?? '')));
                }
            }
        }
        return $out;
    }

    private static function markWithRank(string $v, string $rank): string
    {
        if ($v === '補') {
            return '補';
        }
        return ($v === 'シード' ? 'シ' : '○') . $rank;
    }

    private static function orderKey(array $m): float
    {
        $v = trim((string)($m['display_order'] ?? ''));
        return ($v !== '' && is_numeric($v)) ? (float)$v : 999999.0;
    }

    private static function sexRank(array $m): int
    {
        return (($m['sex'] ?? '') === '女子') ? 1 : 0;
    }

    private static function gradeRank(array $m): int
    {
        return match ((int)($m['grade'] ?? 0)) {
            3 => 0,
            2 => 1,
            1 => 2,
            default => 3,
        };
    }
}
