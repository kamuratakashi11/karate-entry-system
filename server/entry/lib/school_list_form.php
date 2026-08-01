<?php
/**
 * 参加校一覧（新人大会の様式）の型埋め。
 *
 * 実物 `[訂正]2025新人大会参加校一覧.xlsx` と同じ表: 行＝学校、列＝種目のマトリクス。
 * 団体は ○、個人は「人数」と「シ（別枠シード）」の2列組。列の並びは実物どおり:
 *
 *   A 学校番号 / B 学校名 / C 参加校
 *   男子  D 団体形  E 団体組手(5人)  F 団体組手(3人)
 *         G 個人形  H シ    以降 3列おきに 階級・シ（J/K, M/N, P/Q, S/T, V/W）
 *   女子  Y 団体形  Z 団体組手(5人)  AA 団体組手(3人)
 *         AB 個人形 AC シ   以降 3列おきに（AE/AF, AH/AI, AK/AL, AN/AO, AQ/AR）
 *   AT 正選手合計 / AU 正選手＋補欠合計
 *   行75 学校数 / 行76 参加人数 / 行77 別枠シード / 行78 合計人数
 *
 * **数はすべて int で書く**（文字列だと Excel が数えず合計が 0 になる）。
 * **合計行は数式を置かず、ここで数えた結果を入れる**（孝さん・2026-07-31）。
 * 実物の数式は COUNTA が「人数が空でシードだけの学校」を数え落としており、
 * 「=O75+1」のような手直しで埋め合わせてあった。数え方をここに一本化する。
 *
 * AT / AU は**延べ数ではなく人数**（孝さん・2026-07-31）:
 *   AT ＝ 正選手またはシード選手である**人**の数
 *   AU ＝ 正選手・シード選手・補欠のいずれかである**人**の数
 */
final class SchoolListForm
{
    const ROW_FIRST   = 7;
    const ROW_LAST    = 74;
    const ROW_SCHOOLS = 75;   // 学校数
    const ROW_PEOPLE  = 76;   // 参加人数（＝正の人数の合計）
    const ROW_SEED    = 77;   // 別枠シード
    const ROW_TOTAL   = 78;   // 合計人数
    const TITLE       = 'A2';

    const COL_PART   = 3;                        // C 参加校
    const COL_TEAM   = ['m' => 4,  'w' => 25];   // 団体形 / +1 組手5人 / +2 組手3人
    const COL_KATA   = ['m' => 7,  'w' => 28];   // 個人形（+1 がシ）
    const COL_WEIGHT = ['m' => 10, 'w' => 31];   // 階級はここから3列おき
    const COL_TOTAL_REG = 46;   // AT
    const COL_TOTAL_ALL = 47;   // AU

    /**
     * @param array $schools 学校ごと: school_no / name / entries / m_mode / w_mode
     * @return array<string,int|string>
     */
    public static function cells(string $title, array $schools, array $weightsM, array $weightsW): array
    {
        $out = [self::TITLE => $title];
        $put = function (int $r, int $c, $v) use (&$out) {
            if ($v !== null && $v !== '' && $v !== 0) {
                $out[XlsxFill::colName($c) . $r] = $v;
            }
        };

        // 合計用の入れ物。列番号 => 合計
        $sumPeople = [];   // 人数列の合計
        $sumSeed   = [];   // シ列の合計
        $nSchools  = [];   // その列で1以上だった学校数
        $countUp = function (array &$box, int $c, int $v) {
            $box[$c] = ($box[$c] ?? 0) + $v;
        };

        $partSchools = 0;
        $totalReg = 0;
        $totalAll = 0;
        $r = self::ROW_FIRST;

        foreach ($schools as $s) {
            if ($r > self::ROW_LAST) {
                break;                      // 表からあふれた分は載せない
            }
            $t = self::tally($s);
            $out[XlsxFill::colName(1) . $r] = (int)$s['school_no'];
            $out[XlsxFill::colName(2) . $r] = (string)$s['name'];
            if ($t['any']) {
                $put($r, self::COL_PART, '○');
                $partSchools++;
            }

            foreach (['m' => $weightsM, 'w' => $weightsW] as $p => $weights) {
                $base = self::COL_TEAM[$p];
                if ($t[$p]['team_kata'] > 0) {
                    $put($r, $base, '○');
                    $countUp($nSchools, $base, 1);
                }
                if ($t[$p]['team_kumite'] > 0) {
                    $mode = ((string)($s[$p . '_mode'] ?? '5')) === '3' ? 2 : 1;
                    $put($r, $base + $mode, '○');
                    $countUp($nSchools, $base + $mode, 1);
                }
                // 個人形と各階級を同じ形で扱う
                $slots = [[self::COL_KATA[$p], $t[$p]['kata_reg'], $t[$p]['kata_seed']]];
                foreach ($weights as $i => $w) {
                    $slots[] = [self::COL_WEIGHT[$p] + $i * 3,
                                $t[$p]['ku_reg'][$w] ?? 0, $t[$p]['ku_seed'][$w] ?? 0];
                }
                foreach ($slots as [$c, $reg, $seed]) {
                    $put($r, $c, $reg);
                    $put($r, $c + 1, $seed);
                    $countUp($sumPeople, $c, $reg);
                    $countUp($sumSeed, $c, $seed);
                    if ($reg > 0 || $seed > 0) {   // シードだけの学校も数える（実物の数え落とし対策）
                        $countUp($nSchools, $c, 1);
                    }
                }
            }
            $put($r, self::COL_TOTAL_REG, $t['people_reg']);
            $put($r, self::COL_TOTAL_ALL, $t['people_all']);
            $totalReg += $t['people_reg'];
            $totalAll += $t['people_all'];
            $r++;
        }

        // --- 合計行（数式ではなく数えた結果） ---
        $put(self::ROW_SCHOOLS, self::COL_PART, $partSchools);
        foreach ($nSchools as $c => $n) {
            $put(self::ROW_SCHOOLS, $c, $n);
        }
        foreach ($sumPeople as $c => $n) {
            $put(self::ROW_PEOPLE, $c, $n);
            $put(self::ROW_TOTAL, $c, $n + ($sumSeed[$c] ?? 0));
        }
        foreach ($sumSeed as $c => $n) {
            $put(self::ROW_SEED, $c, $n);
            if (!isset($sumPeople[$c])) {
                $put(self::ROW_TOTAL, $c, $n);
            }
        }
        $put(self::ROW_SCHOOLS, self::COL_TOTAL_REG, $totalReg);
        $put(self::ROW_SCHOOLS, self::COL_TOTAL_ALL, $totalAll);
        return $out;
    }

    /** 1校ぶんの数え上げ。$s['entries'] は 1件＝1人 */
    private static function tally(array $s): array
    {
        $z = fn() => ['team_kata' => 0, 'team_kumite' => 0,
                      'kata_reg' => 0, 'kata_seed' => 0, 'ku_reg' => [], 'ku_seed' => []];
        $t = ['m' => $z(), 'w' => $z(), 'any' => false, 'people_reg' => 0, 'people_all' => 0];

        foreach ($s['entries'] as $e) {
            $p = (($e['sex'] ?? '') === '女子') ? 'w' : 'm';
            $isReg = false;    // この人は 正 or シード か
            $isAny = false;    // この人は 何かに出る（補欠を含む）か

            if (!empty($e['team_kata_chk'])) {
                $isAny = true;
                if (($e['team_kata_role'] ?? '') !== '補') {
                    $t[$p]['team_kata']++;
                    $isReg = true;
                }
            }
            if (!empty($e['team_kumi_chk'])) {
                $isAny = true;
                if (($e['team_kumi_role'] ?? '') !== '補') {
                    $t[$p]['team_kumite']++;
                    $isReg = true;
                }
            }
            if (!empty($e['kata_chk'])) {
                $isAny = true;
                $v = (string)($e['kata_val'] ?? '正');
                if ($v === 'シード') {
                    $t[$p]['kata_seed']++;
                    $isReg = true;
                } elseif ($v !== '補') {
                    $t[$p]['kata_reg']++;
                    $isReg = true;
                }
            }
            if (!empty($e['kumi_chk'])) {
                $isAny = true;
                $w = (string)($e['kumi_val'] ?? '');
                $v = (string)($e['kumi_sub_val'] ?? '正');
                if ($v === 'シード') {
                    $t[$p]['ku_seed'][$w] = ($t[$p]['ku_seed'][$w] ?? 0) + 1;
                    $isReg = true;
                } elseif ($v !== '補') {
                    $t[$p]['ku_reg'][$w] = ($t[$p]['ku_reg'][$w] ?? 0) + 1;
                    $isReg = true;
                }
            }
            if ($isAny) {
                $t['any'] = true;
                $t['people_all']++;          // 正・シード・補欠のいずれか＝1人
            }
            if ($isReg) {
                $t['people_reg']++;          // 正またはシード＝1人
            }
        }
        return $t;
    }
}
