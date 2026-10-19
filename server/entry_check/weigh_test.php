<?php
/**
 * 計量チェック表（server/entry/lib/weigh_sheet.php）の回帰テスト。架空のデータだけを使う。
 *
 *   php server/entry_check/weigh_test.php [出力先.xlsx]
 *
 * 終了コード 0 ＝ 全部通過。出力先を渡すと、できた Excel をそこに残す（目で見る用）。
 */
declare(strict_types=1);
error_reporting(E_ALL);

require __DIR__ . '/../entry/lib/weigh_sheet.php';

$fails = 0;
function check(bool $ok, string $what): void
{
    global $fails;
    echo ($ok ? 'OK  ' : 'NG  ') . $what . "\n";
    if (!$ok) {
        $fails++;
    }
}

$weights = ['男子' => ['-55kg級', '-61kg級', '+76kg級'], '女子' => ['-48kg級', '+66kg級']];
$schools = [['school_id' => 's2', 'name' => '二高'], ['school_id' => 's1', 'name' => '一高']];
$members = [
    's1' => [
        ['name' => '一高 補欠', 'sex' => '男子', 'grade' => '1'],
        ['name' => '一高 太郎', 'sex' => '男子', 'grade' => '2'],
        ['name' => '一高 花子', 'sex' => '女子', 'grade' => '1'],
        ['name' => '一高 形だけ', 'sex' => '男子', 'grade' => '2'],
    ],
    's2' => [
        ['name' => '二高 次郎', 'sex' => '男子', 'grade' => '2'],
        ['name' => '二高 三郎', 'sex' => '男子', 'grade' => '1'],
    ],
];
$entries = [
    's1' => [
        '一高 補欠'  => ['kumi_chk' => 1, 'kumi_val' => '-55kg級', 'kumi_sub_val' => '補'],
        '一高 太郎'  => ['kumi_chk' => 1, 'kumi_val' => '-55kg級', 'kumi_sub_val' => 'シード', 'kumi_rank' => '1'],
        '一高 花子'  => ['kumi_chk' => 1, 'kumi_val' => '+66kg級'],
        '一高 形だけ' => ['kata_chk' => 1, 'kumi_val' => '-61kg級'],   // 組手に出ない
    ],
    's2' => [
        '二高 次郎' => ['kumi_chk' => 1, 'kumi_val' => '-55kg級', 'kumi_sub_val' => '正'],
        '二高 三郎' => ['kumi_chk' => 1, 'kumi_val' => '-99kg級'],         // 大会に無い階級
    ],
];

$classes = WeighSheet::collect($weights, $schools,
    fn($sid) => $members[$sid], fn($sid) => $entries[$sid]);

check(count($classes) === 5, '階級は大会の数だけ（エントリーが無い階級も入る）');
check(array_map(fn($c) => "{$c['sex']}{$c['class']}", $classes)
      === ['男子-55kg級', '男子-61kg級', '男子+76kg級', '女子-48kg級', '女子+66kg級'], '男子の軽い順 → 女子');
$m55 = array_map(fn($p) => $p['name'], $classes[0]['players']);
check($m55 === ['二高 次郎', '一高 太郎', '一高 補欠'],
      '学校番号順・補欠は学校の最後: ' . implode(' / ', $m55));
check($classes[0]['players'][1]['kubun'] === 'シード' && $classes[0]['players'][0]['kubun'] === '正'
      && $classes[0]['players'][2]['kubun'] === '補', '区分（正・シード・補）');
check($classes[1]['players'] === [], '組手にチェックの無い人は載らない');
check(count($classes[4]['players']) === 1 && $classes[4]['players'][0]['grade'] === '1', '女子 +66kg級 に 1 人・学年つき');
$all = array_merge(...array_map(fn($c) => $c['players'], $classes));
check(!in_array('二高 三郎', array_column($all, 'name'), true), '大会に無い階級の人は載らない');

// ---- Excel にする
$out = $argv[1] ?? (sys_get_temp_dir() . '/weigh_test_' . bin2hex(random_bytes(4)) . '.xlsx');
WeighSheet::book('令和8年度　新人大会', $classes)->save($out);
$zip = new ZipArchive();
check($zip->open($out) === true, 'xlsx が開ける');
$wb = (string)$zip->getFromName('xl/workbook.xml');
preg_match_all('/<sheet name="([^"]+)"/', $wb, $mm);
$names = array_map(fn($s) => html_entity_decode($s, ENT_QUOTES | ENT_XML1, 'UTF-8'), $mm[1]);
check($names === ['一覧', '男子 -55kg級', '女子 +66kg級'], 'シートは 一覧 と人のいる階級だけ: ' . implode(' / ', $names));
$sum = (string)$zip->getFromName('xl/worksheets/sheet1.xml');
check(str_contains($sum, 'COUNTIF(&apos;男子 -55kg級&apos;!F5:F7,&quot;○&quot;)'), '一覧の計量済みは階級のシートの F 列を数える');
check(str_contains($sum, 'SUM(C5:C9)'), '一覧の計');
check(substr_count($sum, '—') === 6, '人のいない階級は計量済み・残りが「—」');
$s2 = (string)$zip->getFromName('xl/worksheets/sheet2.xml');
check(str_contains($s2, '二高 次郎') && str_contains($s2, '一高 補欠'), '階級のシートに選手');
check(str_contains($s2, 'COUNTIF(F5:F7,&quot;○&quot;)&amp;&quot;名&quot;'), '階級のシートにも計量済みの数');
check(preg_match('/<c r="F5" s="6"\/>/', $s2) === 1 && preg_match('/<c r="H7" s="6"\/>/', $s2) === 1,
      '計量・体重・備考は書き込む欄（黄色）');
check(str_contains($wb, "_xlnm.Print_Titles"), '印刷の各ページに見出しを繰り返す');
$zip->close();
if (!isset($argv[1])) {
    unlink($out);
}

// ---- 階級の無い大会
$none = WeighSheet::collect(['男子' => [], '女子' => []], $schools,
    fn($sid) => $members[$sid], fn($sid) => $entries[$sid]);
check($none === [], '階級の無い大会は空');

echo $fails ? "\n{$fails} 件 NG\n" : "\n全部通過\n";
exit($fails ? 1 : 0);
