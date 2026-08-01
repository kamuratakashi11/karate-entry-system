<?php
/**
 * 手元での回帰テスト用CLI。check.php と同じデータで型埋めし、
 * 書いたセルの一覧を JSON で標準出力に出す（Python 側の検証が読む）。
 *
 *   php cli_test.php [出力先.xlsx]
 */
declare(strict_types=1);
error_reporting(E_ALL);

require __DIR__ . '/xlsx_fill.php';
require __DIR__ . '/shinjin_form.php';
$data = require __DIR__ . '/test_data.php';
$coords = json_decode((string)file_get_contents(__DIR__ . '/coords_shinjin.json'), true, 512, JSON_THROW_ON_ERROR);

$out = $argv[1] ?? (__DIR__ . '/out_test.xlsx');
$cells = ShinjinForm::cells($coords, $data['school'], $data['members']);
$written = XlsxFill::fill(__DIR__ . '/template_shinjin.xlsx', $out, $cells);

fwrite(STDERR, sprintf("書いたセル: %d 箇所 → %s\n", count($written), $out));
echo json_encode($written, JSON_UNESCAPED_UNICODE | JSON_PRETTY_PRINT), "\n";
