<?php
/**
 * 初回だけ使う設置スクリプト。
 *
 * さくらのファイルマネージャーは www より上の階層を開けないため、DBを置く
 * ~/entry_data/ を「ブラウザから」作る。PHP はホームに書き込めるので、
 * このファイルを www/entry/ に置いて開くだけで次の3つが済む:
 *
 *   1. ~/entry_data/ と ~/entry_data/uploads/ を作る（公開ディレクトリの外）
 *   2. www/entry/ に仮置きした entry.sqlite3 をそこへ移す
 *   3. 権限を本人だけに絞る（DB=600 / フォルダ=700）
 *
 * 何度実行しても壊れない。**済んだら setup.php は削除すること。**
 */
declare(strict_types=1);
error_reporting(E_ALL);
ini_set('display_errors', '1');
header('Content-Type: text/html; charset=utf-8');

$home     = dirname(__DIR__, 2);          // …/www/entry → ホーム
$dataDir  = $home . '/entry_data';
$uploads  = $dataDir . '/uploads';
$dbTarget = $dataDir . '/entry.sqlite3';
$dbHere   = __DIR__ . '/entry.sqlite3';   // 仮置き場所（www の中）

$steps = [];
$ok = true;
$step = function (string $label, bool $good, string $note = '') use (&$steps, &$ok) {
    $steps[] = [$label, $good, $note];
    if (!$good) {
        $ok = false;
    }
};

// 0. 場所の確認（想定と違うところに置かれていないか）
$step('ホームの場所', is_dir($home) && is_writable($home), $home);

// 1. フォルダを作る
if (!is_dir($dataDir)) {
    @mkdir($dataDir, 0700, true);
}
$step('entry_data フォルダ', is_dir($dataDir), $dataDir);
if (!is_dir($uploads)) {
    @mkdir($uploads, 0700, true);
}
$step('uploads フォルダ', is_dir($uploads));

// 2. DBを公開ディレクトリの外へ移す
if (is_file($dbHere)) {
    if (is_file($dbTarget)) {
        // 2回目以降: 既にあるものを上書き（受付直前のデータ入れ替えにも使う）
        @unlink($dbTarget);
    }
    $moved = @rename($dbHere, $dbTarget);
    if (!$moved) {                        // 別ファイルシステムなら copy で代替
        $moved = @copy($dbHere, $dbTarget) && @unlink($dbHere);
    }
    $step('entry.sqlite3 を entry_data へ移動', $moved && is_file($dbTarget));
} else {
    $step('entry.sqlite3 を entry_data へ移動', is_file($dbTarget),
        is_file($dbTarget) ? '移動済み（このフォルダには残っていません）'
                           : '★ entry.sqlite3 が見つかりません。先にこのフォルダへアップロードしてください');
}

// 3. 権限を絞る（PHPは本人の権限で動くので 600/700 で足りる）
if (is_file($dbTarget)) {
    @chmod($dbTarget, 0600);
    @chmod($dataDir, 0700);
    @chmod($uploads, 0700);
    $step('DBを読み書きできる', is_readable($dbTarget) && is_writable($dbTarget),
        'サイズ ' . number_format((int)filesize($dbTarget)) . ' バイト');
}

// 4. www の中に残骸が無いか（実名入りのDBが公開されたままになっていないか）
$step('www の中にDBが残っていない', !is_file($dbHere));
?>
<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>申込窓口 初回設置</title>
<style>
  body { font-family:"Yu Gothic Medium","Hiragino Kaku Gothic ProN",Meiryo,sans-serif;
         max-width:680px; margin:2rem auto; padding:0 1rem; line-height:1.7; color:#1a1f2b; }
  h1 { font-size:1.2rem; border-bottom:3px solid #1a1f2b; padding-bottom:.4rem; }
  .banner { padding:.8rem 1rem; border-radius:6px; font-weight:bold; margin:1rem 0; }
  .ok { background:#dff3e7; color:#135f13; }
  .ng { background:#fbeae8; color:#8f1414; }
  table { border-collapse:collapse; width:100%; }
  td, th { border:1px solid #ccc; padding:.3rem .7rem; text-align:left; font-size:.95rem; }
  .miss { color:#a8322a; font-weight:bold; }
  .note { color:#667; font-size:.9rem; }
  ol { padding-left:1.3rem; }
</style>
</head>
<body>
<h1>申込窓口 初回設置</h1>
<?php if ($ok): ?>
  <div class="banner ok">✅ できました — 次は <a href="check.php">check.php で配備チェック</a></div>
<?php else: ?>
  <div class="banner ng">❌ 途中で止まりました — 赤い行をそのまま伝えてください</div>
<?php endif; ?>
<table>
<?php foreach ($steps as [$label, $good, $note]): ?>
  <tr><th><?= htmlspecialchars($label) ?></th>
      <td><?= $good ? 'OK' : '<span class="miss">★問題★</span>' ?>
          <?= $note !== '' ? ' <span class="note">' . htmlspecialchars($note) . '</span>' : '' ?></td></tr>
<?php endforeach; ?>
</table>
<?php if ($ok): ?>
<p class="note"><strong>このあと必ず: setup.php を削除してください。</strong>
（受付開始前にデータを入れ直すときは、また置いて使えます）</p>
<?php endif; ?>
</body>
</html>
