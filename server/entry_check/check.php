<?php
/**
 * さくらのサーバーで「申込書の型埋め」が動くかを確かめる診断ページ（④の技術検証）。
 *
 * 置き方: このフォルダ（check.php / xlsx_fill.php / shinjin_form.php / test_data.php /
 *         coords_shinjin.json / template_shinjin.xlsx）を ~/www/entry_check/ に置いて
 *         https://…/entry_check/check.php を開くだけ。
 * 後始末: 確認が済んだらフォルダごと削除してよい（実名は含まれていない）。
 */
declare(strict_types=1);
error_reporting(E_ALL);
ini_set('display_errors', '1');
header('Content-Type: text/html; charset=utf-8');

// --- 1. 環境 ---
$exts = [];
foreach (['zip', 'dom', 'libxml', 'mbstring'] as $e) {
    $exts[$e] = extension_loaded($e);
}
$dirWritable = is_writable(__DIR__);
$envOk = !in_array(false, $exts, true) && $dirWritable;

// --- 2. 必要ファイルの点呼（コピー失敗の原因はたいていここ） ---
$required = ['check.php', 'xlsx_fill.php', 'shinjin_form.php', 'test_data.php',
             'coords_shinjin.json', 'template_shinjin.xlsx'];
$files = [];
foreach ($required as $f) {
    $p = __DIR__ . '/' . $f;
    $files[$f] = is_file($p) ? (int)filesize($p) : null;
}
$filesOk = !in_array(null, $files, true);
$listing = [];
foreach (scandir(__DIR__) ?: [] as $f) {
    if ($f !== '.' && $f !== '..') {
        $listing[$f] = is_file(__DIR__ . '/' . $f) ? (int)filesize(__DIR__ . '/' . $f) : -1;
    }
}

// --- 3. 生成テスト ---
// 出力名は英数字にする（多バイトのファイル名で書けない環境の可能性を消す）。
// ダウンロード時の名前は <a download> 属性で日本語に戻す。
$outName = 'output_test.xlsx';
$dlName  = '申込書_確認用テスト.xlsx';

$result = null;
$error = null;
if ($envOk && $filesOk) {
    try {
        require __DIR__ . '/xlsx_fill.php';
        require __DIR__ . '/shinjin_form.php';
        $data = require __DIR__ . '/test_data.php';
        $coords = json_decode((string)file_get_contents(__DIR__ . '/coords_shinjin.json'), true, 512, JSON_THROW_ON_ERROR);

        $t0 = microtime(true);
        $cells = ShinjinForm::cells($coords, $data['school'], $data['members']);
        $written = XlsxFill::fill(__DIR__ . '/template_shinjin.xlsx', __DIR__ . '/' . $outName, $cells);
        $ms = (int)round((microtime(true) - $t0) * 1000);

        // 忠実性: テンプレートと比べて、変わった zip エントリを列挙する。
        // シートXML（値を入れた場所）だけが変わっているのが正しい状態。
        $crc = function (string $path): array {
            $z = new ZipArchive();
            $z->open($path);
            $map = [];
            for ($i = 0; $i < $z->numFiles; $i++) {
                $st = $z->statIndex($i);
                $map[$st['name']] = $st['crc'];
            }
            $z->close();
            return $map;
        };
        $a = $crc(__DIR__ . '/template_shinjin.xlsx');
        $b = $crc(__DIR__ . '/' . $outName);
        $changed = [];
        foreach ($b as $name => $c) {
            if (!array_key_exists($name, $a)) {
                $changed[] = "$name（追加）";
            } elseif ($a[$name] !== $c) {
                $changed[] = $name;
            }
        }
        foreach ($a as $name => $c) {
            if (!array_key_exists($name, $b)) {
                $changed[] = "$name（消失）";
            }
        }
        $result = [
            'cells'   => count($written),
            'ms'      => $ms,
            'peak'    => round(memory_get_peak_usage(true) / 1048576, 1),
            'size'    => filesize(__DIR__ . '/' . $outName),
            'changed' => $changed,
            'fidelity_ok' => ($changed === ['xl/worksheets/sheet1.xml']),
        ];
    } catch (Throwable $e) {
        $error = $e;
    }
}
$allOk = $envOk && $filesOk && $result && $result['fidelity_ok'] && !$error;
?>
<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>申込書の型埋め 動作確認</title>
<style>
  body { font-family: "Hiragino Kaku Gothic ProN", "Yu Gothic", Meiryo, sans-serif;
         max-width: 720px; margin: 2rem auto; padding: 0 1rem; line-height: 1.7; color: #223; }
  h1 { font-size: 1.3rem; border-bottom: 3px solid #223; padding-bottom: .4rem; }
  .banner { padding: .8rem 1rem; border-radius: 8px; font-weight: bold; font-size: 1.1rem; margin: 1rem 0; }
  .ok  { background: #e6f6e6; color: #135f13; border: 1px solid #a5d6a5; }
  .ng  { background: #fdeaea; color: #8f1414; border: 1px solid #efb3b3; }
  table { border-collapse: collapse; margin: .5rem 0 1rem; }
  td, th { border: 1px solid #ccc; padding: .3rem .8rem; text-align: left; }
  .miss { color: #b00; font-weight: bold; }
  .dl { display: inline-block; padding: .7rem 1.2rem; background: #235; color: #fff;
        border-radius: 8px; text-decoration: none; font-weight: bold; }
  ul { padding-left: 1.3rem; }
  .note { color: #667; font-size: .92rem; }
</style>
</head>
<body>
<h1>申込書の型埋め 動作確認（新人戦テンプレート）</h1>

<?php if ($allOk): ?>
  <div class="banner ok">✅ 成功 — この方式でいけます</div>
<?php else: ?>
  <div class="banner ng">❌ 問題があります — 下の表とエラーをそのまま伝えてください</div>
<?php endif; ?>

<h2>1. 環境</h2>
<table>
  <tr><th>PHP</th><td><?= htmlspecialchars(PHP_VERSION) ?></td></tr>
  <?php foreach ($exts as $name => $ok): ?>
  <tr><th><?= htmlspecialchars($name) ?></th><td><?= $ok ? 'あり' : '★無い★' ?></td></tr>
  <?php endforeach; ?>
  <tr><th>フォルダ書き込み</th><td><?= $dirWritable ? '可' : '★不可★' ?></td></tr>
  <tr><th>memory_limit</th><td><?= htmlspecialchars((string)ini_get('memory_limit')) ?></td></tr>
</table>

<h2>2. 必要ファイル</h2>
<table>
  <?php foreach ($files as $name => $size): ?>
  <tr><th><?= htmlspecialchars($name) ?></th>
      <td><?= $size !== null ? number_format($size) . ' バイト' : '<span class="miss">★無い★ アップロードし直してください</span>' ?></td></tr>
  <?php endforeach; ?>
</table>
<?php $others = array_diff_key($listing, $files, [$outName => 0]); if ($others): ?>
<p class="note">このフォルダにある他のファイル（名前違いのアップロードが混ざっていないか）:
  <?= htmlspecialchars(implode(' / ', array_keys($others))) ?></p>
<?php endif; ?>

<h2>3. 生成テスト</h2>
<?php if ($error): ?>
  <div class="banner ng">エラー: <?= htmlspecialchars($error->getMessage()) ?></div>
  <pre class="note"><?= htmlspecialchars((string)$error) ?></pre>
<?php elseif ($result): ?>
  <table>
    <tr><th>書いたセル</th><td><?= $result['cells'] ?> 箇所</td></tr>
    <tr><th>処理時間</th><td><?= $result['ms'] ?> ミリ秒</td></tr>
    <tr><th>メモリ</th><td><?= $result['peak'] ?> MB</td></tr>
    <tr><th>出力サイズ</th><td><?= number_format((int)$result['size']) ?> バイト</td></tr>
    <tr><th>変わった中身</th><td><?= htmlspecialchars(implode(' / ', $result['changed'])) ?>
        <?= $result['fidelity_ok'] ? '（＝セル値だけ。正常）' : '（★想定外。体裁が変わっている可能性★）' ?></td></tr>
  </table>
  <p><a class="dl" href="<?= rawurlencode($outName) ?>" download="<?= htmlspecialchars($dlName) ?>">📄 できた申込書をダウンロード</a></p>
  <p>Excel で開いて、見慣れた申込書と見比べてください:</p>
  <ul>
    <li>罫線・結合セル・列幅が崩れていないか</li>
    <li>学校名「確認用テスト高校」・校長・顧問4名（引率の○×）が正しい欄にあるか</li>
    <li>部員24名: 男女の各種目の列に ○・補・シ2 などが入っているか</li>
    <li>23人目・24人目が2ブロック目（用紙の後半）に入っているか</li>
    <li>印刷プレビューが今までの申込書と同じか</li>
  </ul>
<?php else: ?>
  <p>環境またはファイルに不足があるため、生成テストは実行していません。</p>
<?php endif; ?>

<p class="note">確認が済んだら、この entry_check フォルダは丸ごと削除して構いません（実名は含まれていません）。</p>
</body>
</html>
