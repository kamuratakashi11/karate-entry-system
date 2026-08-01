<?php
/**
 * /entry/ の配備セルフチェック。さくらに置いたら最初に1回開く。
 * DBの場所・読み書き・PHP拡張・セッションを点検して緑/赤で答える。
 * 秘密は表示しない（件数と大会名だけ）。**確認できたらこのファイルは削除する。**
 */
declare(strict_types=1);
error_reporting(E_ALL);
ini_set('display_errors', '1');
header('Content-Type: text/html; charset=utf-8');

require_once __DIR__ . '/config.php';

$rows = [];
$ok = true;
$add = function (string $label, bool $good, string $note = '') use (&$rows, &$ok) {
    $rows[] = [$label, $good, $note];
    if (!$good) {
        $ok = false;
    }
};

$add('PHP ' . PHP_VERSION, version_compare(PHP_VERSION, '8.0', '>='));
foreach (['pdo_sqlite', 'zip', 'dom', 'mbstring', 'session'] as $e) {
    $add("拡張 {$e}", extension_loaded($e));
}

// セッション（ログインが動く前提条件）
try {
    session_name('KARATEENTRYCHK');
    @session_start();
    $_SESSION['t'] = '1';
    $add('セッションの書き込み', session_status() === PHP_SESSION_ACTIVE && ($_SESSION['t'] ?? '') === '1',
        ini_get('session.save_path') ?: '(既定の場所)');
    session_destroy();
} catch (Throwable $e) {
    $add('セッションの書き込み', false, $e->getMessage());
}

// アプリのファイル
foreach (['index.html', 'style.css', 'app.js', 'api.php', 'auth.php', 'db.php',
          'config.php', 'uploads.php', 'admin.html', 'admin_api.php',
          'lib/xlsx_fill.php', 'lib/entry_form.php', 'lib/school_list_form.php',
          'lib/coords_shinjin.json', 'lib/coords_standard.json',
          'lib/template_shinjin.xlsx', 'lib/template_kantou.xlsx',
          'lib/template_interhigh.xlsx', 'lib/template_schools_shinjin.xlsx'] as $f) {
    $add("ファイル {$f}", is_file(__DIR__ . '/' . $f));
}

// 不足があったとき用: 実際に何が置かれているかを見せる（名前違い・入れ場所違いの切り分け）
$listing = null;
if (!$ok) {
    $ls = function (string $dir): string {
        if (!is_dir($dir)) {
            return '（フォルダがありません）';
        }
        $n = array_values(array_diff(scandir($dir) ?: [], ['.', '..']));
        return $n ? implode('  ', $n) : '（空）';
    };
    $listing = ['entry の中' => $ls(__DIR__), 'entry/lib の中' => $ls(__DIR__ . '/lib')];
}

// データベース
$db = ENTRY_DB_PATH;
$add('DBの場所が公開フォルダの外', !str_contains(str_replace('\\', '/', $db), '/www/'), $db);
$add('DBファイルがある', is_file($db));
$add('DBを読める', is_file($db) && is_readable($db));
$add('DBを書ける（保存に必要）', is_file($db) && is_writable($db));
$add('DBのフォルダを書ける（SQLiteの作業ファイル用）', is_dir(dirname($db)) && is_writable(dirname($db)));

// 押印済み申込書の置き場（学校からの提出先。DBと同じくwwwの外）
if (is_file(__DIR__ . '/uploads.php')) {
    require_once __DIR__ . '/uploads.php';
    if (!is_dir(ENTRY_UPLOAD_DIR)) {
        @mkdir(ENTRY_UPLOAD_DIR, 0700, true);
    }
    $add('提出フォルダがある', is_dir(ENTRY_UPLOAD_DIR), ENTRY_UPLOAD_DIR);
    $add('提出フォルダを書ける', is_dir(ENTRY_UPLOAD_DIR) && is_writable(ENTRY_UPLOAD_DIR));
    // 押印済みの申込書は写真1枚で数MBになる。上限が小さいと先生方が提出できない
    $add('アップロードの上限（3MB以上）', upload_limit_bytes() >= 3 * 1024 * 1024,
        upload_limit_text() . '｜php.ini: upload_max_filesize=' . ini_get('upload_max_filesize')
        . ' post_max_size=' . ini_get('post_max_size')
        . '（小さいときは entry フォルダに php.ini を置いて上げる）');
    $add('提出済みの申込書', true, count(upload_scan()) . '件');
}

$counts = null;
$active = null;
if (is_file($db) && extension_loaded('pdo_sqlite')) {
    try {
        require_once __DIR__ . '/db.php';
        $counts = [];
        foreach (['schools', 'members', 'entries', 'config'] as $t) {
            $counts[$t] = (int)db()->query("SELECT COUNT(*) FROM {$t}")->fetchColumn();
        }
        $add('表がそろっている', min($counts) >= 0,
            "学校{$counts['schools']}・名簿{$counts['members']}・エントリー{$counts['entries']}・設定{$counts['config']}");
        $active = active_tournament();
        $add('受付中の大会', $active !== null,
            $active ? ($active['name'] . '（締切 ' . ($active['deadline'] ?? '未設定') . '）') : '無し');
    } catch (Throwable $e) {
        $add('DBの読み取り', false, $e->getMessage());
    }
}
?>
<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>申込窓口 配備チェック</title>
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
</style>
</head>
<body>
<h1>申込窓口（/entry/）配備チェック</h1>
<?php if ($ok): ?>
  <div class="banner ok">✅ 配備OK — <a href="./">ここから申込画面を開けます</a></div>
<?php else: ?>
  <div class="banner ng">❌ 問題があります — 赤い行をそのまま伝えてください</div>
<?php endif; ?>
<table>
<?php foreach ($rows as [$label, $good, $note]): ?>
  <tr><th><?= htmlspecialchars($label) ?></th>
      <td><?= $good ? 'OK' : '<span class="miss">★問題★</span>' ?>
          <?= $note !== '' ? '<span class="note">' . htmlspecialchars($note) . '</span>' : '' ?></td></tr>
<?php endforeach; ?>
</table>
<?php if ($listing): ?>
<h2 style="font-size:1rem">いま置かれているもの</h2>
<table>
<?php foreach ($listing as $where => $names): ?>
  <tr><th><?= htmlspecialchars($where) ?></th><td class="note"><?= htmlspecialchars($names) ?></td></tr>
<?php endforeach; ?>
</table>
<?php endif; ?>
<p class="note">このページはサーバーの中身を映すので、確認できたら <strong>check.php を削除</strong>してください（アプリの動作には不要です）。</p>
</body>
</html>
