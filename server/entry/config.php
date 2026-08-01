<?php
/**
 * /entry/ の設定。秘密情報は持たない（DBの場所と定数だけ）。
 *
 * DB はさくらでは ~/entry_data/entry.sqlite3（公開ディレクトリの外）。
 * ~/www/entry/ に置かれたとき __DIR__ の2つ上がホームになる。
 * 開発時は環境変数 ENTRY_DB で差し替える。
 */
declare(strict_types=1);

// 提出日時・締切の計算は日本時間で行う（さくらの既定は UTC のことがあり、
// 何も指定しないと「提出したのに9時間前の表示」になる）
date_default_timezone_set('Asia/Tokyo');

$env = getenv('ENTRY_DB');
define('ENTRY_DB_PATH', ($env !== false && $env !== '') ? $env : dirname(__DIR__, 2) . '/entry_data/entry.sqlite3');
define('ENTRY_UPLOAD_DIR', dirname(ENTRY_DB_PATH) . '/uploads');
define('ENTRY_LIB_DIR', __DIR__ . '/lib');
