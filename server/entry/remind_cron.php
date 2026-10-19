<?php
/**
 * 締め切りの確認メールを、送る日が来たものだけ送る（remind.php）。さくらの cron から毎日 1 回動かす。
 *
 *   さくらのコントロールパネル → スクリプト設定（CRON 設定）に、例えば毎朝 7 時:
 *     /usr/local/bin/php /home/<アカウント>/www/entry/remind_cron.php
 *
 * ブラウザから開かれても何もしない（コマンドから動かしたときだけ送る）。
 */
declare(strict_types=1);

if (PHP_SAPI !== 'cli') {
    http_response_code(404);
    exit;
}
date_default_timezone_set('Asia/Tokyo');
require_once __DIR__ . '/remind.php';

$res = remind_run(date('Y-m-d'));
foreach ($res['sent'] as $d) {
    echo date('Y-m-d H:i'), " 送った: {$d['kind']} {$d['name']}（締め切り {$d['deadline']}）{$d['result']}\n";
}
foreach ($res['errors'] as $e) {
    fwrite(STDERR, date('Y-m-d H:i') . " {$e}\n");
}
exit($res['errors'] ? 1 : 0);
