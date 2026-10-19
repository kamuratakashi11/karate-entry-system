<?php
/**
 * 締め切りの確認メール（server/entry/remind.php）の回帰テスト。架空のデータだけ。
 * メールは送らず、一時フォルダに .eml を書く（ENTRY_MAIL_DIR）。
 *
 *   php server/entry_check/remind_test.php
 */
declare(strict_types=1);
error_reporting(E_ALL);

$tmp = sys_get_temp_dir() . '/remind_test_' . bin2hex(random_bytes(4));
mkdir($tmp);
mkdir("{$tmp}/mail");
putenv("ENTRY_DB={$tmp}/entry.sqlite3");
putenv("ENTRY_MAIL_DIR={$tmp}/mail");

$pdo = new PDO("sqlite:{$tmp}/entry.sqlite3");
$pdo->exec("CREATE TABLE schools (school_id TEXT PRIMARY KEY, base_name TEXT, short_name TEXT, school_no INTEGER,
  principal TEXT, advisors_json TEXT, password_hash TEXT, raw_json TEXT);
CREATE TABLE members (school_id TEXT, name TEXT, sex TEXT, grade TEXT, dob TEXT, jkf_no TEXT, display_order TEXT, active TEXT);
CREATE TABLE entries (tournament_id TEXT, entry_key TEXT, data_json TEXT, PRIMARY KEY (tournament_id, entry_key));
CREATE TABLE config (key TEXT PRIMARY KEY, value_json TEXT);");
$pdo->exec("INSERT INTO schools VALUES ('a','見本一',null,1,'','[]','','{}'), ('b','見本二高校',null,2,'','[]','','{}'), ('c','見本三',null,3,'','[]','','{}')");
$today = '2026-10-19';
$put = $pdo->prepare('INSERT INTO config VALUES (?,?)');
$put->execute(['tournaments', json_encode([
    't1' => ['name' => '大会イ', 'active' => true, 'deadline' => '2026-10-25'],    // 送る日（10-11）を過ぎている＝取りこぼしを送る
    't2' => ['name' => '大会ロ', 'active' => false, 'deadline' => '2026-10-30'],   // 受付前は送らない
    't3' => ['name' => '大会ハ', 'active' => false, 'deadline' => '2026-12-01'],
], JSON_UNESCAPED_UNICODE)]);
$put->execute(['shinsa_rounds', json_encode([
    'r1' => ['year' => 8, 'no' => 2, 'date' => '2026-11-23', 'deadline' => '2026-11-02', 'open' => true, 'lc' => true],  // ちょうど 14 日前
    'r2' => ['year' => 8, 'no' => 3, 'date' => '2027-02-14', 'deadline' => '2027-01-25', 'open' => false, 'lc' => true],
], JSON_UNESCAPED_UNICODE)]);
$pdo = null;

require __DIR__ . '/../entry/remind.php';

$fails = 0;
function check(string $what, bool $ok, string $detail = ''): void
{
    global $fails;
    echo ($ok ? 'OK  ' : 'NG  ') . $what . ($ok || $detail === '' ? '' : " … {$detail}") . "\n";
    if (!$ok) {
        $fails++;
    }
}
function setcfg(string $k, $v): void
{
    db()->prepare('INSERT OR REPLACE INTO config VALUES (?,?)')->execute([$k, json_encode($v, JSON_UNESCAPED_UNICODE)]);
}
$mails = fn() => glob("{$tmp}/mail/*.eml") ?: [];

$plan = array_column(remind_plan($today), null, 'name');
check('送る日は締め切りの 14 日前', $plan['第2回 級位審査会']['send_on'] === '2026-10-19');
check('ちょうど 14 日前は送る', $plan['第2回 級位審査会']['due']);
check('送る日を過ぎていても、締め切り前なら送る（cron が止まっていた日の取りこぼし）', $plan['大会イ']['due']);
check('受付中でない大会には送らない', !$plan['大会ロ']['due']);
check('まだ先の締め切りは送らない', !$plan['大会ハ']['due']);
check('過ぎた締め切りは予定に出ない', !isset(array_column(remind_plan('2026-10-26'), null, 'name')['大会イ']));

$r = remind_run($today);
check('送り主が無いと送らない', !$r['sent'] && (bool)preg_grep('/送り主/u', $r['errors']));

setcfg('mail_settings', ['from' => 'info@example.com', 'reply_to' => 'kanri@example.com']);
$r = remind_run($today);
check('アドレスが 1 つも無いと送らず、送った記録も付けない', !$r['sent'] && !$mails() && (array)config_get('mail_sent', []) === []);

[$ok, $bad] = split_emails("x@example.com, y@example.com　z＠example.com;x@example.com");
check('アドレスを区切って確かめる（重複は 1 つに・おかしいものは分ける）', $ok === ['x@example.com', 'y@example.com'] && $bad === ['z＠example.com']);

setcfg('school_emails', ['a' => ['a1@example.com', 'a2@example.com'], 'b' => ['b@example.com']]);
$r = remind_run($today);
check('受付中で送る日の来た 2 件を送る', count($r['sent']) === 2, json_encode($r, JSON_UNESCAPED_UNICODE));
check('アドレスの数だけ送る（2 件 × 3 通）', count($mails()) === 6, (string)count($mails()));
check('アドレスの無い学校は記録に残す', str_contains($r['sent'][0]['result'], 'アドレス無し1校'));
$r2 = remind_run($today);
check('同じ締め切りには 2 回送らない', !$r2['sent'] && count($mails()) === 6);

// 文面
$m = '';
foreach ($mails() as $f) {
    $raw = (string)file_get_contents($f);
    if (str_contains($raw, 'b@example.com')) {
        [, $body] = explode("\r\n\r\n", $raw, 2);
        $dec = base64_decode($body);
        if (str_contains($dec, '級位審査会')) {
            $m = $dec;
            // 長い件名は折り返されている（「\r\n 」で続く）ので、つないでから読む
            preg_match('/\r\nSubject: (.*?)\r\n(?=[^ \t])/s', "\r\n" . $raw, $mm);
            $subj = mb_decode_mimeheader(str_replace(["\r\n ", "\r\n\t"], '', $mm[1] ?? ''));
        }
    }
}
check('宛名は学校名（「高校」を重ねない）', str_contains($m, '見本二高等学校 空手道部 顧問の先生'));
check('締め切りと残り日数', str_contains($m, '2026年11月2日（月）です（あと14日）'));
check('審査期日も書く', str_contains($m, '審査期日 11月23日（月）'));
check('級位審査はいまの申込の人数を添える', str_contains($m, '申込はまだ保存されていません'));
check('件名', ($subj ?? '') === '【締め切りの確認】第2回 級位審査会の申込は11月2日（月）まで', $subj ?? '');

echo $fails ? "\n{$fails} 件 NG\n" : "\n全部通過\n";
exit($fails ? 1 : 0);
