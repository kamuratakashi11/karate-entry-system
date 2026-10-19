<?php
/**
 * 級位審査の申込（server/entry/shinsa.php）の回帰テスト。架空のデータだけを使う。
 * 一時フォルダに DB を作るので、本物の DB には触れない。
 *
 *   php server/entry_check/shinsa_test.php
 *
 * 終了コード 0 ＝ 全部通過。落ちたものは「NG」と理由を出す。
 */
declare(strict_types=1);
error_reporting(E_ALL);

$tmp = sys_get_temp_dir() . '/shinsa_test_' . bin2hex(random_bytes(4));
mkdir($tmp);
$dbPath = "{$tmp}/entry.sqlite3";
putenv("ENTRY_DB={$dbPath}");

$pdo = new PDO('sqlite:' . $dbPath);
$pdo->exec("CREATE TABLE schools (school_id TEXT PRIMARY KEY, base_name TEXT, short_name TEXT, school_no INTEGER,
  principal TEXT, advisors_json TEXT, password_hash TEXT, raw_json TEXT);
CREATE TABLE members (school_id TEXT, name TEXT, sex TEXT, grade TEXT, dob TEXT, jkf_no TEXT, display_order TEXT, active TEXT);
CREATE TABLE entries (tournament_id TEXT, entry_key TEXT, data_json TEXT, PRIMARY KEY (tournament_id, entry_key));
CREATE TABLE config (key TEXT PRIMARY KEY, value_json TEXT);");
$pdo->exec("INSERT INTO schools VALUES ('s1','見本高校','見本',1,'校長 太郎','[{\"name\":\"顧問 一郎\"}]','','{}')");
$ins = $pdo->prepare('INSERT INTO members VALUES (?,?,?,?,?,?,?,?)');
for ($i = 1; $i <= 23; $i++) {
    $ins->execute(['s1', "部員{$i}", $i % 2 ? '男子' : '女子', (string)(1 + $i % 3), '', $i === 23 ? '' : "9{$i}", '', 'True']);
}
$ins->execute(['s1', '退部者', '男子', '2', '', '999', '', 'False']);
$today = date('Y-m-d');
$future = date('Y-m-d', strtotime('+10 days'));
$past = date('Y-m-d', strtotime('-1 day'));
$rounds = [
    'r1' => ['year' => 8, 'no' => 1, 'date' => date('Y-m-d', strtotime('+30 days')), 'deadline' => $future, 'open' => true, 'lc' => true],
    'r2' => ['year' => 8, 'no' => 2, 'date' => date('Y-m-d', strtotime('+90 days')), 'deadline' => date('Y-m-d', strtotime('+60 days')), 'open' => false, 'lc' => false],
];
$pdo->prepare('INSERT INTO config VALUES (?,?)')->execute(['shinsa_rounds', json_encode($rounds, JSON_UNESCAPED_UNICODE)]);
$pdo->prepare('INSERT INTO config VALUES (?,?)')->execute(['tournaments', json_encode(
    ['a' => ['name' => '大会A', 'deadline' => $past], 'b' => ['name' => '大会B', 'deadline' => $future, 'active' => true]],
    JSON_UNESCAPED_UNICODE)]);
$pdo = null;

require __DIR__ . '/../entry/shinsa.php';
require __DIR__ . '/../entry/lib/xlsx_fill.php';

$fails = 0;
function check(string $what, bool $ok, string $detail = ''): void
{
    global $fails;
    echo ($ok ? 'OK  ' : 'NG  ') . $what . ($ok || $detail === '' ? '' : " … {$detail}") . "\n";
    if (!$ok) {
        $fails++;
    }
}
function row(string $name, string $kubun = 'jokyu', string $kyu = '3級', string $kata = 'ジオン'): array
{
    return ['name' => $name, 'kubun' => $kubun, 'kyu' => $kyu, 'kata' => $kata, 'years' => '1', 'months' => '4'];
}

$school = school_by_id('s1');
$r = shinsa_open_round();
check('学校に見せる回は open の回', ($r['id'] ?? '') === 'r1');
check('締め切り前は直せる', shinsa_can_edit($r));
check('締め切りの翌日からは直せない', !shinsa_can_edit(['open' => true, 'deadline' => $past]));
check('締め切りの当日は直せる', shinsa_can_edit(['open' => true, 'deadline' => $today]));
check('見せていない回は直せない', !shinsa_can_edit(['open' => false, 'deadline' => $future]));

// --- 保存の決まり ---
[$e] = shinsa_save($r, $school, [row('いない人')]);
check('名簿に無い人は断る', (bool)preg_grep('/部員名簿にありません/u', $e), implode(' / ', $e));
[$e] = shinsa_save($r, $school, [row('退部者')]);
check('在籍でない部員は断る', (bool)preg_grep('/部員名簿にありません/u', $e), implode(' / ', $e));
[$e] = shinsa_save($r, $school, [row('部員23')]);
check('JKF 番号の無い部員は断る（会員登録が無いと受審できない）', (bool)preg_grep('/JKF/u', $e), implode(' / ', $e));
[$e] = shinsa_save($r, $school, [row('部員23', 'koshu', '無級', 'ジオン')]);
check('形講習会のみは JKF 番号が無くても申し込める', !$e, implode(' / ', $e));
shinsa_save($r, $school, []);
[$e] = shinsa_save($r, $school, [row('部員1'), row('部員1')]);
check('同じ人を 2 回は断る', (bool)preg_grep('/2 回/u', $e), implode(' / ', $e));
[$e] = shinsa_save($r, $school, [row('部員1', 'jokyu', '初段')]);
check('上級に段は選べない', (bool)preg_grep('/公認級/u', $e), implode(' / ', $e));
[$e] = shinsa_save($r, $school, [row('部員1', 'jokyu', '3級', '知らない形')]);
check('一覧に無い形は断る', (bool)preg_grep('/形/u', $e), implode(' / ', $e));
[$e] = shinsa_save($r, $school, [['name' => '部員1', 'kubun' => 'jokyu', 'kyu' => '3級', 'kata' => 'ジオン', 'years' => '', 'months' => '4']]);
check('修業期間が空なら断る', (bool)preg_grep('/修業期間/u', $e), implode(' / ', $e));
[$e] = shinsa_save($rounds['r2'] + ['id' => 'r2'], $school, [row('部員1', 'lc', '初段', 'エンピ')]);
check('レベルチェックの無い回では選べない', (bool)preg_grep('/レベルチェック/u', $e), implode(' / ', $e));
[$e] = shinsa_save($rounds['r2'] + ['id' => 'r2'], $school, [row('部員1', 'koshu', '3級', 'エンピ')]);
check('形講習会の無い回では形講習会のみも選べない', (bool)preg_grep('/形講習会/u', $e), implode(' / ', $e));
check('断ったときは何も書かない', shinsa_apps_of('r1', 's1') === []);

// --- 21 人（申込書 2 枚）。最後の 1 人はレベルチェック ---
$rows = [];
for ($i = 1; $i <= 20; $i++) {
    $rows[] = row("部員{$i}", $i % 2 ? 'jokyu' : 'shokyu', $i % 2 ? '3級' : '無級', $i % 2 ? 'ジオン' : '松濤館平安初段');
}
$rows[] = row('部員21', 'lc', '初段', 'エンピ');
[$e, $n] = shinsa_save($r, $school, $rows);
check('21 人を保存できる', !$e && $n === 21, implode(' / ', $e));
$apps = shinsa_apps_of('r1', 's1');
check('並び順は送った順', array_column($apps, 'name') === array_column($rows, 'name'));
check('性別・学年・会員番号は名簿から写す', $apps[0]['sex'] === '男子' && $apps[0]['grade'] === '2' && $apps[0]['jkf_no'] === '91');
check('申込書は 2 枚', shinsa_form_pages($apps) === 2);

$c1 = shinsa_form_cells($r, $school, $apps, 1);
$c2 = shinsa_form_cells($r, $school, $apps, 2);
check('1 枚目の先頭', $c1['A17'] === 1 && $c1['H17'] === '部員1' && $c1['N17'] === '男子上級' && $c1['B17'] === '91');
check('1 枚目は 20 人まで', isset($c1['H36']) && !isset($c1['H37']));
check('2 枚目は 21 番から', $c2['A17'] === 21 && $c2['H17'] === '部員21' && !isset($c2['H18']));
check('レベルチェックは受審級を空けて〇', !isset($c2['N17']) && $c2['Z17'] === '〇' && $c2['P17'] === '初段');
check('学校名は「高校」を除く（様式に刷ってある）', $c1['G10'] === '見本');
[$ry] = shinsa_reiwa($r['date']);
check('審査期日を令和で書く', $c1['V6'] === $ry && $c1['E6'] === 8 && $c1['J6'] === 1);
check('修業期間は数で書く', $c1['V17'] === 1 && $c1['X17'] === 4);

// 様式に実際に書けるか（受審申込書 シート）
$out = "{$tmp}/form.xlsx";
$written = XlsxFill::fill(__DIR__ . '/../entry/lib/template_shinsa.xlsx', $out, $c1, '受審申込書 ');
check('様式の「受審申込書 」に書ける', ($written['H17'] ?? '') === '部員1' && is_file($out));
$zip = new ZipArchive();
$zip->open($out);
$wb = (string)$zip->getFromName('xl/workbook.xml');
check('外部リンクが無い（開くたびに更新を聞かれない）', !str_contains($wb, 'externalReference') && $zip->locateName('xl/externalLinks/externalLink1.xml') === false);
check('開いたときに人数の式を計算し直す', str_contains($wb, 'fullCalcOnLoad="1"'));
$zip->close();

// --- 直しても受審番号は残る（同じ区分のとき） ---
$pdo = shinsa_db();
$d = json_decode((string)$pdo->query("SELECT data_json FROM shinsa_apps WHERE name='部員1'")->fetchColumn(), true);
$d['no'] = '7';
$pdo->prepare("UPDATE shinsa_apps SET data_json=? WHERE name='部員1'")->execute([json_encode($d, JSON_UNESCAPED_UNICODE)]);
$rows[0]['kata'] = 'エンピ';
$rows[2]['kubun'] = 'jokyu';
$rows[2]['kyu'] = '4級';
shinsa_save($r, $school, $rows);
$by = array_column(shinsa_apps_of('r1', 's1'), null, 'name');
check('形を直しても受審番号は残る', ($by['部員1']['no'] ?? '') === '7' && $by['部員1']['kata'] === 'エンピ');

// --- これからの締め切り ---
$dl = upcoming_deadlines();
$names = array_column($dl, 'name');
check('過ぎた締め切りは出さない', !in_array('大会A', $names, true));
check('大会と審査を締め切り順に並べる', $names === ['大会B', '第1回 級位審査会', '第2回 級位審査会'], implode(' / ', $names));

// =====================================================================
// 段階 2: 審査の管理画面
// =====================================================================
require __DIR__ . '/../entry/shinsa_day.php';

// --- 2 校目を足して、受審番号の並びを見る ---
$pdo = shinsa_db();
$pdo->exec("INSERT INTO schools VALUES ('s0','見本第零',0,0,'','[]','','{}')");   // 学校番号 0＝s1 より前
$ins = $pdo->prepare('INSERT INTO members VALUES (?,?,?,?,?,?,?,?)');
foreach (['零A' => '男子', '零B' => '男子', '零C' => '女子', '零D' => '女子'] as $n => $sx) {
    $ins->execute(['s0', $n, $sx, '1', '', "8{$n}", '', 'True']);
}
[$e0] = shinsa_save($r, school_by_id('s0'), [row('零A'), row('零B'), row('零C', 'shokyu', '無級', 'ピンアン二段'),
    row('零D', 'koshu', '無級', 'ピンアン二段')]);
check('形講習会のみを申し込める', !$e0, implode(' / ', $e0));
$counts = shinsa_assign_numbers('r1');
$ordered = shinsa_ordered(shinsa_round_apps('r1'));
$names = fn(string $k) => array_column($ordered[$k], 'name');
// s1 の上級男子は 部員1,3,5,…（奇数＝男子・jokyu）。s0 は 零A, 零B
check('上級は各校の 1 人目 → 2 人目の順（学校番号順）',
    array_slice($names('jokyu_M'), 0, 4) === ['零A', '部員1', '零B', '部員3'], implode(',', $names('jokyu_M')));
check('初級は学校番号順・学校の中は申込書の順',
    array_slice($names('shokyu_F'), 0, 3) === ['零C', '部員2', '部員4'], implode(',', $names('shokyu_F')));
$by = array_column(shinsa_round_apps('r1'), null, 'name');
check('組ごとに 1 番から', $by['零A']['no'] === 1 && $by['零C']['no'] === 1 && $by['部員21']['no'] === 1);
check('受審番号は形講習会のみの人以外の全員に', count(array_filter($by, fn($a) => ($a['no'] ?? '') !== '')) === count($by) - 1);
check('形講習会のみの人には受審番号を振らない', ($by['零D']['no'] ?? '') === '');
$c0 = shinsa_form_cells($r, school_by_id('s0'), shinsa_apps_of('r1', 's0'), 1);
check('形講習会のみは申込書で受審級を空けて〇', !isset($c0['N20']) && $c0['Z20'] === '〇');
$st0 = array_column(shinsa_status('r1'), null, 'school_id')['s0'];
check('申込状況で形講習会のみを別に数える', ($st0['koshu'] ?? 0) === 1 && $st0['to_number'] === 3 && $st0['numbered'] === 3);
check('形講習会のみは審査用紙が無い', shinsa_kubun()['koshu']['sheet'] === null && shinsa_kubun()['lc']['sheet'] === 'lc');
check('振った日時を回に残す', (shinsa_rounds()['r1']['numbered_at'] ?? '') !== '');

// --- 回の保存 ---
$rows = array_values(array_map(fn($x) => $x, shinsa_rounds()));
$e = shinsa_save_rounds(array_merge($rows, [['year' => 8, 'no' => 3, 'date' => '2027-02-14', 'deadline' => '2027-03-01']]));
check('締め切りが審査期日より後は断る', (bool)preg_grep('/締め切りが審査期日より後/u', $e), implode(' / ', $e));
$e = shinsa_save_rounds(array_merge($rows, [['year' => 8, 'no' => 1, 'date' => '2027-02-14', 'deadline' => '2027-01-25']]));
check('同じ年度・同じ回は断る', (bool)preg_grep('/2 つあります/u', $e), implode(' / ', $e));
$e = shinsa_save_rounds([$rows[1]]);
check('申込のある回は消せない', (bool)preg_grep('/消せません/u', $e), implode(' / ', $e));
$e = shinsa_save_rounds(array_merge($rows, [['year' => 9, 'no' => 1, 'date' => '2027-06-20', 'deadline' => '2027-05-31', 'lc' => true]]));
$after = shinsa_rounds();
check('回を足せる・前の回の受審番号の記録は残る', !$e && count($after) === 3 && ($after['r1']['numbered_at'] ?? '') !== '', implode(' / ', $e));

// --- 領収書 ---
$e = shinsa_save_results(shinsa_rounds()['r1'], $school, ['部員1' => 'pass']);
check('領収書を出す前は合否を受け付けない', (bool)$e);
check('領収書の日付が無いと出せない', (bool)shinsa_set_receipt('r1', true, ''));
shinsa_set_receipt('r1', true, '2026-11-23');
$e = shinsa_save_results(shinsa_rounds()['r1'], $school, ['部員1' => 'pass', '部員2' => 'fail', '部員3' => 'absent', '部員21' => 'pass']);
$by = array_column(shinsa_apps_of('r1', 's1'), null, 'name');
check('合否を残す', !$e && $by['部員1']['result'] === 'pass' && $by['部員2']['result'] === 'fail' && $by['部員3']['result'] === 'absent');
check('レベルチェック（無料）には合否を付けない', ($by['部員21']['result'] ?? '') === '');
check('知らない値は断る', (bool)shinsa_save_results(shinsa_rounds()['r1'], $school, ['部員1' => '○']));
$st = shinsa_state_for($school);
check('領収書の日付を令和で渡す', $st['receipt_reiwa'] === [8, 11, 23]);

// --- 当日用の Excel ---
check('合格級の式（上級）', shinsa_pass_formula('jokyu', 'I4')
    === 'IF(I4="","",IF(I4>=42,"1級",IF(I4>=39,"2級",IF(I4>=36,"3級",IF(I4>=30,"4級","不合格")))))');
check('合格級の式（初級は 22.5 点など小数も）', str_contains(shinsa_pass_formula('shokyu', 'J4'), 'IF(J4>=22.5,"6級"'));
$day = shinsa_day_book(shinsa_rounds()['r1'], "{$tmp}/day.xlsx");
$zip = new ZipArchive();
$zip->open($day);
$wb = (string)$zip->getFromName('xl/workbook.xml');
$sheetNames = [];
preg_match_all('/<sheet name="([^"]+)"/u', $wb, $m);
check('当日用の Excel のシート', $m[1] === ['受付名簿', '上級集計', '初級集計', '受審順（掲示）', '形講習・レベルチェック'], implode(',', $m[1]));
$s5 = (string)$zip->getFromName('xl/worksheets/sheet5.xml');
check('形講習・レベルチェックのシートに両方の人', str_contains($s5, '部員21') && str_contains($s5, '零D') && str_contains($s5, '形講習のみ'));
$s4 = (string)$zip->getFromName('xl/worksheets/sheet4.xml');
check('受審順（掲示）に形講習会のみの人は出さない', !str_contains($s4, '零D'));
$s1 = (string)$zip->getFromName('xl/worksheets/sheet1.xml');
check('受付名簿は全員（見出し 3 行＋人数）', substr_count($s1, '<row ') === 3 + count(shinsa_round_apps('r1')));
check('開いたときに式を計算し直す', str_contains($wb, 'fullCalcOnLoad="1"'));
$zip->close();

echo $fails ? "\n{$fails} 件 NG\n" : "\n全部通過\n";
exit($fails ? 1 : 0);
