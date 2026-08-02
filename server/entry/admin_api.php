<?php
/**
 * 管理者（専門部）用のAPI。学校側の api.php とは分けてある。
 * 受付状況の集計・パスワード対応・大会設定・データ出力だけを持ち、規則は持たない。
 */
declare(strict_types=1);
error_reporting(E_ALL);
ini_set('display_errors', '0');

require_once __DIR__ . '/auth.php';
require_once __DIR__ . '/uploads.php';

function aout(array $data, int $code = 200): never
{
    http_response_code($code);
    header('Content-Type: application/json; charset=utf-8');
    echo json_encode($data, JSON_UNESCAPED_UNICODE);
    exit;
}

function afail(string $msg, int $code = 400): never
{
    aout(['ok' => false, 'error' => $msg], $code);
}

function abody(): array
{
    $d = json_decode(file_get_contents('php://input') ?: '{}', true);
    return is_array($d) ? $d : [];
}

/** Excelでそのまま開けるCSV（BOM付きUTF-8）を返して終了する */
function csv_out(string $filename, array $header, array $rows): never
{
    header('Content-Type: text/csv; charset=UTF-8');
    header("Content-Disposition: attachment; filename*=UTF-8''" . rawurlencode($filename));
    $fh = fopen('php://output', 'w');
    fwrite($fh, "\xEF\xBB\xBF");
    fputcsv($fh, $header);
    foreach ($rows as $r) {
        fputcsv($fh, $r);
    }
    fclose($fh);
    exit;
}

function weight_list(?string $csv): array
{
    $out = [];
    foreach (explode(',', (string)$csv) as $w) {
        $w = trim($w);
        if ($w === '') {
            continue;
        }
        // 「-55」→「-55kg級」。選抜の「選抜の部」などはそのまま
        $out[] = preg_match('/^[+\-0-9]/', $w) ? "{$w}kg級" : $w;
    }
    return $out;
}

/** 全学校 × 受付中の大会の状況を1回のクエリ群でまとめる */
function collect_status(?array $t): array
{
    $schools = db()->query('SELECT * FROM schools')->fetchAll();
    usort($schools, fn($a, $b) => [(int)$a['school_no'], $a['base_name']] <=> [(int)$b['school_no'], $b['base_name']]);

    // 名簿の人数（学校ごと・学年で絞る）
    $grades = $t ? array_map('intval', (array)($t['grades'] ?? [1, 2, 3])) : [1, 2, 3];
    $memCount = [];
    $memSex = [];
    foreach (db()->query('SELECT school_id, sex, grade FROM members')->fetchAll() as $m) {
        $sid = $m['school_id'];
        $memCount[$sid] = ($memCount[$sid] ?? 0) + 1;
        if (in_array((int)$m['grade'], $grades, true)) {
            $memSex[$sid][$m['sex']] = ($memSex[$sid][$m['sex']] ?? 0) + 1;
        }
    }

    // エントリー（受付中の大会）
    $ent = [];
    $meta = [];
    if ($t) {
        $st = db()->prepare('SELECT entry_key, data_json FROM entries WHERE tournament_id = ?');
        $st->execute([$t['id']]);
        foreach ($st->fetchAll() as $row) {
            $k = (string)$row['entry_key'];
            $d = (array)json_decode($row['data_json'], true);
            if (str_starts_with($k, '_meta_')) {
                $meta[substr($k, 6)] = $d;
            } else {
                $ent[$k] = $d;
            }
        }
    }
    // 押印済み申込書の提出（ファイル名から。DBには持たない＝uploads.php の説明）
    $ups = upload_by_school($t['id'] ?? null);

    // 選手の性別を引くための索引
    $sexOf = [];
    foreach (db()->query('SELECT school_id, name, sex FROM members')->fetchAll() as $m) {
        $sexOf[$m['school_id'] . '_' . $m['name']] = $m['sex'];
    }

    $rows = [];
    foreach ($schools as $s) {
        $sid = (string)$s['school_id'];
        $c = ['男子' => [], '女子' => []];
        foreach (['tk', 'tku', 'k', 'ku'] as $key) {
            $c['男子'][$key] = 0;
            $c['女子'][$key] = 0;
        }
        $total = 0;
        foreach ($ent as $k => $e) {
            if (!str_starts_with($k, $sid . '_')) {
                continue;
            }
            $sex = $sexOf[$k] ?? '男子';
            $any = false;
            if (!empty($e['team_kata_chk'])) { $c[$sex]['tk']++;  $any = true; }
            if (!empty($e['team_kumi_chk'])) { $c[$sex]['tku']++; $any = true; }
            if (!empty($e['kata_chk']))      { $c[$sex]['k']++;   $any = true; }
            if (!empty($e['kumi_chk']))      { $c[$sex]['ku']++;  $any = true; }
            if ($any) {
                $total++;
            }
        }
        $m = $meta[$sid] ?? [];
        $rows[] = [
            'school_id'  => $sid,
            'name'       => $s['base_name'],
            'short_name' => $s['short_name'],
            'school_no'  => (int)$s['school_no'],
            'principal'  => $s['principal'],
            'advisors'   => count((array)json_decode((string)$s['advisors_json'], true)),
            'members'    => $memCount[$sid] ?? 0,
            'eligible_m' => $memSex[$sid]['男子'] ?? 0,
            'eligible_w' => $memSex[$sid]['女子'] ?? 0,
            'counts'     => $c,
            'entered'    => $total,
            'm_mode'     => $m['m_kumite_mode'] ?? '',
            'w_mode'     => $m['w_kumite_mode'] ?? '',
            'touched'    => $meta[$sid] !== null && isset($meta[$sid]),
            'uploads'    => $ups[$sid] ?? [],
        ];
    }
    return $rows;
}

entry_session_start();
$in = abody();
$action = (string)($_GET['action'] ?? ($in['action'] ?? ''));

if ($_SERVER['REQUEST_METHOD'] === 'POST' && ($_SERVER['HTTP_X_ENTRY_API'] ?? '') !== '1') {
    afail('不正なリクエスト', 403);
}

try {
    switch ($action) {

        case 'admin_login':
            if (!admin_login((string)($in['password'] ?? ''))) {
                afail('パスワードが違います', 401);
            }
            aout(['ok' => true]);

        case 'admin_logout':
            unset($_SESSION['admin_ok']);
            aout(['ok' => true]);

        case 'admin_state': {
            require_admin();
            $t = active_tournament();
            aout(['ok' => true,
                  'tournaments' => (array)config_get('tournaments', []),
                  'active'      => $t['id'] ?? null,
                  'year'        => (string)config_get('year', ''),
                  'weak_password' => admin_password_is_weak(),
                  'status'      => collect_status($t),
                  'upload_groups' => upload_groups(),
                  'tournament'  => $t]);
        }

        case 'admin_set_tournament': {
            require_admin();
            $tid = (string)($in['tournament_id'] ?? '');
            $deadline = trim((string)($in['deadline'] ?? ''));
            $ts = (array)config_get('tournaments', []);
            if ($tid !== '' && !isset($ts[$tid])) {
                afail('その大会がありません');
            }
            if ($deadline !== '' && !preg_match('/^\d{4}-\d{2}-\d{2}$/', $deadline)) {
                afail('締切は 2026-09-25 の形で入れてください');
            }
            foreach ($ts as $k => $v) {
                $ts[$k]['active'] = ($k === $tid);
            }
            if ($tid !== '') {
                $ts[$tid]['deadline'] = $deadline;
            }
            db()->prepare('INSERT OR REPLACE INTO config VALUES (?,?)')
                ->execute(['tournaments', json_encode($ts, JSON_UNESCAPED_UNICODE)]);
            aout(['ok' => true]);
        }

        case 'admin_reset_password': {
            require_admin();
            $sid = (string)($in['school_id'] ?? '');
            $pw  = (string)($in['password'] ?? '');
            if (mb_strlen($pw) < 4) {
                afail('パスワードは4文字以上にしてください');
            }
            if (!reset_school_password($sid, $pw)) {
                afail('その学校がありません');
            }
            aout(['ok' => true]);
        }

        case 'admin_set_password': {
            require_admin();
            $pw = (string)($in['password'] ?? '');
            if (mb_strlen($pw) < 8) {
                afail('管理者パスワードは8文字以上にしてください');
            }
            set_admin_password($pw);
            aout(['ok' => true]);
        }

        // ---- データ出力（CSV・BOM付きUTF-8） ----
        case 'export': {
            require_admin();
            $kind = (string)($_GET['kind'] ?? '');
            $t = active_tournament();
            if (!$t) {
                afail('受付中の大会がありません');
            }
            $status = collect_status($t);
            $short = [];
            foreach ($status as $r) {
                $short[$r['school_id']] = $r['short_name'] ?: $r['name'];
            }

            if ($kind === 'schools') {
                $rows = [];
                foreach ($status as $r) {
                    $c = $r['counts'];
                    $rows[] = [$r['school_no'], $r['name'], $r['short_name'], $r['principal'],
                               $r['advisors'], $r['members'], $r['entered'],
                               $c['男子']['tk'], $c['女子']['tk'],
                               $r['m_mode'], $c['男子']['tku'], $r['w_mode'], $c['女子']['tku'],
                               $c['男子']['k'], $c['女子']['k'], $c['男子']['ku'], $c['女子']['ku']];
                }
                csv_out('参加校一覧.csv',
                    ['学校番号', '学校名', '略称', '校長', '顧問数', '部員数', 'エントリー人数',
                     '団体形男', '団体形女', '団体組手男の制', '団体組手男', '団体組手女の制', '団体組手女',
                     '個人形男', '個人形女', '個人組手男', '個人組手女'], $rows);
            }

            if ($kind === 'entries') {
                // 種目ごとに1行ずつの平たい表（Excelで絞り込んで使う）
                $members = db()->query('SELECT * FROM members')->fetchAll();
                $byKey = [];
                foreach ($members as $m) {
                    $byKey[$m['school_id'] . '_' . $m['name']] = $m;
                }
                $st = db()->prepare('SELECT entry_key, data_json FROM entries WHERE tournament_id = ?');
                $st->execute([$t['id']]);
                $meta = [];
                $ent = [];
                foreach ($st->fetchAll() as $row) {
                    $k = (string)$row['entry_key'];
                    if (str_starts_with($k, '_meta_')) {
                        $meta[substr($k, 6)] = (array)json_decode($row['data_json'], true);
                    } else {
                        $ent[$k] = (array)json_decode($row['data_json'], true);
                    }
                }
                $rows = [];
                foreach ($ent as $k => $e) {
                    $m = $byKey[$k] ?? null;
                    if (!$m) {
                        continue;
                    }
                    $sid = $m['school_id'];
                    $sc = $short[$sid] ?? '';
                    $sex = $m['sex'];
                    $base = [$sc, $m['name'], $m['grade'], $m['jkf_no']];
                    if (!empty($e['team_kata_chk'])) {
                        $rows[] = array_merge(["{$sex}団体形", ''], $base,
                            [$e['team_kata_role'] ?: '正', '']);
                    }
                    if (!empty($e['team_kumi_chk'])) {
                        $mode = $meta[$sid][($sex === '女子' ? 'w' : 'm') . '_kumite_mode'] ?? '';
                        $rows[] = array_merge(["{$sex}団体組手", $mode ? "{$mode}人制" : ''], $base,
                            [$e['team_kumi_role'] ?: '正', '']);
                    }
                    if (!empty($e['kata_chk'])) {
                        $rows[] = array_merge(["{$sex}個人形", ''], $base,
                            [$e['kata_val'], $e['kata_rank']]);
                    }
                    if (!empty($e['kumi_chk'])) {
                        $rows[] = array_merge(["{$sex}個人組手", $e['kumi_val']], $base,
                            [$e['kumi_sub_val'] ?: '正', $e['kumi_rank']]);
                    }
                }
                usort($rows, fn($a, $b) => [$a[0], $a[1], $a[2]] <=> [$b[0], $b[1], $b[2]]);
                csv_out('出場者一覧.csv',
                    ['種目', '階級・人数制', '学校', '選手名', '学年', 'JKF番号', '区分', '順位'], $rows);
            }

            if ($kind === 'advisors') {
                // 実物「顧問出欠」と同じ並び: 通番 / 学校名（略称）/ 顧問名 / 1日目 / 2日目
                // 学校は学校番号順、同じ学校の中は登録順。細部は人が手を入れる前提
                $rows = [];
                $n = 0;
                foreach ($status as $r) {
                    $s = school_by_id($r['school_id']);
                    foreach ((array)json_decode((string)$s['advisors_json'], true) as $a) {
                        $rows[] = [++$n, $r['short_name'] ?: $r['name'], $a['name'] ?? '',
                                   !empty($a['d1']) ? '○' : '×', !empty($a['d2']) ? '○' : '×',
                                   $a['role'] ?? ''];
                    }
                }
                csv_out('顧問出欠.csv', ['通番', '学校名', '顧問名', '1日目', '2日目', '役割'], $rows);
            }

            if ($kind === 'schoollist') {
                // 参加校一覧（新人大会の様式・Excel）。実物と同じマトリクスの表
                if (($t['type'] ?? '') !== 'shinjin') {
                    afail('参加校一覧の様式はいまのところ新人大会のぶんだけです');
                }
                require_once ENTRY_LIB_DIR . '/xlsx_fill.php';
                require_once ENTRY_LIB_DIR . '/school_list_form.php';
                $tpl = ENTRY_LIB_DIR . '/template_schools_shinjin.xlsx';
                if (!is_file($tpl)) {
                    afail('参加校一覧のテンプレートがありません');
                }
                // 学校ごとに「選手＋エントリー」を束ねる
                $sexOf = [];
                foreach (db()->query('SELECT school_id, name, sex FROM members')->fetchAll() as $m) {
                    $sexOf[$m['school_id'] . '_' . $m['name']] = $m['sex'];
                }
                $st = db()->prepare('SELECT entry_key, data_json FROM entries WHERE tournament_id = ?');
                $st->execute([$t['id']]);
                $bySchool = [];
                $meta = [];
                foreach ($st->fetchAll() as $row) {
                    $k = (string)$row['entry_key'];
                    $d = (array)json_decode($row['data_json'], true);
                    if (str_starts_with($k, '_meta_')) {
                        $meta[substr($k, 6)] = $d;
                        continue;
                    }
                    foreach ($status as $r) {
                        if (str_starts_with($k, $r['school_id'] . '_')) {
                            $d['sex'] = $sexOf[$k] ?? '男子';
                            $bySchool[$r['school_id']][] = $d;
                            break;
                        }
                    }
                }
                $schools = [];
                foreach ($status as $r) {
                    $schools[] = [
                        'school_no' => $r['school_no'],
                        'name'      => $r['short_name'] ?: $r['name'],
                        'entries'   => $bySchool[$r['school_id']] ?? [],
                        'm_mode'    => $meta[$r['school_id']]['m_kumite_mode'] ?? '5',
                        'w_mode'    => $meta[$r['school_id']]['w_kumite_mode'] ?? '5',
                    ];
                }
                $title = '令和' . config_get('year', '') . '年度　埼玉県空手道新人大会';
                $cells = SchoolListForm::cells($title, $schools,
                    weight_list($t['weights_m'] ?? null), weight_list($t['weights_w'] ?? null));
                $tmp = tempnam(sys_get_temp_dir(), 'slist_') . '.xlsx';
                XlsxFill::fill($tpl, $tmp, $cells);
                header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
                header("Content-Disposition: attachment; filename*=UTF-8''" . rawurlencode('参加校一覧.xlsx'));
                header('Content-Length: ' . (string)filesize($tmp));
                readfile($tmp);
                unlink($tmp);
                exit;
            }

            if ($kind === 'roster_m' || $kind === 'roster_w') {
                // タイマーの team_roster.csv（Team, Player の2列）。補欠も出す
                $sex = $kind === 'roster_m' ? '男子' : '女子';
                $members = db()->query('SELECT * FROM members')->fetchAll();
                $byKey = [];
                foreach ($members as $m) {
                    $byKey[$m['school_id'] . '_' . $m['name']] = $m;
                }
                $st = db()->prepare('SELECT entry_key, data_json FROM entries WHERE tournament_id = ?');
                $st->execute([$t['id']]);
                $rows = [];
                foreach ($st->fetchAll() as $row) {
                    $k = (string)$row['entry_key'];
                    if (str_starts_with($k, '_meta_')) {
                        continue;
                    }
                    $e = (array)json_decode($row['data_json'], true);
                    $m = $byKey[$k] ?? null;
                    if (!$m || $m['sex'] !== $sex || empty($e['team_kumi_chk'])) {
                        continue;
                    }
                    $rows[] = [$short[$m['school_id']] ?? '', $m['name']];
                }
                sort($rows);
                csv_out("team_roster_{$sex}.csv", ['Team', 'Player'], $rows);
            }

            if ($kind === 'draw') {
                // 組合せ抽選システムへ渡す候補一式（JSON・手で直せる中間フォーマット）。
                // 種目・階級ごとに シード選手（順位順）と 学校別の正選手（番手順）を出す。
                // 補欠は表に載らないので入れない。読む側は
                // karate-timer-system/tools/draw_prototype/from_entry.py
                $sexOf = [];
                foreach (db()->query('SELECT school_id, name, sex FROM members')->fetchAll() as $m) {
                    $sexOf[$m['school_id'] . '_' . $m['name']] = $m['sex'];
                }
                $st = db()->prepare('SELECT entry_key, data_json FROM entries WHERE tournament_id = ?');
                $st->execute([$t['id']]);
                $meta = [];
                $ents = [];   // [sid, name, sex, entry]
                foreach ($st->fetchAll() as $row) {
                    $k = (string)$row['entry_key'];
                    $d = (array)json_decode($row['data_json'], true);
                    if (str_starts_with($k, '_meta_')) {
                        $meta[substr($k, 6)] = $d;
                        continue;
                    }
                    foreach ($status as $r) {
                        if (str_starts_with($k, $r['school_id'] . '_')) {
                            $name = substr($k, strlen($r['school_id']) + 1);
                            $ents[] = [$r['school_id'], $name, $sexOf[$k] ?? '男子', $d];
                            break;
                        }
                    }
                }
                $isWeight = weight_list($t['weights_m'] ?? null) !== []
                         || weight_list($t['weights_w'] ?? null) !== [];
                $rankNum = fn($v) => is_numeric(trim((string)$v)) ? (int)trim((string)$v) : 9999;

                // 個人種目1カテゴリぶんを組み立てる共通処理
                $indCat = function (string $key, string $name, string $sex, ?string $class,
                                    string $chkK, string $valK, string $subK, string $rankK)
                        use ($ents, $status, $isWeight, $rankNum) {
                    $seeds = [];
                    $bySchool = [];
                    foreach ($ents as [$sid, $pname, $psex, $e]) {
                        if ($psex !== $sex || empty($e[$chkK])) {
                            continue;
                        }
                        if ($valK === 'kumi_val' && $isWeight) {
                            if (($e['kumi_val'] ?? '') !== $class) {
                                continue;
                            }
                            $kubun = (string)(($e[$subK] ?? '') ?: '正');
                        } else {
                            $kubun = (string)(($e[$valK] ?? '') ?: '正');
                        }
                        $rank = trim((string)($e[$rankK] ?? ''));
                        if ($kubun === 'シード') {
                            $seeds[] = ['seed' => $rankNum($rank), 'rank' => $rank, 'name' => $pname];
                        } elseif ($kubun !== '補') {
                            $bySchool[$sid][] = ['band' => $rankNum($rank), 'name' => $pname];
                        }
                    }
                    usort($seeds, fn($a, $b) => $a['seed'] <=> $b['seed']);
                    $schools = [];
                    $nPlayers = 0;
                    foreach ($status as $r) {          // 学校番号順
                        $sid = $r['school_id'];
                        if (empty($bySchool[$sid])) {
                            continue;
                        }
                        usort($bySchool[$sid], fn($a, $b) => [$a['band'], $a['name']] <=> [$b['band'], $b['name']]);
                        $players = [];
                        foreach ($bySchool[$sid] as $i => $p) {
                            $players[] = ['band' => $i + 1, 'rank' => $p['band'] === 9999 ? '' : (string)$p['band'],
                                          'name' => $p['name']];
                        }
                        $schools[] = ['school' => $r['short_name'] ?: $r['name'], 'players' => $players];
                        $nPlayers += count($players);
                    }
                    // シード側にも学校名を（候補の書式「氏名 (学校名)」に要る）
                    foreach ($seeds as &$s) {
                        foreach ($ents as [$sid2, $pname2, , ]) {
                            if ($pname2 === $s['name']) {
                                foreach ($status as $r2) {
                                    if ($r2['school_id'] === $sid2) {
                                        $s['school'] = $r2['short_name'] ?: $r2['name'];
                                        break;
                                    }
                                }
                                break;
                            }
                        }
                        $s['school'] = $s['school'] ?? '';
                        unset($s['seed']);
                    }
                    unset($s);
                    return ['key' => $key, 'name' => $name, 'kind' => 'individual', 'sex' => $sex,
                            'class' => $class, 'seeds' => $seeds, 'schools' => $schools,
                            'count' => ['seeds' => count($seeds), 'players' => $nPlayers,
                                        'total' => count($seeds) + $nPlayers]];
                };

                $cats = [];
                foreach (['男子' => 'm', '女子' => 'w'] as $sex => $p) {
                    $classes = weight_list($t["weights_{$p}"] ?? null);
                    if ($classes) {
                        foreach ($classes as $w) {
                            $cats[] = $indCat("{$p}_kumite_{$w}", "{$sex}個人組手 {$w}", $sex, $w,
                                              'kumi_chk', 'kumi_val', 'kumi_sub_val', 'kumi_rank');
                        }
                    } else {
                        $cats[] = $indCat("{$p}_kumite", "{$sex}個人組手", $sex, null,
                                          'kumi_chk', 'kumi_val', 'kumi_sub_val', 'kumi_rank');
                    }
                    $cats[] = $indCat("{$p}_kata", "{$sex}個人形", $sex, null,
                                      'kata_chk', 'kata_val', '', 'kata_rank');
                    // 団体（チーム＝学校。表の単位が学校名なので名簿ごと出す）
                    foreach ([['team_kumite', '団体組手', 'team_kumi_chk', 'team_kumi_role'],
                              ['team_kata', '団体形', 'team_kata_chk', 'team_kata_role']] as
                             [$tk, $tn, $chk, $role]) {
                        $teams = [];
                        foreach ($status as $r) {
                            $roster = [];
                            foreach ($ents as [$sid, $pname, $psex, $e]) {
                                if ($sid === $r['school_id'] && $psex === $sex && !empty($e[$chk])) {
                                    $roster[] = ['name' => $pname,
                                                 'role' => ($e[$role] ?? '') === '補' ? '補' : '正'];
                                }
                            }
                            if ($roster) {
                                $team = ['school' => $r['short_name'] ?: $r['name'], 'players' => $roster];
                                if ($tk === 'team_kumite') {
                                    $team['mode'] = $meta[$r['school_id']][($p) . '_kumite_mode'] ?? '5';
                                }
                                $teams[] = $team;
                            }
                        }
                        $cats[] = ['key' => "{$p}_{$tk}", 'name' => "{$sex}{$tn}", 'kind' => $tk,
                                   'sex' => $sex, 'teams' => $teams,
                                   'count' => ['teams' => count($teams)]];
                    }
                }

                $out = ['tournament_id' => $t['id'], 'tournament' => (string)($t['name'] ?? ''),
                        'year' => (string)config_get('year', ''),
                        'generated_at' => date('Y-m-d H:i'),
                        'categories' => $cats];
                header('Content-Type: application/json; charset=utf-8');
                header("Content-Disposition: attachment; filename*=UTF-8''"
                    . rawurlencode("抽選用データ_{$t['id']}.json"));
                echo json_encode($out, JSON_UNESCAPED_UNICODE | JSON_PRETTY_PRINT);
                exit;
            }

            afail('不明な出力: ' . $kind, 404);
        }

        // ---- 抽選会の盤面の受け取り（publish_draw.py が使う） ----
        case 'draw_publish': {
            require_admin();
            $files = (array)($in['files'] ?? []);
            if (!$files || count($files) > 40) {
                afail('ファイルがありません（または多すぎます）');
            }
            // /entry/ の隣の /draw/t/ に置く（drawにはBASIC認証が掛かっている）
            $dir = dirname(__DIR__) . '/draw/t';
            if (!is_dir($dir) && !@mkdir($dir, 0755, true)) {
                afail('draw/t フォルダを作れません: ' . $dir);
            }
            foreach (glob($dir . '/*.html') ?: [] as $old) {
                @unlink($old);          // 前の大会の盤面を残さない
            }
            $n = 0;
            foreach ($files as $f) {
                $name = (string)($f['name'] ?? '');
                if (!preg_match('/^[A-Za-z0-9_\-]{1,64}\.html$/', $name)) {
                    afail('ファイル名が不正: ' . $name);
                }
                $body = base64_decode((string)($f['content_b64'] ?? ''), true);
                if ($body === false || strlen($body) > 2_000_000) {
                    afail('中身が不正: ' . $name);
                }
                if (file_put_contents($dir . '/' . $name, $body) === false) {
                    afail('書き込めません: ' . $name);
                }
                $n++;
            }
            aout(['ok' => true, 'count' => $n]);
        }

        // ---- 提出された押印済み申込書を1件だけ取り出す ----
        case 'admin_download_upload': {
            require_admin();
            $name = (string)($_GET['name'] ?? '');
            $u = upload_parse($name);
            $path = upload_path($name);
            if (!$u || !$path) {
                afail('その提出物はありません', 404);
            }
            $s = school_by_id($u['school_id']);
            $fname = upload_download_name($u, (string)($s['base_name'] ?? $u['school_id']));
            header('Content-Type: application/octet-stream');
            header("Content-Disposition: attachment; filename*=UTF-8''" . rawurlencode($fname));
            header('Content-Length: ' . (string)filesize($path));
            readfile($path);
            exit;
        }

        // ---- まとめてZIP（従来「Googleドライブを開いて一括で落とす」だったもの） ----
        case 'admin_download_zip': {
            require_admin();
            if (!class_exists('ZipArchive')) {
                afail('このサーバーでは ZIP を作れません（1件ずつ取り出してください）');
            }
            $t = active_tournament();
            $all = ($_GET['all'] ?? '') === '1';   // 既定は学校ごとの最新1件だけ
            $info = [];
            foreach (db()->query('SELECT school_id, base_name, school_no FROM schools')->fetchAll() as $r) {
                $info[$r['school_id']] = ['name' => (string)$r['base_name'], 'no' => (int)$r['school_no']];
            }
            $items = [];
            foreach (upload_by_school($t['id'] ?? null) as $sid => $list) {
                $sname = $info[$sid]['name'] ?? $sid;
                foreach (($all ? $list : array_slice($list, 0, 1)) as $u) {
                    $items[] = ['sort'  => [$info[$sid]['no'] ?? 9999, $sname, $u['ts']],
                                'path'  => ENTRY_UPLOAD_DIR . '/' . $u['name'],
                                'entry' => upload_download_name($u, $sname)];
                }
            }
            if (!$items) {
                afail('提出された申込書はまだありません');
            }
            usort($items, fn($a, $b) => $a['sort'] <=> $b['sort']);

            $tmp = (string)tempnam(sys_get_temp_dir(), 'entryzip_');
            $zip = new ZipArchive();
            if ($zip->open($tmp, ZipArchive::OVERWRITE) !== true) {
                afail('ZIPを作れませんでした');
            }
            $used = [];
            foreach ($items as $it) {
                $entry = $it['entry'];
                for ($i = 2; isset($used[$entry]); $i++) {   // 同じ秒の2件目
                    $entry = preg_replace('/(\.[a-z]+)$/', "-{$i}$1", $it['entry']);
                }
                $used[$entry] = true;
                $zip->addFile($it['path'], $entry);
            }
            $zip->close();

            $fname = '申込書_' . ($t['id'] ?? 'all') . '_' . date('Ymd') . '.zip';
            header('Content-Type: application/zip');
            header("Content-Disposition: attachment; filename*=UTF-8''" . rawurlencode($fname));
            header('Content-Length: ' . (string)filesize($tmp));
            readfile($tmp);
            unlink($tmp);
            exit;
        }

        // ---- 提出物を1件消す（間違って上げたものの後始末） ----
        case 'admin_delete_upload': {
            require_admin();
            $name = (string)($in['name'] ?? '');
            $u = upload_parse($name);
            if (!$u || !upload_path($name)) {
                afail('その提出物はありません', 404);
            }
            $s = school_by_id($u['school_id']);
            if (!upload_delete($name)) {
                afail('消せませんでした');
            }
            upload_log_deleted([$name], '1件削除');
            aout(['ok' => true, 'school' => (string)($s['base_name'] ?? $u['school_id']),
                  'at' => $u['at']]);
        }

        // ---- 年度×大会のぶんをまとめて消す（たまったものの整理） ----
        case 'admin_delete_group': {
            require_admin();
            $nendo = (int)($in['nendo'] ?? 0);
            $tid   = (string)($in['tid'] ?? '');
            $expect = (int)($in['expect_count'] ?? -1);
            if ($nendo < 2000 || $tid === '') {
                afail('消す対象が指定されていません');
            }
            // 画面に出ていた件数と今の件数が違えば止める（表示のあとに増えたぶんを
            // 巻き添えで消さないため。もう一度開いて確かめてもらう）
            $now = 0;
            foreach (upload_scan($tid) as $u) {
                if (upload_nendo($u['ts']) === $nendo) {
                    $now++;
                }
            }
            if ($now === 0) {
                afail('その年度・大会の提出物はもうありません');
            }
            if ($expect >= 0 && $expect !== $now) {
                afail("画面を開いたあとに件数が変わりました（{$expect}件 → {$now}件）。"
                    . '画面を読み込み直してから、もう一度確かめてください');
            }
            $done = upload_delete_group($nendo, $tid);
            upload_log_deleted($done, "{$nendo}年度 {$tid} をまとめて削除");
            aout(['ok' => true, 'deleted' => count($done)]);
        }

        default:
            afail('不明な操作: ' . $action, 404);
    }
} catch (Throwable $e) {
    aout(['ok' => false, 'error' => 'サーバー内部の問題: ' . $e->getMessage()], 500);
}
