<?php
/**
 * /entry/ の薄いAPI。保存・認証・申込書の型埋めだけを持ち、規則ロジックは持たない。
 * 入出力は JSON（action で分岐）。エントリーの検証規則は Streamlit 版 app.py の
 * 保存時チェックの移植（重複選択の禁止／正・シードは順位必須／補は順位禁止／
 * 同じ順位の重複禁止／人数上限）。
 */
declare(strict_types=1);
error_reporting(E_ALL);
ini_set('display_errors', '0');

require_once __DIR__ . '/auth.php';
require_once __DIR__ . '/uploads.php';

// 既定の人数制限（config に無い場合の予備。app.py の DEFAULT_LIMITS と同値）
const DEFAULT_LIMITS = [
    'team_kata'     => ['min' => 3, 'max' => 3, 'sub_max' => 1],
    'team_kumite_5' => ['min' => 3, 'max' => 5, 'sub_max' => 2],
    'team_kumite_3' => ['min' => 2, 'max' => 3, 'sub_max' => 1],
    'ind_kata_reg'  => ['max' => 4], 'ind_kata_sub' => ['max' => 2],
    'ind_kumi_reg'  => ['max' => 4], 'ind_kumi_sub' => ['max' => 2],
];

function out(array $data, int $code = 200): never
{
    http_response_code($code);
    header('Content-Type: application/json; charset=utf-8');
    echo json_encode($data, JSON_UNESCAPED_UNICODE);
    exit;
}

function fail(string $msg, int $code = 400): never
{
    out(['ok' => false, 'error' => $msg], $code);
}

function fail_list(array $errors): never
{
    out(['ok' => false, 'errors' => $errors], 422);
}

function body(): array
{
    $raw = file_get_contents('php://input');
    $d = json_decode($raw ?: '{}', true);
    return is_array($d) ? $d : [];
}

/** 階級名の一覧（"-55,-61" → ["-55kg級", "-61kg級"]） */
function weight_names(?string $csv): array
{
    if (!$csv) {
        return [];
    }
    $out = [];
    foreach (explode(',', $csv) as $w) {
        $w = trim($w);
        if ($w !== '') {
            $out[] = "{$w}kg級";
        }
    }
    return $out;
}

function member_is_active(array $m): bool
{
    return !in_array(strtolower(trim((string)($m['active'] ?? ''))), ['false', '0'], true);
}

entry_session_start();
$in = body();
$action = (string)($_GET['action'] ?? ($in['action'] ?? ''));

// POST はヘッダで CSRF を防ぐ（同一オリジンの fetch だけが付けられる）
if ($_SERVER['REQUEST_METHOD'] === 'POST'
    && ($_SERVER['HTTP_X_ENTRY_API'] ?? '') !== '1') {
    fail('不正なリクエスト', 403);
}

try {
    switch ($action) {

        // ---- 公開: 学校名の一覧（ログイン画面のプルダウン用） ----
        case 'schools': {
            $rows = db()->query('SELECT base_name, school_no FROM schools')->fetchAll();
            usort($rows, fn($a, $b) => [(int)$a['school_no'], $a['base_name']] <=> [(int)$b['school_no'], $b['base_name']]);
            out(['ok' => true, 'schools' => array_column($rows, 'base_name')]);
        }

        case 'login': {
            $row = login_school((string)($in['school'] ?? ''), (string)($in['password'] ?? ''));
            if (!$row) {
                fail('学校名またはパスワードが違います', 401);
            }
            out(['ok' => true]);
        }

        case 'logout': {
            $_SESSION = [];
            session_destroy();
            out(['ok' => true]);
        }

        // ---- ログイン後: 画面が必要とする全部 ----
        case 'state': {
            $s = require_school();
            $t = active_tournament();
            $tid = $t['id'] ?? null;
            $ent = $tid ? entries_of($tid, (string)$s['school_id']) : ['meta' => [], 'by_name' => []];
            if ($t) {
                $t['weights_m_names'] = weight_names($t['weights_m'] ?? null);
                $t['weights_w_names'] = weight_names($t['weights_w'] ?? null);
            }
            out(['ok' => true,
                 'school' => [
                     'name'      => $s['base_name'],
                     'principal' => $s['principal'],
                     'advisors'  => (array)json_decode((string)$s['advisors_json'], true),
                 ],
                 'members'    => members_of((string)$s['school_id']),
                 'tournament' => $t,
                 'limits'     => (array)config_get('limits', DEFAULT_LIMITS),
                 'year'       => (string)config_get('year', ''),
                 'entries'    => $ent['by_name'],
                 'meta'       => $ent['meta'],
                 'uploads'    => $tid ? upload_scan($tid, (string)$s['school_id']) : [],
                 'upload_limit' => upload_limit_text(),
                 'upload_limit_bytes' => upload_limit_bytes(),
            ]);
        }

        // ---- 顧問登録（校長名＋顧問一覧の全上書き） ----
        case 'save_advisors': {
            $s = require_school();
            $advisors = [];
            foreach ((array)($in['advisors'] ?? []) as $a) {
                $name = trim((string)($a['name'] ?? ''));
                if ($name === '') {
                    fail_list(['氏名が未入力の顧問行があります']);
                }
                $advisors[] = [
                    'name' => $name,
                    'role' => (string)($a['role'] ?? ''),
                    'd1'   => !empty($a['d1']),
                    'd2'   => !empty($a['d2']),
                ];
            }
            $raw = (array)json_decode((string)$s['raw_json'], true);
            $raw['principal'] = trim((string)($in['principal'] ?? ''));
            $raw['advisors'] = $advisors;
            db()->prepare('UPDATE schools SET principal = ?, advisors_json = ?, raw_json = ? WHERE school_id = ?')
                ->execute([$raw['principal'], json_encode($advisors, JSON_UNESCAPED_UNICODE),
                           json_encode($raw, JSON_UNESCAPED_UNICODE), $s['school_id']]);
            out(['ok' => true]);
        }

        // ---- 部員名簿（学校ぶんの全上書き） ----
        case 'save_members': {
            $s = require_school();
            $rows = [];
            $seen = [];
            $errors = [];
            foreach ((array)($in['members'] ?? []) as $i => $m) {
                $name = trim((string)($m['name'] ?? ''));
                $sex = (string)($m['sex'] ?? '');
                $grade = (string)($m['grade'] ?? '');
                if ($name === '') {
                    continue;   // 空行は無視
                }
                if (!in_array($sex, ['男子', '女子'], true)) {
                    $errors[] = "「{$name}」: 性別を選んでください";
                }
                if (!in_array($grade, ['1', '2', '3'], true)) {
                    $errors[] = "「{$name}」: 学年を選んでください";
                }
                if (isset($seen[$name])) {
                    $errors[] = "「{$name}」が名簿に2回あります（エントリーは名前で選手を区別するため、同姓同名は「山田 太郎B」のように区別してください）";
                }
                $seen[$name] = true;
                $do = trim((string)($m['display_order'] ?? ''));
                if ($do !== '' && !is_numeric($do)) {
                    $errors[] = "「{$name}」: No. は数字で入れてください";
                }
                $rows[] = [$s['school_id'], $name, $sex, $grade,
                           trim((string)($m['dob'] ?? '')), trim((string)($m['jkf_no'] ?? '')),
                           $do === '' ? '' : (string)(int)$do, 'True'];
            }
            if ($errors) {
                fail_list($errors);
            }
            $pdo = db();
            $pdo->beginTransaction();
            $pdo->prepare('DELETE FROM members WHERE school_id = ?')->execute([$s['school_id']]);
            $st = $pdo->prepare('INSERT INTO members VALUES (?,?,?,?,?,?,?,?)');
            foreach ($rows as $r) {
                $st->execute($r);
            }
            $pdo->commit();
            out(['ok' => true, 'count' => count($rows)]);
        }

        // ---- 大会エントリー（学校ぶんの全上書き＋検証） ----
        case 'save_entry': {
            $s = require_school();
            $t = active_tournament();
            if (!$t) {
                fail('受付中の大会がありません');
            }
            $tid = $t['id'];
            $grades = array_map('intval', (array)($t['grades'] ?? [1, 2, 3]));
            $limits = (array)config_get('limits', DEFAULT_LIMITS) + DEFAULT_LIMITS;

            // 名簿（$byName は出場できる学年だけ。$allNames は学年違いの説明用）
            $byName = [];
            $allNames = [];
            foreach (members_of((string)$s['school_id']) as $m) {
                if (!member_is_active($m)) {
                    continue;
                }
                $allNames[$m['name']] = $m;
                if (in_array((int)$m['grade'], $grades, true)) {
                    $byName[$m['name']] = $m;
                }
            }

            // 全員ぶんを一度まっさらにしてから、表の内容を当てる（app.py と同じ全上書き）
            $processed = [];
            foreach ($byName as $name => $_) {
                $processed[$name] = [
                    'team_kata_chk' => false, 'team_kata_role' => '',
                    'team_kumi_chk' => false, 'team_kumi_role' => '',
                    'kata_chk' => false, 'kata_val' => '', 'kata_rank' => '',
                    'kumi_chk' => false, 'kumi_val' => '', 'kumi_rank' => '', 'kumi_sub_val' => '',
                ];
            }

            $errors = [];
            $tables = (array)($in['tables'] ?? []);
            $meta_in = (array)($in['meta'] ?? []);
            $mMode = in_array($meta_in['m_kumite_mode'] ?? '', ['5', '3', 'none'], true) ? $meta_in['m_kumite_mode'] : '5';
            $wMode = in_array($meta_in['w_kumite_mode'] ?? '', ['5', '3', 'none'], true) ? $meta_in['w_kumite_mode'] : '5';
            $rankMap = [];

            $checkName = function (string $n, string $sex, string $label) use ($byName, $allNames, &$errors): bool {
                if (!isset($byName[$n])) {
                    $errors[] = isset($allNames[$n])
                        ? "{$label}: 「{$n}」は{$allNames[$n]['grade']}年生のため、この大会には出場できません"
                        : "{$label}: 「{$n}」は名簿にありません（先に名簿を保存してください）";
                    return false;
                }
                if ($byName[$n]['sex'] !== $sex) {
                    $errors[] = "{$label}: 「{$n}」は{$byName[$n]['sex']}です";
                    return false;
                }
                return true;
            };

            $applyTeam = function (string $key, string $label, string $sex, string $chkK, string $roleK, array $lim)
                    use (&$processed, &$errors, $tables, $checkName) {
                $used = [];
                $reg = $sub = 0;
                foreach ((array)($tables[$key] ?? []) as $row) {
                    $n = trim((string)($row['name'] ?? ''));
                    if ($n === '') {
                        continue;
                    }
                    if (isset($used[$n])) {
                        $errors[] = "{$label}: 「{$n}」が重複して選択されています";
                        continue;
                    }
                    $used[$n] = true;
                    if (!$checkName($n, $sex, $label)) {
                        continue;
                    }
                    $role = ($row['role'] ?? '') === '補' ? '補' : '正';
                    $role === '補' ? $sub++ : $reg++;
                    $processed[$n][$chkK] = true;
                    $processed[$n][$roleK] = $role;
                }
                if ($reg > (int)$lim['max']) {
                    $errors[] = "{$label}: 正選手が{$lim['max']}名を超えています";
                }
                if ($sub > (int)$lim['sub_max']) {
                    $errors[] = "{$label}: 補欠が{$lim['sub_max']}名を超えています";
                }
            };

            $applyInd = function (array $rows, string $label, string $sex, string $chkK, string $valK,
                                  string $rankK, string $subK, bool $isWeight, ?string $weight)
                    use (&$processed, &$errors, &$rankMap, $checkName) {
                $used = [];
                foreach ($rows as $row) {
                    $n = trim((string)($row['name'] ?? ''));
                    $v = (string)($row['kubun'] ?? '');
                    $rk = trim(mb_convert_kana((string)($row['rank'] ?? ''), 'n'));
                    if ($n === '' || in_array($v, ['なし', '出場しない', ''], true)) {
                        continue;
                    }
                    if (isset($used[$n])) {
                        $errors[] = "{$label}: 「{$n}」が重複して選択されています";
                        continue;
                    }
                    $used[$n] = true;
                    if (!$checkName($n, $sex, $label)) {
                        continue;
                    }
                    $needRank = in_array($v, ['正', 'シード'], true);
                    if ($needRank && $rk === '') {
                        $errors[] = "{$label}: 「{$n}」（{$v}）の順位が入力されていません";
                    }
                    if (!$needRank && $rk !== '') {
                        $errors[] = "{$label}: 「{$n}」は補ですが順位が入力されています。順位を消してください";
                    }
                    if ($needRank && $rk !== '') {
                        $rankMap["{$label}|{$v}"][$rk][] = $n;
                    }
                    $processed[$n][$chkK] = true;
                    if ($isWeight && $weight !== null) {
                        $processed[$n][$valK] = $weight;
                        $processed[$n][$subK] = $v;
                    } else {
                        $processed[$n][$valK] = $v;
                    }
                    $processed[$n][$rankK] = $rk;
                }
            };

            $hasNames = fn(string $key) => !empty(array_filter((array)($tables[$key] ?? []),
                fn($r) => trim((string)($r['name'] ?? '')) !== ''));

            // 団体形（「出場しない」を選べる。省略時は出場扱い＝後方互換）
            $tkM = !array_key_exists('part_m_tk', $meta_in) || !empty($meta_in['part_m_tk']);
            $tkW = !array_key_exists('part_w_tk', $meta_in) || !empty($meta_in['part_w_tk']);
            if (!$tkM && $hasNames('m_tk')) {
                $errors[] = '男子 団体形: 「出場しない」なのに選手が選ばれています';
            } elseif ($tkM) {
                $applyTeam('m_tk', '男子 団体形', '男子', 'team_kata_chk', 'team_kata_role', $limits['team_kata']);
            }
            if (!$tkW && $hasNames('w_tk')) {
                $errors[] = '女子 団体形: 「出場しない」なのに選手が選ばれています';
            } elseif ($tkW) {
                $applyTeam('w_tk', '女子 団体形', '女子', 'team_kata_chk', 'team_kata_role', $limits['team_kata']);
            }
            // 団体組手（人数制ごとの上限。出場しないのに選手がいれば誤り）
            if ($mMode === 'none' && $hasNames('m_tku')) {
                $errors[] = '男子 団体組手: 「出場しない」なのに選手が選ばれています';
            } elseif ($mMode !== 'none') {
                $applyTeam('m_tku', '男子 団体組手', '男子', 'team_kumi_chk', 'team_kumi_role', $limits["team_kumite_{$mMode}"]);
            }
            if ($wMode === 'none' && $hasNames('w_tku')) {
                $errors[] = '女子 団体組手: 「出場しない」なのに選手が選ばれています';
            } elseif ($wMode !== 'none') {
                $applyTeam('w_tku', '女子 団体組手', '女子', 'team_kumi_chk', 'team_kumi_role', $limits["team_kumite_{$wMode}"]);
            }
            // 個人形
            $applyInd((array)($tables['m_k'] ?? []), '男子 個人形', '男子', 'kata_chk', 'kata_val', 'kata_rank', '', false, null);
            $applyInd((array)($tables['w_k'] ?? []), '女子 個人形', '女子', 'kata_chk', 'kata_val', 'kata_rank', '', false, null);
            // 個人組手。**階級があるかどうかで持ち方が変わる**（画面側 isWeightTournament と同じ判定）:
            //   階級あり … kumi_val=階級、kumi_sub_val=区分
            //   階級なし … kumi_val=区分
            // 判定を大会の type ではなく「階級が定義されているか」にしてあるのは、
            // 選抜が type=division なのに階級（部）を持っており、type で見ると取り違えるため
            $isWeight = weight_names($t['weights_m'] ?? null) !== []
                     || weight_names($t['weights_w'] ?? null) !== [];
            if ($isWeight) {
                foreach (weight_names($t['weights_m'] ?? null) as $w) {
                    $rows = array_values(array_filter((array)($tables['m_ku'] ?? []), fn($r) => ($r['weight'] ?? '') === $w));
                    $applyInd($rows, "男子 個人組手 {$w}", '男子', 'kumi_chk', 'kumi_val', 'kumi_rank', 'kumi_sub_val', true, $w);
                }
                foreach (weight_names($t['weights_w'] ?? null) as $w) {
                    $rows = array_values(array_filter((array)($tables['w_ku'] ?? []), fn($r) => ($r['weight'] ?? '') === $w));
                    $applyInd($rows, "女子 個人組手 {$w}", '女子', 'kumi_chk', 'kumi_val', 'kumi_rank', 'kumi_sub_val', true, $w);
                }
            } else {
                $applyInd((array)($tables['m_ku'] ?? []), '男子 個人組手', '男子', 'kumi_chk', 'kumi_val', 'kumi_rank', '', false, null);
                $applyInd((array)($tables['w_ku'] ?? []), '女子 個人組手', '女子', 'kumi_chk', 'kumi_val', 'kumi_rank', '', false, null);
            }
            // 同じ順位の重複
            foreach ($rankMap as $key => $ranks) {
                [$label, $v] = explode('|', $key);
                foreach ($ranks as $rk => $names) {
                    if (count($names) > 1) {
                        $errors[] = "{$label}: {$v}の順位 {$rk} が重複しています（" . implode('・', $names) . '）';
                    }
                }
            }
            if ($errors) {
                fail_list(array_values(array_unique($errors)));
            }

            // 正の順位が 1 から連番か。**保存は止めない（警告だけ）**。
            //   - シードは対象外。シードの「順位」は学校内の優先順位ではなく**大会の
            //     シード順位**なので、1人でも「3」があり得る（app.py の案内文と同じ）
            //   - 補も対象外（順位を持たない）
            //   - 抽選側は順位を読まず**人数**で番手を決める（from_entry.py は rank を
            //     参照しない／rules_check.effective_rank は 1人校の番手を無効にする）＝
            //     ここが連番でなくてもデータは壊れない。だからエラーにはしない。
            //     拾いたいのは「1人しか出さないのに順位2」のような**入力の取り違え**
            $warnings = [];
            foreach ($rankMap as $key => $ranks) {
                [$label, $v] = explode('|', $key);
                if ($v !== '正') {
                    continue;
                }
                // 重複はこの時点で無い（上で弾いてある）ので、件数＝人数
                $got  = array_map('strval', array_keys($ranks));
                $want = array_map('strval', range(1, count($got)));
                if (array_diff($want, $got)) {
                    sort($got, SORT_NATURAL);
                    $warnings[] = "{$label}: 正の順位が1から連番になっていません"
                        . '（いまは ' . implode('・', $got) . '／' . count($got) . '名）。'
                        . '保存はしましたが、入力の間違いがないか確認してください';
                }
            }

            // メタ（Streamlit に切り戻しても画面が読めるよう part_* も立てておく）
            $meta = entries_of($tid, (string)$s['school_id'])['meta'];
            $meta['m_kumite_mode'] = $mMode;
            $meta['w_kumite_mode'] = $wMode;
            $any = fn(string $sex, string $chk) => (bool)array_filter($processed,
                fn($p, $n) => $byName[$n]['sex'] === $sex && $p[$chk], ARRAY_FILTER_USE_BOTH);
            $meta['part_m_tk'] = $tkM;
            $meta['part_w_tk'] = $tkW;
            $meta['part_m_tku'] = $mMode !== 'none';
            $meta['part_w_tku'] = $wMode !== 'none';
            $meta['part_m_k'] = $any('男子', 'kata_chk');
            $meta['part_w_k'] = $any('女子', 'kata_chk');
            $meta['part_m_ku'] = $any('男子', 'kumi_chk');
            $meta['part_w_ku'] = $any('女子', 'kumi_chk');
            if ($isWeight) {
                foreach (weight_names($t['weights_m'] ?? null) as $w) {
                    $meta["part_m_ku_{$w}"] = (bool)array_filter($processed,
                        fn($p, $n) => $byName[$n]['sex'] === '男子' && $p['kumi_chk'] && $p['kumi_val'] === $w, ARRAY_FILTER_USE_BOTH);
                }
                foreach (weight_names($t['weights_w'] ?? null) as $w) {
                    $meta["part_w_ku_{$w}"] = (bool)array_filter($processed,
                        fn($p, $n) => $byName[$n]['sex'] === '女子' && $p['kumi_chk'] && $p['kumi_val'] === $w, ARRAY_FILTER_USE_BOTH);
                }
            }

            // 保存（学校ぶんを消してから入れ直す）
            $pdo = db();
            $pdo->beginTransaction();
            $pdo->prepare("DELETE FROM entries WHERE tournament_id = ? AND (entry_key = ? OR entry_key LIKE ? ESCAPE '\\')")
                ->execute([$tid, '_meta_' . $s['school_id'], like_escape((string)$s['school_id']) . '\_%']);
            $st = $pdo->prepare('INSERT OR REPLACE INTO entries VALUES (?,?,?)');
            foreach ($processed as $name => $p) {
                $st->execute([$tid, $s['school_id'] . '_' . $name, json_encode($p, JSON_UNESCAPED_UNICODE)]);
            }
            $st->execute([$tid, '_meta_' . $s['school_id'], json_encode($meta, JSON_UNESCAPED_UNICODE)]);
            $pdo->commit();
            out($warnings ? ['ok' => true, 'warnings' => array_values(array_unique($warnings))]
                          : ['ok' => true]);
        }

        // ---- 申込書のダウンロード（④のエンジンで型埋め） ----
        case 'download_form': {
            $s = require_school();
            $t = active_tournament();
            if (!$t) {
                fail('受付中の大会がありません');
            }
            require_once ENTRY_LIB_DIR . '/xlsx_fill.php';
            require_once ENTRY_LIB_DIR . '/entry_form.php';

            // 大会ごとの様式（config の template / coords）。無ければ標準に落ちる
            $coordsFile = basename((string)($t['coords'] ?? 'coords_standard.json'));
            $tplFile    = basename((string)($t['template'] ?? 'template.xlsx'));
            $coordsPath = ENTRY_LIB_DIR . '/' . $coordsFile;
            $tplPath    = ENTRY_LIB_DIR . '/' . $tplFile;
            if (!is_file($coordsPath)) {
                fail("この大会の座標ファイルがありません（{$coordsFile}）");
            }
            if (!is_file($tplPath)) {
                fail("この大会の申込書テンプレートがありません（{$tplFile}）");
            }
            $coords = json_decode((string)file_get_contents($coordsPath), true, 512, JSON_THROW_ON_ERROR);

            $ent = entries_of($t['id'], (string)$s['school_id']);
            $members = [];
            foreach (members_of((string)$s['school_id']) as $m) {
                if (member_is_active($m)) {
                    $members[] = array_merge($m, $ent['by_name'][$m['name']] ?? []);
                }
            }
            $school = [
                'year'            => (string)config_get('year', ''),
                'tournament_name' => (string)($t['name'] ?? ''),
                'school_name'     => (string)$s['base_name'],
                'principal'       => (string)$s['principal'],
                'advisors'        => (array)json_decode((string)$s['advisors_json'], true),
                'm_kumite_mode'   => (string)($ent['meta']['m_kumite_mode'] ?? '5'),
                'w_kumite_mode'   => (string)($ent['meta']['w_kumite_mode'] ?? '5'),
            ];
            $grades = array_map('intval', (array)($t['grades'] ?? [1, 2, 3]));
            $cells = EntryForm::cells($coords, $school, $members, $grades);
            $tmp = tempnam(sys_get_temp_dir(), 'entry_') . '.xlsx';
            XlsxFill::fill($tplPath, $tmp, $cells);

            $fname = '申込書_' . $s['base_name'] . '.xlsx';
            header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
            header("Content-Disposition: attachment; filename*=UTF-8''" . rawurlencode($fname));
            header('Content-Length: ' . (string)filesize($tmp));
            readfile($tmp);
            unlink($tmp);
            exit;
        }

        // ---- 押印済み申込書の提出（従来の GAS→Googleドライブ の置き換え） ----
        case 'upload_form': {
            $s = require_school();
            $t = active_tournament();
            if (!$t) {
                fail('受付中の大会がありません');
            }
            // 上限を超えた POST は本文ごと捨てられ、$_FILES が空で届く。
            // 「何も起きない」ように見えるので、ここで理由を返す
            if (!$_FILES && (int)($_SERVER['CONTENT_LENGTH'] ?? 0) > 0) {
                fail('ファイルが大きすぎます（上限 ' . upload_limit_text() . '）', 413);
            }
            $res = upload_store((string)$t['id'], (string)$s['school_id'], (array)($_FILES['file'] ?? []));
            if (!$res['ok']) {
                fail($res['error']);
            }
            out(['ok' => true, 'upload' => $res['upload'],
                 'uploads' => upload_scan((string)$t['id'], (string)$s['school_id'])]);
        }

        // ---- 提出したものを学校が取り戻す（自分のぶんだけ） ----
        case 'download_upload': {
            $s = require_school();
            $name = (string)($_GET['name'] ?? '');
            $u = upload_parse($name);
            $path = upload_path($name);
            if (!$u || !$path || $u['school_id'] !== (string)$s['school_id']) {
                fail('その提出物はありません', 404);
            }
            $fname = upload_download_name($u, (string)$s['base_name']);
            header('Content-Type: application/octet-stream');
            header("Content-Disposition: attachment; filename*=UTF-8''" . rawurlencode($fname));
            header('Content-Length: ' . (string)filesize($path));
            readfile($path);
            exit;
        }

        default:
            fail('不明な操作: ' . $action, 404);
    }
} catch (Throwable $e) {
    out(['ok' => false, 'error' => 'サーバー内部の問題: ' . $e->getMessage()], 500);
}
