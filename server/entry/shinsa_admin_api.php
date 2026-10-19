<?php
/**
 * 審査の管理画面（shinsa_admin.html）の API。審査担当の先生が使う。
 * 大会の管理画面（admin_api.php）とはパスワードが別（孝さん 2026-10-14
 * 「出欠席の管理や合否の管理を別の先生に渡したいので管理者画面は分ける」）。
 * 大会の管理者（専門部）はそのまま入れる。決まりは shinsa.php にあり、ここは受け渡しだけ。
 */
declare(strict_types=1);
error_reporting(E_ALL);
ini_set('display_errors', '0');

require_once __DIR__ . '/shinsa.php';

function sout(array $data, int $code = 200): never
{
    http_response_code($code);
    header('Content-Type: application/json; charset=utf-8');
    echo json_encode($data, JSON_UNESCAPED_UNICODE);
    exit;
}

function sfail(string $msg, int $code = 400): never
{
    sout(['ok' => false, 'error' => $msg], $code);
}

function require_shinsa_admin(): void
{
    if (!is_shinsa_admin()) {
        sout(['ok' => false, 'error' => 'ログインしてください'], 401);
    }
}

function round_or_fail(string $rid): array
{
    $r = shinsa_rounds()[$rid] ?? null;
    if (!$r) {
        sfail('その回がありません', 404);
    }
    return $r;
}

entry_session_start();
$raw = file_get_contents('php://input');
$in = json_decode($raw ?: '{}', true);
$in = is_array($in) ? $in : [];
$action = (string)($_GET['action'] ?? ($in['action'] ?? ''));

if ($_SERVER['REQUEST_METHOD'] === 'POST' && ($_SERVER['HTTP_X_ENTRY_API'] ?? '') !== '1') {
    sfail('不正なリクエスト', 403);
}

try {
    switch ($action) {
        case 'login': {
            if (!shinsa_admin_password_set()) {
                sfail('審査の管理画面のパスワードがまだ決まっていません。専門部の管理画面（大会設定）で決めてください');
            }
            if (!shinsa_admin_login((string)($in['password'] ?? ''))) {
                sfail('パスワードが違います', 401);
            }
            sout(['ok' => true]);
        }

        case 'logout': {
            unset($_SESSION['shinsa_ok']);
            sout(['ok' => true, 'still_admin' => is_admin()]);
        }

        // 回の一覧（全部）と、それぞれの申込の人数
        case 'state': {
            require_shinsa_admin();
            $counts = shinsa_round_counts();
            $rounds = array_values(shinsa_rounds());
            usort($rounds, fn($a, $b) => strcmp((string)($a['date'] ?? ''), (string)($b['date'] ?? '')));
            foreach ($rounds as &$r) {
                $r['title'] = shinsa_round_title($r);
                $r['counts'] = $counts[$r['id']] ?? [];
            }
            unset($r);
            sout(['ok' => true, 'rounds' => $rounds, 'today' => date('Y-m-d'),
                  'as_admin' => empty($_SESSION['shinsa_ok']) && is_admin(),
                  'groups' => shinsa_groups(), 'kubun' => shinsa_options()['kubun'], 'fee' => SHINSA_FEE]);
        }

        // 選んだ回の、学校ごとの申込状況
        case 'round': {
            require_shinsa_admin();
            $rid = (string)($_GET['round_id'] ?? '');
            $r = round_or_fail($rid);
            $ordered = shinsa_ordered(shinsa_round_apps($rid));
            sout(['ok' => true, 'round' => $r + ['title' => shinsa_round_title($r)],
                  'status' => shinsa_status($rid),
                  'group_counts' => array_map('count', $ordered)]);
        }

        case 'save_rounds': {
            require_shinsa_admin();
            $errors = shinsa_save_rounds((array)($in['rounds'] ?? []));
            if ($errors) {
                sout(['ok' => false, 'errors' => $errors], 422);
            }
            sout(['ok' => true]);
        }

        case 'number': {
            require_shinsa_admin();
            $rid = (string)($in['round_id'] ?? '');
            round_or_fail($rid);
            sout(['ok' => true, 'counts' => shinsa_assign_numbers($rid)]);
        }

        case 'receipt': {
            require_shinsa_admin();
            $errors = shinsa_set_receipt((string)($in['round_id'] ?? ''), !empty($in['on']), (string)($in['date'] ?? ''));
            if ($errors) {
                sout(['ok' => false, 'errors' => $errors], 422);
            }
            sout(['ok' => true]);
        }

        case 'set_password': {
            require_shinsa_admin();
            $pw = (string)($in['password'] ?? '');
            if (mb_strlen($pw) < 8) {
                sfail('パスワードは 8 文字以上にしてください');
            }
            set_shinsa_admin_password($pw);
            sout(['ok' => true]);
        }

        // 当日用の Excel（受付名簿・集計・受審順・レベルチェック）
        case 'day': {
            require_shinsa_admin();
            $rid = (string)($_GET['round_id'] ?? '');
            $r = round_or_fail($rid);
            require_once __DIR__ . '/shinsa_day.php';
            $tmp = tempnam(sys_get_temp_dir(), 'shinsa_day_') . '.xlsx';
            shinsa_day_book($r + ['id' => $rid], $tmp);
            $fname = '級位審査_当日用_令和' . ($r['year'] ?? '') . '年度第' . ($r['no'] ?? '') . '回.xlsx';
            header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
            header("Content-Disposition: attachment; filename*=UTF-8''" . rawurlencode($fname));
            header('Content-Length: ' . (string)filesize($tmp));
            readfile($tmp);
            unlink($tmp);
            exit;
        }

        default:
            sfail('不明な操作: ' . $action, 404);
    }
} catch (Throwable $e) {
    sout(['ok' => false, 'error' => 'サーバー内部の問題: ' . $e->getMessage()], 500);
}
