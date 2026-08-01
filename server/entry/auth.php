<?php
/**
 * 学校ログイン（現行と同じ「学校名＋パスワード」）。
 * パスワードは移行時に pbkdf2_sha256$反復$salt$hash 形式で入っている
 * （tools/migrate_from_sheets.py の docstring と対）。
 */
declare(strict_types=1);

require_once __DIR__ . '/db.php';

function entry_session_start(): void
{
    if (session_status() === PHP_SESSION_ACTIVE) {
        return;
    }
    session_name('KARATEENTRY');
    session_set_cookie_params([
        'lifetime' => 0,
        'path'     => '/',
        'httponly' => true,
        'samesite' => 'Lax',
        'secure'   => !empty($_SERVER['HTTPS']),
    ]);
    session_start();
}

function verify_password(string $plain, string $stored): bool
{
    $p = explode('$', $stored);
    if (count($p) !== 4 || $p[0] !== 'pbkdf2_sha256') {
        return false;
    }
    $raw = hash_pbkdf2('sha256', $plain, hex2bin($p[2]), (int)$p[1], 0, true);
    return hash_equals($p[3], bin2hex($raw));
}

/** 学校名（基本名）でログイン。成功したら school_id をセッションへ */
function login_school(string $schoolName, string $password): ?array
{
    $st = db()->prepare('SELECT * FROM schools WHERE base_name = ?');
    $st->execute([trim($schoolName)]);
    $row = $st->fetch();
    if (!$row || !verify_password($password, (string)$row['password_hash'])) {
        usleep(500_000);   // 総当たりをゆっくりにする
        return null;
    }
    session_regenerate_id(true);
    $_SESSION['school_id'] = $row['school_id'];
    return $row;
}

function current_school(): ?array
{
    if (empty($_SESSION['school_id'])) {
        return null;
    }
    return school_by_id((string)$_SESSION['school_id']);
}

/* ---- 管理者（専門部） ---- */

/**
 * 移行直後の admin_password は Sheets からの平文なので、ハッシュ・平文の
 * どちらでも照合できるようにしておく（パスワード変更でハッシュに移る）。
 */
function verify_admin_password(string $plain): bool
{
    $stored = (string)config_get('admin_password', '');
    if ($stored === '') {
        return false;
    }
    if (str_starts_with($stored, 'pbkdf2_sha256$')) {
        return verify_password($plain, $stored);
    }
    return hash_equals($stored, $plain);
}

function admin_password_is_weak(): bool
{
    $stored = (string)config_get('admin_password', '');
    return !str_starts_with($stored, 'pbkdf2_sha256$');
}

function set_admin_password(string $plain): void
{
    $salt = random_bytes(16);
    $hash = hash_pbkdf2('sha256', $plain, $salt, 100000, 0, true);
    $stored = 'pbkdf2_sha256$100000$' . bin2hex($salt) . '$' . bin2hex($hash);
    db()->prepare('INSERT OR REPLACE INTO config VALUES (?,?)')
        ->execute(['admin_password', json_encode($stored, JSON_UNESCAPED_UNICODE)]);
}

function admin_login(string $password): bool
{
    if (!verify_admin_password($password)) {
        usleep(500_000);
        return false;
    }
    session_regenerate_id(true);
    $_SESSION['admin_ok'] = true;
    return true;
}

function is_admin(): bool
{
    return !empty($_SESSION['admin_ok']);
}

function require_admin(): void
{
    if (!is_admin()) {
        http_response_code(401);
        header('Content-Type: application/json; charset=utf-8');
        echo json_encode(['ok' => false, 'error' => 'ログインしてください'], JSON_UNESCAPED_UNICODE);
        exit;
    }
}

/** 学校のパスワードを作り直す（管理者用） */
function reset_school_password(string $schoolId, string $plain): bool
{
    $salt = random_bytes(16);
    $hash = hash_pbkdf2('sha256', $plain, $salt, 100000, 0, true);
    $stored = 'pbkdf2_sha256$100000$' . bin2hex($salt) . '$' . bin2hex($hash);
    $st = db()->prepare('UPDATE schools SET password_hash = ? WHERE school_id = ?');
    $st->execute([$stored, $schoolId]);
    return $st->rowCount() > 0;
}

function require_school(): array
{
    $s = current_school();
    if (!$s) {
        http_response_code(401);
        echo json_encode(['ok' => false, 'error' => 'ログインしてください'], JSON_UNESCAPED_UNICODE);
        exit;
    }
    return $s;
}
