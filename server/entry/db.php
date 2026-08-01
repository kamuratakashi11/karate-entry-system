<?php
/**
 * SQLite への薄いアクセス層。スキーマは tools/migrate_from_sheets.py が作ったもの
 * （schools / members / entries / config）をそのまま使う。
 */
declare(strict_types=1);

require_once __DIR__ . '/config.php';

function db(): PDO
{
    static $pdo = null;
    if ($pdo === null) {
        if (!is_file(ENTRY_DB_PATH)) {
            throw new RuntimeException('データベースがありません: ' . ENTRY_DB_PATH);
        }
        $pdo = new PDO('sqlite:' . ENTRY_DB_PATH);
        $pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
        $pdo->setAttribute(PDO::ATTR_DEFAULT_FETCH_MODE, PDO::FETCH_ASSOC);
        $pdo->exec('PRAGMA busy_timeout = 5000');
    }
    return $pdo;
}

function config_get(string $key, $default = null)
{
    $st = db()->prepare('SELECT value_json FROM config WHERE key = ?');
    $st->execute([$key]);
    $row = $st->fetch();
    if (!$row) {
        return $default;
    }
    $v = json_decode($row['value_json'], true);
    return $v === null ? $default : $v;
}

/** 受付中の大会（config.tournaments の active が立っている最初の1つ） */
function active_tournament(): ?array
{
    foreach ((array)config_get('tournaments', []) as $tid => $t) {
        if (!empty($t['active'])) {
            return ['id' => (string)$tid] + (array)$t;
        }
    }
    return null;
}

function school_by_id(string $sid): ?array
{
    $st = db()->prepare('SELECT * FROM schools WHERE school_id = ?');
    $st->execute([$sid]);
    $row = $st->fetch();
    return $row ?: null;
}

/** 学校の部員（表示順 → 男子先 → 学年降順 → 名前。app.py の確認リストと同じ感覚の並び） */
function members_of(string $sid): array
{
    $st = db()->prepare('SELECT * FROM members WHERE school_id = ?');
    $st->execute([$sid]);
    $rows = $st->fetchAll();
    usort($rows, function ($a, $b) {
        $oa = is_numeric(trim((string)$a['display_order'])) ? (float)$a['display_order'] : 999999.0;
        $ob = is_numeric(trim((string)$b['display_order'])) ? (float)$b['display_order'] : 999999.0;
        if ($oa !== $ob) {
            return $oa <=> $ob;
        }
        $sa = $a['sex'] === '女子' ? 1 : 0;
        $sb = $b['sex'] === '女子' ? 1 : 0;
        if ($sa !== $sb) {
            return $sa <=> $sb;
        }
        $cmp = (int)$b['grade'] <=> (int)$a['grade'];
        if ($cmp !== 0) {
            return $cmp;
        }
        return strcmp((string)$a['name'], (string)$b['name']);
    });
    return $rows;
}

/** 学校のエントリー（entry_key "{sid}_{name}" → 中身。_meta_ は 'meta' キーで返す） */
function entries_of(string $tid, string $sid): array
{
    $st = db()->prepare('SELECT entry_key, data_json FROM entries WHERE tournament_id = ?');
    $st->execute([$tid]);
    $out = ['meta' => [], 'by_name' => []];
    $prefix = $sid . '_';
    foreach ($st->fetchAll() as $row) {
        $k = (string)$row['entry_key'];
        if ($k === '_meta_' . $sid) {
            $out['meta'] = (array)json_decode($row['data_json'], true);
        } elseif (str_starts_with($k, $prefix)) {
            $out['by_name'][substr($k, strlen($prefix))] = (array)json_decode($row['data_json'], true);
        }
    }
    return $out;
}

/** LIKE 用エスケープ（school_id に _ % が入っていても安全に前方一致させる） */
function like_escape(string $s): string
{
    return str_replace(['\\', '%', '_'], ['\\\\', '\\%', '\\_'], $s);
}
