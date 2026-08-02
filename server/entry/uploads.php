<?php
/**
 * 押印済み申込書の置き場（~/entry_data/uploads/）を扱う共通部品。
 * 学校側 api.php と管理者側 admin_api.php の両方から使う。
 *
 * 置き場所は公開ディレクトリの外なので URL では取れない。取り出しは API 経由だけ
 * （学校は自分のぶん・管理者は全部）。従来の GAS→Googleドライブ の置き換え。
 *
 * **記録をDBに持たないのは意図**: 受付切替や不具合時に entry.sqlite3 を丸ごと
 * 入れ替える運用（prepare_deploy.py → setup.php）があるため、提出の記録をDBに
 * 置くと入れ替えで消える。ファイル名だけで完結させてある:
 *
 *     <大会ID>__<学校ID>__<日時>.<拡張子>
 *     shinjin__sch_20260120103625__20260801_141500.pdf
 *
 * 学校名はファイル名に入れない（学校名が変わると追えなくなるため）。表示名・
 * ダウンロード名は school_id から引き直して組み立てる。
 */
declare(strict_types=1);

require_once __DIR__ . '/db.php';

/**
 * 一覧・取り出し・削除で「うちの提出物」と認めるもの。
 * **新しく受け取るのは PDF だけ**（下の UPLOAD_ACCEPT_EXTS）だが、ここを狭めると
 * 以前に受け取った写真が一覧から消えて取り出せなくなるので、旧形式も残す。
 */
const UPLOAD_EXTS = ['pdf', 'jpg', 'jpeg', 'png'];

/**
 * 新しく受け取る拡張子。**PDF だけ**（2026-08-02 決定・孝さん）。
 * 申込書は22名を超えると2ページになる。写真だと2枚バラバラの提出になり、
 * 「いちばん新しいものが提出物」の決まりと噛み合わず1ページ目が抜ける。
 * PDFなら複数ページが1つにまとまり、iPhone の HEIC 変換にも頼らずに済む。
 */
const UPLOAD_ACCEPT_EXTS = ['pdf'];

/** 1ファイルの上限。php.ini の上限がこれより小さければそちらが効く */
const UPLOAD_CAP_BYTES = 20 * 1024 * 1024;

/** php.ini の "2M" 形式をバイト数に直す（0・空は「制限なし」） */
function upload_ini_bytes(string $v): int
{
    $v = trim($v);
    if ($v === '') {
        return 0;
    }
    $n = (int)$v;
    return match (strtolower(substr($v, -1))) {
        'g' => $n * 1024 * 1024 * 1024,
        'm' => $n * 1024 * 1024,
        'k' => $n * 1024,
        default => $n,
    };
}

/**
 * 実際に受け取れる1ファイルの上限。
 * php.ini（upload_max_filesize / post_max_size）とこちらの上限の小さい方。
 * post_max_size は本文全体の上限なので、境界の書式ぶん少し引いておく。
 */
function upload_limit_bytes(): int
{
    $limit = UPLOAD_CAP_BYTES;
    $up = upload_ini_bytes((string)ini_get('upload_max_filesize'));
    if ($up > 0) {
        $limit = min($limit, $up);
    }
    $post = upload_ini_bytes((string)ini_get('post_max_size'));
    if ($post > 0) {
        $limit = min($limit, max(0, $post - 8 * 1024));
    }
    return $limit;
}

function upload_limit_text(): string
{
    return number_format(upload_limit_bytes() / 1024 / 1024, 1) . ' MB';
}

function upload_size_text(int $bytes): string
{
    return $bytes >= 1024 * 1024
        ? number_format($bytes / 1024 / 1024, 1) . ' MB'
        : max(1, (int)round($bytes / 1024)) . ' KB';
}

/** 中身が拡張子どおりか（先頭の数バイトで見る。別物の取り違えを弾くため） */
function upload_looks_like(string $path, string $ext): bool
{
    $fh = @fopen($path, 'rb');
    if (!$fh) {
        return false;
    }
    $head = (string)fread($fh, 8);
    fclose($fh);
    return match ($ext) {
        'pdf'         => str_starts_with($head, '%PDF-'),
        'jpg', 'jpeg' => str_starts_with($head, "\xFF\xD8\xFF"),
        'png'         => str_starts_with($head, "\x89PNG\r\n\x1A\n"),
        default       => false,
    };
}

/** ファイル名 → 中身。うちの流儀で無いものは null（他のファイルを混ぜても無視される） */
function upload_parse(string $name): ?array
{
    if (!preg_match('/^([A-Za-z0-9\-]+)__([A-Za-z0-9_\-]+)__(\d{8}_\d{6})(?:-(\d+))?\.([a-z]+)$/', $name, $m)) {
        return null;
    }
    if (!in_array($m[5], UPLOAD_EXTS, true)) {
        return null;
    }
    return [
        'name'      => $name,
        'tid'       => $m[1],
        'school_id' => $m[2],
        'ts'        => $m[3],
        'seq'       => (int)($m[4] ?? 0) ?: 1,   // 同じ秒の2件目以降
        'ext'       => $m[5],
        'at'        => upload_at_text($m[3]),
    ];
}

/** "20260801_141500" → "2026-08-01 14:15" */
function upload_at_text(string $ts): string
{
    return sprintf('%s-%s-%s %s:%s',
        substr($ts, 0, 4), substr($ts, 4, 2), substr($ts, 6, 2),
        substr($ts, 9, 2), substr($ts, 11, 2));
}

/** 保存されている提出物の一覧（新しい順）。tid / sid で絞れる */
function upload_scan(?string $tid = null, ?string $sid = null): array
{
    if (!is_dir(ENTRY_UPLOAD_DIR)) {
        return [];
    }
    $out = [];
    foreach ((array)scandir(ENTRY_UPLOAD_DIR) as $name) {
        $u = upload_parse((string)$name);
        if (!$u
            || ($tid !== null && $u['tid'] !== $tid)
            || ($sid !== null && $u['school_id'] !== $sid)) {
            continue;
        }
        $u['size'] = (int)@filesize(ENTRY_UPLOAD_DIR . '/' . $name);
        $u['size_text'] = upload_size_text($u['size']);
        $out[] = $u;
    }
    usort($out, fn($a, $b) => [$b['ts'], $b['seq']] <=> [$a['ts'], $a['seq']]);
    return $out;
}

/** 学校ごとにまとめる（[school_id => 新しい順の配列]） */
function upload_by_school(?string $tid = null): array
{
    $out = [];
    foreach (upload_scan($tid) as $u) {
        $out[$u['school_id']][] = $u;
    }
    return $out;
}

/**
 * 提出年度（4月始まり）。"20260801_141500" → 2026、"20270210_..." → 2026。
 * 大会ID（shinjin など）は毎年おなじなので、年度で分けないと去年のぶんを整理できない。
 */
function upload_nendo(string $ts): int
{
    $y = (int)substr($ts, 0, 4);
    return (int)substr($ts, 4, 2) >= 4 ? $y : $y - 1;
}

/** 年度×大会でまとめる（整理用）。新しいものが先 */
function upload_groups(): array
{
    $g = [];
    foreach (upload_scan() as $u) {
        $nendo = upload_nendo($u['ts']);
        $k = "{$nendo}|{$u['tid']}";
        if (!isset($g[$k])) {
            $g[$k] = ['nendo' => $nendo, 'tid' => $u['tid'], 'count' => 0,
                      'bytes' => 0, 'schools' => [], 'first' => $u['ts'], 'last' => $u['ts']];
        }
        $g[$k]['count']++;
        $g[$k]['bytes'] += $u['size'];
        $g[$k]['schools'][$u['school_id']] = true;
        $g[$k]['first'] = min($g[$k]['first'], $u['ts']);
        $g[$k]['last']  = max($g[$k]['last'], $u['ts']);
    }
    $out = [];
    foreach ($g as $v) {
        $v['schools']   = count($v['schools']);
        $v['size_text'] = upload_size_text($v['bytes']);
        $v['first_at']  = upload_at_text($v['first']);
        $v['last_at']   = upload_at_text($v['last']);
        $out[] = $v;
    }
    usort($out, fn($a, $b) => [$b['nendo'], $b['last']] <=> [$a['nendo'], $a['last']]);
    return $out;
}

/** 消したものを控える（www の外・追記のみ）。「去年のぶんはどこへ？」に答えられるように */
function upload_log_deleted(array $names, string $why): void
{
    $line = date('Y-m-d H:i') . "\t{$why}\t" . count($names) . "件\t" . implode(' ', $names) . "\n";
    @file_put_contents(dirname(ENTRY_UPLOAD_DIR) . '/uploads_deleted.log', $line, FILE_APPEND);
}

/** 1件消す。名前がうちの流儀でなければ何もしない（経路を作らせない） */
function upload_delete(string $name): bool
{
    $path = upload_path($name);
    return $path !== null && @unlink($path);
}

/** 年度×大会のぶんをまとめて消す。消した名前を返す */
function upload_delete_group(int $nendo, string $tid): array
{
    $done = [];
    foreach (upload_scan($tid) as $u) {
        if (upload_nendo($u['ts']) === $nendo && upload_delete($u['name'])) {
            $done[] = $u['name'];
        }
    }
    return $done;
}

/** 保存名 → 実体の場所。名前がうちの流儀で無ければ null（経路を作らせない） */
function upload_path(string $name): ?string
{
    if (!upload_parse($name)) {
        return null;
    }
    $path = ENTRY_UPLOAD_DIR . '/' . $name;
    return is_file($path) ? $path : null;
}

/** ダウンロード時の名前。従来（GAS）と同じ「【学校名】_申込書_日時.拡張子」 */
function upload_download_name(array $u, string $schoolName): string
{
    return "【{$schoolName}】_申込書_{$u['ts']}.{$u['ext']}";
}

/**
 * 受け取って保存する。$f は $_FILES の1件。
 * 失敗は例外にせず ['ok'=>false,'error'=>...] で返す（そのまま画面に出すため）。
 */
function upload_store(string $tid, string $sid, array $f): array
{
    $err = (int)($f['error'] ?? UPLOAD_ERR_NO_FILE);
    if ($err !== UPLOAD_ERR_OK) {
        return ['ok' => false, 'error' => match ($err) {
            UPLOAD_ERR_INI_SIZE, UPLOAD_ERR_FORM_SIZE =>
                'ファイルが大きすぎます（上限 ' . upload_limit_text() . '）',
            UPLOAD_ERR_NO_FILE    => 'ファイルが選ばれていません',
            UPLOAD_ERR_PARTIAL    => '送信が途中で切れました。もう一度お試しください',
            UPLOAD_ERR_NO_TMP_DIR, UPLOAD_ERR_CANT_WRITE =>
                'サーバーが一時ファイルを書けませんでした（専門部へご連絡ください）',
            default               => '受け取れませんでした（コード ' . $err . '）',
        }];
    }
    $tmp = (string)($f['tmp_name'] ?? '');
    if ($tmp === '' || !is_uploaded_file($tmp)) {
        return ['ok' => false, 'error' => '受け取れませんでした'];
    }

    $ext = strtolower((string)pathinfo((string)($f['name'] ?? ''), PATHINFO_EXTENSION));
    if (!in_array($ext, UPLOAD_ACCEPT_EXTS, true)) {
        return ['ok' => false, 'error' =>
            '申込書は PDF で提出してください（選ばれたもの: '
            . ($ext === '' ? '拡張子なし' : ".{$ext}") . '）。'
            . 'スマホで撮る場合は、写真ではなく「書類をスキャン」でPDFにしてください'
            . '（画面の「スマートフォンでPDFにするには」をご覧ください）。'
            . 'どうしてもPDFにできない場合は、加村までメールをください'];
    }
    $size = (int)($f['size'] ?? 0);
    if ($size <= 0) {
        return ['ok' => false, 'error' => '中身が空のファイルです'];
    }
    if ($size > upload_limit_bytes()) {
        return ['ok' => false, 'error' => 'ファイルが大きすぎます（上限 ' . upload_limit_text()
            . '／選ばれたもの ' . upload_size_text($size) . '）'];
    }
    if (!upload_looks_like($tmp, $ext)) {
        return ['ok' => false, 'error' => "中身が {$ext} ではないようです。"
            . '拡張子を書き換えたファイルは受け取れません'];
    }

    if (!is_dir(ENTRY_UPLOAD_DIR) && !@mkdir(ENTRY_UPLOAD_DIR, 0700, true)) {
        return ['ok' => false, 'error' => '提出先のフォルダがありません（専門部へご連絡ください）'];
    }
    // 大会IDは名前の区切りに使うので英数字だけにする（区切りの __ が壊れないように）
    $tidSafe = preg_replace('/[^A-Za-z0-9\-]/', '-', $tid) ?: 'x';
    $base = "{$tidSafe}__{$sid}__" . date('Ymd_His');
    $name = "{$base}.{$ext}";
    for ($i = 2; is_file(ENTRY_UPLOAD_DIR . '/' . $name); $i++) {   // 同じ秒に2つ来たとき
        $name = "{$base}-{$i}.{$ext}";
    }
    $dest = ENTRY_UPLOAD_DIR . '/' . $name;
    if (!@move_uploaded_file($tmp, $dest)) {
        return ['ok' => false, 'error' => '保存できませんでした（専門部へご連絡ください）'];
    }
    @chmod($dest, 0600);

    $u = upload_parse($name);
    $u['size'] = $size;
    $u['size_text'] = upload_size_text($size);
    return ['ok' => true, 'upload' => $u];
}
