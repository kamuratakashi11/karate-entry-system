<?php
/**
 * 級位審査の申込（2026-10-14 孝さん「大会申込と級位審査申込を選べるようにする」）。
 *
 * **審査の決まり（選べる級・形・区分、受付できるか、申込書のどのセルに何を書くか）は
 * このファイルの 1 か所に置く。** 画面（app.js・shinsa_print.js）は API が返す
 * options を描くだけで、選択肢を自分では持たない。
 *
 * 決まりの出どころ:
 *   - 選択肢 … 審査部会の様式「級位審査申込書」（受審申込書シートの入力規則と AD17:AD49 の形の一覧）
 *   - 区分 … 1〜4級を受けるのが上級、5〜8級が初級。形講習会・レベルチェックテストは
 *            様式の注意事項⑤「【第1・3回限定】…希望する場合は〇」（受講料は無料）
 *   - 全空連の会員登録が無いと受審できない（注意事項③）
 *
 * 審査の回は config の shinsa_rounds（回の ID → 設定）。設定するのは審査の管理画面（段階 2）。
 *   year … 令和の年度（数字）   no … 第何回   date … 審査期日 YYYY-MM-DD
 *   deadline … 申込の締め切り YYYY-MM-DD   open … 学校に見せるか   lc … 形講習・レベルチェックがある回か
 *
 * 申込は表 shinsa_apps に、回・学校・氏名ごとに 1 行。並び（order）は学校が並べた順で、
 * 受審番号を振るときに「各校の 1 人目 → 2 人目 → …」の順になる（受付の Excel の★シート）。
 */
declare(strict_types=1);

require_once __DIR__ . '/auth.php';

/**
 * 合格の基準（合計点 → 合格級。上から順に見て、最初に届いたもの。どれにも届かなければ不合格）。
 * 審査部会の集計表（00.第1回まとめ.xlsx の ⑦上級集計・⑧初級集計）の式から写した（2026-10-14）。
 * 当日用の Excel の「合格級」の式はここから作る。
 */
const SHINSA_PASS = [
    'jokyu'  => [[42, '1級'], [39, '2級'], [36, '3級'], [30, '4級']],
    'shokyu' => [[27, '5級'], [22.5, '6級'], [18, '7級'], [17.5, '8級']],
];

/** 受審料（級位審査だけ。形講習・レベルチェックは無料）。不合格は 1,000 円を返して 2,000 円 */
const SHINSA_FEE = ['pass' => 3000, 'fail' => 2000];

const SHINSA_KATA = [
    'ゲキサイ第一', 'ゲキサイ第二',
    '糸東流平安初段', '糸東流平安二段', '糸東流平安三段', '糸東流平安四段', '糸東流平安五段',
    '松濤館平安初段', '松濤館平安二段', '松濤館平安三段', '松濤館平安四段', '松濤館平安五段',
    'ピンアン初段', 'ピンアン二段', 'ピンアン三段', 'ピンアン四段', 'ピンアン五段',
    'サイファ', 'セーパイ', 'クルルンファ', 'セイサン', 'ジオン', 'カンクウダイ', 'カンクウショウ',
    'エンピ', 'バッサイ(大)', 'セイエンチン', 'マツムラローハイ', 'ニーパイポ', 'セイシャン',
    'チントウ', 'ニーセーシー', 'クーシャンク―',
];

/** 区分 → 画面の表示名・選べる現在の級・審査用紙 */
function shinsa_kubun(): array
{
    $kyu = ['2級', '3級', '4級', '5級', '6級', '7級', '8級', '無級'];
    return [
        // fee … 級位審査として受ける（受審料がかかる＝領収書を出す。申込書では受審級を書く）か。
        //        false の区分は無料で、申込書では受審級を空けて「形講習のみ希望」に〇（注意事項⑤）
        // sheet … 審査用紙（shinsa/sheet_*.png）。null は刷らない
        // lc_round … 形講習会・レベルチェックテストのある回（第1・3回）だけ選べる
        'jokyu'  => ['label' => '上級（1〜4級を受ける）', 'short' => '上級', 'kyu' => $kyu,
                     'sheet' => 'jokyu', 'fee' => true, 'lc_round' => false],
        'shokyu' => ['label' => '初級（5〜8級を受ける）', 'short' => '初級', 'kyu' => $kyu,
                     'sheet' => 'shokyu', 'fee' => true, 'lc_round' => false],
        // 黒帯の子が今のレベルを確かめる（級位審査とは別。無料・領収書なし。受審番号も別に振る）
        'lc'     => ['label' => 'レベルチェックテスト（黒帯）', 'short' => 'レベルチェック',
                     'kyu' => ['1級', '初段', '二段', '三段'], 'sheet' => 'lc', 'fee' => false, 'lc_round' => true],
        // 形講習会だけを受ける（審査は受けない。審査用紙は刷らない。孝さん 2026-10-19）
        'koshu'  => ['label' => '形講習会のみ', 'short' => '形講習のみ',
                     'kyu' => array_merge($kyu, ['1級', '初段', '二段', '三段']), 'sheet' => null, 'fee' => false,
                     'lc_round' => true],
    ];
}

/** 画面に渡す選択肢（画面はこれを描くだけ） */
function shinsa_options(): array
{
    $k = [];
    foreach (shinsa_kubun() as $id => $v) {
        $k[] = ['id' => $id, 'label' => $v['label'], 'short' => $v['short'], 'kyu' => $v['kyu'],
                'sheet' => $v['sheet'], 'fee' => $v['fee'], 'lc_round' => $v['lc_round']];
    }
    return ['kubun' => $k, 'kata' => SHINSA_KATA, 'years' => range(0, 15), 'months' => range(0, 11),
            'fee' => SHINSA_FEE];
}

function shinsa_db(): PDO
{
    static $ready = false;
    $pdo = db();
    if (!$ready) {
        // 本番の DB にあとから足す表。無ければ作る（大会の表には触れない）
        $pdo->exec('CREATE TABLE IF NOT EXISTS shinsa_apps (
            round_id TEXT, school_id TEXT, name TEXT, data_json TEXT,
            PRIMARY KEY (round_id, school_id, name))');
        $ready = true;
    }
    return $pdo;
}

function shinsa_rounds(): array
{
    $out = [];
    foreach ((array)config_get('shinsa_rounds', []) as $rid => $r) {
        $out[(string)$rid] = ['id' => (string)$rid] + (array)$r;
    }
    return $out;
}

/** 学校に見せている回（open。いくつもあれば審査期日の早いほう） */
function shinsa_open_round(): ?array
{
    $open = array_values(array_filter(shinsa_rounds(), fn($r) => !empty($r['open'])));
    usort($open, fn($a, $b) => strcmp((string)($a['date'] ?? ''), (string)($b['date'] ?? '')));
    return $open[0] ?? null;
}

/** 申込を直せるか（見せていて、締め切りの日を過ぎていない） */
function shinsa_can_edit(?array $r): bool
{
    if (!$r || empty($r['open'])) {
        return false;
    }
    $d = (string)($r['deadline'] ?? '');
    return $d === '' || date('Y-m-d') <= $d;
}

function shinsa_round_title(array $r): string
{
    return '令和' . ($r['year'] ?? '') . '年度 第' . ($r['no'] ?? '') . '回 級位審査会';
}

/** 学校の申込（並び順） */
function shinsa_apps_of(string $rid, string $sid): array
{
    $st = shinsa_db()->prepare('SELECT name, data_json FROM shinsa_apps WHERE round_id = ? AND school_id = ?');
    $st->execute([$rid, $sid]);
    $out = [];
    foreach ($st->fetchAll() as $row) {
        $out[] = ['name' => (string)$row['name']] + (array)json_decode((string)$row['data_json'], true);
    }
    usort($out, fn($a, $b) => (int)($a['order'] ?? 0) <=> (int)($b['order'] ?? 0));
    return $out;
}

/**
 * 学校の申込を全部置き換える（大会のエントリーと同じ全上書き）。
 * 性別・学年・会員番号は名簿から写して持つ（申込の時点の内容を残す。4 月に学年が
 * 上がっても、出した申込書と審査用紙は変わらない）。
 *
 * @return array{0: array<int,string>, 1: int} [エラー, 保存した人数]
 */
function shinsa_save(array $round, array $school, array $rows): array
{
    $kubun = shinsa_kubun();
    $roster = [];
    foreach (members_of((string)$school['school_id']) as $m) {
        if (member_is_active($m)) {
            $roster[(string)$m['name']] = $m;
        }
    }
    $before = [];
    foreach (shinsa_apps_of((string)$round['id'], (string)$school['school_id']) as $a) {
        $before[$a['name']] = $a;
    }

    $errors = [];
    $seen = [];
    $save = [];
    foreach (array_values($rows) as $i => $r) {
        $name = trim((string)($r['name'] ?? ''));
        if ($name === '') {
            continue;
        }
        $who = "「{$name}」";
        if (isset($seen[$name])) {
            $errors[] = "{$who}が 2 回選ばれています";
            continue;
        }
        $seen[$name] = true;
        $m = $roster[$name] ?? null;
        if (!$m) {
            $errors[] = "{$who}は部員名簿にありません（先に「部員名簿」で保存してください）";
            continue;
        }
        $k = (string)($r['kubun'] ?? '');
        if (!isset($kubun[$k])) {
            $errors[] = "{$who}: 区分を選んでください";
            continue;
        }
        if ($kubun[$k]['lc_round'] && empty($round['lc'])) {
            $errors[] = "{$who}: この回は形講習会・レベルチェックテストがありません（第1回・第3回だけ）";
        }
        $kyu = (string)($r['kyu'] ?? '');
        if (!in_array($kyu, $kubun[$k]['kyu'], true)) {
            $errors[] = "{$who}: " . ($kubun[$k]['fee'] ? '現在の公認級' : '現在の級・段') . 'を選んでください';
        }
        $kata = (string)($r['kata'] ?? '');
        if (!in_array($kata, SHINSA_KATA, true)) {
            $errors[] = "{$who}: 演武する形を選んでください";
        }
        $y = (string)($r['years'] ?? '');
        $mo = (string)($r['months'] ?? '');
        if (!preg_match('/^\d{1,2}$/', $y) || (int)$y > 15 || !preg_match('/^\d{1,2}$/', $mo) || (int)$mo > 11) {
            $errors[] = "{$who}: 通算修業期間（年・か月）を選んでください";
        }
        // 会員登録が要るのは審査を受ける人（注意事項③「受審できません」）。形講習会のみは要らない
        if (trim((string)$m['jkf_no']) === '' && $kubun[$k]['sheet'] !== null) {
            $errors[] = "{$who}: 部員名簿に JKF 番号（全空連会員番号）がありません。会員登録が無いと受審できません";
        }
        $save[$name] = [
            'order'  => count($save) + 1,
            'kubun'  => $k, 'kyu' => $kyu, 'kata' => $kata,
            'years'  => (int)$y, 'months' => (int)$mo,
            'sex'    => (string)$m['sex'], 'grade' => (string)$m['grade'],
            'jkf_no' => trim((string)$m['jkf_no']),
            // 受審番号（締め切りのあとに審査の管理画面で振る）と、領収書のための合否
            // （審査のあとに顧問が選ぶ）は、直しても同じ人・同じ区分なら残す
            'no'     => (($before[$name]['kubun'] ?? '') === $k) ? ($before[$name]['no'] ?? '') : '',
            'result' => (($before[$name]['kubun'] ?? '') === $k) ? ($before[$name]['result'] ?? '') : '',
        ];
    }
    if ($errors) {
        return [array_values(array_unique($errors)), 0];
    }
    $pdo = shinsa_db();
    $pdo->beginTransaction();
    $pdo->prepare('DELETE FROM shinsa_apps WHERE round_id = ? AND school_id = ?')
        ->execute([$round['id'], $school['school_id']]);
    $st = $pdo->prepare('INSERT INTO shinsa_apps VALUES (?,?,?,?)');
    foreach ($save as $name => $d) {
        $st->execute([$round['id'], $school['school_id'], $name, json_encode($d, JSON_UNESCAPED_UNICODE)]);
    }
    $pdo->commit();
    return [[], count($save)];
}

/** 'YYYY-MM-DD' → [令和の年, 月, 日]（読めなければ空） */
function shinsa_reiwa(string $ymd): array
{
    if (!preg_match('/^(\d{4})-(\d{2})-(\d{2})$/', $ymd, $m)) {
        return ['', '', ''];
    }
    return [(int)$m[1] - 2018, (int)$m[2], (int)$m[3]];
}

/** 申込書・審査用紙に書く学校名（様式の側に「高等学校」「高校」と刷ってあるので、それを除く） */
function shinsa_school_label(array $school): string
{
    return (string)preg_replace('/(高等学校|高校)$/u', '', (string)$school['base_name']);
}

/** 紙の申込書 1 枚に入る人数（様式の 17〜36 行目） */
const SHINSA_FORM_ROWS = 20;

/**
 * 紙の申込書（審査部会の様式「受審申込書 」シート）に書くセル。
 * $page は 1 から。21 人目からは 2 枚目（同じ様式をもう 1 つ落とす）。
 * 形講習・レベルチェックの人は、受審級を空けて「形講習のみ希望」に〇（様式の注意事項⑤）。
 * 引率顧問名（39〜41 行目）は「来られない顧問名は記載しない」ので書かない（手で書く）。
 */
function shinsa_form_cells(array $round, array $school, array $apps, int $page): array
{
    [$ry, $rm, $rd] = shinsa_reiwa((string)($round['date'] ?? ''));
    [$dy, $dm, $dd] = shinsa_reiwa((string)($round['deadline'] ?? ''));
    $advisors = (array)json_decode((string)($school['advisors_json'] ?? ''), true);
    $c = [
        'E6' => is_numeric($round['year'] ?? '') ? (int)$round['year'] : (string)($round['year'] ?? ''),
        'J6' => is_numeric($round['no'] ?? '') ? (int)$round['no'] : (string)($round['no'] ?? ''),
        'V6' => $ry, 'X6' => $rm, 'Z6' => $rd,
        'I8' => $dy, 'K8' => $dm, 'M8' => $dd,
        'G10' => shinsa_school_label($school),
        'G12' => (string)($school['principal'] ?? ''),
        'T12' => (string)($advisors[0]['name'] ?? ''),
    ];
    $kubun = shinsa_kubun();
    $slice = array_slice(array_values($apps), ($page - 1) * SHINSA_FORM_ROWS, SHINSA_FORM_ROWS);
    foreach ($slice as $i => $a) {
        $r = 17 + $i;
        $lc = !$kubun[$a['kubun']]['fee'];   // 形講習・レベルチェック（無料）
        $c["A{$r}"] = ($page - 1) * SHINSA_FORM_ROWS + $i + 1;
        $c["B{$r}"] = (string)($a['jkf_no'] ?? '');
        $c["H{$r}"] = (string)$a['name'];
        $c["L{$r}"] = is_numeric($a['grade'] ?? '') ? (int)$a['grade'] : (string)($a['grade'] ?? '');
        $c["N{$r}"] = $lc ? '' : (string)($a['sex'] ?? '') . $kubun[$a['kubun']]['short'];
        $c["P{$r}"] = (string)($a['kyu'] ?? '');
        $c["R{$r}"] = (string)($a['kata'] ?? '');
        $c["V{$r}"] = (int)($a['years'] ?? 0);
        $c["X{$r}"] = (int)($a['months'] ?? 0);
        $c["Z{$r}"] = $lc ? '〇' : '';
    }
    return array_filter($c, fn($v) => $v !== '');
}

/** 申込書が何枚になるか */
function shinsa_form_pages(array $apps): int
{
    return max(1, (int)ceil(count($apps) / SHINSA_FORM_ROWS));
}

/**
 * これからの締め切り（大会と級位審査）。ログインのあとの選ぶ画面に並べる。
 * 大会の締め切りは大会設定（tournaments[大会]['deadline']）、審査は回の設定にある。
 * 締め切りの持ち主はそれぞれ 1 か所で、ここは集めて並べるだけ。
 */
function upcoming_deadlines(?string $today = null): array
{
    $today ??= date('Y-m-d');
    $out = [];
    foreach ((array)config_get('tournaments', []) as $tid => $t) {
        $d = (string)($t['deadline'] ?? '');
        if (preg_match('/^\d{4}-\d{2}-\d{2}$/', $d) && $d >= $today) {
            $out[] = ['kind' => '大会', 'name' => (string)($t['name'] ?? $tid), 'deadline' => $d,
                      'open' => !empty($t['active'])];
        }
    }
    foreach (shinsa_rounds() as $r) {
        $d = (string)($r['deadline'] ?? '');
        if (preg_match('/^\d{4}-\d{2}-\d{2}$/', $d) && $d >= $today) {
            $out[] = ['kind' => '級位審査', 'name' => '第' . ($r['no'] ?? '') . '回 級位審査会', 'deadline' => $d,
                      'date' => (string)($r['date'] ?? ''), 'open' => !empty($r['open'])];
        }
    }
    usort($out, fn($a, $b) => strcmp($a['deadline'], $b['deadline']));
    return $out;
}

/** 学校の画面に渡す審査のぶん（state に入れる） */
function shinsa_state_for(array $school): array
{
    $r = shinsa_open_round();
    $apps = $r ? shinsa_apps_of((string)$r['id'], (string)$school['school_id']) : [];
    return [
        'round'    => $r ? $r + ['title' => shinsa_round_title($r)] : null,
        'school_label' => shinsa_school_label($school),
        // 領収書に刷る日付（令和の 年・月・日）。出せるようにした回だけ
        'receipt_reiwa' => $r && !empty($r['receipt_on']) ? shinsa_reiwa((string)($r['receipt_date'] ?? '')) : null,
        'can_edit' => shinsa_can_edit($r),
        'apps'     => $apps,
        'form_pages' => shinsa_form_pages($apps),
        'options'  => shinsa_options(),
    ];
}

/* =====================================================================
 * 審査の管理画面（審査担当の先生。大会の管理画面とは別のパスワード）
 * ===================================================================== */

/** 審査の管理者か。大会の管理者（専門部）も入れる */
function is_shinsa_admin(): bool
{
    return !empty($_SESSION['shinsa_ok']) || is_admin();
}

function shinsa_admin_password_set(): bool
{
    return str_starts_with((string)config_get('shinsa_admin_password', ''), 'pbkdf2_sha256$');
}

function shinsa_admin_login(string $plain): bool
{
    $stored = (string)config_get('shinsa_admin_password', '');
    if ($stored === '' || !verify_password($plain, $stored)) {
        usleep(500_000);
        return false;
    }
    session_regenerate_id(true);
    $_SESSION['shinsa_ok'] = true;
    return true;
}

function set_shinsa_admin_password(string $plain): void
{
    db()->prepare('INSERT OR REPLACE INTO config VALUES (?,?)')
        ->execute(['shinsa_admin_password', json_encode(make_password_hash($plain), JSON_UNESCAPED_UNICODE)]);
}

function shinsa_put_rounds(array $rounds): void
{
    $store = [];
    foreach ($rounds as $id => $r) {
        unset($r['id']);
        $store[(string)$id] = $r;
    }
    db()->prepare('INSERT OR REPLACE INTO config VALUES (?,?)')
        ->execute(['shinsa_rounds', json_encode($store, JSON_UNESCAPED_UNICODE)]);
}

/** 回ごとの申込の人数（区分ごと） */
function shinsa_round_counts(): array
{
    $out = [];
    foreach (shinsa_db()->query('SELECT round_id, data_json FROM shinsa_apps')->fetchAll() as $row) {
        $d = (array)json_decode((string)$row['data_json'], true);
        $rid = (string)$row['round_id'];
        $out[$rid]['total'] = ($out[$rid]['total'] ?? 0) + 1;
        $k = (string)($d['kubun'] ?? '');
        $out[$rid][$k] = ($out[$rid][$k] ?? 0) + 1;
    }
    return $out;
}

/**
 * 回の一覧をまとめて保存する（1 年分を 3 月に登録する。孝さん 2026-10-14）。
 * 送られなかった回は消す。ただし申込のある回は消さない（誤操作で申込が消えないように）。
 * 受審番号・領収書の設定（numbered_at・receipt_*）は、ここでは変えずに残す。
 *
 * @return array<int,string> エラー（空なら保存した）
 */
function shinsa_save_rounds(array $rows): array
{
    $old = shinsa_rounds();
    $counts = shinsa_round_counts();
    $ymd = '/^\d{4}-\d{2}-\d{2}$/';
    $errors = [];
    $new = [];
    $seen = [];
    foreach (array_values($rows) as $i => $r) {
        $label = '（' . ($i + 1) . '行目）';
        $year = trim(mb_convert_kana((string)($r['year'] ?? ''), 'n'));
        $no = trim(mb_convert_kana((string)($r['no'] ?? ''), 'n'));
        $date = trim((string)($r['date'] ?? ''));
        $dl = trim((string)($r['deadline'] ?? ''));
        if (!preg_match('/^\d{1,2}$/', $year) || (int)$year < 1) {
            $errors[] = "{$label} 年度（令和）を数字で入れてください";
        }
        if (!preg_match('/^\d$/', $no) || (int)$no < 1) {
            $errors[] = "{$label} 第何回かを数字で入れてください";
        }
        if (!preg_match($ymd, $date)) {
            $errors[] = "{$label} 審査期日を選んでください";
        }
        if (!preg_match($ymd, $dl)) {
            $errors[] = "{$label} 申込の締め切りを選んでください";
        } elseif (preg_match($ymd, $date) && $dl > $date) {
            $errors[] = "{$label} 締め切りが審査期日より後になっています";
        }
        $key = "{$year}-{$no}";
        if (isset($seen[$key])) {
            $errors[] = "令和{$year}年度 第{$no}回 が 2 つあります";
        }
        $seen[$key] = true;
        $id = (string)($r['id'] ?? '');
        if ($id === '' || !isset($old[$id])) {
            $id = 'r' . bin2hex(random_bytes(4));
        }
        $keep = $old[$id] ?? [];
        unset($keep['id']);
        $new[$id] = array_merge($keep, [
            'year' => (int)$year, 'no' => (int)$no, 'date' => $date, 'deadline' => $dl,
            'open' => !empty($r['open']), 'lc' => !empty($r['lc']),
        ]);
    }
    foreach ($old as $id => $r) {
        if (!isset($new[$id]) && !empty($counts[$id]['total'])) {
            $errors[] = shinsa_round_title($r) . ' は申込が ' . $counts[$id]['total']
                . ' 人あるので消せません（学校に見せないなら「学校に見せる」を外してください）';
        }
    }
    if ($errors) {
        return array_values(array_unique($errors));
    }
    shinsa_put_rounds($new);
    return [];
}

/** 領収書の設定（出すか・日付）。日付は令和で領収書に刷る */
function shinsa_set_receipt(string $rid, bool $on, string $date): array
{
    $rounds = shinsa_rounds();
    if (!isset($rounds[$rid])) {
        return ['その回がありません'];
    }
    if (!preg_match('/^\d{4}-\d{2}-\d{2}$/', $date)) {
        return ['領収書の日付を選んでください'];
    }
    $rounds[$rid]['receipt_on'] = $on;
    $rounds[$rid]['receipt_date'] = $date;
    shinsa_put_rounds($rounds);
    return [];
}

/** 回の申込を全部（学校番号順 → 学校が並べた順） */
function shinsa_round_apps(string $rid): array
{
    $schools = [];
    foreach (db()->query('SELECT school_id, base_name, school_no FROM schools')->fetchAll() as $s) {
        $schools[(string)$s['school_id']] = $s;
    }
    $st = shinsa_db()->prepare('SELECT school_id, name, data_json FROM shinsa_apps WHERE round_id = ?');
    $st->execute([$rid]);
    $out = [];
    foreach ($st->fetchAll() as $row) {
        $s = $schools[(string)$row['school_id']] ?? ['base_name' => '?', 'school_no' => 9999];
        $out[] = ['name' => (string)$row['name'], 'school_id' => (string)$row['school_id'],
                  'school' => (string)$s['base_name'], 'school_label' => shinsa_school_label($s),
                  'school_no' => (int)$s['school_no']]
               + (array)json_decode((string)$row['data_json'], true);
    }
    usort($out, fn($a, $b) => [$a['school_no'], $a['school'], (int)($a['order'] ?? 0)]
                          <=> [$b['school_no'], $b['school'], (int)($b['order'] ?? 0)]);
    return $out;
}

/** 受審番号を振る組（この組ごとに 1 番から）と、組の中の並べ方 */
function shinsa_groups(): array
{
    return [
        // number … 受審番号を振るか（形講習会のみの人は審査を受けないので振らない）
        ['key' => 'jokyu_F',  'label' => '上級 女子', 'kubun' => 'jokyu',  'sex' => '女子', 'order' => 'interleave', 'number' => true],
        ['key' => 'jokyu_M',  'label' => '上級 男子', 'kubun' => 'jokyu',  'sex' => '男子', 'order' => 'interleave', 'number' => true],
        ['key' => 'shokyu_F', 'label' => '初級 女子', 'kubun' => 'shokyu', 'sex' => '女子', 'order' => 'school', 'number' => true],
        ['key' => 'shokyu_M', 'label' => '初級 男子', 'kubun' => 'shokyu', 'sex' => '男子', 'order' => 'school', 'number' => true],
        ['key' => 'lc',       'label' => 'レベルチェック', 'kubun' => 'lc', 'sex' => null, 'order' => 'school', 'number' => true],
        ['key' => 'koshu',    'label' => '形講習会のみ', 'kubun' => 'koshu', 'sex' => null, 'order' => 'school', 'number' => false],
    ];
}

/**
 * 組ごとの受審順。受付の Excel（00.第1回まとめ.xlsx）のやり方を写した（2026-10-14）:
 *   上級 … 各校の 1 人目を学校番号順に → 2 人目を → …（自由組手で 1 番どうしが当たる。★シート）
 *   初級・レベルチェック … 学校番号順に、学校の中は申込書の順
 * 男女は別に 1 番から数え、女子が先（集計表の並び）。
 *
 * @param array $apps shinsa_round_apps() の結果（学校番号順・学校の中は申込書の順）
 * @return array<string, list<array>> 組の key → 受審順の申込
 */
function shinsa_ordered(array $apps): array
{
    $out = [];
    foreach (shinsa_groups() as $g) {
        $mine = array_values(array_filter($apps, fn($a) => ($a['kubun'] ?? '') === $g['kubun']
            && ($g['sex'] === null || ($a['sex'] ?? '') === $g['sex'])));
        if ($g['order'] === 'interleave') {
            $bySchool = [];
            foreach ($mine as $a) {
                $bySchool[$a['school_id']][] = $a;   // 入れた順＝学校番号順のまま
            }
            $list = [];
            for ($k = 0; $bySchool; $k++) {
                foreach ($bySchool as $sid => $l) {
                    if (!isset($l[$k])) {
                        unset($bySchool[$sid]);
                        continue;
                    }
                    $list[] = $l[$k];
                }
            }
            $mine = $list;
        }
        $out[$g['key']] = $mine;
    }
    return $out;
}

/** 受審番号を振る（振り直すと前の番号は上書き）。@return array<string,int> 組ごとの人数 */
function shinsa_assign_numbers(string $rid): array
{
    $ordered = shinsa_ordered(shinsa_round_apps($rid));
    $numbered = array_column(shinsa_groups(), 'number', 'key');
    $pdo = shinsa_db();
    $pdo->beginTransaction();
    $get = $pdo->prepare('SELECT data_json FROM shinsa_apps WHERE round_id = ? AND school_id = ? AND name = ?');
    $put = $pdo->prepare('UPDATE shinsa_apps SET data_json = ? WHERE round_id = ? AND school_id = ? AND name = ?');
    $counts = [];
    foreach ($ordered as $key => $list) {
        foreach ($list as $i => $a) {
            $get->execute([$rid, $a['school_id'], $a['name']]);
            $d = (array)json_decode((string)$get->fetchColumn(), true);
            $d['no'] = $numbered[$key] ? $i + 1 : '';
            $put->execute([json_encode($d, JSON_UNESCAPED_UNICODE), $rid, $a['school_id'], $a['name']]);
        }
        $counts[$key] = count($list);
    }
    $pdo->commit();
    $rounds = shinsa_rounds();
    $rounds[$rid]['numbered_at'] = date('Y-m-d H:i');
    shinsa_put_rounds($rounds);
    return $counts;
}

/** 学校ごとの申込状況（申込の無い学校も出す） */
function shinsa_status(string $rid): array
{
    $by = [];
    foreach (shinsa_round_apps($rid) as $a) {
        $sid = $a['school_id'];
        // 級位審査は区分と男女で、形講習・レベルチェックは区分だけで数える
        $k = shinsa_has_fee($a) ? ($a['kubun'] . '_' . (($a['sex'] ?? '') === '女子' ? 'F' : 'M')) : (string)$a['kubun'];
        $by[$sid][$k] = ($by[$sid][$k] ?? 0) + 1;
        $by[$sid]['total'] = ($by[$sid]['total'] ?? 0) + 1;
        $by[$sid]['numbered'] = ($by[$sid]['numbered'] ?? 0) + ((string)($a['no'] ?? '') !== '' ? 1 : 0);
        $by[$sid]['to_number'] = ($by[$sid]['to_number'] ?? 0) + (shinsa_needs_number($a) ? 1 : 0);
        if (in_array($a['result'] ?? '', ['pass', 'fail'], true)) {
            $by[$sid]['receipts'] = ($by[$sid]['receipts'] ?? 0) + 1;
        }
    }
    $schools = db()->query('SELECT school_id, base_name, school_no FROM schools')->fetchAll();
    usort($schools, fn($a, $b) => [(int)$a['school_no'], $a['base_name']] <=> [(int)$b['school_no'], $b['base_name']]);
    $rows = [];
    foreach ($schools as $s) {
        $rows[] = ['school_id' => (string)$s['school_id'], 'name' => (string)$s['base_name']]
                + ($by[(string)$s['school_id']] ?? []);
    }
    return $rows;
}

/**
 * 顧問が領収書のために合否を選ぶ（審査のあと、管理画面で領収書を ON にしてから）。
 * 合格 pass ／不合格 fail ／欠席 absent ／未選択 ''。級位審査（上級・初級）の人だけ
 * （形講習・レベルチェックは無料で領収書なし）。
 */
function shinsa_save_results(array $round, array $school, array $results): array
{
    if (empty($round['receipt_on'])) {
        return ['領収書はまだ出せません（審査のあとに審査部会が出せるようにします）'];
    }
    $pdo = shinsa_db();
    $apps = shinsa_apps_of((string)$round['id'], (string)$school['school_id']);
    foreach ($results as $name => $v) {
        if (!in_array((string)$v, ['pass', 'fail', 'absent', ''], true)) {
            return ["「{$name}」: 合格・不合格・欠席から選んでください"];
        }
    }
    $put = $pdo->prepare('UPDATE shinsa_apps SET data_json = ? WHERE round_id = ? AND school_id = ? AND name = ?');
    $pdo->beginTransaction();
    foreach ($apps as $a) {
        if (!shinsa_has_fee($a) || !array_key_exists($a['name'], $results)) {
            continue;
        }
        $name = $a['name'];
        unset($a['name']);
        $a['result'] = (string)$results[$name];
        $put->execute([json_encode($a, JSON_UNESCAPED_UNICODE), $round['id'], $school['school_id'], $name]);
    }
    $pdo->commit();
    return [];
}

/** 受審番号を振る申込か（形講習会のみは振らない） */
function shinsa_needs_number(array $a): bool
{
    foreach (shinsa_groups() as $g) {
        if ($g['kubun'] === ($a['kubun'] ?? '')) {
            return $g['number'];
        }
    }
    return false;
}

/** 受審料がかかる（＝領収書を出す）申込か。級位審査の上級・初級だけ */
function shinsa_has_fee(array $a): bool
{
    return !empty(shinsa_kubun()[$a['kubun'] ?? '']['fee']);
}
