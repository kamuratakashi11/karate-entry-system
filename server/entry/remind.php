<?php
/**
 * 締め切りの確認メール（孝さん 2026-10-13「締め切りが近くなると自動で確認のメールが送れるといい」。
 * 10-13 に決めたこと: **締め切りの 2 週間前**に、**大会申込に登録されている学校すべて**へ送る。
 * アドレスは孝さんが持っているので、専門部の管理画面「メール」で学校ごとに入れる）。
 *
 * 送るもの: 「これからの締め切り」（shinsa.php の upcoming_deadlines。大会と級位審査）のうち、
 *   受付中（大会は受付中の大会・審査は学校に見せている回）で、締め切りまで REMIND_DAYS 日以内のもの。
 *   同じ締め切りには 1 回だけ（送った記録 config の mail_sent）。毎日 1 回、さくらの cron が
 *   remind_cron.php を動かす。動かし忘れた日があっても、締め切りの日までに動けば送る。
 *
 * 送り方: PHP の mail()（さくらのレンタルサーバーの sendmail）。送り主はさくらのドメインの
 *   アドレスにする（他のドメインを名乗ると迷惑メール扱いになりやすい）。返信先は孝さんのアドレス。
 *   手元で試すときは環境変数 ENTRY_MAIL_DIR を付けると、送らずにそのフォルダへ .eml を書く。
 */
declare(strict_types=1);

require_once __DIR__ . '/shinsa.php';

const REMIND_DAYS = 14;

/** メールの設定（専門部の管理画面で入れる） */
function mail_settings(): array
{
    return (array)config_get('mail_settings', []) + [
        'from'     => '',   // 送り主（さくらのドメインのアドレス）
        'reply_to' => '',   // 返信先（孝さん）
        'sign'     => '埼玉県高等学校体育連盟 空手道専門部',
        'url'      => 'https://saitamahs-karate.sakura.ne.jp/entry/',
    ];
}

/** 学校ごとのアドレス（school_id → [アドレス…]。1 校に何人でも） */
function school_emails(): array
{
    $out = [];
    foreach ((array)config_get('school_emails', []) as $sid => $list) {
        $out[(string)$sid] = array_values(array_filter(array_map('trim', (array)$list)));
    }
    return $out;
}

/** 「a@x, b@y」のような書き方を分けて確かめる。@return array{0: list<string>, 1: list<string>} [正しい, おかしい] */
function split_emails(string $text): array
{
    $ok = $bad = [];
    foreach (preg_split('/[\s,，、;；]+/u', $text) ?: [] as $a) {
        $a = trim($a);
        if ($a === '') {
            continue;
        }
        filter_var($a, FILTER_VALIDATE_EMAIL) ? $ok[] = $a : $bad[] = $a;
    }
    return [array_values(array_unique($ok)), $bad];
}

/** 締め切りごとの記録の鍵（同じ締め切りに 2 回送らない） */
function remind_key(array $d): string
{
    return $d['kind'] . '|' . $d['name'] . '|' . $d['deadline'];
}

function jp_date(string $ymd, bool $withYear = false): string
{
    $t = strtotime($ymd);
    if ($t === false) {
        return $ymd;
    }
    $w = ['日', '月', '火', '水', '木', '金', '土'][(int)date('w', $t)];
    return ($withYear ? date('Y', $t) . '年' : '') . (int)date('n', $t) . '月' . (int)date('j', $t) . "日（{$w}）";
}

/**
 * これから送る予定と、送ったもの（管理画面に出す）。
 * @return list<array{kind:string,name:string,deadline:string,send_on:string,open:bool,sent:?string,due:bool}>
 */
function remind_plan(string $today): array
{
    $sent = (array)config_get('mail_sent', []);
    $out = [];
    foreach (upcoming_deadlines($today) as $d) {
        $sendOn = date('Y-m-d', strtotime($d['deadline'] . ' -' . REMIND_DAYS . ' days'));
        $k = remind_key($d);
        $out[] = $d + [
            'send_on' => $sendOn,
            'sent'    => isset($sent[$k]) ? (string)$sent[$k] : null,
            // 送る日になっていて、受付中で、まだ送っていない
            'due'     => $today >= $sendOn && $today <= $d['deadline'] && !empty($d['open']) && !isset($sent[$k]),
        ];
    }
    return $out;
}

/** この学校の、その締め切りについてのひとこと（級位審査はいまの申込の人数） */
function remind_status_line(array $d, array $school): string
{
    if ($d['kind'] !== '級位審査') {
        return '';
    }
    $r = shinsa_open_round();
    if (!$r) {
        return '';
    }
    $n = count(shinsa_apps_of((string)$r['id'], (string)$school['school_id']));
    return $n ? "いま保存されている申込: {$n}名（直す場合は締め切りまでに）" : 'いまのところ、申込はまだ保存されていません。';
}

/** @return array{0:string,1:string} [件名, 本文] */
function remind_message(array $d, array $school, array $set, int $daysLeft): array
{
    $what = $d['kind'] === '級位審査'
        ? $d['name'] . ((string)($d['date'] ?? '') !== '' ? '（審査期日 ' . jp_date((string)$d['date']) . '）' : '')
        : $d['name'];
    $subject = '【締め切りの確認】' . $d['name'] . 'の申込は' . jp_date($d['deadline']) . 'まで';
    $status = remind_status_line($d, $school);
    $lines = [
        shinsa_school_label($school) . '高等学校 空手道部 顧問の先生',
        '',
        $set['sign'] . 'です。',
        "{$what}の申込の締め切りは、" . jp_date($d['deadline'], true) . "です（あと{$daysLeft}日）。",
    ];
    if ($status !== '') {
        $lines[] = $status;
    }
    $lines = array_merge($lines, [
        '',
        '申込は、いつもの申込システムからお願いします（ログインのあとのメニューで選びます）。',
        $set['url'],
        '',
        'このメールは、締め切りが近づいたとき（' . REMIND_DAYS . '日前から）に、登録されているすべての学校へ自動でお送りしています。',
        'すでにお済みの場合や、行き違いの場合はご容赦ください。',
        '',
        '――',
        $set['sign'],
    ]);
    return [$subject, implode("\n", $lines) . "\n"];
}

/**
 * 1 通送る。ENTRY_MAIL_DIR があれば送らずに .eml を書く（手元の確かめ用）。
 */
function send_mail(string $to, string $subject, string $body, array $set): bool
{
    mb_language('uni');
    mb_internal_encoding('UTF-8');
    $from = (string)$set['from'];
    $headers = [
        'From: ' . mb_encode_mimeheader($set['sign'], 'UTF-8') . " <{$from}>",
        'MIME-Version: 1.0',
        'Content-Type: text/plain; charset=UTF-8',
        'Content-Transfer-Encoding: base64',
    ];
    if ($set['reply_to'] !== '') {
        $headers[] = 'Reply-To: ' . $set['reply_to'];
    }
    $encSubject = mb_encode_mimeheader($subject, 'UTF-8');
    $encBody = chunk_split(base64_encode($body));
    $dir = getenv('ENTRY_MAIL_DIR');
    if ($dir !== false && $dir !== '') {
        $name = $dir . '/' . date('Ymd_His') . '_' . bin2hex(random_bytes(3)) . '.eml';
        return file_put_contents($name, "To: {$to}\r\nSubject: {$encSubject}\r\n" . implode("\r\n", $headers)
            . "\r\n\r\n" . $encBody) !== false;
    }
    // -f で封筒の差出人も送り主にそろえる（さくらの推奨。そろえないと届きにくい）
    return mail($to, $encSubject, $encBody, implode("\r\n", $headers), '-f' . $from);
}

/**
 * 送る日が来たものを送る（cron と、管理画面の「いま送る」）。
 * @return array{sent: list<array>, errors: list<string>}
 */
function remind_run(string $today): array
{
    $set = mail_settings();
    if (!filter_var($set['from'], FILTER_VALIDATE_EMAIL)) {
        return ['sent' => [], 'errors' => ['送り主のアドレスが決まっていません（管理画面の「メール」で入れてください）']];
    }
    $emails = school_emails();
    $schools = db()->query('SELECT * FROM schools')->fetchAll();
    $sentLog = (array)config_get('mail_sent', []);
    $done = [];
    $errors = [];
    foreach (remind_plan($today) as $d) {
        if (!$d['due']) {
            continue;
        }
        $daysLeft = (int)round((strtotime($d['deadline']) - strtotime($today)) / 86400);
        $ok = $ng = $none = 0;
        foreach ($schools as $s) {
            $to = $emails[(string)$s['school_id']] ?? [];
            if (!$to) {
                $none++;
                continue;
            }
            [$subject, $body] = remind_message($d, $s, $set, $daysLeft);
            foreach ($to as $addr) {
                send_mail($addr, $subject, $body, $set) ? $ok++ : $ng++;
            }
        }
        if ($ng) {
            $errors[] = "{$d['name']}: {$ng} 通を送れませんでした";
        }
        if ($ok === 0 && $ng === 0) {
            // どの学校にもアドレスが無い。送ったことにせず、アドレスが入ってから送る
            $errors[] = "{$d['name']}: 学校のアドレスが 1 つも入っていないので送っていません";
            continue;
        }
        $sentLog[remind_key($d)] = date('Y-m-d H:i') . " {$ok}通" . ($ng ? "・失敗{$ng}" : '') . ($none ? "・アドレス無し{$none}校" : '');
        $done[] = $d + ['result' => $sentLog[remind_key($d)]];
    }
    if ($done) {
        db()->prepare('INSERT OR REPLACE INTO config VALUES (?,?)')
            ->execute(['mail_sent', json_encode($sentLog, JSON_UNESCAPED_UNICODE)]);
    }
    return ['sent' => $done, 'errors' => $errors];
}
