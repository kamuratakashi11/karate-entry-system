<?php
/**
 * 審査の当日に使う Excel（審査の管理画面から書き出す）。
 *
 * 会場はネットにつながらないことが多いので、出欠・点数・合否は**この Excel で手元で付ける**
 * （孝さん 2026-10-14「オンライン環境でない場合が多いので、エクセルで出力できるほうが良い」）。
 * 結果をシステムに読み込ませることはしない（領収書の合否は顧問が画面で選ぶ）。
 *
 * シートは、これまで手で作っていた受付の Excel（00.第1回まとめ.xlsx）に合わせた:
 *   受付名簿 … 学校ごと・申込書の順。出欠の欄つき
 *   上級集計・初級集計 … 受審番号順。点を入れると合計点と合格級が出る（基準は shinsa.php の SHINSA_PASS）
 *   受審順（掲示） … 組ごとの受審番号・氏名・学校
 *   形講習・レベルチェック … レベルチェックテスト（番号あり）と形講習会のみ（番号なし）の人
 */
declare(strict_types=1);

require_once __DIR__ . '/shinsa.php';
require_once __DIR__ . '/lib/xlsx_book.php';

/** 合計点のセル → 合格級の式（SHINSA_PASS から作る。空なら空、どれにも届かなければ不合格） */
function shinsa_pass_formula(string $kubun, string $cell): string
{
    $f = '"不合格"';
    foreach (array_reverse(SHINSA_PASS[$kubun]) as [$min, $kyu]) {
        $f = "IF({$cell}>={$min},\"{$kyu}\",{$f})";
    }
    return "IF({$cell}=\"\",\"\",{$f})";
}

function shinsa_pass_text(string $kubun): string
{
    $p = SHINSA_PASS[$kubun];
    $parts = [];
    foreach ($p as $i => [$min, $kyu]) {
        $parts[] = "{$kyu} {$min}点以上" . ($i > 0 ? "（{$p[$i - 1][0]}点未満）" : '');
    }
    return '基準: ' . implode('、', $parts) . '。これに届かなければ不合格';
}

function shinsa_kubun_text(array $a): string
{
    $k = shinsa_kubun()[$a['kubun']];
    return $k['fee'] ? (string)($a['sex'] ?? '') . $k['short'] : $k['short'];
}

function shinsa_period_text(array $a): string
{
    return (int)($a['years'] ?? 0) . '年' . (int)($a['months'] ?? 0) . 'か月';
}

/** 受審番号（まだ振っていなければ空） */
function shinsa_no(array $a): int|string
{
    return (string)($a['no'] ?? '') === '' ? '' : (int)$a['no'];
}

/** @return string 書き出したファイルのパス（$path） */
function shinsa_day_book(array $round, string $path): string
{
    $B = XlsxBook::class;
    $apps = shinsa_round_apps((string)$round['id']);
    $title = shinsa_round_title($round);
    [$ry, $rm, $rd] = shinsa_reiwa((string)($round['date'] ?? ''));
    $when = $ry !== '' ? "（審査期日 令和{$ry}年{$rm}月{$rd}日）" : '';
    $book = new XlsxBook();

    // 受付番号（受付名簿の通し番号）。集計表から受付名簿を引けるよう、どのシートにも出す
    $ukeNo = [];
    foreach ($apps as $i => $a) {
        $ukeNo[$a['school_id'] . "\t" . $a['name']] = $i + 1;
    }
    $uke = fn(array $a) => $ukeNo[$a['school_id'] . "\t" . $a['name']];

    // --- 受付名簿 ---
    $s = $book->sheet('受付名簿', [6, 13, 15, 5, 5, 17, 7, 8, 15, 10, 11, 7]);
    $s->landscape = true;
    $s->put(1, 1, "{$title}　受付名簿{$when}", $B::TITLE);
    $fee = number_format(SHINSA_FEE['pass']);
    $s->put(2, 1, "受審料: 上級・初級 {$fee}円（不合格は " . number_format(SHINSA_FEE['pass'] - SHINSA_FEE['fail'])
        . '円を返す）／形講習・レベルチェック 無料。出欠の欄に ○ を書き込む', $B::NOTE);
    $head = ['受付№', '学校名', '受審者氏名', '学年', '男女', '区分', '受審№', '現公認級', '演武形名', '修業期間', '会員番号', '出欠'];
    foreach ($head as $c => $h) {
        $s->put(3, $c + 1, $h, $B::HEAD);
    }
    $s->headerRows = '3:3';
    foreach ($apps as $i => $a) {
        $r = 4 + $i;
        $vals = [$i + 1, $a['school'], $a['name'], (int)$a['grade'], ($a['sex'] ?? '') === '女子' ? '女' : '男',
                 shinsa_kubun_text($a), shinsa_no($a), $a['kyu'], $a['kata'],
                 shinsa_period_text($a), (string)$a['jkf_no'], ''];
        foreach ($vals as $c => $v) {
            $s->put($r, $c + 1, $v, $c === 11 ? $B::INPUT : (in_array($c, [1, 2, 8], true) ? $B::LEFT : $B::CELL));
        }
    }

    $ordered = shinsa_ordered($apps);
    $groups = shinsa_groups();

    // --- 上級集計・初級集計（点を入れると合格級が出る） ---
    $score = ['jokyu' => ['指定形', '自由組手'], 'shokyu' => ['基本', '基本形', '約束組手']];
    foreach (['jokyu' => '上級集計', 'shokyu' => '初級集計'] as $kubun => $name) {
        $items = $score[$kubun];
        $n = count($items);
        $widths = array_merge([6, 5, 15, 13, 5, 6], array_fill(0, $n, 8), [8, 8, 8, 15, 10, 11]);
        $s = $book->sheet($name, $widths);
        $s->landscape = true;
        $s->put(1, 1, "{$title}（" . shinsa_kubun()[$kubun]['short'] . "）結果一覧{$when}", $B::TITLE);
        $s->put(2, 1, shinsa_pass_text($kubun) . '。黄色の欄に点を入れる', $B::NOTE);
        $head = array_merge(['受審№', '男女', '受審者氏名', '学校名', '学年', '受付№'], $items,
                            ['合計点', '現在級', '合格級', '演武形名', '修業期間', '会員番号']);
        foreach ($head as $c => $h) {
            $s->put(3, $c + 1, $h, $B::HEAD);
        }
        $s->headerRows = '3:3';
        $r = 4;
        foreach ($groups as $g) {
            if ($g['kubun'] !== $kubun) {
                continue;
            }
            foreach ($ordered[$g['key']] as $a) {
                $first = XlsxBook::col(7);
                $lastItem = XlsxBook::col(6 + $n);
                $sumCol = XlsxBook::col(7 + $n);
                $vals = [shinsa_no($a), ($a['sex'] ?? '') === '女子' ? '女' : '男',
                         $a['name'], $a['school'], (int)$a['grade'], $uke($a)];
                foreach ($vals as $c => $v) {
                    $s->put($r, $c + 1, $v, in_array($c, [2, 3], true) ? $B::LEFT : $B::CELL);
                }
                for ($k = 0; $k < $n; $k++) {
                    $s->put($r, 7 + $k, '', $B::INPUT);
                }
                $s->put($r, 7 + $n, '', $B::CELL, "IF(COUNT({$first}{$r}:{$lastItem}{$r})=0,\"\",SUM({$first}{$r}:{$lastItem}{$r}))");
                $s->put($r, 8 + $n, (string)$a['kyu'], $B::CELL);
                $s->put($r, 9 + $n, '', $B::CELL, shinsa_pass_formula($kubun, "{$sumCol}{$r}"));
                $s->put($r, 10 + $n, (string)$a['kata'], $B::LEFT);
                $s->put($r, 11 + $n, shinsa_period_text($a), $B::CELL);
                $s->put($r, 12 + $n, (string)$a['jkf_no'], $B::CELL);
                $r++;
            }
        }
    }

    // --- 受審順（掲示） ---
    $s = $book->sheet('受審順（掲示）', [7, 18, 16, 3, 7, 18, 16]);
    $s->put(1, 1, "{$title}　受審順{$when}", $B::TITLE);
    $r = 3;
    foreach ($groups as $g) {
        $list = $ordered[$g['key']];
        if (!$list || !$g['number']) {
            continue;
        }
        $s->put($r, 1, $g['label'] . '（' . count($list) . '名）', $B::TITLE);
        $r++;
        foreach (['受審№', '受審者氏名', '学校名'] as $c => $h) {
            $s->put($r, $c + 1, $h, $B::HEAD);
        }
        $r++;
        foreach ($list as $a) {
            $s->put($r, 1, shinsa_no($a), $B::CELL);
            $s->put($r, 2, $a['name'], $B::LEFT);
            $s->put($r, 3, $a['school'], $B::LEFT);
            $r++;
        }
        $r++;
    }

    // --- 形講習・レベルチェック（レベルチェックの人が先、形講習会のみの人はあと・番号なし） ---
    $free = array_merge($ordered['lc'], $ordered['koshu']);
    if ($free) {
        $s = $book->sheet('形講習・レベルチェック', [6, 12, 5, 15, 13, 5, 6, 9, 15, 8, 11]);
        $s->put(1, 1, "{$title}　形講習会・レベルチェックテスト{$when}", $B::TITLE);
        $s->put(2, 1, '無料（領収書なし）。レベルチェックの点は審査用紙 3 枚の合計。形講習会のみの人は審査用紙なし', $B::NOTE);
        foreach (['受審№', '区分', '男女', '氏名', '学校名', '学年', '受付№', '現級段位', '指定形', '点', '会員番号'] as $c => $h) {
            $s->put(3, $c + 1, $h, $B::HEAD);
        }
        $s->headerRows = '3:3';
        foreach ($free as $i => $a) {
            $r = 4 + $i;
            $test = shinsa_needs_number($a);
            $vals = [shinsa_no($a), shinsa_kubun_text($a), ($a['sex'] ?? '') === '女子' ? '女' : '男', $a['name'],
                     $a['school'], (int)$a['grade'], $uke($a), $a['kyu'], $a['kata'], '', (string)$a['jkf_no']];
            foreach ($vals as $c => $v) {
                $style = $c === 9 ? ($test ? $B::INPUT : $B::CELL) : (in_array($c, [3, 4, 8], true) ? $B::LEFT : $B::CELL);
                $s->put($r, $c + 1, $v, $style);
            }
        }
    }

    $book->save($path);
    return $path;
}
