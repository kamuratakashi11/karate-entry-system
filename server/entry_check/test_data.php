<?php
/**
 * 動作確認用の架空データ（実名なし）。
 * 22の種目列すべてに1回以上マークが入り、2ブロック目（23人目以降）も使うように
 * 組んである。check.php（さくら上）と cli_test.php（手元）の両方がこれを読む。
 */

$mk = function (int $i, string $sex, array $extra) {
    return array_merge([
        'name'          => sprintf('テスト%s%02d', $sex === '男子' ? '太郎' : '花子', $i),
        'grade'         => ($i % 2) + 1,
        'dob'           => sprintf('2010/%d/%d', ($i % 12) + 1, ($i % 27) + 1),
        'jkf_no'        => sprintf('%07d', 1000000 + $i),
        'sex'           => $sex,
        'display_order' => $i,
    ], $extra);
};

$members = [
    // 男子: 団体形・団体組手(5人制)・個人形・個人組手5階級
    $mk(1, '男子', ['team_kata_chk' => 1, 'kata_chk' => 1, 'kata_val' => '正', 'kata_rank' => '1']),
    $mk(2, '男子', ['team_kata_chk' => 1, 'team_kata_role' => '補', 'team_kumi_chk' => 1]),
    $mk(3, '男子', ['team_kumi_chk' => 1, 'team_kumi_role' => '補', 'kata_chk' => 1, 'kata_val' => 'シード', 'kata_rank' => '2']),
    $mk(4, '男子', ['kata_chk' => 1, 'kata_val' => '補', 'kumi_chk' => 1, 'kumi_val' => '-55kg級', 'kumi_sub_val' => '正', 'kumi_rank' => '1']),
    $mk(5, '男子', ['kumi_chk' => 1, 'kumi_val' => '-61kg級', 'kumi_sub_val' => 'シード', 'kumi_rank' => '3']),
    $mk(6, '男子', ['kumi_chk' => 1, 'kumi_val' => '-68kg級', 'kumi_sub_val' => '補']),
    $mk(7, '男子', ['kumi_chk' => 1, 'kumi_val' => '-76kg級']),
    $mk(8, '男子', ['kumi_chk' => 1, 'kumi_val' => '+76kg級', 'kumi_rank' => '2']),
    // 女子: 団体組手は3人制の列に入ることを検証する（w_kumite_mode='3'）
    $mk(9, '女子', ['team_kata_chk' => 1, 'kata_chk' => 1, 'kata_val' => '正', 'kata_rank' => '']),
    $mk(10, '女子', ['team_kata_chk' => 1, 'team_kata_role' => '補', 'team_kumi_chk' => 1]),
    $mk(11, '女子', ['team_kumi_chk' => 1, 'team_kumi_role' => '補', 'kata_chk' => 1, 'kata_val' => 'シード', 'kata_rank' => '1']),
    $mk(12, '女子', ['kata_chk' => 1, 'kata_val' => '補', 'kumi_chk' => 1, 'kumi_val' => '-48kg級']),
    $mk(13, '女子', ['kumi_chk' => 1, 'kumi_val' => '-53kg級', 'kumi_sub_val' => 'シード', 'kumi_rank' => '2']),
    $mk(14, '女子', ['kumi_chk' => 1, 'kumi_val' => '-59kg級', 'kumi_sub_val' => '補']),
    $mk(15, '女子', ['kumi_chk' => 1, 'kumi_val' => '-66kg級', 'kumi_rank' => '1']),
    $mk(16, '女子', ['kumi_chk' => 1, 'kumi_val' => '+66kg級']),
];
// 2ブロック目（23人目・24人目が start_row+offset の行に入る）まで埋める
for ($i = 17; $i <= 24; $i++) {
    $members[] = $mk($i, ($i % 2 === 1) ? '男子' : '女子', []);
}

return [
    'school' => [
        'year'            => '8年',   // 「年」付きで渡して「8」になることを検証
        'tournament_name' => '新人大会（動作確認）',
        'date'            => '令和8年8月1日',
        'school_name'     => '確認用テスト高校',
        'principal'       => '確認 校長',
        'm_kumite_mode'   => '5',
        'w_kumite_mode'   => '3',
        'advisors'        => [
            ['name' => '顧問 一郎', 'd1' => true,  'd2' => false],
            ['name' => '顧問 二郎', 'd1' => false, 'd2' => true],
            ['name' => '顧問 三子', 'd1' => true,  'd2' => true],
            ['name' => '顧問 四子', 'd1' => false, 'd2' => false],
        ],
    ],
    'members' => $members,
];
