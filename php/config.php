<?php
return [
    // Google Sheetsの設定スプレッドシートURL。環境変数 SETTINGS_URL が設定されていればそちらを優先します。
    'settings_url' => 'https://docs.google.com/spreadsheets/d/e/2PACX-1vTKnnQY1d5BXnOstLwIhJOn7IX8aqHXC98XzreJoFscTUFPJXhef7jO2-0KKvZ7_fPF0uZwpbdcEpcV/pub?output=xlsx',

    // キャッシュ保存先のベースディレクトリ
    'cache_root' => __DIR__ . '/../php-cache',

    // ストア一覧を取得する設定シート名
    'settings_sheet_name' => '給与計算_設定',

    // 店舗一覧の開始行と列の指定（行番号は1始まり）
    'store_list_start_row' => 11,
    'store_name_column' => 'A',
    'store_url_column' => 'B',

    // HTTPダウンロードのタイムアウト（秒）
    'http_timeout' => 45,
];
