<?php
declare(strict_types=1);

return [
    // Google Sheets URL that hosts the shared settings sheet.
    'SETTINGS_URL' => 'https://docs.google.com/spreadsheets/d/e/2PACX-1vTKnnQY1d5BXnOstLwIhJOn7IX8aqHXC98XzreJoFscTUFPJXhef7jO2-0KKvZ7_fPF0uZwpbdcEpcV/pub?output=xlsx',
    'SETTINGS_SHEET_NAME' => '給与計算_設定',

    // Directory to store cached files and generated JSON.
    'DATA_DIR' => __DIR__ . '/data',

    // Spreadsheet definitions for each store. The key is used as the cache file name.
    'STORES' => [
        'night' => [
            'name' => '夜勤',
            'url' => 'https://docs.google.com/spreadsheets/d/1gCGyxiXXxOOhgHG2tk3BlzMpXuaWQULacySlIhhoWRY/edit?gid=601593061#gid=601593061',
        ],
        'sagamihara_higashi' => [
            'name' => '相模原東大沼店',
            'url' => 'https://docs.google.com/spreadsheets/d/1fEMEasqSGU30DuvCx6O6D0nJ5j6m6WrMkGTAaSQuqBY/edit?gid=358413717#gid=358413717',
        ],
        'kobuchi' => [
            'name' => '古淵駅前店',
            'url' => 'https://docs.google.com/spreadsheets/d/1hSD3sdIQftusWcNegZnGbCtJmByZhzpAvLJegDoJckQ/edit?gid=946573079#gid=946573079',
        ],
        'hashimoto' => [
            'name' => '相模原橋本五丁目店',
            'url' => 'https://docs.google.com/spreadsheets/d/1YYvWZaF9Li_RHDLevvOm2ND8ASJ3864uHRkDAiWBEDc/edit?gid=2000770170#gid=2000770170',
        ],
        'isehara' => [
            'name' => '伊勢原高森七丁目店',
            'url' => 'https://docs.google.com/spreadsheets/d/1PfEQRnvHcKS5hJ6gkpJQc0VFjDoJUBhHl7JTTyJheZc/edit?gid=34390331#gid=34390331',
        ],
    ],
];
