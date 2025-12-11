<?php
require_once __DIR__ . '/lib/helpers.php';

$config = load_config();
try {
    $download = fetch_workbook_binary($config['settings_url'], (int) $config['http_timeout']);
    $paths = save_settings_cache($config, $download);
    $time = date('Y-m-d H:i:s', (int) floor($download['fetchedAt'] / 1000));
    echo "設定シートを取得しました: {$paths['settings_json']} ({$time})\n";
} catch (Throwable $e) {
    fwrite(STDERR, "設定シートの取得に失敗しました: " . $e->getMessage() . "\n");
    exit(1);
}
