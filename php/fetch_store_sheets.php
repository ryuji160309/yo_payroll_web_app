<?php
require_once __DIR__ . '/lib/helpers.php';

$config = load_config();
try {
    $settingsPath = ensure_settings_workbook($config);
    $stores = parse_store_rows($settingsPath, $config);
    if (empty($stores)) {
        throw new RuntimeException('設定シートから店舗URLを取得できませんでした。');
    }

    $results = [];
    foreach ($stores as $store) {
        $results[] = cache_store_workbook($config, $store);
    }

    echo "店舗シートを更新しました (" . count($results) . "件)。\n";
    foreach ($results as $item) {
        $time = date('Y-m-d H:i:s', (int) floor($item['fetchedAt'] / 1000));
        echo "- {$item['store']['name']} => {$item['jsonPath']} ({$time})\n";
    }
} catch (Throwable $e) {
    fwrite(STDERR, "店舗シートの更新に失敗しました: " . $e->getMessage() . "\n");
    exit(1);
}
