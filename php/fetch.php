<?php
declare(strict_types=1);

$config = require __DIR__ . '/config.php';
require __DIR__ . '/helpers.php';

$dataDir = $config['DATA_DIR'] ?? (__DIR__ . '/data');
ensure_directory($dataDir);

$settingsResult = fetch_settings_payload(
    $config['SETTINGS_URL'],
    $config['SETTINGS_SHEET_NAME'] ?? '給与計算_設定',
    $dataDir
);

$targets = [];

foreach ($config['STORES'] as $key => $store) {
    $targets[$key] = [
        'url' => $store['url'],
        'format' => 'xlsx',
        'label' => $store['name'] ?? $key,
    ];
}

$results = [];
foreach ($targets as $key => $info) {
    $results[] = fetch_and_cache_sheet($key, $info['url'], $dataDir, $info['format']);
}

$workbookCache = [];
foreach ($results as $result) {
    if (!empty($result['sourceUrl']) && !empty($result['filename'])) {
        $workbookCache[$result['sourceUrl']] = 'php/data/' . $result['filename'];
    }
}

if (!empty($settingsResult['record'])) {
    $settingsResult['record']['workbookCache'] = $workbookCache;
    $settingsJson = json_encode($settingsResult['record'], JSON_PRETTY_PRINT | JSON_UNESCAPED_UNICODE);
    file_put_contents($dataDir . '/settings.json', $settingsJson);
    $settingsResult['jsonPath'] = $dataDir . '/settings.json';
}

$index = [
    'generatedAt' => gmdate('c'),
    'dataDir' => $dataDir,
    'items' => $results,
    'settings' => [
        'status' => $settingsResult['status'] ?? 'unknown',
        'path' => $settingsResult['path'] ?? null,
        'jsonPath' => $settingsResult['jsonPath'] ?? null,
        'message' => $settingsResult['message'] ?? null,
    ],
];

file_put_contents($dataDir . '/cache-index.json', json_encode($index, JSON_PRETTY_PRINT | JSON_UNESCAPED_UNICODE));

$asJson = php_sapi_name() !== 'cli';
if (!$asJson && isset($argv) && in_array('--json', $argv, true)) {
    $asJson = true;
}

if ($asJson) {
    header('Content-Type: application/json; charset=UTF-8');
    echo json_encode($index, JSON_PRETTY_PRINT | JSON_UNESCAPED_UNICODE);
    exit;
}

if (!empty($settingsResult['status'])) {
    $settingsLine = sprintf('[%s] settings', strtoupper((string) $settingsResult['status']));
    if (!empty($settingsResult['jsonPath'])) {
        $settingsLine .= ' (json: ' . $settingsResult['jsonPath'] . ')';
    } elseif (!empty($settingsResult['path'])) {
        $settingsLine .= ' (' . $settingsResult['path'] . ')';
    }
    if (!empty($settingsResult['message'])) {
        $settingsLine .= ' - ' . $settingsResult['message'];
    }
    echo $settingsLine . PHP_EOL;
}

foreach ($results as $result) {
    $label = $result['key'] ?? 'unknown';
    $status = $result['status'] ?? 'unknown';
    $line = sprintf('[%s] %s', strtoupper((string) $status), $label);
    if (isset($result['message'])) {
        $line .= ' - ' . $result['message'];
    }
    if (isset($result['path'])) {
        $line .= ' (' . $result['path'] . ')';
    }
    echo $line . PHP_EOL;
}
