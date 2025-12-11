<?php
/**
 * 共通ヘルパー関数
 */

function load_config(): array
{
    $config = require __DIR__ . '/../config.php';
    if ($envUrl = getenv('SETTINGS_URL')) {
        $config['settings_url'] = $envUrl;
    }
    return $config;
}

function ensure_directory(string $path): void
{
    if (is_dir($path)) {
        return;
    }
    if (!mkdir($path, 0775, true) && !is_dir($path)) {
        throw new RuntimeException("ディレクトリを作成できませんでした: {$path}");
    }
}

function write_json(string $path, array $payload): void
{
    $json = json_encode($payload, JSON_UNESCAPED_UNICODE | JSON_UNESCAPED_SLASHES | JSON_PRETTY_PRINT);
    if (false === file_put_contents($path, $json)) {
        throw new RuntimeException("JSONの書き込みに失敗しました: {$path}");
    }
}

function write_binary(string $path, string $contents): void
{
    if (false === file_put_contents($path, $contents)) {
        throw new RuntimeException("バイナリの書き込みに失敗しました: {$path}");
    }
}

function extract_file_id(string $url): ?string
{
    if (preg_match('#/d/([a-zA-Z0-9_-]+)#', $url, $m)) {
        return $m[1];
    }
    return null;
}

function to_xlsx_export_url(string $url): string
{
    // すでにxlsx指定がある場合はそのまま返す
    if (strpos($url, 'output=xlsx') !== false || strpos($url, 'format=xlsx') !== false) {
        return $url;
    }
    $fileId = extract_file_id($url);
    if ($fileId) {
        $params = ['id' => $fileId, 'format' => 'xlsx'];
        if (preg_match('/[?&]gid=(\d+)/', $url, $gidMatch)) {
            $params['gid'] = $gidMatch[1];
        }
        return sprintf('https://docs.google.com/spreadsheets/d/%s/export?%s', $fileId, http_build_query($params));
    }
    return $url;
}

function slugify(string $value): string
{
    $normalized = preg_replace('/[^a-zA-Z0-9]+/u', '-', $value);
    $normalized = trim($normalized ?? '', '-');
    return $normalized ?: 'sheet';
}

function fetch_workbook_binary(string $url, int $timeoutSeconds = 45): array
{
    $targetUrl = to_xlsx_export_url($url);
    $ch = curl_init($targetUrl);
    curl_setopt_array($ch, [
        CURLOPT_RETURNTRANSFER => true,
        CURLOPT_FOLLOWLOCATION => true,
        CURLOPT_HEADER => true,
        CURLOPT_USERAGENT => 'yo-payroll-php-cache/1.0',
        CURLOPT_CONNECTTIMEOUT => 10,
        CURLOPT_TIMEOUT => $timeoutSeconds,
    ]);

    $response = curl_exec($ch);
    if ($response === false) {
        $error = curl_error($ch);
        curl_close($ch);
        throw new RuntimeException($error ?: 'ダウンロードに失敗しました。');
    }

    $headerSize = curl_getinfo($ch, CURLINFO_HEADER_SIZE);
    $statusCode = curl_getinfo($ch, CURLINFO_HTTP_CODE);
    $contentType = curl_getinfo($ch, CURLINFO_CONTENT_TYPE) ?: 'application/octet-stream';
    $filetime = curl_getinfo($ch, CURLINFO_FILETIME);
    curl_close($ch);

    $headersRaw = substr($response, 0, $headerSize);
    $body = substr($response, $headerSize);

    if ($statusCode < 200 || $statusCode >= 300 || !$body) {
        throw new RuntimeException("HTTP {$statusCode} でスプレッドシートを取得できませんでした: {$targetUrl}");
    }

    $fetchedAt = $filetime && $filetime > 0 ? ($filetime * 1000) : (int) round(microtime(true) * 1000);

    return [
        'body' => $body,
        'fetchedAt' => $fetchedAt,
        'contentType' => $contentType,
        'sourceUrl' => $targetUrl,
        'headers' => $headersRaw,
    ];
}

function get_cache_paths(array $config): array
{
    $root = rtrim($config['cache_root'], '/');
    $paths = [
        'root' => $root,
        'settings_json' => $root . '/settings.json',
        'settings_xlsx' => $root . '/settings.xlsx',
        'sheets_dir' => $root . '/sheets',
    ];
    return $paths;
}

function save_settings_cache(array $config, array $download): array
{
    $paths = get_cache_paths($config);
    ensure_directory($paths['root']);
    ensure_directory($paths['sheets_dir']);

    $payload = [
        'source' => $download['sourceUrl'],
        'fetchedAt' => $download['fetchedAt'],
        'contentType' => $download['contentType'],
        'base64' => base64_encode($download['body']),
    ];

    write_json($paths['settings_json'], $payload);
    write_binary($paths['settings_xlsx'], $download['body']);

    return $paths;
}

function ensure_settings_workbook(array $config): string
{
    $paths = get_cache_paths($config);
    if (file_exists($paths['settings_xlsx'])) {
        return $paths['settings_xlsx'];
    }
    if (file_exists($paths['settings_json'])) {
        $json = json_decode((string) file_get_contents($paths['settings_json']), true);
        if (!empty($json['base64'])) {
            $binary = base64_decode($json['base64']);
            write_binary($paths['settings_xlsx'], $binary);
            return $paths['settings_xlsx'];
        }
    }

    $download = fetch_workbook_binary($config['settings_url'], (int) $config['http_timeout']);
    save_settings_cache($config, $download);
    return $paths['settings_xlsx'];
}

function column_letter_to_index(string $letter): int
{
    $letter = strtoupper(trim($letter));
    $index = 0;
    for ($i = 0; $i < strlen($letter); $i++) {
        $index = $index * 26 + (ord($letter[$i]) - ord('A') + 1);
    }
    return $index - 1; // zero-based
}

function parse_store_rows(string $settingsXlsxPath, array $config): array
{
    require_once __DIR__ . '/SimpleXLSX.php';

    $xlsx = SimpleXLSX::parse($settingsXlsxPath);
    if (!$xlsx) {
        throw new RuntimeException('設定スプレッドシートのパースに失敗しました: ' . SimpleXLSX::parseError());
    }

    $sheetIndex = 0;
    $targetSheetName = $config['settings_sheet_name'] ?? '';
    if ($targetSheetName) {
        foreach ($xlsx->sheetNames() as $idx => $name) {
            if ($name === $targetSheetName) {
                $sheetIndex = $idx;
                break;
            }
        }
    }

    $rows = $xlsx->rows($sheetIndex);
    $startRow = max(0, (int) ($config['store_list_start_row'] ?? 11) - 1);
    $nameIndex = column_letter_to_index($config['store_name_column'] ?? 'A');
    $urlIndex = column_letter_to_index($config['store_url_column'] ?? 'B');

    $stores = [];
    for ($r = $startRow; $r < count($rows); $r++) {
        $row = $rows[$r];
        $name = isset($row[$nameIndex]) ? trim((string) $row[$nameIndex]) : '';
        $url = isset($row[$urlIndex]) ? trim((string) $row[$urlIndex]) : '';
        if ($name === '' && $url === '') {
            if (!empty($stores)) {
                break; // 末尾の空行
            }
            continue;
        }
        if ($name === '' || $url === '') {
            continue;
        }
        $stores[] = ['name' => $name, 'url' => $url];
    }

    return $stores;
}

function cache_store_workbook(array $config, array $store): array
{
    $paths = get_cache_paths($config);
    ensure_directory($paths['sheets_dir']);

    $download = fetch_workbook_binary($store['url'], (int) $config['http_timeout']);
    $fileId = extract_file_id($store['url']) ?? slugify($store['name']);
    $jsonPath = sprintf('%s/%s.json', $paths['sheets_dir'], $fileId);
    $xlsxPath = sprintf('%s/%s.xlsx', $paths['sheets_dir'], $fileId);

    $payload = [
        'storeName' => $store['name'],
        'source' => $download['sourceUrl'],
        'fetchedAt' => $download['fetchedAt'],
        'contentType' => $download['contentType'],
        'base64' => base64_encode($download['body']),
    ];

    write_json($jsonPath, $payload);
    write_binary($xlsxPath, $download['body']);

    return [
        'store' => $store,
        'jsonPath' => $jsonPath,
        'xlsxPath' => $xlsxPath,
        'fetchedAt' => $download['fetchedAt'],
        'sourceUrl' => $download['sourceUrl'],
    ];
}
