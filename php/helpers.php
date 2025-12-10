<?php
declare(strict_types=1);

/**
 * Convert a Google Sheets URL into an export URL (xlsx or csv) that can be downloaded
 * without authentication.
 */
function to_export_url(string $url, string $format = 'xlsx'): ?string
{
    $format = strtolower($format) === 'csv' ? 'csv' : 'xlsx';

    if (!preg_match('#^https?://#i', $url)) {
        return null;
    }

    // Already an export URL with output=...
    if (strpos($url, 'output=') !== false) {
        $parsed = parse_url($url);
        $query = [];
        if (isset($parsed['query'])) {
            parse_str($parsed['query'], $query);
        }
        $query['output'] = $format;
        $parsed['query'] = http_build_query($query);
        return build_url($parsed);
    }

    if (preg_match('#/spreadsheets/d/([^/]+)/edit#i', $url, $m)) {
        $docId = $m[1];
        $gid = '0';
        if (preg_match('/[?&#]gid=([0-9]+)/', $url, $g)) {
            $gid = $g[1];
        }
        return "https://docs.google.com/spreadsheets/d/{$docId}/export?format={$format}&id={$docId}&gid={$gid}";
    }

    return null;
}

/**
 * Build URL from parse_url components.
 * @param array<string,mixed> $parts
 */
function build_url(array $parts): string
{
    $scheme   = $parts['scheme'] ?? 'https';
    $host     = $parts['host'] ?? '';
    $port     = isset($parts['port']) ? ':' . $parts['port'] : '';
    $path     = $parts['path'] ?? '';
    $query    = isset($parts['query']) ? '?' . $parts['query'] : '';
    $fragment = isset($parts['fragment']) ? '#' . $parts['fragment'] : '';
    return $scheme . '://' . $host . $port . $path . $query . $fragment;
}

/**
 * Ensure that a directory exists.
 */
function ensure_directory(string $path): void
{
    if (!is_dir($path)) {
        mkdir($path, 0775, true);
    }
}

/**
 * Download a spreadsheet and cache it locally.
 * Returns an associative array with status information.
 *
 * @param string $key       Identifier used for cache filename
 * @param string $url       Original Google Sheets URL
 * @param string $dataDir   Directory where files are saved
 * @param string $format    Either 'xlsx' (default) or 'csv'
 * @return array<string,mixed>
 */
function fetch_and_cache_sheet(string $key, string $url, string $dataDir, string $format = 'xlsx'): array
{
    $exportUrl = to_export_url($url, $format);
    if ($exportUrl === null) {
        return [
            'key' => $key,
            'status' => 'error',
            'message' => 'URL をエクスポート用に変換できませんでした',
        ];
    }

    ensure_directory($dataDir);

    $ext = $format === 'csv' ? 'csv' : 'xlsx';
    $destPath = rtrim($dataDir, '/\') . '/' . $key . '.' . $ext;

    $headers = [
        'User-Agent: yo-payroll-cache/1.0',
    ];
    if (file_exists($destPath)) {
        $mtime = filemtime($destPath);
        if ($mtime !== false) {
            $headers[] = 'If-Modified-Since: ' . gmdate('D, d M Y H:i:s', $mtime) . ' GMT';
        }
    }

    $ch = curl_init($exportUrl);
    curl_setopt_array($ch, [
        CURLOPT_RETURNTRANSFER => true,
        CURLOPT_FOLLOWLOCATION => true,
        CURLOPT_TIMEOUT => 60,
        CURLOPT_HTTPHEADER => $headers,
        CURLOPT_HEADER => true,
    ]);

    $response = curl_exec($ch);
    if ($response === false) {
        $error = curl_error($ch);
        curl_close($ch);
        return [
            'key' => $key,
            'status' => 'error',
            'message' => $error ?: 'curl error',
        ];
    }

    $headerSize = curl_getinfo($ch, CURLINFO_HEADER_SIZE) ?: 0;
    $statusCode = curl_getinfo($ch, CURLINFO_HTTP_CODE) ?: 0;
    $headersText = substr($response, 0, $headerSize);
    $body = substr($response, $headerSize);

    $lastModified = null;
    foreach (explode("\r\n", $headersText) as $line) {
        if (stripos($line, 'Last-Modified:') === 0) {
            $lastModified = trim(substr($line, strlen('Last-Modified:')));
            break;
        }
    }

    curl_close($ch);

    if ($statusCode === 304 && file_exists($destPath)) {
        return [
            'key' => $key,
            'status' => 'cached',
            'path' => $destPath,
            'filename' => basename($destPath),
            'exportUrl' => $exportUrl,
            'bytes' => filesize($destPath) ?: 0,
            'lastModified' => $lastModified,
            'sourceUrl' => $url,
        ];
    }

    if ($statusCode !== 200) {
        return [
            'key' => $key,
            'status' => 'error',
            'message' => 'HTTP ' . $statusCode . ' で失敗しました',
            'exportUrl' => $exportUrl,
            'sourceUrl' => $url,
        ];
    }

    file_put_contents($destPath, $body);
    if ($lastModified) {
        $timestamp = strtotime($lastModified);
        if ($timestamp !== false) {
            touch($destPath, $timestamp);
        }
    }

    return [
        'key' => $key,
        'status' => 'updated',
        'path' => $destPath,
        'filename' => basename($destPath),
        'exportUrl' => $exportUrl,
        'bytes' => strlen($body),
        'lastModified' => $lastModified,
        'sourceUrl' => $url,
    ];
}

/**
 * Fetch and parse the settings sheet as CSV to produce JSON usable by the frontend.
 *
 * @return array<string,mixed>
 */
function fetch_settings_payload(string $url, string $sheetName, string $dataDir): array
{
    $csvResult = fetch_and_cache_sheet('settings', $url, $dataDir, 'csv');
    $csvPath = $csvResult['path'] ?? null;

    if (!$csvPath || !file_exists($csvPath)) {
        return array_merge($csvResult, ['record' => null]);
    }

    $rows = [];
    if (($handle = fopen($csvPath, 'rb')) !== false) {
        while (($row = fgetcsv($handle)) !== false) {
            $rows[] = $row;
        }
        fclose($handle);
    }

    $statusCell = trim((string)($rows[3][1] ?? ''));
    $baseWage = isset($rows[10][3]) ? floatval($rows[10][3]) : null;
    $overtime = isset($rows[10][5]) ? floatval($rows[10][5]) : null;
    $password = isset($rows[10][9]) ? (string)$rows[10][9] : null;

    $excludeWords = [];
    $excludeCount = 0;
    for ($r = 10; $r < count($rows); $r++) {
        $val = $rows[$r][7] ?? null;
        if ($val === null || $val === '') {
            if ($excludeCount > 0) {
                break;
            }
            continue;
        }
        $excludeWords[] = (string)$val;
        $excludeCount++;
    }

    $stores = [];
    $idx = 1;
    for ($r = 10; $r < count($rows); $r++) {
        $name = $rows[$r][0] ?? '';
        $urlCell = $rows[$r][1] ?? '';
        if ($name === '' && $urlCell === '') {
            if ($idx > 1) {
                break;
            }
            continue;
        }
        if ($name !== '' && $urlCell !== '') {
            $stores['store' . $idx] = [
                'name' => (string)$name,
                'url' => (string)$urlCell,
                'baseWage' => $baseWage,
                'overtime' => $overtime,
                'excludeWords' => $excludeWords,
            ];
            $idx++;
        }
    }

    $settingsError = strtoupper($statusCell) !== 'ALL_OK';
    $details = [];
    if ($settingsError) {
        $details[] = $statusCell !== '' ? $statusCell : 'ALL_OK ではありません';
        $cells = [
            ['row' => 4, 'col' => 1, 'label' => 'ステータス'],
            ['row' => 10, 'col' => 1, 'label' => 'URL設定'],
            ['row' => 10, 'col' => 3, 'label' => '基本時給設定'],
            ['row' => 10, 'col' => 5, 'label' => '時間外倍率設定'],
            ['row' => 10, 'col' => 9, 'label' => 'パスワード'],
        ];
        foreach ($cells as $cell) {
            $val = $rows[$cell['row']][$cell['col']] ?? null;
            if ($val !== null && $val !== '') {
                $details[] = $cell['label'] . '：' . $val;
            }
        }
    }

    $record = [
        'settingsSheetName' => $sheetName,
        'stores' => $stores,
        'password' => $password,
        'baseWage' => $baseWage,
        'overtime' => $overtime,
        'excludeWords' => $excludeWords,
        'settingsError' => $settingsError ?: null,
        'settingsErrorDetails' => $settingsError ? $details : null,
        'source' => 'php-cache',
        'fetchedAt' => gmdate('c'),
    ];

    return array_merge($csvResult, ['record' => $record]);
}
