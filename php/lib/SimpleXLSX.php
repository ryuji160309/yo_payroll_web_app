<?php
/**
 * シンプルなXLSXリーダー（外部依存なし）
 * 必要最低限の文字列・数値読み取りに対応
 */
class SimpleXLSX
{
    private array $sheets = [];
    private array $sheetNames = [];
    private static string $lastError = '';

    public static function parse(string $filename): ?self
    {
        $xlsx = new self();
        if ($xlsx->load($filename)) {
            return $xlsx;
        }
        return null;
    }

    public static function parseError(): string
    {
        return self::$lastError;
    }

    public function sheetNames(): array
    {
        return $this->sheetNames;
    }

    public function rows(int $sheetIndex = 0): array
    {
        return $this->sheets[$sheetIndex] ?? [];
    }

    private function setError(string $message): void
    {
        self::$lastError = $message;
    }

    private function load(string $filename): bool
    {
        if (!is_readable($filename)) {
            $this->setError('ファイルを読み込めません: ' . $filename);
            return false;
        }

        $zip = new ZipArchive();
        if ($zip->open($filename) !== true) {
            $this->setError('Zipの展開に失敗しました。');
            return false;
        }

        $workbookXml = $zip->getFromName('xl/workbook.xml');
        if ($workbookXml === false) {
            $this->setError('workbook.xmlが見つかりません。');
            $zip->close();
            return false;
        }

        $relsXml = $zip->getFromName('xl/_rels/workbook.xml.rels');
        if ($relsXml === false) {
            $this->setError('workbook.xml.relsが見つかりません。');
            $zip->close();
            return false;
        }

        $workbook = simplexml_load_string($workbookXml);
        $rels = simplexml_load_string($relsXml);
        if (!$workbook || !$rels) {
            $this->setError('workbookの解析に失敗しました。');
            $zip->close();
            return false;
        }

        $relationships = [];
        foreach ($rels->Relationship as $rel) {
            $rid = (string) $rel['Id'];
            $target = (string) $rel['Target'];
            $relationships[$rid] = $target;
        }

        $sharedStrings = $this->loadSharedStrings($zip);

        $this->sheetNames = [];
        foreach ($workbook->sheets->sheet as $sheet) {
            $this->sheetNames[] = (string) $sheet['name'];
        }

        $this->sheets = [];
        $sheetIndex = 0;
        foreach ($workbook->sheets->sheet as $sheet) {
            $rid = (string) $sheet['r:id'];
            $path = $relationships[$rid] ?? null;
            if (!$path) {
                $this->sheets[$sheetIndex] = [];
                $sheetIndex++;
                continue;
            }
            $sheetPath = str_starts_with($path, '/') ? substr($path, 1) : 'xl/' . $path;
            $sheetXml = $zip->getFromName($sheetPath);
            if ($sheetXml === false) {
                $this->sheets[$sheetIndex] = [];
                $sheetIndex++;
                continue;
            }
            $this->sheets[$sheetIndex] = $this->parseSheet($sheetXml, $sharedStrings);
            $sheetIndex++;
        }

        $zip->close();
        return true;
    }

    private function loadSharedStrings(ZipArchive $zip): array
    {
        $shared = [];
        $xml = $zip->getFromName('xl/sharedStrings.xml');
        if ($xml === false) {
            return $shared;
        }
        $data = simplexml_load_string($xml);
        if (!$data) {
            return $shared;
        }
        foreach ($data->si as $si) {
            $texts = [];
            if (isset($si->t)) {
                $texts[] = (string) $si->t;
            }
            if (isset($si->r)) {
                foreach ($si->r as $run) {
                    if (isset($run->t)) {
                        $texts[] = (string) $run->t;
                    }
                }
            }
            $shared[] = implode('', $texts);
        }
        return $shared;
    }

    private function parseSheet(string $xml, array $sharedStrings): array
    {
        $sheet = simplexml_load_string($xml);
        if (!$sheet || !isset($sheet->sheetData)) {
            return [];
        }

        $rows = [];
        foreach ($sheet->sheetData->row as $rowNode) {
            $rowIndex = ((int) $rowNode['r']) - 1;
            $row = [];
            foreach ($rowNode->c as $cell) {
                $ref = (string) $cell['r'];
                $colIndex = $this->columnIndexFromRef($ref);
                $row[$colIndex] = $this->extractCellValue($cell, $sharedStrings);
            }
            if (!empty($row)) {
                $maxCol = max(array_keys($row));
                for ($i = 0; $i <= $maxCol; $i++) {
                    if (!array_key_exists($i, $row)) {
                        $row[$i] = '';
                    }
                }
                ksort($row);
            }
            $rows[$rowIndex] = $row;
        }
        ksort($rows);
        return array_values($rows);
    }

    private function columnIndexFromRef(string $cellRef): int
    {
        if (preg_match('/([A-Z]+)\d+/i', $cellRef, $m)) {
            $letters = strtoupper($m[1]);
            $index = 0;
            for ($i = 0; $i < strlen($letters); $i++) {
                $index = $index * 26 + (ord($letters[$i]) - ord('A') + 1);
            }
            return $index - 1;
        }
        return 0;
    }

    private function extractCellValue(SimpleXMLElement $cell, array $sharedStrings)
    {
        $type = (string) $cell['t'];
        if ($type === 's') {
            $idx = (int) $cell->v;
            return $sharedStrings[$idx] ?? '';
        }
        if ($type === 'inlineStr' && isset($cell->is->t)) {
            return (string) $cell->is->t;
        }
        if ($type === 'b') {
            return ((string) $cell->v) === '1' ? true : false;
        }
        if (!isset($cell->v)) {
            return '';
        }
        $value = (string) $cell->v;
        if ($value === '') {
            return '';
        }
        if (is_numeric($value)) {
            return $value + 0; // cast to int/float
        }
        return $value;
    }
}
