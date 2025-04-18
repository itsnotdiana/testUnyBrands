<?php

require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;

$inputFile = 'OZON_2024_review_Boxraw.xlsx';
$outputFile = 'output.sql';

$spreadsheet = IOFactory::load($inputFile);
$sheet = $spreadsheet->getActiveSheet();
$rows = $sheet->toArray(null, true, true, true);

$headers = array_shift($rows);
$headers = array_flip($headers);

$fieldMap = [
    'UF_ACTIVE'         => ['required' => false, 'const' => 1, 'type' => 'int'],
    'UF_AGREE'          => ['required' => false, 'inputKey' => 'likes', 'type' => 'int'],
    'UF_DISAGREE'       => ['required' => false, 'inputKey' => 'unlikes', 'type' => 'int'],
    'UF_USER_ID'        => ['required' => false, 'const' => 'NULL', 'type' => 'int'],
    'UF_PHOTOS'         => ['required' => false, 'const' => "''", 'type' => 'text'],
    'UF_NAME'           => ['required' => true,  'inputKey' => 'author', 'type' => 'text'],
    'UF_EMAIL'          => ['required' => false, 'const' => "''", 'type' => 'text'],
    'UF_DATETIME'       => ['required' => true,  'inputKey' => 'created_at', 'type' => 'datetime'],
    'UF_TEXT'           => ['required' => true,  'inputKey' => 'text', 'type' => 'text'],
    'UF_GRADE'          => ['required' => true,  'inputKey' => 'rating', 'type' => 'int'],
    'UF_BITRIX_ITEM_ID' => ['required' => true,  'inputKey' => 'product_id', 'type' => 'int'],
    'UF_MEDIA'          => ['required' => false, 'inputKey' => 'photo', 'type' => 'text'],
    'UF_MARKET_ID'      => ['required' => true,  'inputKey' => 'comment_id', 'type' => 'market_id'],
    'UF_SITE_ID'        => ['required' => false, 'const' => "'0f'", 'type' => 'text'],
    'UF_SOURCE'         => ['required' => false, 'const' => 15, 'type' => 'int']
];

$sql = '';

foreach ($rows as $row)
{
    $values = [];
    $skip = false;

    foreach ($fieldMap as $column => $meta)
    {
        if (isset($meta['const']))
        {
            $values[$column] = $meta['const'];
            continue;
        }

        if(empty($headers[$meta['inputKey']]))
        {
            echo "‼️ Ошибка! Ключ {$meta['inputKey']} в файле не найден";
            exit();
        }

        $key = $headers[$meta['inputKey']];
        $raw = $row[$headers[$meta['inputKey']]] ?? '';

        if ($meta['required'] && empty($raw))
        {
            $skip = true;
            break;
        }

        switch ($meta['type'])
        {
            case 'datetime':
                $fmt = new \IntlDateFormatter('ru_RU', \IntlDateFormatter::LONG, \IntlDateFormatter::NONE, null, null, 'd MMMM yyyy');
                $timestamp = $fmt->parse($raw);
                $values[$column] = "'" .date('Y-m-d', $timestamp) . "'";
                break;
            case 'market_id':
                $values[$column] = "'ozon:" . addslashes($raw) . "'";
                break;
            case 'int':
                $values[$column] = (int)$raw;
                break;
            case 'text':
                $values[$column] = "'" . addslashes($raw) . "'";
                break;
            default:
                $values[$column] = "''";
        }
    }

    if($skip)
        continue;

    $columns = implode(', ', array_keys($values));
    $vals    = implode(', ', array_values($values));

    $sql .= "INSERT INTO `app_product_review` ($columns) VALUES ($vals);\n";
}

file_put_contents($outputFile, $sql);
echo "✅ Готово! SQL-запросы сохранены в $outputFile\n";
