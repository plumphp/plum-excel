<?php

/**
 * This file is part of plumphp/plum-excel.
 *
 * (c) Florian Eckerstorfer <florian@eckerstorfer.co>
 *
 * For the full copyright and license information, please view the LICENSE
 * file that was distributed with this source code.
 */

use Plum\PlumExcel\ExcelReader;

require_once __DIR__.'/../vendor/autoload.php';

if (!file_exists(__DIR__.'/example.xlsx')) {
    echo "Please run the write-excel.php first to generate a Excel file.\n";
    exit(1);
}

$excel = PHPExcel_IOFactory::load(__DIR__.'/example.xlsx');
$reader = new ExcelReader($excel);
foreach ($reader->getIterator() as $row) {
    print_r($row);
}
