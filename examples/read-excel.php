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

$excel = PHPExcel_IOFactory::load(__DIR__.'/example.xlsx');
$reader = new ExcelReader($excel);
$reader->setHeaderRow(0);
foreach ($reader->getIterator() as $row) {
    print_r($row);
}
