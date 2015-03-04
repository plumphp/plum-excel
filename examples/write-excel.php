<?php

/**
 * This file is part of plumphp/plum-excel.
 *
 * (c) Florian Eckerstorfer <florian@eckerstorfer.co>
 *
 * For the full copyright and license information, please view the LICENSE
 * file that was distributed with this source code.
 */

require_once __DIR__.'/../vendor/autoload.php';

use Plum\PlumExcel\ExcelWriter;

$writer = new ExcelWriter(__DIR__.'/example.xlsx');
$writer->autoDetectHeader();
$writer->prepare();
$writer->writeItem(['Town' => 'Vienna', 'Country' => 'Austria', 'District' => 'Alsergrund', 'DistrictNumber' => 1090]);
$writer->writeItem(['Town' => 'Vienna', 'Country' => 'Austria', 'District' => 'Mariahilf', 'DistrictNumber' => 1060]);
$writer->finish();
