<?php

require_once __DIR__.'/../vendor/autoload.php';

use Plum\PlumExcel\ExcelWriter;

$writer = new ExcelWriter(__DIR__.'/example.xlsx');
$writer->autoDetectHeader();
$writer->prepare();
$writer->writeItem(['Town' => 'Vienna', 'Country' => 'Austria', 'District' => 'Alsergrund', 'DistrictNumber' => 1090]);
$writer->finish();
