<?php

/**
 * This file is part of plumphp/plum-excel.
 *
 * (c) Florian Eckerstorfer <florian@eckerstorfer.co>
 *
 * For the full copyright and license information, please view the LICENSE
 * file that was distributed with this source code.
 */

namespace Plum\PlumExcel;

use ArrayIterator;
use PHPExcel;
use PHPExcel_IOFactory;
use Plum\Plum\Reader\ReaderInterface;

/**
 * ExcelReader
 *
 * @package   Plum\PlumExcel
 * @author    Florian Eckerstorfer <florian@eckerstorfer.co>
 * @copyright 2015 Florian Eckerstorfer
 */
class ExcelReader implements ReaderInterface
{
    /**
     * @var PHPExcel
     */
    private $excel;

    /**
     * @var array[][]
     */
    private $data;

    /**
     * @param PHPExcel|string $excel
     *
     * @codeCoverageIgnore
     */
    public function __construct($excel)
    {
        if (!$excel instanceof PHPExcel) {
            $excel = PHPExcel_IOFactory::load($excel);
        }
        $this->excel = $excel;
    }

    /**
     * @return ArrayIterator
     */
    public function getIterator()
    {
        return new ArrayIterator($this->getData());
    }

    /**
     * @return int
     */
    public function count()
    {
        return count($this->getData());
    }

    /**
     * @return array[][]
     */
    protected function getData()
    {
        if ($this->data) {
            return $this->data;
        }

        $this->data = [];
        $sheet = $this->excel->getActiveSheet();
        $rowIndex = 0;
        foreach ($sheet->getRowIterator(1) as $excelRow) {
            $this->data[$rowIndex] = [];
            $columnIndex = 0;
            $cellIterator = $excelRow->getCellIterator();
            $cellIterator->setIterateOnlyExistingCells(false);
            foreach ($cellIterator as $excelCell) {
                $this->data[$rowIndex][] = $excelCell->getValue();
                $columnIndex++;
            }
            $rowIndex++;
        }

        return $this->data;
    }

    /**
     * @param mixed $input
     *
     * @return bool
     */
    public static function accepts($input)
    {
        return $input instanceof PHPExcel || (is_string($input) && preg_match('/\.(xls|xlsx)$/', $input));
    }
}
