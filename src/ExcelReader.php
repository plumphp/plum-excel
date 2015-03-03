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
     * @var int
     */
    private $headerRow;

    /**
     * @var array[][]
     */
    private $data;

    /**
     * @param PHPExcel $excel
     *
     * @codeCoverageIgnore
     */
    public function __construct(PHPExcel $excel)
    {
        $this->excel = $excel;
    }

    /**
     * @param int $headerRow
     *
     * @return ExcelReader
     */
    public function setHeaderRow($headerRow)
    {
        $this->headerRow = $headerRow;

        return $this;
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
        if ($this->headerRow !== null) {
            $header = $this->getHeader($this->headerRow + 1);
        } else {
            $header = [];
        }
        $startRow = $this->headerRow !== null ? $this->headerRow+2 : 1;
        $rowIndex = 0;
        foreach ($sheet->getRowIterator($startRow) as $excelRow) {
            $this->data[$rowIndex] = [];
            $columnIndex = 0;
            foreach ($excelRow->getCellIterator() as $excelCell) {
                $this->data[$rowIndex][$this->getKey($header, $columnIndex)] = $excelCell->getValue();
                $columnIndex++;
            }
            $rowIndex++;
        }

        return $this->data;
    }

    /**
     * @param int $headerRow
     *
     * @return string[]
     */
    protected function getHeader($headerRow)
    {
        $header = [];

        $sheet = $this->excel->getActiveSheet();
        foreach ($sheet->getRowIterator($headerRow)->current()->getCellIterator() as $cell) {
            $header[] = $cell->getValue();
        }

        return $header;
    }

    /**
     * @param string[] $header
     * @param int      $index
     *
     * @return string|int
     */
    protected function getKey(array $header, $index)
    {
        return isset($header[$index]) ? $header[$index] : $index;
    }
}
