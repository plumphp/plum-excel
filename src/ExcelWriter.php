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

use PHPExcel;
use PHPExcel_IOFactory;
use PHPExcel_Writer_IWriter;
use Plum\Plum\Writer\WriterInterface;

/**
 * ExcelWriter
 *
 * @package   Plum\PlumExcel
 * @author    Florian Eckerstorfer <florian@eckerstorfer.co>
 * @copyright 2015 Florian Eckerstorfer
 */
class ExcelWriter implements WriterInterface
{
    /**
     * @var string
     */
    private $filename;

    /**
     * @var string[]
     */
    private $header;

    /**
     * @var bool
     */
    private $autoDetectHeader = false;

    /**
     * @var PHPExcel
     */
    private $excel;

    /**
     * @var string
     */
    private $format;

    /**
     * @var PHPExcel_Writer_IWriter
     */
    private $writer;

    /**
     * @var int
     */
    private $currentRow = 1;

    /**
     * @param string   $filename
     * @param PHPExcel $excel
     * @param string   $format   Format, defaults to `Excel2007`
     *
     * @codeCoverageIgnore
     */
    public function __construct(
        $filename,
        PHPExcel $excel = null,
        $format = 'Excel2007',
        PHPExcel_Writer_IWriter $writer = null
    ) {
        $this->filename = $filename;
        $this->excel    = $excel;
        $this->format   = $format;
        $this->writer   = $writer;
    }

    /**
     * @param string[] $header
     *
     * @return ExcelWriter
     */
    public function setHeader(array $header)
    {
        $this->header = $header;

        return $this;
    }

    /**
     * @param bool $autoDetectHeader
     *
     * @return ExcelWriter
     */
    public function autoDetectHeader($autoDetectHeader = true)
    {
        $this->autoDetectHeader = $autoDetectHeader;

        return $this;
    }

    /**
     * Write the given item.
     *
     * @param mixed $item
     *
     * @return void
     */
    public function writeItem($item)
    {
        if ($this->autoDetectHeader && !$this->header) {
            $this->header = array_keys($item);
            $this->writeItem($this->header);
        }

        $column = 0;
        foreach ($item as $value) {
            $this->excel->getActiveSheet()->setCellValueByColumnAndRow($column, $this->currentRow, $value);
            $column++;
        }
        $this->currentRow++;
    }

    /**
     * Prepare the writer.
     *
     * @return void
     */
    public function prepare()
    {
        if ($this->excel === null) {
            $this->excel = new PHPExcel();
        }

        if ($this->header !== null) {
            $this->writeItem($this->header);
        }
    }

    /**
     * Finish the writer.
     *
     * @return void
     */
    public function finish()
    {
        $writer = $this->writer;
        if (!$writer) {
            $writer = PHPExcel_IOFactory::createWriter($this->excel, $this->format);
        }
        $writer->save($this->filename);
    }
}
