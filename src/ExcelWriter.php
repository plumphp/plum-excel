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
use Cocur\Vale\Vale;

/**
 * ExcelWriter.
 *
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
     * @var string[]|null
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
     * @param string                       $filename
     * @param PHPExcel|null                $excel
     * @param string                       $format   Format, defaults to `Excel2007`
     * @param PHPExcel_Writer_IWriter|null $writer
     *
     * @codeCoverageIgnore
     */
    public function __construct(
        $filename,
        PHPExcel $excel                 = null,
        $format                         = 'Excel2007',
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
     */
    public function writeItem($item)
    {
        if ($this->autoDetectHeader && $this->header === null) {
            $this->handleAutoDetectHeaders($item);
        }

        if (is_array($item)) {
            $keys = array_keys($item);
        } elseif ($this->header && is_object($item)) {
            $keys = $this->header;
        } else {
            throw new \InvalidArgumentException(sprintf(
                'Plum\PlumExcel\ExcelWriter currently only supports array items or objects if headers are set using '.
                'the setHeader() method. You have passed an item of type "%s" to writeItem().',
                gettype($item)
            ));
        }

        $column = 0;
        foreach ($keys as $key) {
            $value = Vale::get($item, $key);
            $this->excel->getActiveSheet()->setCellValueByColumnAndRow($column, $this->currentRow, $value);
            ++$column;
        }
        ++$this->currentRow;
    }

    /**
     * Prepare the writer.
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
     */
    public function finish()
    {
        $writer = $this->writer;
        if (!$writer) {
            $writer = PHPExcel_IOFactory::createWriter($this->excel, $this->format);
        }
        $writer->save($this->filename);
    }

    protected function handleAutoDetectHeaders($item)
    {
        if (!is_array($item)) {
            throw new \InvalidArgumentException(sprintf(
                'Plum\PlumExcel\ExcelWriter currently only supports header detection if the item passed to '.
                'writeItem() is an array. "%s" was passed writeItem().',
                gettype($item)
            ));
        }
        $this->header = array_keys($item);
        $this->writeItem($this->header);
    }
}
