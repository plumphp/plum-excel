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

use PHPExcel_IOFactory;

/**
 * ExcelReaderTest.
 *
 * @author    Florian Eckerstorfer <florian@eckerstorfer.co>
 * @copyright 2015 Florian Eckerstorfer
 * @group     unit
 */
class ExcelReaderTest extends \PHPUnit_Framework_TestCase
{
    /**
     * @var ExcelReader
     */
    private $reader;

    /**
     * @var \PHPExcel|\Mockery\MockInterface
     */
    private $excel;

    public function setUp()
    {
        $this->excel = PHPExcel_IOFactory::load(__DIR__.'/fixtures/test.xlsx');

        $this->reader = new ExcelReader($this->excel);
    }

    /**
     * @test
     * @covers Plum\PlumExcel\ExcelReader::getIterator()
     * @covers Plum\PlumExcel\ExcelReader::getData()
     */
    public function getIteratorReturnsIterator()
    {
        $iterator = $this->reader->getIterator();

        $this->assertEquals('col A', $iterator[0][0]);
        $this->assertEquals('col B', $iterator[0][1]);
        $this->assertEquals('line 1A', $iterator[1][0]);
        $this->assertEquals('line 1B', $iterator[1][1]);
        $this->assertCount(4, $iterator);
    }

    /**
     * @test
     * @covers Plum\PlumExcel\ExcelReader::count()
     * @covers Plum\PlumExcel\ExcelReader::getData()
     */
    public function countReturnsNumberOfRows()
    {
        $this->assertEquals(4, $this->reader->count());
    }

    /**
     * @test
     * @covers Plum\PlumExcel\ExcelReader::accepts()
     */
    public function acceptsReturnsTrueIfPHPExcelIsGiven()
    {
        $this->assertTrue(ExcelReader::accepts($this->excel));
    }

    /**
     * @test
     * @covers Plum\PlumExcel\ExcelReader::accepts()
     */
    public function acceptsReturnsTrueIfExcelFilenameIsGiven()
    {
        $this->assertTrue(ExcelReader::accepts('foo.xls'));
        $this->assertTrue(ExcelReader::accepts('foo.xlsx'));
    }

    /**
     * @test
     * @covers Plum\PlumExcel\ExcelReader::accepts()
     */
    public function acceptsReturnsFalseIfInvalidInputIsGiven()
    {
        $this->assertFalse(ExcelReader::accepts('foo.csv'));
        $this->assertFalse(ExcelReader::accepts([]));
    }
}
