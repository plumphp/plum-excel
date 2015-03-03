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

use Mockery;
use org\bovigo\vfs\vfsStream;

/**
 * ExcelWriterTest
 *
 * @package   Plum\PlumExcel
 * @author    Florian Eckerstorfer <florian@eckerstorfer.co>
 * @copyright 2015 Florian Eckerstorfer
 * @group     unit
 */
class ExcelWriterTest extends \PHPUnit_Framework_TestCase
{
    /**
     * @var ExcelWriter
     */
    private $writer;

    /**
     * @var \PHPExcel|\Mockery\MockInterface
     */
    private $excel;

    /**
     * @var \PHPExcel_Writer_IWriter|\Mockery\MockInterface
     */
    private $excelWriter;

    public function setUp()
    {
        $this->excel = Mockery::mock('\PHPExcel');
        $this->excel->shouldReceive('getID');
        $this->excel->shouldReceive('disconnectWorksheets');

        $this->excelWriter = Mockery::mock('\PHPExcel_Writer_IWriter');

        vfsStream::setup('fixtures', null, ['text.xlsx' => []]);

        $this->writer = new ExcelWriter(
            vfsStream::url('fixtures/test.xlsx'),
            $this->excel,
            'Excel2007',
            $this->excelWriter
        );
    }

    /**
     * @test
     * @covers Plum\PlumExcel\ExcelWriter::setHeader()
     * @covers Plum\PlumExcel\ExcelWriter::prepare()
     */
    public function prepareWritesHeader()
    {
        $sheet = $this->getMockWorksheet();
        $sheet->shouldReceive('setCellValueByColumnAndRow')->with(0, 1, 'City')->once();
        $sheet->shouldReceive('setCellValueByColumnAndRow')->with(1, 1, 'Country')->once();

        $this->excel->shouldReceive('getActiveSheet')->andReturn($sheet);

        $this->writer->setHeader(['City', 'Country']);
        $this->writer->prepare();
    }
    
    /**
     * @test
     * @covers Plum\PlumExcel\ExcelWriter::writeItem()
     */
    public function writeItemWritesItemToExcel()
    {
        $sheet = $this->getMockWorksheet();
        $sheet->shouldReceive('setCellValueByColumnAndRow')->with(0, 1, 'Vienna')->once();
        $sheet->shouldReceive('setCellValueByColumnAndRow')->with(1, 1, 'Austria')->once();

        $this->excel->shouldReceive('getActiveSheet')->andReturn($sheet);

        $this->writer->writeItem(['City' => 'Vienna', 'Country' => 'Austria']);
    }

    /**
     * @test
     * @covers Plum\PlumExcel\ExcelWriter::autoDetectHeader()
     * @covers Plum\PlumExcel\ExcelWriter::writeItem()
     */
    public function writeItemWritesHeaderIfAutoDetectHeaderIsTrueAndItemToExcel()
    {
        $sheet = $this->getMockWorksheet();
        $sheet->shouldReceive('setCellValueByColumnAndRow')->with(0, 1, 'City')->once();
        $sheet->shouldReceive('setCellValueByColumnAndRow')->with(1, 1, 'Country')->once();
        $sheet->shouldReceive('setCellValueByColumnAndRow')->with(0, 2, 'Vienna')->once();
        $sheet->shouldReceive('setCellValueByColumnAndRow')->with(1, 2, 'Austria')->once();

        $this->excel->shouldReceive('getActiveSheet')->andReturn($sheet);

        $this->writer->autoDetectHeader();
        $this->writer->writeItem(['City' => 'Vienna', 'Country' => 'Austria']);
    }

    /**
     * @test
     * @covers Plum\PlumExcel\ExcelWriter::finish()
     */
    public function finishWritesExcelFileToDisk()
    {
        $this->excelWriter->shouldReceive('save')->with(vfsStream::url('fixtures/test.xlsx'))->once();

        $this->writer->finish();
    }

    /**
     * @return Mockery\MockInterface|\PHPExcel_Worksheet
     */
    private function getMockWorksheet()
    {
        $sheet = Mockery::mock('\PHPExcel_Worksheet');
        $sheet->shouldReceive('disconnectCells');

        return $sheet;
    }
}
