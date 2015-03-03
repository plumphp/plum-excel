<?php


namespace Plum\PlumExcel;

use Mockery;

/**
 * ExcelReaderTest
 *
 * @package   Plum\PlumExcel
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
        $this->excel  = Mockery::mock('\PHPExcel');
        $this->excel->shouldReceive('getID');
        $this->excel->shouldReceive('disconnectWorksheets');

        $this->reader = new ExcelReader($this->excel);
    }

    /**
     * @test
     * @covers Plum\PlumExcel\ExcelReader::getIterator()
     * @covers Plum\PlumExcel\ExcelReader::getData()
     */
    public function getIteratorReturnsIterator()
    {
        $this->excel
            ->shouldReceive('getActiveSheet')
            ->andReturn($this->getMockWorksheet([['vienna', 'austria'], ['hamburg', 'germany']]));

        $iterator = $this->reader->getIterator();

        $this->assertEquals('vienna', $iterator[0][0]);
        $this->assertEquals('austria', $iterator[0][1]);
        $this->assertEquals('hamburg', $iterator[1][0]);
        $this->assertEquals('germany', $iterator[1][1]);
        $this->assertCount(2, $iterator);
    }

    /**
     * @test
     * @covers Plum\PlumExcel\ExcelReader::getIterator()
     * @covers Plum\PlumExcel\ExcelReader::getData()
     */
    public function getIteratorDoesNotCallExcelTwice()
    {
        $this->excel
            ->shouldReceive('getActiveSheet')
            ->andReturn($this->getMockWorksheet([['vienna', 'austria'], ['hamburg', 'germany']]));

        $this->reader->getIterator();
        $this->reader->getIterator();
    }

    /**
     * @test
     * @covers Plum\PlumExcel\ExcelReader::count()
     * @covers Plum\PlumExcel\ExcelReader::getData()
     */
    public function countReturnsNumberOfRows()
    {
        $this->excel
            ->shouldReceive('getActiveSheet')
            ->andReturn($this->getMockWorksheet([['vienna', 'austria'], ['hamburg', 'germany']]));

        $this->assertEquals(2, $this->reader->count());
    }

    /**
     * @test
     * @covers Plum\PlumExcel\ExcelReader::setHeaderRow()
     * @covers Plum\PlumExcel\ExcelReader::getIterator()
     * @covers Plum\PlumExcel\ExcelReader::getData()
     * @covers Plum\PlumExcel\ExcelReader::getHeader()
     * @covers Plum\PlumExcel\ExcelReader::getKey()
     */
    public function getDataReturnsArrayWithKeyIfHeaderRowIsSet()
    {
        $this->excel
            ->shouldReceive('getActiveSheet')
            ->andReturn($this->getMockWorksheet(
                [['city', 'country'], ['vienna', 'austria'], ['hamburg', 'germany']],
                true
            ));

        $this->reader->setHeaderRow(0);
        $iterator = $this->reader->getIterator();

        $this->assertEquals('vienna', $iterator[0]['city']);
        $this->assertEquals('austria', $iterator[0]['country']);
        $this->assertEquals('hamburg', $iterator[1]['city']);
        $this->assertEquals('germany', $iterator[1]['country']);
        $this->assertCount(2, $iterator);
    }

    /**
     * @param array $data
     *
     * @return Mockery\MockInterface|\PHPExcel_Worksheet
     */
    protected function getMockWorksheet($data, $header = false)
    {
        $rowValid   = [];
        $cellValid  = [];
        $cellValues = [];
        $cellCount  = 0;
        $rowCount   = 0;

        if ($header) {
            $rowCount++;
            foreach ($data[0] as $headerCell) {
                $cellValid[] = true;
                $cellValues[] = $headerCell;
                $cellCount++;
            }
            $cellValid[] = false;
            array_shift($data);
        }

        foreach ($data as $dataRow) {
            $rowValid[] = true;
            $rowCount++;
            foreach ($dataRow as $dataCell) {
                $cellValid[]  = true;
                $cellValues[] = $dataCell;
                $cellCount++;
            }
            $cellValid[] = false;
        }
        $rowValid[] = false;

        $cell = Mockery::mock('\PHPExcel_Cell');
        $cell->shouldReceive('getValue')->andReturnValues($cellValues);

        $cellIterator = Mockery::mock('\PHPExcel_Worksheet_CellIterator', ['rewind' => null, 'next' => Mockery::self()]);
        $cellIterator->shouldReceive('valid')->times(count($cellValid))->andReturnValues($cellValid);
        $cellIterator->shouldReceive('current')->times($cellCount)->andReturn($cell);

        $row = Mockery::mock('PHPExcel_Worksheet_Row');
        $row->shouldReceive('getCellIterator')->andReturn($cellIterator);

        $rowIterator = Mockery::mock('\PHPExcel_Worksheet_RowIterator', ['rewind' => null, 'next' => null]);
        $rowIterator->shouldReceive('valid')->times(count($rowValid))->andReturnValues($rowValid);
        $rowIterator->shouldReceive('current')->times($rowCount)->andReturn($row);

        $sheet = Mockery::mock('\PHPExcel_Worksheet');
        $sheet->shouldReceive('getRowIterator')->andReturn($rowIterator);
        $sheet->shouldReceive('disconnectCells');

        return $sheet;
    }
}
