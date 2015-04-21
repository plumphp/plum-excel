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

use \PHPExcel_IOFactory;
use Plum\Plum\Converter\HeaderConverter;
use Plum\Plum\Filter\SkipFirstFilter;
use Plum\Plum\Workflow;
use Plum\Plum\Writer\ArrayWriter;

/**
 * ExcelTest
 *
 * @package   Plum\PlumExcel
 * @author    Florian Eckerstorfer <florian@eckerstorfer.co>
 * @copyright 2015 Florian Eckerstorfer
 * @group     functional
 */
class ExcelHeaderTest extends \PHPUnit_Framework_TestCase
{
    /**
     * @test
     */
    public function testWorkflow()
    {
        $reader = new ExcelReader(PHPExcel_IOFactory::load(__DIR__.'/fixtures/test.xlsx'));
        $writer = new ArrayWriter();

        $workflow = new Workflow();
        $workflow
            ->addConverter(new HeaderConverter())
            ->addFilter(new SkipFirstFilter(1))
            ->addWriter($writer);
        $result = $workflow->process($reader);

        $this->assertSame(0, $result->getErrorCount());
        $this->assertSame(4, $result->getReadCount());
        $this->assertSame(3, $result->getItemWriteCount());
        $this->assertSame('line 1A', $writer->getData()[0]['col A']);
        $this->assertSame('line 2B', $writer->getData()[1]['col B']);
        $this->assertSame('line 3C', $writer->getData()[2]['col C']);
    }
}
