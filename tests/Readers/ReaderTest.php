<?php

namespace EasyExcelTests\Readers;

use EasyExcel\Factory;
use EasyExcel\Interfaces\ExcelInterface;
use EasyExcel\Readers\Csv\CsvReader;
use PHPUnit\Framework\TestCase;

class ReaderTest extends TestCase
{
    public function excelDataProvider(): array
    {
        return [
            ["tests/data/sample.csv"],
            ["tests/data/sample.ods"],
            ["tests/data/sample.xls"],
            ["tests/data/sample.xlsx"],
        ];
    }

    /**
     * @dataProvider excelDataProvider
     *
     * @param  string  $filename
     *
     * @return void
     */
    public function test_load(string $filename)
    {
        $reader = Factory::load($filename);

        $actualSheetName = $reader instanceof CsvReader ? (pathinfo($filename, PATHINFO_FILENAME)) : "Sheet1";
        $this->assertEquals($actualSheetName, $reader->getActiveSheet()->getName());
        $this->assertEquals(2, $reader->getActiveSheet()->getTotalRows());
        $this->assertEquals(3, $reader->getActiveSheet()->getTotalColumns());
    }

    /**
     * @dataProvider excelDataProvider
     *
     * @param  string  $filename
     *
     * @return void
     */
    public function test_readable(string $filename)
    {
        $reader = Factory::load($filename);
        $this->assertTrue($reader::readable($filename));
        $this->assertFalse($reader::readable("/foo"));
    }

    /**
     * @dataProvider excelDataProvider
     *
     * @param  string  $filename
     *
     * @return void
     */
    public function test_read(string $filename)
    {
        $reader = Factory::load($filename);
        $actualRows = array_map(function ($item) {
            return $item->toArray();
        }, iterator_to_array($reader->getRowIterator()));
        $expectedRows = [
            ['Title1', 'Title2', 'Title3'],
            ['Desc1', 'Desc2', 'Desc3'],
        ];
        $this->assertEquals($expectedRows, $actualRows);
    }
}