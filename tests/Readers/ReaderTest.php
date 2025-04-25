<?php

namespace EasyExcelTests\Readers;

use EasyExcel\Factory;
use EasyExcel\Readers\Csv\Reader as CsvReader;
use PHPUnit\Framework\TestCase;

class ReaderTest extends TestCase
{
    public function excelDataProvider(): array
    {
        return [
            [__DIR__."/../data/sample.csv"],
            [__DIR__."/../data/sample.ods"],
            [__DIR__."/../data/sample.xls"],
            [__DIR__."/../data/sample.xlsx"],
        ];
    }

    /**
     * @dataProvider excelDataProvider
     *
     * @param  string  $filename
     *
     * @return void
     *
     * @throws \Exception
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
     *
     * @throws \Exception
     */
    public function test_readable(string $filename)
    {
        $reader = Factory::load($filename);
        $this->assertTrue($reader->readable($filename));
        $this->assertFalse($reader->readable("/foo"));
    }

    /**
     * @dataProvider excelDataProvider
     *
     * @param  string  $filename
     *
     * @return void
     *
     * @throws \Exception
     */
    public function test_read(string $filename)
    {
        $reader = Factory::load($filename);
        $actualRows = [];
        foreach ($reader->getActiveSheet()->getRowIterator() as $row) {
            $actualRows[] = $row->toArray();
        }
        $expectedRows = [
            ['Title1', 'Title2', 'Title3'],
            ['Desc1', 'Desc2', 'Desc3'],
        ];
        $this->assertEquals($expectedRows, $actualRows);
    }
}