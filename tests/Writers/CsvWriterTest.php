<?php

namespace EasyExcelTests\Writers;

use EasyExcel\Factory;
use PHPUnit\Framework\TestCase;

class CsvWriterTest extends TestCase
{
    public function test_open()
    {
        $filename = "tests/data/temp/test.csv";
        $writer = Factory::open($filename);
        $writer->addRow([1, 2, 3])
            ->close();
        $reader = Factory::load($filename);
        $actualRows = [];
        foreach ($reader->getRowIterator() as $row) {
            $actualRows[] = $row->toArray();
        }
        $this->assertEquals([[1, 2, 3]], $actualRows);
        @unlink($filename);
    }
}