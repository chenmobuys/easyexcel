<?php

namespace EasyExcelTests\Writers;

use EasyExcel\Factory;
use PHPUnit\Framework\TestCase;

class XlsxWriterTest extends TestCase
{
    public function test_open()
    {
        $filename = "tests/data/temp/test.xlsx";
        $writer = Factory::open($filename);
        $writer->addRow([1, 2, 3])
            ->close();
        $reader = Factory::load($filename);
        $actualRows = array_map(function ($item) {
            return $item->toArray();
        }, iterator_to_array($reader->getRowIterator()));
        $this->assertEquals([[1, 2, 3]], $actualRows);
        @unlink($filename);
    }
}