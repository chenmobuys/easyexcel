<?php

namespace EasyExcelTests\Writers;

use EasyExcel\Factory;
use PHPUnit\Framework\TestCase;

class XlsxWriterTest extends TestCase
{
    public function test_open()
    {
        $filename = __DIR__."/../data/temp/test.xlsx";
        if (file_exists($filename)) {
            @unlink($filename);
        }
        if (!is_dir(dirname($filename))) {
            mkdir(dirname($filename));
        }
        $easyExcel = Factory::open($filename);
        $easyExcel->getActiveSheet()->getRowWriter()->writes([1, 2, 3]);
        $easyExcel->close();
        $reader = Factory::load($filename);
        $actualRows = [];
        foreach ($reader->getActiveSheet()->getRowIterator() as $row) {
            $actualRows[] = $row->toArray();
        }
        $this->assertEquals([[1, 2, 3]], $actualRows);
        @unlink($filename);
        @rmdir(dirname($filename));
    }
}