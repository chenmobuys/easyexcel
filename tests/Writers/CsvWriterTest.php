<?php

namespace EasyExcelTests\Writers;

use EasyExcel\Factory;
use PHPUnit\Framework\TestCase;

class CsvWriterTest extends TestCase
{
    /**
     * @throws \EasyExcel\Exceptions\UnknownExcelTypeException
     * @throws \EasyExcel\Exceptions\NotSupportedExcelTypeException
     * @throws \EasyExcel\Exceptions\FileNotWriteableException
     * @throws \EasyExcel\Exceptions\FileNotReadableException
     */
    public function test_open()
    {
        $filename = __DIR__."/../data/temp/test.csv";
        if (file_exists($filename)) {
            @unlink($filename);
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
    }
}