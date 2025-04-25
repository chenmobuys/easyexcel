<?php

namespace EasyExcelTests\Metadata;

use EasyExcel\Factory;
use EasyExcel\Interfaces\ReaderSheet;
use PHPUnit\Framework\TestCase;

class ExcelTest extends TestCase
{
    /**
     * @throws \EasyExcel\Exceptions\NotSupportedExcelTypeException
     * @throws \EasyExcel\Exceptions\FileNotReadableException
     * @throws \EasyExcel\Exceptions\UnknownExcelTypeException
     */
    public function test_hasSheetName()
    {
        $excel = Factory::load(__DIR__."/../data/sample.xlsx");
        $this->assertTrue($excel->hasSheetName("Sheet1"));
        $this->assertFalse($excel->hasSheetName("foo"));
    }

    /**
     * @throws \EasyExcel\Exceptions\NotSupportedExcelTypeException
     * @throws \EasyExcel\Exceptions\FileNotReadableException
     * @throws \EasyExcel\Exceptions\UnknownExcelTypeException
     */
    public function test_hasSheetIndex()
    {
        $excel = Factory::load(__DIR__."/../data/sample.xlsx");
        $this->assertTrue($excel->hasSheetIndex(0));
        $this->assertFalse($excel->hasSheetIndex(3));
    }

    /**
     * @throws \EasyExcel\Exceptions\NotSupportedExcelTypeException
     * @throws \EasyExcel\Exceptions\FileNotReadableException
     * @throws \EasyExcel\Exceptions\UnknownExcelTypeException
     */
    public function test_getActiveSheet()
    {
        $excel = Factory::load(__DIR__."/../data/sample.xlsx");
        $this->assertInstanceOf(ReaderSheet::class, $excel->getActiveSheet());
    }

    /**
     * @throws \EasyExcel\Exceptions\NotSupportedExcelTypeException
     * @throws \EasyExcel\Exceptions\FileNotReadableException
     * @throws \EasyExcel\Exceptions\UnknownExcelTypeException
     */
    public function test_setActiveSheetByName()
    {
        $excel = Factory::load(__DIR__."/../data/sample.xlsx");
        $this->assertInstanceOf(ReaderSheet::class, $excel->getSheetByName("Sheet1"));
        $this->assertEquals("Sheet1", $excel->getActiveSheet()->getName());
    }

    /**
     * @throws \EasyExcel\Exceptions\NotSupportedExcelTypeException
     * @throws \EasyExcel\Exceptions\FileNotReadableException
     * @throws \EasyExcel\Exceptions\UnknownExcelTypeException
     */
    public function test_setActiveSheetByIndex()
    {
        $excel = Factory::load(__DIR__."/../data/sample.xlsx");
        $this->assertInstanceOf(ReaderSheet::class, $excel->getSheetByIndex(1));
    }

    /**
     * @throws \EasyExcel\Exceptions\NotSupportedExcelTypeException
     * @throws \EasyExcel\Exceptions\FileNotReadableException
     * @throws \EasyExcel\Exceptions\UnknownExcelTypeException
     */
    public function test_getAllSheets()
    {
        $excel = Factory::load(__DIR__."/../data/sample.xlsx");
        $this->assertCount(3, $excel->getAllSheets());
    }

    /**
     * @throws \EasyExcel\Exceptions\NotSupportedExcelTypeException
     * @throws \EasyExcel\Exceptions\FileNotReadableException
     * @throws \EasyExcel\Exceptions\UnknownExcelTypeException
     */
    public function test_removeSheetByName()
    {
        $excel = Factory::load(__DIR__."/../data/sample.xlsx");
        $excel->removeSheetByName("Sheet1");
        $this->assertCount(2, $excel->getAllSheets());
    }

    /**
     * @throws \EasyExcel\Exceptions\NotSupportedExcelTypeException
     * @throws \EasyExcel\Exceptions\FileNotReadableException
     * @throws \EasyExcel\Exceptions\UnknownExcelTypeException
     */
    public function test_removeSheetByIndex()
    {
        $excel = Factory::load(__DIR__."/../data/sample.xlsx");
        $excel->removeSheetByIndex(0);
        $this->assertCount(2, $excel->getAllSheets());
    }
}