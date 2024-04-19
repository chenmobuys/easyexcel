<?php

namespace EasyExcelTests\Metadata;

use EasyExcel\Exceptions\SheetIndexNotExistsException;
use EasyExcel\Exceptions\SheetNameNotExistsException;
use EasyExcel\Factory;
use EasyExcel\Interfaces\SheetInterface;
use PHPUnit\Framework\TestCase;

class ExcelTest extends TestCase
{
    public function test_hasSheetName()
    {
        $excel = Factory::load("tests/data/sample.xlsx");
        $this->assertTrue($excel->hasSheetName("Sheet1"));
        $this->assertFalse($excel->hasSheetName("foo"));
    }

    public function test_hasSheetIndex()
    {
        $excel = Factory::load("tests/data/sample.xlsx");
        $this->assertTrue($excel->hasSheetIndex(0));
        $this->assertFalse($excel->hasSheetIndex(3));
    }

    public function test_getActiveSheet()
    {
        $excel = Factory::load("tests/data/sample.xlsx");
        $this->assertInstanceOf(SheetInterface::class, $excel->getActiveSheet());
    }

    public function test_setActiveSheetByName()
    {
        $excel = Factory::load("tests/data/sample.xlsx");
        $excel->setActiveSheetByName("foo", true);
        $this->assertInstanceOf(SheetInterface::class, $excel->getActiveSheet());
        $this->assertEquals("foo", $excel->getActiveSheet()->getName());
    }

    public function test_setActiveSheetByNameError()
    {
        $this->expectException(SheetNameNotExistsException::class);
        $excel = Factory::load("tests/data/sample.xlsx");
        try {
            $excel->setActiveSheetByName("foo");
        } catch (SheetNameNotExistsException $e) {
            $this->assertEquals("foo", $e->getSheetName());
            throw $e;
        }
    }

    public function test_setActiveSheetByIndex()
    {
        $excel = Factory::load("tests/data/sample.xlsx");
        $excel->setActiveSheetByIndex(1);
        $this->assertInstanceOf(SheetInterface::class, $excel->getActiveSheet());
    }

    public function test_setActiveSheetByIndexError()
    {
        $this->expectException(SheetIndexNotExistsException::class);
        $excel = Factory::load("tests/data/sample.xlsx");
        try {
            $excel->setActiveSheetByIndex(3);
        } catch (SheetIndexNotExistsException $e) {
            $this->assertEquals(3, $e->getSheetIndex());
            throw $e;
        }
    }

    public function test_getAllSheets()
    {
        $excel = Factory::load("tests/data/sample.xlsx");
        $this->assertCount(3, $excel->getAllSheets());
    }

    public function test_removeSheetByName()
    {
        $excel = Factory::load("tests/data/sample.xlsx");
        $excel->removeSheetByName("Sheet1");
        $this->assertCount(2, $excel->getAllSheets());
    }

    public function test_removeSheetByIndex()
    {
        $excel = Factory::load("tests/data/sample.xlsx");
        $excel->removeSheetByIndex(0);
        $this->assertCount(2, $excel->getAllSheets());
    }
}