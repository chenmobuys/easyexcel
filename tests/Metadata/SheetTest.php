<?php

namespace EasyExcelTests\Metadata;

use EasyExcel\Factory;
use EasyExcel\Interfaces\ExcelInterface;
use EasyExcel\Interfaces\SheetInterface;
use EasyExcel\Metadata\AutoFilter;
use EasyExcel\Metadata\Hyperlink;
use PHPUnit\Framework\TestCase;

class SheetTest extends TestCase
{
    public function sheet(): array
    {
        return [
            [Factory::load('tests/data/sample.csv')->getActiveSheet()]
        ];
    }

    /**
     * @dataProvider  sheet
     * @return void
     */
    public function test_getExcel(SheetInterface $sheet)
    {
        $this->assertInstanceOf(ExcelInterface::class, $sheet->getExcel());
    }

    /**
     * @dataProvider sheet
     * @return void
     */
    public function test_getName(SheetInterface $sheet)
    {
        $this->assertEquals("sample", $sheet->getName());
    }

    /**
     * @dataProvider sheet
     * @return void
     */
    public function test_setName(SheetInterface $sheet)
    {
        $sheet->setName("Foo");
        $this->assertEquals("Foo", $sheet->getName());
    }

    /**
     * @dataProvider sheet
     * @return void
     */
    public function test_getIndex(SheetInterface $sheet)
    {
        $this->assertEquals(0, $sheet->getIndex());
    }

    /**
     * @dataProvider sheet
     * @return void
     */
    public function test_setIndex(SheetInterface $sheet)
    {
        $sheet->setIndex(4);
        $this->assertEquals(4, $sheet->getIndex());
    }

    /**
     * @dataProvider sheet
     * @return void
     */
    public function test_getTotalRows(SheetInterface $sheet)
    {
        $this->assertEquals(2, $sheet->getTotalRows());
    }

    /**
     * @dataProvider sheet
     * @return void
     */
    public function test_setTotalRows(SheetInterface $sheet)
    {
        $sheet->setTotalRows(3);
        $this->assertEquals(3, $sheet->getTotalRows());
    }

    /**
     * @dataProvider sheet
     * @return void
     */
    public function test_getTotalColumns(SheetInterface $sheet)
    {
        $this->assertEquals(3, $sheet->getTotalColumns());
    }

    /**
     * @dataProvider sheet
     * @return void
     */
    public function test_setTotalColumns(SheetInterface $sheet)
    {
        $sheet->setTotalColumns(4);
        $this->assertEquals(4, $sheet->getTotalColumns());
    }

    /**
     * @dataProvider sheet
     * @return void
     */
    public function test_getLastColumnIndex(SheetInterface $sheet)
    {
        $this->assertEquals(2, $sheet->getLastColumnIndex());
    }

    /**
     * @dataProvider sheet
     * @return void
     */
    public function test_getLastColumnLetter(SheetInterface $sheet)
    {
        $this->assertEquals("C", $sheet->getLastColumnLetter());
    }

    /**
     * @dataProvider sheet
     * @return void
     */
    public function test_getAutoFilter(SheetInterface $sheet)
    {
        $this->assertNull($sheet->getAutoFilter());
    }

    /**
     * @dataProvider sheet
     * @return void
     */
    public function test_setAutoFilter(SheetInterface $sheet)
    {
        $sheet->setAutoFilter(new AutoFilter("A1:A2"));
        $this->assertNotNull($sheet->getAutoFilter());
    }

    /**
     * @dataProvider sheet
     * @return void
     */
    public function test_getHyperlink(SheetInterface $sheet)
    {
        $this->assertNull($sheet->getHyperlink("A1"));
    }

    /**
     * @dataProvider sheet
     * @return void
     */
    public function test_setHyperlink(SheetInterface $sheet)
    {
        $sheet->setHyperlink("A1", new Hyperlink("https://example.com"));
        $this->assertNotNull($sheet->getHyperlink("A1"));
        $sheet->setHyperlink("A1", null);
        $this->assertNull($sheet->getHyperlink("A1"));
    }

    /**
     * @dataProvider sheet
     * @return void
     */
    public function test_hasHyperlink(SheetInterface $sheet)
    {
        $this->assertFalse($sheet->hasHyperlink("A1"));
        $sheet->setHyperlink("A1", new Hyperlink("https://example.com"));
        $this->assertTrue($sheet->hasHyperlink("A1"));
    }

    /**
     * @dataProvider sheet
     * @return void
     */
    public function test_setHyperlinks(SheetInterface $sheet)
    {
        $hyperlinks = [
            'A1' => new Hyperlink("https://example.com", "com"),
            'B1' => new Hyperlink("https://example.org", "org")
        ];
        $sheet->setHyperlinks($hyperlinks);
        $this->assertNotNull($sheet->getHyperlink("A1"));
        $this->assertNotNull($sheet->getHyperlink("B1"));
    }

    /**
     * @dataProvider sheet
     * @return void
     */
    public function test_getHyperlinks(SheetInterface $sheet)
    {
        $this->assertEmpty($sheet->getHyperlinks());
    }

    /**
     * @dataProvider sheet
     * @return void
     */
    public function test_getMergeCell(SheetInterface $sheet)
    {
        $this->assertNull($sheet->getMergeCell("A1"));
    }

    /**
     * @dataProvider sheet
     * @return void
     */
    public function test_setMergeCell(SheetInterface $sheet)
    {
        $sheet->setMergeCell("A1", "A1:B1");
        $this->assertNotNull($sheet->getMergeCell("A1"));
        $sheet->setMergeCell("A1", null);
        $this->assertNull($sheet->getMergeCell("A1"));
    }

    /**
     * @dataProvider sheet
     * @return void
     */
    public function test_getMergeCells(SheetInterface $sheet)
    {
        $this->assertEmpty($sheet->getMergeCells());
    }

    /**
     * @dataProvider sheet
     * @return void
     */
    public function test_setMergeCells(SheetInterface $sheet)
    {
        $mergeCells = ["A1" => "A1:B1",];
        $sheet->setMergeCells($mergeCells);
        $this->assertCount(1, $sheet->getMergeCells());
    }
}