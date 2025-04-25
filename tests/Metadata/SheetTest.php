<?php

namespace EasyExcelTests\Metadata;

use EasyExcel\Factory;
use EasyExcel\Interfaces\ReaderSheet;
use EasyExcel\Metadata\AutoFilter;
use EasyExcel\Metadata\Hyperlink;
use EasyExcel\Readers\ReaderExcel;
use PHPUnit\Framework\TestCase;

class SheetTest extends TestCase
{
    /**
     * @throws \EasyExcel\Exceptions\NotSupportedExcelTypeException
     * @throws \EasyExcel\Exceptions\FileNotReadableException
     * @throws \EasyExcel\Exceptions\UnknownExcelTypeException
     */
    public function sheet(): array
    {
        return [
            [Factory::load(__DIR__.'/../data/sample.csv')->getActiveSheet()]
        ];
    }

    /**
     * @dataProvider  sheet
     * @return void
     */
    public function test_getExcel(ReaderSheet $sheet)
    {
        $this->assertInstanceOf(ReaderExcel::class, $sheet->getExcel());
    }

    /**
     * @dataProvider sheet
     * @return void
     */
    public function test_getName(ReaderSheet $sheet)
    {
        $this->assertEquals("sample", $sheet->getName());
    }

    /**
     * @dataProvider sheet
     * @return void
     */
    public function test_setName(ReaderSheet $sheet)
    {
        $sheet->setName("Foo");
        $this->assertEquals("Foo", $sheet->getName());
    }

    /**
     * @dataProvider sheet
     * @return void
     */
    public function test_getIndex(ReaderSheet $sheet)
    {
        $this->assertEquals(0, $sheet->getIndex());
    }

    /**
     * @dataProvider sheet
     * @return void
     */
    public function test_setIndex(ReaderSheet $sheet)
    {
        $sheet->setIndex(4);
        $this->assertEquals(4, $sheet->getIndex());
    }

    /**
     * @dataProvider sheet
     * @return void
     */
    public function test_getTotalRows(ReaderSheet $sheet)
    {
        $this->assertEquals(2, $sheet->getTotalRows());
    }

    /**
     * @dataProvider sheet
     * @return void
     */
    public function test_setTotalRows(ReaderSheet $sheet)
    {
        $sheet->setTotalRows(3);
        $this->assertEquals(3, $sheet->getTotalRows());
    }

    /**
     * @dataProvider sheet
     * @return void
     */
    public function test_getTotalColumns(ReaderSheet $sheet)
    {
        $this->assertEquals(3, $sheet->getTotalColumns());
    }

    /**
     * @dataProvider sheet
     * @return void
     */
    public function test_setTotalColumns(ReaderSheet $sheet)
    {
        $sheet->setTotalColumns(4);
        $this->assertEquals(4, $sheet->getTotalColumns());
    }

    /**
     * @dataProvider sheet
     * @return void
     */
    public function test_getLastColumnIndex(ReaderSheet $sheet)
    {
        $this->assertEquals(2, $sheet->getLastColumnIndex());
    }

    /**
     * @dataProvider sheet
     * @return void
     */
    public function test_getLastColumnLetter(ReaderSheet $sheet)
    {
        $this->assertEquals("C", $sheet->getLastColumnLetter());
    }

    /**
     * @dataProvider sheet
     * @return void
     */
    public function test_getAutoFilter(ReaderSheet $sheet)
    {
        $this->assertNull($sheet->getAutoFilter());
    }

    /**
     * @dataProvider sheet
     * @return void
     */
    public function test_setAutoFilter(ReaderSheet $sheet)
    {
        $sheet->setAutoFilter(new AutoFilter("A1:A2"));
        $this->assertNotNull($sheet->getAutoFilter());
    }

    /**
     * @dataProvider sheet
     * @return void
     */
    public function test_getHyperlink(ReaderSheet $sheet)
    {
        $this->assertNotNull($sheet->getHyperlink("A1"));
    }

    /**
     * @dataProvider sheet
     * @return void
     */
    public function test_setHyperlink(ReaderSheet $sheet)
    {
        $sheet->setHyperlink("A1", new Hyperlink("https://example.com"));
        $this->assertNotNull($sheet->getHyperlink("A1"));
        $sheet->setHyperlink("A1", null);
        $this->assertNotNull($sheet->getHyperlink("A1"));
    }

    /**
     * @dataProvider sheet
     * @return void
     */
    public function test_hasHyperlink(ReaderSheet $sheet)
    {
        $this->assertFalse($sheet->hasHyperlink("A1"));
        $sheet->setHyperlink("A1", new Hyperlink("https://example.com"));
        $this->assertTrue($sheet->hasHyperlink("A1"));
    }

    /**
     * @dataProvider sheet
     * @return void
     */
    public function test_setHyperlinks(ReaderSheet $sheet)
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
    public function test_getHyperlinks(ReaderSheet $sheet)
    {
        $this->assertEmpty($sheet->getHyperlinks());
    }

    /**
     * @dataProvider sheet
     * @return void
     */
    public function test_getMergeCell(ReaderSheet $sheet)
    {
        $this->assertNull($sheet->getMergeCell("A1"));
    }

    /**
     * @dataProvider sheet
     * @return void
     */
    public function test_setMergeCell(ReaderSheet $sheet)
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
    public function test_getMergeCells(ReaderSheet $sheet)
    {
        $this->assertEmpty($sheet->getMergeCells());
    }

    /**
     * @dataProvider sheet
     * @return void
     */
    public function test_setMergeCells(ReaderSheet $sheet)
    {
        $mergeCells = ["A1" => "A1:B1",];
        $sheet->setMergeCells($mergeCells);
        $this->assertCount(1, $sheet->getMergeCells());
    }
}