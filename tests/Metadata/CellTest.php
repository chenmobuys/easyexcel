<?php

namespace EasyExcelTests\Metadata;

use EasyExcel\Metadata\Cell;
use EasyExcel\Metadata\Hyperlink;
use EasyExcel\Metadata\Style;
use PHPUnit\Framework\TestCase;

class CellTest extends TestCase
{
    public function test_rowIndex()
    {
        $cell = new Cell();
        $this->assertNull($cell->getRowIndex());
        $cell->setRowIndex(1);
        $this->assertEquals(1, $cell->getRowIndex());
    }

    public function test_columnIndex()
    {
        $cell = new Cell();
        $this->assertNull($cell->getColumnIndex());
        $cell->setColumnIndex(1);
        $this->assertEquals(1, $cell->getColumnIndex());
    }

    public function test_coordinate()
    {
        $cell = new Cell(null, null, 0, 0);
        $this->assertEquals("A1", $cell->getCoordinate());
    }

    public function test_xfIndex()
    {
        $cell = new Cell();
        $this->assertEquals(0, $cell->getXfIndex());
        $cell->setXfIndex(1);
        $this->assertEquals(1, $cell->getXfIndex());
    }

    public function test_value()
    {
        $cell = new Cell();
        $this->assertNull($cell->getValue());
        $cell->setValue(1);
        $this->assertEquals(1, $cell->getValue());
    }

    public function test_formulaValue()
    {
        $cell = new Cell();
        $this->assertNull($cell->getFormulaValue());
        $cell->setFormulaValue(1);
        $this->assertEquals(1, $cell->getFormulaValue());
    }

    public function test_formattedValue()
    {
        $cell = new Cell();
        $this->assertNull($cell->getFormattedValue());
        $cell->setFormattedValue(1);
        $this->assertEquals(1, $cell->getFormattedValue());

        $style = Style::builder()
            ->setFormatCode(Style\Format::FORMAT_DATE_DDMMYYYY)
            ->build();
        $cell = new Cell(1000);
        $cell->setStyle($style);
        $this->assertEquals("26/09/1902", $cell->getFormattedValue());
    }

    public function test_mergeCell()
    {
        $cell = new Cell();
        $this->assertNull($cell->getMergeCell());
        $cell->setMergeCell("A1:B1");
        $this->assertEquals("A1:B1", $cell->getMergeCell());
    }

    public function test_hyperlink()
    {
        $cell = new Cell();
        $this->assertNotNull($cell->getHyperlink());
        $this->assertTrue($cell->hasHyperlink());
        $cell->setHyperlink(new Hyperlink("https://foo.com"));
        $this->assertEquals("https://foo.com", $cell->getHyperlink()->getUrl());
    }

    public function test_style()
    {
        $cell = new Cell();
        $this->assertNotNull($cell->getStyle());
        $this->assertTrue($cell->hasStyle());
        $style = new Style();
        $cell->setStyle($style);
        $this->assertNotNull($cell->getStyle());
        $this->assertTrue($cell->hasStyle());
    }
}