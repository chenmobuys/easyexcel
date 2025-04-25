<?php

namespace EasyExcelTests\Metadata;

use EasyExcel\Metadata\Cell;
use EasyExcel\Metadata\Row;
use PHPUnit\Framework\TestCase;

class RowTest extends TestCase
{
    public function test_rowIndex()
    {
        $row = new Row();
        $this->assertNull($row->getRowIndex());
        $row->setRowIndex(1);
        $this->assertEquals(1, $row->getRowIndex());
    }

    public function test_cells()
    {
        $row = new Row();
        $this->assertCount(0, $row->getCells());
        $row->setCells([1, 2, new Cell(3)]);
        $this->assertCount(3, $row->getCells());
    }

    public function test_toArray()
    {
        $row = new Row([1, 2, 3]);
        $this->assertEquals([1, 2, 3], $row->toArray());
        $this->assertEquals([1, 2, 3], $row->toArray(true));
    }
}