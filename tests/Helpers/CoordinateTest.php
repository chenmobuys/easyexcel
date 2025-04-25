<?php

namespace EasyExcelTests\Helpers;

use EasyExcel\Helpers\Coordinate;
use PHPUnit\Framework\TestCase;

class CoordinateTest extends TestCase
{
    public function test_columnIndexFromColumnLetter(): void
    {
        $this->assertEquals(0, Coordinate::columnIndexFromColumnLetter('A'));
        $this->assertNull(Coordinate::columnIndexFromColumnLetter(']'));
    }

    public function test_columnLetterFromColumnIndex(): void
    {
        $this->assertEquals('A', Coordinate::columnLetterFromColumnIndex(0));
    }

    public function test_rowNumberAndColumnLetterFromCoordinate(): void
    {
        $this->assertEquals([1, 'A'], Coordinate::rowNumberAndColumnLetterFromCoordinate('A1'));
    }

    public function test_coordinateFromRowNumberAndColumnLetter(): void
    {
        $this->assertEquals('A1', Coordinate::coordinateFromRowNumberAndColumnLetter(1, 'A'));
    }

    public function test_rowIndexAndColumnIndexFromCoordinate(): void
    {
        $this->assertEquals([0, 0], Coordinate::rowIndexAndColumnIndexFromCoordinate('A1'));
    }

    public function test_coordinateFromRowIndexAndColumnIndex(): void
    {
        $this->assertEquals('A1', Coordinate::coordinateFromRowIndexAndColumnIndex(0, 0));
    }

    public function test_coordinatesFromRange(): void
    {
        $this->assertEquals(['A1'], Coordinate::coordinatesFromRange('A1'));
        $this->assertEquals(['A1', 'B1', 'A2', 'B2'], Coordinate::coordinatesFromRange('A1:B2'));
    }

    public function test_coordinateIsInRange(): void
    {
        $this->assertTrue(Coordinate::coordinateIsInRange('A1', 'A1:B1'));
        $this->assertFalse(Coordinate::coordinateIsInRange('A3', 'A1:B1'));
    }
}
