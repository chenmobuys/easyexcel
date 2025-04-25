<?php

namespace EasyExcel\Readers;

use EasyExcel\Helpers\Coordinate;
use EasyExcel\Interfaces\ReaderRow;
use EasyExcel\Interfaces\ReaderSheet;
use EasyExcel\Metadata\Cell;
use EasyExcel\Metadata\Row;

abstract class BaseReaderRow implements ReaderRow
{
    /**
     * @var \EasyExcel\Interfaces\ReaderSheet
     */
    protected $sheet;

    /**
     * @var \EasyExcel\Metadata\Cell[]
     */
    protected $row = [];

    /**
     * @var int
     */
    protected $position = 0;

    /**
     * @var int
     */
    protected $startRow = 1;

    /**
     * @var ?int
     */
    protected $endRow = null;

    /**
     * @param \EasyExcel\Interfaces\ReaderSheet $sheet
     * @param int $startRow
     * @param int|null $endRow
     */
    public function __construct(ReaderSheet $sheet, int $startRow = 1, ?int $endRow = null)
    {
        $this->sheet = $sheet;
        $this->startRow = $startRow;
        $this->endRow = $endRow;
    }

    /**
     * @return \EasyExcel\Metadata\Row
     */
    public function current(): Row
    {
        // Fill styles, hyperlinks, mergeCell.
        foreach ($this->row as $cell) {
            if ($style = $this->sheet->getExcel()->getCellXf($cell->getXfIndex())) {
                $cell->setStyle($style);
            }
            if ($this->sheet->hasHyperlink($cell->getCoordinate())) {
                $cell->setHyperlink($this->sheet->getHyperlink($cell->getCoordinate()));
            }
            foreach ($this->sheet->getMergeCells() as $mergeCell) {
                if (!Coordinate::coordinateIsInRange($cell->getCoordinate(), $mergeCell)) {
                    continue;
                }
                $cell->setMergeCell($mergeCell);
                break;
            }
        }

        return new Row($this->row, $this->position);
    }

    /**
     * @return void
     */
    public function next(): void
    {
        $this->position++;
    }

    /**
     * @return mixed|null
     */
    public function key(): int
    {
        return $this->position;
    }

    /**
     * @return bool
     */
    public function valid(): bool
    {
        return $this->position >= $this->startRow - 1 && $this->position < $this->endRow;
    }

    /**
     * @return void
     */
    public function rewind(): void
    {
        $this->position = 0;
    }

    /**
     * @return int
     */
    public function count(): int
    {
        return count($this->row);
    }

    /**
     * Get empty row.
     *
     * @return array
     */
    protected function getEmptyRow(): array
    {
        $row = [];

        for ($i = 0; $i < $this->sheet->getTotalColumns(); $i++) {
            $row[$i] = new Cell('', $this->position, $i);
        }

        return $row;
    }
}