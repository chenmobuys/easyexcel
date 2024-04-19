<?php

namespace EasyExcel\Metadata;

use EasyExcel\Helpers\Coordinate;
use EasyExcel\Interfaces\ExcelInterface;
use EasyExcel\Interfaces\SheetInterface;

class Sheet implements SheetInterface
{
    /**
     * @var \EasyExcel\Interfaces\ExcelInterface
     */
    protected $excel;

    /**
     * @var
     */
    protected $name;

    /**
     * @var
     */
    protected $index;

    /**
     * @var int
     */
    protected $totalRows = 0;

    /**
     * @var int
     */
    protected $totalColumns = 0;

    /**
     * @var AutoFilter
     */
    protected $autoFilter;

    /**
     * @var array
     */
    protected $hyperlinks = [];

    /**
     * @var array
     */
    protected $mergeCells = [];

    public function __construct(string $name, int $index, ExcelInterface $excel)
    {
        $this->name = $name;
        $this->index = $index;
        $this->excel = $excel;
    }

    public function getExcel(): ExcelInterface
    {
        return $this->excel;
    }

    public function getName(): ?string
    {
        return $this->name;
    }

    public function setName(string $name): SheetInterface
    {
        $this->name = $name;

        return $this;
    }

    public function getIndex(): ?int
    {
        return $this->index;
    }

    public function setIndex(int $index): SheetInterface
    {
        $this->index = $index;

        return $this;
    }

    public function getTotalRows(): int
    {
        return $this->totalRows;
    }

    public function setTotalRows(int $totalRows): SheetInterface
    {
        $this->totalRows = $totalRows;

        return $this;
    }

    public function getTotalColumns(): int
    {
        return $this->totalColumns;
    }

    public function setTotalColumns(int $totalColumns): SheetInterface
    {
        $this->totalColumns = $totalColumns;

        return $this;
    }

    public function getLastColumnIndex(): int
    {
        return $this->totalColumns > 0 ? $this->totalColumns - 1 : 0;
    }

    public function getLastColumnLetter(): ?string
    {
        return $this->totalColumns > 0 ? Coordinate::columnLetterFromColumnIndex($this->totalColumns - 1) : null;
    }

    public function getAutoFilter(): ?AutoFilter
    {
        return $this->autoFilter;
    }

    public function setAutoFilter(?AutoFilter $autoFilter = null): SheetInterface
    {
        $this->autoFilter = $autoFilter;

        return $this;
    }

    public function getHyperlink(string $coordinate): ?Hyperlink
    {
        return $this->hyperlinks[$coordinate] ?? null;
    }

    public function setHyperlink(string $coordinate, ?Hyperlink $hyperlink): SheetInterface
    {
        if (is_null($hyperlink)) {
            unset($this->hyperlinks[$coordinate]);
        } else {
            $this->hyperlinks[$coordinate] = $hyperlink;
        }

        return $this;
    }

    public function hasHyperlink(string $coordinate): bool
    {
        return ($this->hyperlinks[$coordinate] ?? null) instanceof Hyperlink;
    }

    public function getHyperlinks(): array
    {
        return $this->hyperlinks;
    }

    public function setHyperlinks(array $hyperlinks): SheetInterface
    {
        $this->hyperlinks = $hyperlinks;

        return $this;
    }

    public function getMergeCell(string $coordinate): ?string
    {
        return $this->mergeCells[$coordinate] ?? null;
    }

    public function setMergeCell(string $coordinate, ?string $mergeCell): SheetInterface
    {
        if (is_null($mergeCell)) {
            unset($this->mergeCells[$coordinate]);
        } else {
            $this->mergeCells[$coordinate] = $mergeCell;
        }

        return $this;
    }

    public function getMergeCells(): array
    {
        return $this->mergeCells;
    }

    public function setMergeCells(array $mergeCells): SheetInterface
    {
        $this->mergeCells = $mergeCells;

        return $this;
    }
}
