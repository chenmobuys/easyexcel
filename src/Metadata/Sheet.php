<?php

namespace EasyExcel\Metadata;

use EasyExcel\Helpers\Coordinate;
use EasyExcel\Interfaces\BaseSheet;

abstract class Sheet implements BaseSheet
{
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

    /**
     * Get sheet name.
     *
     * @return ?string
     */
    public function getName(): ?string
    {
        return $this->name;
    }

    /**
     * Set sheet name.
     *
     * @param string $name
     * @return $this
     */
    public function setName(string $name): BaseSheet
    {
        $this->name = $name;

        return $this;
    }


    /**
     * Get sheet index
     *
     * @return ?int
     */
    public function getIndex(): ?int
    {
        return $this->index;
    }

    /**
     * Set sheet index.
     *
     * @param int $index
     * @return $this
     */
    public function setIndex(int $index): BaseSheet
    {
        $this->index = $index;

        return $this;
    }

    /**
     * Get total rows.
     *
     * @return int
     */
    public function getTotalRows(): int
    {
        return $this->totalRows;
    }

    /**
     * Set total rows.
     *
     * @param int $totalRows
     * @return $this
     */
    public function setTotalRows(int $totalRows): BaseSheet
    {
        $this->totalRows = $totalRows;

        return $this;
    }

    /**
     * Get total columns.
     *
     * @return int
     */
    public function getTotalColumns(): int
    {
        return $this->totalColumns;
    }

    /**
     * Set total columns.
     *
     * @param int $totalColumns
     * @return $this
     */
    public function setTotalColumns(int $totalColumns): BaseSheet
    {
        $this->totalColumns = $totalColumns;

        return $this;
    }

    /**
     * Get last column index.
     *
     * @return int
     */
    public function getLastColumnIndex(): int
    {
        return $this->totalColumns > 0 ? $this->totalColumns - 1 : 0;
    }

    /**
     * Get last column letter.
     *
     * @return string|null
     */
    public function getLastColumnLetter(): ?string
    {
        return $this->totalColumns > 0 ? Coordinate::columnLetterFromColumnIndex($this->totalColumns - 1) : null;
    }

    /**
     * Get auto filter.
     *
     * @return ?AutoFilter
     */
    public function getAutoFilter(): ?AutoFilter
    {
        return $this->autoFilter;
    }

    /**
     * Set auto filter.
     *
     * @param  ?\EasyExcel\Metadata\AutoFilter $autoFilter
     * @return $this
     */
    public function setAutoFilter(?AutoFilter $autoFilter = null): BaseSheet
    {
        $this->autoFilter = $autoFilter;

        return $this;
    }

    /**
     * Get hyperlink.
     *
     * @param string $coordinate
     * @return \EasyExcel\Metadata\Hyperlink
     */
    public function getHyperlink(string $coordinate): Hyperlink
    {
        if (!isset($this->hyperlinks[$coordinate])) {
            $this->hyperlinks[$coordinate] = new Hyperlink();
        }

        return $this->hyperlinks[$coordinate];
    }

    /**
     * Set hyperlink.
     *
     * @param string $coordinate
     * @param \EasyExcel\Metadata\Hyperlink|null $hyperlink
     * @return $this
     */
    public function setHyperlink(string $coordinate, ?Hyperlink $hyperlink): BaseSheet
    {
        if (is_null($hyperlink)) {
            unset($this->hyperlinks[$coordinate]);
        } else {
            $this->hyperlinks[$coordinate] = $hyperlink;
        }

        return $this;
    }


    /**
     * @param string $coordinate
     * @return bool
     */
    public function hasHyperlink(string $coordinate): bool
    {
        return ($this->hyperlinks[$coordinate] ?? null) instanceof Hyperlink;
    }

    /**
     * Get hyperlinks.
     *
     * @return Hyperlink[]
     */
    public function getHyperlinks(): array
    {
        return $this->hyperlinks;
    }

    /**
     * Set hyperlinks.
     *
     * @return $this
     */
    public function setHyperlinks(array $hyperlinks): BaseSheet
    {
        $this->hyperlinks = $hyperlinks;

        return $this;
    }

    /**
     * Get merge cell.
     *
     * @param string $coordinate
     * @return string|null
     */
    public function getMergeCell(string $coordinate): ?string
    {
        return $this->mergeCells[$coordinate] ?? null;
    }

    /**
     * Set merge cell.
     *
     * @param string $coordinate
     * @param string|null $mergeCell
     * @return $this
     */
    public function setMergeCell(string $coordinate, ?string $mergeCell): BaseSheet
    {
        if (is_null($mergeCell)) {
            unset($this->mergeCells[$coordinate]);
        } else {
            $this->mergeCells[$coordinate] = $mergeCell;
        }

        return $this;
    }

    /**
     * Get merge cells.
     *
     * @return array
     */
    public function getMergeCells(): array
    {
        return $this->mergeCells;
    }

    /**
     * Set merge cells.
     *
     * @param array $mergeCells
     * @return $this;
     */
    public function setMergeCells(array $mergeCells): BaseSheet
    {
        $this->mergeCells = $mergeCells;

        return $this;
    }
}