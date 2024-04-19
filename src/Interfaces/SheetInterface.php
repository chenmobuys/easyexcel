<?php

namespace EasyExcel\Interfaces;

use EasyExcel\Metadata\AutoFilter;
use EasyExcel\Metadata\Hyperlink;

interface SheetInterface
{
    /**
     * Get excel.
     *
     * @return ExcelInterface
     */
    public function getExcel(): ExcelInterface;

    /**
     * Get sheet name.
     *
     * @return ?string
     */
    public function getName(): ?string;

    /**
     * Set sheet name.
     *
     * @param  string  $name
     *
     * @return $this
     */
    public function setName(string $name): SheetInterface;

    /**
     * Get sheet index
     *
     * @return ?int
     */
    public function getIndex(): ?int;

    /**
     * Set sheet index.
     *
     * @param  int  $index
     *
     * @return $this
     */
    public function setIndex(int $index): SheetInterface;

    /**
     * Get total rows.
     *
     * @return int
     */
    public function getTotalRows(): int;

    /**
     * Set total rows.
     *
     * @param  int  $totalRows
     *
     * @return $this
     */
    public function setTotalRows(int $totalRows): SheetInterface;

    /**
     * Get total columns.
     *
     * @return int
     */
    public function getTotalColumns(): int;

    /**
     * Set total columns.
     *
     * @param  int  $totalColumns
     *
     * @return $this
     */
    public function setTotalColumns(int $totalColumns): SheetInterface;

    /**
     * Get last column index.
     *
     * @return int
     */
    public function getLastColumnIndex(): int;

    /**
     * Get last column letter.
     *
     * @return string|null
     */
    public function getLastColumnLetter(): ?string;

    /**
     * Get auto filter.
     *
     * @return ?AutoFilter
     */
    public function getAutoFilter(): ?AutoFilter;

    /**
     * Set auto filter.
     *
     * @param  ?\EasyExcel\Metadata\AutoFilter  $autoFilter
     *
     * @return $this
     */
    public function setAutoFilter(?AutoFilter $autoFilter = null): SheetInterface;

    /**
     * Get hyperlink.
     *
     * @param  string  $coordinate
     *
     * @return \EasyExcel\Metadata\Hyperlink|null
     */
    public function getHyperlink(string $coordinate): ?Hyperlink;

    /**
     * Set hyperlink.
     *
     * @param  string                              $coordinate
     * @param  \EasyExcel\Metadata\Hyperlink|null  $hyperlink
     *
     * @return $this
     */
    public function setHyperlink(string $coordinate, ?Hyperlink $hyperlink): SheetInterface;


    /**
     * @param  string  $coordinate
     *
     * @return bool
     */
    public function hasHyperlink(string $coordinate): bool;

    /**
     * Get hyperlinks.
     *
     * @return Hyperlink[]
     */
    public function getHyperlinks(): array;

    /**
     * Set hyperlinks.
     *
     * @return $this
     */
    public function setHyperlinks(array $hyperlinks): SheetInterface;

    /**
     * Get merge cell.
     *
     * @param  string  $coordinate
     *
     * @return string|null
     */
    public function getMergeCell(string $coordinate): ?string;

    /**
     * Set merge cell.
     *
     * @param  string       $coordinate
     * @param  string|null  $mergeCell
     *
     * @return $this
     */
    public function setMergeCell(string $coordinate, ?string $mergeCell): SheetInterface;

    /**
     * Get merge cells.
     *
     * @return array
     */
    public function getMergeCells(): array;

    /**
     * Set merge cells.
     *
     * @param  array  $mergeCells
     *
     * @return $this;
     */
    public function setMergeCells(array $mergeCells): SheetInterface;
}