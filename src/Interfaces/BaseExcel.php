<?php

namespace EasyExcel\Interfaces;

use EasyExcel\Metadata\Style;

interface BaseExcel
{
    /**
     * Add cellXf.
     *
     * @param \EasyExcel\Metadata\Style $style
     * @param int|null $index
     * @return $this
     */
    public function addCellXf(Style $style, ?int $index = null): self;

    /**
     * Get cellXf.
     *
     * @param int $index
     * @return \EasyExcel\Metadata\Style|null
     */
    public function getCellXf(int $index): ?Style;

    /**
     * Get cellXf.
     *
     * @param string $hashCode
     * @return \EasyExcel\Metadata\Style|null
     */
    public function getCellXfByHashCode(string $hashCode): ?Style;

    /**
     * Get cellXfs.
     *
     * @return \EasyExcel\Metadata\Style[]
     */
    public function getCellXfs(): array;

    /**
     * Get cellXfs count.
     *
     * @return int
     */
    public function getCellXfsCount(): int;

    /**
     * Add shared string.
     *
     * @param string $string
     * @param int|null $index
     * @return $this
     */
    public function addSharedString(string $string, ?int $index = null): self;

    /**
     * Get shared string by index.
     *
     * @param int $index
     * @return string|null
     */
    public function getSharedString(int $index): ?string;

    /**
     * Get shared strings.
     *
     * @return array
     */
    public function getSharedStrings(): array;

    /**
     * Get shared strings count.
     *
     * @return int
     */
    public function getSharedStringsCount(): int;

    /**
     * Close excel.
     *
     * @return void
     */
    public function close(): void;
}