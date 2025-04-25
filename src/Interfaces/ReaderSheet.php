<?php

namespace EasyExcel\Interfaces;

use EasyExcel\Metadata\Cell;

interface ReaderSheet extends BaseSheet
{
    /**
     * Get reader excel.
     *
     * @return \EasyExcel\Interfaces\ReaderExcel
     */
    public function getExcel(): ReaderExcel;

    /**
     * Get cell.
     *
     * @param string $coordinate
     * @return \EasyExcel\Metadata\Cell|null
     */
    public function getCell(string $coordinate): ?Cell;

    /**
     * Get row iterator.
     *
     * @param int $startRow
     * @param int|null $endRow
     * @return \EasyExcel\Interfaces\ReaderRow
     */
    public function getRowIterator(int $startRow = 1, ?int $endRow = null): ReaderRow;

    /**
     * @param bool $formatValue
     * @param bool $formatDate
     * @param int $startRow
     * @param int|null $endRow
     * @return array
     */
    public function toArray(bool $formatValue = false, bool $formatDate = true, int $startRow = 1, ?int $endRow = null): array;
}
