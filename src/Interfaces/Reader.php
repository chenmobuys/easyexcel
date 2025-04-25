<?php

namespace EasyExcel\Interfaces;

/**
 * @mixin \EasyExcel\Interfaces\ReaderExcel
 */
interface Reader
{
    /**
     * Determine whether the file is readable.
     *
     * @param string $filename
     * @return bool
     */
    public function readable(string $filename): bool;

    /**
     * Load file.
     *
     * @param string $filename
     * @return \EasyExcel\Interfaces\Reader
     */
    public function load(string $filename): Reader;

    /**
     * Close file.
     *
     * @return void
     */
    public function close(): void;

    /**
     * Get excel.
     *
     * @return \EasyExcel\Interfaces\ReaderExcel
     */
    public function getExcel(): ?ReaderExcel;

    /**
     * Get row iterator.
     *
     * @param \EasyExcel\Interfaces\ReaderSheet $sheet
     * @param int $startRow
     * @param int|null $endRow
     * @return \EasyExcel\Interfaces\ReaderRow
     */
    public function getRowIterator(ReaderSheet $sheet, int $startRow = 1, ?int $endRow = null): ReaderRow;
}