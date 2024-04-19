<?php

namespace EasyExcel\Interfaces;

interface ReaderInterface extends ExcelInterface
{
    /**
     * Determine whether the file is readable.
     *
     * @param  string  $filename
     *
     * @return bool
     */
    public static function readable(string $filename): bool;

    /**
     * Load file.
     *
     * @param  string  $filename
     *
     * @return $this
     */
    public static function load(string $filename): ReaderInterface;

    /**
     * Get row iterator.
     *
     * @param  int       $startRow
     * @param  int|null  $endRow
     *
     * @return ReaderRowInterface
     */
    public function getRowIterator(int $startRow = 1, int $endRow = null): ReaderRowInterface;

    /**
     * Close file.
     *
     * @return void
     */
    public function close(): void;
}
