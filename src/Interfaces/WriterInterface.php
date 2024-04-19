<?php

namespace EasyExcel\Interfaces;

use EasyExcel\Metadata\Style;

interface WriterInterface extends ExcelInterface
{
    /**
     * Determine whether the file is writeable.
     *
     * @param  string  $filename
     *
     * @return bool
     */
    public static function writeable(string $filename): bool;

    /**
     * Open file.
     *
     * @param  string  $filename
     *
     * @return $this
     */
    public static function open(string $filename): WriterInterface;

    /**
     * Add row.
     *
     * @param  array       $row
     * @param  Style|null  $style
     *
     * @return $this
     */
    public function addRow(array $row, ?Style $style = null): WriterInterface;

    /**
     * Add rows.
     *
     * @param  array       $rows
     * @param  Style|null  $style
     *
     * @return $this
     */
    public function addRows(array $rows, ?Style $style = null): WriterInterface;

    /**
     * Close and save file.
     *
     * @return void
     */
    public function close(): void;
}
