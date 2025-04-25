<?php

namespace EasyExcel\Interfaces;

/**
 * @mixin \EasyExcel\Interfaces\WriterExcel
 */
interface Writer
{
    /**
     * Determine whether the file is writeable.
     *
     * @param string $filename
     * @return bool
     */
    public function writeable(string $filename): bool;

    /**
     * Open file.
     *
     * @param string $filename
     * @return \EasyExcel\Interfaces\Writer
     */
    public function open(string $filename): Writer;

    /**
     * Close and save file.
     *
     * @return void
     */
    public function close(): void;

    /**
     * Get excel.
     *
     * @return \EasyExcel\Interfaces\WriterExcel
     */
    public function getExcel(): ?WriterExcel;

    /**
     * Get row writer.
     *
     * @param \EasyExcel\Interfaces\WriterSheet $sheet
     * @return \EasyExcel\Interfaces\WriterRow
     */
    public function getRowWriter(WriterSheet $sheet): WriterRow;
}