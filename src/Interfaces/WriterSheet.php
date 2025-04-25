<?php

namespace EasyExcel\Interfaces;

interface WriterSheet extends BaseSheet
{
    /**
     * Get excel.
     *
     * @return \EasyExcel\Interfaces\WriterExcel
     */
    public function getExcel(): WriterExcel;

    /**
     * Get row writer.
     *
     * @return \EasyExcel\Interfaces\WriterRow
     */
    public function getRowWriter(): WriterRow;
}