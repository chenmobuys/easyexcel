<?php

namespace EasyExcel\Writers\Csv;

use EasyExcel\Interfaces\Writer;
use EasyExcel\Interfaces\WriterSheet;
use EasyExcel\Metadata\Cell;
use EasyExcel\Metadata\Row;
use EasyExcel\Metadata\Style;
use EasyExcel\Writers\BaseWriterRow;
use SplFileObject;

class WriterRow extends BaseWriterRow
{
    /**
     * @var \SplFileObject
     */
    protected $handler;

    /**
     * @param \SplFileObject $handler
     * @param \EasyExcel\Interfaces\Writer $writer
     * @param \EasyExcel\Interfaces\WriterSheet $sheet
     */
    public function __construct(SplFileObject $handler, Writer $writer, WriterSheet $sheet)
    {
        parent::__construct($writer, $sheet);
        $this->handler = $handler;
    }

    /**
     * Write row.
     *
     * @param \EasyExcel\Metadata\Row $row
     * @param \EasyExcel\Metadata\Style|null $style
     * @return void
     */
    protected function writeRow(Row $row, ?Style $style = null): void
    {
        $this->handler->fputcsv($row->toArray());
    }

    /**
     * Write array.
     *
     * @param array $row
     * @param \EasyExcel\Metadata\Style|null $style
     * @return void
     */
    protected function writeArray(array $row, ?Style $style = null): void
    {
        $row = array_map(function ($cell) {
            return $cell instanceof Cell ? $cell->getValue() : $cell;
        }, $row);
        $this->handler->fputcsv($row);
    }

    /**
     * Close row writer.
     *
     * @return void
     */
    public function close(): void
    {
        $this->handler = null;
    }
}