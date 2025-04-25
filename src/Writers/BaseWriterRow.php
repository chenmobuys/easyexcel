<?php

namespace EasyExcel\Writers;

use EasyExcel\Interfaces\Writer;
use EasyExcel\Interfaces\WriterRow;
use EasyExcel\Interfaces\WriterSheet;
use EasyExcel\Metadata\Row;
use EasyExcel\Metadata\Style;

abstract class BaseWriterRow implements WriterRow
{
    /**
     * @var \EasyExcel\Interfaces\Writer
     */
    protected $writer;

    /**
     * @var \EasyExcel\Interfaces\WriterSheet
     */
    protected $sheet;

    /**
     * @var int
     */
    protected $rowIndex = 0;

    /**
     * @param \EasyExcel\Interfaces\Writer $writer
     * @param \EasyExcel\Interfaces\WriterSheet $sheet
     */
    public function __construct(Writer $writer, WriterSheet $sheet)
    {
        $this->writer = $writer;
        $this->sheet = $sheet;
    }

    /**
     * @param array $rows
     * @param \EasyExcel\Metadata\Style|null $style
     * @return void
     */
    public function writes(array $rows, ?Style $style = null): void
    {
        $firstRow = reset($rows);
        $rows = (is_array($firstRow) || $firstRow instanceof Row) ? $rows : [$rows];

        foreach ($rows as $row) {
            if (is_array($row)) {
                $this->writeArray($row, $style);
            } else if ($row instanceof Row) {
                $this->writeRow($row, $style);
            }
            $this->rowIndex++;
        }
    }

    /**
     * Write row.
     *
     * @param \EasyExcel\Metadata\Row $row
     * @param \EasyExcel\Metadata\Style|null $style
     * @return void
     */
    abstract protected function writeRow(Row $row, ?Style $style = null): void;

    /**
     * Write array.
     *
     * @param array $row
     * @param \EasyExcel\Metadata\Style|null $style
     * @return void
     */
    abstract protected function writeArray(array $row, ?Style $style = null): void;
}