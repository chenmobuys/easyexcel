<?php

namespace EasyExcel\Writers;

use EasyExcel\Interfaces\SheetInterface;
use EasyExcel\Interfaces\WriterRowInterface;
use EasyExcel\Metadata\Row;
use EasyExcel\Metadata\Style;

abstract class WriterRow implements WriterRowInterface
{
    /**
     * @var SheetInterface
     */
    protected $sheet;

    /**
     * @var int
     */
    protected $rowIndex = 0;

    /**
     * @param  SheetInterface  $sheet
     */
    public function __construct(SheetInterface $sheet)
    {
        $this->sheet = $sheet;
    }

    /**
     * @param  array  $rows
     * @param  \EasyExcel\Metadata\Style|null  $style
     * @return void
     */
    public function writes(array $rows, ?Style $style = null): void
    {
        $firstRow = reset($rows);
        $rows = (is_array($firstRow) || $firstRow instanceof Row) ? $rows : [$rows];

        foreach ($rows as $row) {
            if (is_array($row)) {
                $this->writeArray($row, $style);
            } else {
                if ($row instanceof Row) {
                    $this->writeRow($row, $style);
                }
            }
            $this->rowIndex++;
        }
    }

    /**
     * Write row.
     *
     * @param  \EasyExcel\Metadata\Row  $row
     * @param  \EasyExcel\Metadata\Style|null  $style
     * @return void
     */
    abstract protected function writeRow(Row $row, ?Style $style = null): void;

    /**
     * Write array.
     *
     * @param  array  $row
     * @param  \EasyExcel\Metadata\Style|null  $style
     * @return void
     */
    abstract protected function writeArray(array $row, ?Style $style = null): void;
}
