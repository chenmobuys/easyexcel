<?php

namespace EasyExcel\Readers\Csv;

use EasyExcel\Interfaces\SheetInterface;
use EasyExcel\Metadata\Cell;
use EasyExcel\Metadata\Row;
use EasyExcel\Readers\ReaderRow;
use SplFileObject;

class CsvReaderRow extends ReaderRow
{
    /**
     * @var \SplFileObject
     */
    protected $handler;

    /**
     * @var string
     */
    protected $encoding;

    /**
     * @param  SplFileObject  $handler
     * @param  string  $encoding
     * @param  SheetInterface  $sheet
     * @param  int  $startRow
     * @param  int|null  $endRow
     */
    public function __construct(
        SplFileObject $handler,
        string $encoding,
        SheetInterface $sheet,
        int $startRow = 1,
        int $endRow = null
    ) {
        parent::__construct($sheet, $startRow, $endRow);
        $this->handler = $handler;
        $this->encoding = $encoding;
    }

    /**
     * @return \EasyExcel\Metadata\Row
     */
    public function current(): Row
    {
        $this->row = $this->getEmptyRow();

        foreach ((array) $this->handler->current() as $columnIndex => $cellValue) {
            $this->row[$columnIndex] = new Cell($cellValue, null, $this->position, $columnIndex);
        }

        return parent::current();
    }

    /**
     * @return void
     */
    public function rewind(): void
    {
        parent::rewind();

        $this->handler->rewind();

        while (true) {
            if ($this->position >= $this->startRow - 1) {
                break;
            }

            $this->next();
        }
    }

    /**
     * @return void
     */
    public function next(): void
    {
        parent::next();

        $this->handler->next();
    }

    /**
     * Close row iterator.
     *
     * @return void
     */
    public function close(): void
    {
        $this->handler = null;
    }
}
