<?php

namespace EasyExcel\Readers;

use EasyExcel\Interfaces\ReaderExcel;
use EasyExcel\Interfaces\ReaderRow;
use EasyExcel\Interfaces\ReaderSheet as ReaderSheetInterface;
use EasyExcel\Metadata\Cell;
use EasyExcel\Metadata\Sheet;

class ReaderSheet extends Sheet implements ReaderSheetInterface
{
    /**
     * Excel.
     *
     * @var \EasyExcel\Interfaces\ReaderExcel
     */
    protected $excel;

    /**
     * @var \EasyExcel\Interfaces\ReaderRow
     */
    protected $rowIterator;

    /**
     * Constructor.
     *
     * @param \EasyExcel\Interfaces\ReaderExcel $excel
     * @param string $name
     * @param int $index
     */
    public function __construct(ReaderExcel $excel, string $name, int $index)
    {
        $this->excel = $excel;
        $this->name = $name;
        $this->index = $index;
    }

    /**
     * Get reader excel.
     *
     * @return \EasyExcel\Interfaces\ReaderExcel
     */
    public function getExcel(): ReaderExcel
    {
        return $this->excel;
    }

    /**
     * Get cell.
     *
     * @param string $coordinate
     * @return \EasyExcel\Metadata\Cell|null
     */
    public function getCell(string $coordinate): ?Cell
    {
        $rowNumber = preg_replace('/[A-Z]+/', '', $coordinate);
        foreach ($this->getRowIterator($rowNumber, $rowNumber) as $rowItem) {
            foreach ($rowItem->getCells() as $cell) {
                if ($cell->getCoordinate() == $coordinate) {
                    return $cell;
                }
            }
        }
        return null;
    }

    /**
     * Get row iterator.
     *
     * @param int $startRow
     * @param int|null $endRow
     * @return \EasyExcel\Interfaces\ReaderRow
     */
    public function getRowIterator(int $startRow = 1, ?int $endRow = null): ReaderRow
    {
        $this->close();
        $endRow = ($endRow > 0 && $endRow <= $this->totalRows) ? $endRow : $this->totalRows;
        $this->rowIterator = $this->excel->getReader()->getRowIterator($this, $startRow, $endRow);

        return $this->rowIterator;
    }

    /**
     * @param bool $formatValue
     * @param bool $formatDate
     * @param int $startRow
     * @param int|null $endRow
     * @return array
     */
    public function toArray(bool $formatValue = false, bool $formatDate = true, int $startRow = 1, ?int $endRow = null): array
    {
        $data = [];
        foreach ($this->getRowIterator($startRow, $endRow) as $row) {
            $data[] = $row->toArray($formatValue, $formatDate);
        }

        return $data;
    }

    /**
     * Close sheet.
     *
     * @return void
     */
    public function close(): void
    {
        if ($this->rowIterator) {
            $this->rowIterator->close();
        }
    }

    /**
     * Clone sheet and clean rowIterator.
     *
     * @return void
     */
    public function __clone()
    {
        $this->rowIterator = null;
    }
}
