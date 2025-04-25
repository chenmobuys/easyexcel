<?php

namespace EasyExcel\Metadata;

class Row
{

    /**
     * @var \EasyExcel\Metadata\Cell[]
     */
    private $cells;

    /**
     * @var ?int
     */
    private $rowIndex;

    /**
     * Constructor.
     *
     * @param array $cells
     * @param int|null $rowIndex
     */
    public function __construct(array $cells = [], ?int $rowIndex = null)
    {
        $this->rowIndex = $rowIndex;
        $this->setCells($cells);
    }

    /**
     * @return ?int
     */
    public function getRowIndex(): ?int
    {
        return $this->rowIndex;
    }

    /**
     * @param int $rowIndex
     * @return $this
     */
    public function setRowIndex(int $rowIndex): self
    {
        $this->rowIndex = $rowIndex;

        return $this;
    }

    /**
     * @return array
     */
    public function getCells(): array
    {
        return $this->cells;
    }

    /**
     * @param array $cells
     * @return $this
     */
    public function setCells(array $cells): self
    {
        $this->cells = [];

        foreach ($cells as $cell) {
            $this->cells[] = $cell instanceof Cell ? $cell : new Cell($cell);
        }

        return $this;
    }

    /**
     * @param bool $formatValue
     * @param bool $formatDate
     * @return array
     */
    public function toArray(bool $formatValue = false, bool $formatDate = true): array
    {
        return array_map(function (Cell $cell) use ($formatValue, $formatDate) {
            return $formatValue ? $cell->getFormattedValue() : $cell->getValue($formatDate);
        }, $this->cells);
    }
}
