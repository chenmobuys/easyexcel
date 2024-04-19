<?php

namespace EasyExcel\Metadata;

use EasyExcel\Helpers\Coordinate;
use EasyExcel\Metadata\Style\Formatter;

class Cell
{
    /**
     * @var ?int
     */
    private $rowIndex;

    /**
     * @var ?int
     */
    private $columnIndex;

    /**
     * @var int
     */
    private $xfIndex = 0;

    /**
     * @var string
     */
    private $value;

    /**
     * @var string
     */
    private $formulaValue;

    /**
     * @var string
     */
    private $formattedValue;

    /**
     * @var string
     */
    private $mergeCell;

    /**
     * @var \EasyExcel\Metadata\Hyperlink
     */
    private $hyperlink;

    /**
     * @var \EasyExcel\Metadata\Style
     */
    private $style;

    /**
     * @param  string|null  $value
     * @param  Style|null  $style
     * @param  int|null  $rowIndex
     * @param  int|null  $columnIndex
     */
    public function __construct(
        ?string $value = null,
        ?Style $style = null,
        ?int $rowIndex = null,
        ?int $columnIndex = null,
    ) {
        $this->value = $value;
        $this->style = $style;
        $this->rowIndex = $rowIndex;
        $this->columnIndex = $columnIndex;
    }

    public function getCoordinate(): string
    {
        return Coordinate::columnLetterFromColumnIndex($this->columnIndex).($this->rowIndex + 1);
    }

    /**
     * @return int|null
     */
    public function getRowIndex(): ?int
    {
        return $this->rowIndex;
    }

    /**
     * @param  int  $rowIndex
     * @return \EasyExcel\Metadata\Cell
     */
    public function setRowIndex(int $rowIndex): Cell
    {
        $this->rowIndex = $rowIndex;

        return $this;
    }

    /**
     * @return int|null
     */
    public function getColumnIndex(): ?int
    {
        return $this->columnIndex;
    }

    /**
     * @param  int  $columnIndex
     * @return \EasyExcel\Metadata\Cell
     */
    public function setColumnIndex(int $columnIndex): Cell
    {
        $this->columnIndex = $columnIndex;

        return $this;
    }

    /**
     * @return int
     */
    public function getXfIndex(): int
    {
        return $this->xfIndex;
    }

    /**
     * @param  int  $xfIndex
     * @return \EasyExcel\Metadata\Cell
     */
    public function setXfIndex(int $xfIndex): Cell
    {
        $this->xfIndex = $xfIndex;

        return $this;
    }

    /**
     * @return string|null
     */
    public function getValue(): ?string
    {
        return $this->value;
    }

    /**
     * @param  string|null  $value
     * @return \EasyExcel\Metadata\Cell
     */
    public function setValue(?string $value): Cell
    {
        $this->value = $value;

        return $this;
    }

    /**
     * @return string
     */
    public function getFormulaValue(): ?string
    {
        return $this->formulaValue;
    }

    /**
     * @param  string  $formulaValue
     * @return \EasyExcel\Metadata\Cell
     */
    public function setFormulaValue(string $formulaValue): Cell
    {
        $this->formulaValue = $formulaValue;

        return $this;
    }

    /**
     * @return string|null
     */
    public function getFormattedValue(): ?string
    {
        if ($this->formattedValue) {
            return $this->formattedValue;
        }

        if ($this->style && is_numeric($this->value) && $this->style->getFormat()->getFormatCode()) {
            $this->formattedValue = Formatter::getFormattedValue((float) $this->value,
                $this->style->getFormat()->getFormatCode());
        } else {
            $this->formattedValue = $this->value;
        }

        return $this->formattedValue;
    }

    /**
     * @param  string  $formattedValue
     * @return \EasyExcel\Metadata\Cell
     */
    public function setFormattedValue(string $formattedValue): Cell
    {
        $this->formattedValue = $formattedValue;

        return $this;
    }

    /**
     * @return string|null
     */
    public function getMergeCell(): ?string
    {
        return $this->mergeCell;
    }

    /**
     * @param  string  $mergeCell
     * @return \EasyExcel\Metadata\Cell
     */
    public function setMergeCell(string $mergeCell): Cell
    {
        $this->mergeCell = $mergeCell;

        return $this;
    }

    /**
     * @return bool
     */
    public function hasHyperlink(): bool
    {
        return !is_null($this->hyperlink);
    }

    /**
     * @return \EasyExcel\Metadata\Hyperlink|null
     */
    public function getHyperlink(): ?Hyperlink
    {
        return $this->hyperlink;
    }

    /**
     * @param  \EasyExcel\Metadata\Hyperlink  $hyperlink
     * @return \EasyExcel\Metadata\Cell
     */
    public function setHyperlink(Hyperlink $hyperlink): Cell
    {
        $this->hyperlink = $hyperlink;

        return $this;
    }

    /**
     * @return bool
     */
    public function hasStyle(): bool
    {
        return !is_null($this->style);
    }

    /**
     * @return \EasyExcel\Metadata\Style
     */
    public function getStyle(): ?Style
    {
        return $this->style;
    }

    /**
     * @param  \EasyExcel\Metadata\Style  $style
     * @return \EasyExcel\Metadata\Cell
     */
    public function setStyle(Style $style): Cell
    {
        $this->style = $style;

        return $this;
    }
}
