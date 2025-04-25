<?php

namespace EasyExcel\Writers\Xlsx;

use EasyExcel\Helpers\Coordinate;
use EasyExcel\Interfaces\Writer;
use EasyExcel\Interfaces\WriterSheet;
use EasyExcel\Metadata\Cell;
use EasyExcel\Metadata\Row;
use EasyExcel\Metadata\Style;
use EasyExcel\Writers\BaseWriterRow;
use EasyExcel\Writers\Xlsx\Parts\Sheet;

class WriterRow extends BaseWriterRow
{
    /**
     * @var \EasyExcel\Writers\Xlsx\Parts\Sheet
     */
    protected $handler;

    /**
     * @param \EasyExcel\Writers\Xlsx\Parts\Sheet $handler
     * @param \EasyExcel\Interfaces\Writer $writer
     * @param \EasyExcel\Interfaces\WriterSheet $sheet
     */
    public function __construct(Sheet $handler, Writer $writer, WriterSheet $sheet)
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
        $this->handler->getXml()->startElement('row');
        $this->handler->getXml()->writeAttribute('r', $this->rowIndex + 1);

        foreach ($row->getCells() as $columnIndex => $cell) {
            $this->writeCellObject($columnIndex, $cell, $style);
        }

        $this->handler->getXml()->endElement();
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
        $this->handler->getXml()->startElement('row');
        $this->handler->getXml()->writeAttribute('r', $this->rowIndex + 1);

        foreach ($row as $columnIndex => $cell) {
            if ($cell instanceof Cell) {
                $this->writeCellObject($columnIndex, $cell, $style);
            } else {
                $this->writeCellString($columnIndex, $cell, $style);
            }
        }

        $this->handler->getXml()->endElement();
    }

    /**
     * Write cell object.
     *
     * @param int $columnIndex
     * @param \EasyExcel\Metadata\Cell $cell
     * @param \EasyExcel\Metadata\Style|null $style
     * @return void
     */
    protected function writeCellObject(int $columnIndex, Cell $cell, ?Style $style): void
    {
        $cell->setRowIndex($this->rowIndex)->setColumnIndex($columnIndex);
        $this->handler->getXml()->startElement('c');
        $this->handler->getXml()->writeAttribute('r', $cell->getCoordinate());

        $xfIndex = 0;
        if ($style || $cell->hasStyle()) {
            $style = $style ?: $cell->getStyle();
            $existsStyle = $this->sheet->getExcel()->getCellXfByHashCode($style->getHashCode());
            if (!$existsStyle) {
                $this->sheet->getExcel()->addCellXf($style);
            } else {
                $style = $existsStyle;
            }
            $xfIndex = $style->getIndex();
        }
        $this->handler->getXml()->writeAttribute('s', $xfIndex);

        $stringFormats = [Style\Format::FORMAT_TEXT, Style\Format::FORMAT_GENERAL];
        if ($style && !in_array($style->getFormat()->getFormatCode(), $stringFormats)) {
            $this->handler->getXml()->writeAttribute('t', 'n');
            $this->handler->getXml()->writeElement('v', $cell->getValue(false));
        } else if ($cell->getFormulaValue()) {
            $this->handler->getXml()->writeAttribute('t', 'str');
            $this->handler->getXml()->writeElement('f', $cell->getFormulaValue());
            $this->handler->getXml()->writeElement('v', $cell->getValue(false));
        } else {
            $this->handler->getXml()->writeAttribute('t', 'inlineStr');
            $this->handler->getXml()->startElement('is');
            $this->handler->getXml()->writeElement('t', $cell->getValue(false));
            $this->handler->getXml()->endElement();
        }

        if ($cell->hasHyperlink()) {
            $this->sheet->setHyperlink($cell->getCoordinate(), $cell->getHyperlink());
        }

        if ($cell->getMergeCell()) {
            $this->sheet->setMergeCell($cell->getCoordinate(), $cell->getMergeCell());
        }

        $this->handler->getXml()->endElement();
    }

    /**
     * Write cell string.
     *
     * @param int $columnIndex
     * @param string|null $cellValue
     * @param \EasyExcel\Metadata\Style|null $style
     * @return void
     */
    protected function writeCellString(int $columnIndex, ?string $cellValue, ?Style $style): void
    {
        $coordinate = Coordinate::coordinateFromRowIndexAndColumnIndex($this->rowIndex, $columnIndex);
        $this->handler->getXml()->startElement('c');
        $this->handler->getXml()->writeAttribute('r', $coordinate);

        $xfIndex = 0;
        if ($style) {
            $existsStyle = $this->sheet->getExcel()->getCellXfByHashCode($style->getHashCode());
            if (!$existsStyle) {
                $this->sheet->getExcel()->addCellXf($style);
            } else {
                $style = $existsStyle;
            }
            $xfIndex = $style->getIndex();
        }
        $this->handler->getXml()->writeAttribute('s', $xfIndex);

        $stringFormats = [Style\Format::FORMAT_TEXT, Style\Format::FORMAT_GENERAL];
        if ($style && !in_array($style->getFormat()->getFormatCode(), $stringFormats)) {
            $this->handler->getXml()->writeAttribute('t', 'n');
            $this->handler->getXml()->writeElement('v', $cellValue);
        } else {
            $this->handler->getXml()->writeAttribute('t', 'inlineStr');
            $this->handler->getXml()->startElement('is');
            $this->handler->getXml()->writeElement('t', $cellValue);
            $this->handler->getXml()->endElement();
        }

        $this->handler->getXml()->endElement();
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
