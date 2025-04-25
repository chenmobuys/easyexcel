<?php

namespace EasyExcel\Readers\Xlsx;

use EasyExcel\Helpers\Coordinate;
use EasyExcel\Interfaces\ReaderSheet;
use EasyExcel\Metadata\Cell;
use EasyExcel\Metadata\Row;
use EasyExcel\Readers\BaseReaderRow;
use EasyExcel\Settings;
use XMLReader;

class ReaderRow extends BaseReaderRow
{
    /**
     * @var string
     */
    protected $handler;

    /**
     * @var \XMLReader
     */
    protected $xml;

    /**
     * @param string $handler
     * @param \EasyExcel\Interfaces\ReaderSheet $sheet
     * @param int $startRow
     * @param int|null $endRow
     */
    public function __construct(string $handler, ReaderSheet $sheet, int $startRow = 1, ?int $endRow = null)
    {
        parent::__construct($sheet, $startRow, $endRow);
        $this->handler = $handler;
    }

    /**
     * @return \EasyExcel\Metadata\Row
     */
    public function current(): Row
    {
        $this->row = $this->getEmptyRow();

        $rowNumber = $this->xml->getAttribute('r');
        if ($this->xml->isEmptyElement || $this->position < $rowNumber - 1) {
            return parent::current();
        }

        // Read until cell node.
        do {
            $this->xml->read();
        } while ($this->xml->localName != Reader::EL_C);

        do {
            $valueType = $this->xml->getAttribute('t');
            $coordinate = $this->xml->getAttribute('r');
            $xfIndex = (int) $this->xml->getAttribute('s');
            $columnLetter = preg_replace('/\d+/', '', $coordinate);
            $columnIndex = Coordinate::columnIndexFromColumnLetter($columnLetter);
            if(!isset($this->row[$columnIndex])) {
                $this->row[$columnIndex] = new Cell(null, $columnIndex);
            }
            $this->row[$columnIndex]->setXfIndex($xfIndex);
            if (!$this->xml->isEmptyElement) {
                while ($this->xml->read()) {
                    if ($this->xml->localName == Reader::EL_C) {
                        break;
                    }
                    if ($this->xml->nodeType == XMLReader::END_ELEMENT) {
                        continue;
                    }
                    switch ($this->xml->localName) {
                        case Reader::EL_V:
                            $value = $this->xml->readString();
                            if ($valueType == 's') {
                                $value = $this->sheet->getExcel()->getSharedString((int) $value);
                            } elseif (is_float($value) || is_numeric($value)) {
                                $value = 0 + $value;
                            }
                            if (isset($this->row[$columnIndex])) {
                                $this->row[$columnIndex]->setValue($value);
                            }
                            break;
                        case Reader::EL_IS:
                            $value = $this->xml->readString();
                            $this->row[$columnIndex]->setValue($value);
                            break;
                        case Reader::EL_F:
                            $value = $this->xml->readString();
                            $this->row[$columnIndex]->setFormulaValue($value);
                            break;
                    }
                }
            }
        } while ($this->xml->next() && $this->xml->localName == Reader::EL_C);

        return parent::current();
    }

    /**
     * @return void
     */
    public function next(): void
    {
        $rowNumber = $this->xml->getAttribute('r');
        if ($this->position >= $rowNumber - 1) {
            $this->xml->next();
        }

        parent::next();
    }

    /**
     * @return void
     */
    public function rewind(): void
    {
        parent::rewind();

        $this->xml = new XMLReader();
        $this->xml->open($this->handler, null, Settings::getLibXmlLoaderOptions());

        while ($this->xml->read()) {
            if ($this->xml->localName == Reader::EL_ROW) {
                break;
            }
        }

        if ($this->startRow > 1) {
            do {
                if ($this->startRow == $this->position + 1) {
                    break;
                }
                $rowNumber = $this->xml->getAttribute('r');
                if ($this->startRow < $rowNumber) {
                    $this->position = $this->startRow - 1;
                    break;
                }
                $this->position++;
            } while ($this->xml->next() && $this->xml->localName == Reader::EL_ROW);
        }
    }

    /**
     * Close row iterator.
     *
     * @return void
     */
    public function close(): void
    {
        $this->xml->close();
    }
}
