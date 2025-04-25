<?php

namespace EasyExcel\Readers\Ods;

use EasyExcel\Interfaces\ReaderSheet;
use EasyExcel\Metadata\Row;
use EasyExcel\Readers\BaseReaderRow;
use EasyExcel\Settings;
use XMLReader;

class ReaderRow extends BaseReaderRow
{
    /**
     * XML path.
     *
     * @var string
     */
    protected $handler;

    /**
     * @var \XMLReader
     */
    protected $xml;

    /**
     * Prepare for repeated row.
     *
     * @var Row|null
     */
    protected $lastRow;

    /**
     * @var int
     */
    protected $repeatedRows = 0;

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

        if ($this->repeatedRows > 0) {
            if (!$this->lastRow) {
                $this->lastRow = $this->getNextRow();
            }
            $this->row = $this->lastRow;
            return parent::current();
        }

        if ($this->xml->isEmptyElement) {
            return parent::current();
        }

        $numberRowsRepeated = ((int) $this->xml->getAttributeNS('number-rows-repeated', Reader::NS_TABLE)) ?: 1;

        $this->row = $this->getNextRow();

        if ($numberRowsRepeated > 1) {
            $this->lastRow = $this->row;
            $this->repeatedRows = $numberRowsRepeated;
        }

        return parent::current();
    }

    private function getNextRow(): array
    {
        $columnIndex = 0;
        $row = $this->getEmptyRow();

        $this->xml->read();
        do {
            $numberColumnsRepeated = ((int) $this->xml->getAttributeNS('number-columns-repeated', Reader::NS_TABLE)) ?: 1;
            if (!$this->xml->isEmptyElement) {
                $valueType = $this->xml->getAttributeNs('value-type', Reader::NS_OFFICE);
                $formulaValue = $this->xml->getAttributeNs('formula', Reader::NS_TABLE);
                $formattedValue = $this->xml->readString();
                switch ($valueType) {
                    case 'date':
                    case 'time':
                        $value = $this->xml->getAttributeNS($valueType . '-value', Reader::NS_OFFICE);
                        break;
                    default:
                        $value = $this->xml->getAttributeNs('value', Reader::NS_OFFICE) ?: $formattedValue;
                        break;
                }

                for ($i = 0; $i < $numberColumnsRepeated; $i++) {
                    if (!is_null($formulaValue)) {

                        $formulaValue = substr($formulaValue, strpos($formulaValue, ':=') + 2);
                        $temp = explode('"', $formulaValue);
                        $tKey = false;
                        foreach ($temp as &$tempValue) {
                            // Only replace in alternate array entries (i.e. non-quoted blocks)
                            if ($tKey = !$tKey) {
                                // Cell range reference in another sheet
                                $tempValue = preg_replace('/\[([^.]+)\.([^.]+):\.([^.]+)]/U', '$1!$2:$3', $tempValue);
                                // Cell reference in another sheet
                                $tempValue = preg_replace('/\[([^.]+)\.([^.]+)]/U', '$1!$2', $tempValue);
                                // Cell range reference
                                $tempValue = preg_replace('/\[\.([^.]+):\.([^.]+)]/U', '$1:$2', $tempValue);
                                // Simple cell reference
                                $tempValue = preg_replace('/\[\.([^.]+)]/U', '$1', $tempValue);
                                // Separator
                                $tempValue = str_replace(';', ',', $tempValue);
                            }
                        }
                        unset($tempValue);

                        // Then rebuild the formula string
                        $formulaValue = implode('"', $temp);

                        $row[$columnIndex + $i]->setFormulaValue($formulaValue);
                    }
                    if (!is_null($value)) {
                        $row[$columnIndex + $i]->setValue($value);
                    }
                    if (!is_null($formattedValue)) {
                        $row[$columnIndex + $i]->setFormattedValue($formattedValue);
                    }
                }

                while ($this->xml->read()) {
                    if ($this->xml->localName == Reader::EL_TABLE_CELL) {
                        break;
                    }
                }
            }
            $columnIndex += $numberColumnsRepeated;
        } while ($this->xml->next() && $this->xml->localName == Reader::EL_TABLE_CELL);

        return $row;
    }

    /**
     * @return void
     */
    public function next(): void
    {
        if ($this->repeatedRows > 0) {
            $this->repeatedRows--;
        }

        if ($this->repeatedRows == 0) {
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
        $sheetIndex = 0;
        while ($this->xml->read()) {
            if ($this->xml->localName != Reader::EL_TABLE) {
                continue;
            }
            do {
                if ($sheetIndex == $this->sheet->getIndex()) {
                    if (!$this->xml->isEmptyElement) {
                        while ($this->xml->read()) {
                            if ($this->xml->localName == Reader::EL_TABLE_ROW) {
                                break;
                            }
                        }
                    }
                    break 2;
                }
                $sheetIndex++;
            } while ($this->xml->next() && $this->xml->localName == Reader::EL_TABLE);
        }

        if ($this->startRow > 1) {
            do {
                if ($this->startRow == $this->position + 1) {
                    break;
                }
                $numberRowsRepeated = ((int) $this->xml->getAttributeNs('number-rows-repeated', Reader::NS_TABLE)) ?: 1;
                if ($numberRowsRepeated > 1) {
                    if ($this->startRow > $this->position + $numberRowsRepeated + 1) {
                        $this->position += $numberRowsRepeated;
                    } else {
                        // 2 , 0 , 3
                        $this->repeatedRows = $this->position + $numberRowsRepeated - $this->startRow;
                        $this->position = $this->startRow - 1;
                        break;
                    }
                } else {
                    $this->position++;
                }
            } while ($this->xml->next() && $this->xml->localName == Reader::EL_TABLE_ROW);
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