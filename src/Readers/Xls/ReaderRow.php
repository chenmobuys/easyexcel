<?php

namespace EasyExcel\Readers\Xls;

use EasyExcel\Interfaces\ReaderSheet;
use EasyExcel\Metadata\Row;
use EasyExcel\Readers\BaseReaderRow;

class ReaderRow extends BaseReaderRow
{
    /**
     * @var string
     */
    protected $handler;

    /**
     * @var int
     */
    protected $version;

    /**
     * @var string
     */
    protected $codepage;

    /**
     * @var array
     */
    protected $rowOffsets = [];

    /**
     * @param string $handler
     * @param string $codepage
     * @param int $version
     * @param array $rowOffsets
     * @param \EasyExcel\Interfaces\ReaderSheet $sheet
     * @param int $startRow
     * @param int|null $endRow
     */
    public function __construct(
        string      $handler,
        string      $codepage,
        int         $version,
        array       $rowOffsets,
        ReaderSheet $sheet,
        int         $startRow = 1,
        ?int         $endRow = null
    )
    {
        parent::__construct($sheet, $startRow, $endRow);
        $this->handler = $handler;
        $this->codepage = $codepage;
        $this->version = $version;
        $this->rowOffsets = $rowOffsets;
    }

    /**
     * @return \EasyExcel\Metadata\Row
     */
    public function current(): Row
    {
        $this->row = $this->getEmptyRow();

        $currentRowLength = 0;
        $currentRowOffset = $this->rowOffsets[$this->position] ?? null;
        if (!is_numeric($currentRowOffset)) {
            return parent::current();
        }

        // Find next valid row.
        if ($this->position < $this->endRow - 1) {
            $nextRowStep = 1;
            do {
                $nextRowOffset = $this->rowOffsets[$this->position + $nextRowStep] ?? null;
                $nextRowStep++;
            } while (!is_numeric($nextRowOffset));
        } else {
            $nextRowOffset = $this->rowOffsets[$this->position + 1] ?? null;
        }

        do {
            $position = $currentRowOffset + $currentRowLength;
            $code = Helper::getInt($position, $this->handler, 2);
            $length = Helper::getInt($position + 2, $this->handler, 2);
            $position += 4;

            if (is_numeric($nextRowOffset) && $position >= $nextRowOffset) {
                break;
            }

            switch ($code) {
                case Reader::RECORD_BLANK:
                    $columnIndex = Helper::getInt($position + 2, $this->handler, 2);
                    $xfIndex = Helper::getInt($position + 4, $this->handler, 2);
                    $this->row[$columnIndex]->setXfIndex($xfIndex);
                    break;
                case Reader::RECORD_MULBLANK:
                    $columnStartIndex = Helper::getInt($position + 2, $this->handler, 2);
                    $columnEndIndex = Helper::getInt($position + $length - 2, $this->handler, 2);
                    $columnCount = $columnEndIndex - $columnStartIndex + 1;
                    for ($i = 0; $i < $columnCount; $i++) {
                        $xfIndex = Helper::getInt($position + $i * 2 + 4, $this->handler, 2);
                        $this->row[$columnStartIndex + $i]->setXfIndex($xfIndex);
                    }
                    break;
                case Reader::RECORD_RK:
                    $columnIndex = Helper::getInt($position + 2, $this->handler, 2);
                    $xfIndex = Helper::getInt($position + 4, $this->handler, 2);
                    $rkNum = Helper::getInt($position + 6, $this->handler);
                    $value = Helper::getIEEE754($rkNum);
                    $this->row[$columnIndex]->setValue($value)->setXfIndex($xfIndex);
                    break;
                case Reader::RECORD_MULRK:
                    $columnStartIndex = Helper::getInt($position + 2, $this->handler, 2);
                    $columnEndIndex = Helper::getInt($position + $length - 2, $this->handler, 2);
                    $columnCount = $columnEndIndex - $columnStartIndex + 1;
                    for ($i = 0; $i < $columnCount; $i++) {
                        $xfIndex = Helper::getInt($position + $i * 6 + 4, $this->handler, 2);
                        $rkNum = Helper::getInt($position + ($i + 1) * 6, $this->handler);
                        $value = Helper::getIEEE754($rkNum);
                        $this->row[$columnStartIndex + $i]->setValue($value)->setXfIndex($xfIndex);
                    }
                    break;
                case Reader::RECORD_NUMBER:
                    $columnIndex = Helper::getInt($position + 2, $this->handler, 2);
                    $xfIndex = Helper::getInt($position + 4, $this->handler, 2);
                    $tmpValue = unpack("d", substr($this->handler, $position + 6, 8));
                    $value = current($tmpValue);
                    $this->row[$columnIndex]->setValue($value)->setXfIndex($xfIndex);
                    break;
                case Reader::RECORD_BOOLERR:
                    $columnIndex = Helper::getInt($position + 2, $this->handler, 2);
                    $xfIndex = Helper::getInt($position + 4, $this->handler, 2);
                    $value = Helper::getInt($position + 6, $this->handler, 1);
                    $isError = Helper::getInt($position + 7, $this->handler, 1);
                    $value = $isError ? ErrorCode::indexedCode($value) : ($value ? 'TRUE' : 'FALSE');
                    $this->row[$columnIndex]->setValue($value)->setXfIndex($xfIndex);
                    break;
                case Reader::RECORD_LABEL:
                    $columnIndex = Helper::getInt($position + 2, $this->handler, 2);
                    $xfIndex = Helper::getInt($position + 4, $this->handler, 2);
                    $data = substr($this->handler, 6, $length - 6);
                    if ($this->version == Reader::BIFF_VERSION_8) {
                        $value = Helper::readUnicodeStringLong($data, $this->codepage);
                    } else {
                        $value = Helper::readByteStringLong($data, $this->codepage);
                    }
                    $this->row[$columnIndex]->setValue($value)->setXfIndex($xfIndex);
                    break;
                case Reader::RECORD_LABELSST:
                    $columnIndex = Helper::getInt($position + 2, $this->handler, 2);
                    $xfIndex = Helper::getInt($position + 4, $this->handler, 2);
                    $sharedStringIndex = Helper::getInt($position + 6, $this->handler);
                    $value = $this->sheet->getExcel()->getSharedString($sharedStringIndex);
                    $this->row[$columnIndex]->setValue($value)->setXfIndex($xfIndex);
                    break;
                case Reader::RECORD_FORMULA:
                    $columnIndex = Helper::getInt($position + 2, $this->handler, 2);
                    $xfIndex = Helper::getInt($position + 4, $this->handler, 2);
                    $identifier = Helper::getInt($position + 6, $this->handler, 1);
                    if (
                        Helper::getInt($position + 12, $this->handler, 1) == 255
                        && Helper::getInt($position + 13, $this->handler, 1) == 255
                        && in_array($identifier, [0, 1, 2, 3])
                    ) {
                        $value = Helper::getInt($position + 8, $this->handler, 1) == 1;
                        switch ($identifier) {
                            case 0:
                                // Str formula. Result follows in a STRING record
                                // This row/col are stored to be referenced in that record
                                $position += $length;
                                $code = Helper::getInt($position, $this->handler, 2);
                                $length = Helper::getInt($position + 2, $this->handler, 2);
                                $position += 4;
                                if ($code == Reader::RECORD_STRING) {
                                    $data = substr($this->handler, $position);
                                    if ($this->version == Reader::BIFF_VERSION_8) {
                                        $value = Helper::readUnicodeStringLong($data, $this->codepage);
                                    } else {
                                        $value = Helper::readByteStringLong($data, $this->codepage);
                                    }
                                }
                                break;
                            case 1:
                                // Boolean formula. Result is in +2; 0=false,1=true
                                $value = $value == 1 ? 'TRUE' : 'FALSE';
                                break;
                            case 2:
                                // Error formula. Error code is in +2;
                                $value = ErrorCode::indexedCode($value);
                                break;
                            case 3:
                                // Formula result is a null string.
                                $value = '';
                                break;
                        }
                    } else {
                        // result is a number, so first 14 bytes are just like a _NUMBER record
                        $value = current(unpack('d', substr($this->handler, $position + 6, 8)));
                    }

                    $this->row[$columnIndex]->setValue($value)->setXfIndex($xfIndex);
                    break;
            }

            $currentRowOffset = $position;
            $currentRowLength = $length;
        } while ($code != Reader::RECORD_EOF);

        return parent::current();
    }

    /**
     * @return void
     */
    public function rewind(): void
    {
        parent::rewind();

        if ($this->startRow > 1) {
            do {
                $this->position++;
            } while ($this->startRow > $this->position + 1);
        }
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