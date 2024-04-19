<?php

namespace EasyExcel\Metadata;

use EasyExcel\Exceptions\SheetIndexNotExistsException;
use EasyExcel\Exceptions\SheetNameNotExistsException;
use EasyExcel\Interfaces\ExcelInterface;
use EasyExcel\Interfaces\SheetInterface;

class Excel implements ExcelInterface
{
    /**
     * Sheets.
     *
     * @var \EasyExcel\Interfaces\SheetInterface[]
     */
    protected $sheets = [];

    /**
     * @var \EasyExcel\Metadata\Style[]
     */
    protected $cellXfs = [];

    /**
     * @var array
     */
    protected $sharedStrings = [];

    /**
     * @var int
     */
    protected $activeSheetIndex = 0;

    /**
     * Protected Constructor.
     */
    protected function __construct()
    {
    }

    /**
     * Add cellXf.
     *
     * @param  \EasyExcel\Metadata\Style  $style
     * @param  int|null  $index
     * @return $this
     */
    public function addCellXf(Style $style, int $index = null): ExcelInterface
    {
        if (isset($index)) {
            $this->cellXfs[$index] = $style->setIndex($index);
        } else {
            $this->cellXfs[] = $style->setIndex($this->getCellXfsCount());
        }

        return $this;
    }

    /**
     * Get cellXf.
     *
     * @param  int  $index
     * @return \EasyExcel\Metadata\Style|null
     */
    public function getCellXf(int $index): ?Style
    {
        return $this->cellXfs[$index] ?? null;
    }

    /**
     * Get cellXf.
     *
     * @param  string  $hashCode
     * @return \EasyExcel\Metadata\Style|null
     */
    public function getCellXfByHashCode(string $hashCode): ?Style
    {
        foreach ($this->cellXfs as $cellXf) {
            if ($cellXf->getHashCode() == $hashCode) {
                return $cellXf;
            }
        }
        return null;
    }

    /**
     * Get cellXfs.
     *
     * @return \EasyExcel\Metadata\Style[]
     */
    public function getCellXfs(): array
    {
        return $this->cellXfs;
    }

    /**
     * Get cellXfs count.
     *
     * @return int
     */
    public function getCellXfsCount(): int
    {
        return count($this->cellXfs);
    }

    /**
     * Add shared string.
     *
     * @param  string  $string
     * @param  int|null  $index
     * @return $this
     */
    public function addSharedString(string $string, ?int $index = null): ExcelInterface
    {
        if (is_null($index)) {
            $this->sharedStrings[] = $string;
        } else {
            $this->sharedStrings[$index] = $string;
        }

        return $this;
    }

    /**
     * Get shared string by index.
     *
     * @param  int  $index
     * @return string|null
     */
    public function getSharedString(int $index): ?string
    {
        return $this->sharedStrings[$index] ?? null;
    }

    /**
     * Get shared strings.
     *
     * @return array
     */
    public function getSharedStrings(): array
    {
        return $this->sharedStrings;
    }

    /**
     * Get shared strings count.
     *
     * @return int
     */
    public function getSharedStringsCount(): int
    {
        return count($this->sharedStrings);
    }

    public function hasSheetName(string $sheetName): bool
    {
        return !is_null($this->getSheetByName($sheetName));
    }

    public function hasSheetIndex(int $sheetIndex): bool
    {
        return !is_null($this->getSheetByIndex($sheetIndex));
    }

    public function getActiveSheet(): ?SheetInterface
    {
        return $this->sheets[$this->activeSheetIndex] ?? null;
    }

    public function setActiveSheetByName(string $sheetName, bool $createIfNotExists = false): ExcelInterface
    {
        $sheet = $this->getSheetByName($sheetName, $createIfNotExists);
        if (is_null($sheet)) {
            throw new SheetNameNotExistsException($sheetName);
        }
        $this->activeSheetIndex = $sheet->getIndex();
        return $this;
    }

    public function setActiveSheetByIndex(int $sheetIndex): ExcelInterface
    {
        $sheet = $this->getSheetByIndex($sheetIndex);
        if (is_null($sheet)) {
            throw new SheetIndexNotExistsException($sheetIndex);
        }
        $this->activeSheetIndex = $sheet->getIndex();
        return $this;
    }

    public function getAllSheets(): array
    {
        return $this->sheets;
    }

    public function getSheetByName(string $sheetName, bool $createIfNotExists = false): ?SheetInterface
    {
        $sheet = current(array_filter(
            $this->sheets,
            function (SheetInterface $sheet) use ($sheetName) {
                return $sheet->getName() == $sheetName;
            }
        )) ?: null;

        if ($createIfNotExists) {
            $sheetIndex = count($this->sheets);
            $this->sheets[] = new Sheet($sheetName, $sheetIndex, $this);
            return $this->sheets[count($this->sheets) - 1];
        }

        return $sheet;
    }

    public function getSheetByIndex(int $sheetIndex): ?SheetInterface
    {
        return $this->sheets[$sheetIndex] ?? null;
    }

    public function removeSheetByName(string $sheetName): ExcelInterface
    {
        foreach ($this->sheets as $index => $sheet) {
            if ($sheet->getName() == $sheetName) {
                $this->removeSheetByIndex($index);
                break;
            }
        }
        return $this;
    }

    public function removeSheetByIndex(int $sheetIndex): ExcelInterface
    {
        array_splice($this->sheets, $sheetIndex, 1);
        return $this;
    }
}
