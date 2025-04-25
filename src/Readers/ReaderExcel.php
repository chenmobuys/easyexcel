<?php

namespace EasyExcel\Readers;

use EasyExcel\Exceptions\SheetNameExistsException;
use EasyExcel\Interfaces\Reader;
use EasyExcel\Interfaces\ReaderExcel as ReaderExcelInterface;
use EasyExcel\Interfaces\ReaderSheet as ReaderSheetInterface;
use EasyExcel\Metadata\Excel;

class ReaderExcel extends Excel implements ReaderExcelInterface
{
    /**
     * Reader.
     *
     * @var \EasyExcel\Interfaces\Reader
     */
    protected $reader;

    /**
     * Sheets.
     *
     * @var \EasyExcel\Interfaces\ReaderSheet[]
     */
    protected $sheets = [];

    /**
     * Constructor.
     *
     * @param \EasyExcel\Interfaces\Reader $reader
     */
    public function __construct(Reader $reader)
    {
        $this->reader = $reader;
    }

    /**
     * Get reader.
     *
     * @return \EasyExcel\Interfaces\Reader
     */
    public function getReader(): Reader
    {
        return $this->reader;
    }

    /**
     * Detect whether exists sheet name.
     *
     * @param string $sheetName
     * @return bool
     */
    public function hasSheetName(string $sheetName): bool
    {
        return !is_null($this->getSheetByName($sheetName));
    }

    /**
     * Detect whether exists sheet name.
     *
     * @param int $sheetIndex
     * @return bool
     */
    public function hasSheetIndex(int $sheetIndex): bool
    {
        return !is_null($this->getSheetByIndex($sheetIndex));
    }

    /**
     * Get sheet by name.
     *
     * @param string $sheetName
     * @return \EasyExcel\Interfaces\ReaderSheet|null
     */
    public function getSheetByName(string $sheetName): ?ReaderSheetInterface
    {
        return current(array_filter(
            $this->sheets,
            function (ReaderSheetInterface $sheet) use ($sheetName) {
                return $sheet->getName() == $sheetName;
            }
        )) ?: null;
    }

    /**
     * Get sheet by index.
     *
     * @param int $sheetIndex
     * @return \EasyExcel\Interfaces\ReaderSheet|null
     */
    public function getSheetByIndex(int $sheetIndex): ?ReaderSheetInterface
    {
        return $this->sheets[$sheetIndex] ?? null;
    }

    /**
     * Get active sheet.
     *
     * @return \EasyExcel\Interfaces\ReaderSheet|null
     */
    public function getActiveSheet(): ?ReaderSheetInterface
    {
        return current($this->sheets) ?: null;
    }

    /**
     * Get all sheets.
     *
     * @return \EasyExcel\Interfaces\ReaderSheet[]
     */
    public function getAllSheets(): array
    {
        return $this->sheets;
    }

    /**
     * Add sheet.
     *
     * @param string|null $sheetName
     * @param int|null $sheetIndex
     * @return \EasyExcel\Interfaces\ReaderSheet
     * @throws \EasyExcel\Exceptions\SheetNameExistsException
     */
    public function addSheet(string $sheetName, ?int $sheetIndex = null): ReaderSheetInterface
    {
        if ($this->hasSheetName($sheetName)) {
            throw new SheetNameExistsException($sheetName);
        }

        $sheetIndex = is_null($sheetIndex) ? count($this->sheets) : $sheetIndex;
        $sheet = new ReaderSheet($this, $sheetName, $sheetIndex);

        if ($this->hasSheetIndex($sheetIndex)) {
            array_splice($this->sheets, $sheetIndex, 1, [$sheet]);
        } else {
            $this->sheets[] = $sheet;
        }

        return $sheet;
    }

    /**
     * Remove sheet by name.
     *
     * @param string $sheetName
     * @return void
     */
    public function removeSheetByName(string $sheetName): void
    {
        foreach ($this->sheets as $index => $sheet) {
            if ($sheet->getName() == $sheetName) {
                $this->removeSheetByIndex($index);
                break;
            }
        }
    }

    /**
     * Remove sheet by index.
     *
     * @param int $sheetIndex
     * @return void
     */
    public function removeSheetByIndex(int $sheetIndex): void
    {
        array_splice($this->sheets, $sheetIndex, 1);
    }

    /**
     * Close excel.
     *
     * @return void
     */
    public function close(): void
    {
        foreach ($this->sheets as $sheet) {
            $sheet->close();
        }
    }
}