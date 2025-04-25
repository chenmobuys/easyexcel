<?php

namespace EasyExcel\Writers;

use EasyExcel\Exceptions\SheetNameExistsException;
use EasyExcel\Interfaces\Writer;
use EasyExcel\Interfaces\WriterExcel as WriterExcelInterface;
use EasyExcel\Interfaces\WriterSheet as WriterSheetInterface;
use EasyExcel\Metadata\Excel;

class WriterExcel extends Excel implements WriterExcelInterface
{
    /**
     * Writer.
     *
     * @var \EasyExcel\Interfaces\Writer
     */
    protected $writer;

    /**
     * Sheets.
     *
     * @var \EasyExcel\Interfaces\WriterSheet[]
     */
    protected $sheets = [];

    /**
     * Constructor.
     *
     * @param \EasyExcel\Interfaces\Writer $writer
     */
    public function __construct(Writer $writer)
    {
        $this->writer = $writer;
    }

    /**
     * Get writer.
     *
     * @return \EasyExcel\Interfaces\Writer
     */
    public function getWriter(): Writer
    {
        return $this->writer;
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
     * @return \EasyExcel\Interfaces\WriterSheet|null
     */
    public function getSheetByName(string $sheetName): ?WriterSheetInterface
    {
        return current(array_filter(
            $this->sheets,
            function (WriterSheetInterface $sheet) use ($sheetName) {
                return $sheet->getName() == $sheetName;
            }
        )) ?: null;
    }

    /**
     * Get sheet by index.
     *
     * @param int $sheetIndex
     * @return \EasyExcel\Interfaces\WriterSheet|null
     */
    public function getSheetByIndex(int $sheetIndex): ?WriterSheetInterface
    {
        return $this->sheets[$sheetIndex] ?? null;
    }

    /**
     * Get active sheet.
     *
     * @return \EasyExcel\Interfaces\WriterSheet|null
     */
    public function getActiveSheet(): ?WriterSheetInterface
    {
        return current($this->sheets) ?: null;
    }

    /**
     * Get all sheets.
     *
     * @return \EasyExcel\Interfaces\WriterSheet[]
     */
    public function getAllSheets(): array
    {
        return $this->sheets;
    }

    /**
     *  Add sheet.
     *
     * @param string $sheetName
     * @param int|null $sheetIndex
     * @return \EasyExcel\Interfaces\WriterSheet
     * @throws \EasyExcel\Exceptions\SheetNameExistsException
     */
    public function addSheet(string $sheetName, ?int $sheetIndex = null): WriterSheetInterface
    {
        if ($this->hasSheetName($sheetName)) {
            throw new SheetNameExistsException($sheetName);
        }

        $sheetIndex = is_null($sheetIndex) ? count($this->sheets) : $sheetIndex;
        $sheet = new WriterSheet($this, $sheetName, $sheetIndex);

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