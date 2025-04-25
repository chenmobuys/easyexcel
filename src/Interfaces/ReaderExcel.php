<?php

namespace EasyExcel\Interfaces;

interface ReaderExcel extends BaseExcel
{
    /**
     * Get reader.
     *
     * @return \EasyExcel\Interfaces\Reader
     */
    public function getReader(): Reader;

    /**
     * Detect whether exists sheet name.
     *
     * @param string $sheetName
     * @return bool
     */
    public function hasSheetName(string $sheetName): bool;

    /**
     * Detect whether exists sheet name.
     *
     * @param int $sheetIndex
     * @return bool
     */
    public function hasSheetIndex(int $sheetIndex): bool;

    /**
     * Get sheet by name.
     *
     * @param string $sheetName
     * @return \EasyExcel\Interfaces\ReaderSheet|null
     */
    public function getSheetByName(string $sheetName): ?ReaderSheet;

    /**
     * Get sheet by index.
     *
     * @param int $sheetIndex
     * @return \EasyExcel\Interfaces\ReaderSheet|null
     */
    public function getSheetByIndex(int $sheetIndex): ?ReaderSheet;

    /**
     * Get active sheet.
     *
     * @return \EasyExcel\Interfaces\ReaderSheet|null
     */
    public function getActiveSheet(): ?ReaderSheet;

    /**
     * Get all sheets.
     *
     * @return \EasyExcel\Interfaces\ReaderSheet[]
     */
    public function getAllSheets(): array;

    /**
     * Add sheet.
     *
     * @param string $sheetName
     * @param int|null $sheetIndex
     * @return \EasyExcel\Interfaces\ReaderSheet
     */
    public function addSheet(string $sheetName, ?int $sheetIndex = null): ReaderSheet;

    /**
     * Remove sheet by name.
     *
     * @param string $sheetName
     * @return void
     */
    public function removeSheetByName(string $sheetName): void;

    /**
     * Remove sheet by index.
     *
     * @param int $sheetIndex
     * @return void
     */
    public function removeSheetByIndex(int $sheetIndex): void;
}