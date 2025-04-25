<?php

namespace EasyExcel\Interfaces;

interface WriterExcel extends BaseExcel
{
    /**
     * Get writer.
     *
     * @return \EasyExcel\Interfaces\Writer
     */
    public function getWriter(): Writer;

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
     * @return \EasyExcel\Interfaces\WriterSheet|null
     */
    public function getSheetByName(string $sheetName): ?WriterSheet;

    /**
     * Get sheet by index.
     *
     * @param int $sheetIndex
     * @return \EasyExcel\Interfaces\WriterSheet|null
     */
    public function getSheetByIndex(int $sheetIndex): ?WriterSheet;

    /**
     * Get active sheet.
     *
     * @return \EasyExcel\Interfaces\WriterSheet|null
     */
    public function getActiveSheet(): ?WriterSheet;

    /**
     * Get all sheets.
     *
     * @return \EasyExcel\Interfaces\WriterSheet[]
     */
    public function getAllSheets(): array;

    /**
     * Add sheet.
     *
     * @param string $sheetName
     * @param int|null $sheetIndex
     * @return \EasyExcel\Interfaces\WriterSheet
     */
    public function addSheet(string $sheetName, ?int $sheetIndex = null): WriterSheet;

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