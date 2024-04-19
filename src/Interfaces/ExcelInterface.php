<?php

namespace EasyExcel\Interfaces;

interface ExcelInterface
{
    public function hasSheetName(string $sheetName): bool;

    public function hasSheetIndex(int $sheetIndex): bool;

    public function getActiveSheet(): ?SheetInterface;

    public function setActiveSheetByName(string $sheetName, bool $createIfNotExists = false): ExcelInterface;

    public function setActiveSheetByIndex(int $sheetIndex): ExcelInterface;

    public function getAllSheets(): array;

    public function getSheetByName(string $sheetName, bool $createIfNotExists = false): ?SheetInterface;

    public function getSheetByIndex(int $sheetIndex): ?SheetInterface;

    public function removeSheetByName(string $sheetName): ExcelInterface;

    public function removeSheetByIndex(int $sheetIndex): ExcelInterface;
}