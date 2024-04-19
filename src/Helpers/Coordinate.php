<?php

namespace EasyExcel\Helpers;

class Coordinate
{
    /**
     * Convert column letter to column index.
     *
     * @param string $letter
     * @return ?int
     */
    public static function columnIndexFromColumnLetter(string $letter): ?int
    {
        $result = 0;
        $letter = strtoupper($letter);
        for ($i = strlen($letter) - 1, $j = 0; $i >= 0; $i--, $j++) {
            $ord = ord($letter[$i]) - 64;
            if ($ord > 26 || $ord < 1) {
                return null;
            }
            $result += $ord * pow(26, $j);
        }
        return $result - 1;
    }

    /**
     * Convert column index to column letter.
     *
     * @param int $index
     * @return string
     */
    public static function columnLetterFromColumnIndex(int $index): string
    {
        $index++;
        $letter = null;
        do {
            $characterValue = ($index % 26) ?: 26;
            $index = ($index - $characterValue) / 26;
            $letter = chr($characterValue + 64) . ($letter ?: '');
        } while ($index > 0);
        return $letter;
    }

    /**
     * Convert coordinate string to row number and column letter.
     *
     * @param string $coordinate
     * @return array [rowNumber, columnLetter]
     */
    public static function rowNumberAndColumnLetterFromCoordinate(string $coordinate): array
    {
        preg_match('/^([$]?[A-Z]{1,3})([$]?\\d{1,7})$/', $coordinate, $matches);

        return [(int) $matches[2], $matches[1]];
    }

    /**
     * Convert row number and column letter to coordinate.
     *
     * @param int $rowNumber
     * @param string $columnLetter
     * @return string
     */
    public static function coordinateFromRowNumberAndColumnLetter(int $rowNumber, string $columnLetter): string
    {
        return $columnLetter . $rowNumber;
    }

    /**
     * Convert coordinate string to row index and column index.
     *
     * @param string $coordinate
     * @return array [rowIndex, columnIndex]
     */
    public static function rowIndexAndColumnIndexFromCoordinate(string $coordinate): array
    {
        [$rowNumber, $columnLetter] = Coordinate::rowNumberAndColumnLetterFromCoordinate($coordinate);

        $rowIndex = $rowNumber - 1;
        $columnIndex = Coordinate::columnIndexFromColumnLetter($columnLetter);

        return [$rowIndex, $columnIndex];
    }

    /**
     * Convert row index and column index to coordinate.
     *
     * @param int $rowIndex
     * @param int $columnIndex
     * @return string
     */
    public static function coordinateFromRowIndexAndColumnIndex(int $rowIndex, int $columnIndex): string
    {
        return Coordinate::columnLetterFromColumnIndex($columnIndex) . ($rowIndex + 1);
    }

    /**
     * Convert coordinate range to coordinate array.
     *
     * @param string $range
     * @return array
     */
    public static function coordinatesFromRange(string $range): array
    {
        if (strpos($range, ':') === false) {
            return [$range];
        }

        $coordinates = [];
        [$coordinateStart, $coordinateEnd] = explode(':', $range);
        [$rowStartIndex, $columnStartIndex] = Coordinate::rowIndexAndColumnIndexFromCoordinate($coordinateStart);
        [$rowEndIndex, $columnEndIndex] = Coordinate::rowIndexAndColumnIndexFromCoordinate($coordinateEnd);

        for ($i = $rowStartIndex; $i <= $rowEndIndex; $i++) {
            for ($j = $columnStartIndex; $j <= $columnEndIndex; $j++) {
                $coordinates[] = Coordinate::coordinateFromRowIndexAndColumnIndex($i, $j);
            }
        }

        return $coordinates;
    }

    /**
     * Detect coordinate is in range.
     *
     * @param string $coordinate
     * @param string $range
     * @return bool
     */
    public static function coordinateIsInRange(string $coordinate, string $range): bool
    {
        [$rowIndex, $columnIndex] = Coordinate::rowIndexAndColumnIndexFromCoordinate($coordinate);
        [[$rowStartIndex, $columnStartIndex], [$rowEndIndex, $columnEndIndex]] = array_map(function ($item) {
            return Coordinate::rowIndexAndColumnIndexFromCoordinate($item);
        }, explode(':', $range));

        return ($columnStartIndex <= $columnIndex) && ($columnEndIndex >= $columnIndex) && ($rowStartIndex <= $rowIndex) && ($rowEndIndex >= $rowIndex);
    }
}
