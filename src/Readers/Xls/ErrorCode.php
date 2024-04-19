<?php

namespace EasyExcel\Readers\Xls;

class ErrorCode
{
    private const INDEXED_CODES = [
        0x00 => '#NULL!',
        0x07 => '#DIV/0!',
        0x0F => '#VALUE!',
        0x17 => '#REF!',
        0x1D => '#NAME?',
        0x24 => '#NUM!',
        0x2A => '#N/A',
    ];

    /**
     * Map error code, e.g. '#N/A'.
     *
     * @param int $code
     *
     * @return string|null
     */
    public static function indexedCode(int $code): ?string
    {
        return self::INDEXED_CODES[$code] ?? null;
    }
}