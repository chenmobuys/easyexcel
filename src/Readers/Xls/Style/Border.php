<?php

namespace EasyExcel\Readers\Xls\Style;

use EasyExcel\Metadata\Style\Border as StyleBorder;

class Border
{
    private const INDEXED_BORDERS = [
        0x00 => StyleBorder::BORDER_NONE,
        0x01 => StyleBorder::BORDER_THIN,
        0x02 => StyleBorder::BORDER_MEDIUM,
        0x03 => StyleBorder::BORDER_DASHED,
        0x04 => StyleBorder::BORDER_DOTTED,
        0x05 => StyleBorder::BORDER_THICK,
        0x06 => StyleBorder::BORDER_DOUBLE,
        0x07 => StyleBorder::BORDER_HAIR,
        0x08 => StyleBorder::BORDER_MEDIUMDASHED,
        0x09 => StyleBorder::BORDER_DASHDOT,
        0x0A => StyleBorder::BORDER_MEDIUMDASHDOT,
        0x0B => StyleBorder::BORDER_DASHDOTDOT,
        0x0C => StyleBorder::BORDER_MEDIUMDASHDOTDOT,
        0x0D => StyleBorder::BORDER_SLANTDASHDOT,
    ];

    /**
     * @param int $index
     * @return string
     */
    public static function indexedBorder(int $index): string
    {
        if (isset(self::INDEXED_BORDERS[$index])) {
            return self::INDEXED_BORDERS[$index];
        }

        return StyleBorder::BORDER_NONE;
    }
}