<?php

namespace EasyExcel\Readers\Xls\Style;

use EasyExcel\Readers\Xls\Reader;

class Color
{
    private const BUILT_IN_INDEXED_COLORS = [
        0x00 => '000000',
        0x01 => 'FFFFFF',
        0x02 => 'FF0000',
        0x03 => '00FF00',
        0x04 => '0000FF',
        0x05 => 'FFFF00',
        0x06 => 'FF00FF',
        0x07 => '00FFFF',
        0x40 => '000000', // system window text color
        0x41 => 'FFFFFF', // system window background color
    ];

    private const BIFF8_INDEXED_COLORS = [
        0x08 => '000000',
        0x09 => 'FFFFFF',
        0x0A => 'FF0000',
        0x0B => '00FF00',
        0x0C => '0000FF',
        0x0D => 'FFFF00',
        0x0E => 'FF00FF',
        0x0F => '00FFFF',
        0x10 => '800000',
        0x11 => '008000',
        0x12 => '000080',
        0x13 => '808000',
        0x14 => '800080',
        0x15 => '008080',
        0x16 => 'C0C0C0',
        0x17 => '808080',
        0x18 => '9999FF',
        0x19 => '993366',
        0x1A => 'FFFFCC',
        0x1B => 'CCFFFF',
        0x1C => '660066',
        0x1D => 'FF8080',
        0x1E => '0066CC',
        0x1F => 'CCCCFF',
        0x20 => '000080',
        0x21 => 'FF00FF',
        0x22 => 'FFFF00',
        0x23 => '00FFFF',
        0x24 => '800080',
        0x25 => '800000',
        0x26 => '008080',
        0x27 => '0000FF',
        0x28 => '00CCFF',
        0x29 => 'CCFFFF',
        0x2A => 'CCFFCC',
        0x2B => 'FFFF99',
        0x2C => '99CCFF',
        0x2D => 'FF99CC',
        0x2E => 'CC99FF',
        0x2F => 'FFCC99',
        0x30 => '3366FF',
        0x31 => '33CCCC',
        0x32 => '99CC00',
        0x33 => 'FFCC00',
        0x34 => 'FF9900',
        0x35 => 'FF6600',
        0x36 => '666699',
        0x37 => '969696',
        0x38 => '003366',
        0x39 => '339966',
        0x3A => '003300',
        0x3B => '333300',
        0x3C => '993300',
        0x3D => '993366',
        0x3E => '333399',
        0x3F => '333333',
    ];

    private const BIFF5_INDEXED_COLORS = [
        0x08 => '000000',
        0x09 => 'FFFFFF',
        0x0A => 'FF0000',
        0x0B => '00FF00',
        0x0C => '0000FF',
        0x0D => 'FFFF00',
        0x0E => 'FF00FF',
        0x0F => '00FFFF',
        0x10 => '800000',
        0x11 => '008000',
        0x12 => '000080',
        0x13 => '808000',
        0x14 => '800080',
        0x15 => '008080',
        0x16 => 'C0C0C0',
        0x17 => '808080',
        0x18 => '8080FF',
        0x19 => '802060',
        0x1A => 'FFFFC0',
        0x1B => 'A0E0F0',
        0x1C => '600080',
        0x1D => 'FF8080',
        0x1E => '0080C0',
        0x1F => 'C0C0FF',
        0x20 => '000080',
        0x21 => 'FF00FF',
        0x22 => 'FFFF00',
        0x23 => '00FFFF',
        0x24 => '800080',
        0x25 => '800000',
        0x26 => '008080',
        0x27 => '0000FF',
        0x28 => '00CFFF',
        0x29 => '69FFFF',
        0x2A => 'E0FFE0',
        0x2B => 'FFFF80',
        0x2C => 'A6CAF0',
        0x2D => 'DD9CB3',
        0x2E => 'B38FEE',
        0x2F => 'E3E3E3',
        0x30 => '2A6FF9',
        0x31 => '3FB8CD',
        0x32 => '488436',
        0x33 => '958C41',
        0x34 => '8E5E42',
        0x35 => 'A0627A',
        0x36 => '624FAC',
        0x37 => '969696',
        0x38 => '1D2FBE',
        0x39 => '286676',
        0x3A => '004500',
        0x3B => '453E01',
        0x3C => '6A2813',
        0x3D => '85396A',
        0x3E => '4A3285',
        0x3F => '424242',
    ];

    /**
     * @param int $index
     * @param array $palette
     * @param int $version
     * @return string
     */
    public static function indexedColor(int $index, array $palette, int $version): string
    {
        if ($index <= 0x07 || $index >= 0x40) {
            $colorPalette = self::BUILT_IN_INDEXED_COLORS;
        } elseif (isset($palette[$index - 8])) {
            $index = $index - 8;
            $colorPalette = $palette;
        } elseif ($version == Reader::BIFF_VERSION_8) {
            $colorPalette = self::BIFF8_INDEXED_COLORS;
        } else {
            $colorPalette = self::BIFF5_INDEXED_COLORS;
        }

        return $colorPalette[$index] ?? '000000';
    }
}