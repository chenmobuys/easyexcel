<?php

namespace EasyExcel\Readers\Xls;

use EasyExcel\Helpers\Encoding;

class Helper
{
    /**
     * Get int.
     *
     * @param int $position
     * @param string $data
     * @param int $length
     * @return int
     */
    public static function getInt(int $position, string $data, int $length = 4): int
    {
        $value = ord($data[$position]);
        for ($i = 1; $i < $length; $i++) {
            $ord = ord($data[$position + $i]);
            $ord = $i == 3 ? ($ord >= 128 ? -abs(256 - $ord) : ($ord & 127)) : $ord;
            $value = $value | ($ord << (8 * $i));
        }

        return $value >= 4294967294 ? -2 : $value;
    }

    /**
     * Get IEEE754.
     *
     * @param float $rkNum
     * @return float
     */
    public static function getIEEE754(float $rkNum): float
    {
        if (($rkNum & 0x02) != 0) {
            $value = $rkNum >> 2;
        } else {
            // changes by mmp, info on IEEE754 encoding from
            // research.microsoft.com/~hollasch/cgindex/coding/ieeefloat.html
            // The RK format calls for using only the most significant 30 bits
            // of the 64 bit floating point value. The other 34 bits are assumed
            // to be 0 so we use the upper 30 bits of $rkNum as follows...
            $sign = ($rkNum & 0x80000000) >> 31;
            $exp = ($rkNum & 0x7ff00000) >> 20;
            $mantissa = (0x100000 | ($rkNum & 0x000ffffc));
            $value = $mantissa / 2 ** (20 - ($exp - 1023));
            if ($sign) {
                $value = -1 * $value;
            }
            //end of changes by mmp
        }
        if (($rkNum & 0x01) != 0) {
            $value /= 100;
        }

        return $value;
    }

    // /**
    //  * Extract number.
    //  *
    //  * @param string $data
    //  * @return float
    //  */
    // public static function extractNumber(string $data): float
    // {
    //     $rknumhigh = Helper::getInt(4, $data);
    //     $rknumlow = Helper::getInt(0, $data);
    //     $sign = ($rknumhigh & 0x80000000) >> 31;
    //     $exp = (($rknumhigh & 0x7ff00000) >> 20) - 1023;
    //     $mantissa = (0x100000 | ($rknumhigh & 0x000fffff));
    //     $mantissalow1 = ($rknumlow & 0x80000000) >> 31;
    //     $mantissalow2 = ($rknumlow & 0x7fffffff);
    //     $value = $mantissa / 2 ** (20 - $exp);
    //
    //     if ($mantissalow1 != 0) {
    //         $value += 1 / 2 ** (21 - $exp);
    //     }
    //
    //     $value += $mantissalow2 / 2 ** (52 - $exp);
    //     if ($sign) {
    //         $value *= -1;
    //     }
    //
    //     return $value;
    // }

    /**
     * Read rgb.
     *
     * @param string $rgb
     * @return string
     */
    public static function readRGB(string $rgb): string
    {
        return sprintf('%02X%02X%02X', ord($rgb[0]), ord($rgb[1]), ord($rgb[2]));
    }

    /**
     * Read short byte string.
     *
     * @param string $data
     * @param string $encoding
     * @return string
     */
    public static function readByteStringShort(string $data, string $encoding): string
    {
        // offset: 0; size: 1; length of the string (character count)
        $ln = ord($data[0]);

        // offset: 1: size: var; character array (8-bit characters)
        return Encoding::convertEncoding(substr($data, 1, $ln), $encoding);
    }

    /**
     * Read long byte string.
     *
     * @param string $data
     * @param string $encoding
     * @return string
     */
    public static function readByteStringLong(string $data, string $encoding): string
    {
        // offset: 0; size: 2; length of the string (character count)
        $characterCount = Helper::getInt(0, $data, 2);

        // offset: 2: size: var; character array (8-bit characters)
        return Encoding::convertEncoding(substr($data, 2), $encoding);
    }

    /**
     * Read short unicode string.
     *
     * @param string $data
     * @param string $encoding
     * @return string
     */
    public static function readUnicodeStringShort(string $data, string $encoding): string
    {
        // offset: 0: size: 1; length of the string (character count)
        $characterCount = ord($data[0]);

        return self::readUnicodeString(substr($data, 1), $characterCount, $encoding);
    }

    /**
     * Read long unicode string.
     *
     * @param string $data
     * @param string $encoding
     * @return string
     */
    public static function readUnicodeStringLong(string $data, string $encoding): string
    {
        // offset: 0: size: 2; length of the string (character count)
        $characterCount = Helper::getInt(0, $data, 2);

        return self::readUnicodeString(substr($data, 2), $characterCount, $encoding);
    }

    /**
     * Read unicode string.
     *
     * @param string $data
     * @param int $length
     * @param string $encoding
     * @return string
     */
    public static function readUnicodeString(string $data, int $length, string $encoding): string
    {
        $position = 0;
        $optionFlags = ord($data[0]);

        // offset: 0: size: 1; option flags
        // bit: 0; mask: 0x01; character compression (0 = compressed 8-bit, 1 = uncompressed 16-bit)
        $isUncompressed = (0x01 & $optionFlags) >> 0;
        $position += 1;

        // bit: 2; mask: 0x04; Asian phonetic settings
        $hasAsian = (0x04 & $optionFlags) >> 2;
        $position += $hasAsian ? 4 : 0;

        // bit: 3; mask: 0x08; Rich-Text settings
        $hasRichText = (0x08 & $optionFlags) >> 3;
        $position += $hasRichText ? 2 : 0;

        $data = substr($data, $position, $isUncompressed ? 2 * $length : $length);

        $data = $isUncompressed ? $data : Helper::uncompressByteString($data);

        return Encoding::convertEncoding($data, $encoding);
    }

    /**
     * Convert UTF-16 string in compressed notation to uncompressed form. Only used for BIFF8.
     *
     * @param string $string
     *
     * @return string
     */
    private static function uncompressByteString(string $string): string
    {
        $uncompressedString = '';
        $strLen = strlen($string);
        for ($i = 0; $i < $strLen; ++$i) {
            $uncompressedString .= $string[$i] . "\0";
        }

        return $uncompressedString;
    }
}