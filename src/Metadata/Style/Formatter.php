<?php

namespace EasyExcel\Metadata\Style;

use DateInterval;
use DateTime;
use Exception;

class Formatter
{
    /**
     * @param float $value
     * @param string $formatCode
     * @return string
     */
    public static function getFormattedValue(float $value, string $formatCode): string
    {
        if (in_array($formatCode, [Format::FORMAT_GENERAL, Format::FORMAT_TEXT])) {
            return $value;
        }

        $format = self::getFormat($value, $formatCode);

        if (preg_match('/^(\[\$[\da-zA-Z]-[\dA-F]*])*[hmsdy]/i', $format)) {
            $value = self::getDatetimeFormatValue($value, $format);
        } else if (preg_match('/#?.*\?\/\?/', $format)) {
            $value = self::getFractionFormatValue($value, $format);
        } else if (preg_match('/(0+)(\\.?)(0*)/', $format, $matches)) {
            $value = self::getNumberFormatValue($value, $format, $matches);
        } else {
            $value = str_replace('?', '', $format);
        }

        return $value;
    }

    public static function isDateValue(float $value, string $formatCode): bool
    {
        $format = self::getFormat($value, $formatCode);

        return preg_match('/^(\[\$[\da-zA-Z]-[\dA-F]*])*[hmsdy]/i', $format);
    }

    /**
     * @param float $value
     * @param string $formatCode
     * @return string
     */
    private static function getFormat(float $value, string $formatCode): string
    {
        [, $format] = self::parseFormatCode($value, $formatCode);

        return (string)preg_replace('/_.?/ui', '', $format);
    }

    /**
     * @param float $value
     * @param string $formatCode
     * @return array
     */
    private static function parseFormatCode(float $value, string $formatCode): array
    {
        $formatCode = str_replace(['"', '*'], '', $formatCode);

        $formatCode = preg_replace('/\[\$-.*]/', '', $formatCode);

        $formatCode = preg_replace('/\\\\/', '', $formatCode);

        $sections = preg_split('/(;)(?=(?:[^"]|"[^"]*")*$)/u', $formatCode);

        $cnt = count($sections);
        $colorRegex = '/\\[(' . implode('|', self::$colors) . ')]/mui';
        $condRegex = '/\\[(>|>=|<|<=|=|<>)([+-]?\\d+([.]\\d+)?)]/';
        $colors = ['', '', '', '', ''];
        $condOps = ['', '', '', '', ''];
        $condValues = [0, 0, 0, 0, 0];

        for ($idx = 0; $idx < $cnt; ++$idx) {
            if (preg_match($colorRegex, $sections[$idx], $matches)) {
                $colors[$idx] = $matches[0];
                $sections[$idx] = (string)preg_replace($colorRegex, '', $sections[$idx]);
            }
            if (preg_match($condRegex, $sections[$idx], $matches)) {
                $condOps[$idx] = $matches[1];
                $condValues[$idx] = $matches[2];
                $sections[$idx] = (string)preg_replace($condRegex, '', $sections[$idx]);
            }
        }

        $color = $colors[0];
        $format = $sections[0];
        $absValue = $value;
        switch ($cnt) {
            case 2:
                $absValue = abs($value);
                if (!self::splitFormatCompare($value, $condOps[0], $condValues[0], '>=')) {
                    $color = $colors[1];
                    $format = $sections[1];
                }
                break;
            case 3:
            case 4:
                $absValue = abs($value);
                if (!self::splitFormatCompare($value, $condOps[0], $condValues[0], '>')) {
                    if (self::splitFormatCompare($value, $condOps[1], $condValues[1], '<')) {
                        $color = $colors[1];
                        $format = $sections[1];
                    } else {
                        $color = $colors[2];
                        $format = $sections[2];
                    }
                }
                break;
        }

        return [$color, $format, $absValue];
    }

    /**
     * @param float $value
     * @param string $cond
     * @param float $val
     * @param string $dfCond
     * @return bool
     */
    private static function splitFormatCompare(float $value, string $cond, float $val, string $dfCond): bool
    {
        if (!$cond) {
            $val = 0;
            $cond = $dfCond;
        }
        switch ($cond) {
            case '>':
                return $value > $val;

            case '<':
                return $value < $val;

            case '<=':
                return $value <= $val;

            case '<>':
                return $value != $val;

            case '=':
                return $value == $val;
        }

        return $value >= $val;
    }

    /**
     * @param float $value
     * @param $format
     * @return string
     */
    private static function getDatetimeFormatValue(float $value, $format): string
    {
        if (Format::getCalendar() === Format::CALENDAR_WINDOWS_1900) {
            $baseDate = ($value < 1) ? new DateTime('1970-01-01')
                : (($value < 60) ? new DateTime('1899-12-31') : new DateTime('1899-12-30'));
        } else {
            $baseDate = new DateTime('1904-01-01');
        }

        $dateFormat = strtolower($format);
        $dateFormat = strtr($dateFormat, self::$dateReplacements);
        if (strpos($dateFormat, 'A') === false) {
            $dateFormat = strtr($dateFormat, self::$dateReplacements24);
        } else {
            $dateFormat = strtr($dateFormat, self::$dateReplacements12);
        }

        try {
            $days = (int)$value;
            $seconds = (int)(abs($value - $days) * 86400);
            $duration = sprintf('P%sD%s', $days, $seconds ? 'T' . $seconds . 'S' : '');
            return $baseDate->add(new DateInterval($duration))->format($dateFormat);
        } catch (Exception $e) {
        }

        return $format;
    }

    /**
     * @param float $value
     * @param string $format
     * @return string
     */
    private static function getFractionFormatValue(float $value, string $format): string
    {
        $absValue = abs($value);
        $sign = $value < 0.0 ? '-' : '';
        $integerPart = floor($absValue);
        $decimalPart = '0';
        if (preg_match('/^\\d*[.](\\d*[1-9])0*$/', $value, $matches) === 1) {
            $decimalPart = $matches[1];
        }
        if ($decimalPart === '0') {
            return "$sign$integerPart";
        }

        $decimalLength = strlen($decimalPart);
        $decimalDivisor = 10 ** $decimalLength;

        $GCD = self::GCD($decimalPart, $decimalDivisor);

        $adjustedDecimalPart = $decimalPart / $GCD;
        $adjustedDecimalDivisor = $decimalDivisor / $GCD;

        if ((strpos($format, '0') !== false)) {
            return "$sign$integerPart $adjustedDecimalPart/$adjustedDecimalDivisor";
        } elseif ((strpos($format, '#') !== false)) {
            if ($integerPart == 0) {
                return "$sign$adjustedDecimalPart/$adjustedDecimalDivisor";
            }
            return "$sign$integerPart $adjustedDecimalPart/$adjustedDecimalDivisor";
        } elseif ((substr($format, 0, 3) == '? ?')) {
            if ($integerPart == 0) {
                $integerPart = '';
            }
            return "$sign$integerPart $adjustedDecimalPart/$adjustedDecimalDivisor";
        }

        $adjustedDecimalPart += $integerPart * $adjustedDecimalDivisor;

        return "$sign$adjustedDecimalPart/$adjustedDecimalDivisor";
    }

    /**
     * @param float $value
     * @param string $format
     * @param array $matches
     * @return string
     */
    private static function getNumberFormatValue(float $value, string $format, array $matches): string
    {
        $left = $matches[1];
        $dec = $matches[2];
        $right = $matches[3];
        $scale = strlen($right);
        $minWidth = strlen($left) + strlen($dec) + strlen($right);
        $useThousands = (bool)preg_match('/(#,#|0,0)/', $format);
        $format = preg_replace('/#/', '0', $format);
        if ($useThousands) {
            $format = preg_replace(['/0,0/', '/#,#/'], ['00', '##'], $format);
            $value = number_format($value, $scale);
            $value = preg_replace('/(0+)(\\.?)(0*)/', $value, $format);
        } else if (preg_match('/^0(\.)?0*%$/', $format)) {
            $value = sprintf('%.' . $scale . 'f', $value) . '%';
        } else if (preg_match('/[0#]E[+-]0/i', $format)) {
            $value = sprintf('%5.' . $scale . 'E', $value);
        } else {
            $value = sprintf('%0' . $minWidth . '.' . strlen($right) . 'f', round($value, strlen($right)));
        }

        if (preg_match('/\[\$(.*)]/u', $format, $matches)) {
            [$currencyCode] = explode('-', $matches[1]);
            $value = preg_replace('/\[\$([^]]*)]/u', $currencyCode, (string)$value);
        }

        return $value;
    }

    /**
     * @param int $a
     * @param int $b
     * @return int
     */
    public static function GCD(int $a, int $b): int
    {
        $a = abs($a);
        $b = abs($b);
        $c = 1;
        if ($a + $b == 0) {
            return 0;
        }
        while ($a > 0) {
            $c = $a;
            $a = $b % $a;
            $b = $c;
        }
        return $c;
    }


    /**
     * @var array
     */
    private static $dateReplacements = [
        '\\'    => '',
        'am/pm' => 'A',
        'yyyy'  => 'Y',
        'yy'    => 'y',
        'mmmmm' => 'M',
        'mmmm'  => 'F',
        'mmm'   => 'M',
        ':mm'   => ':i',
        'mm'    => 'm',
        'm'     => 'n',
        'dddd'  => 'l',
        'ddd'   => 'D',
        'dd'    => 'd',
        'd'     => 'j',
        'ss'    => 's',
        '.s'    => '',
        '12H'   => [],
    ];

    /**
     * @var array
     */
    private static $dateReplacements12 = [
        'hh' => 'h',
        'h'  => 'G',
    ];

    /**
     * @var array
     */
    private static $dateReplacements24 = [
        'hh' => 'H',
        'h'  => 'G',
    ];

    /**
     * @var array
     */
    private static $colors = [
        'Black',
        'White',
        'Red',
        'Green',
        'Blue',
        'Yellow',
        'Magenta',
        'Cyan',
    ];
}
