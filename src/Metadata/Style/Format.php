<?php

namespace EasyExcel\Metadata\Style;

class Format
{
    // Pre-defined formats
    public const FORMAT_GENERAL = 'General';
    public const FORMAT_TEXT = '@';

    public const FORMAT_NUMBER = '0';
    public const FORMAT_NUMBER_0 = '0.0';
    public const FORMAT_NUMBER_00 = '0.00';
    public const FORMAT_NUMBER_COMMA_SEPARATED1 = '#,##0.00';
    public const FORMAT_NUMBER_COMMA_SEPARATED2 = '#,##0.00_-';

    public const FORMAT_PERCENTAGE = '0%';
    public const FORMAT_PERCENTAGE_0 = '0.0%';
    public const FORMAT_PERCENTAGE_00 = '0.00%';

    public const FORMAT_DATE_YYYYMMDD2 = 'yyyy-mm-dd';
    public const FORMAT_DATE_YYYYMMDD = 'yyyy-mm-dd';
    public const FORMAT_DATE_DDMMYYYY = 'dd/mm/yyyy';
    public const FORMAT_DATE_DMYSLASH = 'd/m/yy';
    public const FORMAT_DATE_DMYMINUS = 'd-m-yy';
    public const FORMAT_DATE_DMMINUS = 'd-m';
    public const FORMAT_DATE_MYMINUS = 'm-yy';
    public const FORMAT_DATE_XLSX14 = 'mm-dd-yy';
    public const FORMAT_DATE_XLSX15 = 'd-mmm-yy';
    public const FORMAT_DATE_XLSX16 = 'd-mmm';
    public const FORMAT_DATE_XLSX17 = 'mmm-yy';
    public const FORMAT_DATE_XLSX22 = 'm/d/yy h:mm';
    public const FORMAT_DATE_DATETIME = 'd/m/yy h:mm';
    public const FORMAT_DATE_TIME1 = 'h:mm AM/PM';
    public const FORMAT_DATE_TIME2 = 'h:mm:ss AM/PM';
    public const FORMAT_DATE_TIME3 = 'h:mm';
    public const FORMAT_DATE_TIME4 = 'h:mm:ss';
    public const FORMAT_DATE_TIME5 = 'mm:ss';
    public const FORMAT_DATE_TIME6 = 'h:mm:ss';
    public const FORMAT_DATE_TIME7 = 'i:s.S';
    public const FORMAT_DATE_TIME8 = 'h:mm:ss;@';
    public const FORMAT_DATE_YYYYMMDDSLASH = 'yyyy/mm/dd;@';

    public const FORMAT_CURRENCY_USD_SIMPLE = '"$"#,##0.00_-';
    public const FORMAT_CURRENCY_USD = '$#,##0_-';
    public const FORMAT_CURRENCY_EUR_SIMPLE = '#,##0.00_-"€"';
    public const FORMAT_CURRENCY_EUR = '#,##0_-"€"';
    public const FORMAT_ACCOUNTING_USD = '_("$"* #,##0.00_);_("$"* \(#,##0.00\);_("$"* "-"??_);_(@_)';
    public const FORMAT_ACCOUNTING_EUR = '_("€"* #,##0.00_);_("€"* \(#,##0.00\);_("€"* "-"??_);_(@_)';

    public const CALENDAR_WINDOWS_1900 = 1900; //    Base date of 1st Jan 1900 = 1.0
    public const CALENDAR_MAC_1904 = 1904;     //    Base date of 2nd Jan 1904 = 1.0

    /**
     * @var array
     */
    private static $builtinFormats = [
        0 => '',
        1 => '0',
        2 => '0.00',
        3 => '#,##0',
        4 => '#,##0.00',

        9 => '0%',
        10 => '0.00%',
        11 => '0.00E+00',
        12 => '# ?/?',
        13 => '# ??/??',
        14 => 'mm-dd-yy',
        15 => 'd-mmm-yy',
        16 => 'd-mmm',
        17 => 'mmm-yy',
        18 => 'h:mm AM/PM',
        19 => 'h:mm:ss AM/PM',
        20 => 'h:mm',
        21 => 'h:mm:ss',
        22 => 'm/d/yy h:mm',

        37 => '#,##0 ;(#,##0)',
        38 => '#,##0 ;[Red](#,##0)',
        39 => '#,##0.00;(#,##0.00)',
        40 => '#,##0.00;[Red](#,##0.00)',

        44 => '_("$"* #,##0.00_);_("$"* \(#,##0.00\);_("$"* "-"??_);_(@_)',
        45 => 'mm:ss',
        46 => '[h]:mm:ss',
        47 => 'mmss.0',
        48 => '##0.0E+0',
        49 => '@',

        // CHT & CHS
        27 => '[$-404]e/m/d',
        30 => 'm/d/yy',
        36 => '[$-404]e/m/d',
        50 => '[$-404]e/m/d',
        57 => '[$-404]e/m/d',

        // THA
        59 => 't0',
        60 => 't0.00',
        61 => 't#,##0',
        62 => 't#,##0.00',
        67 => 't0%',
        68 => 't0.00%',
        69 => 't# ?/?',
        70 => 't# ??/??',

        // JPN
        28 => '[$-411]ggge"年"m"月"d"日"',
        29 => '[$-411]ggge"年"m"月"d"日"',
        31 => 'yyyy"年"m"月"d"日"',
        32 => 'h"時"mm"分"',
        33 => 'h"時"mm"分"ss"秒"',
        34 => 'yyyy"年"m"月"',
        35 => 'm"月"d"日"',
        51 => '[$-411]ggge"年"m"月"d"日"',
        52 => 'yyyy"年"m"月"',
        53 => 'm"月"d"日"',
        54 => '[$-411]ggge"年"m"月"d"日"',
        55 => 'yyyy"年"m"月"',
        56 => 'm"月"d"日"',
        58 => '[$-411]ggge"年"m"月"d"日"',
    ];

    /**
     * Calendar.
     *
     * @var int
     */
    private static $calendar = self::CALENDAR_WINDOWS_1900;

    /**
     * Format code.
     *
     * @var string
     */
    protected $formatCode = self::FORMAT_GENERAL;

    /**
     * Get Calendar.
     *
     * @return int
     */
    public static function getCalendar(): int
    {
        return self::$calendar;
    }

    /**
     * Set Calendar.
     *
     * @var int $calendar
     */
    public static function setCalendar(int $calendar): void
    {
        self::$calendar = $calendar;
    }

    /**
     * @param int $index
     * @return string
     */
    public static function builtInFormatCode(int $index): string
    {
        return self::$builtinFormats[$index] ?? '';
    }

    /**
     * @return string
     */
    public function getFormatCode(): string
    {
        return $this->formatCode;
    }

    /**
     * @param string $formatCode
     * @return $this
     */
    public function setFormatCode(string $formatCode): self
    {
        $this->formatCode = $formatCode;

        return $this;
    }

    /**
     * Get hash code.
     *
     * @return string
     */
    public function getHashCode(): string
    {
        return md5(
            __CLASS__ .
            $this->getFormatCode()
        );
    }
}