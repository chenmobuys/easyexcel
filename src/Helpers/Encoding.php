<?php

namespace EasyExcel\Helpers;

use Closure;
use EasyExcel\Exceptions\NotSupportConverterException;

class Encoding
{
    public const CONVERTER_MB = 'mb';
    public const CONVERTER_ICONV = 'iconv';
    public const CONVERTER_CLOSURE = 'closure';

    /**
     * Encoding converter.
     *
     * @var string
     */
    private static $converter = self::CONVERTER_MB;

    /**
     * Closure converter.
     *
     * @var \Closure|null
     */
    private static $closureConverter;

    /**
     * Get encoding converter.
     *
     * @return string
     */
    public static function getConverter(): string
    {
        return self::$converter;
    }

    /**
     * Set encoding converter.
     *
     * @param string $converter
     * @return void
     * @throws NotSupportConverterException
     */
    public static function setConverter(string $converter): void
    {
        switch ($converter) {
            case self::CONVERTER_MB:
            case self::CONVERTER_ICONV:
            case self::CONVERTER_CLOSURE:
                self::$converter = $converter;
                break;
            default:
                throw new NotSupportConverterException(__CLASS__, $converter);
        }
    }

    /**
     * Get closure converter.
     *
     * @return Closure|null
     */
    public static function getClosureConverter(): ?Closure
    {
        return self::$closureConverter;
    }

    /**
     * Set closure converter.
     *
     * @param \Closure|null $closureConverter
     * @return void
     */
    public static function setClosureConverter(?Closure $closureConverter): void
    {
        self::$closureConverter = $closureConverter;
    }

    /**
     * Covert string encoding.
     *
     * @param string $string
     * @param string $from_encoding
     * @param string $to_encoding
     * @return string
     */
    public static function convertEncoding(string $string, string $from_encoding, string $to_encoding = 'UTF-8'): string
    {
        switch (self::$converter) {
            case self::CONVERTER_MB:
                $string = mb_convert_encoding($string, $to_encoding, $from_encoding);
                break;
            case self::CONVERTER_ICONV:
                $string = iconv($from_encoding, $to_encoding, $string);
                break;
            case self::CONVERTER_CLOSURE:
                if (self::$closureConverter instanceof Closure) {
                    $string = call_user_func(self::$closureConverter, $string, $from_encoding, $to_encoding);
                }
                break;
        }

        return $string;
    }
}