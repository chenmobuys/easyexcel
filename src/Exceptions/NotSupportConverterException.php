<?php

namespace EasyExcel\Exceptions;

use Throwable;

class NotSupportConverterException extends EasyExcelException
{
    /**
     * Converter.
     *
     * @var string
     */
    private $converter;

    /**
     * Constructor.
     *
     * @param  string           $converter
     * @param  int              $code
     * @param  \Throwable|null  $previous
     */
    public function __construct(string $converter, int $code = 0, Throwable $previous = null)
    {
        parent::__construct("Not supported the converter '$converter'.", $code, $previous);

        $this->converter = $converter;
    }

    /**
     * Get converter.
     *
     * @return string
     */
    public function getConverter(): string
    {
        return $this->converter;
    }
}