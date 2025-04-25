<?php

namespace EasyExcel\Exceptions;

use Throwable;

class NotSupportConverterException extends EasyExcelException
{
    /**
     * Class name.
     *
     * @var string
     */
    private $className;

    /**
     * Converter.
     *
     * @var string
     */
    private $converter;

    /**
     * Constructor.
     *
     * @param string $className
     * @param string $converter
     * @param int $code
     * @param \Throwable|null $previous
     */
    public function __construct(string $className, string $converter, int $code = 0, ?Throwable $previous = null)
    {
        parent::__construct("$className not supported the converter '$converter'.", $code, $previous);

        $this->className = $className;

        $this->converter = $converter;
    }

    /**
     * Get class name.
     *
     * @return string
     */
    public function getClassName(): string
    {
        return $this->className;
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