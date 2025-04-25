<?php

namespace EasyExcel\Exceptions;

use Throwable;

class NotSupportedExcelTypeException extends EasyExcelException
{
    /**
     * Class name.
     *
     * @var string
     */
    private $className;

    /**
     * Excel type.
     *
     * @var string
     */
    private $excelType;

    /**
     * Constructor.
     *
     * @param string $className
     * @param string $excelType
     * @param int $code
     * @param \Throwable|null $previous
     */
    public function __construct(string $className, string $excelType, int $code = 0, ?Throwable $previous = null)
    {
        parent::__construct("$className not supported the excel type '$excelType'.", $code, $previous);

        $this->className = $className;

        $this->excelType = $excelType;
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
     * Get excel type.
     *
     * @return string
     */
    public function getExcelType(): string
    {
        return $this->excelType;
    }
}