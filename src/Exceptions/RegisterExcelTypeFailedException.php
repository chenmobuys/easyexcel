<?php

namespace EasyExcel\Exceptions;

use Throwable;

class RegisterExcelTypeFailedException extends EasyExcelException
{
    /**
     * Class name.
     *
     * @var string
     */
    private $className;

    /**
     * Constructor.
     *
     * @param string $className
     * @param int $code
     * @param \Throwable|null $previous
     */
    public function __construct(string $className, int $code = 0, Throwable $previous = null)
    {
        parent::__construct("Register class must implements interface '$className'.", $code, $previous);

        $this->className = $className;
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
}