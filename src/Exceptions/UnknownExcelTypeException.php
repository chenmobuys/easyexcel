<?php

namespace EasyExcel\Exceptions;

use Throwable;

class UnknownExcelTypeException extends EasyExcelException
{
    /**
     * Filename.
     *
     * @var string
     */
    private $filename;

    /**
     * Constructor.
     *
     * @param string $filename
     * @param int $code
     * @param \Throwable|null $previous
     */
    public function __construct(string $filename, int $code = 0, ?Throwable $previous = null)
    {
        parent::__construct("Can not detect excel type for file '$filename'.", $code, $previous);

        $this->filename = $filename;
    }

    /**
     * Get filename.
     *
     * @return string
     */
    public function getFilename(): string
    {
        return $this->filename;
    }
}