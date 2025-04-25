<?php

namespace EasyExcel\Exceptions;

use Throwable;

class FileNotWriteableException extends EasyExcelException
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
        parent::__construct("Can not write file '$filename'.", $code, $previous);

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