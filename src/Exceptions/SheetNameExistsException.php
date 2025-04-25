<?php

namespace EasyExcel\Exceptions;

use Throwable;

class SheetNameExistsException extends EasyExcelException
{
    /**
     * Sheet name.
     *
     * @var string
     */
    private $sheetName;

    /**
     * Constructor.
     *
     * @param string $sheetName
     * @param int $code
     * @param \Throwable|null $previous
     */
    public function __construct(string $sheetName, int $code = 0, Throwable $previous = null)
    {
        parent::__construct("Sheet name '$sheetName' already exists, rename this sheet name.", $code, $previous);

        $this->sheetName = $sheetName;
    }

    /**
     * Get sheet name.
     *
     * @return string
     */
    public function getSheetName(): string
    {
        return $this->sheetName;
    }
}