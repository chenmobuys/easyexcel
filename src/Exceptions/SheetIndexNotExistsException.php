<?php

namespace EasyExcel\Exceptions;

use Throwable;

class SheetIndexNotExistsException extends EasyExcelException
{
    /**
     * Sheet name.
     *
     * @var int
     */
    private $sheetIndex;

    /**
     * Constructor.
     *
     * @param int $sheetIndex
     * @param int $code
     * @param \Throwable|null $previous
     */
    public function __construct(int $sheetIndex, int $code = 0, Throwable $previous = null)
    {
        parent::__construct("Sheet index '$sheetIndex' not exists.", $code, $previous);

        $this->sheetIndex = $sheetIndex;
    }

    /**
     * Get sheet index.
     *
     * @return int
     */
    public function getSheetIndex(): int
    {
        return $this->sheetIndex;
    }
}
