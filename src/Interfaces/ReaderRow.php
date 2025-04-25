<?php

namespace EasyExcel\Interfaces;

use Countable;
use EasyExcel\Metadata\Row;
use Iterator;

interface ReaderRow extends Iterator, Countable
{
    /**
     * Get current row.
     *
     * @return \EasyExcel\Metadata\Row
     */
    public function current(): Row;

    /**
     * Close row iterator.
     *
     * @return void
     */
    public function close(): void;
}
