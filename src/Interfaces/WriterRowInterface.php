<?php

namespace EasyExcel\Interfaces;

use EasyExcel\Metadata\Style;

interface WriterRowInterface
{
    /**
     * Write rows.
     *
     * @param array $rows
     * @param \EasyExcel\Metadata\Style|null $style
     * @return void
     */
    public function writes(array $rows, ?Style $style = null): void;

    /**
     * Close row writer.
     *
     * @return void
     */
    public function close(): void;
}