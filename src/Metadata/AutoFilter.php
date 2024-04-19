<?php

namespace EasyExcel\Metadata;

class AutoFilter
{
    /**
     * @var string
     */
    private $range;

    /**
     * Constructor.
     *
     * @param string $range
     */
    public function __construct(string $range = '')
    {
        $this->range = $range;
    }

    /**
     * @return string
     */
    public function getRange(): string
    {
        return $this->range;
    }

    /**
     * @param string $range
     */
    public function setRange(string $range): void
    {
        $this->range = $range;
    }
}