<?php

namespace EasyExcel\Metadata;

use EasyExcel\Interfaces\BaseExcel;

abstract class Excel implements BaseExcel
{
    /**
     * @var \EasyExcel\Metadata\Style[]
     */
    protected $cellXfs = [];

    /**
     * @var array
     */
    protected $sharedStrings = [];

    /**
     * Add cellXf.
     *
     * @param \EasyExcel\Metadata\Style $style
     * @param int|null $index
     * @return $this
     */
    public function addCellXf(Style $style, ?int $index = null): BaseExcel
    {
        if (isset($index)) {
            $this->cellXfs[$index] = $style->setIndex($index);
        } else {
            $this->cellXfs[] = $style->setIndex($this->getCellXfsCount());
        }

        return $this;
    }

    /**
     * Get cellXf.
     *
     * @param int $index
     * @return \EasyExcel\Metadata\Style|null
     */
    public function getCellXf(int $index): ?Style
    {
        return $this->cellXfs[$index] ?? null;
    }

    /**
     * Get cellXf.
     *
     * @param string $hashCode
     * @return \EasyExcel\Metadata\Style|null
     */
    public function getCellXfByHashCode(string $hashCode): ?Style
    {
        foreach ($this->cellXfs as $cellXf) {
            if ($cellXf->getHashCode() == $hashCode) {
                return $cellXf;
            }
        }
        return null;
    }

    /**
     * Get cellXfs.
     *
     * @return \EasyExcel\Metadata\Style[]
     */
    public function getCellXfs(): array
    {
        return $this->cellXfs;
    }

    /**
     * Get cellXfs count.
     *
     * @return int
     */
    public function getCellXfsCount(): int
    {
        return count($this->cellXfs);
    }

    /**
     * Add shared string.
     *
     * @param string $string
     * @param int|null $index
     * @return $this
     */
    public function addSharedString(string $string, ?int $index = null): BaseExcel
    {
        if (is_null($index)) {
            $this->sharedStrings[] = $string;
        } else {
            $this->sharedStrings[$index] = $string;
        }

        return $this;
    }

    /**
     * Get shared string by index.
     *
     * @param int $index
     * @return string|null
     */
    public function getSharedString(int $index): ?string
    {
        return $this->sharedStrings[$index] ?? null;
    }

    /**
     * Get shared strings.
     *
     * @return array
     */
    public function getSharedStrings(): array
    {
        return $this->sharedStrings;
    }

    /**
     * Get shared strings count.
     *
     * @return int
     */
    public function getSharedStringsCount(): int
    {
        return count($this->sharedStrings);
    }
}