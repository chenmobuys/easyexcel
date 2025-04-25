<?php

namespace EasyExcel\Metadata;

use EasyExcel\Metadata\Style\Alignment;
use EasyExcel\Metadata\Style\Borders;
use EasyExcel\Metadata\Style\Fill;
use EasyExcel\Metadata\Style\Font;
use EasyExcel\Metadata\Style\Format;
use EasyExcel\Metadata\Style\Protection;

class Style
{
    /**
     * @var int|null
     */
    protected $index;

    /**
     * @var \EasyExcel\Metadata\Style\Font
     */
    protected $font;

    /**
     * @var \EasyExcel\Metadata\Style\Fill
     */
    protected $fill;

    /**
     * @var \EasyExcel\Metadata\Style\Format
     */
    protected $format;

    /**
     * @var \EasyExcel\Metadata\Style\Borders
     */
    protected $borders;

    /**
     * @var \EasyExcel\Metadata\Style\Alignment
     */
    protected $alignment;

    /**
     * @var \EasyExcel\Metadata\Style\Protection
     */
    protected $protection;

    /**
     * @var bool
     */
    protected $quotePrefix = false;

    /**
     * Constructor.
     */
    public function __construct()
    {
        $this->font = new Font();
        $this->fill = new Fill();
        $this->format = new Format();
        $this->borders = new Borders();
        $this->alignment = new Alignment();
        $this->protection = new Protection();
    }

    public static function builder(): StyleBuilder
    {
        return new StyleBuilder(new static());
    }

    /**
     * @return int
     */
    public function getIndex(): ?int
    {
        return $this->index;
    }

    /**
     * @param string $index
     * @return $this
     */
    public function setIndex(string $index): self
    {
        $this->index = $index;

        return $this;
    }

    /**
     * @return \EasyExcel\Metadata\Style\Font
     */
    public function getFont(): Font
    {
        return $this->font;
    }

    /**
     * @param \EasyExcel\Metadata\Style\Font $font
     * @return $this
     */
    public function setFont(Font $font): self
    {
        $this->font = $font;

        return $this;
    }

    /**
     * @return \EasyExcel\Metadata\Style\Fill
     */
    public function getFill(): Fill
    {
        return $this->fill;
    }

    /**
     * @param \EasyExcel\Metadata\Style\Fill $fill
     * @return $this
     */
    public function setFill(Fill $fill): self
    {
        $this->fill = $fill;

        return $this;
    }

    /**
     * @return \EasyExcel\Metadata\Style\Format
     */
    public function getFormat(): Format
    {
        return $this->format;
    }

    /**
     * @param \EasyExcel\Metadata\Style\Format $format
     * @return $this
     */
    public function setFormat(Format $format): self
    {
        $this->format = $format;

        return $this;
    }

    /**
     * @return \EasyExcel\Metadata\Style\Borders
     */
    public function getBorders(): Borders
    {
        return $this->borders;
    }

    /**
     * @param \EasyExcel\Metadata\Style\Borders $borders
     * @return $this
     */
    public function setBorders(Borders $borders): self
    {
        $this->borders = $borders;

        return $this;
    }

    /**
     * @return \EasyExcel\Metadata\Style\Alignment
     */
    public function getAlignment(): Alignment
    {
        return $this->alignment;
    }

    /**
     * @param \EasyExcel\Metadata\Style\Alignment $alignment
     * @return $this
     */
    public function setAlignment(Alignment $alignment): self
    {
        $this->alignment = $alignment;

        return $this;
    }

    /**
     * @return \EasyExcel\Metadata\Style\Protection
     */
    public function getProtection(): Protection
    {
        return $this->protection;
    }

    /**
     * @param \EasyExcel\Metadata\Style\Protection $protection
     * @return $this
     */
    public function setProtection(Protection $protection): self
    {
        $this->protection = $protection;

        return $this;
    }

    /**
     * @return bool
     */
    public function getQuotePrefix(): bool
    {
        return $this->quotePrefix;
    }

    /**
     * @param bool $quotePrefix
     * @return $this
     */
    public function setQuotePrefix(bool $quotePrefix): self
    {
        $this->quotePrefix = $quotePrefix;

        return $this;
    }

    /**
     * @return string
     */
    public function getHashCode(): string
    {
        return md5(
            __CLASS__ .
            $this->getFont()->getHashCode() .
            $this->getFill()->getHashCode() .
            $this->getFormat()->getHashCode() .
            $this->getBorders()->getHashCode() .
            $this->getAlignment()->getHashCode() .
            $this->getProtection()->getHashCode() .
            (int) $this->getQuotePrefix()
        );
    }
}