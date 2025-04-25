<?php

namespace EasyExcel\Metadata\Style;

class Alignment
{
    // Horizontal alignment styles
    public const HORIZONTAL_GENERAL = 'general';
    public const HORIZONTAL_LEFT = 'left';
    public const HORIZONTAL_RIGHT = 'right';
    public const HORIZONTAL_CENTER = 'center';
    public const HORIZONTAL_CENTER_CONTINUOUS = 'centerContinuous';
    public const HORIZONTAL_JUSTIFY = 'justify';
    public const HORIZONTAL_FILL = 'fill';
    public const HORIZONTAL_DISTRIBUTED = 'distributed'; // Excel2007 only

    // Vertical alignment styles
    public const VERTICAL_BOTTOM = 'bottom';
    public const VERTICAL_TOP = 'top';
    public const VERTICAL_CENTER = 'center';
    public const VERTICAL_JUSTIFY = 'justify';
    public const VERTICAL_DISTRIBUTED = 'distributed'; // Excel2007 only

    // Read order
    public const READORDER_CONTEXT = 0;
    public const READORDER_LTR = 1;
    public const READORDER_RTL = 2;

    // Special value for Text Rotation
    public const TEXTROTATION_STACK_EXCEL = 255;
    public const TEXTROTATION_STACK_PHPSPREADSHEET = -165; // 90 - 255

    /**
     * Horizontal alignment.
     *
     * @var string
     */
    protected $horizontal = self::HORIZONTAL_GENERAL;

    /**
     * Vertical alignment.
     *
     * @var string
     */
    protected $vertical = self::VERTICAL_BOTTOM;

    /**
     * Text rotation.
     *
     * @var int
     */
    protected $textRotation = 0;

    /**
     * Wrap text.
     *
     * @var bool
     */
    protected $wrapText = false;

    /**
     * Shrink to fit.
     *
     * @var bool
     */
    protected $shrinkToFit = false;

    /**
     * Indent - only possible with horizontal alignment left and right.
     *
     * @var int
     */
    protected $indent = 0;

    /**
     * Read order.
     *
     * @var string
     */
    protected $readOrder = self::READORDER_CONTEXT;

    /**
     * @return string
     */
    public function getHorizontal(): string
    {
        return $this->horizontal;
    }

    /**
     * @param string $horizontal
     * @return $this
     */
    public function setHorizontal(string $horizontal): self
    {
        $this->horizontal = $horizontal;

        return $this;
    }

    /**
     * @return string
     */
    public function getVertical(): string
    {
        return $this->vertical;
    }

    /**
     * @param string $vertical
     * @return $this
     */
    public function setVertical(string $vertical): self
    {
        $this->vertical = $vertical;

        return $this;
    }

    /**
     * @return int
     */
    public function getTextRotation(): int
    {
        return $this->textRotation;
    }

    /**
     * @param int $textRotation
     * @return $this
     */
    public function setTextRotation(int $textRotation): self
    {
        $this->textRotation = $textRotation;

        return $this;
    }

    /**
     * @return bool
     */
    public function isWrapText(): bool
    {
        return $this->wrapText;
    }

    /**
     * @param bool $wrapText
     * @return $this
     */
    public function setWrapText(bool $wrapText): self
    {
        $this->wrapText = $wrapText;

        return $this;
    }

    /**
     * @return bool
     */
    public function isShrinkToFit(): bool
    {
        return $this->shrinkToFit;
    }

    /**
     * @param bool $shrinkToFit
     * @return $this
     */
    public function setShrinkToFit(bool $shrinkToFit): self
    {
        $this->shrinkToFit = $shrinkToFit;

        return $this;
    }

    /**
     * @return int
     */
    public function getIndent(): int
    {
        return $this->indent;
    }

    /**
     * @param int $indent
     * @return $this
     */
    public function setIndent(int $indent): self
    {
        $this->indent = $indent;

        return $this;
    }

    /**
     * @return string
     */
    public function getReadOrder(): string
    {
        return $this->readOrder;
    }

    /**
     * @param string $readOrder
     * @return $this
     */
    public function setReadOrder(string $readOrder): self
    {
        $this->readOrder = $readOrder;

        return $this;
    }

    /**
     * @return string
     */
    public function getHashCode(): string
    {
        return md5(
            __CLASS__ .
            $this->getHorizontal() .
            $this->getVertical() .
            $this->getReadOrder() .
            $this->getIndent() .
            $this->getTextRotation()
        );
    }
}