<?php

namespace EasyExcel\Metadata\Style;

class Font
{
    // Underline types
    public const UNDERLINE_NONE = 'none';
    public const UNDERLINE_DOUBLE = 'double';
    public const UNDERLINE_DOUBLEACCOUNTING = 'doubleAccounting';
    public const UNDERLINE_SINGLE = 'single';
    public const UNDERLINE_SINGLEACCOUNTING = 'singleAccounting';

    /**
     * Font Name.
     *
     * @var string
     */
    protected $name = 'Calibri';

    /**
     * Font Size.
     *
     * @var float
     */
    protected $size = 11.00;

    /**
     * Bold.
     *
     * @var bool
     */
    protected $bold = false;

    /**
     * Italic.
     *
     * @var bool
     */
    protected $italic = false;

    /**
     * Superscript.
     *
     * @var bool
     */
    protected $superscript = false;

    /**
     * Subscript.
     *
     * @var bool
     */
    protected $subscript = false;

    /**
     * Strikethrough.
     *
     * @var bool
     */
    protected $strikethrough = false;

    /**
     * Underline.
     *
     * @var string
     */
    protected $underline = self::UNDERLINE_NONE;

    /**
     * Foreground color.
     *
     * @var Color
     */
    protected $color;

    /**
     * Constructor.
     */
    public function __construct()
    {
        $this->color = new Color();
    }

    /**
     * @return string
     */
    public function getName(): string
    {
        return $this->name;
    }

    /**
     * @param string $name
     * @return $this
     */
    public function setName(string $name): self
    {
        $this->name = $name;

        return $this;
    }

    /**
     * @return float
     */
    public function getSize(): float
    {
        return $this->size;
    }

    /**
     * @param float $size
     * @return $this
     */
    public function setSize(float $size): self
    {
        $this->size = $size;

        return $this;
    }

    /**
     * @return bool
     */
    public function getBold(): bool
    {
        return $this->bold;
    }

    /**
     * @param bool $bold
     * @return $this
     */
    public function setBold(bool $bold): self
    {
        $this->bold = $bold;

        return $this;
    }

    /**
     * @return bool
     */
    public function getItalic(): bool
    {
        return $this->italic;
    }

    /**
     * @param bool $italic
     * @return $this
     */
    public function setItalic(bool $italic): self
    {
        $this->italic = $italic;

        return $this;
    }

    /**
     * @return bool
     */
    public function getSuperscript(): bool
    {
        return $this->superscript;
    }

    /**
     * @param bool $superscript
     * @return $this
     */
    public function setSuperscript(bool $superscript): self
    {
        $this->superscript = $superscript;

        return $this;
    }

    /**
     * @return bool
     */
    public function getSubscript(): bool
    {
        return $this->subscript;
    }

    /**
     * @param bool $subscript
     * @return $this
     */
    public function setSubscript(bool $subscript): self
    {
        $this->subscript = $subscript;

        return $this;
    }

    /**
     * @return string
     */
    public function getUnderline(): string
    {
        return $this->underline;
    }

    /**
     * @param string $underline
     * @return $this
     */
    public function setUnderline(string $underline): self
    {
        $this->underline = $underline;

        return $this;
    }

    /**
     * @return bool
     */
    public function getStrikethrough(): bool
    {
        return $this->strikethrough;
    }

    /**
     * @param bool $strikethrough
     * @return $this
     */
    public function setStrikethrough(bool $strikethrough): self
    {
        $this->strikethrough = $strikethrough;

        return $this;
    }

    /**
     * @return \EasyExcel\Metadata\Style\Color
     */
    public function getColor(): Color
    {
        return $this->color;
    }

    /**
     * @param \EasyExcel\Metadata\Style\Color $color
     * @return $this
     */
    public function setColor(Color $color): self
    {
        $this->color = $color;

        return $this;
    }

    /**
     * @return string
     */
    public function getHashCode(): string
    {
        return md5(
            __CLASS__ .
            $this->getName() .
            $this->getSize() .
            (int) $this->getBold() .
            (int) $this->getItalic() .
            (int) $this->getSuperscript() .
            (int) $this->getSubscript() .
            $this->getUnderline() .
            $this->getStrikethrough() .
            $this->getColor()->getHashCode()
        );
    }
}