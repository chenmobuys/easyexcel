<?php

namespace EasyExcel\Metadata\Style;

class Border
{
    // Border style
    public const BORDER_NONE = 'none';
    public const BORDER_DASHDOT = 'dashDot';
    public const BORDER_DASHDOTDOT = 'dashDotDot';
    public const BORDER_DASHED = 'dashed';
    public const BORDER_DOTTED = 'dotted';
    public const BORDER_DOUBLE = 'double';
    public const BORDER_HAIR = 'hair';
    public const BORDER_MEDIUM = 'medium';
    public const BORDER_MEDIUMDASHDOT = 'mediumDashDot';
    public const BORDER_MEDIUMDASHDOTDOT = 'mediumDashDotDot';
    public const BORDER_MEDIUMDASHED = 'mediumDashed';
    public const BORDER_SLANTDASHDOT = 'slantDashDot';
    public const BORDER_THICK = 'thick';
    public const BORDER_THIN = 'thin';

    /**
     * Border style.
     *
     * @var string
     */
    protected $style = self::BORDER_NONE;

    /**
     * Border color.
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
    public function getStyle(): string
    {
        return $this->style;
    }

    /**
     * @param string $style
     * @return $this
     */
    public function setStyle(string $style): self
    {
        $this->style = $style;

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
            $this->getStyle() .
            $this->getColor()->getHashCode()
        );
    }
}