<?php

namespace EasyExcel\Metadata\Style;

class Fill
{
    // Fill types
    public const FILL_NONE = 'none';
    public const FILL_SOLID = 'solid';
    public const FILL_GRADIENT_LINEAR = 'linear';
    public const FILL_GRADIENT_PATH = 'path';
    public const FILL_PATTERN_DARKDOWN = 'darkDown';
    public const FILL_PATTERN_DARKGRAY = 'darkGray';
    public const FILL_PATTERN_DARKGRID = 'darkGrid';
    public const FILL_PATTERN_DARKHORIZONTAL = 'darkHorizontal';
    public const FILL_PATTERN_DARKTRELLIS = 'darkTrellis';
    public const FILL_PATTERN_DARKUP = 'darkUp';
    public const FILL_PATTERN_DARKVERTICAL = 'darkVertical';
    public const FILL_PATTERN_GRAY0625 = 'gray0625';
    public const FILL_PATTERN_GRAY125 = 'gray125';
    public const FILL_PATTERN_LIGHTDOWN = 'lightDown';
    public const FILL_PATTERN_LIGHTGRAY = 'lightGray';
    public const FILL_PATTERN_LIGHTGRID = 'lightGrid';
    public const FILL_PATTERN_LIGHTHORIZONTAL = 'lightHorizontal';
    public const FILL_PATTERN_LIGHTTRELLIS = 'lightTrellis';
    public const FILL_PATTERN_LIGHTUP = 'lightUp';
    public const FILL_PATTERN_LIGHTVERTICAL = 'lightVertical';
    public const FILL_PATTERN_MEDIUMGRAY = 'mediumGray';

    /**
     * Fill type.
     *
     * @var string
     */
    protected $type = self::FILL_NONE;
    /**
     * Rotation.
     *
     * @var float
     */
    protected $rotation = 0.0;

    /**
     * @var \EasyExcel\Metadata\Style\Color
     */
    protected $startColor;

    /**
     * @var \EasyExcel\Metadata\Style\Color
     */
    protected $endColor;

    /**
     * Constructor.
     */
    public function __construct()
    {
        $this->startColor = new Color(Color::WHITE);
        $this->endColor = new Color(Color::BLACK);
    }

    /**
     * @return string
     */
    public function getType(): string
    {
        return $this->type;
    }

    /**
     * @param string $type
     * @return $this
     */
    public function setType(string $type): self
    {
        $this->type = $type;

        return $this;
    }

    /**
     * @return float
     */
    public function getRotation(): float
    {
        return $this->rotation;
    }

    /**
     * @param float $rotation
     * @return $this
     */
    public function setRotation(float $rotation): self
    {
        $this->rotation = $rotation;

        return $this;
    }

    /**
     * @return \EasyExcel\Metadata\Style\Color
     */
    public function getStartColor(): Color
    {
        return $this->startColor;
    }

    /**
     * @param \EasyExcel\Metadata\Style\Color $startColor
     * @return $this
     */
    public function setStartColor(Color $startColor): self
    {
        $this->startColor = $startColor;

        return $this;
    }

    /**
     * @return \EasyExcel\Metadata\Style\Color
     */
    public function getEndColor(): Color
    {
        return $this->endColor;
    }

    /**
     * @param \EasyExcel\Metadata\Style\Color $endColor
     * @return $this
     */
    public function setEndColor(Color $endColor): self
    {
        $this->endColor = $endColor;

        return $this;
    }

    /**
     * @return string
     */
    public function getHashCode(): string
    {
        return md5(
            __CLASS__ .
            $this->getType() .
            $this->getRotation() .
            $this->getStartColor()->getHashCode() .
            $this->getEndColor()->getHashCode()
        );
    }
}