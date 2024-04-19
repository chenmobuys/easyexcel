<?php

namespace EasyExcel\Metadata\Style;

class Borders
{
    // Diagonal directions
    public const DIAGONAL_NONE = 0;
    public const DIAGONAL_UP = 1;
    public const DIAGONAL_DOWN = 2;
    public const DIAGONAL_BOTH = 3;

    /**
     * Left.
     *
     * @var Border
     */
    protected $left;

    /**
     * Right.
     *
     * @var Border
     */
    protected $right;

    /**
     * Top.
     *
     * @var Border
     */
    protected $top;

    /**
     * Bottom.
     *
     * @var Border
     */
    protected $bottom;

    /**
     * Diagonal.
     *
     * @var Border
     */
    protected $diagonal;

    /**
     * DiagonalDirection.
     *
     * @var int
     */
    protected $diagonalDirection = self::DIAGONAL_NONE;

    /**
     * Vertical pseudo-border. Only applies to supervisor.
     *
     * @var Border
     */
    protected $vertical;

    /**
     * Horizontal pseudo-border. Only applies to supervisor.
     *
     * @var Border
     */
    protected $horizontal;

    /**
     * Constructor.
     */
    public function __construct()
    {
        $this->left = new Border();
        $this->right = new Border();
        $this->top = new Border();
        $this->bottom = new Border();
        $this->diagonal = new Border();
        $this->vertical = new Border();
        $this->horizontal = new Border();
        $this->diagonalDirection = self::DIAGONAL_NONE;
    }

    /**
     * @return \EasyExcel\Metadata\Style\Border
     */
    public function getLeft(): ?Border
    {
        return $this->left;
    }

    /**
     * @param \EasyExcel\Metadata\Style\Border $left
     * @return $this
     */
    public function setLeft(Border $left): self
    {
        $this->left = $left;

        return $this;
    }

    /**
     * @return \EasyExcel\Metadata\Style\Border
     */
    public function getRight(): ?Border
    {
        return $this->right;
    }

    /**
     * @param \EasyExcel\Metadata\Style\Border $right
     * @return $this
     */
    public function setRight(Border $right): self
    {
        $this->right = $right;

        return $this;
    }

    /**
     * @return \EasyExcel\Metadata\Style\Border
     */
    public function getTop(): ?Border
    {
        return $this->top;
    }

    /**
     * @param \EasyExcel\Metadata\Style\Border $top
     * @return $this
     */
    public function setTop(Border $top): self
    {
        $this->top = $top;

        return $this;
    }

    /**
     * @return \EasyExcel\Metadata\Style\Border
     */
    public function getBottom(): ?Border
    {
        return $this->bottom;
    }

    /**
     * @param \EasyExcel\Metadata\Style\Border $bottom
     * @return $this
     */
    public function setBottom(Border $bottom): self
    {
        $this->bottom = $bottom;

        return $this;
    }

    /**
     * @return \EasyExcel\Metadata\Style\Border
     */
    public function getDiagonal(): ?Border
    {
        return $this->diagonal;
    }

    /**
     * @param \EasyExcel\Metadata\Style\Border $diagonal
     * @return $this
     */
    public function setDiagonal(Border $diagonal): self
    {
        $this->diagonal = $diagonal;

        return $this;
    }

    /**
     * @return int
     */
    public function getDiagonalDirection(): int
    {
        return $this->diagonalDirection;
    }

    /**
     * @param int $diagonalDirection
     * @return $this
     */
    public function setDiagonalDirection(int $diagonalDirection): self
    {
        $this->diagonalDirection = $diagonalDirection;

        return $this;
    }

    /**
     * @return \EasyExcel\Metadata\Style\Border
     */
    public function getVertical(): ?Border
    {
        return $this->vertical;
    }

    /**
     * @param \EasyExcel\Metadata\Style\Border $vertical
     * @return $this
     */
    public function setVertical(Border $vertical): self
    {
        $this->vertical = $vertical;

        return $this;
    }

    /**
     * @return \EasyExcel\Metadata\Style\Border
     */
    public function getHorizontal(): ?Border
    {
        return $this->horizontal;
    }

    /**
     * @param \EasyExcel\Metadata\Style\Border $horizontal
     * @return $this
     */
    public function setHorizontal(Border $horizontal): self
    {
        $this->horizontal = $horizontal;

        return $this;
    }

    /**
     * @return string
     */
    public function getHashCode(): string
    {
        return md5(
            __CLASS__ .
            $this->getLeft()->getHashCode() .
            $this->getRight()->getHashCode() .
            $this->getTop()->getHashCode() .
            $this->getBottom()->getHashCode() .
            $this->getDiagonal()->getHashCode() .
            $this->getDiagonalDirection()
        );
    }
}