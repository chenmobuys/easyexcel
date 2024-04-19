<?php

namespace EasyExcel\Metadata;

use EasyExcel\Metadata\Style\Alignment;
use EasyExcel\Metadata\Style\Border;
use EasyExcel\Metadata\Style\Borders;
use EasyExcel\Metadata\Style\Color;
use EasyExcel\Metadata\Style\Fill;
use EasyExcel\Metadata\Style\Font;
use EasyExcel\Metadata\Style\Format;

class StyleBuilder
{
    /**
     * @var Style
     */
    private $style;

    /**
     * Constructor.
     *
     * @param  Style  $style
     */
    public function __construct(Style $style)
    {
        $this->style = $style;
    }

    /**
     * Set font name.
     *
     * @param  string  $fontName  {@see Font}
     *
     * @return $this
     */
    public function setFontName(string $fontName): StyleBuilder
    {
        $this->style->getFont()->setName($fontName);

        return $this;
    }

    /**
     * Set font size.
     *
     * @param  float  $fontSize  {@see Font}
     *
     * @return $this
     */
    public function setFontSize(float $fontSize): StyleBuilder
    {
        $this->style->getFont()->setSize($fontSize);

        return $this;
    }

    /**
     * Set font color.
     *
     * @param  string  $fontColor  {@see Font}
     *
     * @return $this
     */
    public function setFontColor(string $fontColor): StyleBuilder
    {
        $this->style->getFont()->setColor(new Color($fontColor));

        return $this;
    }

    /**
     * Set font bold.
     *
     * @param  bool  $fontBold  {@see Font}
     *
     * @return $this
     */
    public function setFontBold(bool $fontBold): StyleBuilder
    {
        $this->style->getFont()->setBold($fontBold);

        return $this;
    }

    /**
     * Set font italic.
     *
     * @param  bool  $fontItalic  {@see Font}
     *
     * @return $this
     */
    public function setFontItalic(bool $fontItalic): StyleBuilder
    {
        $this->style->getFont()->setItalic($fontItalic);

        return $this;
    }

    /**
     * Set font superscript.
     *
     * @param  bool  $fontSuperscript  {@see Font}
     *
     * @return $this
     */
    public function setFontSuperscript(bool $fontSuperscript): StyleBuilder
    {
        $this->style->getFont()->setSuperscript($fontSuperscript);

        return $this;
    }

    /**
     * Set font subscript.
     *
     * @param  bool  $fontSubscript  {@see Font}
     *
     * @return $this
     */
    public function setFontSubscript(bool $fontSubscript): StyleBuilder
    {
        $this->style->getFont()->setSubscript($fontSubscript);

        return $this;
    }

    /**
     * Set font strikethrough.
     *
     * @param  bool  $fontStrikethrough  {@see Font}
     *
     * @return $this
     */
    public function setFontStrikethrough(bool $fontStrikethrough): StyleBuilder
    {
        $this->style->getFont()->setStrikethrough($fontStrikethrough);

        return $this;
    }

    /**
     * Set font underline.
     *
     * @param  string  $fontUnderline  {@see Font}
     *
     * @return $this
     */
    public function setFontUnderline(string $fontUnderline): StyleBuilder
    {
        $this->style->getFont()->setUnderline($fontUnderline);

        return $this;
    }

    /**
     * Set fill type.
     *
     * @param  string  $fillType  {@see Fill}
     *
     * @return $this
     */
    public function setFillType(string $fillType): StyleBuilder
    {
        $this->style->getFill()->setType($fillType);

        return $this;
    }

    /**
     * Set fill rotation.
     *
     * @param  float  $fillRotation  {@see Fill}
     *
     * @return $this
     */
    public function setFillRotation(float $fillRotation): StyleBuilder
    {
        $this->style->getFill()->setRotation($fillRotation);

        return $this;
    }

    /**
     * Set fill start color.
     *
     * @param  string  $startColor  {@see Fill}
     *
     * @return $this
     */
    public function setFillStartColor(string $startColor): StyleBuilder
    {
        $this->style->getFill()->setStartColor(new Color($startColor));

        return $this;
    }

    /**
     * Set fill end color.
     *
     * @param  string  $endColor  {@see Fill}
     *
     * @return $this
     */
    public function setFillEndColor(string $endColor): StyleBuilder
    {
        $this->style->getFill()->setEndColor(new Color($endColor));

        return $this;
    }

    /**
     * Set format code.
     *
     * @param  string  $formatCode  {@see Format}
     *
     * @return $this
     */
    public function setFormatCode(string $formatCode): StyleBuilder
    {
        $this->style->getFormat()->setFormatCode($formatCode);

        return $this;
    }

    /**
     * Set left border style.
     *
     * @param  string  $borderStyle  {@see Border}
     *
     * @return $this
     */
    public function setLeftBorderStyle(string $borderStyle): StyleBuilder
    {
        $this->style->getBorders()->getLeft()->setStyle($borderStyle);

        return $this;
    }

    /**
     * Set left border color.
     *
     * @param  string  $borderColor  {@see Border}
     *
     * @return $this
     */
    public function setLeftBorderColor(string $borderColor): StyleBuilder
    {
        $this->style->getBorders()->getLeft()->setColor(new Color($borderColor));

        return $this;
    }

    /**
     * Set right border style.
     *
     * @param  string  $borderStyle  {@see Border}
     *
     * @return $this
     */
    public function setRightBorderStyle(string $borderStyle): StyleBuilder
    {
        $this->style->getBorders()->getRight()->setStyle($borderStyle);

        return $this;
    }

    /**
     * Set right border color.
     *
     * @param  string  $borderColor  {@see Border}
     *
     * @return $this
     */
    public function setRightBorderColor(string $borderColor): StyleBuilder
    {
        $this->style->getBorders()->getRight()->setColor(new Color($borderColor));

        return $this;
    }

    /**
     * Set top border style.
     *
     * @param  string  $borderStyle  {@see Border}
     *
     * @return $this
     */
    public function setTopBorderStyle(string $borderStyle): StyleBuilder
    {
        $this->style->getBorders()->getTop()->setStyle($borderStyle);

        return $this;
    }

    /**
     * Set top border color.
     *
     * @param  string  $borderColor  {@see Border}
     *
     * @return $this
     */
    public function setTopBorderColor(string $borderColor): StyleBuilder
    {
        $this->style->getBorders()->getTop()->setColor(new Color($borderColor));

        return $this;
    }

    /**
     * Set bottom border style.
     *
     * @param  string  $borderStyle  {@see Border}
     *
     * @return $this
     */
    public function setBottomBorderStyle(string $borderStyle): StyleBuilder
    {
        $this->style->getBorders()->getBottom()->setStyle($borderStyle);

        return $this;
    }

    /**
     * Set bottom border color.
     *
     * @param  string  $borderColor  {@see Border}
     *
     * @return $this
     */
    public function setBottomBorderColor(string $borderColor): StyleBuilder
    {
        $this->style->getBorders()->getBottom()->setColor(new Color($borderColor));

        return $this;
    }

    /**
     * Set horizontal border style.
     *
     * @param  string  $borderStyle  {@see Border}
     *
     * @return $this
     */
    public function setHorizontalBorderStyle(string $borderStyle): StyleBuilder
    {
        $this->style->getBorders()->getHorizontal()->setStyle($borderStyle);

        return $this;
    }

    /**
     * Set horizontal border color.
     *
     * @param  string  $borderColor  {@see Border}
     *
     * @return $this
     */
    public function setHorizontalBorderColor(string $borderColor): StyleBuilder
    {
        $this->style->getBorders()->getHorizontal()->setColor(new Color($borderColor));

        return $this;
    }

    /**
     * Set vertical border style.
     *
     * @param  string  $borderStyle  {@see Border}
     *
     * @return $this
     */
    public function setVerticalBorderStyle(string $borderStyle): StyleBuilder
    {
        $this->style->getBorders()->getVertical()->setStyle($borderStyle);

        return $this;
    }

    /**
     * Set vertical border color.
     *
     * @param  string  $borderColor  {@see Border}
     *
     * @return $this
     */
    public function setVerticalBorderColor(string $borderColor): StyleBuilder
    {
        $this->style->getBorders()->getVertical()->setColor(new Color($borderColor));

        return $this;
    }

    /**
     * Set diagonal border style.
     *
     * @param  string  $borderStyle  {@see Border}
     *
     * @return $this
     */
    public function setDiagonalBorderStyle(string $borderStyle): StyleBuilder
    {
        $this->style->getBorders()->getDiagonal()->setStyle($borderStyle);

        return $this;
    }

    /**
     * Set diagonal border color.
     *
     * @param  string  $borderColor  {@see Border}
     *
     * @return $this
     */
    public function setDiagonalBorderColor(string $borderColor): StyleBuilder
    {
        $this->style->getBorders()->getDiagonal()->setColor(new Color($borderColor));

        return $this;
    }

    /**
     * Set diagonal border direction.
     *
     * @param  int  $borderDirection  {@see Borders}
     *
     * @return $this
     */
    public function setDiagonalBorderDirection(int $borderDirection): StyleBuilder
    {
        $this->style->getBorders()->setDiagonalDirection($borderDirection);

        return $this;
    }

    /**
     * Set alignment horizontal.
     *
     * @param  string  $alignmentHorizontal  {@see Alignment}
     *
     * @return $this
     */
    public function setAlignmentHorizontal(string $alignmentHorizontal): StyleBuilder
    {
        $this->style->getAlignment()->setHorizontal($alignmentHorizontal);

        return $this;
    }

    /**
     * Set alignment vertical.
     *
     * @param  string  $alignmentVertical  {@see Alignment}
     *
     * @return $this
     */
    public function setAlignmentVertical(string $alignmentVertical): StyleBuilder
    {
        $this->style->getAlignment()->setVertical($alignmentVertical);

        return $this;
    }

    /**
     * Set alignment text rotation.
     *
     * @param  int  $alignmentTextRotation
     *
     * @return $this
     */
    public function setAlignmentTextRotation(int $alignmentTextRotation): StyleBuilder
    {
        $this->style->getAlignment()->setTextRotation($alignmentTextRotation);

        return $this;
    }

    /**
     * Set alignment wrap text.
     *
     * @param  bool  $alignmentWrapText
     *
     * @return $this
     */
    public function setAlignmentWrapText(bool $alignmentWrapText): StyleBuilder
    {
        $this->style->getAlignment()->setWrapText($alignmentWrapText);

        return $this;
    }

    /**
     * Set alignment shrink to fit.
     *
     * @param  bool  $alignmentShrinkToFit
     *
     * @return $this
     */
    public function setAlignmentShrinkToFit(bool $alignmentShrinkToFit): StyleBuilder
    {
        $this->style->getAlignment()->setShrinkToFit($alignmentShrinkToFit);

        return $this;
    }

    /**
     * Set alignment indent.
     *
     * @param  int  $alignmentIndent
     *
     * @return $this
     */
    public function setAlignmentIndent(int $alignmentIndent): StyleBuilder
    {
        $this->style->getAlignment()->setIndent($alignmentIndent);

        return $this;
    }

    /**
     * Set alignment read order.
     *
     * @param  string  $alignmentReadOrder  {@link Alignment::READORDER_CONTEXT, Alignment::READORDER_LTR, Alignment::READORDER_RTL}
     *
     * @return $this
     */
    public function setAlignmentReadOrder(string $alignmentReadOrder): StyleBuilder
    {
        $this->style->getAlignment()->setReadOrder($alignmentReadOrder);

        return $this;
    }

    /**
     * Set protection locked.
     *
     * @param  string  $protectionLocked
     *
     * @return $this
     */
    public function setProtectionLocked(string $protectionLocked): StyleBuilder
    {
        $this->style->getProtection()->setLocked($protectionLocked);

        return $this;
    }

    /**
     * Set protection hidden.
     *
     * @param  string  $protectionHidden
     *
     * @return $this
     */
    public function setProtectionHidden(string $protectionHidden): StyleBuilder
    {
        $this->style->getProtection()->setHidden($protectionHidden);

        return $this;
    }

    /**
     * Set quote prefix.
     *
     * @param  bool  $quotePrefix
     *
     * @return $this
     */
    public function setQuotePrefix(bool $quotePrefix): StyleBuilder
    {
        $this->style->setQuotePrefix($quotePrefix);

        return $this;
    }

    /**
     * Build style.
     *
     * @return Style
     */
    public function build(): Style
    {
        return $this->style;
    }
}