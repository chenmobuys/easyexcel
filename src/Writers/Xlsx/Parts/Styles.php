<?php

namespace EasyExcel\Writers\Xlsx\Parts;

use EasyExcel\Metadata\Style;
use EasyExcel\Metadata\Style\Border;
use EasyExcel\Metadata\Style\Borders;
use EasyExcel\Metadata\Style\Fill;
use EasyExcel\Metadata\Style\Protection;

class Styles extends AbstractPart
{
    /**
     * @var array
     */
    protected $numFmtIndexes = [];

    /**
     * @var array
     */
    protected $fontIndexes = [];

    /**
     * @var array
     */
    protected $fillIndexes = [];

    /**
     * @var array
     */
    protected $borderIndexes = [];

    /**
     * @var array
     */
    protected $cellXfs = [];

    /**
     * Write start.
     *
     * @return $this
     */
    protected function writeStart(): parent
    {
        $this->xml->startElement('styleSheet');
        $this->xml->writeAttribute('xmlns', self::NS_MAIN);

        $this->writeNumFmts();

        $this->writeFonts();

        $this->writeFills();

        $this->writeBorders();

        $this->writeCellXfs();

        return $this;
    }

    /**
     * @return void
     */
    protected function writeNumFmts(): void
    {
        $numFmts = [];
        foreach ($this->excel->getCellXfs() as $cellXf) {
            $hashCode = $cellXf->getFormat()->getHashCode();
            $numFmts[$hashCode] = $cellXf->getFormat();
            $this->numFmtIndexes[$hashCode] = array_search($hashCode, array_keys($numFmts));
        }
        $this->xml->startElement('numFmts');
        $this->xml->writeAttribute('count', count($this->numFmtIndexes));

        foreach ($this->numFmtIndexes as $hashCode => $index) {
            $this->xml->startElement('numFmt');
            $this->xml->writeAttribute('numFmtId', $index + 164);
            $this->xml->writeAttribute('formatCode', $numFmts[$hashCode]->getFormatCode());
            $this->xml->endElement();
        }
        $this->xml->endElement();
    }

    /**
     * @return void
     */
    protected function writeFonts(): void
    {
        $fonts = [];
        foreach ($this->excel->getCellXfs() as $cellXf) {
            $hashCode = $cellXf->getFont()->getHashCode();
            $fonts[$hashCode] = $cellXf->getFont();
            $this->fontIndexes[$hashCode] = array_search($hashCode, array_keys($fonts));
        }
        $this->xml->startElement('fonts');
        $this->xml->writeAttribute('count', count($this->fontIndexes));

        foreach ($fonts as $font) {

            $this->xml->startElement('font');

            if ($font->getBold()) {
                $this->writeElementWithAttributes('b', ['val' => 1]);
            }

            if ($font->getItalic()) {
                $this->writeElementWithAttributes('i', ['val' => 1]);
            }

            if ($font->getStrikethrough()) {
                $this->writeElementWithAttributes('strike', ['val' => 1]);
            }

            if ($font->getUnderline()) {
                $this->writeElementWithAttributes('u', ['val' => $font->getUnderline()]);
            }

            if ($font->getSuperscript() || $font->getSubscript()) {
                $this->writeElementWithAttributes('vertAlign', [
                    'val' => $font->getSuperscript() ? 'superscript' : 'subscript',
                ]);
            }

            if ($font->getName()) {
                $this->writeElementWithAttributes('name', ['val' => $font->getName()]);
            }

            if ($font->getSize()) {
                $this->writeElementWithAttributes('sz', ['val' => $font->getSize()]);
            }

            if ($font->getColor()->getArgb()) {
                $this->writeElementWithAttributes('color', ['rgb' => $font->getColor()->getArgb()]);
            }

            $this->xml->endElement();
        }
        $this->xml->endElement();
    }

    /**
     * @return void
     */
    protected function writeFills(): void
    {
        $fills = [];
        foreach ($this->excel->getCellXfs() as $cellXf) {
            $hashCode = $cellXf->getFill()->getHashCode();
            $fills[$hashCode] = $cellXf->getFill();
            $this->fillIndexes[$hashCode] = array_search($hashCode, array_keys($fills));
        }
        $this->xml->startElement('fills');
        $this->xml->writeAttribute('count', count($this->fillIndexes));

        foreach ($fills as $fill) {
            if (in_array($fill->getType(), [
                Fill::FILL_GRADIENT_PATH,
                Fill::FILL_GRADIENT_LINEAR,
            ])) {
                // fill
                $this->xml->startElement('fill');

                // gradientFill
                $this->xml->startElement('gradientFill');
                $this->xml->writeAttribute('type', $fill->getType());
                $this->xml->writeAttribute('degree', $fill->getRotation());

                // stop
                $this->xml->startElement('stop');
                $this->xml->writeAttribute('position', '0');

                // color
                $this->xml->startElement('color');
                $this->xml->writeAttribute('rgb', $fill->getStartColor()->getArgb());
                $this->xml->endElement();

                $this->xml->endElement();

                // stop
                $this->xml->startElement('stop');
                $this->xml->writeAttribute('position', '1');

                // color
                $this->xml->startElement('color');
                $this->xml->writeAttribute('rgb', $fill->getEndColor()->getArgb());
                $this->xml->endElement();

                $this->xml->endElement();

                $this->xml->endElement();

                $this->xml->endElement();
            } elseif ($fill->getType()) {

                // fill
                $this->xml->startElement('fill');

                // patternFill
                $this->xml->startElement('patternFill');
                $this->xml->writeAttribute('patternType', $fill->getType());

                if ($fill->getType() !== Fill::FILL_NONE) {
                    // fgColor
                    if ($fill->getStartColor()->getARGB()) {
                        $this->xml->startElement('fgColor');
                        $this->xml->writeAttribute('rgb', $fill->getStartColor()->getARGB());
                        $this->xml->endElement();
                    }
                }
                if ($fill->getType() !== Fill::FILL_NONE) {
                    // bgColor
                    if ($fill->getEndColor()->getARGB()) {
                        $this->xml->startElement('bgColor');
                        $this->xml->writeAttribute('rgb', $fill->getEndColor()->getARGB());
                        $this->xml->endElement();
                    }
                }

                $this->xml->endElement();

                $this->xml->endElement();
            }
        }

        $this->xml->endElement();
    }

    /**
     * @return void
     */
    protected function writeBorders()
    {
        $allBorders = [];
        foreach ($this->excel->getCellXfs() as $cellXf) {
            $hashCode = $cellXf->getBorders()->getHashCode();
            $allBorders[$hashCode] = $cellXf->getBorders();
            $this->borderIndexes[$hashCode] = array_search($hashCode, array_keys($allBorders));
        }

        $this->xml->startElement('borders');

        foreach ($allBorders as $borders) {
            $this->xml->startElement('border');

            switch ($borders->getDiagonalDirection()) {
                case Borders::DIAGONAL_UP:
                    $this->xml->writeAttribute('diagonalUp', 1);
                    $this->xml->writeAttribute('diagonalDown', 0);
                    break;
                case Borders::DIAGONAL_DOWN:
                    $this->xml->writeAttribute('diagonalUp', 0);
                    $this->xml->writeAttribute('diagonalDown', 1);
                    break;
                case Borders::DIAGONAL_BOTH:
                    $this->xml->writeAttribute('diagonalUp', 1);
                    $this->xml->writeAttribute('diagonalDown', 1);
                    break;
            }

            $this->writeBorderPr($borders->getLeft(), 'left');
            $this->writeBorderPr($borders->getRight(), 'right');
            $this->writeBorderPr($borders->getTop(), 'top');
            $this->writeBorderPr($borders->getBottom(), 'bottom');
            $this->writeBorderPr($borders->getDiagonal(), 'diagonal');

            $this->xml->endElement();
        }

        $this->xml->endElement();
    }

    /**
     * @param \EasyExcel\Metadata\Style\Border $border
     * @param $name
     * @return void
     */
    protected function writeBorderPr(Border $border, $name)
    {
        $this->xml->startElement($name);
        // Write BorderPr
        if ($border->getStyle() != Border::BORDER_NONE) {
            $this->xml->writeAttribute('style', $border->getStyle());

            // color
            $this->xml->startElement('color');
            $this->xml->writeAttribute('rgb', $border->getColor()->getARGB());
            $this->xml->endElement();
        }
        $this->xml->endElement();
    }

    /**
     * @return void
     */
    protected function writeCellXfs(): void
    {
        $defaultStyle = new Style();

        $cellXfs = [];
        foreach ($this->excel->getCellXfs() as $cellXf) {
            $hashCode = $cellXf->getHashCode();
            $cellXfs[$hashCode] = $cellXf;
            $this->cellXfs[$hashCode] = array_search($hashCode, array_keys($cellXfs));
        }

        $this->xml->startElement('cellXfs');

        foreach ($cellXfs as $cellXf) {
            $this->xml->startElement('xf');
            $this->xml->writeAttribute('xfId', 0);
            if ($cellXf->getQuotePrefix()) {
                $this->xml->writeAttribute('quotePrefix', 1);
            }
            $this->xml->writeAttribute('numFmtId', (int) $this->numFmtIndexes[$cellXf->getFormat()->getHashCode()] + 164);
            $this->xml->writeAttribute('fontId', (int) $this->fontIndexes[$cellXf->getFont()->getHashCode()]);
            $this->xml->writeAttribute('fillId', (int) $this->fillIndexes[$cellXf->getFill()->getHashCode()]);
            $this->xml->writeAttribute('borderId', (int) $this->borderIndexes[$cellXf->getBorders()->getHashCode()]);

            $this->xml->writeAttribute('applyNumberFormat', $defaultStyle->getFormat()->getHashCode() != $cellXf->getFormat()->getHashCode() ? '1' : '0');
            $this->xml->writeAttribute('applyFont', $defaultStyle->getFont()->getHashCode() != $cellXf->getFont()->getHashCode() ? '1' : '0');
            $this->xml->writeAttribute('applyFill', $defaultStyle->getFill()->getHashCode() != $cellXf->getFill()->getHashCode() ? '1' : '0');
            $this->xml->writeAttribute('applyBorder', $defaultStyle->getBorders()->getHashCode() != $cellXf->getBorders()->getHashCode() ? '1' : '0');
            $this->xml->writeAttribute('applyAlignment', $defaultStyle->getAlignment()->getHashCode() != $cellXf->getAlignment()->getHashCode() ? '1' : '0');
            if ($cellXf->getProtection()->getLocked() != Protection::PROTECTION_INHERIT || $cellXf->getProtection()->getHidden() != Protection::PROTECTION_INHERIT) {
                $this->xml->writeAttribute('applyProtection', 'true');
            }

            $this->xml->startElement('alignment');
            $this->xml->writeAttribute('horizontal', $cellXf->getAlignment()->getHorizontal());
            $this->xml->writeAttribute('vertical', $cellXf->getAlignment()->getVertical());

            if ($cellXf->getAlignment()->getTextRotation() >= 0) {
                $textRotation = $cellXf->getAlignment()->getTextRotation();
            } else {
                $textRotation = 90 - $cellXf->getAlignment()->getTextRotation();
            }
            $this->xml->writeAttribute('textRotation', $textRotation);

            $this->xml->writeAttribute('wrapText', (int) $cellXf->getAlignment()->isWrapText());
            $this->xml->writeAttribute('shrinkToFit', (int) $cellXf->getAlignment()->isShrinkToFit());

            if ($cellXf->getAlignment()->getIndent() > 0) {
                $this->xml->writeAttribute('indent', $cellXf->getAlignment()->getIndent());
            }
            if ($cellXf->getAlignment()->getReadOrder() > 0) {
                $this->xml->writeAttribute('readingOrder', $cellXf->getAlignment()->getReadOrder());
            }
            $this->xml->endElement();

            if ($cellXf->getProtection()->getLocked() != Protection::PROTECTION_INHERIT || $cellXf->getProtection()->getHidden() != Protection::PROTECTION_INHERIT) {
                $this->xml->startElement('protection');
                if ($cellXf->getProtection()->getLocked() != Protection::PROTECTION_INHERIT) {
                    $this->xml->writeAttribute('locked', (int) $cellXf->getProtection()->getLocked() == Protection::PROTECTION_PROTECTED);
                }
                if ($cellXf->getProtection()->getHidden() != Protection::PROTECTION_INHERIT) {
                    $this->xml->writeAttribute('hidden', (int) $cellXf->getProtection()->getHidden() == Protection::PROTECTION_PROTECTED);
                }
                $this->xml->endElement();
            }

            $this->xml->endElement();
        }

        $this->xml->endElement();
    }

    /**
     * Write end.
     *
     * @return $this
     */
    protected function writeEnd(): parent
    {
        $this->xml->endElement();

        return $this;
    }
}