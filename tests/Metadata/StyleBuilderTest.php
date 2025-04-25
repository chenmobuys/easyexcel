<?php

namespace EasyExcelTests\Metadata;

use EasyExcel\Metadata\Style;
use PHPUnit\Framework\TestCase;

class StyleBuilderTest extends TestCase
{

    public function test_font()
    {
        $style = Style::builder()
            ->setFontName("Arial")
            ->setFontSize(12)
            ->setFontColor(Style\Color::BLACK)
            ->setFontUnderline(Style\Font::UNDERLINE_DOUBLE)
            ->setFontBold(true)
            ->setFontItalic(true)
            ->setFontSuperscript(true)
            ->setFontSubscript(true)
            ->setFontStrikethrough(true)
            ->build();
        $this->assertEquals("Arial", $style->getFont()->getName());
        $this->assertEquals(12, $style->getFont()->getSize());
        $this->assertEquals(Style\Color::BLACK, $style->getFont()->getColor()->getArgb());
        $this->assertEquals(Style\Font::UNDERLINE_DOUBLE, $style->getFont()->getUnderline());
        $this->assertTrue($style->getFont()->getBold());
        $this->assertTrue($style->getFont()->getItalic());
        $this->assertTrue($style->getFont()->getSuperscript());
        $this->assertTrue($style->getFont()->getSubscript());
        $this->assertTrue($style->getFont()->getStrikethrough());
        $this->assertNotNull($style->getFont()->getHashCode());
    }

    public function test_fill()
    {
        $style = Style::builder()
            ->setFillType(Style\Fill::FILL_GRADIENT_LINEAR)
            ->setFillRotation(5)
            ->setFillStartColor(Style\Color::BLACK)
            ->setFillEndColor(Style\Color::RED)
            ->build();
        $this->assertEquals(Style\Fill::FILL_GRADIENT_LINEAR, $style->getFill()->getType());
        $this->assertEquals(5, $style->getFill()->getRotation());
        $this->assertEquals(Style\Color::BLACK, $style->getFill()->getStartColor()->getArgb());
        $this->assertEquals(Style\Color::RED, $style->getFill()->getEndColor()->getArgb());
        $this->assertNotNull($style->getFill()->getHashCode());
    }

    public function test_format()
    {
        $style = Style::builder()
            ->setFormatCode(Style\Format::FORMAT_DATE_DDMMYYYY)
            ->build();
        $this->assertEquals(Style\Format::FORMAT_DATE_DDMMYYYY, $style->getFormat()->getFormatCode());
        $this->assertEquals(Style\Format::CALENDAR_WINDOWS_1900, $style->getFormat()->getCalendar());
        Style\Format::setCalendar(Style\Format::CALENDAR_MAC_1904);
        $this->assertEquals(Style\Format::CALENDAR_MAC_1904, Style\Format::getCalendar());
        $this->assertNotNull($style->getFormat()->getHashCode());
    }

    public function test_borders()
    {
        $style = Style::builder()
            ->setLeftBorderStyle(Style\Border::BORDER_DASHED)
            ->setLeftBorderColor(Style\Color::BLACK)
            ->setRightBorderStyle(Style\Border::BORDER_DASHED)
            ->setRightBorderColor(Style\Color::BLACK)
            ->setTopBorderStyle(Style\Border::BORDER_DASHED)
            ->setTopBorderColor(Style\Color::BLACK)
            ->setBottomBorderStyle(Style\Border::BORDER_DASHED)
            ->setBottomBorderColor(Style\Color::BLACK)
            ->setHorizontalBorderStyle(Style\Border::BORDER_DASHED)
            ->setHorizontalBorderColor(Style\Color::BLACK)
            ->setVerticalBorderStyle(Style\Border::BORDER_DASHED)
            ->setVerticalBorderColor(Style\Color::BLACK)
            ->setDiagonalBorderStyle(Style\Border::BORDER_DASHED)
            ->setDiagonalBorderColor(Style\Color::BLACK)
            ->setDiagonalBorderDirection(Style\Borders::DIAGONAL_BOTH)
            ->build();
        $this->assertEquals(Style\Border::BORDER_DASHED, $style->getBorders()->getLeft()->getStyle());
        $this->assertEquals(Style\Color::BLACK, $style->getBorders()->getLeft()->getColor()->getArgb());
        $this->assertEquals(Style\Border::BORDER_DASHED, $style->getBorders()->getRight()->getStyle());
        $this->assertEquals(Style\Color::BLACK, $style->getBorders()->getRight()->getColor()->getArgb());
        $this->assertEquals(Style\Border::BORDER_DASHED, $style->getBorders()->getTop()->getStyle());
        $this->assertEquals(Style\Color::BLACK, $style->getBorders()->getTop()->getColor()->getArgb());
        $this->assertEquals(Style\Border::BORDER_DASHED, $style->getBorders()->getBottom()->getStyle());
        $this->assertEquals(Style\Color::BLACK, $style->getBorders()->getBottom()->getColor()->getArgb());
        $this->assertEquals(Style\Border::BORDER_DASHED, $style->getBorders()->getHorizontal()->getStyle());
        $this->assertEquals(Style\Color::BLACK, $style->getBorders()->getHorizontal()->getColor()->getArgb());
        $this->assertEquals(Style\Border::BORDER_DASHED, $style->getBorders()->getVertical()->getStyle());
        $this->assertEquals(Style\Color::BLACK, $style->getBorders()->getVertical()->getColor()->getArgb());
        $this->assertEquals(Style\Border::BORDER_DASHED, $style->getBorders()->getDiagonal()->getStyle());
        $this->assertEquals(Style\Color::BLACK, $style->getBorders()->getDiagonal()->getColor()->getArgb());
        $this->assertEquals(Style\Borders::DIAGONAL_BOTH, $style->getBorders()->getDiagonalDirection());
        $this->assertNotNull($style->getBorders()->getHashCode());
    }

    public function test_alignment()
    {
        $style = Style::builder()
            ->setAlignmentHorizontal(Style\Alignment::HORIZONTAL_LEFT)
            ->setAlignmentVertical(Style\Alignment::VERTICAL_CENTER)
            ->setAlignmentTextRotation(1)
            ->setAlignmentWrapText(true)
            ->setAlignmentShrinkToFit(true)
            ->setAlignmentIndent(1)
            ->setAlignmentReadOrder(Style\Alignment::READORDER_LTR)
            ->build();
        $this->assertEquals(Style\Alignment::HORIZONTAL_LEFT, $style->getAlignment()->getHorizontal());
        $this->assertEquals(Style\Alignment::VERTICAL_CENTER, $style->getAlignment()->getVertical());
        $this->assertEquals(1, $style->getAlignment()->getTextRotation());
        $this->assertTrue($style->getAlignment()->isWrapText());
        $this->assertTrue($style->getAlignment()->isShrinkToFit());
        $this->assertEquals(1, $style->getAlignment()->getIndent());
        $this->assertEquals(Style\Alignment::READORDER_LTR, $style->getAlignment()->getReadOrder());
        $this->assertNotNull($style->getAlignment()->getHashCode());
    }

    public function test_protection()
    {
        $style = Style::builder()
            ->setProtectionLocked(Style\Protection::PROTECTION_PROTECTED)
            ->setProtectionHidden(Style\Protection::PROTECTION_PROTECTED)
            ->build();
        $this->assertEquals(Style\Protection::PROTECTION_PROTECTED, $style->getProtection()->getLocked());
        $this->assertEquals(Style\Protection::PROTECTION_PROTECTED, $style->getProtection()->getHidden());
        $this->assertNotNull($style->getProtection()->getHashCode());
    }

    public function test_quotePrefix()
    {
        $style = Style::builder()
            ->setQuotePrefix(true)
            ->build();
        $this->assertTrue($style->getQuotePrefix());
    }
}