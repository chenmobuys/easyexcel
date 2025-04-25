<?php

namespace EasyExcelTests\Metadata;

use EasyExcel\Metadata\Style;
use PHPUnit\Framework\TestCase;

class StyleTest extends TestCase
{
    public function test_index()
    {
        $style = new Style();
        $style->setIndex(-1);
        $this->assertEquals(-1, $style->getIndex());
    }

    public function test_font()
    {
        $style = new Style();
        $font = new Style\Font();
        $font->setName("foo");
        $style->setFont($font);
        $this->assertEquals("foo", $style->getFont()->getName());
    }

    public function test_fill()
    {
        $style = new Style();
        $fill = new Style\Fill();
        $fill->setType(Style\Fill::FILL_GRADIENT_LINEAR);
        $style->setFill($fill);
        $this->assertEquals(Style\Fill::FILL_GRADIENT_LINEAR, $style->getFill()->getType());
    }

    public function test_format()
    {
        $style = new Style();
        $format = new Style\Format();
        $format->setFormatCode(Style\Format::FORMAT_DATE_DDMMYYYY);
        $style->setFormat($format);
        $this->assertEquals(Style\Format::FORMAT_DATE_DDMMYYYY, $style->getFormat()->getFormatCode());
    }

    public function test_borders()
    {
        $style = new Style();
        $borders = new Style\Borders();
        $borders->setDiagonalDirection(Style\Borders::DIAGONAL_BOTH);
        $style->setBorders($borders);
        $this->assertEquals(Style\Borders::DIAGONAL_BOTH, $style->getBorders()->getDiagonalDirection());
    }

    public function test_alignment()
    {
        $style = new Style();
        $alignment = new Style\Alignment();
        $alignment->setHorizontal(Style\Alignment::HORIZONTAL_LEFT);
        $style->setAlignment($alignment);
        $this->assertEquals(Style\Alignment::HORIZONTAL_LEFT, $style->getAlignment()->getHorizontal());
    }

    public function test_protection()
    {
        $style = new Style();
        $protection = new Style\Protection();
        $protection->setLocked(Style\Protection::PROTECTION_PROTECTED);
        $style->setProtection($protection);
        $this->assertEquals(Style\Protection::PROTECTION_PROTECTED, $style->getProtection()->getLocked());
    }

    public function test_hashCode()
    {
        $style = new Style();
        $this->assertNotNull($style->getHashCode());
    }
}