<?php

namespace EasyExcelTests\Helpers;

use EasyExcel\Helpers\CodePage;
use PHPUnit\Framework\TestCase;

class CodePageTest extends TestCase
{
    public function test_numberToName(): void
    {
        $this->assertEquals("UTF-16LE", CodePage::numberToName(1200));
    }
}