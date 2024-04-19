<?php

namespace EasyExcelTests\Metadata;

use EasyExcel\Metadata\AutoFilter;
use PHPUnit\Framework\TestCase;

class AutoFilterTest extends TestCase
{
    public function test_getRange()
    {
        $autoFilter = new AutoFilter();
        $this->assertEmpty($autoFilter->getRange());
    }

    public function test_setRange()
    {
        $autoFilter = new AutoFilter();
        $autoFilter->setRange("A1:B1");
        $this->assertEquals("A1:B1", $autoFilter->getRange());
    }
}