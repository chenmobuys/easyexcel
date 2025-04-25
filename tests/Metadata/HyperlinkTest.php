<?php

namespace EasyExcelTests\Metadata;

use EasyExcel\Metadata\Hyperlink;
use PHPUnit\Framework\TestCase;

class HyperlinkTest extends TestCase
{
    public function test_url()
    {
        $hyperlink = new Hyperlink("https://foo.com");
        $this->assertEquals("https://foo.com", $hyperlink->getUrl());
        $hyperlink->setUrl("https://bar.com");
        $this->assertEquals("https://bar.com", $hyperlink->getUrl());
    }

    public function test_tooltip()
    {
        $hyperlink = new Hyperlink("", "foo");
        $this->assertEquals("foo", $hyperlink->getTooltip());
        $hyperlink->setTooltip("bar");
        $this->assertEquals("bar", $hyperlink->getTooltip());
    }

    public function test_internal()
    {
        $hyperlink = new Hyperlink("https://foo.com");
        $this->assertFalse($hyperlink->isInternal());
        $hyperlink->setUrl("sheet://Sheet1");
        $this->assertTrue($hyperlink->isInternal());
    }
}