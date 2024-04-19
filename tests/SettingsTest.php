<?php

namespace EasyExcelTests;

use EasyExcel\Settings;
use PHPUnit\Framework\TestCase;

class SettingsTest extends TestCase
{
    public function test_getLibXmlLoaderOptions(): void
    {
        $this->assertEquals(12, Settings::getLibXmlLoaderOptions());
        Settings::setLibXmlLoaderOptions(LIBXML_DTDLOAD);
        $this->assertEquals(4, Settings::getLibXmlLoaderOptions());
        Settings::setLibXmlLoaderOptions(null);
        $this->assertEquals(12, Settings::getLibXmlLoaderOptions());
    }

    public function test_setTempDir(): void
    {
        Settings::setTempDir(sys_get_temp_dir());
        $this->assertEquals(sys_get_temp_dir(), Settings::getTempDir());
    }

}