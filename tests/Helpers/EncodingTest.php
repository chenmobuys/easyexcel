<?php

namespace EasyExcelTests\Helpers;

use EasyExcel\Exceptions\NotSupportConverterException;
use EasyExcel\Helpers\Encoding;
use PHPUnit\Framework\TestCase;

class EncodingTest extends TestCase
{
    /**
     * @throws \EasyExcel\Exceptions\NotSupportConverterException
     */
    public function test_setConvert(): void
    {
        Encoding::setConverter(Encoding::CONVERTER_CLOSURE);
        Encoding::setClosureConverter(function ($string, $from_encoding, $to_encoding) {
            return "bar";
        });
        $this->assertEquals(Encoding::CONVERTER_CLOSURE, Encoding::getConverter());
        $this->assertNotNull(Encoding::getClosureConverter());
        $this->assertEquals("bar", Encoding::convertEncoding("foo", "UTF-8", "GBK"));
        Encoding::setConverter(Encoding::CONVERTER_ICONV);
        $this->assertEquals("bar", Encoding::convertEncoding("bar", "UTF-8", "GBK"));
        Encoding::setConverter(Encoding::CONVERTER_MB);
    }

    public function test_setConvertError(): void
    {
        $this->expectException(NotSupportConverterException::class);
        try {
            Encoding::setConverter("Foo");
        } catch (NotSupportConverterException $e) {
            $this->assertEquals("Foo", $e->getConverter());
            throw $e;
        }
    }
}