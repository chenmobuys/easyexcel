<?php

namespace EasyExcelTests;

use EasyExcel\Exceptions\FileNotReadableException;
use EasyExcel\Exceptions\FileNotWriteableException;
use EasyExcel\Exceptions\NotSupportedExcelTypeException;
use EasyExcel\Exceptions\UnknownExcelTypeException;
use EasyExcel\Factory;
use EasyExcel\Interfaces\Reader;
use EasyExcel\Interfaces\Writer;
use PHPUnit\Framework\TestCase;

class FactoryTest extends TestCase
{
    /**
     * @throws \EasyExcel\Exceptions\NotSupportedExcelTypeException
     * @throws \EasyExcel\Exceptions\FileNotReadableException
     * @throws \EasyExcel\Exceptions\UnknownExcelTypeException
     */
    public function test_load(): void
    {
        $reader = Factory::load(__DIR__."/data/sample.csv");
        $this->assertInstanceOf(Reader::class, $reader);
        $reader->close();
    }

    /**
     * @throws \EasyExcel\Exceptions\NotSupportedExcelTypeException
     */
    public function test_createReader(): void
    {
        $reader = Factory::createReader(Factory::TYPE_XLSX);
        $this->assertTrue(is_a($reader, Reader::class, true));
    }

    public function test_createReaderError1(): void
    {
        $this->expectException(NotSupportedExcelTypeException::class);
        try {
            Factory::createReader("Foo");
        } catch (NotSupportedExcelTypeException $e) {
            $this->assertEquals("Foo", $e->getExcelType());
            $this->assertEquals(Reader::class, $e->getClassName());
            throw $e;
        }
    }

    /**
     * @throws \EasyExcel\Exceptions\NotSupportedExcelTypeException
     * @throws \EasyExcel\Exceptions\FileNotReadableException
     * @throws \EasyExcel\Exceptions\UnknownExcelTypeException
     */
    public function test_createReaderForFile(): void
    {
        $reader = Factory::createReaderForFile(__DIR__."/data/sample.foo");
        $this->assertTrue(is_a($reader, Reader::class, true));
    }

    /**
     * @throws \EasyExcel\Exceptions\NotSupportedExcelTypeException
     * @throws \EasyExcel\Exceptions\UnknownExcelTypeException
     */
    public function test_createReaderForFileError1(): void
    {
        $this->expectException(FileNotReadableException::class);
        try {
            Factory::createReaderForFile("/foo");
        } catch (FileNotReadableException $e) {
            $this->assertEquals("/foo", $e->getFilename());
            throw $e;
        }
    }

    /**
     * @throws \EasyExcel\Exceptions\NotSupportedExcelTypeException
     * @throws \EasyExcel\Exceptions\FileNotReadableException
     */
    public function test_createReaderForFileError2(): void
    {
        $this->expectException(UnknownExcelTypeException::class);
        try {
            Factory::createReaderForFile(__DIR__."/data/sample.zip");
        } catch (UnknownExcelTypeException $e) {
            throw $e;
        }
    }

    /**
     * @throws \EasyExcel\Exceptions\FileNotWriteableException
     * @throws \EasyExcel\Exceptions\NotSupportedExcelTypeException
     */
    public function test_open(): void
    {
        $temp = sys_get_temp_dir()."/sample.csv";
        $writer = Factory::open($temp);
        $writer->close();
        $this->assertInstanceOf(Writer::class, $writer);
        @unlink($temp);
    }

    public function test_createWriterError(): void
    {
        $this->expectException(NotSupportedExcelTypeException::class);
        try {
            Factory::createWriter("Foo");
        } catch (NotSupportedExcelTypeException $e) {
            $this->assertEquals("Foo", $e->getExcelType());
            throw $e;
        }
    }

    /**
     * @throws \EasyExcel\Exceptions\NotSupportedExcelTypeException
     */
    public function test_createWriterForFileError(): void
    {
        $this->expectException(FileNotWriteableException::class);
        try {
            Factory::createWriterForFile("/foo");
        } catch (FileNotWriteableException $e) {
            $this->assertEquals("/foo", $e->getFilename());
            throw $e;
        }
    }

    public function test_registerExcelTypeResolver(): void
    {
        $excelTypeResolver = function (string $filename, string $extension) {
            if ($extension == 'foo') {
                return 'bar';
            }
            return null;
        };
        Factory::registerExcelTypeResolver($excelTypeResolver);
        $guessExcelTypeFromFilename = Factory::guessExcelTypeFromFilename("bar.foo");
        $this->assertEquals("bar", $guessExcelTypeFromFilename);
        Factory::registerExcelTypeResolver(null);
    }

    public function test_guessExcelTypeFromFilename(): void
    {
        $this->assertEquals(Factory::TYPE_XLSX, Factory::guessExcelTypeFromFilename("foo.xlsx"));
        $this->assertEquals(Factory::TYPE_XLS, Factory::guessExcelTypeFromFilename("foo.xls"));
        $this->assertEquals(Factory::TYPE_ODS, Factory::guessExcelTypeFromFilename("foo.ods"));
        $this->assertEquals(Factory::TYPE_CSV, Factory::guessExcelTypeFromFilename("foo.csv"));
    }
}
