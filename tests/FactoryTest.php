<?php

namespace EasyExcelTests;

use EasyExcel\Exceptions\FileNotReadableException;
use EasyExcel\Exceptions\FileNotWriteableException;
use EasyExcel\Exceptions\NotSupportedExcelTypeException;
use EasyExcel\Exceptions\RegisterExcelTypeFailedException;
use EasyExcel\Exceptions\UnknownExcelTypeException;
use EasyExcel\Factory;
use EasyExcel\Interfaces\ReaderInterface;
use EasyExcel\Interfaces\ReaderRowInterface;
use EasyExcel\Interfaces\SheetInterface;
use EasyExcel\Interfaces\WriterInterface;
use EasyExcel\Interfaces\WriterRowInterface;
use EasyExcel\Metadata\Row;
use EasyExcel\Metadata\Style;
use EasyExcel\Readers\Reader;
use EasyExcel\Readers\ReaderRow;
use EasyExcel\Writers\Writer;
use EasyExcel\Writers\WriterRow;
use PHPUnit\Framework\TestCase;
use stdClass;

class FactoryTest extends TestCase
{
    public function test_load(): void
    {
        $reader = Factory::load("tests/data/sample.csv");
        $this->assertInstanceOf(ReaderInterface::class, $reader);
        $reader->close();
    }

    public function test_createReader(): void
    {
        $reader = Factory::createReader(Factory::TYPE_XLSX);
        $this->assertTrue(is_a($reader, ReaderInterface::class, true));
    }

    public function test_createReaderError1(): void
    {
        $this->expectException(NotSupportedExcelTypeException::class);
        try {
            Factory::createReader("Foo");
        } catch (NotSupportedExcelTypeException $e) {
            $this->assertEquals("Foo", $e->getExcelType());
            $this->assertEquals(ReaderInterface::class, $e->getClassName());
            throw $e;
        }
    }

    public function test_createReaderForFile(): void
    {
        $reader = Factory::createReaderForFile("tests/data/sample.foo");
        $this->assertTrue(is_a($reader, ReaderInterface::class, true));
    }

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

    public function test_createReaderForFileError2(): void
    {
        $this->expectException(UnknownExcelTypeException::class);
        try {
            Factory::createReaderForFile("tests/data/sample.zip");
        } catch (UnknownExcelTypeException $e) {
            $this->assertEquals("tests/data/sample.zip", $e->getFilename());
            throw $e;
        }
    }

    public function test_registerReader(): void
    {
        Factory::registerReader("Foo", FooReader::class);
        $reader = Factory::createReader("Foo");
        $this->assertTrue(is_a($reader, ReaderInterface::class, true));
    }

    public function test_registerReaderError1(): void
    {
        $this->expectException(RegisterExcelTypeFailedException::class);
        try {
            Factory::registerReader(Factory::TYPE_XLSX, stdClass::class);
        } catch (RegisterExcelTypeFailedException $e) {
            $this->assertEquals(ReaderInterface::class, $e->getClassName());
            throw $e;
        }
    }

    public function test_open(): void
    {
        $temp = sys_get_temp_dir()."/sample.csv";
        $writer = Factory::open($temp);
        $writer->close();
        $this->assertInstanceOf(WriterInterface::class, $writer);
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

    public function test_registerWriter(): void
    {
        Factory::registerWriter("Foo", FooWriter::class);
        $writer = Factory::createWriter("Foo");
        $this->assertTrue(is_a($writer, WriterInterface::class, true));
    }

    public function test_registerWriterError(): void
    {
        $this->expectException(RegisterExcelTypeFailedException::class);
        try {
            Factory::registerWriter("Foo", stdClass::class);
        } catch (RegisterExcelTypeFailedException $e) {
            $this->assertEquals(WriterInterface::class, $e->getClassName());
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

class FooReader extends Reader
{
    public static function readable(string $filename): bool
    {
        return true;
    }

    protected function loadFromFile(string $filename): ReaderInterface
    {
        return $this;
    }

    protected function getRowIteratorBySheet(
        SheetInterface $sheet,
        int $startRow = 1,
        int $endRow = null
    ): ReaderRowInterface {
        return new class($this->getActiveSheet()) extends ReaderRow {
            public function close(): void
            {
            }
        };
    }

    protected function closeReader(): void
    {
    }
}

class FooWriter extends Writer
{
    protected function openFromFile(string $filename): WriterInterface
    {
        return $this;
    }

    protected function getRowWriterBySheet(SheetInterface $sheet): WriterRowInterface
    {
        return new class($sheet) extends WriterRow {
            protected function writeRow(Row $row, ?Style $style = null): void
            {
            }

            protected function writeArray(array $row, ?Style $style = null): void
            {
            }

            public function close(): void
            {
            }
        };
    }

    protected function closeWriter(): void
    {
    }
}