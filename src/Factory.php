<?php

namespace EasyExcel;

use Closure;
use EasyExcel\Exceptions\FileNotReadableException;
use EasyExcel\Exceptions\FileNotWriteableException;
use EasyExcel\Exceptions\NotSupportedExcelTypeException;
use EasyExcel\Exceptions\RegisterExcelTypeFailedException;
use EasyExcel\Exceptions\UnknownExcelTypeException;
use EasyExcel\Interfaces\Reader;
use EasyExcel\Interfaces\Writer;

final class Factory
{
    public const TYPE_CSV = 'Csv';
    public const TYPE_ODS = 'Ods';
    public const TYPE_XLS = 'Xls';
    public const TYPE_XLSX = 'Xlsx';
    public const TYPE_XML = 'Xml';
    public const TYPE_SLK = 'Slk';
    public const TYPE_HTML = 'Html';
    public const TYPE_GNUMERIC = 'Gnumeric';

    /**
     * Default support readers.
     *
     * @var array
     */
    private static $readers = [
        self::TYPE_CSV => Readers\Csv\Reader::class,
        self::TYPE_ODS => Readers\Ods\Reader::class,
        self::TYPE_XLS => Readers\Xls\Reader::class,
        self::TYPE_XLSX => Readers\Xlsx\Reader::class,
    ];

    /**
     * Default support writers.
     *
     * @var array
     */
    private static $writers = [
        self::TYPE_CSV => Writers\Csv\Writer::class,
        self::TYPE_XLSX => Writers\Xlsx\Writer::class,
    ];

    /**
     * Excel type resolver.
     *
     * @var \Closure|null
     */
    private static $excelTypeResolver;

    /**
     * Load file.
     *
     * @param string $filename
     * @return \EasyExcel\Interfaces\Reader
     * @throws \EasyExcel\Exceptions\FileNotReadableException
     * @throws \EasyExcel\Exceptions\NotSupportedExcelTypeException
     * @throws \EasyExcel\Exceptions\UnknownExcelTypeException
     */
    public static function load(string $filename): Reader
    {
        $reader = self::createReaderForFile($filename);

        return $reader->load($filename);
    }

    /**
     * Create reader For File.
     *
     * @param string $filename
     * @return \EasyExcel\Interfaces\Reader
     * @throws \EasyExcel\Exceptions\FileNotReadableException
     * @throws \EasyExcel\Exceptions\UnknownExcelTypeException
     * @throws \EasyExcel\Exceptions\NotSupportedExcelTypeException
     */
    public static function createReaderForFile(string $filename): Reader
    {
        if (!is_readable($filename)) {
            throw new FileNotReadableException($filename);
        }

        $guessedExcelType = self::guessExcelTypeFromFilename($filename);

        if ($guessedExcelType) {
            $reader = self::createReader($guessedExcelType);
            if ($reader->readable($filename)) {
                return $reader;
            }
        }

        foreach (self::$readers as $excelType => $className) {
            if ($excelType !== $guessedExcelType) {
                $reader = self::createReader($excelType);
                if ($reader->readable($filename)) {
                    return $reader;
                }
            }
        }

        throw new UnknownExcelTypeException($filename);
    }

    /**
     * Create reader for Excel type.
     *
     * @param string $excelType
     * @return \EasyExcel\Interfaces\Reader
     * @throws \EasyExcel\Exceptions\NotSupportedExcelTypeException
     * @see \EasyExcel\Factory
     * @see \EasyExcel\Factory::registerReader()
     */
    public static function createReader(string $excelType): Reader
    {
        if (!isset(self::$readers[$excelType])) {
            throw new NotSupportedExcelTypeException(Reader::class, $excelType);
        }

        $className = self::$readers[$excelType];

        return new $className();
    }

    /**
     * Register a reader with Excel type and class name.
     *
     * @param string $excelType
     * @param string $className
     * @return void
     * @throws \EasyExcel\Exceptions\RegisterExcelTypeFailedException
     */
    public static function registerReader(string $excelType, string $className): void
    {
        if (!is_a($className, Reader::class, true)) {
            throw new RegisterExcelTypeFailedException(Reader::class);
        }

        self::$readers[$excelType] = $className;
    }

    /**
     * Open file.
     *
     * @param string $filename
     * @return \EasyExcel\Interfaces\Writer
     * @throws \EasyExcel\Exceptions\FileNotWriteableException
     * @throws \EasyExcel\Exceptions\NotSupportedExcelTypeException
     */
    public static function open(string $filename): Writer
    {
        $writer = self::createWriterForFile($filename);

        return $writer->open($filename);
    }

    /**
     * Create writer For File.
     *
     * @param string $filename
     * @return \EasyExcel\Interfaces\Writer
     * @throws \EasyExcel\Exceptions\FileNotWriteableException
     * @throws \EasyExcel\Exceptions\NotSupportedExcelTypeException
     */
    public static function createWriterForFile(string $filename): Writer
    {
        $guessedExcelType = self::guessExcelTypeFromFilename($filename);

        if ($guessedExcelType) {
            $writer = self::createWriter($guessedExcelType);
            $dirname = dirname($filename);
            if (!is_dir($dirname)) {
                mkdir($dirname);
            }
            if ($writer->writeable($filename)) {
                return $writer;
            }
        }

        throw new FileNotWriteableException($filename);
    }

    /**
     * Create writer for Excel type.
     *
     * @param string $excelType
     * @return \EasyExcel\Interfaces\Writer
     * @throws \EasyExcel\Exceptions\NotSupportedExcelTypeException
     * @see \EasyExcel\Factory
     * @see \EasyExcel\Factory::registerWriter()
     */
    public static function createWriter(string $excelType): Writer
    {
        if (!isset(self::$writers[$excelType])) {
            throw new NotSupportedExcelTypeException(Writer::class, $excelType);
        }

        $className = self::$writers[$excelType];

        return new $className();
    }

    /**
     * Register a writer with Excel type and class name.
     *
     * @param string $excelType
     * @param string $className
     * @return void
     * @throws \EasyExcel\Exceptions\RegisterExcelTypeFailedException
     */
    public static function registerWriter(string $excelType, string $className): void
    {
        if (!is_a($className, Writer::class, true)) {
            throw new RegisterExcelTypeFailedException(Writer::class);
        }

        self::$writers[$excelType] = $className;
    }

    /**
     * Register excel type resolver.
     *
     * @param \Closure|null $excelTypeResolver
     * @return void
     */
    public static function registerExcelTypeResolver(?Closure $excelTypeResolver): void
    {
        self::$excelTypeResolver = $excelTypeResolver;
    }

    /**
     * Unregister excel type resolver.
     *
     * @return void
     */
    public static function unregisterExcelTypeResolver(): void
    {
        self::$excelTypeResolver = null;
    }

    /**
     * Detect excel type for filename.
     *
     * @param string $filename
     * @return ?string
     */
    public static function guessExcelTypeFromFilename(string $filename): ?string
    {
        $extension = strtolower(pathinfo($filename, PATHINFO_EXTENSION));

        if (self::$excelTypeResolver instanceof Closure) {
            return call_user_func(self::$excelTypeResolver, $filename, $extension);
        }

        switch ($extension) {
            case 'xlsx': // Excel (OfficeOpenXML) Excel
            case 'xlsm': // Excel (OfficeOpenXML) Macro Excel (macros will be discarded)
            case 'xltx': // Excel (OfficeOpenXML) Template
            case 'xltm': // Excel (OfficeOpenXML) Macro Template (macros will be discarded)
                return self::TYPE_XLSX;
            case 'xls': // Excel (BIFF) Excel
            case 'xlt': // Excel (BIFF) Template
                return self::TYPE_XLS;
            case 'ods': // Open/Libre Office Calc
            case 'ots': // Open/Libre Office Calc Template
                return self::TYPE_ODS;
            case 'slk':
                return self::TYPE_SLK;
            case 'xml': // Excel 2003 SpreadSheetML
                return self::TYPE_XML;
            case 'gnumeric':
                return self::TYPE_GNUMERIC;
            case 'htm':
            case 'html':
                return self::TYPE_HTML;
            case 'csv':
            case 'tsv':
                return self::TYPE_CSV;
            default:
                return null;
        }
    }
}
