<?php

namespace EasyExcel;

use Closure;
use EasyExcel\Exceptions\FileNotReadableException;
use EasyExcel\Exceptions\FileNotWriteableException;
use EasyExcel\Exceptions\NotSupportedExcelTypeException;
use EasyExcel\Exceptions\RegisterExcelTypeFailedException;
use EasyExcel\Exceptions\UnknownExcelTypeException;
use EasyExcel\Interfaces\ReaderInterface;
use EasyExcel\Interfaces\WriterInterface;

final class Factory
{
    public const TYPE_CSV = 'Csv';
    public const TYPE_ODS = 'Ods';
    public const TYPE_XLS = 'Xls';
    public const TYPE_XLSX = 'Xlsx';
    // public const TYPE_XML = 'Xml';
    // public const TYPE_SLK = 'Slk';
    // public const TYPE_HTML = 'Html';
    // public const TYPE_GNUMERIC = 'Gnumeric';

    /**
     * Default support readers.
     *
     * @var array
     */
    private static $readers = [
        self::TYPE_CSV  => Readers\Csv\CsvReader::class,
        self::TYPE_ODS  => Readers\Ods\OdsReader::class,
        self::TYPE_XLS  => Readers\Xls\XlsReader::class,
        self::TYPE_XLSX => Readers\Xlsx\XlsxReader::class,
    ];

    /**
     * Default support writers.
     *
     * @var array
     */
    private static $writers = [
        self::TYPE_CSV  => Writers\Csv\CsvWriter::class,
        self::TYPE_XLSX => Writers\Xlsx\XlsxWriter::class,
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
     * @param  string  $filename
     *
     * @return ReaderInterface
     */
    public static function load(string $filename): ReaderInterface
    {
        $reader = self::createReaderForFile($filename);

        return $reader::load($filename);
    }

    /**
     * Create reader For File.
     *
     * @param  string  $filename
     *
     * @return mixed
     */
    public static function createReaderForFile(string $filename): string
    {
        if (!is_readable($filename)) {
            throw new FileNotReadableException($filename);
        }

        $guessedExcelType = self::guessExcelTypeFromFilename($filename);

        if ($guessedExcelType) {
            $reader = self::createReader($guessedExcelType);
            if ($reader::readable($filename)) {
                return $reader;
            }
        }

        foreach (self::$readers as $excelType => $className) {
            if ($excelType !== $guessedExcelType) {
                $reader = self::createReader($excelType);
                if ($reader::readable($filename)) {
                    return $reader;
                }
            }
        }

        throw new UnknownExcelTypeException($filename);
    }

    /**
     * Detect excel type for filename.
     *
     * @param  string  $filename
     *
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
            case 'csv':
            case 'tsv':
                return self::TYPE_CSV;
            // case 'slk':
            //     return self::TYPE_SLK;
            // case 'xml': // Excel 2003 SpreadSheetML
            //     return self::TYPE_XML;
            // case 'gnumeric':
            //     return self::TYPE_GNUMERIC;
            // case 'htm':
            // case 'html':
            //     return self::TYPE_HTML;
            default:
                return null;
        }
    }

    /**
     * Create reader for Excel type.
     *
     * @param  string  $excelType
     *
     * @return mixed
     */
    public static function createReader(string $excelType): string
    {
        if (!isset(self::$readers[$excelType])) {
            throw new NotSupportedExcelTypeException(ReaderInterface::class, $excelType);
        }

        return self::$readers[$excelType];
    }

    /**
     * Register a reader with Excel type and class name.
     *
     * @param  string  $excelType
     * @param  string  $className
     *
     * @return void
     */
    public static function registerReader(string $excelType, string $className): void
    {
        if (!is_a($className, ReaderInterface::class, true)) {
            throw new RegisterExcelTypeFailedException(ReaderInterface::class);
        }

        self::$readers[$excelType] = $className;
    }

    /**
     * Open file.
     *
     * @param  string  $filename
     *
     * @return WriterInterface
     */
    public static function open(string $filename): WriterInterface
    {
        $writer = self::createWriterForFile($filename);

        return $writer::open($filename);
    }

    /**
     * Create writer For File.
     *
     * @param  string  $filename
     *
     * @return mixed
     */
    public static function createWriterForFile(string $filename): string
    {
        $guessedExcelType = self::guessExcelTypeFromFilename($filename);

        if ($guessedExcelType) {
            $writer = self::createWriter($guessedExcelType);
            if ($writer::writeable($filename)) {
                return $writer;
            }
        }

        throw new FileNotWriteableException($filename);
    }

    /**
     * Create writer for Excel type.
     *
     * @param  string  $excelType
     *
     * @return mixed
     */
    public static function createWriter(string $excelType): string
    {
        if (!isset(self::$writers[$excelType])) {
            throw new NotSupportedExcelTypeException(WriterInterface::class, $excelType);
        }

        return self::$writers[$excelType];
    }

    /**
     * Register a writer with Excel type and class name.
     *
     * @param  string  $excelType
     * @param  string  $className
     *
     * @return void
     */
    public static function registerWriter(string $excelType, string $className): void
    {
        if (!is_a($className, WriterInterface::class, true)) {
            throw new RegisterExcelTypeFailedException(WriterInterface::class);
        }

        self::$writers[$excelType] = $className;
    }

    /**
     * Register excel type resolver.
     *
     * @param  \Closure|null  $excelTypeResolver
     *
     * @return void
     */
    public static function registerExcelTypeResolver(?Closure $excelTypeResolver): void
    {
        self::$excelTypeResolver = $excelTypeResolver;
    }
}
