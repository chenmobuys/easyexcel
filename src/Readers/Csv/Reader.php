<?php

namespace EasyExcel\Readers\Csv;

use EasyExcel\Helpers\Encoding;
use EasyExcel\Interfaces\ReaderExcel as ReaderExcelInterface;
use EasyExcel\Interfaces\ReaderRow as ReaderRowInterface;
use EasyExcel\Interfaces\ReaderSheet;
use EasyExcel\Readers\BaseReader;
use EasyExcel\Readers\ReaderExcel;
use SplFileObject;
use SplTempFileObject;

class Reader extends BaseReader
{
    public const DEFAULT_ENCODING = 'CP1252';
    public const UTF8_BOM = "\xef\xbb\xbf";
    public const UTF16BE_BOM = "\xfe\xff";
    public const UTF16BE_LF = "\x00\x0a";
    public const UTF16LE_BOM = "\xff\xfe";
    public const UTF16LE_LF = "\x0a\x00";
    public const UTF32BE_BOM = "\x00\x00\xfe\xff";
    public const UTF32BE_LF = "\x00\x00\x00\x0a";
    public const UTF32LE_BOM = "\xff\xfe\x00\x00";
    public const UTF32LE_LF = "\x0a\x00\x00\x00";

    /**
     * File handler.
     *
     * @var \SplFileObject
     */
    protected $handler;

    /**
     * File encoding.
     *
     * @var string
     */
    protected $encoding;

    /**
     * Separator character.
     *
     * @var string
     */
    protected $separator = ',';

    /**
     * Enclosure character.
     *
     * @var string
     */
    protected $enclosure = '"';

    /**
     * Escape character.
     *
     * @var string
     */
    protected $escape = '\\';

    /**
     * SplFileObject flags.
     *
     * @var int
     */
    protected $flags = SplFileObject::READ_CSV | SplFileObject::READ_AHEAD | SplFileObject::SKIP_EMPTY;

    /**
     * Get file encoding.
     *
     * @return string
     */
    public function getEncoding(): string
    {
        return $this->encoding;
    }

    /**
     * Set file encoding.
     *
     * @param string $encoding
     * @return $this
     */
    public function setEncoding(string $encoding): self
    {
        $this->encoding = $encoding;

        return $this;
    }

    /**
     * Get SplFileObject flags.
     *
     * @return int
     */
    public function getFlags(): int
    {
        return $this->flags;
    }

    /**
     * Set SplFileObject flags.
     *
     * @param int $flags
     * @return $this
     */
    public function setFlags(int $flags): self
    {
        $this->flags = $flags;

        return $this;
    }

    /**
     * Set SplFileObject csv control.
     *
     * @param string $separator
     * @param string $enclosure
     * @param string $escape
     * @return $this
     */
    public function setCsvControl(string $separator, string $enclosure, string $escape): self
    {
        $this->separator = $separator;
        $this->enclosure = $enclosure;
        $this->escape = $escape;

        return $this;
    }

    /**
     * Get SplFileObject csv control.
     *
     * @return array
     */
    public function getCsvControl(): array
    {
        return [$this->separator, $this->enclosure, $this->escape];
    }

    /**
     * Determine whether the file is readable.
     *
     * @param string $filename
     * @return bool
     */
    public function readable(string $filename): bool
    {
        $extension = strtolower(pathinfo($filename, PATHINFO_EXTENSION));
        if (in_array($extension, ['csv', 'tsv'])) {
            return true;
        }

        $mimeType = @mime_content_type($filename);
        $supportTypes = ['application/csv', 'text/csv', 'text/plain', 'inode/x-empty', 'application/x-empty'];

        return in_array($mimeType, $supportTypes);
    }

    /**
     * Get row iterator.
     *
     * @param \EasyExcel\Interfaces\ReaderSheet $sheet
     * @param int $startRow
     * @param int|null $endRow
     * @return \EasyExcel\Interfaces\ReaderRow
     */
    public function getRowIterator(ReaderSheet $sheet, int $startRow = 1, ?int $endRow = null): ReaderRowInterface
    {
        return new ReaderRow($this->handler, $this->encoding, $sheet, $startRow, $endRow);
    }

    /**
     * Load from file.
     *
     * @param string $filename
     * @return \EasyExcel\Interfaces\ReaderExcel
     * @throws \EasyExcel\Exceptions\SheetNameExistsException
     */
    protected function loadFromFile(string $filename): ReaderExcelInterface
    {
        if (!$this->encoding) {
            $this->encoding = $this->guessEncoding($filename);
        }

        $this->handler = new SplTempFileObject();
        $this->handler->setFlags($this->flags);
        $this->handler->setCsvControl(...$this->getCsvControl());

        $contents = file_get_contents($filename);
        $hasUTF8Bom = substr($contents, 0, 3) == self::UTF8_BOM;
        $contents = substr($contents, $hasUTF8Bom ? 3 : 0);
        $utf8Contents = Encoding::convertEncoding($contents, $this->encoding);

        $this->handler->fwrite($utf8Contents);
        $this->handler->rewind();

        $totalRows = 0;
        $totalColumns = 0;
        while ($row = $this->handler->current()) {
            if (is_array($row)) {
                $totalRows++;
                $totalColumns = max($totalColumns, count($row));
            }
            $this->handler->next();
        }

        $excel = new ReaderExcel($this);
        $sheetName = pathinfo($filename, PATHINFO_FILENAME);
        $sheet = $excel->addSheet($sheetName);
        $sheet->setTotalRows($totalRows);
        $sheet->setTotalColumns($totalColumns);

        return $excel;
    }

    /**
     * Guess file encoding.
     *
     * @param string $filename
     * @return string
     */
    protected function guessEncoding(string $filename): string
    {
        $first4 = file_get_contents($filename, false, null, 0, 4);
        $bomEncodingList = [
            self::UTF8_BOM => 'UTF-8',
            self::UTF16BE_BOM => 'UTF-16BE',
            self::UTF32BE_BOM => 'UTF-32BE',
            self::UTF32LE_BOM => 'UTF-32LE',
            self::UTF16LE_BOM => 'UTF-16LE',
        ];
        foreach ($bomEncodingList as $characters => $encoding) {
            if (substr($first4, 0, strlen($characters)) === $characters) {
                return $encoding;
            }
        }

        $contents = file_get_contents($filename);
        $noBomEncodingList = [
            self::UTF32BE_LF => 'UTF-32BE',
            self::UTF32LE_LF => 'UTF-32LE',
            self::UTF16BE_LF => 'UTF-16BE',
            self::UTF16LE_LF => 'UTF-16LE',
        ];
        foreach ($noBomEncodingList as $characters => $encoding) {
            $pos = strpos($contents, $characters);
            if ($pos !== false && ($pos % strlen($characters) === 0)) {
                return $encoding;
            }
        }

        if (preg_match('//u', $contents) === 1) {
            return 'UTF-8';
        }

        return self::DEFAULT_ENCODING;
    }

    /**
     * Close reader.
     *
     * @return void
     */
    protected function closeReader(): void
    {
        $this->handler = null;
    }
}