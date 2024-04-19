<?php

namespace EasyExcel\Readers;

use EasyExcel\Interfaces\ReaderInterface;
use EasyExcel\Interfaces\ReaderRowInterface;
use EasyExcel\Interfaces\SheetInterface;
use EasyExcel\Metadata\Excel;

abstract class Reader extends Excel implements ReaderInterface
{
    /**
     * @var string
     */
    protected $filename;

    /**
     * @var array
     */
    protected $rowReaders = [];

    /**
     * Close flag.
     *
     * @var bool
     */
    protected $closed = false;

    protected function __construct(string $filename)
    {
        parent::__construct();
        $this->filename = $filename;
        $this->loadFromFile($filename);
    }

    /**
     * Load file.
     *
     * @param  string  $filename
     * @return $this
     */
    public static function load(string $filename): ReaderInterface
    {
        return new static($filename);
    }

    /**
     *
     * @param  int  $startRow
     * @param  int|null  $endRow
     * @return ReaderRowInterface
     */
    public function getRowIterator(int $startRow = 1, int $endRow = null): ReaderRowInterface
    {
        $sheet = $this->getActiveSheet();
        if (!isset($this->rowReaders[$sheet->getIndex()])) {
            $this->rowReaders[$sheet->getIndex()] = $this->getRowIteratorBySheet($sheet);
        }
        return $this->rowReaders[$sheet->getIndex()];
    }

    /**
     * Close file.
     *
     * @return void
     */
    public function close(): void
    {
        if (!$this->closed) {
            $this->closeReader();
            $this->closed = true;
        }
    }

    /**
     * Load from file.
     *
     * @param  string  $filename
     * @return $this
     */
    abstract protected function loadFromFile(string $filename): ReaderInterface;

    /**
     * Get row iterator by sheet.
     *
     * @param  \EasyExcel\Interfaces\SheetInterface  $sheet
     * @param  int  $startRow
     * @param  int|null  $endRow
     * @return \EasyExcel\Interfaces\ReaderRowInterface
     */
    abstract protected function getRowIteratorBySheet(SheetInterface $sheet, int $startRow = 1, int $endRow = null): ReaderRowInterface;

    /**
     * Close reader.
     *
     * @return void
     */
    abstract protected function closeReader(): void;

    /**
     * Clone.
     */
    public function __clone()
    {
        $this->rowReaders = [];
    }

    /**
     * Destructor.
     */
    public function __destruct()
    {
        $this->close();
    }
}
