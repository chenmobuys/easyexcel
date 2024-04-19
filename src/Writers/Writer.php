<?php

namespace EasyExcel\Writers;

use EasyExcel\Interfaces\SheetInterface;
use EasyExcel\Interfaces\WriterInterface;
use EasyExcel\Interfaces\WriterRowInterface;
use EasyExcel\Metadata\Excel;
use EasyExcel\Metadata\Style;

abstract class Writer extends Excel implements WriterInterface
{
    /**
     * @var string
     */
    protected $filename;

    /**
     * @var array
     */
    protected $rowWriters = [];

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
        $this->openFromFile($filename);
    }

    /**
     * Determine whether the file is writeable.
     *
     * @param  string  $filename
     *
     * @return bool
     */
    public static function writeable(string $filename): bool
    {
        if (!is_dir($directory = dirname($filename))) {
            mkdir($directory);
        }
        return is_dir($directory) && is_writeable($directory);
    }

    /**
     * Open file.
     *
     * @param  string  $filename
     *
     * @return $this
     */
    public static function open(string $filename): WriterInterface
    {
        return new static($filename);
    }

    public function addRow(array $row, ?Style $style = null): WriterInterface
    {
        $this->getRowWriter()->writes([$row], $style);
        return $this;
    }

    public function addRows(array $rows, ?Style $style = null): WriterInterface
    {
        $this->getRowWriter()->writes($rows, $style);
        return $this;
    }

    protected function getRowWriter(): WriterRowInterface
    {
        $sheet = $this->getActiveSheet();
        if (!isset($this->rowWriters[$sheet->getIndex()])) {
            $this->rowWriters[$sheet->getIndex()] = $this->getRowWriterBySheet($sheet);
        }
        return $this->rowWriters[$sheet->getIndex()];
    }

    /**
     * Close and save file.
     *
     * @return void
     */
    public function close(): void
    {
        if (!$this->closed) {
            $this->closeWriter();
            $this->closed = true;
        }
    }

    /**
     * @param  string  $filename
     *
     * @return $this
     */
    abstract protected function openFromFile(string $filename): WriterInterface;

    /**
     * @param  SheetInterface  $sheet
     *
     * @return WriterRowInterface
     */
    abstract protected function getRowWriterBySheet(SheetInterface $sheet): WriterRowInterface;

    /**
     * Close writer.
     *
     * @return void
     */
    abstract protected function closeWriter(): void;

    /**
     * Destructor.
     */
    public function __destruct()
    {
        if (!$this->closed) {
            trigger_error("Use close method to save document.");
        }
    }
}
