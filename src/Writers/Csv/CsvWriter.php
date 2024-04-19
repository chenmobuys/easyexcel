<?php

namespace EasyExcel\Writers\Csv;

use EasyExcel\Helpers\Filesystem;
use EasyExcel\Interfaces\SheetInterface;
use EasyExcel\Interfaces\WriterInterface;
use EasyExcel\Interfaces\WriterRowInterface;
use EasyExcel\Writers\Writer;
use SplFileObject;

class CsvWriter extends Writer
{
    /**
     * @var \SplFileObject
     */
    protected $handler;

    /**
     * @var string
     */
    protected $tempFilename;

    /**
     * @param  SheetInterface  $sheet
     * @return WriterRowInterface
     */
    public function getRowWriterBySheet(SheetInterface $sheet): WriterRowInterface
    {
        return new CsvWriterRow($this->handler, $sheet);
    }

    /**
     * @param  string  $filename
     * @return $this
     */
    protected function openFromFile(string $filename): WriterInterface
    {
        $this->tempFilename = Filesystem::getTempName();
        $this->handler = new SplFileObject($this->tempFilename, 'w+b');
        $this->getSheetByName('Sheet', true);

        return $this;
    }

    /**
     * Close writer.
     *
     * @return void
     */
    protected function closeWriter(): void
    {
        @copy($this->tempFilename, $this->filename);

        $this->handler = null;
        @unlink($this->tempFilename);
    }
}
