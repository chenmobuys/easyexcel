<?php

namespace EasyExcel\Writers\Csv;

use EasyExcel\Helpers\Filesystem;
use EasyExcel\Interfaces\WriterExcel as WriterExcelInterface;
use EasyExcel\Interfaces\WriterRow as WriterRowInterface;
use EasyExcel\Interfaces\WriterSheet;
use EasyExcel\Writers\BaseWriter;
use EasyExcel\Writers\WriterExcel;
use SplFileObject;

class Writer extends BaseWriter
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
     * Get row writer.
     *
     * @param \EasyExcel\Interfaces\WriterSheet $sheet
     * @return \EasyExcel\Interfaces\WriterRow
     */
    public function getRowWriter(WriterSheet $sheet): WriterRowInterface
    {
        return new WriterRow($this->handler, $this, $sheet);
    }

    /**
     * Open from file.
     *
     * @param string $filename
     * @return \EasyExcel\Interfaces\WriterExcel
     * @throws \EasyExcel\Exceptions\SheetNameExistsException
     */
    protected function openFromFile(string $filename): WriterExcelInterface
    {
        $excel = new WriterExcel($this);
        $this->tempFilename = Filesystem::getTempName();
        $this->handler = new SplFileObject($this->tempFilename, 'w+b');
        $this->handler->setCsvControl(',', '"', '\\');
        $excel->addSheet('Sheet');

        return $excel;
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