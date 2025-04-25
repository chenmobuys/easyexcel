<?php

namespace EasyExcel\Writers\Xlsx;

use EasyExcel\Interfaces\WriterExcel as WriterExcelInterface;
use EasyExcel\Interfaces\WriterRow as WriterRowInterface;
use EasyExcel\Interfaces\WriterSheet;
use EasyExcel\Metadata\Style;
use EasyExcel\Writers\BaseWriter;
use EasyExcel\Writers\WriterExcel;
use EasyExcel\Writers\Xlsx\Parts\ContentTypes;
use EasyExcel\Writers\Xlsx\Parts\DocPropsApp;
use EasyExcel\Writers\Xlsx\Parts\DocPropsCore;
use EasyExcel\Writers\Xlsx\Parts\Rels;
use EasyExcel\Writers\Xlsx\Parts\SharedStrings;
use EasyExcel\Writers\Xlsx\Parts\Sheet;
use EasyExcel\Writers\Xlsx\Parts\Styles;
use EasyExcel\Writers\Xlsx\Parts\Workbook;
use EasyExcel\Writers\Xlsx\Parts\WorkbookRels;
use ZipArchive;

class Writer extends BaseWriter
{
    /**
     * @var \EasyExcel\Writers\Xlsx\Parts\Sheet[]
     */
    protected $sheetHandlers = [];

    /**
     * Get row writer.
     *
     * @param \EasyExcel\Interfaces\WriterSheet $sheet
     * @return \EasyExcel\Interfaces\WriterRow
     */
    public function getRowWriter(WriterSheet $sheet): WriterRowInterface
    {
        $handler = $this->getHandlerByWriterSheet($sheet);
        return new WriterRow($handler, $this, $sheet);
    }

    /**
     * @param \EasyExcel\Interfaces\WriterSheet $sheet
     * @return \EasyExcel\Writers\Xlsx\Parts\Sheet
     */
    protected function getHandlerByWriterSheet(WriterSheet $sheet): Sheet
    {
        if (!isset($this->sheetHandlers[$sheet->getIndex()])) {
            $this->sheetHandlers[$sheet->getIndex()] = new Sheet($this, $this->excel, $sheet);
        }

        return $this->sheetHandlers[$sheet->getIndex()];
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
        $excel->addCellXf(new Style());
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
        if (!$this->excel) {
            return;
        }

        $zip = new ZipArchive();
        $zip->open($this->filename, ZipArchive::CREATE | ZipArchive::OVERWRITE);

        $writeParts = [
            '[Content_Types].xml' => new ContentTypes($this, $this->excel),
            '_rels/.rels' => new Rels($this, $this->excel),
            'docProps/app.xml' => new DocPropsApp($this, $this->excel),
            'docProps/core.xml' => new DocPropsCore($this, $this->excel),
            'xl/styles.xml' => new Styles($this, $this->excel),
            'xl/workbook.xml' => new Workbook($this, $this->excel),
            'xl/sharedStrings.xml' => new SharedStrings($this, $this->excel),
            'xl/_rels/workbook.xml.rels' => new WorkbookRels($this, $this->excel),
        ];

        foreach ($this->excel->getAllSheets() as $index => $sheet) {
            $entryName = 'xl/worksheets/sheet' . ($index + 1) . '.xml';
            $writeParts[$entryName] = $this->getHandlerByWriterSheet($sheet);
            $relsEntryName = 'xl/worksheets/_rels/sheet' . ($index + 1) . '.xml.rels';
            $writeParts[$relsEntryName] = $writeParts[$entryName]->getRels();
        }

        foreach ($writeParts as $entryName => $writePart) {
            $zip->addFile($writePart->getFilename(), $entryName);
        }

        $zip->close();

        foreach ($writeParts as $writePart) {
            $writePart->close();
        }
    }
}
