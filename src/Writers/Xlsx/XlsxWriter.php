<?php

namespace EasyExcel\Writers\Xlsx;

use EasyExcel\Interfaces\SheetInterface;
use EasyExcel\Interfaces\WriterInterface;
use EasyExcel\Interfaces\WriterRowInterface;
use EasyExcel\Metadata\Style;
use EasyExcel\Writers\Writer;
use EasyExcel\Writers\Xlsx\Parts\ContentTypesPart;
use EasyExcel\Writers\Xlsx\Parts\DocPropsAppPart;
use EasyExcel\Writers\Xlsx\Parts\DocPropsCorePart;
use EasyExcel\Writers\Xlsx\Parts\RelsPart;
use EasyExcel\Writers\Xlsx\Parts\SharedStringsPart;
use EasyExcel\Writers\Xlsx\Parts\SheetPart;
use EasyExcel\Writers\Xlsx\Parts\StylesPart;
use EasyExcel\Writers\Xlsx\Parts\WorkbookPart;
use EasyExcel\Writers\Xlsx\Parts\WorkbookRelsPart;
use ZipArchive;

class XlsxWriter extends Writer
{
    /**
     * @var \EasyExcel\Writers\Xlsx\Parts\SheetPart[]
     */
    protected $sheetHandlers = [];

    /**
     * @param  SheetInterface  $sheet
     * @return WriterRowInterface
     */
    public function getRowWriterBySheet(SheetInterface $sheet): WriterRowInterface
    {
        $handler = $this->getHandlerByWriterSheet($sheet);
        return new XlsxWriterRow($handler, $sheet);
    }

    /**
     * @param  SheetInterface  $sheet
     * @return SheetPart
     */
    protected function getHandlerByWriterSheet(SheetInterface $sheet): SheetPart
    {
        if (!isset($this->sheetHandlers[$sheet->getIndex()])) {
            $this->sheetHandlers[$sheet->getIndex()] = new SheetPart($sheet);
        }

        return $this->sheetHandlers[$sheet->getIndex()];
    }

    /**
     * @param  string  $filename
     * @return $this
     */
    protected function openFromFile(string $filename): WriterInterface
    {
        $this->addCellXf(new Style());
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
        $zip = new ZipArchive();
        $zip->open($this->filename, ZipArchive::CREATE | ZipArchive::OVERWRITE);

        $writeParts = [
            '[Content_Types].xml'        => new ContentTypesPart($this),
            '_rels/.rels'                => new RelsPart($this),
            'docProps/app.xml'           => new DocPropsAppPart($this),
            'docProps/core.xml'          => new DocPropsCorePart($this),
            'xl/styles.xml'              => new StylesPart($this),
            'xl/workbook.xml'            => new WorkbookPart($this),
            'xl/sharedStrings.xml'       => new SharedStringsPart($this),
            'xl/_rels/workbook.xml.rels' => new WorkbookRelsPart($this),
        ];

        foreach ($this->getAllSheets() as $index => $sheet) {
            $entryName = 'xl/worksheets/sheet'.($index + 1).'.xml';
            $writeParts[$entryName] = $this->getHandlerByWriterSheet($sheet);
            $relsEntryName = 'xl/worksheets/_rels/sheet'.($index + 1).'.xml.rels';
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
