<?php

namespace EasyExcel\Writers\Xlsx\Parts;

use EasyExcel\Interfaces\WriterExcel;
use EasyExcel\Writers\WriterSheet;
use EasyExcel\Writers\Xlsx\Writer;

class Sheet extends AbstractPart
{
    /**
     * @var int
     */
    protected $rId = 1;

    /**
     * @var \EasyExcel\Writers\Xlsx\Parts\SheetRels
     */
    protected $rels;

    /**
     * @var \EasyExcel\Writers\WriterSheet
     */
    protected $sheet;

    /**
     * @param \EasyExcel\Writers\Xlsx\Writer $writer
     * @param \EasyExcel\Interfaces\WriterExcel $excel
     * @param \EasyExcel\Writers\WriterSheet $sheet
     */
    public function __construct(Writer $writer, WriterExcel $excel, WriterSheet $sheet)
    {
        $this->sheet = $sheet;
        $this->rels = new SheetRels($writer, $excel, $sheet);
        parent::__construct($writer, $excel);
    }

    /**
     * Get sheet rels.
     *
     * @return \EasyExcel\Writers\Xlsx\Parts\SheetRels
     */
    public function getRels(): SheetRels
    {
        return $this->rels;
    }

    /**
     * Write start.
     *
     * @return $this
     */
    protected function writeStart(): parent
    {
        $this->xml->startElement('worksheet');
        $this->xml->writeAttribute('xmlns', self::NS_MAIN);
        $this->xml->writeAttribute('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');
        $this->xml->writeAttribute('xmlns:xdr', 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing');
        $this->xml->writeAttribute('xmlns:x14', 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main');
        $this->xml->writeAttribute('xmlns:mc', 'http://schemas.openxmlformats.org/markup-compatibility/2006');
        $this->xml->writeAttribute('xmlns:etc', 'http://www.wps.cn/officeDocument/2017/etCustomData');

        $this->xml->startElement('sheetData');

        return $this;
    }

    /**
     * Write end.
     *
     * @return $this
     */
    protected function writeEnd(): parent
    {
        $this->xml->endElement();

        // MergeCells
        $mergeCellsCount = count($this->sheet->getMergeCells());
        if($mergeCellsCount > 0) {
            $this->xml->startElement('mergeCells');
            $this->xml->writeAttribute('count', count($this->sheet->getMergeCells()));
            foreach ($this->sheet->getMergeCells() as $mergeCell) {
                $this->writeElementWithAttributes('mergeCell', ['ref' => $mergeCell]);
            }
            $this->xml->endElement();
        }

        // Hyperlinks
        $hyperlinksCount = count($this->sheet->getHyperlinks());
        if($hyperlinksCount > 0) {
            $this->xml->startElement('hyperlinks');
            foreach ($this->sheet->getHyperlinks() as $coordinate => $hyperlink) {
                $attributes = ['ref' => $coordinate, 'r:id' => 'rId' . $this->rId];
                if ($hyperlink->getTooltip()) {
                    $attributes['tooltip'] = $hyperlink->getTooltip();
                }
                $this->writeElementWithAttributes('hyperlink', $attributes);
                $this->rId++;
            }
            $this->xml->endElement();
        }

        // AutoFilter
        if($this->sheet->getAutoFilter()) {
            $this->writeElementWithAttributes('autoFilter', ['ref' => $this->sheet->getAutoFilter()->getRange()]);
        }

        $this->xml->endElement();

        return $this;
    }
}