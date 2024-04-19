<?php

namespace EasyExcel\Writers\Xlsx\Parts;

use EasyExcel\Interfaces\SheetInterface;
use EasyExcel\Writers\Xlsx\Part;

class SheetPart extends Part
{
    /**
     * @var int
     */
    protected $rId = 1;

    /**
     * @var SheetRelsPart
     */
    protected $rels;

    /**
     * @var SheetInterface
     */
    protected $sheet;

    /**
     * @param  SheetInterface  $sheet
     */
    public function __construct(SheetInterface $sheet)
    {
        $this->sheet = $sheet;
        $this->rels = new SheetRelsPart($sheet);
        parent::__construct($sheet->getExcel());
    }

    /**
     * Get sheet rels.
     *
     * @return \EasyExcel\Writers\Xlsx\Parts\SheetRelsPart
     */
    public function getRels(): SheetRelsPart
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
        if ($mergeCellsCount > 0) {
            $this->xml->startElement('mergeCells');
            $this->xml->writeAttribute('count', count($this->sheet->getMergeCells()));
            foreach ($this->sheet->getMergeCells() as $mergeCell) {
                $this->writeElementWithAttributes('mergeCell', ['ref' => $mergeCell]);
            }
            $this->xml->endElement();
        }

        // Hyperlinks
        $hyperlinksCount = count($this->sheet->getHyperlinks());
        if ($hyperlinksCount > 0) {
            $this->xml->startElement('hyperlinks');
            foreach ($this->sheet->getHyperlinks() as $coordinate => $hyperlink) {
                $attributes = ['ref' => $coordinate, 'r:id' => 'rId'.$this->rId];
                if ($hyperlink->getTooltip()) {
                    $attributes['tooltip'] = $hyperlink->getTooltip();
                }
                $this->writeElementWithAttributes('hyperlink', $attributes);
                $this->rId++;
            }
            $this->xml->endElement();
        }

        // AutoFilter
        if ($this->sheet->getAutoFilter()) {
            $this->writeElementWithAttributes('autoFilter', ['ref' => $this->sheet->getAutoFilter()->getRange()]);
        }

        $this->xml->endElement();

        return $this;
    }
}
