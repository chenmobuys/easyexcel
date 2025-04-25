<?php

namespace EasyExcel\Writers\Xlsx\Parts;

class DocPropsApp extends AbstractPart
{
    /**
     * Write start.
     *
     * @return $this
     */
    protected function writeStart(): parent
    {
        $sheetCount = count($this->excel->getAllSheets());

        $this->xml->startElement('Properties');
        $this->xml->writeAttribute('xmlns', 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties');
        $this->xml->writeAttribute('xmlns:vt', 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes');

        // Application
        $this->xml->writeElement('Application', 'Microsoft Excel');

        // DocSecurity
        $this->xml->writeElement('DocSecurity', '0');

        // ScaleCrop
        $this->xml->writeElement('ScaleCrop', 'false');

        // HeadingPairs
        $this->xml->startElement('HeadingPairs');

        // Vector
        $this->xml->startElement('vt:vector');
        $this->xml->writeAttribute('size', '2');
        $this->xml->writeAttribute('baseType', 'variant');

        // Variant
        $this->xml->startElement('vt:variant');
        $this->xml->writeElement('vt:lpstr', 'Worksheets');
        $this->xml->endElement();

        // Variant
        $this->xml->startElement('vt:variant');
        $this->xml->writeElement('vt:i4', $sheetCount);
        $this->xml->endElement();

        $this->xml->endElement();

        $this->xml->endElement();

        // TitlesOfParts
        $this->xml->startElement('TitlesOfParts');

        // Vector
        $this->xml->startElement('vt:vector');
        $this->xml->writeAttribute('size', $sheetCount);
        $this->xml->writeAttribute('baseType', 'lpstr');

        foreach ($this->excel->getAllSheets() as $sheet) {
            $this->xml->writeElement('vt:lpstr', $sheet->getName());
        }

        $this->xml->endElement();

        $this->xml->endElement();

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

        return $this;
    }
}