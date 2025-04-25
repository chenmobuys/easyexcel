<?php

namespace EasyExcel\Writers\Xlsx\Parts;

class DocPropsCore extends AbstractPart
{
    /**
     * Write start.
     *
     * @return $this
     */
    protected function writeStart(): parent
    {
        $this->xml->startElement('cp:coreProperties');
        $this->xml->writeAttribute('xmlns:cp', 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties');
        $this->xml->writeAttribute('xmlns:dc', 'http://purl.org/dc/elements/1.1/');
        $this->xml->writeAttribute('xmlns:dcterms', 'http://purl.org/dc/terms/');
        $this->xml->writeAttribute('xmlns:dcmitype', 'http://purl.org/dc/dcmitype/');
        $this->xml->writeAttribute('xmlns:xsi', 'http://www.w3.org/2001/XMLSchema-instance');

        $this->xml->writeElement('dc:creator', 'EasyExcel');

        $this->xml->writeElement('cp:lastModifiedBy', 'EasyExcel');

        $this->xml->startElement('dcterms:created');
        $this->xml->writeAttribute('xsi:type', 'dcterms:W3CDTF');
        $this->xml->writeRaw(date(DATE_W3C, time()));
        $this->xml->endElement();

        $this->xml->startElement('dcterms:modified');
        $this->xml->writeAttribute('xsi:type', 'dcterms:W3CDTF');
        $this->xml->writeRaw(date(DATE_W3C, time()));
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