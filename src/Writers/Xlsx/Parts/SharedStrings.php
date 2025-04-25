<?php

namespace EasyExcel\Writers\Xlsx\Parts;

class SharedStrings extends AbstractPart
{
    /**
     * Write start.
     *
     * @return $this
     */
    protected function writeStart(): parent
    {
        $this->xml->startElement('sst');
        $this->xml->writeAttribute('xmlns', self::NS_MAIN);
        $this->xml->writeAttribute('count', $this->excel->getSharedStringsCount());
        $this->xml->writeAttribute('uniqueCount', $this->excel->getSharedStringsCount());

        // Do not use shared strings.
        // foreach ($this->excel->getSharedStrings() as $sharedString) {
        //     $this->xml->startElement('si');
        //     $this->xml->writeElement('t', $sharedString);
        //     $this->xml->endElement();
        // }

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