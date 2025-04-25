<?php

namespace EasyExcel\Writers\Xlsx\Parts;

use EasyExcel\Metadata\Style\Format;

class Workbook extends AbstractPart
{
    /**
     * Write start.
     *
     * @return $this
     */
    protected function writeStart(): parent
    {
        $this->xml->startElement('workbook');
        $this->xml->writeAttribute('xmlns', self::NS_MAIN);
        $this->xml->writeAttribute('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');

        $this->xml->startElement('workbookPr');
        if (Format::getCalendar() === Format::CALENDAR_MAC_1904) {
            $this->xml->writeAttribute('date1904', 1);
        }
        $this->xml->endElement();

        $this->xml->startElement('bookViews');
        $this->xml->startElement('workbookView');
        $this->xml->writeAttribute('activeTab', 0);
        $this->xml->writeAttribute('firstSheet', 0);
        $this->xml->writeAttribute('showSheetTabs', true);
        $this->xml->writeAttribute('showHorizontalScroll', true);
        $this->xml->writeAttribute('showVerticalScroll', true);
        $this->xml->endElement();
        $this->xml->endElement();

        $this->xml->startElement('sheets');
        foreach ($this->excel->getAllSheets() as $index => $sheet) {
            $this->xml->startElement('sheet');
            $this->xml->writeAttribute('name', $sheet->getName());
            $this->xml->writeAttribute('sheetId', $index + 1);
            $this->xml->writeAttribute('r:id', 'rId' . ($index + 1));
            $this->xml->endElement();
        }
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