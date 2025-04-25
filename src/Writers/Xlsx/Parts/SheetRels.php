<?php

namespace EasyExcel\Writers\Xlsx\Parts;

use EasyExcel\Interfaces\WriterExcel;
use EasyExcel\Writers\WriterSheet;
use EasyExcel\Writers\Xlsx\Writer;

class SheetRels extends AbstractPart
{
    /**
     * @var int
     */
    protected $rId = 1;

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
        parent::__construct($writer, $excel);
    }

    /**
     * Write start.
     *
     * @return $this
     */
    protected function writeStart(): parent
    {
        $this->xml->startElement('Relationships');
        $this->xml->writeAttribute('xmlns', self::NS_RELATIONSHIPS);

        return $this;
    }

    /**
     * Write end.
     *
     * @return $this
     */
    protected function writeEnd(): parent
    {
        // Hyperlinks
        foreach ($this->sheet->getHyperlinks() as $hyperlink) {
            $attributes = [
                'Id' => 'rId' . $this->rId,
                'Type' => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
                'Target' => $hyperlink->getUrl(),
                'TargetMode' => $hyperlink->isInternal() ? 'Internal' : 'External'
            ];
            $this->writeElementWithAttributes('Relationship', $attributes);
            $this->rId++;
        }

        $this->xml->endElement();

        return $this;
    }
}