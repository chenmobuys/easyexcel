<?php

namespace EasyExcel\Writers\Xlsx\Parts;

use EasyExcel\Interfaces\SheetInterface;
use EasyExcel\Writers\Xlsx\Part;

class SheetRelsPart extends Part
{
    /**
     * @var int
     */
    protected $rId = 1;

    /**
     * @var \EasyExcel\Interfaces\SheetInterface
     */
    protected $sheet;

    /**
     * @param  SheetInterface  $sheet
     */
    public function __construct(SheetInterface $sheet)
    {
        $this->sheet = $sheet;
        parent::__construct($sheet->getExcel());
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
                'Id'         => 'rId'.$this->rId,
                'Type'       => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
                'Target'     => $hyperlink->getUrl(),
                'TargetMode' => $hyperlink->getInternal() ? 'Internal' : 'External'
            ];
            $this->writeElementWithAttributes('Relationship', $attributes);
            $this->rId++;
        }

        $this->xml->endElement();

        return $this;
    }
}
