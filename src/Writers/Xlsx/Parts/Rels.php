<?php

namespace EasyExcel\Writers\Xlsx\Parts;

class Rels extends AbstractPart
{
    protected $relationshipElements = [
        [
            'Type' => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument',
            'Target' => 'xl/workbook.xml',
        ],
        [
            'Type' => 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties',
            'Target' => 'docProps/core.xml',
        ],
        [
            'Type' => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties',
            'Target' => 'docProps/app.xml',
        ],
    ];

    /**
     * Write start.
     *
     * @return $this
     */
    protected function writeStart(): parent
    {
        $this->xml->startElement('Relationships');
        $this->xml->writeAttribute('xmlns', self::NS_RELATIONSHIPS);

        // Write relationship elements
        foreach ($this->relationshipElements as $index => $attributes) {
            $attributes['Id'] = 'rId' . ($index + 1);
            $this->writeElementWithAttributes('Relationship', $attributes);
        }

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