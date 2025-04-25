<?php

namespace EasyExcel\Writers\Xlsx\Parts;

class WorkbookRels extends AbstractPart
{
    protected $relationshipElements = [
        [
            'Type' => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles',
            'Target' => 'styles.xml',
        ],
        [
            'Type' => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings',
            'Target' => 'sharedStrings.xml',
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

        $relationshipElements = [];
        // Add sheet relationships
        foreach ($this->excel->getAllSheets() as $index => $sheet) {
            $relationshipElements[] = $this->getRelationshipSheetAttributes($index);
        }

        $this->relationshipElements = array_merge($relationshipElements, $this->relationshipElements);

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

    /**
     * Get relationship sheet attributes.
     *
     * @param int $index
     * @return string[]
     */
    private function getRelationshipSheetAttributes(int $index): array
    {
        return [
            'Type' => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',
            'Target' => 'worksheets/sheet' . ($index + 1) . '.xml',
        ];
    }
}