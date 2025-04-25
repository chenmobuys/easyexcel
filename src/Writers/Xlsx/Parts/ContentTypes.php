<?php

namespace EasyExcel\Writers\Xlsx\Parts;

class ContentTypes extends AbstractPart
{
    /**
     * @var \string[][]
     */
    protected $defaultElements = [
        ['Extension' => 'rels', 'ContentType' => 'application/vnd.openxmlformats-package.relationships+xml',],
        ['Extension' => 'xml', 'ContentType' => 'application/xml'],
    ];

    /**
     * @var \string[][]
     */
    protected $overrideElements = [
        [
            'PartName' => '/docProps/app.xml',
            'ContentType' => 'application/vnd.openxmlformats-officedocument.extended-properties+xml',
        ],
        [
            'PartName' => '/docProps/core.xml',
            'ContentType' => 'application/vnd.openxmlformats-package.core-properties+xml',
        ],
        [
            'PartName' => '/xl/styles.xml',
            'ContentType' => 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml',
        ],
        [
            'PartName' => '/xl/sharedStrings.xml',
            'ContentType' => 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml',
        ],
        [
            'PartName' => '/xl/workbook.xml',
            'ContentType' => 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
        ],
    ];

    /**
     * Write start.
     *
     * @return $this
     */
    protected function writeStart(): parent
    {
        $this->xml->startElement('Types');
        $this->xml->writeAttribute('xmlns', 'http://schemas.openxmlformats.org/package/2006/content-types');

        // Add override sheet elements
        foreach ($this->excel->getAllSheets() as $index => $sheet) {
            $this->overrideElements[] = $this->getOverrideSheetAttributes($index);
        }

        // Write default elements
        foreach ($this->defaultElements as $attributes) {
            $this->writeElementWithAttributes('Default', $attributes);
        }

        // Write override elements
        foreach ($this->overrideElements as $attributes) {
            $this->writeElementWithAttributes('Override', $attributes);
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
     * Get override sheet attributes.
     *
     * @param int $index
     * @return string[]
     */
    private function getOverrideSheetAttributes(int $index): array
    {
        return [
            'PartName' => '/xl/worksheets/sheet' . ($index + 1) . '.xml',
            'ContentType' => 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml',
        ];
    }
}