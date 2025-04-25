<?php

namespace EasyExcel\Readers\Xlsx;

use EasyExcel\Helpers\Coordinate;
use EasyExcel\Interfaces\ReaderExcel as ReaderExcelInterface;
use EasyExcel\Interfaces\ReaderRow as ReaderRowInterface;
use EasyExcel\Interfaces\ReaderSheet;
use EasyExcel\Metadata\AutoFilter;
use EasyExcel\Metadata\Style;
use EasyExcel\Metadata\Style\Format;
use EasyExcel\Readers\BaseReader;
use EasyExcel\Readers\ReaderExcel;
use EasyExcel\Settings;
use SimpleXMLElement;
use Throwable;
use XMLReader;
use ZipArchive;

/**
 * @see http://officeopenxml.com/anatomyofOOXML-xlsx.php
 */
class Reader extends BaseReader
{
    // File
    public const FILE_RELS_RELS = '_rels/.rels';

    // Namespace
    public const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
    public const NS_RELATIONSHIPS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
    public const NS_OFFICE_DOCUMENT = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument';
    public const NS_CORE_PROPERTIES = 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties';
    public const NS_EXTENDED_PROPERTIES = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties';
    public const NS_THEME = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme';
    public const NS_STYLES = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles';
    public const NS_WORKSHEET = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet';
    public const NS_SHARED_STRINGS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings';
    public const NS_DRAWINGML = 'http://schemas.openxmlformats.org/drawingml/2006/main';
    public const NS_HYPERLINK = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink';

    // Element
    public const EL_SHEET = 'sheet';
    public const EL_SHEET_DATA = 'sheetData';
    public const EL_WORKBOOK_PR = 'workbookPr';
    public const EL_MERGE_CELL = 'mergeCell';
    public const EL_HYPERLINK = 'hyperlink';
    public const EL_AUTO_FILTER = 'autoFilter';
    public const EL_ROW = 'row';
    public const EL_C = 'c';
    public const EL_V = 'v';
    public const EL_F = 'f';
    public const EL_IS = 'is';

    /**
     * @var \ZipArchive|null
     */
    protected $zip;

    /**
     * @var array
     */
    protected $mainRelationships = [];

    /**
     * @var array
     */
    protected $workbookRelationships = [];

    /**
     * @var array
     */
    protected $workbookXmlRelationships = [];

    /**
     * @var array
     */
    protected $hyperlinkRelationships = [];

    /**
     * @var ?\EasyExcel\Readers\Xlsx\Theme
     */
    protected $theme;

    /**
     * @var array
     */
    protected $palette = [];

    /**
     * @var \SimpleXMLElement[]
     */
    protected $fills = [];

    /**
     * @var \SimpleXMLElement[]
     */
    protected $fonts = [];

    /**
     * @var \SimpleXMLElement[]
     */
    protected $borders = [];

    /**
     * @var \SimpleXMLElement[]
     */
    protected $numFmts = [];

    /**
     * @var \SimpleXMLElement[]
     */
    protected $cellXfs = [];

    /**
     * Get theme.
     *
     * @return \EasyExcel\Readers\Xlsx\Theme|null
     */
    public function getTheme(): ?Theme
    {
        return $this->theme;
    }

    /**
     * Determine whether the file is readable.
     *
     * @param string $filename
     * @return bool
     */
    public function readable(string $filename): bool
    {
        if ((int) @filesize($filename) === 0) {
            return false;
        }

        $result = false;
        $zip = new ZipArchive();

        if ($zip->open($filename) === true) {
            try {
                $relsXml = $this->getXMLReaderFromName(self::FILE_RELS_RELS, $filename);
                while ($relsXml->read()) {
                    if (
                        $relsXml->getAttribute('Type') == self::NS_OFFICE_DOCUMENT
                        && preg_match('/workbook.*\.xml/', basename($relsXml->getAttribute('Target')))
                    ) {
                        $result = true;
                        break;
                    }
                }
            } catch (Throwable $e) {
                // Do nothing.
            } finally {
                if (isset($relsXml)) {
                    $relsXml->close();
                }
            }

            $zip->close();
        }

        return $result;
    }

    /**
     * Get row iterator.
     *
     * @param \EasyExcel\Interfaces\ReaderSheet $sheet
     * @param int $startRow
     * @param int|null $endRow
     * @return \EasyExcel\Interfaces\ReaderRow
     */
    public function getRowIterator(ReaderSheet $sheet, int $startRow = 1, ?int $endRow = null): ReaderRowInterface
    {
        $handler = 'zip://' . $this->zip->filename . '#' . $this->workbookXmlRelationships[$sheet->getIndex()];

        return new ReaderRow($handler, $sheet, $startRow, $endRow);
    }

    /**
     * Load from file.
     *
     * @param string $filename
     * @return \EasyExcel\Interfaces\ReaderExcel
     * @throws \Exception
     */
    protected function loadFromFile(string $filename): ReaderExcelInterface
    {
        $excel = new ReaderExcel($this);

        $this->zip = new ZipArchive();
        $this->zip->open($filename);

        $mainRelationships = $this->getMainRelationships();
        $workbookRelationships = $this->getWorkbookRelationships();

        // Worksheet
        if (
            isset($mainRelationships[self::NS_OFFICE_DOCUMENT])
            && isset($workbookRelationships[self::NS_WORKSHEET])
        ) {
            $workbookXml = $this->getXMLReaderFromName($mainRelationships[self::NS_OFFICE_DOCUMENT]);

            while ($workbookXml->read()) {
                if ($workbookXml->localName == self::EL_WORKBOOK_PR) {
                    Format::setCalendar(
                        $workbookXml->getAttribute('date1904')
                            ? Format::CALENDAR_MAC_1904
                            : Format::CALENDAR_WINDOWS_1900
                    );
                    continue;
                }

                if ($workbookXml->localName != self::EL_SHEET) {
                    continue;
                }

                $id = $workbookXml->getAttributeNs('id', self::NS_RELATIONSHIPS);
                $name = $workbookXml->getAttribute('name');
                $sheet = $excel->addSheet($name);
                $sheetXmlPathname = $workbookRelationships[self::NS_WORKSHEET][$id];
                $this->workbookXmlRelationships[$sheet->getIndex()] = $sheetXmlPathname;
                // Read sheet relationships.
                $hyperlinkRelationships = $this->getHyperlinkRelationshipsBySheetIndex($sheet->getIndex());

                $lastRowNumber = 0;
                $lastColumnLetter = null;
                $mergedCells = [];

                $sheetXml = $this->getXMLReaderFromName($sheetXmlPathname);
                while ($sheetXml->read()) {
                    switch ($sheetXml->localName) {
                        case self::EL_ROW:
                            do {
                                $lastRowNumber = $sheetXml->getAttribute('r');
                                if (!$sheetXml->isEmptyElement) {
                                    $sheetXml->read();
                                    do {
                                        $columnLetter = preg_replace('/\d+/', '', $sheetXml->getAttribute('r'));
                                        $lastColumnLetter = ($lastColumnLetter > $columnLetter) ? $lastColumnLetter : $columnLetter;
                                    } while ($sheetXml->next() && $sheetXml->localName == self::EL_C);
                                }
                            } while ($sheetXml->next() && $sheetXml->localName == self::EL_ROW);
                            break;
                        case self::EL_AUTO_FILTER:
                            if ($ref = (string) $sheetXml->getAttribute('ref')) {
                                $autoFilter = new AutoFilter($ref);
                                $sheet->setAutoFilter($autoFilter);
                            }
                            break;
                        case self::EL_MERGE_CELL:
                            $mergedCells[] = $sheetXml->getAttribute('ref');
                            break;
                        case self::EL_HYPERLINK:
                            $id = (string) $sheetXml->getAttributeNS('id', self::NS_RELATIONSHIPS);
                            $ref = (string) $sheetXml->getAttribute('ref');
                            $tooltip = (string) $sheetXml->getAttribute('tooltip');
                            $location = (string) $sheetXml->getAttribute('location');

                            $coordinates = Coordinate::coordinatesFromRange($ref);
                            $hyperlinkRelationship = $hyperlinkRelationships[$id] ?? null;
                            foreach ($coordinates as $coordinate) {
                                $hyperlink = $sheet->getHyperlink($coordinate);
                                if ($tooltip) {
                                    $hyperlink->setTooltip($tooltip);
                                }
                                if ($hyperlinkRelationship) {
                                    $hyperlink->setUrl($hyperlinkRelationship['Target']);
                                } elseif ($location) {
                                    $hyperlink->setUrl('sheet://' . $location);
                                }
                            }
                            break;
                    }
                }
                $sheetXml->close();
                $totalRows = $lastRowNumber;
                $totalColumns = $lastColumnLetter ? (Coordinate::columnIndexFromColumnLetter($lastColumnLetter) + 1) : 0;

                $sheet->setTotalRows($totalRows)
                    ->setTotalColumns($totalColumns)
                    ->setMergeCells($mergedCells);
            }

            $workbookXml->close();
        }

        // Shared strings
        if (isset($workbookRelationships[self::NS_SHARED_STRINGS])) {
            $sharedStringsXml = $this->getXMLReaderFromName($workbookRelationships[self::NS_SHARED_STRINGS]);
            $cacheValue = '';
            while ($sharedStringsXml->read()) {
                switch ($sharedStringsXml->name) {
                    case 'si':
                        if ($sharedStringsXml->nodeType == XMLReader::END_ELEMENT) {
                            $excel->addSharedString($cacheValue);
                            $cacheValue = '';
                        }
                        break;
                    case 't':
                        if ($sharedStringsXml->nodeType == XMLReader::END_ELEMENT) {
                            break;
                        }
                        $cacheValue .= $sharedStringsXml->readString();
                        break;
                }
            }

            $sharedStringsXml->close();
        }

        // Theme
        if (isset($workbookRelationships[self::NS_THEME])) {
            $themeXml = new SimpleXMLElement(
                $this->zip->getFromName($workbookRelationships[self::NS_THEME]),
                Settings::getLibXmlLoaderOptions()
            );
            $themeOrderArray = ['lt1', 'dk1', 'lt2', 'dk2'];
            $themeOrderAdditional = count($themeOrderArray);
            $drawingNS = self::NS_DRAWINGML;
            $themeName = (string) $themeXml->attributes()['name'];
            $themeXml = $themeXml->children($drawingNS);

            $colourScheme = $themeXml->themeElements->clrScheme->attributes();
            $colourSchemeName = (string) $colourScheme['name'];
            $colourScheme = $themeXml->themeElements->clrScheme->children($drawingNS);

            $themeColours = [];
            foreach ($colourScheme as $k => $xmlColour) {
                $themePos = array_search($k, $themeOrderArray);
                if ($themePos === false) {
                    $themePos = $themeOrderAdditional++;
                }
                if (isset($xmlColour->sysClr)) {
                    $xmlColourData = $xmlColour->sysClr->attributes();
                    $themeColours[$themePos] = (string) $xmlColourData['lastClr'];
                } elseif (isset($xmlColour->srgbClr)) {
                    $xmlColourData = $xmlColour->srgbClr->attributes();
                    $themeColours[$themePos] = (string) $xmlColourData['val'];
                }
            }
            $this->theme = new Theme($themeName, $colourSchemeName, $themeColours);
        }

        // Styles
        if (isset($workbookRelationships[self::NS_STYLES])) {
            $stylesXml = new SimpleXMLElement(
                $this->zip->getFromName($workbookRelationships[self::NS_STYLES]),
                Settings::getLibXmlLoaderOptions()
            );

            if ($stylesXml->colors && $stylesXml->colors->indexedColors && $stylesXml->colors->indexedColors->rgbColor) {
                $palette = [];
                foreach ($stylesXml->colors->indexedColors->rgbColor as $node) {
                    if (!is_null($node)) {
                        $attr = $node->attributes();
                        if (isset($attr['rgb'])) {
                            $palette[] = (string) $attr['rgb'];
                        }
                    }
                }
                $this->palette = (count($palette) === 64) ? $palette : [];
            }

            if ($stylesXml->fills) {
                foreach ($stylesXml->fills->fill as $node) {
                    $this->fills[] = $node;
                }
            }

            if ($stylesXml->fonts) {
                foreach ($stylesXml->fonts->font as $node) {
                    $this->fonts[] = $node;
                }
            }

            if ($stylesXml->borders) {
                foreach ($stylesXml->borders->border as $node) {
                    $this->borders[] = $node;
                }
            }

            if ($stylesXml->cellXfs) {
                foreach ($stylesXml->cellXfs->xf as $node) {
                    $this->cellXfs[] = $node;
                }
            }

            if ($stylesXml->numFmts) {
                foreach ($stylesXml->numFmts->numFmt as $node) {
                    $this->numFmts[(int) $node->attributes()->numFmtId]
                        = (string) $node->attributes()->formatCode;
                }
            }

            foreach ($this->cellXfs as $xf) {
                $style = new Style();

                $numFmtId = (int) $xf->attributes()->numFmtId;
                $numFmt = $this->numFmts[$numFmtId] ?? Format::builtInFormatCode($numFmtId);
                $font = $this->fonts[(int) ($xf->attributes()->fontId)];
                $fill = $this->fills[(int) ($xf->attributes()->fillId)];
                $border = $this->borders[(int) ($xf->attributes()->borderId)];
                $alignment = $xf->alignment;
                $protection = $xf->protection;
                $quotePrefix = (bool) ($xf->attributes()->quotePrefix ?? false);

                $style->getFormat()->setFormatCode($numFmt ?: Format::FORMAT_GENERAL);

                if ($font) {
                    if ($font->name) {
                        $style->getFont()->setName((string) $font->name->attributes()->val);
                    }
                    if ($font->sz) {
                        $style->getFont()->setSize((string) $font->sz->attributes()->val);
                    }
                    if ($font->color) {
                        $style->getFont()->getColor()->setArgb($this->readColor($font->color));
                    }
                    if ($font->b) {
                        $style->getFont()->setBold(!isset($font->b->attributes()->val) || $font->b->attributes()->val);
                    }
                    if ($font->i) {
                        $style->getFont()->setItalic(!isset($font->i->attributes()->val) || $font->i->attributes()->val);
                    }
                    if ($font->strike) {
                        $style->getFont()->setStrikethrough(!isset($font->strike->attributes()->val) || $font->strike->attributes()->val);
                    }
                    if ($font->u) {
                        $style->getFont()->setUnderline($font->u->attributes()->val ?? Style\Font::UNDERLINE_SINGLE);
                    }
                    if (isset($font->vertAlign, $font->vertAlign->attributes()->val)) {
                        $vertAlign = strtolower((string) $font->vertAlign->attributes()->val);
                        $style->getFont()->setSubscript($vertAlign == 'subscript');
                        $style->getFont()->setSuperscript($vertAlign == 'superscript');
                    }
                }

                if ($fill) {
                    if ($fill->gradientFill) {
                        if ($fill->gradientFill->attributes()->type) {
                            $style->getFill()->setType((string) $fill->gradientFill->attributes()->type);
                        }
                        $style->getFill()->setRotation((float) $fill->gradientFill->attributes()->degree);
                        foreach ($fill->gradientFill->stop as $stop) {
                            if ((int) $stop->attributes()->position > 0) {
                                $style->getFill()->getEndColor()->setArgb($this->readColor($stop->color));
                            } else {
                                $style->getFill()->getStartColor()->setArgb($this->readColor($stop->color));
                            }
                        }
                    } elseif ($fill->patternFill) {
                        $patternType = ((string) $fill->patternFill->attributes()->patternType) ?: Style\Fill::FILL_SOLID;
                        $style->getFill()->setType($patternType);
                        if ($fill->patternFill->fgColor) {
                            $style->getFill()->getStartColor()->setArgb($this->readColor($fill->patternFill->fgColor,
                                true));
                        }
                        if ($fill->patternFill->bgColor) {
                            $style->getFill()->getEndColor()->setArgb($this->readColor($fill->patternFill->bgColor,
                                true));
                        }
                    }
                }

                if ($border) {
                    $diagonalUp = (bool) (string) $border->attributes()->diagonalUp;
                    $diagonalDown = (bool) (string) $border->attributes()->diagonalDown;
                    if (!$diagonalUp && !$diagonalDown) {
                        $style->getBorders()->setDiagonalDirection(Style\Borders::DIAGONAL_NONE);
                    } elseif ($diagonalUp && !$diagonalDown) {
                        $style->getBorders()->setDiagonalDirection(Style\Borders::DIAGONAL_UP);
                    } elseif (!$diagonalUp && $diagonalDown) {
                        $style->getBorders()->setDiagonalDirection(Style\Borders::DIAGONAL_DOWN);
                    } else {
                        $style->getBorders()->setDiagonalDirection(Style\Borders::DIAGONAL_BOTH);
                    }
                    $this->readBorder($style->getBorders()->getLeft(), $border->left);
                    $this->readBorder($style->getBorders()->getRight(), $border->right);
                    $this->readBorder($style->getBorders()->getTop(), $border->top);
                    $this->readBorder($style->getBorders()->getBottom(), $border->bottom);
                    $this->readBorder($style->getBorders()->getDiagonal(), $border->diagonal);
                }

                if ($alignment) {
                    $style->getAlignment()->setHorizontal(((string) $alignment->attributes()->horizontal) ?: Style\Alignment::HORIZONTAL_GENERAL);
                    $style->getAlignment()->setVertical(((string) $alignment->attributes()->vertical) ?: Style\Alignment::VERTICAL_CENTER);

                    $textRotation = 0;
                    if ((int) $alignment->attributes()->textRotation <= 90) {
                        $textRotation = (int) $alignment->attributes()->textRotation;
                    } elseif ((int) $alignment->attributes()->textRotation > 90) {
                        $textRotation = 90 - (int) $alignment->attributes()->textRotation;
                    }

                    $style->getAlignment()->setTextRotation($textRotation);
                    $style->getAlignment()->setWrapText((bool) (string) $alignment->attributes()->wrapText);
                    $style->getAlignment()->setShrinkToFit((bool) (string) $alignment->attributes()->shrinkToFit);
                    $style->getAlignment()->setIndent((int) ((string) $alignment->attributes()->indent) > 0 ? (int) ((string) $alignment->attributes()->indent) : 0);
                    $style->getAlignment()->setReadOrder((int) ((string) $alignment->attributes()->readingOrder) > 0 ? (int) ((string) $alignment->attributes()->readingOrder) : 0);
                }

                if ($protection) {
                    $locked = (bool) $protection->attributes()->locked ?? false;
                    $hidden = (bool) $protection->attributes()->hidden ?? false;
                    $style->getProtection()->setLocked($locked ? Style\Protection::PROTECTION_PROTECTED : Style\Protection::PROTECTION_UNPROTECTED);
                    $style->getProtection()->setHidden($hidden ? Style\Protection::PROTECTION_PROTECTED : Style\Protection::PROTECTION_UNPROTECTED);
                }

                if ($quotePrefix) {
                    $style->setQuotePrefix(true);
                }

                $excel->addCellXf($style);
            }
        }

        return $excel;
    }

    /**
     * Get main relationships.
     *
     * @return array
     */
    private function getMainRelationships(): array
    {
        if (empty($this->mainRelationships)) {
            $relsXml = $this->getXMLReaderFromName(self::FILE_RELS_RELS);

            while ($relsXml->read()) {
                switch ($relsXml->getAttribute('Type')) {
                    case self::NS_OFFICE_DOCUMENT:
                        $this->mainRelationships[self::NS_OFFICE_DOCUMENT] = $relsXml->getAttribute('Target');
                        break;
                    case self::NS_CORE_PROPERTIES:
                        $this->mainRelationships[self::NS_CORE_PROPERTIES] = $relsXml->getAttribute('Target');
                        break;
                    case self::NS_EXTENDED_PROPERTIES:
                        $this->mainRelationships[self::NS_EXTENDED_PROPERTIES] = $relsXml->getAttribute('Target');
                        break;
                }
            }

            $relsXml->close();
        }

        return $this->mainRelationships;
    }

    /**
     * Get workbook relationships.
     *
     * @return array
     */
    private function getWorkbookRelationships(): array
    {
        if (empty($this->workbookRelationships)) {
            $mainRelationships = $this->getMainRelationships();
            $officeDocumentTarget = $mainRelationships[self::NS_OFFICE_DOCUMENT];
            $baseDirectory = dirname($officeDocumentTarget);
            $workbookBasename = basename($officeDocumentTarget);
            $workbookRelsFilepath = $baseDirectory . '/_rels/' . $workbookBasename . '.rels';

            $workbookRelsXml = $this->getXMLReaderFromName($workbookRelsFilepath);
            while ($workbookRelsXml->read()) {
                $id = $workbookRelsXml->getAttribute('Id');
                $target = (string) $workbookRelsXml->getAttribute('Target');
                $isAbsolute = strpos($target, '/') === 0;
                $absoluteTarget = $isAbsolute ? trim($target, '/') : ($baseDirectory . '/' . $target);
                switch ($workbookRelsXml->getAttribute('Type')) {
                    case self::NS_THEME:
                        $this->workbookRelationships[self::NS_THEME] = $absoluteTarget;
                        break;
                    case self::NS_STYLES:
                        $this->workbookRelationships[self::NS_STYLES] = $absoluteTarget;
                        break;
                    case self::NS_SHARED_STRINGS:
                        $this->workbookRelationships[self::NS_SHARED_STRINGS] = $absoluteTarget;
                        break;
                    case self::NS_WORKSHEET:
                        $this->workbookRelationships[self::NS_WORKSHEET][$id] = $absoluteTarget;
                        break;
                }
            }

            $workbookRelsXml->close();
        }

        return $this->workbookRelationships;
    }

    /**
     * @param int $sheetIndex
     * @return array
     */
    private function getHyperlinkRelationshipsBySheetIndex(int $sheetIndex): array
    {
        $sheetXmlFilename = $this->workbookXmlRelationships[$sheetIndex] ?? null;
        $sheetXmlBasename = basename($sheetXmlFilename);
        $sheetXmlDirectory = dirname($sheetXmlFilename);
        $sheetRelsXmlFilename = $sheetXmlDirectory . '/_rels/' . $sheetXmlBasename . '.rels';

        if (!$this->zip->statName($sheetRelsXmlFilename)) {
            return $this->hyperlinkRelationships[$sheetIndex] ?? [];
        }

        if (!isset($this->hyperlinkRelationships[$sheetIndex])) {
            $sheetRelsXml = $this->getXMLReaderFromName($sheetRelsXmlFilename);
            while ($sheetRelsXml->read()) {
                if ($sheetRelsXml->getAttribute('Type') != self::NS_HYPERLINK) {
                    continue;
                }
                $id = $sheetRelsXml->getAttribute('Id');
                $this->hyperlinkRelationships[$sheetIndex][$id] = [
                    'Target' => (string) $sheetRelsXml->getAttribute('Target'),
                    'TargetMode' => (string) $sheetRelsXml->getAttribute('TargetMode'),
                ];
            }
        }

        return $this->hyperlinkRelationships[$sheetIndex] ?? [];
    }

    /**
     * Read color.
     *
     * @param \SimpleXMLElement $color
     * @param bool $background
     * @return string
     */
    private function readColor(SimpleXMLElement $color, bool $background = false): string
    {
        if (isset($color->attributes()->rgb)) {
            return (string) $color->attributes()->rgb;
        } elseif (isset($color->attributes()->indexed)) {
            $index = (int) (empty($this->palette) ? ($color->attributes()->indexed - 7) : $color->attributes()->indexed);
            return Style\Color::indexedColor($index, $background, $this->palette)->getARGB() ?? '';
        } elseif (isset($color->attributes()->theme)) {
            if ($this->theme !== null) {
                $returnColour = $this->theme->getColourByIndex((int) $color->attributes()->theme);
                if (isset($color->attributes()->tint)) {
                    $tintAdjust = (float) $color->attributes()->tint;
                    $returnColour = Style\Color::changeBrightness($returnColour, $tintAdjust);
                }

                return 'FF' . $returnColour;
            }
        }

        if ($background) {
            return 'FFFFFFFF';
        }

        return 'FF000000';
    }

    /**
     * Read border.
     *
     * @param \EasyExcel\Metadata\Style\Border $border
     * @param ?\SimpleXMLElement $borderXml
     * @return void
     */
    private function readBorder(Style\Border $border, ?SimpleXMLElement $borderXml): void
    {
        if ($borderXml && isset($borderXml->attributes()->style)) {
            $border->setStyle((string) $borderXml->attributes()->style);
        }
        if ($borderXml && isset($borderXml->color)) {
            $border->getColor()->setARGB(self::readColor($borderXml->color));
        }
    }

    /**
     * Get XMLReader from name.
     *
     * @param string $name
     * @param string|null $filename
     * @return \XMLReader
     */
    private function getXMLReaderFromName(string $name, ?string $filename = null): XMLReader
    {
        $xmlReader = new XMLReader();
        $filename = $filename ?: $this->zip->filename;
        $xmlReader->open(
            'zip://' . $filename . '#' . $name,
            null,
            Settings::getLibXmlLoaderOptions()
        );

        return $xmlReader;
    }

    /**
     * Close reader.
     *
     * @return void
     */
    protected function closeReader(): void
    {
        if ($this->zip) {
            $this->zip->close();
        }
    }
}
