<?php

namespace EasyExcel\Readers\Ods;

use EasyExcel\Helpers\Coordinate;
use EasyExcel\Interfaces\ReaderExcel as ReaderExcelInterface;
use EasyExcel\Interfaces\ReaderRow as ReaderRowInterface;
use EasyExcel\Interfaces\ReaderSheet;
use EasyExcel\Readers\BaseReader;
use EasyExcel\Readers\ReaderExcel;
use EasyExcel\Settings;
use Throwable;
use XMLReader;
use ZipArchive;

/**
 * @see http://docs.oasis-open.org/office/v1.2/os/OpenDocument-v1.2-os-part1.html
 */
class Reader extends BaseReader
{
    // File
    public const FILE_CONTENT = 'content.xml';
    public const FILE_MANIFEST = 'META-INF/manifest.xml';

    // Namespace
    public const NS_MANIFEST = 'urn:oasis:names:tc:opendocument:xmlns:manifest:1.0';
    public const NS_OFFICE = 'urn:oasis:names:tc:opendocument:xmlns:office:1.0';
    public const NS_TABLE = 'urn:oasis:names:tc:opendocument:xmlns:table:1.0';
    public const NS_STYLE = 'urn:oasis:names:tc:opendocument:xmlns:style:1.0';
    public const NS_TEXT = 'urn:oasis:names:tc:opendocument:xmlns:text:1.0';
    public const NS_LINK = 'http://www.w3.org/1999/xlink';
    public const NS_FO = 'urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0';

    // Element
    public const EL_TABLE = 'table';
    public const EL_SPREADSHEET = 'spreadsheet';
    public const EL_TABLE_ROW = 'table-row';
    public const EL_TABLE_CELL = 'table-cell';
    public const EL_COVERED_TABLE_CELL = 'covered-table-cell';
    public const EL_TABLE_CELL_P = 'text:p';
    public const EL_TABLE_COLUMN = 'table-column';
    public const EL_AUTOMATIC_STYLES = 'automatic-styles';
    public const EL_TEXT_PROPERTIES = 'text-properties';
    public const EL_FILE_ENTRY = 'file-entry';
    public const EL_STYLE = 'style';

    /**
     * @var ?\ZipArchive
     */
    protected $zip;

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

        $mimeType = null;
        $zip = new ZipArchive();

        if ($zip->open($filename) === true) {
            $stat = $zip->statName('mimetype');
            if ($stat && ($stat['size'] <= 255)) {
                $mimeType = $zip->getFromName($stat['name']);
            } elseif ($zip->statName(self::FILE_MANIFEST)) {
                try {
                    $manifestXml = $this->getXMLReaderFromName(self::FILE_MANIFEST, $filename);
                    while ($manifestXml->read()) {
                        if (
                            $manifestXml->localName == self::EL_FILE_ENTRY
                            && $manifestXml->nodeType == XMLReader::ELEMENT
                            && $manifestXml->getAttributeNs('full-path', self::NS_MANIFEST) == '/'
                        ) {
                            $mimeType = $manifestXml->getAttributeNs('media-type', self::NS_MANIFEST);
                            break;
                        }
                    }
                } catch (Throwable $e) {
                    // Do nothing.
                } finally {
                    if (isset($manifestXml)) {
                        $manifestXml->close();
                    }
                }
            }
            $zip->close();
        }

        return $mimeType == 'application/vnd.oasis.opendocument.spreadsheet';
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
        $handler = 'zip://' . $this->zip->filename . '#' . self::FILE_CONTENT;

        return new ReaderRow($handler, $sheet, $startRow, $endRow);
    }

    /**
     * Load from file.
     *
     * @param string $filename
     * @return \EasyExcel\Interfaces\ReaderExcel
     * @throws \EasyExcel\Exceptions\SheetNameExistsException
     */
    protected function loadFromFile(string $filename): ReaderExcelInterface
    {
        $excel = new ReaderExcel($this);

        $this->zip = $zip = new ZipArchive();
        $zip->open($filename);

        $xml = $this->getXMLReaderFromName(self::FILE_CONTENT);
        while ($xml->read()) {
            // Tables
            if ($xml->localName == self::EL_TABLE) {
                do {
                    $sheetName = (string) $xml->getAttributeNs('name', self::NS_TABLE);
                    $sheet = $excel->addSheet($sheetName);

                    if ($xml->isEmptyElement) {
                        continue;
                    }

                    while ($xml->read()) {
                        if ($xml->localName == self::EL_TABLE_ROW) {
                            break;
                        }
                    }

                    $totalRows = 0;
                    $totalColumns = 0;

                    // Rows
                    do {
                        $rowColumns = 0;
                        $numberRowsRepeated = ((int) $xml->getAttributeNs('number-rows-repeated', self::NS_TABLE)) ?: 1;
                        if (!$xml->isEmptyElement) {
                            $xml->read();
                            do {
                                $numberColumnsRepeated = ((int) $xml->getAttributeNS('number-columns-repeated', self::NS_TABLE)) ?: 1;
                                if (!$xml->isEmptyElement) {
                                    while ($xml->read()) {
                                        if ($xml->localName == self::EL_TABLE_CELL) {
                                            break;
                                        }
                                        if ($xml->localName == 'a' && $xml->nodeType == XMLReader::ELEMENT) {

                                            for ($i = 1; $i <= $numberColumnsRepeated; $i++) {
                                                $coordinate = Coordinate::columnLetterFromColumnIndex($rowColumns) . ($totalRows + $i);
                                                $url = (string) $xml->getAttributeNs('href', self::NS_LINK);
                                                $tooltip = $xml->readString();
                                                if (strpos($url, '#') === 0) {
                                                    $url = 'sheet://' . substr($url, 1);
                                                }
                                                $sheet->getHyperlink($coordinate)->setUrl($url)->setTooltip($tooltip);
                                            }
                                        }
                                    }
                                }
                                $rowColumns += $numberColumnsRepeated;

                            } while ($xml->next() && $xml->localName == self::EL_TABLE_CELL);
                        }
                        $totalRows += $numberRowsRepeated;
                        $totalColumns = ($totalColumns > $rowColumns) ? $totalColumns : $rowColumns;
                    } while ($xml->next() && $xml->localName == self::EL_TABLE_ROW);

                    $sheet->setTotalRows($totalRows)->setTotalColumns($totalColumns);

                    do {
                        if($xml->localName == 'table') {
                            break;
                        }
                    } while($xml->read());

                } while ($xml->next() && $xml->localName == self::EL_TABLE);
            }
        }

        $xml->close();

        return $excel;
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