<?php

namespace EasyExcel\Readers\Xls;

class OLEReader
{
    // Stream constant
    public const SECTOR_SIZE_POSITION = 0x1e;
    public const SHORT_SECTOR_SIZE_POSITION = 0x20;
    public const SECTOR_COUNT_POSITION = 0x2c;
    public const DIRECTORY_FIRST_POSITION = 0x30;
    public const STANDARD_STREAM_MIN_SIZE_POSITION = 0x38;
    public const SHORT_SECTOR_FIRST_POSITION = 0x3c;
    public const SHORT_SECTOR_COUNT_POSITION = 0x40;
    public const MASTER_SECTOR_FIRST_POSITION = 0x44;
    public const MASTER_SECTOR_COUNT_POSITION = 0x48;
    public const MASTER_SECTOR_POSITION = 0x4c;
    public const DIRECTORY_PROPERTY_LENGTH = 0x80;
    public const DIRECTORY_PROPERTY_NAME_SIZE_POSITION = 0x40;
    public const DIRECTORY_PROPERTY_TYPE_POSITION = 0x42;
    public const DIRECTORY_PROPERTY_FIRST_POSITION = 0x74;
    public const DIRECTORY_PROPERTY_SIZE = 0x78;

    /**
     * Sector size
     *
     * @var int
     */
    private $sectorSize;

    /**
     * Sector chains
     *
     * @var array
     */
    private $sectorChains = [];

    /**
     * Short sector size
     *
     * @var int
     */
    private $shortSectorSize;

    /**
     * Short sector chains
     *
     * @var array
     */
    private $shortSectorChains = [];

    /**
     * Directory stream properties
     *
     * @var array
     */
    private $directoryProperties = [];

    /**
     * Standard stream min size
     *
     * @var int
     */
    private $standardStreamMinSize;

    /**
     * Spreadsheet stream
     *
     * @var string
     */
    private $spreadsheet;

    /**
     * Workbook stream
     *
     * @var string
     */
    private $workbook;

    /**
     * @param string $filename
     */
    public function __construct(string $filename)
    {
        $this->spreadsheet = file_get_contents($filename);

        $this->readSectorChains();

        $this->readDirectoryStream();
    }

    /**
     * Get workbook stream.
     *
     * @return string
     */
    public function getWorkbook(): string
    {
        if (is_null($this->workbook)) {
            $this->workbook = $this->getStreamByName('WORKBOOK');
        }

        return $this->workbook;
    }

    /**
     * Get stream by name.
     *
     * @param string $name
     * @return string
     */
    public function getStreamByName(string $name): string
    {
        $streamProperty = $this->getDirectoryPropertyByName($name);
        if (($streamProperty['size'] < $this->standardStreamMinSize)) {
            $rootEntryProperty = $this->getDirectoryPropertyByName('ROOT ENTRY');
            $rootEntryStreamData = $this->getStreamData($rootEntryProperty['secId']);
            return $this->getShortStreamData($streamProperty['secId'], $rootEntryStreamData);
        }

        return $this->getStreamData($streamProperty['secId']);
    }

    /**
     * Read sector chains.
     *
     * @return void
     */
    private function readSectorChains(): void
    {
        $sectorSizeExponent = Helper::getInt(self::SECTOR_SIZE_POSITION, $this->spreadsheet, 2);
        $sectorSize = pow(2, $sectorSizeExponent);

        $shortSectorSizeExponent = Helper::getInt(self::SHORT_SECTOR_SIZE_POSITION, $this->spreadsheet, 2);
        $shortSectorSize = pow(2, $shortSectorSizeExponent);

        $sectorCountOrigin = Helper::getInt(self::SECTOR_COUNT_POSITION, $this->spreadsheet);

        $standardStreamMinSize = Helper::getInt(self::STANDARD_STREAM_MIN_SIZE_POSITION, $this->spreadsheet);

        $shortSectorFirst = Helper::getInt(self::SHORT_SECTOR_FIRST_POSITION, $this->spreadsheet);

        // $shortSectorCount = Helper::getInt(self::SHORT_SECTOR_COUNT_POSITION, $this->spreadsheet);

        $masterSectorFirst = Helper::getInt(self::MASTER_SECTOR_FIRST_POSITION, $this->spreadsheet);

        $masterSectorCount = Helper::getInt(self::MASTER_SECTOR_COUNT_POSITION, $this->spreadsheet);

        $sectorCount = $sectorCountOrigin;
        if ($masterSectorCount != 0) {
            $sectorCount = ($sectorSize - self::MASTER_SECTOR_POSITION) / 4;
        }

        $sectorSecIds = [];
        $position = self::MASTER_SECTOR_POSITION;
        for ($i = 0; $i < $sectorCount; $i++) {
            $sectorSecIds[$i] = Helper::getInt($position, $this->spreadsheet);
            $position += 4;
        }

        for ($j = 0; $j < $masterSectorCount; $j++) {
            $position = ($masterSectorFirst + 1) * $sectorSize;
            $blocksToRead = min($sectorCountOrigin - $sectorCount, $sectorSize / 4 - 1);

            for ($i = $sectorCount; $i < $sectorCount + $blocksToRead; $i++) {
                $sectorSecIds[$i] = Helper::getInt($position, $this->spreadsheet);
                $position += 4;
            }

            $sectorCount += $blocksToRead;
            if ($sectorCount < $sectorCountOrigin) {
                $masterSectorFirst = Helper::getInt($position, $this->spreadsheet);
            }
        }

        // Read sector chains
        $sectorChains = [];
        for ($i = 0; $i < $sectorCount; $i++) {
            $position = ($sectorSecIds[$i] + 1) * $sectorSize;
            for ($j = 0; $j < $sectorSize / 4; $j++) {
                $sectorChains[] = Helper::getInt($position, $this->spreadsheet);
                $position += 4;
            }
        }

        // Read short-sector chains
        $shortSectorChains = [];
        $shortSectorFirstSecId = $shortSectorFirst;
        while ($shortSectorFirstSecId != -2) {
            $position = ($shortSectorFirstSecId + 1) * $sectorSize;
            for ($j = 0; $j < $sectorSize / 4; $j++) {
                $shortSectorChains[] = Helper::getInt($position, $this->spreadsheet);
                $position += 4;
            }
            $shortSectorFirstSecId = $sectorChains[$shortSectorFirstSecId];
        }

        $this->sectorSize = $sectorSize;
        $this->sectorChains = $sectorChains;
        $this->shortSectorSize = $shortSectorSize;
        $this->shortSectorChains = $shortSectorChains;
        $this->standardStreamMinSize = $standardStreamMinSize;
    }

    /**
     * Read directory stream.
     *
     * @return void
     */
    private function readDirectoryStream(): void
    {
        $position = 0;
        $firstSecId = Helper::getInt(self::DIRECTORY_FIRST_POSITION, $this->spreadsheet);
        $directoryStreamData = $this->getStreamData($firstSecId);
        $directoryStreamLength = strlen($directoryStreamData);

        while ($position < $directoryStreamLength) {
            $propertyBinary = substr($directoryStreamData, $position, self::DIRECTORY_PROPERTY_LENGTH);
            $nameSize = Helper::getInt(self::DIRECTORY_PROPERTY_NAME_SIZE_POSITION, $propertyBinary, 2);
            $type = Helper::getInt(self::DIRECTORY_PROPERTY_TYPE_POSITION, $propertyBinary, 1);
            $secId = Helper::getInt(self::DIRECTORY_PROPERTY_FIRST_POSITION, $propertyBinary);
            $size = Helper::getInt(self::DIRECTORY_PROPERTY_SIZE, $propertyBinary);
            $name = str_replace("\x00", '', substr($propertyBinary, 0, $nameSize));
            $this->directoryProperties[] = [
                'name' => $name,
                'type' => $type,
                'secId' => $secId,
                'size' => $size,
            ];

            $position += self::DIRECTORY_PROPERTY_LENGTH;
        }
    }

    /**
     * Get stream by sector id.
     *
     * @param int $secId
     * @return string
     */
    private function getStreamData(int $secId): string
    {
        $data = '';
        $sourceData = $this->spreadsheet;
        while ($secId != -2) {
            $position = ($secId + 1) * $this->sectorSize;
            $data .= substr($sourceData, $position, $this->sectorSize);
            $secId = $this->sectorChains[$secId];
        }
        return $data;
    }

    /**
     * Get short stream by sector id.
     *
     * @param int $secId
     * @param string $sourceData
     * @return string
     */
    private function getShortStreamData(int $secId, string $sourceData): string
    {
        $data = '';
        while ($secId != -2) {
            $position = $secId * $this->shortSectorSize;
            $data .= substr($sourceData, $position, $this->shortSectorSize);
            $secId = $this->shortSectorChains[$secId];
        }
        return $data;
    }

    /**
     * Get directory property by name.
     *
     * @param string $name
     * @return array
     */
    private function getDirectoryPropertyByName(string $name): array
    {
        return current(array_filter(
            $this->directoryProperties,
            function ($directoryProperty) use ($name) {
                $name = strtolower($name);
                return strtolower($directoryProperty['name']) == $name
                    || (strtolower($directoryProperty['name']) == 'book' && $name == 'workbook');
            }
        ));
    }
}