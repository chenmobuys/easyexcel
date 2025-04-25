<?php

namespace EasyExcel\Readers\Xls;

use EasyExcel\Helpers\CodePage;
use EasyExcel\Helpers\Coordinate;
use EasyExcel\Helpers\Encoding;
use EasyExcel\Interfaces\ReaderExcel as ReaderExcelInterface;
use EasyExcel\Interfaces\ReaderRow as ReaderRowInterface;
use EasyExcel\Interfaces\ReaderSheet;
use EasyExcel\Metadata\Style;
use EasyExcel\Metadata\Style\Format;
use EasyExcel\Readers\BaseReader;
use EasyExcel\Readers\ReaderExcel;
use EasyExcel\Readers\Xls\Style\Border;
use EasyExcel\Readers\Xls\Style\Color;
use EasyExcel\Readers\Xls\Style\FillPattern;

/**
 * @see http://www.openoffice.org/sc/excelfileformat.pdf
 * @see http://www.openoffice.org/sc/compdocfileformat.pdf
 * @see https://en.wikipedia.org/wiki/Microsoft_Excel
 */
class Reader extends BaseReader
{
    // Record constant
    public const RECORD_BOF = 0x809;
    public const RECORD_EOF = 0x0a;
    public const RECORD_PRECISION = 0x0e;
    public const RECORD_CODEPAGE = 0x42;
    public const RECORD_DATEMODE = 0x22;
    public const RECORD_DEFINEDNAME = 0x18;
    public const RECORD_CALCCOUNT = 0x0c;
    public const RECORD_CALCMODE = 0x0d;
    public const RECORD_REFMODE = 0x0f;
    public const RECORD_DELTA = 0x10;
    public const RECORD_ITERATION = 0x11;
    public const RECORD_PROTECT = 0x12;
    public const RECORD_PASSWORD = 0x13;
    public const RECORD_HEADER = 0x14;
    public const RECORD_FOOTER = 0x15;
    public const RECORD_EXTERNSHEET = 0x17;
    public const RECORD_WINDOWPROTECT = 0x19;
    public const RECORD_VERTICALPAGEBREAKS = 0x1a;
    public const RECORD_HORIZONTALPAGEBREAKS = 0x1b;
    public const RECORD_NOTE = 0x1c;
    public const RECORD_SELECTION = 0x1d;
    public const RECORD_EXTERNALNAME = 0x23;
    public const RECORD_LEFTMARGIN = 0x26;
    public const RECORD_RIGHTMARGIN = 0x27;
    public const RECORD_TOPMARGIN = 0x28;
    public const RECORD_BOTTOMMARGIN = 0x29;
    public const RECORD_PRINTHEADERS = 0x2a;
    public const RECORD_PRINTGRIDLINES = 0x2b;
    public const RECORD_FILEPASS = 0x2f;
    public const RECORD_CONTINUE = 0x3c;
    public const RECORD_WINDOW1 = 0x3d;
    public const RECORD_BACKUP = 0x40;
    public const RECORD_PANE = 0x41;
    public const RECORD_PLS = 0x4d;
    public const RECORD_DCONREF = 0x51;
    public const RECORD_DEFCOLWIDTH = 0x55;
    public const RECORD_XCT = 0x59;
    public const RECORD_CRN = 0x5a;
    public const RECORD_FILESHARING = 0x5b;
    public const RECORD_WRITEACCESS = 0x5c;
    public const RECORD_UNCALCED = 0x5e;
    public const RECORD_SAVERECALC = 0x5f;
    public const RECORD_OBJECTPROTECT = 0x63;
    public const RECORD_COLINFO = 0x7d;
    public const RECORD_GUTS = 0x80;
    public const RECORD_SHEETPR = 0x81;
    public const RECORD_GRIDSET = 0x82;
    public const RECORD_HCENTER = 0x83;
    public const RECORD_VCENTER = 0x84;
    public const RECORD_WRITEPROT = 0x86;
    public const RECORD_COUNTRY = 0x8c;
    public const RECORD_HIDEOBJ = 0x8d;
    public const RECORD_SORT = 0x90;
    public const RECORD_STANDARDWIDTH = 0x99;
    public const RECORD_SCL = 0xa0;
    public const RECORD_PAGESETUP = 0xa1;
    public const RECORD_BOOKBOOL = 0xda;
    public const RECORD_SCENPROTECT = 0xdd;
    public const RECORD_BITMAP = 0xe9;
    public const RECORD_PHONETICPR = 0xef;
    public const RECORD_USESELFS = 0x160;
    public const RECORD_DSF = 0x161;
    public const RECORD_EXTERNALBOOK = 0x1ae;
    public const RECORD_CFHEADER = 0x1b0;
    public const RECORD_DATAVALIDATIONS = 0x1b2;
    public const RECORD_DATAVALIDATION = 0x1be;
    public const RECORD_DEFAULTROWHEIGHT = 0x225;
    public const RECORD_DATATABLE = 0x236;
    public const RECORD_DATATABLE2 = 0x37;
    public const RECORD_WINDOW2 = 0x23e;
    public const RECORD_SHEETLAYOUT = 0x862;
    public const RECORD_SHEETPROTECTION = 0x867;
    public const RECORD_RANGEPROTECTION = 0x868;
    public const RECORD_EXTSST = 0xff;
    public const RECORD_LABELRANGES = 0x15f;
    public const RECORD_INDEX = 0x20b;
    public const RECORD_ARRAY = 0x221;
    public const RECORD_STYLE = 0x293;
    public const RECORD_SHAREDFMLA = 0x4bc;
    public const RECORD_RSTRING = 0xd6;

    public const RECORD_XF = 0xe0;
    public const RECORD_PALETTE = 0x92;
    public const RECORD_FONT = 0x31;
    public const RECORD_FORMAT = 0x41e;
    public const RECORD_SST = 0xfc;
    public const RECORD_SHEET = 0x85;
    public const RECORD_ROW = 0x208;
    public const RECORD_DBCELL = 0xd7;
    public const RECORD_DIMENSION = 0x200;
    public const RECORD_HYPERLINK = 0x1b8;
    public const RECORD_QUICKTIP = 0x800;
    public const RECORD_MERGEDCELLS = 0xe5;
    public const RECORD_BLANK = 0x201;
    public const RECORD_MULBLANK = 0xbe;
    public const RECORD_RK = 0x27e;
    public const RECORD_MULRK = 0xbd;
    public const RECORD_NUMBER = 0x203;
    public const RECORD_LABEL = 0x204;
    public const RECORD_BOOLERR = 0x205;
    public const RECORD_FORMULA = 0x06;
    public const RECORD_STRING = 0x207;
    public const RECORD_LABELSST = 0xfd;
    public const RECORD_OBJ = 0x5d;

    // Binary interchange file format versions (BIFF)
    public const BIFF_VERSION_7 = 0x500;
    public const BIFF_VERSION_8 = 0x600;

    /**
     * @var \EasyExcel\Readers\Xls\OLEReader
     */
    protected $OLEReader;

    /**
     * Workbook stream
     *
     * @var string
     */
    protected $workbook;

    /**
     * @var int
     */
    protected $version;

    /**
     * @var string
     */
    protected $codepage;

    /**
     * @var float
     */
    protected $precision;

    /**
     * @var bool
     */
    protected $hasDbcell = false;

    /**
     * @var array
     */
    protected $sheets = [];

    /**
     * @var array
     */
    protected $rowOffsets = [];

    /**
     * @var array
     */
    protected $rowMergeCells = [];

    /**
     * @var array
     */
    protected $palette = [];

    /**
     * Fonts.
     *
     * @var array
     */
    protected $fonts = [];

    /**
     * Formats.
     *
     * @var array
     */
    protected $formats = [];

    /**
     * @var array
     */
    protected $formulas = [];

    /**
     * @var array
     */
    protected $styles = [];

    /**
     * Determine whether the file is readable.
     *
     * @param string $filename
     * @return bool
     */
    public function readable(string $filename): bool
    {
        return @file_get_contents($filename, false, null, 0, 8)
            == pack('CCCCCCCC', 0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1);
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
        $rowOffsets = $this->rowOffsets[$sheet->getIndex()] ?? [];

        return new ReaderRow(
            $this->workbook, $this->codepage, $this->version,
            $rowOffsets, $sheet, $startRow, $endRow
        );
    }

    /**
     * Load from file.
     *
     * @param string $filename
     * @return \EasyExcel\Interfaces\ReaderExcel
     */
    protected function loadFromFile(string $filename): ReaderExcelInterface
    {
        $excel = new ReaderExcel($this);

        $this->OLEReader = new OLEReader($filename);
        $this->workbook = $this->OLEReader->getWorkbook();

        $this->readSheetsAndStyles($excel);

        $this->readSheetInformation($excel);

        $this->readRowOffsets();

        return $excel;
    }

    /**
     * Read sheets and styles.
     *
     * @param \EasyExcel\Interfaces\ReaderExcel $excel
     * @return void
     */
    protected function readSheetsAndStyles(ReaderExcelInterface $excel): void
    {
        $currentXfIndex = 0;
        $position = 0;
        $code = Helper::getInt($position, $this->workbook, 2);
        $length = Helper::getInt($position + 2, $this->workbook, 2);
        $version = Helper::getInt($position + 4, $this->workbook, 2);
        $subStreamType = Helper::getInt($position + 6, $this->workbook, 2);
        $position += 4;
        // Set version
        $this->version = $version;

        do {
            $position += $length;
            $code = Helper::getInt($position, $this->workbook, 2);
            $length = Helper::getInt($position + 2, $this->workbook, 2);
            $position += 4;

            switch ($code) {
                case self::RECORD_CODEPAGE:
                    $codepageNumber = Helper::getInt($position, $this->workbook, 2);
                    $this->codepage = CodePage::numberToName($codepageNumber);
                    break;
                case self::RECORD_DATEMODE:
                    // 0 = Base date is 1899-Dec-31 (the cell value 1 represents 1900-Jan-01)
                    // 1 = Base date is 1904-Jan-01 (the cell value 1 represents 1904-Jan-02)
                    $dateMode = Helper::getInt($position, $this->workbook, 1);
                    $dateMode && Format::setCalendar(Format::CALENDAR_MAC_1904);
                    break;
                case self::RECORD_PRECISION:
                    // This record stores if formulas use the real cell values for calculation or the values displayed on the screen.
                    // 0 = Use displayed values; 1 = Use real cell values
                    $this->precision = Helper::getInt($position, $this->workbook, 2);
                    break;
                case self::RECORD_PALETTE:
                    // Palette
                    $paletteCount = Helper::getInt($position, $this->workbook, 2);
                    for ($i = 0; $i < $paletteCount; ++$i) {
                        $rgb = substr($this->workbook, $position + 2 + 4 * $i, 4);
                        $this->palette[] = Helper::readRGB($rgb);
                    }
                    break;
                case self::RECORD_FONT:
                    // Font
                    $font = new Style\Font();
                    // offset: 0; size: 2; height of the font (in twips = 1/20 of a point)
                    $size = Helper::getInt($position, $this->workbook, 2) / 20;
                    $font->setSize($size);
                    // offset: 2; size: 2; option flags
                    // bit: 0; mask 0x0001; bold (redundant in BIFF5-BIFF8)
                    // bit: 1; mask 0x0002; italic
                    $italic = (0x0002 & Helper::getInt($position + 2, $this->workbook, 2)) >> 1;
                    $font->setItalic((bool) $italic);
                    // bit: 2; mask 0x0004; underlined (redundant in BIFF5-BIFF8)
                    // bit: 3; mask 0x0008; strikethrough
                    $strike = (0x0008 & Helper::getInt($position + 2, $this->workbook, 2)) >> 3;
                    $font->setStrikethrough((bool) $strike);
                    // offset: 4; size: 2; colour index
                    $colorIndex = Helper::getInt($position + 4, $this->workbook, 2);
                    $font->getColor()->colorIndex = $colorIndex;
                    // offset: 6; size: 2; font weight
                    $weight = Helper::getInt($position + 6, $this->workbook, 2) === 0x02bc;
                    $font->setBold($weight);
                    // offset: 8; size: 2; escapement type
                    $escapement = Helper::getInt($position + 8, $this->workbook, 2);
                    $font->setSuperscript($escapement === 0x0001);
                    $font->setSubscript($escapement === 0x0002);
                    // offset: 10; size: 1; underline type
                    $underlineType = ord($this->workbook[$position + 10]);
                    switch ($underlineType) {
                        case 0x00:
                            break; // no underline
                        case 0x01:
                            $font->setUnderline(Style\Font::UNDERLINE_SINGLE);
                            break;
                        case 0x02:
                            $font->setUnderline(Style\Font::UNDERLINE_DOUBLE);
                            break;
                        case 0x21:
                            $font->setUnderline(Style\Font::UNDERLINE_SINGLEACCOUNTING);
                            break;
                        case 0x22:
                            $font->setUnderline(Style\Font::UNDERLINE_DOUBLEACCOUNTING);
                            break;
                    }

                    // offset: 11; size: 1; font family
                    // offset: 12; size: 1; character set
                    // offset: 13; size: 1; not used
                    // offset: 14; size: var; font name
                    $data = substr($this->workbook, $position + 14);
                    if ($version == self::BIFF_VERSION_8) {
                        $name = Helper::readUnicodeStringShort($data, $this->codepage);
                    } else {
                        $name = Helper::readByteStringShort($data, $this->codepage);
                    }
                    $font->setName($name);
                    $this->fonts[] = $font;
                    break;
                case self::RECORD_FORMAT:
                    $formatIndex = Helper::getInt($position, $this->workbook, 2);
                    if ($version == self::BIFF_VERSION_8) {
                        $formatCode = Helper::readUnicodeStringLong(substr($this->workbook, $position + 2),
                            $this->codepage);
                    } else {
                        $formatCode = Helper::readByteStringShort(substr($this->workbook, $position + 2),
                            $this->codepage);
                    }
                    $this->formats[$formatIndex] = $formatCode;
                    break;
                case self::RECORD_XF:
                    $style = new Style();
                    // offset:  0; size: 2; Index to FONT record
                    $fontIndex = Helper::getInt($position, $this->workbook, 2);
                    $fontIndex = $fontIndex >= 4 ? ($fontIndex - 1) : $fontIndex;
                    $style->setFont($this->fonts[$fontIndex]);

                    // offset:  2; size: 2; Index to FORMAT record
                    $formatIndex = Helper::getInt($position + 2, $this->workbook, 2);
                    if (isset($this->formats[$formatIndex])) {
                        $formatCode = $this->formats[$formatIndex];
                    } elseif (($code = Format::builtInFormatCode($formatIndex)) != '') {
                        $formatCode = $code;
                    } else {
                        $formatCode = Format::FORMAT_GENERAL;
                    }
                    $style->getFormat()->setFormatCode($formatCode);

                    // offset:  4; size: 2; XF type, cell protection, and parent style XF
                    // bit 2-0; mask 0x0007; XF_TYPE_PROT
                    $xfTypeProt = Helper::getInt($position + 4, $this->workbook, 2);

                    // bit 0; mask 0x01; 1 = cell is locked
                    $cellLocked = (0x01 & $xfTypeProt) >> 0;
                    $style->getProtection()->setLocked($cellLocked ? Style\Protection::PROTECTION_INHERIT : Style\Protection::PROTECTION_UNPROTECTED);

                    // bit 1; mask 0x02; 1 = FormulaParser is hidden
                    $formulaHidden = (0x02 & $xfTypeProt) >> 1;
                    $style->getProtection()->setHidden($formulaHidden ? Style\Protection::PROTECTION_PROTECTED : Style\Protection::PROTECTION_UNPROTECTED);

                    // bit 2; mask 0x04; 0 = Cell XF, 1 = Cell Style XF
                    $isCellStyleXf = (0x04 & $xfTypeProt) >> 2;

                    // offset: 6; size: 1; Alignment and text break
                    $alignmentAndTextBreakFlag = Helper::getInt($position + 6, $this->workbook, 1);
                    // bit 2-0, mask 0x07; horizontal alignment
                    $horAlign = (0x07 & $alignmentAndTextBreakFlag) >> 0;
                    switch ($horAlign) {
                        case 0:
                            $style->getAlignment()->setHorizontal(Style\Alignment::HORIZONTAL_GENERAL);
                            break;
                        case 1:
                            $style->getAlignment()->setHorizontal(Style\Alignment::HORIZONTAL_LEFT);
                            break;
                        case 2:
                            $style->getAlignment()->setHorizontal(Style\Alignment::HORIZONTAL_CENTER);
                            break;
                        case 3:
                            $style->getAlignment()->setHorizontal(Style\Alignment::HORIZONTAL_RIGHT);
                            break;
                        case 4:
                            $style->getAlignment()->setHorizontal(Style\Alignment::HORIZONTAL_FILL);
                            break;
                        case 5:
                            $style->getAlignment()->setHorizontal(Style\Alignment::HORIZONTAL_JUSTIFY);
                            break;
                        case 6:
                            $style->getAlignment()->setHorizontal(Style\Alignment::HORIZONTAL_CENTER_CONTINUOUS);
                            break;
                    }
                    // bit 3, mask 0x08; wrap text
                    $wrapText = (0x08 & $alignmentAndTextBreakFlag) >> 3;
                    $style->getAlignment()->setWrapText((bool) $wrapText);

                    // bit 6-4, mask 0x70; vertical alignment
                    $vertAlign = (0x70 & $alignmentAndTextBreakFlag) >> 4;
                    switch ($vertAlign) {
                        case 0:
                            $style->getAlignment()->setVertical(Style\Alignment::VERTICAL_TOP);
                            break;
                        case 1:
                            $style->getAlignment()->setVertical(Style\Alignment::VERTICAL_CENTER);
                            break;
                        case 2:
                            $style->getAlignment()->setVertical(Style\Alignment::VERTICAL_BOTTOM);
                            break;
                        case 3:
                            $style->getAlignment()->setVertical(Style\Alignment::VERTICAL_JUSTIFY);
                            break;
                    }

                    if ($this->version == self::BIFF_VERSION_8) {
                        // offset: 7; size: 1; XF_ROTATION: Text rotation angle
                        $xfRotation = Helper::getInt($position + 7, $this->workbook, 1);
                        $textRotation = 0;
                        if ($xfRotation <= 90) {
                            $textRotation = $xfRotation;
                        } elseif ($xfRotation <= 180) {
                            $textRotation = 90 - $xfRotation;
                        } elseif ($xfRotation == 255) {
                            $textRotation = -165;
                        }
                        $style->getAlignment()->setTextRotation($textRotation);

                        // offset: 8; size: 1; Indentation, shrink to cell size, and text direction
                        $indentation = Helper::getInt($position + 8, $this->workbook, 1);
                        // bit: 3-0; mask: 0x0F; indent level
                        $indent = (0x0F & $indentation) >> 0;
                        $style->getAlignment()->setIndent($indent);
                        // bit: 4; mask: 0x10; 1 = shrink content to fit into cell
                        $shrinkToFit = (0x10 & $indentation) >> 4;
                        $style->getAlignment()->setShrinkToFit((bool) $shrinkToFit);

                        // offset: 9; size: 1; XF_USED_ATTRIB: Flags for used attribute groups
                        // $xfUsedAttrib = Helper::getInt($position + 9, $this->workbook, 1);

                        // offset: 10; size: 4; Cell border lines and background area
                        $mixedFlags1 = Helper::getInt($position + 10, $this->workbook);

                        // bit: 3-0; mask: 0x0000000F; left style
                        $borderLeftStyleIndex = (0x0000000F & $mixedFlags1) >> 0;
                        $borderLeftStyle = Border::indexedBorder($borderLeftStyleIndex);
                        $style->getBorders()->getLeft()->setStyle($borderLeftStyle);

                        // bit: 7-4; mask: 0x000000F0; right style
                        $borderRightStyleIndex = (0x000000F0 & $mixedFlags1) >> 4;
                        $borderRightStyle = Border::indexedBorder($borderRightStyleIndex);
                        $style->getBorders()->getRight()->setStyle($borderRightStyle);

                        // bit: 11-8; mask: 0x00000F00; top style
                        $borderTopStyleIndex = (0x00000F00 & $mixedFlags1) >> 8;
                        $borderTopStyle = Border::indexedBorder($borderTopStyleIndex);
                        $style->getBorders()->getTop()->setStyle($borderTopStyle);

                        // bit: 15-12; mask: 0x0000F000; bottom style
                        $borderBottomStyleIndex = (0x0000F000 & $mixedFlags1) >> 12;
                        $borderBottomStyle = Border::indexedBorder($borderBottomStyleIndex);
                        $style->getBorders()->getBottom()->setStyle($borderBottomStyle);

                        // bit: 22-16; mask: 0x007F0000; left color
                        $borderLeftColorIndex = (0x007F0000 & $mixedFlags1) >> 16;
                        $style->getBorders()->getLeft()->getColor()->colorIndex = $borderLeftColorIndex;

                        // bit: 29-23; mask: 0x3F800000; right color
                        $borderRightColorIndex = (0x3F800000 & $mixedFlags1) >> 23;
                        $style->getBorders()->getRight()->getColor()->colorIndex = $borderRightColorIndex;

                        // bit: 30; mask: 0x40000000; 1 = diagonal line from top left to right bottom
                        // bit: 31; mask: 0x80000000; 1 = diagonal line from bottom left to top right
                        $diagonalDown = (bool) ((0x40000000 & $mixedFlags1) >> 30);
                        $diagonalUp = (bool) ((0x80000000 & $mixedFlags1) >> 31);
                        if (!$diagonalUp && !$diagonalDown) {
                            $style->getBorders()->setDiagonalDirection(Style\Borders::DIAGONAL_NONE);
                        } elseif ($diagonalUp && !$diagonalDown) {
                            $style->getBorders()->setDiagonalDirection(Style\Borders::DIAGONAL_UP);
                        } elseif (!$diagonalUp && $diagonalDown) {
                            $style->getBorders()->setDiagonalDirection(Style\Borders::DIAGONAL_DOWN);
                        } elseif ($diagonalUp && $diagonalDown) {
                            $style->getBorders()->setDiagonalDirection(Style\Borders::DIAGONAL_BOTH);
                        }

                        // offset: 14; size: 4; Cell border lines and background area
                        $mixedFlags2 = Helper::getInt($position + 14, $this->workbook);

                        // bit: 13-7; mask: 0x00003F80; bottom color
                        $borderTopColorIndex = (0x0000007F & $mixedFlags2) >> 0;
                        $style->getBorders()->getTop()->getColor()->colorIndex = $borderTopColorIndex;

                        $borderBottomColorIndex = (0x00003F80 & $mixedFlags2) >> 7;
                        $style->getBorders()->getBottom()->getColor()->colorIndex = $borderBottomColorIndex;

                        // bit: 20-14; mask: 0x001FC000; diagonal color
                        $borderDiagonalColorIndex = (0x001FC000 & $mixedFlags2) >> 14;
                        $style->getBorders()->getDiagonal()->getColor()->colorIndex = $borderDiagonalColorIndex;

                        // bit: 24-21; mask: 0x01E00000; diagonal style
                        $borderDiagonalStyleIndex = (0x01E00000 & $mixedFlags2) >> 21;
                        $borderDiagonalStyle = Border::indexedBorder($borderDiagonalStyleIndex);
                        $style->getBorders()->getDiagonal()->setStyle($borderDiagonalStyle);

                        // bit: 31-26; mask: 0xFC000000 fill pattern
                        $fillPatternIndex = (0xFC000000 & $mixedFlags2) >> 26;
                        $fillType = FillPattern::indexedFillPattern($fillPatternIndex);
                        $style->getFill()->setType($fillType);

                        // offset: 18; size: 2; pattern and background colour
                        $mixedFlags3 = Helper::getInt($position + 18, $this->workbook, 2);

                        // bit: 6-0; mask: 0x007F; color index for pattern color
                        $startColorIndex = (0x007F & $mixedFlags3) >> 0;
                        $style->getFill()->getStartColor()->colorIndex = $startColorIndex;

                        // bit: 13-7; mask: 0x3F80; color index for pattern background
                        $endColorIndex = (0x3F80 & $mixedFlags3) >> 7;
                        $style->getFill()->getEndColor()->colorIndex = $endColorIndex;
                    } else {
                        // BIFF5
                        // offset: 7; size: 1; Text orientation and flags
                        $orientationAndFlags = Helper::getInt($position + 7, $this->workbook, 1);

                        // bit: 1-0; mask: 0x03; XF_ORIENTATION: Text orientation
                        $xfOrientation = (0x03 & $orientationAndFlags) >> 0;
                        switch ($xfOrientation) {
                            case 0:
                                $style->getAlignment()->setTextRotation(0);
                                break;
                            case 1:
                                $style->getAlignment()->setTextRotation(-165);
                                break;
                            case 2:
                                $style->getAlignment()->setTextRotation(90);
                                break;
                            case 3:
                                $style->getAlignment()->setTextRotation(-90);
                                break;
                        }

                        // offset: 8; size: 4; cell border lines and background area
                        $borderAndBackground = Helper::getInt($position + 8, $this->workbook);

                        // bit: 6-0; mask: 0x0000007F; color index for pattern color
                        $style->getFill()->getStartColor()->colorIndex = (0x0000007F & $borderAndBackground) >> 0;

                        // bit: 13-7; mask: 0x00003F80; color index for pattern background
                        $style->getFill()->getEndColor()->colorIndex = (0x00003F80 & $borderAndBackground) >> 7;

                        // bit: 21-16; mask: 0x003F0000; fill pattern
                        $style->getFill()->setType(FillPattern::indexedFillPattern((0x003F0000 & $borderAndBackground) >> 16));

                        // bit: 24-22; mask: 0x01C00000; bottom line style
                        $style->getBorders()->getBottom()->setStyle(Border::indexedBorder((0x01C00000 & $borderAndBackground) >> 22));

                        // bit: 31-25; mask: 0xFE000000; bottom line color
                        $style->getBorders()->getBottom()->getColor()->colorIndex = (0xFE000000 & $borderAndBackground) >> 25;

                        // offset: 12; size: 4; cell border lines
                        $borderLines = Helper::getInt($position + 12, $this->workbook);

                        // bit: 2-0; mask: 0x00000007; top line style
                        $style->getBorders()->getTop()->setStyle(Border::indexedBorder((0x00000007 & $borderLines) >> 0));

                        // bit: 5-3; mask: 0x00000038; left line style
                        $style->getBorders()->getLeft()->setStyle(Border::indexedBorder((0x00000038 & $borderLines) >> 3));

                        // bit: 8-6; mask: 0x000001C0; right line style
                        $style->getBorders()->getRight()->setStyle(Border::indexedBorder((0x000001C0 & $borderLines) >> 6));

                        // bit: 15-9; mask: 0x0000FE00; top line color index
                        $style->getBorders()->getTop()->getColor()->colorIndex = (0x0000FE00 & $borderLines) >> 9;

                        // bit: 22-16; mask: 0x007F0000; left line color index
                        $style->getBorders()->getLeft()->getColor()->colorIndex = (0x007F0000 & $borderLines) >> 16;

                        // bit: 29-23; mask: 0x3F800000; right line color index
                        $style->getBorders()->getRight()->getColor()->colorIndex = (0x3F800000 & $borderLines) >> 23;
                    }

                    if (!$isCellStyleXf) {
                        $this->styles[$currentXfIndex] = $style;
                    }
                    $currentXfIndex++;
                    break;
                case self::RECORD_SST:
                    $sPosition = $position;
                    $limitPosition = $position + $length;
                    $uniqueStringsCount = Helper::getInt($sPosition + 4, $this->workbook);
                    $sPosition += 8;
                    for ($i = $uniqueStringsCount; $i--;) {
                        // Read in the number of characters
                        if ($sPosition == $limitPosition) {
                            $opcode = Helper::getInt($sPosition, $this->workbook, 2);
                            $conLength = Helper::getInt(0, $this->workbook, 2);
                            if ($opcode != 0x3c) {
                                break;
                            }
                            $sPosition += 4;
                            $limitPosition = $sPosition + $conLength;
                        }
                        $numChars = Helper::getInt($sPosition, $this->workbook, 2);
                        $sPosition += 2;
                        $optionFlags = Helper::getInt($sPosition, $this->workbook, 1);
                        $sPosition++;
                        $asciiEncoding = (($optionFlags & 0x01) == 0);
                        $extendedString = (($optionFlags & 0x04) != 0);
                        $richString = (($optionFlags & 0x08) != 0);

                        if ($richString) {
                            $formattingRuns = Helper::getInt($sPosition, $this->workbook, 2);
                            $sPosition += 2;
                        }
                        if ($extendedString) {
                            $extendedRunLength = Helper::getInt($sPosition, $this->workbook);
                            $sPosition += 4;
                        }

                        // echo $string;
                        $sLength = $asciiEncoding ? $numChars : $numChars * 2;
                        if ($sPosition + $sLength < $limitPosition) {
                            $string = substr($this->workbook, $sPosition, $sLength);
                            $sPosition += $sLength;
                        } else {
                            // found continue
                            $bytesRead = $limitPosition - $sPosition;
                            $string = substr($this->workbook, $sPosition, $bytesRead);
                            $charsLeft = $numChars - ($asciiEncoding ? $bytesRead : ($bytesRead / 2));
                            $sPosition = $limitPosition;

                            while ($charsLeft > 0) {
                                $opcode = Helper::getInt($sPosition, $this->workbook, 2);
                                $conLength = Helper::getInt($sPosition + 2, $this->workbook, 2);
                                if ($opcode != 0x3c) {
                                    break;
                                }
                                $sPosition += 4;
                                $limitPosition = $sPosition + $conLength;
                                $option = Helper::getInt($sPosition, $this->workbook, 1);
                                $sPosition += 1;

                                if ($asciiEncoding && ($option == 0)) {
                                    $sLength = min($charsLeft, $limitPosition - $sPosition);
                                    $string .= substr($this->workbook, $sPosition, $sLength);
                                    $charsLeft -= $sLength;
                                    $asciiEncoding = true;
                                } elseif (!$asciiEncoding && ($option != 0)) {
                                    $sLength = min($charsLeft * 2, $limitPosition - $sPosition);
                                    $string .= substr($this->workbook, $sPosition, $sLength);
                                    $charsLeft -= $sLength / 2;
                                    $asciiEncoding = false;
                                } elseif (!$asciiEncoding && ($option == 0)) {
                                    // Bummer - the string starts off as Unicode, but after the
                                    // continuation it is in straightforward ASCII encoding
                                    $sLength = min($charsLeft, $limitPosition - $sPosition);
                                    for ($j = 0; $j < $sLength; $j++) {
                                        $string .= substr($this->workbook, $sPosition + $j, 1) . chr(0);
                                    }
                                    $charsLeft -= $sLength;
                                    $asciiEncoding = false;
                                } else {
                                    $newString = '';
                                    for ($j = 0; $j < strlen($string); $j++) {
                                        $newString = $string[$j] . chr(0);
                                    }
                                    $string = $newString;
                                    $sLength = min($charsLeft * 2, $limitPosition - $sPosition);
                                    $string .= substr($this->workbook, $sPosition, $sLength);
                                    $charsLeft -= $sLength / 2;
                                    $asciiEncoding = false;
                                }

                                $sPosition += $sLength;
                            }
                        }

                        if ($richString) {
                            $sPosition += 4 * $formattingRuns;
                        }

                        // For extended strings, skip over the extended string data
                        if ($extendedString) {
                            $sPosition += $extendedRunLength;
                        }

                        $sharedString = $asciiEncoding ? $string : Encoding::convertEncoding($string, $this->codepage);
                        // Append shared string
                        $excel->addSharedString($sharedString);
                    }
                    break;
                case self::RECORD_SHEET:
                    // Absolute stream position of the BOF record of the sheet represented by this record.
                    // This field is never encrypted in protected files.
                    $offset = Helper::getInt($position, $this->workbook);

                    // 0 = Visible
                    // 1 = Hidden
                    // 2 = “Very hidden”. Can only be set programmatically, e.g. with a Visual Basic macro. It is not possible to make such a sheet visible via the user interface.
                    $state = Helper::getInt($position + 4, $this->workbook, 1);

                    // 00H = Worksheet
                    // 02H = Chart
                    // 06H = Visual Basic module
                    $type = Helper::getInt($position + 5, $this->workbook, 1);

                    // Unicode string, 8-bit string length.
                    // From BIFF8 on, strings are always stored using UTF-16LE text encoding.
                    // The character array is a sequence of 16-bit values.
                    // Additionally it is possible to use a compressed format, which omits the high bytes of all characters, if they are all zero.
                    if ($this->version == self::BIFF_VERSION_8) {
                        $name = Helper::readUnicodeStringShort(substr($this->workbook, $position + 6), $this->codepage);
                    } else {
                        $name = Helper::readByteStringShort(substr($this->workbook, $position + 6), $this->codepage);
                    }

                    $excel->addSheet($name);

                    $this->sheets[] = compact('name', 'type', 'offset', 'state');

                    break;
                default:
                    // Default.
                    break;
            }
        } while ($code != self::RECORD_EOF);

        // Colors
        foreach ($this->styles as $index => $style) {
            if (!is_null($colorIndex = $style->getFont()->getColor()->colorIndex)) {
                $rgb = Color::indexedColor($colorIndex, $this->palette, $this->version);
                $style->getFont()->getColor()->setRgb($rgb);
            }

            if (!is_null($colorIndex = $style->getFill()->getStartColor()->colorIndex)) {
                $rgb = Color::indexedColor($colorIndex, $this->palette, $this->version);
                $style->getFill()->getStartColor()->setRgb($rgb);
            }

            if (!is_null($colorIndex = $style->getFill()->getEndColor()->colorIndex)) {
                $rgb = Color::indexedColor($colorIndex, $this->palette, $this->version);
                $style->getFill()->getEndColor()->setRgb($rgb);
            }

            if (!is_null($colorIndex = $style->getBorders()->getLeft()->getColor()->colorIndex)) {
                $rgb = Color::indexedColor($colorIndex, $this->palette, $this->version);
                $style->getBorders()->getLeft()->getColor()->setRgb($rgb);
            }

            if (!is_null($colorIndex = $style->getBorders()->getRight()->getColor()->colorIndex)) {
                $rgb = Color::indexedColor($colorIndex, $this->palette, $this->version);
                $style->getBorders()->getRight()->getColor()->setRgb($rgb);
            }

            if (!is_null($colorIndex = $style->getBorders()->getTop()->getColor()->colorIndex)) {
                $rgb = Color::indexedColor($colorIndex, $this->palette, $this->version);
                $style->getBorders()->getTop()->getColor()->setRgb($rgb);
            }

            if (!is_null($colorIndex = $style->getBorders()->getBottom()->getColor()->colorIndex)) {
                $rgb = Color::indexedColor($colorIndex, $this->palette, $this->version);
                $style->getBorders()->getBottom()->getColor()->setRgb($rgb);
            }

            if (!is_null($colorIndex = $style->getBorders()->getDiagonal()->getColor()->colorIndex)) {
                $rgb = Color::indexedColor($colorIndex, $this->palette, $this->version);
                $style->getBorders()->getDiagonal()->getColor()->setRgb($rgb);
            }

            $excel->addCellXf($style, $index);
        }
    }

    /**
     * Read sheet info.
     *
     * @param \EasyExcel\Interfaces\ReaderExcel $excel
     * @return void
     */
    protected function readSheetInformation(ReaderExcelInterface $excel): void
    {
        foreach ($this->sheets as $index => $sheet) {
            $totalRows = 0;
            $totalColumns = 0;
            $tmpOffsets = [];
            $this->rowOffsets[$index] = [];
            $this->rowMergeCells[$index] = [];

            $position = $sheet['offset'];
            $length = Helper::getInt($position + 2, $this->workbook, 2);
            $position += 4;

            do {
                $position += $length;
                $lowCode = Helper::getInt($position, $this->workbook, 1);

                if ($lowCode == self::RECORD_EOF) {
                    break;
                }

                $code = Helper::getInt($position, $this->workbook, 2);
                $length = Helper::getInt($position + 2, $this->workbook, 2);
                $position += 4;

                switch ($code) {
                    case self::RECORD_DIMENSION:
                        // Index to first used row
                        $firstRowIndex = Helper::getInt($position, $this->workbook);
                        // Index to last used row, increased by 1
                        $lastRowIndex = Helper::getInt($position + 4, $this->workbook);
                        // Index to first used column
                        $firstColumnIndex = Helper::getInt($position + 8, $this->workbook, 2);
                        // Index to last used column, increased by 1
                        $lastColumnIndex = Helper::getInt($position + 10, $this->workbook, 2);

                        $totalRows = max($lastRowIndex, $totalRows);
                        $totalColumns = max($lastColumnIndex, $totalColumns);
                        break;
                    case self::RECORD_ROW:
                        $rowIndex = Helper::getInt($position, $this->workbook, 2);
                        $tmpOffsets[$rowIndex] = 0;
                        break;
                    case self::RECORD_DBCELL:
                        $relativeOffset = Helper::getInt($position, $this->workbook);
                        $firstRowOffset = $position - $relativeOffset + 0x14;
                        $sumOffset = 0;
                        $rowI = 0;
                        foreach ($tmpOffsets as $rowIndex => $tmpOffset) {
                            $sumOffset += Helper::getInt($position + ($rowI % 32) * 2 + 4, $this->workbook, 2);
                            $this->rowOffsets[$index][$rowIndex] = $firstRowOffset + $sumOffset - 4;
                            $rowI++;
                        }
                        $tmpOffsets = [];
                        $this->hasDbcell = true;
                        break;
                    case self::RECORD_MERGEDCELLS:
                        $mergeCells = [];
                        $rowMergeCells = [];
                        $cellRanges = Helper::getInt($position, $this->workbook, 2);
                        for ($i = 0; $i < $cellRanges; $i++) {
                            $firstRowIndex = Helper::getInt($position + 8 * $i + 2, $this->workbook, 2);
                            $lastRowIndex = Helper::getInt($position + 8 * $i + 4, $this->workbook, 2);
                            $firstColumnIndex = Helper::getInt($position + 8 * $i + 6, $this->workbook, 2);
                            $lastColumnIndex = Helper::getInt($position + 8 * $i + 8, $this->workbook, 2);

                            for ($j = $firstRowIndex; $j <= $lastRowIndex; $j++) {
                                $rowMergeCells[$j][] = [
                                    $firstRowIndex, $lastRowIndex,
                                    $firstColumnIndex, $lastColumnIndex,
                                ];
                            }
                            $firstCoordinate = Coordinate::columnLetterFromColumnIndex($firstColumnIndex) . ($firstRowIndex + 1);
                            $lastCoordinate = Coordinate::columnLetterFromColumnIndex($lastColumnIndex) . ($lastRowIndex + 1);
                            $mergeCells[] = $firstCoordinate . ':' . $lastCoordinate;
                        }
                        ksort($mergeCells);
                        ksort($rowMergeCells);
                        $excel->getSheetByIndex($index)->setMergeCells($mergeCells);
                        $this->rowMergeCells[$index] = $rowMergeCells;
                        break;
                    case self::RECORD_HYPERLINK:
                        $firstRowIndex = Helper::getInt($position, $this->workbook, 2);
                        $lastRowIndex = Helper::getInt($position + 2, $this->workbook, 2);
                        $firstColumnIndex = Helper::getInt($position + 4, $this->workbook, 2);
                        $lastColumnIndex = Helper::getInt($position + 6, $this->workbook, 2);
                        $flags = Helper::getInt($position + 28, $this->workbook, 1);

                        // offset: 28, size: 4; option flags
                        // bit: 0; mask: 0x00000001; 0 = no link or extant, 1 = file link or URL
                        $isFileLinkOrUrl = (0x00000001 & $flags) >> 0;

                        // bit: 1; mask: 0x00000002; 0 = relative path, 1 = absolute path or URL
                        $isAbsPathOrUrl = (0x00000001 & $flags) >> 1;

                        // bit: 2 (and 4); mask: 0x00000014; 0 = no description
                        $hasDesc = (0x00000014 & $flags) >> 2;

                        // bit: 3; mask: 0x00000008; 0 = no text, 1 = has text
                        $hasText = (0x00000008 & $flags) >> 3;

                        // bit: 7; mask: 0x00000080; 0 = no target frame, 1 = has target frame
                        $hasFrame = (0x00000080 & $flags) >> 7;

                        // bit: 8; mask: 0x00000100; 0 = file link or URL, 1 = UNC path (inc. server name)
                        $isUNC = (0x00000100 & $flags) >> 8;

                        // offset within record data
                        $offset = 32;

                        if ($hasDesc) {
                            // offset: 32; size: var; character count of description text
                            $dl = Helper::getInt($position + $offset, $this->workbook);
                            // offset: 36; size: var; character array of description text, no Unicode string header, always 16-bit characters, zero terminated
                            $desc = Encoding::convertEncoding(substr($this->workbook, $position + $offset + 4,
                                2 * ($dl - 1)), $this->codepage);
                            $offset += 4 + 2 * $dl;
                        }
                        if ($hasFrame) {
                            $fl = Helper::getInt($position + $offset, $this->workbook);
                            $offset += 4 + 2 * $fl;
                        }

                        // detect type of hyperlink (there are 4 types)
                        $hyperlinkType = null;

                        if ($isUNC) {
                            $hyperlinkType = 'UNC';
                        } elseif (!$isFileLinkOrUrl) {
                            $hyperlinkType = 'workbook';
                        } elseif (Helper::getInt($position + $offset, $this->workbook, 1) == 0x03) {
                            $hyperlinkType = 'local';
                        } elseif (Helper::getInt($position + $offset, $this->workbook, 1) == 0xE0) {
                            $hyperlinkType = 'URL';
                        }

                        switch ($hyperlinkType) {
                            case 'URL':
                                // offset: var; size: 16; GUID of URL Moniker
                                $offset += 16;
                                // offset: var; size: 4; size (in bytes) of character array of the URL including trailing zero word
                                $us = Helper::getInt($position + $offset, $this->workbook);
                                $offset += 4;
                                // offset: var; size: $us; character array of the URL, no Unicode string header, always 16-bit characters, zero-terminated
                                $url = Encoding::convertEncoding(substr($this->workbook, $position + $offset,
                                    $us - 2),
                                    $this->codepage);
                                $nullOffset = strpos($url, chr(0x00));
                                if ($nullOffset) {
                                    $url = substr($url, 0, $nullOffset);
                                }
                                $url .= $hasText ? '#' : '';
                                $offset += $us;
                                break;
                            case 'local':
                                // offset: var; size: 16; GUI of File Moniker
                                $offset += 16;

                                // offset: var; size: 2; directory up-level count.
                                $upLevelCount = Helper::getInt($position + $offset, $this->workbook, 2);
                                $offset += 2;

                                // offset: var; size: 4; character count of the shortened file path and name, including trailing zero word
                                $sl = Helper::getInt($position + $offset, $this->workbook);
                                $offset += 4;

                                // offset: var; size: sl; character array of the shortened file path and name in 8.3-DOS-format (compressed Unicode string)
                                $shortenedFilePath = substr($this->workbook, $position + $offset, $sl);
                                $shortenedFilePath = Encoding::convertEncoding($shortenedFilePath, $this->codepage);
                                $shortenedFilePath = substr($shortenedFilePath, 0, -1); // remove trailing zero

                                $offset += $sl;

                                // offset: var; size: 24; unknown sequence
                                $offset += 24;

                                // extended file path
                                // offset: var; size: 4; size of the following file link field including string length mark
                                $sz = Helper::getInt($position + $offset, $this->workbook);
                                $offset += 4;

                                // only present if $sz > 0
                                if ($sz > 0) {
                                    // offset: var; size: 4; size of the character array of the extended file path and name
                                    $xl = Helper::getInt($position + $offset, $this->workbook);
                                    $offset += 4;

                                    // offset: var; size 2; unknown
                                    $offset += 2;

                                    // offset: var; size $xl; character array of the extended file path and name.
                                    $extendedFilePath = substr($this->workbook, $position + $offset, $xl);
                                    $extendedFilePath = Encoding::convertEncoding($extendedFilePath,
                                        $this->codepage);
                                    $offset += $xl;
                                }

                                // construct the path
                                $url = str_repeat('..\\', $upLevelCount);
                                $url .= $extendedFilePath ?? $shortenedFilePath; // use extended path if available
                                $url .= $hasText ? '#' : '';
                                break;
                            case 'workbook':
                                $url = 'sheet://';
                                break;
                            case 'UNC':
                            default:
                                break;
                        }

                        if ($hasText && isset($url)) {
                            // offset: var; size: 4; character count of text mark including trailing zero word
                            $tl = Helper::getInt($position + $offset, $this->workbook);
                            $offset += 4;
                            // offset: var; size: var; character array of the text mark without the # sign, no Unicode header, always 16-bit characters, zero-terminated
                            $text = Encoding::convertEncoding(substr($this->workbook, $position + $offset,
                                2 * ($tl - 1)),
                                $this->codepage);
                            $url .= $text;
                        }
                        if (isset($url)) {
                            for ($rowIndex = $firstRowIndex; $rowIndex <= $lastRowIndex; $rowIndex++) {
                                for ($columnIndex = $firstColumnIndex; $columnIndex <= $lastColumnIndex; $columnIndex++) {
                                    $coordinate = Coordinate::columnLetterFromColumnIndex($columnIndex) . ($rowIndex + 1);
                                    $excel->getSheetByIndex($index)->getHyperlink($coordinate)->setUrl($url);
                                }
                            }
                        }
                        break;
                    case self::RECORD_QUICKTIP:
                        $repeatedRecordIdentifier = Helper::getInt($position, $this->workbook, 2);
                        // In several cases, BIFF8 still writes the BIFF2-BIFF5 format of a cell range address.
                        $firstRowIndex = Helper::getInt($position + 2, $this->workbook, 2);
                        $lastRowIndex = Helper::getInt($position + 4, $this->workbook, 2);
                        $firstColumnIndex = Helper::getInt($position + 6, $this->workbook, 1);
                        $lastColumnIndex = Helper::getInt($position + 7, $this->workbook, 1);
                        $byteString = substr($this->workbook, $position + 8, $length - 10);
                        $tooltip = Helper::readByteStringLong($byteString, $this->codepage);

                        for ($rowIndex = $firstRowIndex; $rowIndex <= $lastRowIndex; $rowIndex++) {
                            for ($columnIndex = $firstColumnIndex; $columnIndex <= $lastColumnIndex; $columnIndex++) {
                                $coordinate = Coordinate::columnLetterFromColumnIndex($columnIndex) . ($rowIndex + 1);
                                $excel->getSheetByIndex($index)->getHyperlink($coordinate)->setTooltip($tooltip);
                            }
                        }
                        break;
                }
            } while ($code != self::RECORD_EOF);

            $excel->getSheetByIndex($index)
                ->setTotalRows($totalRows)
                ->setTotalColumns($totalColumns);
        }
    }

    /**
     * Read row offsets
     *
     * If does not have dbcell, need read row offsets.
     *
     * @return void
     */
    protected function readRowOffsets(): void
    {
        // If it has dbcell
        if ($this->hasDbcell) {
            return;
        }

        foreach ($this->sheets as $index => $sheet) {
            $currentPosition = $sheet['offset'];
            $code = Helper::getInt($currentPosition, $this->workbook, 2);
            $currentLength = Helper::getInt($currentPosition + 2, $this->workbook, 2);
            $currentPosition += 4;

            do {
                $position = $currentPosition + $currentLength;
                $lowCode = Helper::getInt($position, $this->workbook, 1);

                if ($lowCode == self::RECORD_EOF) {
                    break;
                }

                $code = Helper::getInt($position, $this->workbook, 2);
                $length = Helper::getInt($position + 2, $this->workbook, 2);
                $position += 4;

                switch ($code) {
                    case Reader:: RECORD_BLANK:
                    case Reader:: RECORD_MULBLANK:
                    case Reader:: RECORD_RK:
                    case Reader:: RECORD_MULRK:
                    case Reader:: RECORD_NUMBER:
                    case Reader:: RECORD_BOOLERR:
                    case Reader:: RECORD_LABEL:
                    case Reader:: RECORD_LABELSST:
                    case Reader:: RECORD_FORMULA:
                        $rowIndex = Helper::getInt($position, $this->workbook, 2);
                        $columnIndex = Helper::getInt($position + 2, $this->workbook, 2);
                        if ($columnIndex == 0) {
                            $this->rowOffsets[$index][$rowIndex] = $position - 4;
                        }
                        break;
                }
                $currentPosition = $position;
                $currentLength = $length;
            } while ($code != self::RECORD_EOF);
        }
    }

    /**
     * Close reader.
     *
     * @return void
     */
    protected function closeReader(): void
    {
        $this->OLEReader = null;
        $this->workbook = null;
    }
}