<?php

namespace EasyExcel\Readers\Xlsx;

class Theme
{
    /**
     * Theme Name.
     *
     * @var string
     */
    protected $themeName;

    /**
     * Colour Scheme Name.
     *
     * @var string
     */
    protected $colourSchemeName;

    /**
     * Colour Map.
     *
     * @var string[]
     */
    protected $colourMap;

    /**
     * Create a new Theme.
     *
     * @param string $themeName
     * @param string $colourSchemeName
     * @param string[] $colourMap
     */
    public function __construct(string $themeName, string $colourSchemeName, array $colourMap)
    {
        // Initialise values
        $this->themeName = $themeName;
        $this->colourSchemeName = $colourSchemeName;
        $this->colourMap = $colourMap;
    }

    /**
     * Get Theme Name.
     *
     * @return string
     */
    public function getThemeName(): string
    {
        return $this->themeName;
    }

    /**
     * Get colour Scheme Name.
     *
     * @return string
     */
    public function getColourSchemeName(): string
    {
        return $this->colourSchemeName;
    }

    /**
     * Get colour Map Value by Position.
     *
     * @param int $index
     *
     * @return ?string
     */
    public function getColourByIndex(int $index): ?string
    {
        return $this->colourMap[$index] ?? null;
    }
}
