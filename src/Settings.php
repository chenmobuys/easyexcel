<?php

namespace EasyExcel;

class Settings
{
    /**
     * @var string
     */
    protected static $tempDir;

    /**
     * Default options for libxml loader.
     *
     * @var int
     */
    protected static $libXmlLoaderOptions;

    /**
     * @return string
     */
    public static function getTempDir(): string
    {
        if (!is_dir(self::$tempDir)) {
            self::$tempDir = sys_get_temp_dir();
        }

        return self::$tempDir;
    }

    /**
     * @param string $tempDir
     * @return void
     */
    public static function setTempDir(string $tempDir): void
    {
        self::$tempDir = $tempDir;
    }

    /**
     * Get default options for libxml loader.
     * Defaults to LIBXML_DTDLOAD | LIBXML_DTDATTR when not set explicitly.
     *
     * @return int Default options for libxml loader
     */
    public static function getLibXmlLoaderOptions(): int
    {
        if (self::$libXmlLoaderOptions === null) {
            self::setLibXmlLoaderOptions();
        }
        return self::$libXmlLoaderOptions;
    }

    /**
     * Set default options for libxml loader.
     *
     * @param ?int $options Default options for libxml loader
     */
    public static function setLibXmlLoaderOptions(?int $options = null): void
    {
        if (is_null($options)) {
            $options = LIBXML_DTDLOAD | LIBXML_DTDATTR;
        }
        self::$libXmlLoaderOptions = $options;
    }
}