<?php

namespace EasyExcel\Helpers;

use EasyExcel\Settings;

class Filesystem
{
    public const DEFAULT_TEMP_PREFIX = 'easyexcel';

    /**
     * Get temp filename.
     *
     * @param string $prefix
     * @return string
     */
    public static function getTempName(string $prefix = self::DEFAULT_TEMP_PREFIX): string
    {
        return @tempnam(Settings::getTempDir(), $prefix);
    }
}