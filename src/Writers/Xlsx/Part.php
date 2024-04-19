<?php

namespace EasyExcel\Writers\Xlsx;

use EasyExcel\Helpers\Filesystem;
use EasyExcel\Interfaces\ExcelInterface;
use XMLWriter;

abstract class Part
{
    protected const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
    protected const NS_RELATIONSHIPS = 'http://schemas.openxmlformats.org/package/2006/relationships';

    /**
     * @var XMLWriter
     */
    protected $xml;

    /**
     * @var string
     */
    protected $filename;

    /**
     * @var ExcelInterface
     */
    protected $excel;

    /**
     * @var bool
     */
    protected $flushed = false;

    /**
     * @param  ExcelInterface  $excel
     */
    public function __construct(ExcelInterface $excel)
    {
        $this->excel = $excel;
        $this->filename = Filesystem::getTempName();
        $this->xml = new XMLWriter();
        $this->xml->openUri($this->filename);
        $this->xml->startDocument('1.0', 'UTF-8', 'yes');
        $this->writeStart();
    }

    /**
     * Write start.
     *
     * @return $this
     */
    abstract protected function writeStart(): self;

    /**
     * @return \XMLWriter
     */
    public function getXml(): XMLWriter
    {
        return $this->xml;
    }

    /**
     * Get filename.
     *
     * @return string
     */
    public function getFilename(): string
    {
        if (!$this->flushed) {
            $this->writeEnd();
            $this->xml->flush();
            $this->flushed = true;
        }

        return $this->filename;
    }

    /**
     * Write end.
     *
     * @return $this
     */
    abstract protected function writeEnd(): self;

    /**
     * Close part.
     *
     * @return void
     */
    public function close(): void
    {
        $this->xml = null;
        @unlink($this->filename);
    }

    /**
     * Write element with attributes.
     *
     * @param  string  $name
     * @param  array   $attributes
     *
     * @return $this
     */
    protected function writeElementWithAttributes(string $name, array $attributes = []): self
    {
        $this->xml->startElement($name);
        foreach ($attributes as $name => $value) {
            $this->xml->writeAttribute($name, $value);
        }
        $this->xml->endElement();

        return $this;
    }
}
