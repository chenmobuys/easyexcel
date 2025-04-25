<?php

namespace EasyExcel\Readers;

use EasyExcel\Interfaces\Reader;
use EasyExcel\Interfaces\ReaderExcel;

abstract class BaseReader implements Reader
{
    /**
     * @var string
     */
    protected $filename;

    /**
     * @var \EasyExcel\Interfaces\ReaderExcel
     */
    protected $excel;

    /**
     * Close flag.
     *
     * @var bool
     */
    protected $closed = false;

    /**
     * Load file.
     *
     * @param string $filename
     * @return \EasyExcel\Interfaces\Reader
     */
    public function load(string $filename): Reader
    {
        $this->filename = $filename;
        $this->excel = $this->loadFromFile($filename);

        return $this;
    }

    /**
     * Close file.
     *
     * @return void
     */
    public function close(): void
    {
        if (!$this->closed) {
            if ($this->excel) {
                $this->excel->close();
            }
            $this->closeReader();
            $this->closed = true;
        }
    }

    /**
     * Get excel.
     *
     * @return \EasyExcel\Interfaces\ReaderExcel|null
     */
    public function getExcel(): ?ReaderExcel
    {
        return $this->excel;
    }

    /**
     * Load from file.
     *
     * @param string $filename
     * @return \EasyExcel\Interfaces\ReaderExcel
     */
    abstract protected function loadFromFile(string $filename): ReaderExcel;

    /**
     * Close reader.
     *
     * @return void
     */
    abstract protected function closeReader(): void;

    /**
     * If method not exists, call method in Excel.
     *
     * @param $name
     * @param $arguments
     * @return mixed
     */
    public function __call($name, $arguments)
    {
        return call_user_func([$this->excel, $name], ...$arguments);
    }

    /**
     * Destructor.
     */
    public function __destruct()
    {
        $this->close();
    }
}