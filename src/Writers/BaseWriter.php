<?php

namespace EasyExcel\Writers;

use EasyExcel\Interfaces\Writer;
use EasyExcel\Interfaces\WriterExcel;

abstract class BaseWriter implements Writer
{
    /**
     * @var string
     */
    protected $filename;

    /**
     * @var \EasyExcel\Interfaces\WriterExcel
     */
    protected $excel;

    /**
     * Close flag.
     *
     * @var bool
     */
    protected $closed = false;

    /**
     * Determine whether the file is writeable.
     *
     * @param string $filename
     * @return bool
     */
    public function writeable(string $filename): bool
    {
        return is_dir(dirname($filename)) && is_writeable(dirname($filename));
    }

    /**
     * Open file.
     *
     * @param string $filename
     * @return \EasyExcel\Interfaces\Writer
     */
    public function open(string $filename): Writer
    {
        $this->filename = $filename;

        $this->excel = $this->openFromFile($filename);

        return $this;
    }

    /**
     * Close and save file.
     *
     * @return void
     */
    public function close(): void
    {
        if (!$this->closed) {
            if ($this->excel) {
                $this->excel->close();
            }
            $this->closeWriter();

            $this->closed = true;
        }
    }

    /**
     * Get excel.
     *
     * @return \EasyExcel\Interfaces\WriterExcel|null
     */
    public function getExcel(): ?WriterExcel
    {
        return $this->excel;
    }

    /**
     * Open from file.
     *
     * @param string $filename
     * @return \EasyExcel\Interfaces\WriterExcel
     */
    abstract protected function openFromFile(string $filename): WriterExcel;

    /**
     * Close writer.
     *
     * @return void
     */
    abstract protected function closeWriter(): void;

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