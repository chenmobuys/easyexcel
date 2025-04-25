<?php

namespace EasyExcel\Writers;

use EasyExcel\Interfaces\WriterExcel;
use EasyExcel\Interfaces\WriterRow;
use EasyExcel\Interfaces\WriterSheet as WriterSheetInterface;
use EasyExcel\Metadata\Sheet;

class WriterSheet extends Sheet implements WriterSheetInterface
{
    /**
     * @var \EasyExcel\Interfaces\WriterExcel
     */
    protected $excel;

    /**
     * @var \EasyExcel\Interfaces\WriterRow
     */
    protected $rowWriter;

    /**
     * @param \EasyExcel\Interfaces\WriterExcel $excel
     * @param string $name
     * @param int $index
     */
    public function __construct(WriterExcel $excel, string $name, int $index)
    {
        $this->excel = $excel;
        $this->name = $name;
        $this->index = $index;
    }

    /**
     * Get excel.
     *
     * @return \EasyExcel\Interfaces\WriterExcel
     */
    public function getExcel(): WriterExcel
    {
        return $this->excel;
    }

    /**
     * Get row writer.
     *
     * @return \EasyExcel\Interfaces\WriterRow
     */
    public function getRowWriter(): WriterRow
    {
        if (!$this->rowWriter) {
            $this->rowWriter = $this->getExcel()->getWriter()->getRowWriter($this);
        }

        return $this->rowWriter;
    }

    /**
     * Close sheet.
     *
     * @return void
     */
    public function close(): void
    {
        if ($this->rowWriter) {
            $this->rowWriter->close();
        }
    }
}