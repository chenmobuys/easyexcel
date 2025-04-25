# EasyExcel

English · [中文](./README-zh_CN.md)

Read large spreadsheets using very low memory usage, supports formats CSV, ODS, XLS, XLSX.

## Install

```bash
composer require chenmobuys/easyexcel
```

## Usage

```php
<?php

use EasyExcel\Factory;

$filename = 'Excel.xlsx';
$reader = Factory::createReaderForFile($filename);
$readerExcel = $reader->load($filename);
$activeSheet = $readerExcel->getActiveSheet();
foreach ($activeSheet->getRowIterator() as $row) {
    // Do something...
}
```

## Features implemented

* ✔ Supported
* ● Partially supported
* ✖ Not supported
* N/A Cannot be supported

| Feature            | Csv | Ods | Xls | Xlsx |
|--------------------|:---:|:---:|:---:|:----:|
| CellOriginalValue  |  ✔  |  ✔  |  ✔  |  ✔   |
| CellFormattedValue | N/A |  ✔  |  ✔  |  ✔   |
| CellFormulaValue   | N/A |  ✔  |  ✔  |  ✔   |
| CellStyle          | N/A |  ✖  |  ✔  |  ✔   |
| Hyperlinks         | N/A |  ✔  |  ✔  |  ✔   |
| MergeCells         | N/A |  ✖  |  ✔  |  ✔   |
