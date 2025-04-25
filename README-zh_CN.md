# EasyExcel

[English](./README.md) · 中文

使用非常低的内存读取大体积电子表格，支持格式 CSV、ODS、XLS、XLSX。

## 安装

```bash
composer require chenmobuys/easyexcel
```

## 使用方法

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

## 已实现功能

* ✔ Supported
* ● Partially supported
* ✖ Not supported
* N/A Cannot be supported

| 功能                 | Csv  | Ods  | Xls  | Xlsx  |
|--------------------|:----:|:----:|:----:|:-----:|
| CellOriginalValue  |  ✔   |  ✔   |  ✔   |   ✔   |
| CellFormattedValue | N/A  |  ✔   |  ✔   |   ✔   |
| CellFormulaValue   | N/A  |  ✔   |  ✔   |   ✔   |
| CellStyle          | N/A  |  ✖   |  ✔   |   ✔   |
| Hyperlinks         | N/A  |  ✔   |  ✔   |   ✔   |
| MergeCells         | N/A  |  ✖   |  ✔   |   ✔   |
