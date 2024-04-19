# EasyExcel

使用低内存读取、写入大电子表格，支持格式 CSV、ODS、XLS、XLSX。

Read and write large spreadsheets using low memory, supporting formats CSV, ODS, XLS, XLSX.

## 环境要求 Environment

* PHP version `^7.1`||`^8.0`
* PHP extension `xml` 
* PHP extension `libxml`
* PHP extension `fileinfo`
* PHP extension `simplexml`
* PHP extension `xmlreader`
* PHP extension `xmlwriter`
* PHP extension `iconv` (suggest)
* PHP extension `mbstring` (suggest)

## 安装 Install

```bash
composer require chenmobuys/easyexcel
```

## 使用方法 Usage

```php
<?php

use EasyExcel\Factory;
use EasyExcel\Metadata\Style;

// read excel
$filename = "/path/to/sample.xlsx";
$reader = Factory::load($filename);
foreach ($reader->getRowIterator() as $row) {
    $rowArray = $row->toArray();
}

// write excel
$filename = "/path/to/sample.xlsx";
$writer = Factory::open($filename);
$writer->addRow(["Foo", "Bar"])->close();

// write excel with style
$filename = "/path/to/sample.xlsx";
$writer = Factory::open($filename);
$style = Style::builder()
    ->setFontSize(12)
    ->setFontColor(Style\Color::RED)
    ->build();
$writer->addRow(["Foo", "Bar"], $style)->close();

```

## 参考 Refer

* https://github.com/mk-j/PHP_XLSXWriter
* https://github.com/nuovo/spreadsheet-reader

[//]: # (## 已实现功能)

[//]: # ()

[//]: # (* ✔ Supported)

[//]: # (* ● Partially supported)

[//]: # (* ✖ Not supported)

[//]: # (* N/A Cannot be supported)

[//]: # ()

[//]: # (| 功能                 | Csv  | Ods  | Xls  | Xlsx  |)

[//]: # (|--------------------|:----:|:----:|:----:|:-----:|)

[//]: # (| CellOriginalValue  |  ✔   |  ✔   |  ✔   |   ✔   |)

[//]: # (| CellFormattedValue | N/A  |  ✔   |  ✔   |   ✔   |)

[//]: # (| CellFormulaValue   | N/A  |  ✔   |  ✔   |   ✔   |)

[//]: # (| CellStyle          | N/A  |  ✖   |  ✔   |   ✔   |)

[//]: # (| Hyperlinks         | N/A  |  ✔   |  ✔   |   ✔   |)

[//]: # (| MergeCells         | N/A  |  ✖   |  ✔   |   ✔   |)
