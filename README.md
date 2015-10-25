<h1 align="center">
    <img src="http://cdn.florian.ec/plum-logo.svg" alt="Plum" width="300">
</h1>

> PlumExcel includes readers and writers for Microsoft Excel files for Plum.  Plum is a data processing pipeline for
PHP.

[![Build Status](https://img.shields.io/travis/plumphp/plum-excel.svg?style=flat)](https://travis-ci.org/plumphp/plum-excel)
[![Windows Build status](https://ci.appveyor.com/api/projects/status/835tpvim8hb986ef?svg=true)](https://ci.appveyor.com/project/florianeckerstorfer/plum-excel)
[![Scrutinizer Code Quality](https://img.shields.io/scrutinizer/g/plumphp/plum-excel.svg?style=flat)](https://scrutinizer-ci.com/g/plumphp/plum-excel/?branch=master)
[![Code Coverage](https://img.shields.io/scrutinizer/coverage/g/plumphp/plum-excel.svg?style=flat)](https://scrutinizer-ci.com/g/plumphp/plum-excel/?branch=master)
[![StyleCI](https://styleci.io/repos/31608543/shield)](https://styleci.io/repos/31608543)

Developed by [Florian Eckerstorfer](https://florian.ec) in Vienna, Europe.


Features
-------

- Read Microsoft Excel (`.xlsx` and `.xls`) files
- Write Microsoft Excel (`.xlsx` and `.xls`) files
- Uses [PHPExcel](https://github.com/PHPOffice/PHPExcel)


Installation
------------

You can install `plum-excel` using [Composer](http://getcomposer.org).

```shell
$ composer require plumphp/plum-excel
```


Usage
-----

PlumExcel contains a reader and a writer for Plum. Please refer to the
[Plum documentation](https://github.com/plumphp/plum/blob/master/docs/index.md) for more information about Plum.

You can also find examples of how to use `ExcelReader` and `ExcelWriter` in the
[`examples/`](https://github.com/plumphp/plum-excel/tree/master/examples) folder.

### Write Excel files

Writing Excel files is extremely simply. Just pass the filename of the file to the constructor. If you want to add
a header row call the `autoDetectHeader()` method.

```php
use Plum\PlumExcel\ExcelWriter;

$writer = new ExcelWriter(__DIR__.'/example.xlsx');
$writer->autoDetectHeader();
```

You can manually set the header names by calling the `setHeader()`  method and passing an array with names.

```php
$writer->setHeader(['Country Name', 'ISO 3166-1-alpha-2 code']);
```

However, if you want more control, you can also pass an instance of `PHPExcel` to the constructor and the format
(`Excel2007` or `Excel5`) or an implementation of `PHPExcel_Writer_IWriter`.

```php
$writer = new ExcelWriter(__DIR__.'/example.xlsx', $excel, 'Excel2007', $writer);
```

### Read Excel files

Reading Excel files is also pretty simple.

```php
use Plum\PlumExcel\ExcelReader;

$reader = new ExcelReader(__DIR__.'/example.xlsx');
```

Instead of a filename you can also pass an instance of `PHPExcel` to the constructor.

```php
use Plum\PlumExcel\ExcelReader;

$excel = PHPExcel_IOFactory::load(__DIR__.'/example.xlsx');
$reader = new ExcelReader($excel);
```

Plum can automatically detect the headers by using `Plum\Plum\Converter\HeaderConverter`.

```php
use Plum\Plum\Converter\HeaderConverter;
use Plum\Plum\Filter\SkipFirstFilter;
use Plum\PlumExcel\ExcelReader;

$workflow->addConverter(new HeaderConverter());
$workflow->addFilter(new SkipFirstFilter(1));
$workflow->process(new ExcelReader($filename));
```


Change Log
----------

### Version 0.3 (24 October 2015)

- [#2](https://github.com/plumphp/plum-excel/pull/2) Handle items that are not arrays 

### Version 0.2.1 (28 April 2015)

- Fix Plum version

### Version 0.2 (22 April 2015)

- `ExcelReader` supports filename as input
- Add support for `Plum\Plum\Reader\ReaderFactory`

### Version 0.1 (21 April 2015)

- Initial release


License
-------

The MIT license applies to plumphp/plum-excel. For the full copyright and license information,
please view the LICENSE file distributed with this source code.
