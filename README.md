# XlsxCrealor

Потоковая библиотека для создания xlsx файлов

# Установка

Используйте [composer](https://getcomposer.org/) для установки

composer.json
```json
{
    "repositories": [
        {
            "type": "vcs",
            "url": "https://github.com/topvisor/XlsxCreator"
        }
    ],
    "require": {
        "topvisor/XlsxCreator": "v0.2-dev"
    }
}
```

# Использование

* [Создание книги](#Создание-книги)
* [Параметры книги](#Параметры-книги)
* [Добавление таблицы](#Добавление-таблицы)
* [Доступ к таблице](#Доступ-к-таблице)
* [Параметры таблицы](#Параметры-таблицы)
	* [Основные параметры](#Основные-параметры)
	* [Параметры печати](#Параметры-печати)
* [Добавление строки](#Добавление-строки)
* [Значения ячеек](#Значения-ячеек)
* [Фиксация изменений](#Фиксация-изменений)

## Создание книги

```php
// string $pathToXlsxFile - путь, по которому будет создан xlsx файл
$workbook = new Workbook($pathToXlsxFile);
```

## Параметры книги

```php
// string $tempdir - путь к директории для хранения временных файлов 
// DateTime $created - время создания
// DateTime $modified - время изменения
// string $creator - создатель
// string $lastModifiedBy - автор последнего изменения 
// string $company - компания 
// string $manager - менеджер 
$workbook
	->setTempdir($tempdir)
	->setCreated($created)
	->setModified($modified)
	->setCreator($creator)
	->setLastModifiedBy($lastModifiedBy)
	->setCompany($company)
	->setManager($manager);
```

## Добавление таблицы

```php
// string $name - имя таблицы 
$worksheet = $workbook->addWorksheet($name);
```

## Доступ к таблице

```php
// string $name - имя таблицы 
$worksheet = $workbook->getWorksheetByName($name);
// int $id - id таблицы
$worksheet = $workbook->getWorksheetById($id);
// array $worksheets - массив таблиц
$worksheets = $workbook->getWorksheets();
```

## Параметры таблицы

Следует задавать параметры таблицы до добавления строк, в противном случае параметры не будут применены!

### Основные параметры

```php
// string $tabColor - цвет вкладки в формате ARGB (например, 'FF00FF00')
// int $outlineLevelCol - worksheet column outline level
// int $outlineLevelRow - worksheet row outline level
// int $defaultRowHeight - высота строки по умолчанию
$worksheet
	->setTabColor($tabColor)
	->setOutlineLevelCol($outlineLevelCol)
	->setOutlineLevelRow($outlineLevelRow)
	->setDefaultRowHeight($defaultRowHeight)
```

### Параметры печати

```php
// array $pageSetup - параметры печати
$pageSetup = [
	'margins' => [                 // Пробелы на границах страницы (в дюймах)
		'left' => 0.7, 
		'right' => 0.7, 
		'top' => 0.75, 
		'bottom' => 0.75, 
		'header' => 0.3, 
		'footer' => 0.3
	],
	'orientation' => 'portrait',   // Ориентация страницы ('portrait', 'landscape')
	'horizontalDpi' => 4294967295, // Точек на дюйм по горизонтали
	'verticalDpi' => 4294967295,   // Точек на дюйм по вертикали
	'pageOrder' => null,           // Порядок печати страниц ('downThenOver', 'overThenDown')
	'blackAndWhite' => false,      // Печать без цвета
	'draft' => false,              // Печать с меньшим качеством (и чернилами)
	'cellComments' => null,        // Где разместить комментарии ('atEnd', 'asDisplayed', 'None')
	'errors' => null,              // Где показывать ошибки ('dash', 'blank', 'NA', 'displayed')
	'scale' => 100,                // Процент увеличения/уменьшения размеров печати
	'fitToWidth' => 1,             // Сколько страниц должно помещаться на листе по ширине (активно если нет scale)
	'fitToHeight' => 1,            // Сколько страниц должно помещаться на листе по высоте (активно если нет scale)
	'paperSize' => null,           // Какой размер бумаги использовать (int) (9 - А4)
	'showRowColHeaders' => false,  // Показывать номера строк и столбцов
	'firstPageNumber' => null,     // Какой номер использовать для первой страницы
];
$worksheet->setPageSetup($pageSetup);
```

## Добавление строки

```php
// array|null $values - значения ячеек строки 
$row = $worksheet->addRow($values);
```

## Значения ячеек

```php
// array|null $values - значения ячеек строки 
$row->setCells($values);
```

## Фиксация изменений

Фиксация выгружает изменения из памяти в файл. После фиксации объект становится неизменяемым.

```php
// Фиксирует строку
$row->commit();
// Фиксирует таблицу
$worksheet->commit();
// Фиксирует книгу и создает xlsx файл
$workbook->commit();
```
