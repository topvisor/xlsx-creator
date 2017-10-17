# XlsxCrealor

Потоковая PHP библиотека для создания xlsx файлов

# Установка

Используйте [composer](https://getcomposer.org/) для установки

composer.json
```json
{
    "require": {
        "topvisor/xlsx-creator": "v0.8-alpha"
    }
}
```

# Вступление

Примеры расположены в папке [examples](https://github.com/topvisor/xlsx-creator/tree/master/examples).

Важнейшей особенностью библиотеки является ее потоковость. [Фиксация изменений](#Фиксация-изменений) выгружает данные в файл, и удаляет их 
из оперативной памяти.

# Пример использования библиотеки

```php
$xlsxFilename = __DIR__.'/example1.xlsx'; // путь, по которому будет создан xlsx файл
$workbook = new \Topvisor\XlsxCreator\Workbook($xlsxFilename); // инициализация библиотеки
 
$sheetName = 'Sheet1'; // имя таблицы
$worksheet = $workbook->addWorksheet($sheetName); // создание таблицы
 
$values = ['test1', 'test2', 3, 4]; // значения ячеек строки
$worksheet->addRow($values); // создание строки
 
$workbook->commit(); // создание xlsx файла
```

# Фиксация изменений

Фиксация выгружает изменения из памяти в файл. После фиксации объект становится неизменяемым.

```php
// Фиксирует книгу и создает xlsx файл
$workbook->commit();
 
// Фиксирует таблицу
$worksheet->commit();
 
// Фиксирует строку
$row->commit();
```
