# XlsxCrealor

Потоковая PHP библиотека для создания xlsx файлов

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
        "topvisor/XlsxCreator": "v0.6-alpha"
    }
}
```

# Вступление

Примеры расположены в папке [examples](https://github.com/topvisor/XlsxCreator/tree/master/examples).

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
// Фиксирует строку
$row->commit();
// Фиксирует таблицу
$worksheet->commit();
// Фиксирует книгу и создает xlsx файл
$workbook->commit();
```
