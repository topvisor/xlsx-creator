# XlsxCrealor

Потоковая PHP библиотека для создания xlsx файлов

# Установка

Используйте [composer](https://getcomposer.org/) для установки

composer.json
```json
{
    "require": {
        "topvisor/xlsx-creator": "1.2"
    }
}
```

# Вступление

Примеры расположены в папке [examples](https://github.com/topvisor/xlsx-creator/tree/master/examples).

Запуск примеров
```bash
git clone https://github.com/topvisor/xlsx-creator.git
cd xlsx-creator
composer install
php examples/simple.php
```

Важнейшей особенностью библиотеки является ее потоковость. [Фиксация изменений](#Фиксация-изменений) выгружает данные в файл, и удаляет их 
из оперативной памяти.

# Пример использования библиотеки

```php
$workbook = new \Topvisor\XlsxCreator\Workbook($xlsxFilename); // инициализация библиотеки
 
$sheetName = 'Sheet1'; // имя таблицы
$worksheet = $workbook->addWorksheet($sheetName); // создание таблицы
 
$values = ['test1', 'test2', 3, 4]; // значения ячеек строки
$worksheet->addRow($values); // создание строки
 
$xlsxFilename = __DIR__.'/example1.xlsx'; // путь, по которому будет создан xlsx файл
$workbook->toFile($xlsxFilename); // создание xlsx файла
```

# Фиксация изменений

Фиксация выгружает изменения из памяти в файл. После фиксации объект становится неизменяемым.

```php
// Фиксирует книгу
$workbook->commit();
 
// Фиксирует таблицу
$worksheet->commit();
 
// Фиксирует строку (и все предыдущие)
$row->commit();
```
