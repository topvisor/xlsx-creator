<?php
/**
 * Простейший пример использования библиотеки.
 *
 * @author decaseal <decaseal@gmail.com>
 */

use Topvisor\XlsxCreator\Workbook;

include __DIR__.'/../vendor/autoload.php';

$workbook = new Workbook(); // инициализация библиотеки

$sheetName = 'Sheet1'; // имя таблицы
$worksheet = $workbook->addWorksheet($sheetName); // создание таблицы

$values = ['test1', 'test2', 3, 4]; // значения ячеек строки
$worksheet->addRow($values); // создание строки

$xlsxFilename = __DIR__.'/simple.xlsx'; // путь, по которому будет создан xlsx файл
$workbook->toFile($xlsxFilename); // создание xlsx файла