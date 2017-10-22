<?php
/**
 * Пример использования колонок.
 *
 * @author decaseal <decaseal@gmail.com>
 */

use Topvisor\XlsxCreator\Workbook;

include __DIR__.'/../vendor/autoload.php';

$xlsxFilename = __DIR__.'/simple.xlsx'; // путь, по которому будет создан xlsx файл
$workbook = new Workbook($xlsxFilename); // инициализация библиотеки

$sheetName = 'Sheet1'; // имя таблицы
$worksheet = $workbook->addWorksheet($sheetName); // создание таблицы

$column = $worksheet->addColumn(); // создать колонку
$column = $worksheet->getColumn(2); // получить колонку

$column
    ->setHidden(true) // спрятать колонку
    ->setWidth(20); // задать ширину колонки

$values = ['test1', 'test2', 3, 4]; // значения ячеек строки
$worksheet->addRow($values); // создание строки

$workbook->commit(); // создание xlsx файла