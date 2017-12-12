<?php
/**
 * Пример использования колонок.
 *
 * @author decaseal <decaseal@gmail.com>
 */

use Topvisor\XlsxCreator\Workbook;

include __DIR__.'/../vendor/autoload.php';

$workbook = new Workbook(); // инициализация библиотеки
$worksheet = $workbook->addWorksheet('Sheet1'); // создание таблицы

$column = $worksheet->addColumn(); // создать колонку
$column = $worksheet->getColumn(2); // получить колонку

$column
    ->setHidden(true) // спрятать колонку
    ->setWidth(20); // задать ширину колонки

$worksheet->addRow(['test1', 'test2', 3, 4]); // создание строки

$workbook->toFile(__DIR__.'/columns.xlsx'); // создание xlsx файла