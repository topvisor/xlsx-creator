<?php
/**
 * Пример использования строк.
 *
 * @author decaseal <decaseal@gmail.com>
 */

use Topvisor\XlsxCreator\Workbook;

include __DIR__.'/../vendor/autoload.php';

$workbook = new Workbook(); // инициализация библиотеки
$worksheet = $workbook->addWorksheet('Sheet1'); // создание таблицы

$values = ['test1', 'test2', 3, 4]; // значения ячеек строки

$row = $worksheet->addRow(); // создать пустую строку
$row = $worksheet->addRow($values); // создать строку со значениями ячеек
$row = $worksheet->getRow(3); // полученить строку

$row
    ->setHidden(true) // скрыть строку
    ->setHeight(40); // высота строки

$workbook->toFile(__DIR__.'/rows.xlsx'); // создание xlsx файла