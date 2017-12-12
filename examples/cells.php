<?php
/**
 * Пример использования ячеек.
 *
 * @author decaseal <decaseal@gmail.com>
 */

use Topvisor\XlsxCreator\Workbook;

include __DIR__.'/../vendor/autoload.php';

$workbook = new Workbook(); // инициализация библиотеки
$worksheet = $workbook->addWorksheet('Sheet1'); // создание таблицы
$row = $worksheet->addRow(); // создание строки

$cell = $row->addCell('value'); // добавить ячейку со значением
$cell = $row->getCell(3); // получить ячейку по номеру столбца
$cell = $worksheet->getCell(2, 2); // получить ячейку по номеру строки и столбца

$cell->setValue('value'); // значение ячейки

$workbook->toFile(__DIR__.'/cells.xlsx'); // создание xlsx файла