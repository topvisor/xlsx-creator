<?php
/**
 * Пример объединения ячеек.
 *
 * @author decaseal <decaseal@gmail.com>
 */

use Topvisor\XlsxCreator\Structures\Range\CellsRange;
use Topvisor\XlsxCreator\Workbook;

include __DIR__.'/../vendor/autoload.php';

$workbook = new Workbook(__DIR__.'/merges.xlsx'); // инициализация библиотеки
$worksheet = $workbook->addWorksheet('Sheet1'); // создание таблицы

// добавление строк
$worksheet->addRow([1, 2]);
$worksheet->addRow([3, 4]);

// объединение ячеек
$worksheet->mergeCells(new CellsRange(1, 1, 1, 2));

// отменить объединение ячеек
$worksheet->mergeCells(new CellsRange(2, 1, 2, 2));
$worksheet->unMergeCells(new CellsRange(2, 1, 2, 2));

$workbook->commit(); // создание xlsx файла