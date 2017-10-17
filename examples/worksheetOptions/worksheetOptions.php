<?php
/**
 * Пример задания различных опций таблицы.
 *
 * @author decaseal <decaseal@gmail.com>
 */

use Topvisor\XlsxCreator\Structures\Views\NormalView;
use Topvisor\XlsxCreator\Workbook;

include __DIR__ . '/../../vendor/autoload.php';

$workbook = new Workbook(__DIR__.'/workbookOptions.xlsx'); // инициализация библиотеки
$worksheet = $workbook->addWorksheet('Sheet1'); // создание таблицы

$worksheet
	->setView(new NormalView());

$worksheet->addRow(['test1', 'test2', 3, 4]); // создание строки

$workbook->commit(); // создание xlsx файла