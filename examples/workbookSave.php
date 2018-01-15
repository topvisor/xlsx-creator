<?php
/**
 * Пример сохранения xlsx файла
 *
 * @author decaseal <decaseal@gmail.com>
 */

use Topvisor\XlsxCreator\Workbook;

include __DIR__.'/../vendor/autoload.php';

$workbook = new Workbook(); // инициализация библиотеки

$worksheet = $workbook->addWorksheet('Sheet1'); // создание таблицы
$worksheet->addRow(['test1', 'test2', 3, 4]); // создание строки

// Запись xlsx файла в поток вывода (http)
$xlsxFilename = 'workbookSave.xlsx'; // имя xlsx файла
$exit = false; // не завершать выполение скрипта после записи xlsx файла в поток
$workbook->toHttp($xlsxFilename, $exit);

// Создание xlsx файла
$xlsxFilePath = __DIR__."/$xlsxFilename"; // путь, по которому будет создан xlsx файл
$workbook->toFile($xlsxFilePath); // создание xlsx файла

