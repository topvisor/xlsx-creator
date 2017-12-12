<?php
/**
 * Пример добавления изображений.
 *
 * @author decaseal <decaseal@gmail.com>
 */

use Topvisor\XlsxCreator\Structures\Range\Range;
use Topvisor\XlsxCreator\Workbook;

include __DIR__.'/../vendor/autoload.php';

$workbook = new Workbook(); // инициализация библиотеки
$worksheet = $workbook->addWorksheet('Sheet1'); // создание таблицы
$worksheet->addRow();

$position = new Range(1.5, 1.5, 3.5, 4); // расположение картинки
$filename = __DIR__.'/images/image'; // путь к картинке
$extension = 'png'; // расширение картинки
$name = 'Image'; // имя картинки

$worksheet->addImage($position, $filename, $extension, $name); // добавить картинку

$workbook->toFile(__DIR__.'/images.xlsx'); // создание xlsx файла