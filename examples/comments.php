<?php
/**
 * Пример использования комментариев.
 *
 * @author decaseal <decaseal@gmail.com>
 */

use Topvisor\XlsxCreator\Workbook;

include __DIR__.'/../vendor/autoload.php';

$workbook = new Workbook(); // инициализация библиотеки
$worksheet = $workbook->addWorksheet('Sheet1'); // создание таблицы
$cell = $worksheet->getCell(1, 1); // получить ячейку по номеру строки и столбца

$cell
    ->setCommentWidth(3) // ширина комментария
    ->setCommentHeight(4) // высота комментария
    ->setComment('comment'); // комментарий

$workbook->toFile(__DIR__.'/comments.xlsx'); // создание xlsx файла