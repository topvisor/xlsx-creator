<?php
/**
 * Пример использования стилей.
 *
 * @author decaseal <decaseal@gmail.com>
 */

use Topvisor\XlsxCreator\Structures\Color;
use Topvisor\XlsxCreator\Structures\Styles\Alignment\Alignment;
use Topvisor\XlsxCreator\Structures\Styles\Alignment\TextRotation;
use Topvisor\XlsxCreator\Structures\Styles\Borders\Border;
use Topvisor\XlsxCreator\Structures\Styles\Borders\Borders;
use Topvisor\XlsxCreator\Structures\Styles\Font;
use Topvisor\XlsxCreator\Workbook;

include __DIR__.'/../vendor/autoload.php';

$workbook = new Workbook(__DIR__.'/styles.xlsx'); // инициализация библиотеки
$worksheet = $workbook->addWorksheet('Sheet1'); // создание таблицы

// стили можно назначать колонкам, строкам и ячейкам

// границы
$borderStyle = 'thin'; // тип границы ('thin', 'dotted', 'dashDot', 'hair', 'dashDotDot', 'slantDashDot'
                                    // 'mediumDashed', 'mediumDashDotDot', 'mediumDashDot', 'medium', 'double', 'thick')
$borderColor = Color::fromHex('000FF0'); // цвет границы
$border = (new Border($borderStyle, $borderColor)); // граница
$borders = (new Borders())
    ->setDefaultColor($borderColor) // цвет границ по умолчанию
    ->setTop($border) // верхняя граница
    ->setBottom($border) // нижняя граница
    ->setLeft($border) // левая граница
    ->setRight($border) // правая граница
    ->setDiagonalStyle($border) // диагональная граница
    ->setDiagonalUp(true) // показывать диагональную границу (из левого верхнего угла)
    ->setDiagonalDown(false); // показывать диагональную границу (в левый нижний угол)

// заливка
$fill = Color::fromHex('0FF000', 'A0');

// шрифт
$fontColor = Color::fromHex('F00F00'); // цвет шрифта
$font = (new Font())
    ->setColor($fontColor)
    ->setBold(true) // жирный
    ->setItalic(false) // курсивный
    ->setName('Arial') // название
    ->setSize(20) // размер
    ->setStrike(true) // зачеркнутый
    ->setUnderline('double') // подчеркнутый ('single', 'double', 'singleAccounting', 'doubleAccounting')
    ->setVerticalAlign('subscript'); // надстрочный/подстрочный ('superscript', 'subscript')

// формат чисел
$numFmt = '0%';

// выравнивание текста
$textRotation = TextRotation::fromAngle(-45); // поворот текста (градусы)
$textRotation = TextRotation::vertical(); // текст по вертикали
$alignment = (new Alignment())
    ->setHorizontal('center') // по горизонтали ('left', 'center', 'right', 'fill',
                                                // 'centerContinuous', 'distributed', 'justify')
    ->setVertical('top') // по вертикали ('top', 'center', 'bottom', 'distributed', 'justify')
    ->setIndent(5) // отступ слева
    ->setReadingOrder('leftToRight') // направление чтения ('leftToRight', 'rightToLeft')
    ->setWrapText(true) // заполнить ячейку текстом
    ->setTextRotation($textRotation);

// назначение стилей
$worksheet->getCell(1, 1)
    ->setBorders($borders)
    ->setFill($fill)
    ->setFont($font)
    ->setNumFmt($numFmt)
    ->setAlignment($alignment);

$workbook->commit(); // создание xlsx файла