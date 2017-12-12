<?php
/**
 * Пример задания различных опций таблицы.
 *
 * @author decaseal <decaseal@gmail.com>
 */

use Topvisor\XlsxCreator\Structures\Color;
use Topvisor\XlsxCreator\Structures\PageSetup;
use Topvisor\XlsxCreator\Structures\Views\FrozenView;
use Topvisor\XlsxCreator\Structures\Views\NormalView;
use Topvisor\XlsxCreator\Structures\Views\SplitView;
use Topvisor\XlsxCreator\Workbook;

include __DIR__ . '/../vendor/autoload.php';

$workbook = new Workbook(); // инициализация библиотеки
$worksheet = $workbook->addWorksheet('Sheet1'); // создание таблицы

// Представления таблицы

// Обычное представление
$view = (new NormalView())
	->setView('pageBreakPreview') // cтиль отображения ('pageBreakPreview', 'pageLayout')
	->setActiveCell('B2') // адрес выбранной ячейки
	->setRightToLeft(true) // ориентация справа на лево
	->setShowGridLines(false) // показывать линии сетки
	->setShowRowColHeaders(true) // показывать заголовки строк и столбцов
	->setShowRuler(false) // показывать линейку в макете страницы
	->setZoomScale(90) //процент увеличения
	->setZoomScaleNormal(100); // нормальное увеличение

// Несколько строк и/или столбцов этого представления заморожено на месте
$xSplit = 2; // количество замороженных столбцов
$ySplit = 3; // количество замороженных строк
$view = (new FrozenView($xSplit, $ySplit))
	->setTopLeftCell('E5'); // Левая-верхняя ячейка в "незамороженной" панели

// Представление разделено на 4 секции с независимой прокруткой
$xSplit = 5000; // количество точек слева до границы
$ySplit = 6000; // количество точек сверху до границы
$view = (new SplitView($xSplit, $ySplit))
	->setTopLeftCell('G10') // Левая-верхняя ячейка в нижней правой панели
	->setActivePane('bottomRight'); // активная панель ('topLeft', 'topRight', 'bottomLeft', 'bottomRight')

// /Представления таблицы

// Параметры печати таблицы
$pageSetup = (new PageSetup())
	->setMarginLeft(0.5) // отступы
	->setMarginRight(0.5)
	->setMarginTop(0.5)
	->setMarginBottom(0.5)
	->setMarginHeader(0.5)
	->setMarginFooter(0.5)
	->setOrientation('landscape') // ориентация страницы ('portrait', 'landscape')
	->setHorizontalDpi(1000) // точек на дюйм по горизонтали
	->setVerticalDpi(1000) // точек на дюйм по вертикали
	->setFitToPage(true) // использовать ли настройки fitToWidth и fitToHeight или scale
	->setPageOrder('downThenOver') // порядок печати страниц ('downThenOver', 'overThenDown')
	->setBlackAndWhite(true) // печать без цвета
	->setDraft(true) // печать с меньшим качеством (и чернилами)
	->setCellComments('atEnd') // где разместить комментарии ('atEnd', 'asDisplayed', 'None')
	->setErrors('displayed') // где показывать ошибки ('dash', 'blank', 'NA', 'displayed')
	->setScale(200) // процент увеличения/уменьшения размеров печати
	->setFitToWidth(2) // сколько страниц должно помещаться на листе по ширине
	->setFitToHeight(2) // сколько страниц должно помещаться на листе по высоте
	->setPaperSize(9) // какой размер бумаги использовать (9 - А4)
	->setShowRowColHeaders(false) // показывать номера строк и столбцов
	->setFirstPageNumber(0); // какой номер использовать для первой страницы

$worksheet
	->setView($view) // параметры отображения
	->setPageSetup($pageSetup) // параметры печати
	->setTabColor(Color::fromHex('00FF00')); // цвет вкладки

$worksheet->addRow(['test1', 'test2', 3, 4]); // создание строки

$workbook->toFile(__DIR__.'/worksheetOptions.xlsx'); // создание xlsx файла