<?php
/**
 * Пример различных значений ячеек.
 *
 * @author decaseal <decaseal@gmail.com>
 */

use Topvisor\XlsxCreator\Structures\Styles\Font;
use Topvisor\XlsxCreator\Structures\Values\ErrorValue;
use Topvisor\XlsxCreator\Structures\Values\FormulaValue;
use Topvisor\XlsxCreator\Structures\Values\HyperlinkValue;
use Topvisor\XlsxCreator\Structures\Values\RichText\RichText;
use Topvisor\XlsxCreator\Structures\Values\RichText\RichTextValue;
use Topvisor\XlsxCreator\Workbook;

include __DIR__.'/../vendor/autoload.php';

$workbook = new Workbook(__DIR__.'/values.xlsx'); // инициализация библиотеки
$worksheet = $workbook->addWorksheet('Sheet1'); // создание таблицы

$nullValue = null; // пустая ячейка
$stringValue = 'string'; // строка
$dateValue = new DateTime(); // дата
$boolValue = true; // булево значение
$errorValue = new ErrorValue('#NAME?'); // ошибка ('#N/A', '#REF!', '#NAME?', '#DIV/0!', '#NULL!', '#VALUE!', '#NUM!')
$formulaValue = new FormulaValue('CONCATENATE(A2, A3)'); // формула
$sharedStringValue = $workbook->addSharedString('topvisor.com'); // общая строка

// число
$numberValue = 5;
$numberValue = -4.9;

// гиперссылка
$hyperlinkValue = new HyperlinkValue('https://topvisor.ru', 'topvisor.ru');
$hyperlinkValue = new HyperlinkValue('https://topvisor.com', $sharedStringValue);

// "богатый" текст
$richText1 = new RichText('text1 ');
$richText2 = new RichText('text2', (new Font())->setStrike(true));
$richTextValue = new RichTextValue([
    $richText1,
    $richText2
]);

$worksheet->addRow([
    $nullValue,
    $stringValue,
    $dateValue,
    $boolValue,
    $errorValue,
    $formulaValue,
    $sharedStringValue,
    $numberValue,
    $hyperlinkValue,
    $richTextValue
]);

$workbook->commit(); // создание xlsx файла