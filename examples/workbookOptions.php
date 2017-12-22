<?php
/**
 * Пример задания различных опций xlsx файла
 *
 * @author decaseal <decaseal@gmail.com>
 */

use Topvisor\XlsxCreator\Workbook;
use Topvisor\XlsxCreator\Structures\Values\HyperlinkValue;

include __DIR__.'/../vendor/autoload.php';

$useSharedStrings = true; // использовать общие строки. Данная опция уменьшит размер xlsx файла за счет увеличения потребления оперативной памяти.
$checkRelsDoubles = true; // проверять дубли ссылок. Данная опция уменьшит размер xlsx файла за счет увеличения потребления оперативной памяти.
$workbook = new Workbook($useSharedStrings, $checkRelsDoubles); // инициализация библиотеки

$version = Workbook::VERSION; // верcия библиотеки

$workbook
	->setTempdir(sys_get_temp_dir()) // путь к директории для хранения временных файлов библиотеки
	->setCompany('Topvisor') // компания
	->setCreator('decaseal') // создатель файла
	->setLastModifiedBy('decaseal') // последний изменявший файл
	->setCreated(new DateTime()) // время создания
	->setModified(new DateTime()) // время изменения
	->setManager('decaseal'); // менеджер

$worksheet = $workbook->addWorksheet('Sheet1'); // создание таблицы
$worksheet->addRow(['test1', 'test2', 3, 4]); // создание строки

$worksheet->addRow([new HyperlinkValue('https://topvisor.ru/', 'topvisor.ru'), new HyperlinkValue('https://topvisor.com/', 'topvisor.com')]);
$worksheet->addRow([new HyperlinkValue('https://topvisor.com/', 'topvisor.com')]);
$worksheet->addRow([new HyperlinkValue('https://topvisor.ru/', 'topvisor.ru')]);

$xlsxFilename = __DIR__.'/workbookOptions.xlsx'; // путь, по которому будет создан xlsx файл
$workbook->toFile($xlsxFilename); // создание xlsx файла