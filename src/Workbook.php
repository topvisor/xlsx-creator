<?php

/**
 * Библиотека для создания xlsx файлов
 *
 * @author decaseal <decaseal@gmail.com>
 * @version v1.12
 */

namespace Topvisor\XlsxCreator;

use DateTime;
use Topvisor\XlsxCreator\Exceptions\EmptyObjectException;
use Topvisor\XlsxCreator\Exceptions\InvalidValueException;
use Topvisor\XlsxCreator\Exceptions\ObjectCommittedException;
use Topvisor\XlsxCreator\Helpers\SharedStrings;
use Topvisor\XlsxCreator\Helpers\Styles;
use Topvisor\XlsxCreator\Helpers\Validator;
use Topvisor\XlsxCreator\Structures\Values\RichText\RichTextValue;
use Topvisor\XlsxCreator\Structures\Values\SharedStringValue;
use Topvisor\XlsxCreator\Xml\Book\WorkbookXml;
use Topvisor\XlsxCreator\Xml\Core\App\AppXml;
use Topvisor\XlsxCreator\Xml\Core\ContentTypesXml;
use Topvisor\XlsxCreator\Xml\Core\CoreXml;
use Topvisor\XlsxCreator\Xml\Core\Relationships\RelationshipsXml;
use ZipArchive;

/**
 * Class Workbook. Используйте его для создания xlsx файла.
 *
 * @package Topvisor\XlsxCreator
 */
class Workbook {
	public const VERSION = "v1.18";

	public const VALID_IMAGES_EXTENSION = ['jpeg', 'png', 'gif'];
	public const INVALID_WORKSHEET_NAME = '/[\/\\\?*\[\]]/';

	private $useSharedStrings;
	private $checkRelsDoubles;
	private $tempdir;
	private $created;
	private $modified;
	private $creator;
	private $lastModifiedBy;
	private $company;
	private $manager;

	private $tempFilename;
	private $images;
	private $sharedStrings;
	private $styles;
	private $worksheets;
	private $worksheetsIds;
	private $committed;
	private $tempFilenames;

	/**
	 * Workbook constructor.
	 *
	 * @param bool $useSharedStrings - принудительно записывать строки как общие. Проверять дубликаты
	 * @param bool|null $checkRelsDoubles - проверять дубликаты гиперссылок и картинок. По умолчанию = $useSharedStrings
	 */
	public function __construct(bool $useSharedStrings = false, ?bool $checkRelsDoubles = null) {
		$this->useSharedStrings = $useSharedStrings;
		$this->checkRelsDoubles = $checkRelsDoubles ?? $useSharedStrings;
		$this->tempdir = sys_get_temp_dir();
		$this->created = new DateTime();
		$this->modified = $this->created;
		$this->creator = 'topvisor/xlsx-writer ' . Workbook::VERSION;
		$this->lastModifiedBy = $this->creator;
		$this->company = '';
		$this->manager = null;

		$this->images = [];
		$this->sharedStrings = new SharedStrings($this);
		$this->styles = new Styles();
		$this->worksheets = [];
		$this->worksheetsIds = [];
		$this->committed = false;
		$this->tempFilenames = [];
	}

	public function __destruct() {
		unset($this->created);
		unset($this->modified);
		unset($this->sharedStrings);
		unset($this->styles);
		unset($this->worksheets);

		if ($this->tempFilename && file_exists($this->tempFilename)) unlink($this->tempFilename);

		$this->unlinkTempFiles();
	}

	/**
	 * @return bool - принудительно записывать строки как общие. Проверять дубликаты
	 */
	public function getUseSharedStrings(): bool {
		return $this->useSharedStrings;
	}

	/**
	 * @return bool - проверять дубликаты гиперссылок и картинок
	 */
	public function getCheckRelsDoubles(): bool {
		return $this->checkRelsDoubles;
	}

	/**
	 * @return string - путь к директории для хранения временных файлов
	 */
	public function getTempdir(): string {
		return $this->tempdir;
	}

	/**
	 * @param string $tempdir - путь к директории для хранения временных файлов
	 * @return Workbook - $this
	 * @throws ObjectCommittedException
	 */
	public function setTempdir(string $tempdir): Workbook {
		$this->checkCommitted();

		$this->tempdir = $tempdir;

		return $this;
	}

	/**
	 * @return DateTime - время создания файла
	 */
	public function getCreated(): DateTime {
		return $this->created;
	}

	/**
	 * @param DateTime $created - время создания файла
	 * @return Workbook - $this
	 * @throws ObjectCommittedException
	 */
	public function setCreated(DateTime $created): Workbook {
		$this->checkCommitted();

		$this->created = $created;

		return $this;
	}

	/**
	 * @return DateTime - время изменения файла
	 */
	public function getModified(): DateTime {
		return $this->modified;
	}

	/**
	 * @param DateTime $modified - время изменения файла
	 * @return Workbook - $this
	 * @throws ObjectCommittedException
	 */
	public function setModified(DateTime $modified): Workbook {
		$this->checkCommitted();

		$this->modified = $modified;

		return $this;
	}

	/**
	 * @return string - автор файла
	 */
	public function getCreator(): string {
		return $this->creator;
	}

	/**
	 * @param string $creator - автор файла
	 * @return Workbook - $this
	 * @throws ObjectCommittedException
	 */
	public function setCreator(string $creator): Workbook {
		$this->checkCommitted();

		$this->creator = $creator;

		return $this;
	}

	/**
	 * @return string - автор последнего изменения файла
	 */
	public function getLastModifiedBy(): string {
		return $this->lastModifiedBy;
	}

	/**
	 * @param string $lastModifiedBy - автор последнего изменения файла
	 * @return Workbook - $this
	 * @throws ObjectCommittedException
	 */
	public function setLastModifiedBy(string $lastModifiedBy): Workbook {
		$this->checkCommitted();

		$this->lastModifiedBy = $lastModifiedBy;

		return $this;
	}

	/**
	 * @return string - компания
	 */
	public function getCompany(): string {
		return $this->company;
	}

	/**
	 * @param string $company - компания
	 * @return Workbook - $this
	 * @throws ObjectCommittedException
	 */
	public function setCompany(string $company): Workbook {
		$this->checkCommitted();

		$this->company = $company;

		return $this;
	}

	/**
	 * @return string|null - менеджер
	 */
	public function getManager() {
		return $this->manager;
	}

	/**
	 * @param string|null $manager - менеджер
	 * @return Workbook - $this
	 * @throws ObjectCommittedException
	 */
	public function setManager(?string $manager = null): Workbook {
		$this->checkCommitted();

		$this->manager = $manager;

		return $this;
	}

	/**
	 * @param string|RichTextValue $value - значение
	 * @return SharedStringValue - общая строка
	 * @throws InvalidValueException
	 */
	public function addSharedString($value): SharedStringValue {
		return $this->sharedStrings->add($value);
	}

	/**
	 * Добавить картинку.
	 *
	 * @param string $filename - имя файла
	 * @param string $extension - расширение файла
	 * @return array - картинка
	 * @throws InvalidValueException
	 */
	public function addImage(string $filename, string $extension = ''): array {
		$filename = realpath($filename);
		if (!$filename || !file_exists($filename)) throw new InvalidValueException('Invalid $filename');

		if (!$extension) $extension = pathinfo($filename, PATHINFO_EXTENSION);
		Validator::validate($extension, '$extension', self::VALID_IMAGES_EXTENSION);

		$id = isset($this->images[$filename]) ? $this->images[$filename]['id'] : (count($this->images) + 1);

		$image = ['id' => $id, 'localname' => "image$id.$extension"];
		$this->images[$filename] = $image;

		return $image;
	}

	/**
	 * Добавить таблицу
	 *
	 * @param string $name - имя таблицы
	 * @return Worksheet - таблица
	 * @throws ObjectCommittedException
	 * @throws InvalidValueException
	 */
	public function addWorksheet(string $name): Worksheet {
		$this->checkCommitted();

		if (mb_strlen($name) > 31) throw new InvalidValueException('The length $name must be less than 31');
		Validator::validateString($name, self::INVALID_WORKSHEET_NAME, '$name');

		$id = count($this->worksheets) + 1;
		$this->worksheetsIds[$name] = $id;

		$worksheet = new Worksheet($this, $this->styles, $id, $name);
		$this->worksheets[] = $worksheet;

		return $worksheet;
	}

	/**
	 * Получить таблицу по имени
	 *
	 * @param string $name - имя таблицы
	 * @return Worksheet - таблица
	 */
	public function getWorksheetByName(string $name): Worksheet {
		return $this->worksheets[$this->worksheetsIds[$name] - 1];
	}

	/**
	 * Получить таблицу по id
	 *
	 * @param int $id - id таблицы
	 * @return Worksheet - таблица
	 */
	public function getWorksheetById(int $id): Worksheet {
		return $this->worksheets[$id - 1];
	}

	/**
	 * Получить все таблицы
	 *
	 * @return array - массив таблиц
	 */
	public function getWorksheets(): array {
		return $this->worksheets;
	}

	/**
	 * @return bool - зафиксированы ли изменения
	 */
	public function isCommitted(): bool {
		return $this->committed;
	}

	/**
	 * Зафиксировать файл workbook
	 *
	 * @throws EmptyObjectException
	 * @throws ObjectCommittedException
	 */
	public function commit() {
		$this->checkCommitted();
		if (!count($this->worksheets)) throw new EmptyObjectException('Workbook is empty');

		$this->committed = true;
		$this->tempFilename = $this->genTempFilename(false, 'xlsx');

		$zip = new ZipArchive();
		$zip->open($this->tempFilename, ZipArchive::CREATE | ZipArchive::OVERWRITE);

		$this->addFilesToZip($zip);
		$this->addStringsToZip($zip);

		$zip->close();

		$this->unlinkTempFiles();
	}

	/**
	 * Сохранить workbook в файл. Фиксирует изменения
	 *
	 * @param string $filename - путь для созддания xlsx файла
	 * @throws EmptyObjectException
	 * @throws ObjectCommittedException
	 */
	public function toFile(string $filename) {
		if (!$this->committed) $this->commit();
		copy($this->tempFilename, $filename);
	}

	/**
	 * Записать workbook в stdout и установить необходимые для скачивания заголовки. Фиксирует изменения
	 *
	 * @param string name - имя скачеваемого файла
	 * @param bool exit - завершить выполнение скрипта
	 * @throws EmptyObjectException
	 * @throws ObjectCommittedException
	 */
	public function toHttp(string $filename, bool $exit = true) {
		if (!$this->committed) $this->commit();
		if (!preg_match('/\.xlsx$/', $filename)) $filename .= '.xlsx';
		if (ob_get_level()) ob_end_clean();

		header('Content-Description: File Transfer');
		header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
		header('Content-Disposition: attachment; filename="'.$filename.'"');
		header('Content-Transfer-Encoding: binary');
		header('Expires: 0');
		header('Cache-Control: must-revalidate');
		header('Pragma: public');
		header('Content-Length: '.stat($this->tempFilename)['size']);

		readfile($this->tempFilename);

		if ($exit) {
			unlink($this->tempFilename);
			exit();
		}
	}

	/**
	 * Удалить все временные файлы
	 */
	public function unlinkTempFiles() {
		foreach ($this->tempFilenames as $tempFilename) if (file_exists($tempFilename)) unlink($tempFilename);
	}

	/**
	 * @return string - имя временного файла
	 *
	 * @param bool $autoUnlink - удалять временный файл при создании фиксации workbook
	 * @param string $ext - расширение файла
	 */
	public function genTempFilename(bool $autoUnlink = true, string $ext = 'xml') {
		$filename = $this->tempdir . '/xlsxcreator_' . base64_encode(rand()) . ".$ext";
		if (file_exists($filename)) return $this->genTempFilename();

		if ($autoUnlink) $this->tempFilenames[] = $filename;

		return $filename;
	}

	/**
	 * Добавить временные файлы в xlsx.
	 *
	 * @param ZipArchive $zip - xlsx файл
	 * @throws ObjectCommittedException
	 */
	private function addFilesToZip(ZipArchive $zip) {
		foreach ($this->worksheets as $worksheet) {
			if (!$worksheet->isCommitted()) $worksheet->commit();
			$zip->addFile($worksheet->getFilename(), $worksheet->getLocalname());

			if ($sheetRelsFilename = $worksheet->getSheetRelsFilename())
				$zip->addFile($sheetRelsFilename, 'xl/worksheets/_rels/sheet' . $worksheet->getId() . '.xml.rels');

			if ($commentsFilenames = $worksheet->getCommentsFilenames()) {
				$zip->addFile($commentsFilenames['comments'], 'xl/comments' . $worksheet->getId() . '.xml');
				$zip->addFile($commentsFilenames['vml'], 'xl/drawings/vmlDrawing' . $worksheet->getId() . '.vml');
			}

			if ($drawingFilenames = $worksheet->getDrawingFilenames()) {
				$zip->addFile($drawingFilenames['drawing'], 'xl/drawings/drawing' . $worksheet->getId() . '.xml');
				$zip->addFile($drawingFilenames['rels'], 'xl/drawings/_rels/drawing' . $worksheet->getId() . '.xml.rels');
			}
		}

		foreach ($this->images as $filename => $image) $zip->addFile($filename, "xl/media/$image[localname]");

		if (!$this->sharedStrings->isCommitted()) $this->sharedStrings->commit();
		if (!$this->sharedStrings->isEmpty()) $zip->addFile($this->sharedStrings->getFilename(), 'xl/sharedStrings.xml');

		$stylesFilename = $this->genTempFilename();
		$this->styles->writeToFile($stylesFilename);
		$zip->addFile($stylesFilename, 'xl/styles.xml');
	}

	/**
	 * Сгенерировать файлы из строк и добавить в xlsx.
	 *
	 * @param ZipArchive $zip - xlsx файл
	 */
	private function addStringsToZip(ZipArchive $zip) {
		$zip->addFromString('_rels/.rels', (new RelationshipsXml())->toXml([
			['id' => 'rId1', 'type' => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument', 'target' => 'xl/workbook.xml'],
			['id' => 'rId2', 'type' => 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties', 'target' => 'docProps/core.xml'],
			['id' => 'rId3', 'type' => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties', 'target' => 'docProps/app.xml'],
		]));

		$zip->addFromString('[Content_Types].xml', (new ContentTypesXml())->toXml($this->getWorksheetsModels()));

		$zip->addFromString('docProps/app.xml', (new AppXml())->toXml([
			'worksheets' => $this->getWorksheetsModels(),
			'company' => $this->company,
			'manager' => $this->manager,
		]));

		$zip->addFromString('docProps/core.xml', (new CoreXml())->toXml($this->getModel()));
		$zip->addFromString('xl/_rels/workbook.xml.rels', (new RelationshipsXml())->toXml($this->genRelationships()));
		$zip->addFromString('xl/workbook.xml', (new WorkbookXml())->toXml($this->getWorksheetsModels()));
	}

	/**
	 * Генерирует связи между файлами для workbook. Подготавивает worksheets к коммиту workbook.
	 *
	 * @return array - связи между файлами
	 */
	private function genRelationships(): array {
		$count = 1;

		$relationships = [[
			'id' => 'rId' . $count++,
			'type' => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles',
			'target' => 'styles.xml',
		]];

		if (!$this->sharedStrings->isEmpty()) $relationships[] = [
			'id' => 'rId' . $count++,
			'type' => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings',
			'target' => 'sharedStrings.xml',
		];

		foreach ($this->worksheets as $worksheet) {
			$worksheet->setRId('rId' . $count++);

			$relationships[] = [
				'id' => $worksheet->getRId(),
				'type' => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',
				'target' => 'worksheets/sheet' . $worksheet->getId() . '.xml',
			];
		}

		return $relationships;
	}

	/**
	 * @return array - массив моделей worksheets
	 */
	private function getWorksheetsModels(): array {
		return array_map(
			function ($worksheet) {
				return $worksheet->getModel();
			},
			$this->worksheets
		);
	}

	/**
	 * @return array - модель workbook
	 */
	private function getModel() {
		return [
			'creator' => $this->creator,
			'lastModifiedBy' => $this->lastModifiedBy,
			'created' => $this->created,
			'modified' => $this->modified,
		];
	}

	/**
	 * @throws ObjectCommittedException
	 */
	private function checkCommitted() {
		if ($this->committed) throw new ObjectCommittedException("Workbook is committed");
	}
}
