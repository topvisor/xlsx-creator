<?php

/**
 * Библиотека для создания xlsx файлов
 *
 * @author decaseal <decaseal@gmail.com>
 * @version v0.3-alpha
 */

namespace Topvisor\XlsxCreator;

use DateTime;
use Topvisor\XlsxCreator\Exceptions\ObjectCommittedException;
use Topvisor\XlsxCreator\Exceptions\EmptyObjectException;
use Topvisor\XlsxCreator\Helpers\SharedStrings;
use Topvisor\XlsxCreator\Structures\Values\SharedStringValue;
use Topvisor\XlsxCreator\Xml\Book\WorkbookXml;
use Topvisor\XlsxCreator\Xml\Core\App\AppXml;
use Topvisor\XlsxCreator\Xml\Core\ContentTypesXml;
use Topvisor\XlsxCreator\Xml\Core\CoreXml;
use Topvisor\XlsxCreator\Xml\Core\Relationships\RelationshipsXml;
use Topvisor\XlsxCreator\Xml\Styles\StylesXml;
use ZipArchive;

/**
 * Class Workbook. Используйте его для создания xlsx файла.
 *
 * @package  Topvisor\XlsxCreator
 */
class Workbook{
	private $filename;
	private $useSharedStrings;
	private $tempdir;
	private $created;
	private $modified;
	private $creator;
	private $lastModifiedBy;
	private $company;
	private $manager;

	private $sharedStrings;
	private $stylesXml;
	private $worksheets;
	private $worksheetsIds;
	private $committed;
	private $tempFilenames;

	/**
	 * Workbook constructor.
	 *
	 * @param string $filename - путь к xlsx файлу
	 * @param bool $useSharedStrings - принудительно записывать строки как общие. Проверять дубликаты
	 */
	function __construct(string $filename, bool $useSharedStrings = false){
		$this->filename = $filename;
		$this->useSharedStrings = $useSharedStrings;
		$this->tempdir = sys_get_temp_dir();
		$this->created = new DateTime();
		$this->modified = $this->created;
		$this->creator = 'XlsxWriter';
		$this->lastModifiedBy = $this->creator;
		$this->company = '';
		$this->manager = null;

		$this->sharedStrings = new SharedStrings($this);
		$this->stylesXml = new StylesXml();
		$this->worksheets = [];
		$this->worksheetsIds = [];
		$this->committed = false;
		$this->tempFilenames = [];
	}

	function __destruct(){
		unset($this->created);
		unset($this->modified);
		unset($this->stylesXml);
		unset($this->worksheets);

		$this->unlinkTempFiles();
	}

	/**
	 * @return string - путь к xlsx файлу
	 */
	function getFilename() : string{
		return $this->filename;
	}

	/**
	 * @param string $filename - путь к xlsx файлу
	 * @return Workbook - $this
	 * @throws ObjectCommittedException
	 */
	function setFilename(string $filename) : Workbook{
		$this->checkCommitted();

		$this->filename = $filename;
		return $this;
	}

	/**
	 * @return bool - принудительно записывать строки как общие. Проверять дубликаты
	 */
	function getUseSharedStrings() : bool{
		return $this->useSharedStrings;
	}

	/**
	 * @return string - путь к директории для хранения временных файлов
	 */
	function getTempdir() : string{
		return $this->tempdir;
	}

	/**
	 * @param string $tempdir - путь к директории для хранения временных файлов
	 * @return Workbook - $this
	 * @throws ObjectCommittedException
	 */
	function setTempdir(string $tempdir) : Workbook{
		$this->checkCommitted();

		$this->tempdir = $tempdir;
		return $this;
	}

	/**
	 * @return DateTime - время создания файла
	 */
	function getCreated() : DateTime{
		return $this->created;
	}

	/**
	 * @param DateTime $created - время создания файла
	 * @return Workbook - $this
	 * @throws ObjectCommittedException
	 */
	function setCreated(DateTime $created) : Workbook{
		$this->checkCommitted();

		$this->created = $created;
		return $this;
	}

	/**
	 * @return DateTime - время изменения файла
	 */
	function getModified() : DateTime{
		return $this->modified;
	}

	/**
	 * @param DateTime $modified - время изменения файла
	 * @return Workbook - $this
	 * @throws ObjectCommittedException
	 */
	function setModified(DateTime $modified) : Workbook{
		$this->checkCommitted();

		$this->modified = $modified;
		return $this;
	}

	/**
	 * @return string - автор файла
	 */
	function getCreator() : string{
		return $this->creator;
	}

	/**
	 * @param string $creator - автор файла
	 * @return Workbook - $this
	 * @throws ObjectCommittedException
	 */
	function setCreator(string $creator) : Workbook{
		$this->checkCommitted();

		$this->creator = $creator;
		return $this;
	}

	/**
	 * @return string - автор последнего изменения файла
	 */
	function getLastModifiedBy() : string{
		return $this->lastModifiedBy;
	}

	/**
	 * @param string $lastModifiedBy - автор последнего изменения файла
	 * @return Workbook - $this
	 * @throws ObjectCommittedException
	 */
	function setLastModifiedBy(string $lastModifiedBy) : Workbook{
		$this->checkCommitted();

		$this->lastModifiedBy = $lastModifiedBy;
		return $this;
	}

	/**
	 * @return string - компания
	 */
	function getCompany() : string{
		return $this->company;
	}

	/**
	 * @param string $company - компания
	 * @return Workbook - $this
	 * @throws ObjectCommittedException
	 */
	function setCompany(string $company) : Workbook{
		$this->checkCommitted();

		$this->company = $company;
		return $this;
	}

	/**
	 * @return string|null - менеджер
	 */
	function getManager(){
		return $this->manager;
	}

	/**
	 * @param string|null $manager - менеджер
	 * @return Workbook - $this
	 * @throws ObjectCommittedException
	 */
	function setManager(string $manager = null) : Workbook{
		$this->checkCommitted();

		$this->manager = $manager;
		return $this;
	}

	function addSharedString($value) : SharedStringValue{
		return $this->sharedStrings->add($value);
	}

	/**
	 * Добавить таблицу
	 *
	 * @param string $name - имя таблицы
	 * @return Worksheet - таблица
	 * @throws ObjectCommittedException
	 */
	function addWorksheet(string $name) : Worksheet{
		$this->checkCommitted();

		$id = count($this->worksheets) + 1;
		$this->worksheetsIds[$name] = $id;

		$worksheet = new Worksheet($this, $id, $name);
		$this->worksheets[] = $worksheet;

		return $worksheet;
	}

	/**
	 * Получить таблицу по имени
	 *
	 * @param string $name - имя таблицы
	 * @return Worksheet - таблица
	 */
	function getWorksheetByName(string $name) : Worksheet{
		return $this->worksheets[$this->worksheetsIds[$name] - 1];
	}

	/**
	 * Получить таблицу по id
	 *
	 * @param int $id - id таблицы
	 * @return Worksheet - таблица
	 */
	function getWorksheetById(int $id) : Worksheet{
		return $this->worksheets[$id - 1];
	}

	/**
	 * Получить все таблицы
	 *
	 * @return array - массив таблиц
	 */
	function getWorksheets() : array{
		return $this->worksheets;
	}

	/**
	 * @return bool - зафиксированы ли изменения
	 */
	function isCommitted() : bool{
		return $this->committed;
	}

	/**
	 * Зафиксировать файл workbook. Используйте для создания xlsx файла.
	 * @throws EmptyObjectException
	 * @throws ObjectCommittedException
	 */
	function commit(){
		$this->checkCommitted();
		if (!count($this->worksheets)) throw new EmptyObjectException('Workbook is empty');

		$this->committed = true;

		$zip = new ZipArchive();
		$zip->open($this->filename, ZipArchive::CREATE | ZipArchive::OVERWRITE);

		foreach ($this->worksheets as $worksheet){
			if (!$worksheet->isCommitted()) $worksheet->commit();
			$zip->addFile($worksheet->getFilename(), $worksheet->getLocalname());

			$sheetRelsFilename = $worksheet->getSheetRels()->getFilename();
			if ($sheetRelsFilename)	$zip->addFile($sheetRelsFilename, $worksheet->getSheetRels()->getLocalname());
		}

		$zip->addFromString('_rels/.rels', (new RelationshipsXml())->toXml([
			['Id' => 'rId1', 'Type' => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument', 'Target' => 'xl/workbook.xml'],
			['Id' => 'rId2', 'Type' => 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties', 'Target' => 'docProps/core.xml'],
			['Id' => 'rId3', 'Type' => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties', 'Target' => 'docProps/app.xml']
		]));
		$zip->addFromString('[Content_Types].xml', (new ContentTypesXml())->toXml($this->getWorksheetsModels()));
		$zip->addFromString('docProps/app.xml', (new AppXml())->toXml([
			'worksheets' => $this->getWorksheetsModels(),
			'company' => $this->company,
			'manager' => $this->manager
		]));
		$zip->addFromString('docProps/core.xml', (new CoreXml())->toXml($this->getModel()));
		$zip->addFromString('xl/styles.xml', $this->stylesXml->toXml());
		$zip->addFromString('xl/_rels/workbook.xml.rels', (new RelationshipsXml())->toXml($this->genRelationships()));
		$zip->addFromString('xl/workbook.xml', (new WorkbookXml())->toXml($this->getWorksheetsModels()));

		$zip->close();

		$this->unlinkTempFiles();
	}

	/**
	 * Удалить все временные файлы
	 */
	function unlinkTempFiles(){
		foreach ($this->tempFilenames as $tempFilename) if (file_exists($tempFilename)) unlink($tempFilename);
	}

	/**
	 * @return StylesXml - менеджер стилей
	 */
	function getStyles() : StylesXml{
		return $this->stylesXml;
	}

	/**
	 * @return string - имя временного файла
	 */
	function genTempFilename(){
		$filename = $this->tempdir . '/xlsxcreator_' . base64_encode(rand()) . '.xml';
		if (file_exists($filename)) $filename = $this->genTempFilename();

		$this->tempFilenames[] = $filename;
		return $filename;
	}

	/**
	 * Генерирует связи между файлами для workbook. Подготавивает worksheets к коммиту workbook.
	 *
	 * @return array - связи между файлами
	 */
	private function genRelationships() : array{
		$count = 1;

		$relationships = [
			['Id' => 'rId' . $count++, 'Type' => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles', 'Target' => 'styles.xml']
		];

		foreach ($this->worksheets as $worksheet) {
			$worksheet->setRId('rId' . $count++);

			$relationships[] = [
				'Id' => $worksheet->getRId(),
				'Type' => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',
				'Target' => 'worksheets/sheet' . $worksheet->getId() . '.xml'
			];
		}

		return $relationships;
	}

	/**
	 * @return array - массив моделей worksheets
	 */
	private function getWorksheetsModels() : array{
		return array_map(
			function($worksheet){
				return $worksheet->getModel();
			},
			$this->worksheets
		);
	}

	/**
	 * @return array - модель workbook
	 */
	private function getModel(){
		return [
			'creator' => $this->creator,
			'lastModifiedBy' => $this->lastModifiedBy,
			'created' => $this->created,
			'modified' => $this->modified
		];
	}

	/**
	 * @throws ObjectCommittedException
	 */
	private function checkCommitted(){
		if ($this->committed) throw new ObjectCommittedException("Workbook is committed");
	}
}