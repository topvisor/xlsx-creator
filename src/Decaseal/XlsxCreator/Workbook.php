<?php

/**
 * Библиотека для создания xlsx файлов
 *
 * @author decaseal <decaseal@gmail.com>
 * @version v0.1-alpha
 */

namespace Decaseal\XlsxCreator;

use DateTime;
use Decaseal\XlsxCreator\Xml\Book\WorkbookXml;
use Decaseal\XlsxCreator\Xml\Core\App\AppXml;
use Decaseal\XlsxCreator\Xml\Core\ContentTypesXml;
use Decaseal\XlsxCreator\Xml\Core\CoreXml;
use Decaseal\XlsxCreator\Xml\Core\Relationships\RelationshipsXml;
use Decaseal\XlsxCreator\Xml\Styles\StylesXml;
use ZipArchive;

/**
 * Class Workbook. Используйте его для создания xlsx файла.
 *
 * @package Decaseal\XlsxCreator
 */
class Workbook{
	private $filename;
	private $tempdir;
	private $created;
	private $modified;
	private $creator;
	private $lastModifiedBy;
	private $company;
	private $manager;

	private $stylesXml;
	private $worksheets;
	private $worksheetsIds;
	private $committed;
	private $tempFilenames;

	/**
	 * Workbook constructor.
	 *
	 * @param string $filename - путь к xlsx файлу
	 */
	function __construct(string $filename){
		$this->filename = $filename;
		$this->tempdir = sys_get_temp_dir();
		$this->created = new DateTime();
		$this->modified = $this->created;
		$this->creator = 'XlsxWriter';
		$this->lastModifiedBy = $this->creator;
		$this->company = '';
		$this->manager = null;

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
	 */
	function setFilename(string $filename) : Workbook{
		$this->filename = $filename;
		return $this;
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
	 */
	function setTempdir(string $tempdir) : Workbook{
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
	 */
	function setCreated(DateTime $created) : Workbook{
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
	 */
	function setModified(DateTime $modified) : Workbook{
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
	 */
	function setCreator(string $creator) : Workbook{
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
	 */
	function setLastModifiedBy(string $lastModifiedBy) : Workbook{
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
	 */
	function setCompany(string $company) : Workbook{
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
	 */
	function setManager(string $manager = null) : Workbook{
		$this->manager = $manager;
		return $this;
	}

	/**
	 * Добавить таблицу
	 *
	 * @param string $name - имя таблицы
	 * @return Worksheet - таблица
	 */
	function addWorksheet(string $name) : Worksheet{
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
	 */
	function commit(){
		if ($this->committed || !count($this->worksheets)) return;
		$this->committed = true;

		$zip = new ZipArchive();
		$zip->open($this->filename, ZipArchive::CREATE | ZipArchive::OVERWRITE);

		foreach ($this->worksheets as $worksheet){
			$worksheet->commit();
			$zip->addFile($worksheet->getFilename(), $worksheet->getLocalname());

			$sheetRelsFilename = $worksheet->getSheetRels()->getFilename();
			if ($sheetRelsFilename)	$zip->addFile($sheetRelsFilename, $worksheet->getSheetRels()->getLocalname());
		}

		$zip->addFile(dirname(__FILE__) . '/Xml/Static/theme1.xml', 'xl/theme/theme1.xml');
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
		$zip->addFromString('xl/styles.xml', (new StylesXml())->toXml());
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
			['Id' => 'rId' . $count++, 'Type' => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles', 'Target' => 'styles.xml'],
			['Id' => 'rId' . $count++, 'Type' => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme', 'Target' => 'theme/theme1.xml']
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
}