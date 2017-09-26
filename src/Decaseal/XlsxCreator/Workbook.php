<?php

/**
 * Библиотека для создания xlsx файлов
 *
 * @author decaseal <decaseal@gmail.com>
 * @version 1.0
 */

namespace Decaseal\XlsxCreator;

use DateTime;
use Decaseal\XlsxCreator\Xml\Book\Workbook\WorkbookXml;
use Decaseal\XlsxCreator\Xml\Core\App\AppXml;
use Decaseal\XlsxCreator\Xml\Core\ContentTypesXml;
use Decaseal\XlsxCreator\Xml\Core\CoreXml;
use Decaseal\XlsxCreator\Xml\Core\Relationships\RelationshipsXml;
use Decaseal\XlsxCreator\Xml\Styles\Index\StylesIndex;
use Decaseal\XlsxCreator\Xml\Styles\StylesXml;
use ZipArchive;

/**
 * Class Workbook. Используйте его для создания xlsx файла.
 *
 * @package Decaseal\Workbook
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
	 * Workbook constructor
	 *
	 * @param string $filename - путь к xlsx файлу
	 * @param string|null $tempdir - путь к директории для хранения временных файлов
	 * @param DateTime|null $created - время создания файла
	 * @param DateTime|null $modified - время изменения файла
	 * @param string|null $creator - автор файла
	 * @param string|null $lastModifiedBy - автор последнего изменения файла
	 * @param string|null $company - компания
	 * @param string|null $manager - менеджер
	 */
	function __construct(string $filename, string $tempdir = null, DateTime $created = null, DateTime $modified = null, string $creator = null,
						 string $lastModifiedBy = null, string $company = null, string $manager = null){

		$this->filename = $filename;
		$this->tempdir = $tempdir ?? sys_get_temp_dir();
		$this->created = $created ?? new DateTime();
		$this->modified = $modified ?? $this->created;
		$this->creator = $creator ?? 'XlsxWriter';
		$this->lastModifiedBy = $lastModifiedBy ?? $this->creator;
		$this->company = $company ?? '';
		$this->manager = $manager ?? null;

		$this->stylesXml = new StylesXml();
		$this->worksheets = [];
		$this->worksheetsIds = [];
		$this->committed = false;
		$this->tempFilenames = [];
	}

	public function __destruct(){
		unset($this->created);
		unset($this->modified);
		unset($this->stylesXml);
		unset($this->worksheets);

		$this->unlinkTempFiles();
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
	 * Получить таблицу по имени
	 *
	 * @param int $id - имя таблицы
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

	function getTempdir() : string{
		return $this->tempdir;
	}

	function getCreator() : string{
		return $this->creator;
	}

	function getLastModifiedBy() : string{
		return $this->lastModifiedBy;
	}

	function getCreated() : DateTime{
		return $this->created;
	}

	function getModified() : DateTime{
		return $this->modified;
	}

	function getStyles() : StylesXml{
		return $this->stylesXml;
	}

	function isCommitted() : bool{
		return $this->committed;
	}

	function commit(){
		if ($this->isCommitted() || !count($this->worksheets)) return;
		$this->committed = true;

		$zip = new ZipArchive();
		$zip->open($this->filename, ZipArchive::CREATE | ZipArchive::OVERWRITE);

		foreach ($this->worksheets as $worksheet){
			$worksheet->commit();
			$zip->addFile($worksheet->getFilename(), $worksheet->getLocalname());
			$zip->addFile($worksheet->getSheetRels()->getFilename(), $worksheet->getSheetRels()->getLocalname());
		}

		$zip->addFile(dirname(__FILE__) . '/Xml/Static/theme1.xml', '/xl/theme/theme1.xml');
		$zip->addFromString('/_rels/.rels', (new RelationshipsXml())->toXml([
			['Id' => 'rId1', 'Type' => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument', 'Target' => 'xl/workbook.xml'],
			['Id' => 'rId2', 'Type' => 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties', 'Target' => 'docProps/core.xml'],
			['Id' => 'rId3', 'Type' => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties', 'Target' => 'docProps/app.xml']
		]));
		$zip->addFromString('/[Content_Types].xml', (new ContentTypesXml())->toXml($this->getWorksheetsModels()));
		$zip->addFromString('/docProps/app.xml', (new AppXml())->toXml([
			'worksheets' => $this->getWorksheetsModels(),
			'company' => $this->company,
			'manager' => $this->manager
		]));
		$zip->addFromString('/docProps/core.xml', (new CoreXml())->toXml($this->getModel()));
		$zip->addFromString('/xl/styles.xml', (new StylesXml())->toXml());
		$zip->addFromString('/xl/_rels/workbook.xml.rels', (new RelationshipsXml())->toXml($this->genRelationships()));
		$zip->addFromString('/xl/workbook.xml', (new WorkbookXml())->toXml($this->getWorksheetsModels()));

		$zip->close();

		$this->unlinkTempFiles();
	}

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

	private function getWorksheetsModels() : array{
		return array_map(
			function($worksheet){
				return $worksheet->getModel();
			},
			$this->worksheets
		);
	}

	private function getModel(){
		return [
			'creator' => $this->creator,
			'lastModifiedBy' => $this->lastModifiedBy,
			'created' => $this->created,
			'modified' => $this->modified
		];
	}

	function unlinkTempFiles(){
		foreach ($this->tempFilenames as $tempFilename) if (file_exists($tempFilename)) unlink($tempFilename);
	}

	function genTempFilename(){
		$filename = $this->tempdir . '/xlsxcreator_' . base64_encode(rand()) . '.xml';
		if (file_exists($filename)) $filename = $this->genTempFilename();

		$this->tempFilenames[] = $filename;
		return $filename;
	}
}