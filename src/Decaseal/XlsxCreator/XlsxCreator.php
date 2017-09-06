<?php

/**
 * Библиотека для создания xlsx файлов
 *
 * @author decaseal <decaseal@gmail.com>
 * @version 1.0
 */

namespace Decaseal\XlsxCreator;

use DateTime;
use Decaseal\XlsxCreator\Xml\Core\AppXml;
use Decaseal\XlsxCreator\Xml\Core\ContentTypesXml;
use Decaseal\XlsxCreator\Xml\Core\CoreXml;
use Decaseal\XlsxCreator\Xml\Core\RelationshipsXml;
use Decaseal\XlsxCreator\Xml\Style\StylesXml;
use ZipArchive;

/**
 * Class XlsxCreator. Используйте его для создания xlsx файла.
 *
 * @package Decaseal\XlsxCreator
 */
class XlsxCreator{
	private $tempdir;
	private $created;
	private $modified;
	private $creator;
	private $lastModifiedBy;
	private $company;
	private $manager;

	private $stylesXml;
	private $worksheets;
	private $committed;
	private $nextId;
	private $zip;

	/**
	 * XlsxCreator constructor
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

		$this->tempdir = $tempdir ?? sys_get_temp_dir();
		$this->created = $created ?? new DateTime();
		$this->modified = $modified ?? $this->created;
		$this->creator = $creator ?? 'XlsxWriter';
		$this->lastModifiedBy = $lastModifiedBy ?? $this->creator;
		$this->company = $company ?? '';
		$this->manager = $manager ?? null;

		$this->stylesXml = new StylesXml();
		$this->worksheets = [];
		$this->committed = false;
		$this->nextId = 1;

		$this->zip = new ZipArchive();
		$this->zip->open($filename, ZipArchive::CREATE | ZipArchive::OVERWRITE);
	}

	/**
	 * Добавить таблицу
	 *
	 * @param string $name - имя таблицы
	 * @param string|null $tabColor - цвет вкладки в формате argb (#FFFFFF00)
	 * @param int $defaultRowHeight - высота строки в px
	 * @return Worksheet - таблица
	 */
	function addWorksheet(string $name, string $tabColor = null, int $defaultRowHeight = 15) : Worksheet{
		$worksheet = new Worksheet($this->nextId(), $name, $tabColor, $defaultRowHeight);
		$worksheets[$name] = $worksheet;

		return $worksheet;
	}

	/**
	 * Получить таблицу
	 *
	 * @param string $name - имя таблицы
	 * @return Worksheet - таблица
	 */
	function getWorksheet(string $name) : Worksheet{
		return $this->worksheets[$name];
	}

	/**
	 * Получить все таблицы
	 *
	 * @return array - массив таблиц
	 */
	function getWorksheets() : array{
		return $this->worksheets;
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

	function isCommitted() : bool{
		return $this->committed;
	}

	function commit(){
		if ($this->isCommitted() || !count($this->getWorksheets())) return;

		foreach ($this->worksheets as $worksheet) $worksheet->commit();

		$this->zip->addFile('./Xml/Static/theme1.xml', 'xl/theme/theme1.xml');
		$this->zip->addFromString('/_rels/.rels', (new RelationshipsXml())->toXml([
			['Id' => 'rId1', 'Type' => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument', 'Target' => 'xl/workbook.xml'],
			['Id' => 'rId2', 'Type' => 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties', 'Target' => 'docProps/core.xml'],
			['Id' => 'rId3', 'Type' => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties', 'Target' => 'docProps/app.xml']
		]));
		$this->zip->addFromString('[Content_Types].xml', (new ContentTypesXml())->toXml(['worksheets' => $this->getWorksheets()]));
		$this->zip->addFromString('docProps/app.xml', (new AppXml())->toXml([
			'worksheets' => $this->getWorksheets(),
			'company' => $this->company,
			'manager' => $this->manager
		]));
		$this->zip->addFromString('docProps/core.xml', (new CoreXml())->toXml($this));
		$this->zip->addFromString('xl/styles.xml', (new StylesXml())->toXml());

		$relationships = [
			['Id' => 'rId1', 'Type' => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles', 'Target' => 'styles.xml'],
			['Id' => 'rId2', 'Type' => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme', 'Target' => 'theme/theme1.xml']
		];
		foreach ($this->getWorksheets() as $worksheet) {
			$relationships[] = [
				'Id' => 'rId' . (count($relationships) + 1),
				'Type' => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',
				'Target' => 'worksheets/sheet' . $worksheet->getId() . '.xml'
			];
		}
		$this->zip->addFromString('/xl/_rels/workbook.xml.rels', (new RelationshipsXml())->toXml($relationships));

		### addWorkbookXml
	}

	private function nextId() : int{
		return $this->nextId++;
	}

	private function finalize(){

	}
}