<?php

namespace Topvisor\XlsxCreator\Helpers;

use Topvisor\XlsxCreator\Exceptions\ObjectCommittedException;
use Topvisor\XlsxCreator\Worksheet;
use XMLWriter;

/**
 * Class Worksheet. Служит для добавления связей в таблицу.
 *
 * @package  Topvisor\XlsxCreator
 */
class SheetRels{
	private $worksheet;

	private $indexes;
	private $hyperlinks;
	private $committed;

	private $filename;
	private $xml;

	/**
	 * SheetRels constructor.
	 * @param Worksheet $worksheet - таблица
	 */
	function __construct(Worksheet $worksheet){
		$this->worksheet = $worksheet;

		$this->indexes = [];
		$this->hyperlinks = [];
		$this->committed = false;
	}

	function __destruct(){
		unset($this->xml);

		if ($this->filename && file_exists($this->filename)) unlink($this->filename);
	}

	/**
	 * @return array - список гиперссылок
	 */
	function getHyperlinks() : array{
		return $this->hyperlinks;
	}

	/**
	 * @return null|string - путь к временному файлу связей
	 */
	function getFilename(){
		return $this->filename;
	}

	/**
	 * @return string - путь к файлу связей внутри xlsx файла
	 */
	function getLocalname() : string{
		return 'xl/worksheets/_rels/sheet' . $this->worksheet->getId() . '.xml.rels';
	}

	/**
	 * @param string $target - гиперссылка
	 * @param string $address - ячейка таблицы ('A1', 'B5')
	 */
	function addHyperlink(string $target, string $address){
		$index = ['target' => $target, 'address' => $address];
		if (in_array($index, $this->indexes)) return;

		$this->indexes[] = $index;

		$this->hyperlinks[] = [
			'address' => $address,
			'rId' => $this->writeRelationship('http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink', $target, 'External')
		];
	}

	/**
	 * Зафиксировать файл связей.
	 * @throws ObjectCommittedException
	 */
	function commit() {
		if (!$this->xml || $this->committed) throw new ObjectCommittedException();
		$this->committed = true;

		$this->endSheetRels();
	}

	/**
	 *	Начать файл связей.
	 */
	private function startSheetRels(){
		$this->filename = $this->worksheet->getWorkbook()->genTempFilename();

		$this->xml = new XMLWriter();
		$this->xml->openURI($this->filename);

		$this->xml->startDocument('1.0', 'UTF-8', 'yes');
		$this->xml->startElement('Relationships');
		$this->xml->writeAttribute('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');
	}

	/**
	 *	Завершить файл связей.
	 */
	private function endSheetRels(){
		$this->xml->endElement();
		$this->xml->endDocument();

		$this->xml->flush();
		unset($this->xml);
	}

	/**
	 * Записать связь в файл.
	 *
	 * @param string $type - схема связи
	 * @param string $target - связь
	 * @param string|null $targetMode - тип связи
	 * @return string - id связи
	 */
	private function writeRelationship(string $type, string $target, string $targetMode = null) : string{
		if (!$this->xml) $this->startSheetRels();

		$rId = 'rId' . (count($this->hyperlinks) + 1);

		$this->xml->startElement('Relationship');

		$this->xml->writeAttribute('Id', $rId);
		$this->xml->writeAttribute('Type', $type);
		$this->xml->writeAttribute('Target', $target);
		if ($targetMode) $this->xml->writeAttribute('TargetMode', $targetMode);

		$this->xml->endElement();

		return $rId;
	}
}