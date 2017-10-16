<?php

namespace Topvisor\XlsxCreator\Helpers;

use Topvisor\XlsxCreator\Exceptions\ObjectCommittedException;
use Topvisor\XlsxCreator\Worksheet;
use Topvisor\XlsxCreator\Xml\Sheet\HyperlinkXml;
use XMLWriter;

/**
 * Class Worksheet. Служит для добавления связей в таблицу.
 *
 * @package  Topvisor\XlsxCreator
 */
class SheetRels{
	private $worksheet;

	private $committed;
	private $nextId;

	private $filename;
	private $xml;

	private $hyperlinksFilename;
	private $hyperlinksXml;

	/**
	 * SheetRels constructor.
	 * @param Worksheet $worksheet - таблица
	 */
	function __construct(Worksheet $worksheet){
		$this->worksheet = $worksheet;

		$this->committed = false;
		$this->nextId = 1;
	}

	function __destruct(){
		unset($this->xml);
		unset($this->hyperlinksXml);

		if ($this->filename && file_exists($this->filename)) unlink($this->filename);
		if ($this->hyperlinksFilename && file_exists($this->hyperlinksFilename)) unlink($this->hyperlinksFilename);
	}

	/**
	 * @return null|string - путь к временному файлу связей
	 */
	function getFilename(){
		if ($this->xml ?? false) $this->xml->flush();
		return $this->filename;
	}

	/**
	 * @return null|string - путь к временному файлу гиперссылок
	 */
	function getHyperlinksFilename(){
		if ($this->hyperlinksXml ?? false) $this->hyperlinksXml->flush();
		return $this->hyperlinksFilename;
	}

	/**
	 * @param string $target - гиперссылка
	 * @param string $address - ячейка таблицы ('A1', 'B5')
	 */
	function addHyperlink(string $target, string $address){
		$this->checkCommited();

		$hyperlink = [
			'address' => $address,
			'rId' => $this->writeRelationship(
				'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
				$target,
				'External'
			)
		];

		if (!$this->hyperlinksXml ?? false) $this->startHyperlinks();
		(new HyperlinkXml())->render($this->hyperlinksXml, $hyperlink);
	}

	/**
	 * Связать таблицу с файлом комментариев и его vml представлением
	 * @return string - id связи с vml файлом
	 */
	function addComments() : string{
		$this->checkCommited();

		$this->writeRelationship(
			'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments',
			'../comments' . $this->worksheet->getId() . '.xml'
		);

		return $this->writeRelationship(
			'http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing',
			'../drawings/vmlDrawing' . $this->worksheet->getId() . '.vml'
		);
	}

	/**
	 * Связать таблицу с файлом рисунков
	 * @return string - id связи с файлом рисунков
	 */
	function addDrawing() : string{
		$this->checkCommited();

		return $this->writeRelationship(
			'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing',
			'../drawings/drawing' . $this->worksheet->getId() . '.xml'
		);
	}

	/**
	 * Зафиксировать файл связей.
	 * @throws ObjectCommittedException
	 */
	function commit() {
		$this->checkCommited();
		$this->committed = true;

		$this->endSheetRels();
		$this->endHyperlinks();
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
		if (!($this->xml ?? false)) return;

		$this->xml->endElement();
		$this->xml->endDocument();

		$this->xml->flush();
		unset($this->xml);
	}

	private function startHyperlinks(){
		$this->hyperlinksFilename = $this->worksheet->getWorkbook()->genTempFilename();

		$this->hyperlinksXml = new XMLWriter();
		$this->hyperlinksXml->openURI($this->hyperlinksFilename);
	}

	private function endHyperlinks(){
		if (!($this->hyperlinksXml ?? false)) return;

		$this->hyperlinksXml->flush();
		unset($this->hyperlinksXml);
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
		if (!($this->xml ?? false)) $this->startSheetRels();

		$rId = 'rId' . $this->nextId++;

		$this->xml->startElement('Relationship');

		$this->xml->writeAttribute('Id', $rId);
		$this->xml->writeAttribute('Type', $type);
		$this->xml->writeAttribute('Target', $target);
		if ($targetMode) $this->xml->writeAttribute('TargetMode', $targetMode);

		$this->xml->endElement();

		return $rId;
	}

	/**
	 * @throws ObjectCommittedException
	 */
	private function checkCommited(){
		if ($this->committed) throw new ObjectCommittedException();
	}
}