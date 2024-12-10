<?php

namespace Topvisor\XlsxCreator\Helpers;

use Topvisor\XlsxCreator\Exceptions\ObjectCommittedException;
use Topvisor\XlsxCreator\Structures\Drawing\TwoCellAnchorXml;
use Topvisor\XlsxCreator\Structures\Range\Range;
use Topvisor\XlsxCreator\Worksheet;
use Topvisor\XlsxCreator\Xml\Core\Relationships\RelationshipXml;
use XMLWriter;

/**
 * Class Drawing. Рисунки таблицы.
 *
 * @package Topvisor\XlsxCreator\Helpers
 */
class Drawing {
	private $worksheet;

	private $empty;
	private $committed;
	private $nextRId;
	private $nextImageId;

	private $filename;
	private $xml;

	private $relsFilename;
	private $relsXml;

	/**
	 * Drawing constructor.
	 *
	 * @param Worksheet $worksheet - таблица
	 */
	public function __construct(Worksheet $worksheet) {
		$this->worksheet = $worksheet;

		$this->empty = true;
		$this->committed = false;
		$this->nextRId = 1;
		$this->nextImageId = 2;
	}

	public function __destruct() {
		unset($this->worksheet);
		unset($this->xml);
		unset($this->relsXml);
	}

	/**
	 * @return bool - рисунков нет
	 */
	public function isEmpty(): bool {
		return $this->empty;
	}

	/**
	 * @return string|null - имя файла рисунков
	 */
	public function getFilename() {
		if ($this->xml ?? false) $this->xml->flush();

		return $this->filename;
	}

	/**
	 * @return string|null - имя файла связей рисунков
	 */
	public function getRelsFilename() {
		if ($this->relsXml ?? false) $this->relsXml->flush();

		return $this->relsFilename;
	}

	/**
	 * @param array $image - картинка
	 * @param Range $position - местоположение в таблице
	 * @param string $name - имя картинки
	 */
	public function addImage(array $image, Range $position, string $name = '') {
		$this->checkCommited();
		if ($this->empty) $this->startDrawing();

		if (!$name) $name = "Image$image[id]";
		$rId = 'rId' . $this->nextRId++;

		(new RelationshipXml())->render($this->relsXml, [
			'id' => $rId,
			'type' => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
			'target' => "../media/$image[localname]",
		]);

		(new TwoCellAnchorXml())->render($this->xml, [
			'pic' => [
				'id' => $this->nextImageId++,
				'name' => $name,
				'rId' => $rId,
			],
			'position' => $position->getModel(),
		]);
	}

	/**
	 * Зафиксировать рисунки
	 */
	public function commit() {
		$this->checkCommited();
		$this->committed = true;

		$this->endDrawing();
	}

	/**
	 * Начать файлы рисунков и связей рисунков
	 */
	private function startDrawing() {
		$this->empty = false;

		$this->filename = $this->worksheet->getWorkbook()->genTempFilename();
		$this->xml = new XMLWriter();
		$this->xml->openURI($this->filename);

		$this->xml->startDocument('1.0', 'UTF-8', 'yes');
		$this->xml->startElement('xdr:wsDr');
		$this->xml->writeAttribute('xmlns:xdr', 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing');
		$this->xml->writeAttribute('xmlns:a', 'http://schemas.openxmlformats.org/drawingml/2006/main');

		$this->relsFilename = $this->worksheet->getWorkbook()->genTempFilename();
		$this->relsXml = new XMLWriter();
		$this->relsXml->openURI($this->relsFilename);

		$this->relsXml->startDocument('1.0', 'UTF-8', 'yes');
		$this->relsXml->startElement('Relationships');
		$this->relsXml->writeAttribute('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');
	}

	/**
	 * Закончить файлы рисунков и связей рисунков
	 */
	private function endDrawing() {
		if ($this->empty) return;

		$this->xml->endElement();
		$this->xml->endDocument();

		$this->xml->flush();
		unset($this->xml);

		$this->relsXml->endElement();
		$this->relsXml->endDocument();

		$this->relsXml->flush();
		unset($this->relsXml);
	}

	/**
	 * @throws ObjectCommittedException
	 */
	private function checkCommited() {
		if ($this->committed) throw new ObjectCommittedException();
	}
}
