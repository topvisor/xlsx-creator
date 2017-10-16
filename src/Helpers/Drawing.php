<?php

namespace Topvisor\XlsxCreator\Helpers;

use Topvisor\XlsxCreator\Exceptions\ObjectCommittedException;
use Topvisor\XlsxCreator\Structures\Range\Range;
use Topvisor\XlsxCreator\Worksheet;
use XMLWriter;

/**
 * Class Drawing. Рисунки таблицы.
 *
 * @package Topvisor\XlsxCreator\Helpers
 */
class Drawing{
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
	function __construct(Worksheet $worksheet){
		$this->worksheet = $worksheet;

		$this->empty = true;
		$this->committed = false;
		$this->nextRId = 1;
		$this->nextImageId = 2;
	}

	function __destruct(){
		unset($this->worksheet);
		unset($this->xml);
		unset($this->relsXml);
	}

	/**
	 * @return bool - рисунков нет
	 */
	function isEmpty(): bool{
		return $this->empty;
	}

	/**
	 * @return string|null - имя файла рисунков
	 */
	function getFilename(){
		if ($this->xml ?? false) $this->xml->flush();
		return $this->filename;
	}

	/**
	 * @return string|null - имя файла связей рисунков
	 */
	function getRelsFilename(){
		if ($this->relsXml ?? false) $this->relsXml->flush();
		return $this->relsFilename;
	}

	/**
	 * @param array $image - картинка
	 * @param Range $position - местоположение в таблице
	 * @param string $name - имя картинки
	 */
	function addImage(array $image, Range $position, string $name = ''){
		$this->checkCommited();
		if ($this->empty) $this->startDrawing();

		if (!$name) $name = "Image$image[id]";
		$rId = 'rId' . $this->nextRId++;

		$this->writeRelationships($rId, $image['localname']);
		$this->writeTwoCellAnchor($position, $name, $rId);
	}

	/**
	 * Зафиксировать рисунки
	 */
	function commit(){
		$this->checkCommited();
		$this->committed = true;

		$this->endDrawing();
	}

	/**
	 * @param string $rId - ид связи
	 * @param string $localname - имя файла картинки
	 */
	private function writeRelationships(string $rId, string $localname){
		$this->relsXml->startElement('Relationship');
		$this->relsXml->writeAttribute('Id', $rId);
		$this->relsXml->writeAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image');
		$this->relsXml->writeAttribute('Target', "../media/$localname");
		$this->relsXml->endElement();
	}

	/**
	 * @param Range $position - позиция картинки
	 * @param string $name - имя картинки
	 * @param string $rId - ид связи
	 */
	private function writeTwoCellAnchor(Range $position, string $name, string $rId){
		$this->xml->startElement('xdr:twoCellAnchor');
		$this->xml->writeAttribute('editAs', 'oneCell');

		$this->writeCellPosition('xdr:from', $position->getLeft(), $position->getTop());
		$this->writeCellPosition('xdr:to', $position->getRight(), $position->getBottom());

		$this->writePic($name, $rId);

		$this->xml->writeElement('xdr:clientData');

		$this->xml->endElement();
	}

	/**
	 * @param string $tag - тег
	 * @param float $col - колонка
	 * @param float $row - строка
	 */
	private function writeCellPosition(string $tag, float $col, float $row){
		$this->xml->startElement($tag);

		$this->xml->writeElement('xdr:col', $intCol = (int) floor($col));
		$this->xml->writeElement('xdr:colOff', (int) floor(($col - $intCol) * 640000));
		$this->xml->writeElement('xdr:row', $intRow = (int) floor(($row)));
		$this->xml->writeElement('xdr:rowOff', (int) floor(($row - $intRow) * 180000));

		$this->xml->endElement();
	}

	/**
	 * @param string $name - имя картинки
	 * @param string $rId - ид связи
	 */
	private function writePic(string $name, string $rId){
		$this->xml->startElement('xdr:pic');

		$this->writeNvPicPr($name);
		$this->writeBlipFill($rId);

		$this->xml->startElement('xdr:spPr');
		$this->xml->startElement('a:xfrm');

		$this->xml->startElement('a:off');
		$this->xml->writeAttribute('x', 0);
		$this->xml->writeAttribute('y', 0);
		$this->xml->endElement();

		$this->xml->startElement('a:ext');
		$this->xml->writeAttribute('cx', 2057400);
		$this->xml->writeAttribute('cy', 528034);
		$this->xml->endElement();

		$this->xml->startElement('a:prstGeom');
		$this->xml->writeAttribute('prst', 'rect');

		$this->xml->writeElement('a:avLst');

		$this->xml->endElement();

		$this->xml->endElement();
		$this->xml->endElement();

		$this->xml->endElement();
	}

	/**
	 * @param string $name - имя картинки
	 */
	private function writeNvPicPr(string $name){
		$this->xml->startElement('xdr:nvPicPr');

		$this->xml->startElement('xdr:cNvPr');
		$this->xml->writeAttribute('id', $this->nextImageId++);
		$this->xml->writeAttribute('name', $name);
		$this->xml->endElement();

		$this->xml->startElement('xdr:cNvPicPr');
		$this->xml->startElement('a:picLocks');

		$this->xml->writeAttribute('noChangeAspect', 1);

		$this->xml->endElement();
		$this->xml->endElement();

		$this->xml->endElement();
	}

	/**
	 * @param string $rId - ид связи
	 */
	private function writeBlipFill(string $rId){
		$this->xml->startElement('xdr:blipFill');

		$this->xml->startElement('a:blip');
		$this->xml->writeAttribute('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');
		$this->xml->writeAttribute('r:embed', $rId);

		$this->xml->startElement('a:extLst');

		$this->xml->startElement('a:ext');
		$this->xml->writeAttribute('uri', '{28A0092B-C50C-407E-A947-70E740481C1C}');

		$this->xml->startElement('a14:useLocalDpi');
		$this->xml->writeAttribute('xmlns:a14', 'http://schemas.microsoft.com/office/drawing/2010/main');
		$this->xml->writeAttribute('val', 0);
		$this->xml->endElement();

		$this->xml->endElement();
		$this->xml->endElement();

		$this->xml->endElement();

		$this->xml->startElement('a:stretch');
		$this->xml->writeElement('a:fillRect');
		$this->xml->endElement();

		$this->xml->endElement();
	}

	/**
	 * Начать файлы рисунков и связей рисунков
	 */
	private function startDrawing(){
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
	private function endDrawing(){
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
	private function checkCommited(){
		if ($this->committed) throw new ObjectCommittedException();
	}
}