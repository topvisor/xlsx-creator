<?php

namespace Decaseal\XlsxCreator;

use Decaseal\XlsxCreator\Xml\ListXml;
use Decaseal\XlsxCreator\Xml\Sheet\SheetPropertiesXml;
use Decaseal\XlsxCreator\Xml\Sheet\SheetViewXml;
use XMLWriter;

class Worksheet{
	private $workbook;
	private $id;
	private $name;
//	private $tabColor;
//	private $defaultRowHeight;
//	private $view;

	private $committed;
	private $rows;
	private $merges;
	private $sheetRels;
	private $startedData;

	private $filename;
	private $xml;

	private $rId;

	function __construct(XlsxCreator $workbook, int $id, string $name){
		$this->workbook = $workbook;
		$this->id = $id;
		$this->name = $name;
//		$this->tabColor = $tabColor;
//		$this->defaultRowHeight = $defaultRowHeight;
//		$this->view = $view;
//
		$this->committed = false;
		$this->rows = [];
		$this->merges = [];
		$this->sheetRels = new SheetRels($this);
		$this->startedData = false;

		$this->filename = $this->workbook->genTempFilename();
		$this->xml = new XMLWriter();
		$this->xml->openURI($this->filename);

		$this->startWorksheet();
	}

	function getId() : int{
		return $this->id;
	}

	function getName() : string{
		return $this->name;
	}

	function isCommitted() : bool{
		return $this->committed;
	}

	function setRId(string $rId){
		$this->rId = $rId;
	}

	function getRId() : string{
		return $this->rId ?? '';
	}

	function getWorkbook() : XlsxCreator{
		return $this->workbook;
	}

//	function getAbsFilename() : string{
//		return $this->workbook->getTempdir() . $this->getRelFilename();
//	}

//	function getRelFilename() : string{
//		return '/xl/worksheets/sheet' . $this->getId() . '.xml';
//	}

//	function addRow($values = null) : Row{
//		$row = new Row($this, count($this->rows) + 1);
//		if (!is_null($values)) $row->setValues($values);
//
//		$this->rows[] = $row;
//
//		return $row;
//	}

	function commit(){
		if ($this->isCommitted()) return;
		$this->committed = true;
	}

	function getModel() : array{
		return [
			'id' => $this->getId(),
			'name' => $this->getName(),
			'rId' => $this->getRId(),
			'partName' => '/xl/worksheets/sheet' . $this->getId() . '.xml'
		];
	}

	private function startWorksheet(){
		$this->xml->startDocument('1.0', 'UTF-8', 'yes');

		$this->xml->startElement('worksheet');
		$this->xml->writeAttribute('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');
		$this->xml->writeAttribute('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');
		$this->xml->writeAttribute('xmlns:mc', 'http://schemas.openxmlformats.org/markup-compatibility/2006');
		$this->xml->writeAttribute('mc:Ignorable', 'x14ac');
		$this->xml->writeAttribute('xmlns:x14ac', 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac');

		(new SheetPropertiesXml())->render($this->xml, [

		]);
		(new ListXml('sheetViews', new SheetViewXml()))->render($this->xml, [

		]);
	}

//	private function endWorksheet(){
//
//	}

//	private function writeRows(){
//		if (!$this->startedData) {
//			$this->writeColumns();
//			$this->writeOpenSheetData();
//			$this->startedData = true;
//		}
//
//		foreach ($this->rows as $row) {
//			if ($row->hasValues()) {
//
//			}
//		}
//	}

//	private function writeColumns(){

//	}

//	private function writeOpenSheetData(){

//	}

}