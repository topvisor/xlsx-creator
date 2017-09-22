<?php

namespace Decaseal\XlsxCreator;

use Decaseal\XlsxCreator\Xml\ListXml;
use Decaseal\XlsxCreator\Xml\Sheet\SheetFormatPropertiesXml;
use Decaseal\XlsxCreator\Xml\Sheet\SheetPropertiesXml;
use Decaseal\XlsxCreator\Xml\Sheet\SheetViewXml;
use XMLWriter;

class Worksheet{
	private const DY_DESCENT = 55;

	private $workbook;
	private $id;
	private $name;
	private $tabColor;
	private $outlineLevelCol;
	private $outlineLevelRow;
	private $defaultRowHeight;
	private $view;
	private $autoFilter;
	private $columns;

	private $committed;
	private $lastUncommittedRow;
	private $rows;
	private $merges;
	private $sheetRels;

	private $filename;
	private $xml;

	private $rId;

	function __construct(Workbook $workbook, int $id, string $name, string $tabColor = null, int $outlineLevelCol = 0, int $outlineLevelRow = 0,
						 int $defaultRowHeight = 15, array $view = null, array $autoFilter = null){
		$this->workbook = $workbook;
		$this->id = $id;
		$this->name = $name;

		$this->tabColor = $tabColor;
		$this->outlineLevelCol = $outlineLevelCol;
		$this->outlineLevelRow = $outlineLevelRow;
		$this->defaultRowHeight = $defaultRowHeight;
		$this->view = $view;
		$this->autoFilter = $autoFilter;
		$this->columns = [];

		$this->committed = false;
		$this->lastUncommittedRow = 1;
		$this->rows = [];
		$this->merges = [];
		$this->sheetRels = new SheetRels($this);

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

	function getOutlineLevelRow() : int{
		return $this->outlineLevelRow;
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

	function getWorkbook() : Workbook{
		return $this->workbook;
	}

	function getSheetRels() : SheetRels{
		return $this->sheetRels;
	}

	function addRow(array $values = null) : Row{
		$row = new Row($this, count($this->rows) + $this->lastUncommittedRow);
		if (!is_null($values)) $row->setValues($values);

		$this->rows[] = $row;

		return $row;
	}

	function commit(){
		if ($this->isCommitted()) return;
		$this->committed = true;

		foreach ($this->rows as $row) $row->commit();
	}

	function getModel() : array{
		return [
			'id' => $this->id,
			'name' => $this->name,
			'rId' => $this->rId ?? '',
			'partName' => '/xl/worksheets/sheet' . $this->id . '.xml'
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

		if ($this->tabColor) (new SheetPropertiesXml())->render($this->xml, ['argb' => $this->tabColor]);
		(new SheetFormatPropertiesXml())->render($this->xml, [
			'defaultRowHeight' => $this->defaultRowHeight,
			'outlineLevelCol' => $this->outlineLevelCol,
			'outlineLevelRow' => $this->outlineLevelRow,
			'dyDescent' => Worksheet::DY_DESCENT
		]);

		$this->writeColumns();

		$this->xml->startElement('sheetData');
	}

	private function writeColumns(){
		// Реализовать
	}
}