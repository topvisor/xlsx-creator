<?php

namespace Decaseal\XlsxCreator;

use Decaseal\XlsxCreator\Xml\ListXml;
use Decaseal\XlsxCreator\Xml\Sheet\HyperlinkXml;
use Decaseal\XlsxCreator\Xml\Sheet\PageMargins;
use Decaseal\XlsxCreator\Xml\Sheet\PageSetupXml;
use Decaseal\XlsxCreator\Xml\Sheet\RowXml;
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
	private $pageSetup;
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
						 int $defaultRowHeight = 15, array $view = null, array $pageSetup = null, array $autoFilter = null){
		$this->workbook = $workbook;
		$this->id = $id;
		$this->name = $name;

		$this->tabColor = $tabColor;
		$this->outlineLevelCol = $outlineLevelCol;
		$this->outlineLevelRow = $outlineLevelRow;
		$this->defaultRowHeight = $defaultRowHeight;
		$this->view = $view;
		$this->pageSetup = array_merge([
			'margins' => ['left' => 0.7, 'right' => 0.7, 'top' => 0.75, 'bottom' => 0.75, 'header' => 0.3, 'footer' => 0.3 ],
			'orientation' => 'portrait',
			'horizontalDpi' => 4294967295,
			'verticalDpi' => 4294967295,
			'fitToPage' => (bool) ($pageSetup && (($pageSetup['fitToWidth'] || $pageSetup['fitToHeight']) && !$pageSetup['scale'])),
			'pageOrder' => 'downThenOver',
			'blackAndWhite' => false,
			'draft' => false,
			'cellComments' => 'None',
			'errors' => 'displayed',
			'scale' => 100,
			'fitToWidth' => 1,
			'fitToHeight' => 1,
			'paperSize' => null,
			'showRowColHeaders' => false,
			'showGridLines' => false,
			'horizontalCentered' => false,
			'verticalCentered' => false,
			'rowBreaks' => null,
			'colBreaks' => null
		], $pageSetup);
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

	public function __destruct(){
		if (file_exists($this->filename)) unlink($this->filename);
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

	function getFilename() : string{
		return $this->filename;
	}

	function getLocalname() :string{
		return '/xl/worksheets/sheet' . $this->id . '.xml';
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

		foreach ($this->rows as $row) (new RowXml())->render($this->xml, $row->genModel());
		unset($this->rows);

		$this->endWorksheet();

		$this->xml->flush();
		unset($this->xml);

		$this->sheetRels->commit();
	}

	function commitRow(Row $lastRow){
		if ($this->isCommitted()) return;

		$rowXml = new RowXml();

		$found = false;
		while (count($this->rows) && !$found) {
			$row = array_shift($this->rows);
			$found = (bool) ($row->getNumber() == $lastRow->getNumber());

			$rowXml->render($this->xml, $row->getModel());
			$this->lastUncommittedRow++;
		}
	}

	function getModel() : array{
		return [
			'id' => $this->id,
			'name' => $this->name,
			'rId' => $this->rId ?? '',
			'partName' => $this->getLocalname()
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
			'tabColor' => ['argb' => $this->tabColor],
			'pageSetup' => $this->pageSetup
		]);

		(new SheetFormatPropertiesXml())->render($this->xml, [
			'defaultRowHeight' => $this->defaultRowHeight,
			'outlineLevelCol' => $this->outlineLevelCol,
			'outlineLevelRow' => $this->outlineLevelRow,
			'dyDescent' => Worksheet::DY_DESCENT
		]);

		$this->writeColumns();

		$this->xml->startElement('sheetData');
	}

	private function endWorksheet(){
		$this->xml->endElement();

		// AutoFilter
		// MergeCells

		(new ListXml('hyperlinks', new HyperlinkXml()))->render($this->xml, $this->sheetRels->getHyperlinks());
		(new PageMargins())->render($this->xml, $this->pageSetup['margins'] ?? null);
		(new PageSetupXml())->render($this->xml, $this->pageSetup);

		$this->xml->endElement();
		$this->xml->endDocument();
	}

	private function writeColumns(){
		// Реализовать
	}
}