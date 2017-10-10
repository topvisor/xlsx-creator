<?php

namespace Topvisor\XlsxCreator;

use OutOfBoundsException;
use Topvisor\XlsxCreator\Exceptions\InvalidValueException;
use Topvisor\XlsxCreator\Exceptions\ObjectCommittedException;
use Topvisor\XlsxCreator\Structures\Color;
use Topvisor\XlsxCreator\Structures\PageSetup;
use Topvisor\XlsxCreator\Structures\Views\NormalView;
use Topvisor\XlsxCreator\Structures\Views\View;
use Topvisor\XlsxCreator\Exceptions\EmptyObjectException;
use Topvisor\XlsxCreator\Xml\ListXml;
use Topvisor\XlsxCreator\Xml\Sheet\HyperlinkXml;
use Topvisor\XlsxCreator\Xml\Sheet\PageMargins;
use Topvisor\XlsxCreator\Xml\Sheet\PageSetupXml;
use Topvisor\XlsxCreator\Xml\Sheet\RowXml;
use Topvisor\XlsxCreator\Xml\Sheet\SheetFormatPropertiesXml;
use Topvisor\XlsxCreator\Xml\Sheet\SheetPropertiesXml;
use Topvisor\XlsxCreator\Xml\Sheet\SheetViewsXml;
use XMLWriter;

/**
 * Class Worksheet. Содержит методы для работы с таблицей.
 *
 * @package XlsxCreator
 */
class Worksheet{
	const DY_DESCENT = 55;

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

	/**
	 * Worksheet constructor.
	 *
	 * @param Workbook $workbook - workbook, к которому принадлежит таблица
	 * @param int $id - ID таблицы в $workbook
	 * @param string $name - имя таблицы
	 */
	function __construct(Workbook $workbook, int $id, string $name){
		$this->workbook = $workbook;
		$this->id = $id;
		$this->name = $name;

		$this->tabColor = null;
		$this->outlineLevelCol = 0;
		$this->outlineLevelRow = 0;
		$this->defaultRowHeight = 15;
		$this->view = new NormalView();
		$this->pageSetup = new PageSetup();
		$this->autoFilter = null;

		$this->committed = false;
		$this->columns = [];
		$this->rows = [];
		$this->merges = [];
		$this->lastUncommittedRow = 1;
		$this->sheetRels = new SheetRels($this);
	}

	function __destruct(){
		unset($this->workbook);
		unset($this->view);
		unset($this->pageSetup);
		unset($this->rows);
		unset($this->merges);
		unset($this->sheetRels);
		unset($this->xml);

		if (file_exists($this->filename)) unlink($this->filename);
	}

	/**
	 * @return Workbook - workbook, к которому принадлежит таблица
	 */
	function getWorkbook() : Workbook{
		return $this->workbook;
	}

	/**
	 * @return int - ID таблицы в $workbook
	 */
	function getId() : int{
		return $this->id;
	}

	/**
	 * @return string - имя таблицы
	 */
	function getName() : string{
		return $this->name;
	}

	/**
	 * @return Color|null - цвет вкладки
	 */
	function getTabColor(){
		return $this->tabColor;
	}

	/**
	 * @param Color|null $tabColor - цвет вкладки в формате 'FF00FF00'
	 * @return Worksheet - $this
	 * @throws ObjectCommittedException
	 */
	function setTabColor(Color $tabColor = null) : Worksheet{
		$this->checkCommitted();
		$this->checkStarted();

		$this->tabColor = $tabColor;
		return $this;
	}

	/**
	 * @return int - worksheet column outline level
	 */
	function getOutlineLevelCol() : int{
		return $this->outlineLevelCol;
	}

	/**
	 * @param int $outlineLevelCol - worksheet column outline level
	 * @return Worksheet - $this
	 * @throws ObjectCommittedException
	 * @throws InvalidValueException
	 */
	function setOutlineLevelCol(int $outlineLevelCol) : Worksheet{
		$this->checkCommitted();
		$this->checkStarted();

		Validator::validatePositive($outlineLevelCol, '$outlineLevelCol');

		$this->outlineLevelCol = $outlineLevelCol;
		return $this;
	}

	/**
	 * @return int - worksheet row outline level
	 */
	function getOutlineLevelRow() : int{
		return $this->outlineLevelRow;
	}

	/**
	 * @param int $outlineLevelRow - worksheet row outline level
	 * @return Worksheet - $this
	 * @throws ObjectCommittedException
	 * @throws InvalidValueException
	 */
	function setOutlineLevelRow(int $outlineLevelRow) : Worksheet{
		$this->checkCommitted();
		$this->checkStarted();

		Validator::validatePositive($outlineLevelRow, '$outlineLevelRow');

		$this->outlineLevelRow = $outlineLevelRow;
		return $this;
	}

	/**
	 * @return int - высота строк по умолчанию
	 */
	function getDefaultRowHeight() : int{
		return $this->defaultRowHeight;
	}

	/**
	 * @param int $defaultRowHeight - высота строк по умолчанию
	 * @return Worksheet - $this
	 * @throws ObjectCommittedException
	 * @throws InvalidValueException
	 */
	function setDefaultRowHeight(int $defaultRowHeight) : Worksheet{
		$this->checkCommitted();
		$this->checkStarted();

		Validator::validatePositive($defaultRowHeight, '$defaultRowHeight');

		$this->defaultRowHeight = $defaultRowHeight;
		return $this;
	}

	/**
	 * @return View - представление worksheet
	 */
	function getView() : View{
		return $this->view;
	}

	/**
	 * @param View $view - представление worksheet
	 * @return Worksheet - $this
	 * @throws ObjectCommittedException
	 * @throws InvalidValueException
	 */
	function setView(View $view) : Worksheet{
		$this->checkCommitted();
		$this->checkStarted();

		if (!is_null($view)) Validator::validateView($view);

		$this->view = $view;
		return $this;
	}

	/**
	 * @return PageSetup - параметры печати
	 */
	function getPageSetup() : PageSetup{
		return $this->pageSetup;
	}

	/**
	 * @param PageSetup $pageSetup - параметры печати
	 * @return Worksheet - $this
	 * @throws ObjectCommittedException
	 */
	function setPageSetup(PageSetup $pageSetup) : Worksheet{
		$this->checkCommitted();
		$this->checkStarted();

		$this->pageSetup = $pageSetup;
		return $this;
	}

	/**
	 * @return string|null - автоматический фильтр
	 */
	function getAutoFilter(){
		return $this->autoFilter;
	}

	/**
	 * @param string|null $autoFilter - автоматический фильтр ('A1:A5')
	 * @return Worksheet - $this
	 * @throws ObjectCommittedException
	 * @throws InvalidValueException
	 */
	function setAutoFilter(string $autoFilter = null) : Worksheet{
		$this->checkCommitted();
		$this->checkStarted();

		if (!is_null($autoFilter)) Validator::validateCellsRange($autoFilter);

		$this->autoFilter = $autoFilter;
		return $this;
	}

	/**
	 * @return bool - зафиксированы ли изменения
	 */
	function isCommitted() : bool{
		return $this->committed;
	}

	/**
	 * @return null|string - путь к временному файлу таблицы
	 */
	function getFilename(){
		return $this->filename;
	}

	/**
	 * @return string - путь к файлу таблицы внутри xlsx
	 */
	function getLocalname() : string{
		return 'xl/worksheets/sheet' . $this->id . '.xml';
	}

	/**
	 * @return string - id связи файла таблицы
	 */
	function getRId() : string{
		return $this->rId ?? '';
	}

	/**
	 * @param string $rId - id связи файла таблицы
	 * @return Worksheet - $this
	 */
	function setRId(string $rId) : Worksheet{
		$this->rId = $rId;
		return $this;
	}

	/**
	 * @return SheetRels - гиперссылки внутри таблицы
	 */
	function getSheetRels() : SheetRels{
		return $this->sheetRels;
	}

	/**
	 * @param int $row - номер строки
	 * @return Row - строка
	 * @throws ObjectCommittedException
	 * @throws InvalidValueException
	 */
	function getRow(int $row) : Row{
		Validator::validateInRange($row, 1, 1048576,'$row');
		if ($row < $this->lastUncommittedRow) throw new ObjectCommittedException('Row is committed');
		if ($row >= $this->lastUncommittedRow + count($this->rows)) throw new OutOfBoundsException();

		return $this->rows[$row - $this->lastUncommittedRow];
	}

	/**
	 * @param array|null $values - значения ячеек строки
	 * @return Row - строка таблицы
	 */
	function addRow(array $values = null) : Row{
		$this->checkCommitted();

		if (!$this->xml) $this->startWorksheet();

		$row = new Row($this, count($this->rows) + $this->lastUncommittedRow);
		if (!is_null($values)) $row->setCells($values);

		$this->rows[] = $row;

		return $row;
	}

	/**
	 *	Зафиксировать файл таблицы.
	 *
	 * 	@throws ObjectCommittedException
	 */
	function commit(){
		if (!$this->xml) throw new EmptyObjectException('Worksheet is empty');

		$this->commitRows();
		unset($this->rows);

		$this->committed = true;

		$this->endWorksheet();
	}

	/**
	 * Зафиксировать строки таблицы.
	 *
	 * @param Row|null $lastRow - последняя фиксируемая строка
	 * @throws ObjectCommittedException
	 */
	function commitRows(Row $lastRow = null){
		$this->checkCommitted();
		if (!$this->rows) return;

		$lastRow = $lastRow ?? $this->rows[count($this->rows) - 1];
		$lastRowNumber = $lastRow->getNumber();
		$rowXml = new RowXml();

		$found = false;
		while (count($this->rows) && !$found) {
			$row = array_shift($this->rows);
			$found = (bool) ($row->getNumber() == $lastRowNumber);

			$rowXml->render($this->xml, $row->getModel());
			$this->lastUncommittedRow++;
		}
	}

	/**
	 * @return array - модель таблицы
	 */
	function getModel() : array{
		return [
			'id' => $this->id,
			'name' => $this->name,
			'rId' => $this->rId ?? '',
			'partName' => $this->getLocalname()
		];
	}

	/**
	 *	Начать файл таблицы.
	 */
	private function startWorksheet(){
		$this->filename = $this->workbook->genTempFilename();
		$this->xml = new XMLWriter();
		$this->xml->openURI($this->filename);

		$this->xml->startDocument('1.0', 'UTF-8', 'yes');
		$this->xml->startElement('worksheet');

		$this->xml->writeAttribute('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');
		$this->xml->writeAttribute('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');
		$this->xml->writeAttribute('xmlns:mc', 'http://schemas.openxmlformats.org/markup-compatibility/2006');
		$this->xml->writeAttribute('mc:Ignorable', 'x14ac');
		$this->xml->writeAttribute('xmlns:x14ac', 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac');

		(new SheetPropertiesXml())->render($this->xml, [
			'tabColor' => $this->tabColor ? $this->tabColor->getModel() : null,
			'pageSetup' => $this->pageSetup->getModel()
		]);

		(new SheetViewsXml())->render($this->xml, $this->view->getModel());

		(new SheetFormatPropertiesXml())->render($this->xml, [
			'defaultRowHeight' => $this->defaultRowHeight,
			'outlineLevelCol' => $this->outlineLevelCol,
			'outlineLevelRow' => $this->outlineLevelRow,
			'dyDescent' => Worksheet::DY_DESCENT
		]);

		$this->writeColumns();

		$this->xml->startElement('sheetData');
	}

	/**
	 *	Закончить файл таблицы.
	 */
	private function endWorksheet(){
		$this->xml->endElement();

		// AutoFilter
		// MergeCells

		(new ListXml('hyperlinks', new HyperlinkXml()))->render($this->xml, $this->sheetRels->getHyperlinks());
		(new PageMargins())->render($this->xml, $this->pageSetup->getModel()['margins'] ?? null);
		(new PageSetupXml())->render($this->xml, $this->pageSetup->getModel());

		$this->xml->endElement();
		$this->xml->endDocument();

		$this->xml->flush();
		unset($this->xml);

		$this->sheetRels->commit();
	}

	/**
	 *	Записать описание колонок в файл
	 */
	private function writeColumns(){
		// Реализовать
	}

	/**
	 * @throws ObjectCommittedException
	 */
	private function checkCommitted(){
		if ($this->committed) throw new ObjectCommittedException('Worksheet is committed');
	}

	/**
	 * @throws ObjectCommittedException
	 */
	private function checkStarted(){
		if ($this->rows) throw new ObjectCommittedException('Worksheet properties is committed');
	}
}