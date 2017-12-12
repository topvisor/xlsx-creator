<?php

namespace Topvisor\XlsxCreator;

use OutOfBoundsException;
use Topvisor\XlsxCreator\Exceptions\InvalidValueException;
use Topvisor\XlsxCreator\Exceptions\ObjectCommittedException;
use Topvisor\XlsxCreator\Helpers\Comments;
use Topvisor\XlsxCreator\Helpers\Drawing;
use Topvisor\XlsxCreator\Helpers\SheetRels;
use Topvisor\XlsxCreator\Structures\Color;
use Topvisor\XlsxCreator\Structures\PageSetup;
use Topvisor\XlsxCreator\Structures\Range\CellsRange;
use Topvisor\XlsxCreator\Structures\Range\Range;
use Topvisor\XlsxCreator\Structures\Views\NormalView;
use Topvisor\XlsxCreator\Structures\Views\View;
use Topvisor\XlsxCreator\Exceptions\EmptyObjectException;
use Topvisor\XlsxCreator\Xml\ListXml;
//use Topvisor\XlsxCreator\Xml\Sheet\AutoFilterXml;
use Topvisor\XlsxCreator\Xml\Sheet\ColumnXml;
use Topvisor\XlsxCreator\Xml\Sheet\MergeXml;
use Topvisor\XlsxCreator\Xml\Sheet\PageMargins;
use Topvisor\XlsxCreator\Xml\Sheet\PageSetupXml;
use Topvisor\XlsxCreator\Xml\Sheet\RowXml;
use Topvisor\XlsxCreator\Xml\Sheet\SheetFormatPropertiesXml;
use Topvisor\XlsxCreator\Xml\Sheet\SheetPropertiesXml;
use Topvisor\XlsxCreator\Xml\Sheet\SheetViewsXml;
use Topvisor\XlsxCreator\Helpers\Validator;
use Topvisor\XlsxCreator\Xml\Simple\StringXml;
use Topvisor\XlsxCreator\Helpers\Styles;
use XMLWriter;

/**
 * Class Worksheet. Содержит методы для работы с таблицей.
 *
 * @package  Topvisor\XlsxCreator
 */
class Worksheet{
	const DY_DESCENT = 55;
	const DEFAULT_ROW_HEIGHT = 15;

	private $workbook;
	private $styles;
	private $id;
	private $name;
	private $tabColor;
	private $outlineLevelCol;
	private $outlineLevelRow;
	private $view;
	private $pageSetup;
//	private $autoFilter;

	private $committed;
	private $lastUncommittedRow;
	private $rows;
	private $columns;
	private $merges;
	private $comments;
	private $drawing;
	private $sheetRels;

	private $filename;
	private $xml;

	private $rId;

	/**
	 * Worksheet constructor.
	 *
	 * @param Workbook $workbook - workbook, к которому принадлежит таблица
	 * @param Styles $styles - стили xlsx
	 * @param int $id - ID таблицы в $workbook
	 * @param string $name - имя таблицы
	 */
	function __construct(Workbook $workbook, Styles $styles, int $id, string $name){
		$this->workbook = $workbook;
		$this->styles = $styles;
		$this->id = $id;
		$this->name = $name;

		$this->tabColor = null;
		$this->outlineLevelCol = 0;
		$this->outlineLevelRow = 0;
		$this->view = new NormalView();
		$this->pageSetup = new PageSetup();
//		$this->autoFilter = null;

		$this->committed = false;
		$this->columns = [];
		$this->rows = [];
		$this->merges = [];
		$this->lastUncommittedRow = 1;
		$this->comments = new Comments($this);
		$this->drawing = new Drawing($this);
		$this->sheetRels = new SheetRels($this);
	}

	function __destruct(){
		unset($this->workbook);
		unset($this->styles);
		unset($this->view);
		unset($this->pageSetup);
		unset($this->rows);
		unset($this->merges);
		unset($this->comments);
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
	 * @param Color|null $tabColor - цвет вкладки
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

//	/**
//	 * @return string|null - автоматический фильтр
//	 */
//	function getAutoFilter(){
//		return $this->autoFilter;
//	}
//
//	/**
//	 * @param string|null $autoFilter - автоматический фильтр ('A1:A5')
//	 * @return Worksheet - $this
//	 * @throws ObjectCommittedException
//	 * @throws InvalidValueException
//	 */
//	function setAutoFilter(string $autoFilter = null) : Worksheet{
//		$this->checkCommitted();
//		$this->checkStarted();
//
//		if (!is_null($autoFilter)) Validator::validateCellsRange($autoFilter);
//
//		$this->autoFilter = $autoFilter;
//		return $this;
//	}

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
		if ($this->xml ?? false) $this->xml->flush();
		return $this->filename;
	}

	/**
	 * @return string - путь к файлу таблицы внутри xlsx
	 */
	function getLocalname() : string{
		return 'xl/worksheets/sheet' . $this->id . '.xml';
	}

	/**
	 * Внутреняя функция. Возвращает id связи.
	 *
	 * @return string - id связи файла таблицы
	 */
	function getRId() : string{
		return $this->rId ?? '';
	}

	/**
	 * Внутреняя функция. Назначает id связи.
	 *
	 * @param string $rId - id связи файла таблицы
	 * @return Worksheet - $this
	 */
	function setRId(string $rId) : Worksheet{
		$this->rId = $rId;
		return $this;
	}

	/**
	 * @return null|string - временый файл связей таблицы
	 */
	function getSheetRelsFilename(){
		return $this->sheetRels->getFilename();
	}

	/**
	 * @return array|null - временные файлы комментариев таблицы
	 */
	function getCommentsFilenames(){
		return !$this->comments->isEmpty() ? [
			'comments' => $this->comments->getCommentsFilename(),
			'vml' => $this->comments->getVmlFilename()
		] : null;
	}

	/**
	 * @return array|null - временные файлы рисунков таблицы
	 */
	function getDrawingFilenames(){
		return !$this->drawing->isEmpty() ? [
			'drawing' => $this->drawing->getFilename(),
			'rels' => $this->drawing->getRelsFilename()
		] : null;
	}

	/**
	 * @param int $col - номер колонки
	 * @return Column - колонка
	 * @throws InvalidValueException
	 * @throws ObjectCommittedException
	 */
	function getColumn(int $col) : Column{
		Validator::validateInRange($col, 1, 16384, '$col');

		$this->checkCommitted();
		$this->checkStarted();

		if ($col > count($this->columns))
			for ($i = count($this->columns) + 1; $i <= $col; $i++)
				$this->columns[] = new Column($this, $i);

		return $this->columns[$col - 1];
	}

	/**
	 * @return Column - колонка
	 * @throws ObjectCommittedException
	 */
	function addColumn() : Column{
		$this->checkCommitted();
		$this->checkStarted();
		if (count($this->columns) >= 16384) throw new OutOfBoundsException('Excel supports columns from 1 to 16384');

		$column = new Column($this, count($this->columns) + 1);
		$this->columns[] = $column;

		return $column;
	}

	/**
	 * @param int $row - номер строки
	 * @return Row - строка
	 * @throws ObjectCommittedException
	 * @throws InvalidValueException
	 */
	function getRow(int $row) : Row{
		$this->checkCommitted();

		Validator::validateInRange($row, 1, 1048576,'$row');
		if ($row < $this->lastUncommittedRow) throw new ObjectCommittedException('Row is committed');

		if ($row >= $this->lastUncommittedRow + count($this->rows))
			for ($i = $this->lastUncommittedRow + count($this->rows); $i <= $row; $i++)
				$this->rows[] = new Row($this, $i);

		return $this->rows[$row - $this->lastUncommittedRow];
	}

	/**
	 * @param array|null $values - значения ячеек строки
	 * @return Row - строка таблицы
	 * @throws ObjectCommittedException
	 * @throws InvalidValueException
	 */
	function addRow(array $values = null) : Row{
		$this->checkCommitted();
		if ($this->lastUncommittedRow + count($this->rows) > 1048576) throw new OutOfBoundsException('Excel supports rows from 1 to 1048576');

		$row = new Row($this, count($this->rows) + $this->lastUncommittedRow);
		if (!is_null($values)) $row->setCells($values);

		$this->rows[] = $row;

		return $row;
	}

	/**
	 * @param int $row - номер строки
	 * @param int $col - номер колонки
	 * @return Cell - ячейка
	 * @throws ObjectCommittedException
	 * @throws InvalidValueException
	 */
	function getCell(int $row, int $col) : Cell{
		$this->checkCommitted();

		return $this->getRow($row)->getCell($col);
	}

	/**
	 * Вставить картинку на таблицу
	 *
	 * @param Range $position - месторасположения картинки
	 * @param string $filename - путь к файлу картинки
	 * @param string $extension - расшинение файла картинки
	 * @param string $imageName - имя картинки
	 * @throws InvalidValueException
	 */
	function addImage(Range $position, string $filename, string $extension = '', string $imageName = ''){
		$this->drawing->addImage($this->workbook->addImage($filename, $extension), $position, $imageName);
	}

	/**
	 * Соединить ячейки в одну
	 *
	 * @param CellsRange $range - диапазон ячеек
	 * @throws InvalidValueException
	 * @throws ObjectCommittedException
	 */
	function mergeCells(CellsRange $range) {
		$this->checkCommitted();

		$master = $this->getCell($range->getTop(), $range->getLeft());
		$this->getCell($range->getBottom(), $range->getRight());

		foreach ($this->merges as $merge)
			if ($range->intersection($merge))
				throw new InvalidValueException('Merge intersect');

		for ($i = $range->getTop(); $i <= $range->getBottom(); $i++) {
			$row = $this->getRow($i);
			for ($j = $range->getLeft(); $j <= $range->getRight(); $j++) {
				$cell = $row->getCell($j);
				if ($cell === $master) continue;
				else $cell->setMaster($master);
			}
		}

		$this->merges[] = $range;
	}

	/**
	 * Разъединить ячейки
	 *
	 * @param CellsRange $range - диапазон ячеек
	 * @throws InvalidValueException
	 * @throws ObjectCommittedException
	 */
	function unMergeCells(CellsRange $range) {
		$this->checkCommitted();

		$this->getCell($range->getTop(), $range->getLeft());
		$this->getCell($range->getBottom(), $range->getRight());

		$found = false;
		foreach ($this->merges as $mergeIndex => $merge) {
			if ($merge == $range) {
				$found = $mergeIndex;
				break;
			}
		}

		if ($found === false) throw new InvalidValueException('Merge not found');

		for ($i = $range->getTop(); $i <= $range->getBottom(); $i++) {
			$row = $this->getRow($i);
			for ($j = $range->getLeft(); $j <= $range->getRight(); $j++)
				$row->getCell($j)->setMaster(null);
		}

		unset($this->merges[$found]);
	}

	/**
	 * Зафиксировать файл таблицы.
	 *
	 * @throws ObjectCommittedException
	 * @throws EmptyObjectException
	 */
	function commit(){
		if (!$this->rows && $this->lastUncommittedRow == 1) throw new EmptyObjectException('Worksheet is empty');

		$this->commitRows();
		unset($this->rows);

		$this->committed = true;

		$this->endWorksheet();
	}

	/**
	 * Зафиксировать строки таблицы.
	 *
	 * @param int|null $lastRow - последняя фиксируемая строка
	 * @throws ObjectCommittedException
	 */
	function commitRows(int $lastRow = null){
		$this->checkCommitted();
		if (!$this->rows) return;

		if (!($this->xml ?? false)) $this->startWorksheet();

		$rowXml = new RowXml();
		$found = false;
		while (count($this->rows) && !$found) {
			$row = array_shift($this->rows);
			$found = (bool) ($row->getNumber() == $lastRow);

			$rowXml->render($this->xml, $row->prepareToCommit($this->styles, $this->sheetRels, $this->comments));
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
			'partName' => $this->getLocalname(),
			'useComments' => !$this->comments->isEmpty(),
			'useDrawing' => !$this->drawing->isEmpty()
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
			'defaultRowHeight' => self::DEFAULT_ROW_HEIGHT,
			'outlineLevelCol' => $this->outlineLevelCol,
			'outlineLevelRow' => $this->outlineLevelRow,
			'dyDescent' => self::DY_DESCENT
		]);

		(new ListXml('cols', new ColumnXml()))->render($this->xml, array_map(function($column){
			return $column->prepareToCommit($this->styles);
		}, $this->columns));

		$this->xml->startElement('sheetData');
	}

	/**
	 * Закончить файл таблицы.
	 *
	 * @throws ObjectCommittedException
	 */
	private function endWorksheet(){
		if (!($this->xml ?? false)) return;

		$this->xml->endElement();

//		(new AutoFilterXml())->render($this->xml, [$this->autoFilter]);
		(new ListXml('mergeCells', new MergeXml(), [], false, true))->render($this->xml, array_map(function($merge){
			return [(string) $merge];
		}, $this->merges));

		if ($hyperlinksFilename = $this->sheetRels->getHyperlinksFilename()) {
			$this->xml->startElement('hyperlinks');

			$hyperlinksFile = fopen($hyperlinksFilename, 'r');
			while (!feof($hyperlinksFile)) $this->xml->writeRaw(fread($hyperlinksFile, 8192));
			fclose($hyperlinksFile);

			$this->xml->endElement();
		}

		(new PageMargins())->render($this->xml, $this->pageSetup->getModel()['margins'] ?? null);
		(new PageSetupXml())->render($this->xml, $this->pageSetup->getModel());

		if (!$this->drawing->isEmpty())
			(new StringXml('drawing', [], 'r:id'))
				->render($this->xml, [$this->sheetRels->addDrawing()]);

		if (!$this->comments->isEmpty())
			(new StringXml('legacyDrawing', [], 'r:id'))
				->render($this->xml, [$this->sheetRels->addComments()]);

		$this->xml->endElement();
		$this->xml->endDocument();

		$this->xml->flush();
		unset($this->xml);

		$this->drawing->commit();
		$this->comments->commit();
		$this->sheetRels->commit();
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