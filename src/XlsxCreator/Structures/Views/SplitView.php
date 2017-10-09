<?php

namespace XlsxCreator\Structures\Views;

use XlsxCreator\Structures\Views\View;
use XlsxCreator\Validator;

/**
 * Class SplitView. Представление разделено на 4 секции с независимой прокруткой.
 *
 * @package XlsxCreator\Structures\Views
 */
class SplitView extends View{
	public function __construct(){
		$this->model['state'] = 'split';
	}

	/**
	 * @return int|null - Количество точек слева до границы
	 */
	function getXSplit(){
		return $this->model['xSplit'] ?? null;
	}

	/**
	 * @param int|null $xSplit - Количество точек слева до границы
	 * @return SplitView - $this
	 */
	function setXSplit(int $xSplit = null) : self{
		if (!is_null($xSplit)) Validator::validatePositive($xSplit, '$xSplit');

		$this->model['xSplit'] = $xSplit;
		return $this;
	}

	/**
	 * @return int|null - Количество точек сверху до границы
	 */
	function getYSplit(){
		return $this->model['ySplit'] ?? null;
	}

	/**
	 * @param int|null $ySplit - Количество точек сверху до границы
	 * @return SplitView - $this
	 */
	function setYSplit(int $ySplit = null) : self{
		if (!is_null($ySplit)) Validator::validatePositive($ySplit, '$ySplit');

		$this->model['ySplit'] = $ySplit;
		return $this;
	}

	/**
	 * @return string|null - Левая-верхняя ячейка в нижней правой панели
	 */
	function getTopLeftCell(){
		return $this->model['topLeftCell'] ?? null;
	}

	/**
	 * @param string|null $topLeftCell - Левая-верхняя ячейка в нижней правой панели (например, 'A1', 'B10', и т.д.)
	 * @return SplitView - $this
	 */
	function setTopLeftCell(string $topLeftCell = null) : self{
		if (!is_null($topLeftCell)) Validator::validateAddress($topLeftCell);

		$this->model['topLeftCell'] = $topLeftCell;
		return $this;
	}

	/**
	 * @return string|null - Левая-верхняя ячейка в нижней правой панели
	 */
	function getActivePane(){
		return $this->model['activePane'] ?? null;
	}

	/**
	 * @param string|null $activePane - Левая-верхняя ячейка в нижней правой панели (например, 'A1', 'B10', и т.д.)
	 * @return SplitView - $this
	 */
	function setActivePane(string $activePane = null) : self{
		if (!is_null($activePane)) Validator::validateAddress($activePane);

		$this->model['activePane'] = $activePane;
		return $this;
	}
}