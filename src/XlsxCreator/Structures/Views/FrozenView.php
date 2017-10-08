<?php

namespace XlsxCreator\Structures\Views;

use XlsxCreator\Validator;

class FrozenView extends View{
	public function __construct(){
		$this->model['state'] = 'frozen';
	}

	/**
	 * @return int|null - Сколько столбцов "заморожено"
	 */
	function getXSplit(){
		return $this->model['xSplit'] ?? null;
	}

	/**
	 * @param int|null $xSplit - Сколько столбцов "заморожено"
	 * @return FrozenView - $this
	 */
	function setXSplit(int $xSplit = null) : self{
		if (!is_null($xSplit)) Validator::validateInRange($xSplit, 1, 16384, '$xSplit');

		$this->model['xSplit'] = $xSplit;
		return $this;
	}

	/**
	 * @return int|null - Сколько строк "заморожено"
	 */
	function getYSplit(){
		return $this->model['ySplit'] ?? null;
	}

	/**
	 * @param int|null $ySplit - Сколько строк "заморожено"
	 * @return FrozenView - $this
	 */
	function setYSplit(int $ySplit = null) : self{
		if (!is_null($ySplit)) Validator::validateInRange($ySplit, 1, 1048576, '$ySplit');

		$this->model['ySplit'] = $ySplit;
		return $this;
	}

	/**
	 * @return string|null - Левая-верхняя ячейка в "незамороженной" панели
	 */
	function getTopLeftCell(){
		return $this->model['topLeftCell'] ?? null;
	}

	/**
	 * @param string|null $topLeftCell - Левая-верхняя ячейка в "незамороженной" панели (например, 'D4', 'G15', и т.д.)
	 * @return FrozenView - $this
	 */
	function setTopLeftCell(string $topLeftCell = null) : self{
		if (!is_null($topLeftCell)) Validator::validateAddress($topLeftCell);

		$this->model['topLeftCell'] = $topLeftCell;
		return $this;
	}
}