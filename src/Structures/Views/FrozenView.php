<?php

namespace Topvisor\XlsxCreator\Structures\Views;

use Topvisor\XlsxCreator\Validator;

/**
 * Class FrozenView. Несколько строк и/или столбцов этого представления заморожено на месте.
 *
 * @package XlsxCreator\Structures\Views
 */
class FrozenView extends View{
	public function __construct(int $xSplit, int $ySplit){
		$this->model['state'] = 'frozen';

		$this->model['xSplit'] = $xSplit;
		$this->model['ySplit'] = $ySplit;
	}

	/**
	 * @return int - Сколько столбцов "заморожено"
	 */
	function getXSplit(){
		return $this->model['xSplit'];
	}

	/**
	 * @param int $xSplit - Сколько столбцов "заморожено"
	 * @return FrozenView - $this
	 */
	function setXSplit(int $xSplit) : self{
		Validator::validateInRange($xSplit, 1, 16384, '$xSplit');

		$this->model['xSplit'] = $xSplit;
		return $this;
	}

	/**
	 * @return int - Сколько строк "заморожено"
	 */
	function getYSplit(){
		return $this->model['ySplit'];
	}

	/**
	 * @param int|null $ySplit - Сколько строк "заморожено"
	 * @return FrozenView - $this
	 */
	function setYSplit(int $ySplit) : self{
		Validator::validateInRange($ySplit, 1, 1048576, '$ySplit');

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