<?php

namespace Topvisor\XlsxCreator\Structures\Styles\Borders;

/**
 * Class Borders. Описывает состояния и стили границ ячейки.
 *
 * @package Topvisor\XlsxCreator\Structures\Styles\Borders
 */
class Borders{
	private $model = [];
	private $borders = [];

	public function __destruct(){
		unset($this->borders);
	}

	/**
	 * @return Border|null - левая граница
	 */
	function getLeft(){
		return $this->borders['left'] ?? null;
	}

	/**
	 * @param Border|null $border - левая граница
	 * @return Borders - $this
	 */
	function setLeft(Border $border = null) : self{
		$this->borders['left'] = $border;
		$this->model['left'] = $border->getModel();
		return $this;
	}

	/**
	 * @return Border|null - правая граница
	 */
	function getRight(){
		return $this->borders['right'] ?? null;
	}

	/**
	 * @param Border|null $border - правая граница
	 * @return Borders - $this
	 */
	function setRight(Border $border = null) : self{
		$this->borders['right'] = $border;
		$this->model['right'] = $border->getModel();
		return $this;
	}

	/**
	 * @return Border|null - верхняя граница
	 */
	function getTop(){
		return $this->borders['top'] ?? null;
	}

	/**
	 * @param Border|null $border - верхняя граница
	 * @return Borders - $this
	 */
	function setTop(Border $border = null) : self{
		$this->borders['top'] = $border;
		$this->model['top'] = $border->getModel();
		return $this;
	}

	/**
	 * @return Border|null - нижняя граница
	 */
	function getBottom(){
		return $this->borders['bottom'] ?? null;
	}

	/**
	 * @param Border|null $border - нижняя граница
	 * @return Borders - $this
	 */
	function setBottom(Border $border = null) : self{
		$this->borders['bottom'] = $border;
		$this->model['bottom'] = $border->getModel();
		return $this;
	}

	/**
	 * @return Border|null - стиль диагрнальных границ
	 */
	function getDiagonalStyle(){
		return $this->borders['diagonal'] ?? null;
	}

	/**
	 * @param Border|null $border - стиль диагональных границ
	 * @return Borders - $this
	 */
	function setDiagonalStyle(Border $border = null) : self{
		$this->borders['diagonal'] = $border;

		if (!isset($this->model['diagonal'])) $this->model['diagonal'] = $border->getModel();
		else $this->model['diagonal'] = array_merge($this->model['diagonal'], $border->getModel());

		return $this;
	}

	/**
	 * @return bool - показывать диагональную границу (из левого верхнего угла)
	 */
	function getDiagonalUp() : bool{
		if ($this->model['diagonal']) {
			return $this->model['diagonal']['up'] ?? false;
		}

		return false;
	}

	/**
	 * @param bool $diagonalUp - показывать диагональную границу (из левого верхнего угла)
	 * @return Borders - $this
	 */
	function setDiagonalUp(bool $diagonalUp) : self{
		if (!isset($this->model['diagonal'])) $this->model['diagonal'] = [];

		$this->model['diagonal']['up'] = $diagonalUp;
		return $this;
	}

	/**
	 * @return bool - показывать диагональную границу (из левого нижнего угла)
	 */
	function getDiagonalDown() : bool{
		if ($this->model['diagonal']) {
			return $this->model['diagonal']['down'] ?? false;
		}

		return false;
	}

	/**
	 * @param bool $diagonalDown - показывать диагональную границу (из левого нижнего угла)
	 * @return Borders - $this
	 */
	function setDiagonalDown(bool $diagonalDown) : self{
		if (!isset($this->model['diagonal'])) $this->model['diagonal'] = [];

		$this->model['diagonal']['down'] = $diagonalDown;
		return $this;
	}

	/**
	 * @return array - модель
	 */
	function getModel() : array{
		return $this->model;
	}
}