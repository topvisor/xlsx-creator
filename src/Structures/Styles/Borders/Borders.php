<?php

namespace Topvisor\XlsxCreator\Structures\Styles\Borders;

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
	 * @return Border|null - левая граница
	 */
	function getRight(){
		return $this->borders['right'] ?? null;
	}

	/**
	 * @param Border|null $border - левая граница
	 * @return Borders - $this
	 */
	function setRight(Border $border = null) : self{
		$this->borders['right'] = $border;
		$this->model['right'] = $border->getModel();
		return $this;
	}

	/**
	 * @return Border|null - левая граница
	 */
	function getTop(){
		return $this->borders['top'] ?? null;
	}

	/**
	 * @param Border|null $border - левая граница
	 * @return Borders - $this
	 */
	function setTop(Border $border = null) : self{
		$this->borders['top'] = $border;
		$this->model['top'] = $border->getModel();
		return $this;
	}

	/**
	 * @return Border|null - левая граница
	 */
	function getBottom(){
		return $this->borders['bottom'] ?? null;
	}

	/**
	 * @param Border|null $border - левая граница
	 * @return Borders - $this
	 */
	function setBottom(Border $border = null) : self{
		$this->borders['bottom'] = $border;
		$this->model['bottom'] = $border->getModel();
		return $this;
	}

	/**
	 * @return Border|null - левая граница
	 */
	function getDiagonal(){
		return $this->borders['diagonal'] ?? null;
	}

	/**
	 * @param Border|null $border - левая граница
	 * @return Borders - $this
	 */
	function setDiagonal(Border $border = null) : self{
		$this->borders['diagonal'] = $border;
		$this->model['diagonal'] = $border->getModel();
		return $this;
	}

	function getDiagonalUp() : bool{
		if ($this->model['diagonal']) {
			return $this->model['diagonal']['up'] ?? false;
		}

		return false;
	}

	#####
//	function setDiagonalUp(bool $diagonalUp) : self{
//		if ($this->model['diagonal']) {
//			return $this->model['diagonal']['up'] ?? false;
//		}
//
//		return false;
//	}

	function getDiagonalDown() : bool{
		if ($this->model['diagonal']) {
			return $this->model['diagonal']['down'] ?? false;
		}

		return false;
	}

	/**
	 * @return array - модель
	 */
	function getModel() : array{
		return $this->model;
	}
}