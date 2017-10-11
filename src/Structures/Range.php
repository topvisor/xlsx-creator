<?php

namespace Topvisor\XlsxCreator\Structures;

/**
 * Class Range. Диапазон.
 *
 * @package Topvisor\XlsxCreator\Structures
 */
class Range{
	private $top;
	private $bottom;
	private $left;
	private $right;

	public function __construct(int $row1, int $col1, int $row2, int $col2){
		if ($row1 <= $row2) {
			$this->top = $row1;
			$this->bottom = $row2;
		} else {
			$this->top = $row2;
			$this->bottom = $row1;
		}

		if ($col1 <= $col2) {
			$this->left = $col1;
			$this->right = $col2;
		} else {
			$this->left = $col2;
			$this->right = $col1;
		}
	}

	/**
	 * @return int - левая колонка
	 */
	public function getLeft() : int{
		return $this->left;
	}

	/**
	 * @return int - верхняя строка
	 */
	public function getTop() : int{
		return $this->top;
	}

	/**
	 * @return int - правая колонка
	 */
	public function getRight() : int{
		return $this->right;
	}

	/**
	 * @return int - нижняя строка
	 */
	public function getBottom() : int{
		return $this->bottom;
	}

	/**
	 * @param Range $range - диапазон
	 * @return Range|null - пересечение
	 */
	function intersection(Range $range) {
		if ($this->left <= $range->right && $this->top <= $range->bottom){
			$right = $range->right;
			$bottom = $range->bottom;

			if ($this->left <= $range->left) $left = $range->left;
			else $left = $this->left;

			if ($this->top <= $range->top) $top = $range->top;
			else $top = $this->top;
		} elseif ($this->right >= $range->left && $this->bottom >= $range->top) {
			$left = $range->left;
			$top = $range->bottom;

			if ($this->right <= $range->right) $right = $range->right;
			else $right = $this->right;

			if ($this->bottom <= $range->bottom) $bottom = $range->bottom;
			else $bottom = $this->bottom;
		} else {
			return null;
		}

		return new Range($top, $left, $bottom, $right);
	}
}