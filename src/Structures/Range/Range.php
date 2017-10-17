<?php

namespace Topvisor\XlsxCreator\Structures\Range;

use Topvisor\XlsxCreator\Cell;
use Topvisor\XlsxCreator\Exceptions\InvalidValueException;
use Topvisor\XlsxCreator\Helpers\Validator;

/**
 * Class Range. Диапазон.
 *
 * @package Topvisor\XlsxCreator\Structures
 */
class Range{
	protected $top;
	protected $left;
	protected $bottom;
	protected $right;

	/**
	 * Range constructor.
	 *
	 * @param float $row1 - номер первой строки
	 * @param float $col1 - номер первого столбца
	 * @param float $row2 - номер второй строки
	 * @param float $col2 - номер второго столбца
	 * @throws InvalidValueException
	 */
	function __construct(float $row1, float $col1, float $row2, float $col2){
		Validator::validateInRange($row1, 1, 1048576, '$row1');
		Validator::validateInRange($col1, 1, 16384, '$col1');
		Validator::validateInRange($row2, 1, 1048576, '$row2');
		Validator::validateInRange($col2, 1, 16384, '$col2');

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
	 * @return float - левая колонка
	 */
	function getLeft() : float{
		return $this->left;
	}

	/**
	 * @return float - верхняя строка
	 */
	function getTop() : float{
		return $this->top;
	}

	/**
	 * @return float - правая колонка
	 */
	function getRight() : float{
		return $this->right;
	}

	/**
	 * @return float - нижняя строка
	 */
	function getBottom() : float{
		return $this->bottom;
	}

	/**
	 * @param Range $range - диапазон
	 * @return Range|null - пересечение
	 */
	function intersection(Range $range) {
		if ($range->right >= $this->left && $range->right <= $this->right
			&& $range->bottom >= $this->top && $range->bottom <= $this->bottom){

			$right = $range->right;
			$bottom = $range->bottom;

			if ($this->left <= $range->left) $left = $range->left;
			else $left = $this->left;

			if ($this->top <= $range->top) $top = $range->top;
			else $top = $this->top;
		} elseif ($range->left >= $this->left && $range->left <= $this->right
			&& $range->top >= $this->top && $this->top <= $this->bottom) {

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