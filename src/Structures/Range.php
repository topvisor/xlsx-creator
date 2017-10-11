<?php

namespace Topvisor\XlsxCreator\Structures;
use Topvisor\XlsxCreator\Cell;
use Topvisor\XlsxCreator\Exceptions\InvalidValueException;
use Topvisor\XlsxCreator\Validator;

/**
 * Class Range. Диапазон.
 *
 * @package Topvisor\XlsxCreator\Structures
 */
class Range{
	private $top;
	private $left;
	private $bottom;
	private $right;

	function __construct(int $row1, int $col1, int $row2, int $col2){
		if ($row1 == $row2 && $col1 == $col2) throw new InvalidValueException("It's not range");

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
	 * @return int - левая колонка
	 */
	function getLeft() : int{
		return $this->left;
	}

	/**
	 * @return int - верхняя строка
	 */
	function getTop() : int{
		return $this->top;
	}

	/**
	 * @return int - правая колонка
	 */
	function getRight() : int{
		return $this->right;
	}

	/**
	 * @return int - нижняя строка
	 */
	function getBottom() : int{
		return $this->bottom;
	}

	function __toString(){
		return Cell::genColStr($this->left) . $this->top . ':' . Cell::genColStr($this->right) . $this->bottom;
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

	/**
	 * @param string $range - диапазон в формате (A1:B4)
	 * @return Range - диапазон
	 * @throws InvalidValueException
	 */
	static function fromString(string $range) : self{
		$addresses = explode(':', $range);

		if (count($addresses) !== 2) throw new InvalidValueException("Unavailable cell's range");

		if (!preg_match('/^([A-Z]{1,3})(\d{1,5})$/', $addresses[0], $matches)) throw new InvalidValueException('Unavailable address format');
		$col1 = Cell::genColNum($matches[1]);
		$row1 = (int) $matches[2];

		if (!preg_match('/^([A-Z]{1,3})(\d{1,5})$/', $addresses[1], $matches)) throw new InvalidValueException('Unavailable address format');
		$col2 = Cell::genColNum($matches[1]);
		$row2 = (int) $matches[2];

		return new self($row1, $col1, $row2, $col2);
	}
}