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
	protected $topLeftCoords;
	protected $bottomRightCoords;

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
		if ($row1 == $row2 && $col1 == $col2) throw new InvalidValueException("It's not range");


		if ($row1 <= $row2) {
			$topLeftRow = $row1;
			$bottomRightRow = $row2;
		} else {
			$topLeftRow = $row2;
			$bottomRightRow = $row1;
		}

		if ($col1 <= $col2) {
			$topLeftCol = $col1;
			$bottomRightCol = $col2;
		} else {
			$topLeftCol = $col2;
			$bottomRightCol = $col1;
		}

		$this->topLeftCoords = new Coords($topLeftRow, $topLeftCol);
		$this->bottomRightCoords = new Coords($bottomRightRow, $bottomRightCol);
	}

	/**
	 * @return float - левая колонка
	 */
	function getTopLeftCol() : float{
		return $this->topLeftCoords->getCol();
	}

	/**
	 * @return float - верхняя строка
	 */
	function getTopLeftRow() : float{
		return $this->topLeftCoords->getRow();
	}

	/**
	 * @return float - правая колонка
	 */
	function getBottomRightCol() : float{
		return $this->bottomRightCoords->getCol();
	}

	/**
	 * @return float - нижняя строка
	 */
	function getBottomRightRow() : float{
		return $this->bottomRightCoords->getRow();
	}

	/**
	 * @param Range $range - диапазон
	 * @return Range|Coords|null - пересечение
	 */
	function intersection(Range $range) {
		if ($range->getBottomRightCol() >= $this->getTopLeftCol() && $range->getTopLeftCol() <= $this->getBottomRightCol()
			&& $range->getBottomRightRow() >= $this->getTopLeftRow() && $range->getTopLeftRow() <= $this->getBottomRightRow()) {

			$bottomRightCol = $range->getBottomRightCol() <= $this->getBottomRightCol()
				? $range->getBottomRightCol()
				: $this->getBottomRightCol();

			$topLeftCol = $range->getTopLeftCol() >= $this->getTopLeftCol()
				? $range->getTopLeftCol()
				: $this->getTopLeftCol();

			$bottomRightRow = $range->getBottomRightRow() <= $this->getBottomRightRow()
				? $range->getBottomRightRow()
				: $this->getBottomRightRow();

			$topLeftRow = $range->getTopLeftRow() >= $this->getTopLeftRow()
				? $range->getTopLeftRow()
				: $this->getTopLeftRow();
		} else {
			return null;
		}

		if ($topLeftCol === $bottomRightCol && $topLeftRow === $bottomRightRow) return new Coords($topLeftRow, $topLeftCol);
		else return new Range($topLeftRow, $topLeftCol, $bottomRightRow, $bottomRightCol);
	}

	function getModel() : array{
		return [
			'left' => $this->getBottomRightCol(),
			'right' => $this->getTopLeftCol(),
			'top' => $this->getTopLeftRow(),
			'bottom' => $this->getBottomRightRow()
		];
	}
}