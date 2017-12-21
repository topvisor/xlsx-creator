<?php

namespace Topvisor\XlsxCreator\Structures\Range;
use Topvisor\XlsxCreator\Exceptions\InvalidValueException;
use Topvisor\XlsxCreator\Helpers\Validator;

/**
 * Class Coords. Координаты.
 *
 * @package Topvisor\XlsxCreator\Structures
 */
class Coords{
	private $row;
	private $col;

	/**
	 * Сoords constructor.
	 *
	 * @param float $row - номер строки
	 * @param float $col - номер столбца
	 * @throws InvalidValueException
	 */
	function __construct(float $row, float $col){
		Validator::validateInRange($row, 1, 1048576, '$row');
		Validator::validateInRange($col, 1, 16384, '$col');

		$this->row = $row;
		$this->col = $col;
	}

	/**
	 * @return float - строка
	 */
	function getRow() : float{
		return $this->row;
	}

	/**
	 * @return float - колонка
	 */
	function getCol() : float{
		return $this->col;
	}
}

