<?php

namespace XlsxCreator\Structures\Values;

/**
 * Class ErrorValue. Используется для задания значения ячейки (ошибка).
 *
 * @package XlsxCreator\Structures\Values
 */
class ErrorValue extends Value{
	/**
	 * ErrorValue constructor.
	 *
	 * @param string $error - текст ошибки
	 */
	function __construct(string $error){
		parent::__construct($error, parent::TYPE_ERROR);
	}

	/**
	 * @param $value - модель
	 * @return Value - значение ячейки
	 */
	static function parse($value): Value{
		return new self($value);
	}
}