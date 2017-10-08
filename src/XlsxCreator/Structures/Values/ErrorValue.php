<?php

namespace XlsxCreator\Structures\Values;
use XlsxCreator\Exceptions\InvalidValueException;

/**
 * Class ErrorValue. Используется для задания значения ячейки (ошибка).
 *
 * @package XlsxCreator\Structures\Values
 */
class ErrorValue extends Value{
	const VALID_ERRORS = [
		'notApplicable' => '#N/A',
		'ref' => '#REF!',
		'name' => '#NAME?',
		'divZero' => '#DIV/0!',
		'null' => '#NULL!',
		'value' => '#VALUE!',
		'num' => '#NUM!'
	];

	/**
	 * ErrorValue constructor.
	 *
	 * @param string $error - текст ошибки
	 * @throws InvalidValueException
	 */
	function __construct(string $error){
		if (!in_array($error, self::VALID_ERRORS))
			throw new InvalidValueException('$error must be in [\'' . implode("','",self::VALID_ERRORS) . "']");

		parent::__construct($error, parent::TYPE_ERROR);
	}

	/**
	 * @param $value - модель
	 * @return Value - значение ячейки
	 * @throws InvalidValueException
	 */
	static function parse($value): Value{
		return new self($value);
	}
}