<?php

namespace Topvisor\XlsxCreator\Structures\Values;

use Topvisor\XlsxCreator\Exceptions\InvalidValueException;
use Topvisor\XlsxCreator\Helpers\Validator;

/**
 * Class ErrorValue. Используется для задания значения ячейки (ошибка).
 *
 * @package  Topvisor\XlsxCreator\Structures\Values
 */
class ErrorValue extends Value {
	public const VALID_ERRORS = [
		'notApplicable' => '#N/A',
		'ref' => '#REF!',
		'name' => '#NAME?',
		'divZero' => '#DIV/0!',
		'null' => '#NULL!',
		'value' => '#VALUE!',
		'num' => '#NUM!',
	];

	/**
	 * ErrorValue constructor.
	 *
	 * @param string $error - текст ошибки
	 * @throws InvalidValueException
	 */
	public function __construct(string $error) {
		Validator::validate($error, '$error', self::VALID_ERRORS);

		parent::__construct($error, parent::TYPE_ERROR);
	}

	/**
	 * @param $value - модель
	 * @return Value - значение ячейки
	 * @throws InvalidValueException
	 */
	public static function parse($value): Value {
		return new self($value);
	}
}
