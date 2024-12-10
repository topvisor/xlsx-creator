<?php

namespace Topvisor\XlsxCreator\Structures\Values;

use Topvisor\XlsxCreator\Helpers\Validator;

/**
 * Class SharedStringValue. Описывает общую строку.
 *
 * @package Topvisor\XlsxCreator\Structures\Values
 */
class SharedStringValue extends Value {
	/**
	 * SharedStringValue constructor.
	 *
	 * @param int $id - ид общей строки
	 */
	public function __construct(int $id) {
		Validator::validatePositive($id, '$id');

		parent::__construct($id, Value::TYPE_SHARED_STRING);
	}

	/**
	 * @param $value - модель
	 * @return Value - значение ячейки
	 */
	public static function parse($value): Value {
		return new self($value);
	}
}
