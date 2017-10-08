<?php

namespace XlsxCreator\Structures\Values;

use DateTime;
use XlsxCreator\Exceptions\InvalidValueException;

/**
 * Class Value. Используется для задания значения ячейки.
 *
 * @package XlsxCreator\Structures\Values
 */
class Value{
	// Типы значений ячеек
	const TYPE_NULL = 0;
//	const TYPE_MERGE = 1;
	const TYPE_NUMBER = 2;
	const TYPE_STRING = 3;
	const TYPE_DATE = 4;
	const TYPE_HYPERLINK = 5;
	const TYPE_FORMULA = 6;
//	const TYPE_RICH_TEXT = 7;
	const TYPE_BOOL = 8;
	const TYPE_ERROR = 9;

	protected $type;
	protected $value;

	/**
	 * Value constructor.
	 *
	 * @param $value - Значение ячейки
	 * @param int $type - Тип ячейки
	 */
	protected function __construct($value, int $type){
		$this->value = $value;
		$this->type = $type;
	}

	/**
	 * Создает объект Value, определяет его тип
	 *
	 * @param $value - значение
	 * @return Value
	 * @throws InvalidValueException
	 */
	static function parse($value) : self{
		switch (true) {
			case is_null($value):
				$type = self::TYPE_NULL;
				break;

			case is_string($value):
				$type = self::TYPE_STRING;
				break;

			case is_numeric($value):
				$type = self::TYPE_NUMBER;
				break;

			case is_bool($value):
				$type = self::TYPE_BOOL;
				break;

			case ($value instanceof DateTime):
				$type = self::TYPE_DATE;
				break;

			default:
				throw new InvalidValueException('$value must be null, string, numeric, bool or DateTime');
				break;
		}

		return new self($value, $type);
	}

	/**
	 * @return mixed - значение
	 */
	function getValue(){
		return $this->value;
	}

	/**
	 * @return int - тип
	 */
	function getType() : int{
		return $this->type;
	}
}