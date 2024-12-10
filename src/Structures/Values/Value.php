<?php

namespace Topvisor\XlsxCreator\Structures\Values;

use DateTime;
use Topvisor\XlsxCreator\Exceptions\InvalidValueException;

/**
 * Class Value. Используется для задания значения ячейки.
 *
 * @package  Topvisor\XlsxCreator\Structures\Values
 */
class Value {
	// Типы значений ячеек
	public const TYPE_NULL = 0;
	public const TYPE_NUMBER = 1;
	public const TYPE_STRING = 2;
	public const TYPE_DATE = 3;
	public const TYPE_HYPERLINK = 4;
	public const TYPE_FORMULA = 5;
	public const TYPE_SHARED_STRING = 6;
	public const TYPE_RICH_TEXT = 7;
	public const TYPE_BOOL = 8;
	public const TYPE_ERROR = 9;

	protected $type;
	protected $value;

	/**
	 * Value constructor.
	 *
	 * @param $value - Значение ячейки
	 * @param int $type - Тип ячейки
	 */
	protected function __construct($value, int $type) {
		$this->value = $value;
		$this->type = $type;
	}

	/**
	 * Создает объект Value, определяет его тип
	 *
	 * @param $value - значение
	 * @throws InvalidValueException
	 */
	public static function parse($value): self {
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
	public function getValue() {
		return $this->value;
	}

	/**
	 * @return int - тип
	 */
	public function getType(): int {
		return $this->type;
	}
}
