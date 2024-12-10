<?php

namespace Topvisor\XlsxCreator\Structures\Values;

/**
 * Class FormulaValue. Используется для задания значения ячейки (формула).
 *
 * @package  Topvisor\XlsxCreator\Structures\Values
 */
class FormulaValue extends Value {
	/**
	 * FormulaValue constructor.
	 *
	 * @param string $formula - формула
	 */
	public function __construct(string $formula) {
		parent::__construct($formula, parent::TYPE_FORMULA);
	}

	/**
	 * @param $model - модель
	 * @return Value - значение ячейки
	 */
	public static function parse($model): Value {
		return new self($model);
	}
}
