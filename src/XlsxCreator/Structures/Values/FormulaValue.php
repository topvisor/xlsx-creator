<?php

namespace XlsxCreator\Structures\Values;

/**
 * Class FormulaValue. Используется для задания значения ячейки (формула).
 *
 * @package XlsxCreator\Structures\Values
 */
class FormulaValue extends Value{
	/**
	 * FormulaValue constructor.
	 *
	 * @param string $formula - формула
	 */
	function __construct(string $formula){
		parent::__construct($formula, parent::TYPE_FORMULA);
	}

	/**
	 * @param $model - модель
	 * @return Value - значение ячейки
	 */
	static function parse($model): Value{
		return new self($model);
	}
}