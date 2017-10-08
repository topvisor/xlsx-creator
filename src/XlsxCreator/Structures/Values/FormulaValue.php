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
	 * @param string|null $result - результат вычисления формулы (если известен)
	 */
	function __construct(string $formula, string $result = null){
		$value = ['formula' => $formula];
		if (!is_null($result)) $value['result'] = $result;

		parent::__construct($value, parent::TYPE_FORMULA);
	}

	/**
	 * @param $model - модель
	 * @return Value - значение ячейки
	 */
	static function parse($model): Value{
		return new self($model['formula'], $model['result'] ?? null);
	}
}