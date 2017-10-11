<?php

namespace Topvisor\XlsxCreator\Structures\Styles\Fills;

/**
 * Class Fill. Заливка ячейки.
 *
 * @package Topvisor\XlsxCreator\Structures\Styles\Fills
 */
abstract class Fill{
	protected $model = [];

	/**
	 * @return string - тип заливки
	 */
	function getType() : string{
		return $this->model['type'];
	}

	/**
	 * @return array - модель
	 */
	public function getModel(): array{
		return $this->model;
	}
}