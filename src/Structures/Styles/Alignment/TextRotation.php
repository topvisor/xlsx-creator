<?php

namespace Topvisor\XlsxCreator\Structures\Styles\Alignment;

use Topvisor\XlsxCreator\Validator;

/**
 * Class TextRotation. Описывает поворот текста.
 *
 * @package Topvisor\XlsxCreator\Structures\Styles
 */
class TextRotation{
	private $model;

	/**
	 * TextRotation constructor.
	 *
	 * @param int|string $model - модель
	 */
	private function __construct($model){
		$this->model = $model;
	}

	/**
	 * @param int $angle - угол (-90;90) поворота текста
	 * @return TextRotation
	 */
	static function fromAngle(int $angle) : self{
		Validator::validateInRange($angle, -90, 90, '$angle');
		if ($angle < 0) $angle = 90 - $angle;

		return new self($angle);
	}

	/**
	 * Текст по вертикали. Например:
	 *
	 * Т
	 * е
	 * к
	 * с
	 * т
	 *
	 * @return TextRotation
	 */
	static function vertical() : self{
		return new self('vertical');
	}

	/**
	 * @return int|string - модель
	 */
	function getModel(){
		return $this->model;
	}
}