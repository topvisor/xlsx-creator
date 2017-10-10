<?php

namespace Topvisor\XlsxCreator\Structures\Styles\Borders;

use Topvisor\XlsxCreator\Structures\Color;
use Topvisor\XlsxCreator\Validator;

class Border{
	const VALID_STYLE = [
		'thin', 'dotted', 'dashDot', 'hair', 'dashDotDot', 'slantDashDot', 'mediumDashed',
		'mediumDashDotDot', 'mediumDashDot', 'medium', 'double', 'thick'
	];

	private $model;

	function __construct(string $style, Color $color = null){
		Validator::validate($style, '$style', self::VALID_STYLE);

		$this->model = [
			'style' => $style,
			'color' => $color
		];
	}

	/**
	 * @return string - стиль границы
	 */
	function getStyle() : string{
		return $this->model['style'];
	}

	/**
	 * @return Color|null - цвет границы
	 */
	function getColor() : Color{
		return $this->model['color'];
	}

	/**
	 * @return array - модель
	 */
	function getModel() : array{
		return $this->model;
	}
}