<?php

namespace Topvisor\XlsxCreator\Structures\Styles\Borders;

use Topvisor\XlsxCreator\Structures\Color;
use Topvisor\XlsxCreator\Validator;

/**
 * Class Border. Описывает стиль границы ячейки.
 *
 * @package Topvisor\XlsxCreator\Structures\Styles\Borders
 */
class Border{
	const VALID_STYLE = [
		'thin', 'dotted', 'dashDot', 'hair', 'dashDotDot', 'slantDashDot', 'mediumDashed',
		'mediumDashDotDot', 'mediumDashDot', 'medium', 'double', 'thick'
	];

	private $style;
	private $color;

	function __construct(string $style, Color $color = null){
		Validator::validate($style, '$style', self::VALID_STYLE);

		$this->style = $style;
		$this->color = $color;
	}

	/**
	 * @return string - стиль границы
	 */
	function getStyle() : string{
		return $this->style;
	}

	/**
	 * @return Color|null - цвет границы
	 */
	function getColor() : Color{
		return $this->color;
	}

	/**
	 * @return array - модель
	 */
	function getModel() : array{
		return [
			'style' => $this->style,
			'color' => $this->color
		];
	}
}