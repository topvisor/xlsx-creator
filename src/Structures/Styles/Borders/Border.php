<?php

namespace Topvisor\XlsxCreator\Structures\Styles\Borders;

use Serializable;
use Topvisor\XlsxCreator\Structures\Color;
use Topvisor\XlsxCreator\Helpers\Validator;

/**
 * Class Border. Описывает стиль границы ячейки.
 *
 * @package Topvisor\XlsxCreator\Structures\Styles\Borders
 */
class Border implements Serializable{
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
			'color' => $this->color->getModel()
		];
	}

	public function serialize(){
		$serialized = array_search($this->style, self::VALID_STYLE);
		if ($this->color) $serialized .= ';' . $this->color->serialize();
		return $serialized;
	}

	public function unserialize($serialized){
		$params = explode(';', $serialized);

		$this->style = self::VALID_STYLE[(int) $params[0]];

		if (isset($params[1])) {
			$this->color = Color::fromHex();
			$this->color->unserialize($params[1]);
		} else {
			$this->color = null;
		}
	}
}