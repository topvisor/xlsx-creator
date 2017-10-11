<?php

namespace Topvisor\XlsxCreator\Structures\Styles\Fills;

use Topvisor\XlsxCreator\Structures\Color;
use Topvisor\XlsxCreator\Validator;

/**
 * Class PatternFill. Заливка ячейки по шаблону.
 *
 * @package Topvisor\XlsxCreator\Structures\Styles\Fills
 */
class PatternFill extends Fill{
	const VALID_PATTERN_TYPE = [
		'none', 'solid', 'darkVertical', 'darkGray', 'mediumGray', 'lightGray', 'gray125', 'gray0625', 'darkHorizontal', 'darkVertical', 'darkDown',
		'darkUp', 'darkGrid', 'darkTrellis', 'lightHorizontal', 'lightVertical', 'lightDown', 'lightUp', 'lightGrid', 'lightTrellis', 'lightGrid'
	];

	private $fgColor;
	private $bgColor;

	/**
	 * PatternFill constructor.
	 *
	 * @param string $patternType - тип шаблона заливки
	 */
	public function __construct(string $patternType = 'none'){
		Validator::validate($patternType, '$patternType', self::VALID_PATTERN_TYPE);

		$this->model['type'] = 'pattern';
		$this->model['pattern'] = $patternType;
	}

	public function __destruct(){
		unset($this->fgColor);
		unset($this->bgColor);
	}

	/**
	 * @return string - тип шаблона заливки
	 */
	function getPatternType() : string{
		return $this->model['pattern'];
	}

	/**
	 * @param string $patternType - тип шаблона заливки
	 * @return PatternFill - $this
	 */
	function setPattenType(string $patternType) : self{
		Validator::validate($patternType, '$patternType', self::VALID_PATTERN_TYPE);

		$this->model['pattern'] = $patternType;
		return $this;
	}

	/**
	 * @return Color|null - цвет переднего фона
	 */
	function getFgColor(){
		return $this->fgColor;
	}

	/**
	 * @param Color|null $fgColor - цвет переднего фона
	 * @return PatternFill - $this
	 */
	function setFgColor(Color $fgColor = null) : self{
		$this->fgColor = $fgColor;
		$this->model['fgColor'] = $fgColor ? $fgColor->getModel() : null;
		return $this;
	}

	/**
	 * @return Color|null - цвет заднего фона
	 */
	function getBgColor(){
		return $this->bgColor;
	}

	/**
	 * @param Color|null $bgColor - цвет заднего фона
	 * @return PatternFill - $this
	 */
	function setBgColor(Color $bgColor = null) : self{
		$this->bgColor = $bgColor;
		$this->model['bgColor'] = $bgColor ? $bgColor->getModel() : null;
		return $this;
	}
}