<?php

namespace Topvisor\XlsxCreator\Structures\Styles;

use Topvisor\XlsxCreator\Structures\Color;
use Topvisor\XlsxCreator\Validator;

/**
 * Class Font. Описывает стили шрифта.
 *
 * @package Topvisor\XlsxCreator\Structures\Styles
 */
class Font{
	const VALID_UNDERLINE = ['single', 'double', 'singleAccounting', 'doubleAccounting'];
	const VALID_SCHEME = ['minor', 'major', 'none'];

	private $model = [];
	private $color;

	public function __destruct(){
		unset($this->color);
	}

	/**
	 * @return string|null - название
	 */
	function getName(){
		return $this->model['name'] ?? null;
	}

	/**
	 * @param string|null $name - название
	 * @return Font - $this
	 */
	function setName(string $name = null) : self{
		$this->model['name'] = $name;
		return $this;
	}

	/**
	 * @return int|null - размер
	 */
	function getSize(){
		return $this->model['sz'] ?? null;
	}

	/**
	 * @param int|null $size - размер
	 * @return Font - $this
	 */
	function setSize(int $size = null) : self{
		if (!is_null($size)) Validator::validateInRange($size, 1, 409, '$size');

		$this->model['sz'] = $size;
		return $this;
	}

	/**
	 * @return Color|null - цвет
	 */
	function getColor(){
		return $this->color;
	}

	/**
	 * @param Color|null $color - цвет
	 * @return Font - $this
	 */
	function setColor(Color $color = null) : self{
		$this->color = $color;
		$this->model['color'] = $color ? $color->getModel() : null;
		return $this;
	}

	/**
	 * @return bool - жирный
	 */
	function getBold() : bool{
		return $this->model['b'] ?? false;
	}

	/**
	 * @param bool $bold - жирный
	 * @return Font -  $this
	 */
	function setBold(bool $bold) : self{
		$this->model['b'] = $bold;
		return $this;
	}

	/**
	 * @return bool - курсивный
	 */
	function getItalic() : bool{
		return $this->model['i'] ?? false;
	}

	/**
	 * @param bool $italic - курсивный
	 * @return Font - $this
	 */
	function setItalic(bool $italic) : self{
		$this->model['i'] = $italic;
		return $this;
	}

	/**
	 * @return string|null - тип подчеркивания
	 */
	function getUnderline(){
		return $this->model['u'] ?? null;
	}

	/**
	 * @param string|null $underline - тип подчеркивания
	 * @return Font - $this
	 */
	function setUnderline(string $underline = null) : self{
		if (!is_null($underline)) Validator::validate($underline, '$underline', self::VALID_UNDERLINE);

		$this->model['u'] = $underline;
		return $this;
	}

	/**
	 * @return string|null - надстрочный/подстрочный
	 */
	function getScheme(){
		return $this->model['scheme'] ?? null;
	}

	/**
	 * @param string|null $scheme - надстрочный/подстрочный
	 * @return Font - $this
	 */
	function setScheme(string $scheme = null) : self{
		if (!is_null($scheme)) Validator::validate($scheme, '$scheme', self::VALID_SCHEME);

		$this->model['scheme'] = $scheme;
		return $this;
	}

	/**
	 * @return bool - зачеркнутый
	 */
	function getStrike() : bool{
		return $this->model['strike'] ?? false;
	}

	/**
	 * @param bool $strike - зачеркнутый
	 * @return Font - $this
	 */
	function setStrike(bool $strike) : self{
		$this->model['strike'] = $strike;
		return $this;
	}

	/**
	 * @return array - модель
	 */
	function getModel(): array{
		return $this->model;
	}
}