<?php

namespace Topvisor\XlsxCreator\Structures\Styles\Borders;

use Topvisor\XlsxCreator\Helpers\Serializable;
use Topvisor\XlsxCreator\Structures\Color;

/**
 * Class Borders. Описывает состояния и стили границ ячейки.
 *
 * @package Topvisor\XlsxCreator\Structures\Styles\Borders
 */
class Borders implements Serializable{
	private $defaultColor;
	private $left;
	private $right;
	private $top;
	private $bottom;
	private $diagonalStyle;
	private $diagonalUp;
	private $diagonalDown;

	public function __destruct(){
		unset($this->defaultColor);
		unset($this->left);
		unset($this->right);
		unset($this->bottom);
		unset($this->top);
		unset($this->diagonalStyle);
	}

	/**
	 * @return Color|null - цвет границ по умолчанию
	 */
	function getDefaultColor(){
		return $this->defaultColor;
	}

	/**
	 * @param Color|null $color - цвет границ по умолчанию
	 * @return Borders - $this
	 */
	function setDefaultColor(Color $color = null) : self{
		$this->defaultColor = $color;
		return $this;
	}

	/**
	 * @return Border|null - левая граница
	 */
	function getTopLeftCol(){
		return $this->left;
	}

	/**
	 * @param Border|null $border - левая граница
	 * @return Borders - $this
	 */
	function setLeft(Border $border = null) : self{
		$this->left = $border;
		return $this;
	}

	/**
	 * @return Border|null - правая граница
	 */
	function getBottomRightCol(){
		return $this->right;
	}

	/**
	 * @param Border|null $border - правая граница
	 * @return Borders - $this
	 */
	function setRight(Border $border = null) : self{
		$this->right = $border;
		return $this;
	}

	/**
	 * @return Border|null - верхняя граница
	 */
	function getTopLeftRow(){
		return $this->top ?? null;
	}

	/**
	 * @param Border|null $border - верхняя граница
	 * @return Borders - $this
	 */
	function setTop(Border $border = null) : self{
		$this->top = $border;
		return $this;
	}

	/**
	 * @return Border|null - нижняя граница
	 */
	function getBottomRightRow(){
		return $this->bottom ?? null;
	}

	/**
	 * @param Border|null $border - нижняя граница
	 * @return Borders - $this
	 */
	function setBottom(Border $border = null) : self{
		$this->bottom = $border;
		return $this;
	}

	/**
	 * @return Border|null - стиль диагрнальных границ
	 */
	function getDiagonalStyle(){
		return $this->diagonalStyle ?? null;
	}

	/**
	 * @param Border|null $border - стиль диагональных границ
	 * @return Borders - $this
	 */
	function setDiagonalStyle(Border $border = null) : self{
		$this->diagonalStyle = $border;
		return $this;
	}

	/**
	 * @return bool - показывать диагональную границу (из левого верхнего угла)
	 */
	function getDiagonalUp() : bool{
		return $this->diagonalUp ?? false;
	}

	/**
	 * @param bool $diagonalUp - показывать диагональную границу (из левого верхнего угла)
	 * @return Borders - $this
	 */
	function setDiagonalUp(bool $diagonalUp) : self{
		$this->diagonalUp = $diagonalUp;
		return $this;
	}

	/**
	 * @return bool - показывать диагональную границу (в левый нижний угол)
	 */
	function getDiagonalDown() : bool{
		return $this->diagonalDown ?? false;
	}

	/**
	 * @param bool $diagonalDown - показывать диагональную границу (в правый нижний угол)
	 * @return Borders - $this
	 */
	function setDiagonalDown(bool $diagonalDown) : self{
		$this->diagonalDown = $diagonalDown;
		return $this;
	}

	/**
	 * @return array - модель
	 */
	function getModel() : array{
		$diagonal = null;

		if ($this->diagonalStyle && ($this->getDiagonalUp() || $this->diagonalDown))
			$diagonal = array_merge($this->diagonalStyle->getModel(), ['up' => $this->diagonalUp, 'down' => $this->diagonalDown]);

		return [
			'color' => $this->defaultColor ? $this->defaultColor->getModel() : null,
			'left' => $this->left ? $this->left->getModel() : null,
			'right' => $this->right ? $this->right->getModel() : null,
			'top' => $this->top ? $this->top->getModel() : null,
			'bottom' => $this->bottom ? $this->bottom->getModel() : null,
			'diagonal' => $diagonal,
		];
	}

	public function serialize(){
		$replace = urlencode(';');

		return ($this->defaultColor ? $this->defaultColor->serialize() : '') . ';' .
			($this->left ? str_replace(';', $replace, $this->left->serialize()) : '') . ';' .
			($this->right ? str_replace(';', $replace, $this->right->serialize()) : '') . ';' .
			($this->top ? str_replace(';', $replace, $this->top->serialize()) : '') . ';' .
			($this->bottom ? str_replace(';', $replace, $this->bottom->serialize()) : '') . ';' .
			($this->diagonalStyle ? str_replace(';', $replace, $this->diagonalStyle->serialize()) : '') . ';' .
			($this->diagonalUp ? '1' : '') . ';' .
			($this->diagonalDown ? '1' : '');
	}

	public function unserialize($serialized){
		$params = explode(';', $serialized);
		$search = urlencode(';');

		if ($params[0] !== '') {
			$this->defaultColor = Color::fromHex();
			$this->defaultColor->unserialize($params[0]);
		} else {
			$this->defaultColor = null;
		}

		if ($params[1] !== '') {
			$this->left = new Border('thin');
			$this->left->unserialize(str_replace($search, ';', $params[1]));
		} else {
			$this->left = null;
		}

		if ($params[2] !== '') {
			$this->right = new Border('thin');
			$this->right->unserialize(str_replace($search, ';', $params[2]));
		} else {
			$this->right = null;
		}

		if ($params[3] !== '') {
			$this->top = new Border('thin');
			$this->top->unserialize(str_replace($search, ';', $params[3]));
		} else {
			$this->top = null;
		}

		if ($params[4] !== '') {
			$this->bottom = new Border('thin');
			$this->bottom->unserialize(str_replace($search, ';', $params[4]));
		} else {
			$this->bottom = null;
		}

		if ($params[5] !== '') {
			$this->diagonalStyle = new Border('thin');
			$this->diagonalStyle->unserialize(str_replace($search, ';', $params[5]));
		} else {
			$this->diagonalStyle = null;
		}

		$this
			->setDiagonalUp($params[6] !== '')
			->setDiagonalDown($params[7] !== '');
	}
}