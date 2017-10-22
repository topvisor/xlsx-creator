<?php

namespace Topvisor\XlsxCreator\Structures\Styles\Alignment;

use Serializable;
use Topvisor\XlsxCreator\Helpers\Validator;

/**
 * Class Alignment. Описывает выравнивание текста.
 *
 * @package Topvisor\XlsxCreator\Structures\Styles
 */
class Alignment implements Serializable{
	const VALID_HORIZONTAL = ['left', 'center', 'right', 'fill', 'centerContinuous', 'distributed', 'justify'];
	const VALID_VERTICAL = ['top', 'center', 'bottom', 'distributed', 'justify'];
	const VALID_READING_ORDER = ['leftToRight', 'rightToLeft'];

	private $horizontal;
	private $vertical;
	private $wrapText;
	private $indent;
	private $readingOrder;
	private $textRotation;

	public function __destruct(){
		unset($this->textRotation);
	}

	/**
	 * @return string|null - выравнивание по горизонтали
	 */
	function getHorizontal(){
		return $this->horizontal;
	}

	/**
	 * @param string|null $horizontal - выравнивание по горизонтали
	 * @return Alignment - $this
	 */
	function setHorizontal(string $horizontal = null) : self{
		if (!is_null($horizontal)) {
			Validator::validate($horizontal, '$horizontal', self::VALID_HORIZONTAL);
			$this->setIndent(null);
		}

		$this->horizontal = $horizontal;
		return $this;
	}

	/**
	 * @return string|null - выравнивание по вертикали
	 */
	function getVertical(){
		return $this->vertical;
	}

	/**
	 * @param string|null $vertical - выравнивание по вертикали
	 * @return Alignment - $this
	 */
	function setVertical(string $vertical = null) : self{
		if (!is_null($vertical)) Validator::validate($vertical, '$vertical', self::VALID_VERTICAL);

		$this->vertical = $vertical;
		return $this;
	}

	/**
	 * @return bool - распределять по ширине
	 */
	function getWrapText() : bool{
		return $this->wrapText ?? false;
	}

	/**
	 * @param bool $wrapText - заполнить ячейку текстом
	 * @return Alignment - $this
	 */
	function setWrapText(bool $wrapText) : self{
		$this->wrapText = $wrapText;
		return $this;
	}

	/**
	 * @return int|null - отступ слева
	 */
	function getIndent(){
		return $this->indent;
	}

	/**
	 * @param int|null $indent - отступ слева
	 * @return Alignment - $this
	 */
	function setIndent(int $indent = null) : self{
		if (!is_null($indent)) {
			Validator::validateInRange($indent, 0, 250, '$indent');
			$this->setHorizontal(null);
		}

		$this->indent = $indent;
		return $this;
	}

	/**
	 * @return string|null - направление чтения
	 */
	function getReadingOrder(){
		return $this->readingOrder;
	}

	/**
	 * @param string|null $readingOrder - направление чтения
	 * @return Alignment - $this
	 */
	function setReadingOrder(string $readingOrder = null) : self{
		if (!is_null($readingOrder)) Validator::validate($readingOrder, '$readingOrder', self::VALID_READING_ORDER);

		$this->readingOrder = $readingOrder;
		return $this;
	}

	/**
	 * @return TextRotation|null - поворот текста
	 */
	function getTextRotation(){
		return $this->textRotation;
	}

	/**
	 * @param TextRotation|null $textRotation - поворот текста
	 * @return Alignment - $this
	 */
	function setTextRotation(TextRotation $textRotation = null) : self{
		$this->textRotation = $textRotation;
		return $this;
	}

	/**
	 * @return array - модель
	 */
	function getModel(): array{
		return [
			'horizontal' => $this->horizontal,
			'vertical' => $this->vertical,
			'wrapText' => $this->wrapText,
			'indent' => $this->indent,
			'readingOrder' => $this->readingOrder,
			'textRotation' => $this->textRotation ? $this->textRotation->getModel() : null
		];
	}

	public function serialize(){
		return ($this->horizontal ? array_search($this->horizontal, self::VALID_HORIZONTAL) : '') . ';' .
			($this->vertical ? array_search($this->vertical, self::VALID_VERTICAL) : '') . ';' .
			($this->wrapText ? '1' : '') . ';' .
			($this->indent ? $this->indent : '') . ';' .
			($this->readingOrder ? array_search($this->readingOrder, self::VALID_READING_ORDER) : '') . ';' .
			($this->textRotation ? $this->textRotation->serialize() : '');
	}

	public function unserialize($serialized){
		$params = explode(';', $serialized);

		if ($params[5]) {
			$this->textRotation = TextRotation::vertical();
			$this->textRotation->unserialize($params[5]);
		} else {
			$this->textRotation = null;
		}

		$this
			->setHorizontal($params[0] ? self::VALID_HORIZONTAL[(int) $params[0]] : null)
			->setVertical($params[1] ? self::VALID_VERTICAL[(int) $params[1]] : null)
			->setWrapText((bool) $params[2])
			->setIndent($params[3] ? (int) $params[3] : null)
			->setReadingOrder($params[4] ? self::VALID_READING_ORDER[(int) $params[4]] : null);
	}
}