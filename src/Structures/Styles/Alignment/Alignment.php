<?php

namespace Topvisor\XlsxCreator\Structures\Styles\Alignment;

use Topvisor\XlsxCreator\Validator;

/**
 * Class Alignment. Описывает выравнивание текста.
 *
 * @package Topvisor\XlsxCreator\Structures\Styles
 */
class Alignment{
	const VALID_HORIZONTAL = ['left', 'center', 'right', 'fill', 'centerContinuous', 'distributed', 'justify'];
	const VALID_VERTICAL = ['top', 'center', 'bottom', 'distributed', 'justify'];
	const VALID_READING_ORDER = ['leftToRight', 'rightToLeft'];

	private $model = [];
	private $textRotation;

	public function __destruct(){
		unset($this->textRotation);
	}

	/**
	 * @return string|null - выравнивание по горизонтали
	 */
	function getHorizontal(){
		return $this->model['horizontal'] ?? null;
	}

	/**
	 * @param string|null $horizontal - выравнивание по горизонтали
	 * @return Alignment - $this
	 */
	function setHorizontal(string $horizontal = null) : self{
		if (!is_null($horizontal)) {
			Validator::validate($horizontal, '$horizontal', self::VALID_HORIZONTAL);
			$this->setIndent(null);

			if ($horizontal !== 'distributed') $this->setWrapText(false);
		}

		$this->model['horizontal'] = $horizontal;
		return $this;
	}

	/**
	 * @return string|null - выравнивание по вертикали
	 */
	function getVertical(){
		return $this->model['vertical'] ?? null;
	}

	/**
	 * @param string|null $vertical - выравнивание по вертикали
	 * @return Alignment - $this
	 */
	function setVertical(string $vertical = null) : self{
		if (!is_null($vertical)) Validator::validate($vertical, '$vertical', self::VALID_VERTICAL);

		$this->model['vertical'] = $vertical;
		return $this;
	}

	/**
	 * @return bool - распределять по ширине
	 */
	function getWrapText() : bool{
		return $this->model['wrapText'] ?? false;
	}

	/**
	 * @param bool $wrapText - распределять по ширине
	 * @return Alignment - $this
	 */
	function setWrapText(bool $wrapText) : self{
		if ($wrapText) $this->setHorizontal('distributed');

		$this->model['wrapText'] = $wrapText;
		return $this;
	}

	/**
	 * @return int|null - отступ слева
	 */
	function getIndent(){
		return $this->model['indent'] ?? null;
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

		$this->model['indent'] = $indent;
		return $this;
	}

	/**
	 * @return string|null - направление чтения
	 */
	function getReadingOrder(){
		return $this->model['readingOrder'] ?? null;
	}

	/**
	 * @param string|null $readingOrder - направление чтения
	 * @return Alignment - $this
	 */
	function setReadingOrder(string $readingOrder = null) : self{
		if (!is_null($readingOrder)) Validator::validate($readingOrder, '$readingOrder', self::VALID_READING_ORDER);

		$this->model['readingOrder'] = $readingOrder;
		return $this;
	}

	/**
	 * @return TextRotation|null - поворот текста
	 */
	function getTextRotation(){
		return $this->textRotation ?? null;
	}

	/**
	 * @param TextRotation|null $textRotation - поворот текста
	 * @return Alignment - $this
	 */
	function setTextRotation(TextRotation $textRotation = null) : self{
		$this->textRotation = $textRotation;
		$this->model['textRotation'] = $textRotation->getModel();
		return $this;
	}

	/**
	 * @return array - модель
	 */
	function getModel(): array{
		return $this->model;
	}
}