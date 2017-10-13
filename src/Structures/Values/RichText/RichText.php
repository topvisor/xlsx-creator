<?php

namespace Topvisor\XlsxCreator\Structures\Values\RichText;

use Topvisor\XlsxCreator\Exceptions\InvalidValueException;
use Topvisor\XlsxCreator\Structures\Styles\Font;

/**
 * Class RichText. Описывает отрывок текста и его шрифт.
 *
 * @package Topvisor\XlsxCreator\Structures\Values\RichText
 */
class RichText{
	private $text;
	private $font;

	/**
	 * RichText constructor.
	 *
	 * @param string $text - текст
	 * @param Font|null $font - шрифт
	 */
	function __construct(string $text, Font $font = null){
		$this->setText($text);
		$this->font = $font;
	}

	/**
	 * @return string - текст
	 */
	function getText(): string{
		return $this->text;
	}

	/**
	 * @param string $text - текст
	 * @return RichText
	 * @throws InvalidValueException
	 */
	function setText(string $text) : self{
		if (!$text) throw new InvalidValueException('$text must be');
		$this->text = $text;
		return $this;
	}

	/**
	 * @return Font|null - шрифт
	 */
	function getFont(){
		return $this->font;
	}

	/**
	 * @param Font|null $font - шрифт
	 * @return RichText
	 */
	function setFont(Font $font = null) : self{
		$this->font = $font;
		return $this;
	}

	/**
	 * @return array - модель
	 */
	function getModel() : array{
		$model = ['text' => $this->text];
		if ($this->font) $model['font'] = $this->font->getModel();

		return $model;
	}
}