<?php

namespace Topvisor\XlsxCreator\Structures\Values\RichText;

use Topvisor\XlsxCreator\Exceptions\InvalidValueException;
use Topvisor\XlsxCreator\Structures\Styles\Font;

/**
 * Class RichText. Описывает отрывок текста и его шрифт.
 *
 * @package Topvisor\XlsxCreator\Structures\Values\RichText
 */
class RichText {
	private $text;
	private $font;

	/**
	 * RichText constructor.
	 *
	 * @param string $text - текст
	 * @param Font|null $font - шрифт
	 */
	public function __construct(string $text, ?Font $font = null) {
		$this->setText($text);
		$this->font = $font;
	}

	public function __destruct() {
		unset($this->font);
	}

	/**
	 * @return string - текст
	 */
	public function getText(): string {
		return $this->text;
	}

	/**
	 * @param string $text - текст
	 * @throws InvalidValueException
	 */
	public function setText(string $text): self {
		if (!$text) throw new InvalidValueException('$text must be');
		$this->text = $text;

		return $this;
	}

	/**
	 * @return Font|null - шрифт
	 */
	public function getFont() {
		return $this->font;
	}

	/**
	 * @param Font|null $font - шрифт
	 */
	public function setFont(?Font $font = null): self {
		$this->font = $font;

		return $this;
	}

	/**
	 * @return array - модель
	 */
	public function getModel(): array {
		$model = ['text' => $this->text];
		if ($this->font) $model['font'] = $this->font->getModel();

		return $model;
	}
}
