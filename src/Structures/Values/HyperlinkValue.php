<?php

namespace Topvisor\XlsxCreator\Structures\Values;

/**
 * Class HyperlinkValue. Используется для задания значения ячейки (гиперссылка).
 *
 * @package  Topvisor\XlsxCreator\Structures\Values
 */
class HyperlinkValue extends Value{
	/**
	 * HyperlinkValue constructor.
	 *
	 * @param string $hyperlink - ссылка
	 * @param string|null $text - текст ссылки
	 */
	function __construct(string $hyperlink, string $text = null){
		parent::__construct(['text' => $text ?? $hyperlink, 'hyperlink' => $hyperlink], parent::TYPE_HYPERLINK);
	}

	/**
	 * @param $value - модель
	 * @return Value - значение ячейки
	 */
	static function parse($value): Value{
		return new self($value['hyperlink'], $value['text']);
	}
}