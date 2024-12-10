<?php

namespace Topvisor\XlsxCreator\Structures\Values;

/**
 * Class HyperlinkValue. Используется для задания значения ячейки (гиперссылка).
 *
 * @package  Topvisor\XlsxCreator\Structures\Values
 */
class HyperlinkValue extends Value {
	/**
	 * HyperlinkValue constructor.
	 *
	 * @param string $hyperlink - ссылка
	 * @param string|SharedStringValue|null $text - текст ссылки
	 */
	public function __construct(string $hyperlink, $text = null) {
		if (is_numeric($text)) $text = (string) $text;
		$text ??= $hyperlink;
		$model = ['hyperlink' => $hyperlink];

		if ($text instanceof SharedStringValue) $model['ssId'] = $text->value;
		else $model['text'] = $text;

		parent::__construct($model, parent::TYPE_HYPERLINK);
	}

	/**
	 * @param $value - модель
	 * @return Value - значение ячейки
	 */
	public static function parse($value): Value {
		return new self($value['hyperlink'], $value['text'] ?? new SharedStringValue($value['ssId']));
	}
}
