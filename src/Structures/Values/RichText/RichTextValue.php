<?php

namespace Topvisor\XlsxCreator\Structures\Values\RichText;

use Topvisor\XlsxCreator\Structures\Values\Value;

/**
 * Class RichTextValue. Строка с разными шрифтами
 *
 * @package Topvisor\XlsxCreator\Structures\Values
 */
class RichTextValue extends Value{
	/**
	 * RichTextValue constructor.
	 *
	 * @param array $richTexts - массив RichText
	 */
	function __construct(array $richTexts){
		parent::__construct(array_map(function(RichText $richText){
			return $richText->getModel();
		}, $richTexts), Value::TYPE_RICH_TEXT);
	}

	/**
	 * @param $value - модель
	 * @return Value - значение ячейки
	 */
	static function parse($value): Value{
		return new RichTextValue(array_map(function(array $richTextModel){
			return new RichText($richTextModel['text'], $richTextModel['font'] ?? null);
		}, $value));
	}
}