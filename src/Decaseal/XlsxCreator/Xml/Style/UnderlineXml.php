<?php

namespace Decaseal\XlsxCreator\Xml\Style;

use Decaseal\XlsxCreator\XlsxCreator;
use Decaseal\XlsxCreator\Xml\BaseXml;
use TypeError;
use XMLWriter;

class UnderlineXml extends BaseXml{
	const TAG = 'u';
	const VAL = 'val';

	function render(XMLWriter $xml, $model = null){
		if (!$model) return;

		$xml->startElement(UnderlineXml::TAG);

		switch ($model) {
			case XlsxCreator::FONT_UNDERLINE_DOUBLE: $attributes = [UnderlineXml::VAL => XlsxCreator::FONT_UNDERLINE_DOUBLE]; break;
			case XlsxCreator::FONT_UNDERLINE_SINGLE_ACCOUNTING: $attributes = [UnderlineXml::VAL => XlsxCreator::FONT_UNDERLINE_SINGLE_ACCOUNTING]; break;
			case XlsxCreator::FONT_UNDERLINE_DOUBLE_ACCOUNTING: $attributes = [UnderlineXml::VAL => XlsxCreator::FONT_UNDERLINE_DOUBLE_ACCOUNTING]; break;
			default: $attributes = []; break;
		}
		foreach ($attributes as $name => $value) $xml->writeAttribute($name, $value);

		$xml->endElement();
	}
}