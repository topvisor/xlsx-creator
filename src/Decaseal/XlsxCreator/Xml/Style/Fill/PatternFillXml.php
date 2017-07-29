<?php

namespace Decaseal\XlsxCreator\Xml\Style\Fill;

use Decaseal\XlsxCreator\XlsxCreator;
use Decaseal\XlsxCreator\Xml\BaseXml;
use Decaseal\XlsxCreator\Xml\Style\ColorXml;
use XMLWriter;

class PatternFillXml extends BaseXml{
	const TAG = 'patternFill';
	const PATTERN_TYPE = 'patternType';

	function render(XMLWriter $xml, $model = null){
		if(is_null($model)) return;

		$xml->startElement(PatternFillXml::TAG);

		$xml->writeAttribute(PatternFillXml::PATTERN_TYPE, $model[XlsxCreator::FILL_PATTERN]);
		if (isset($model[XlsxCreator::FILL_FG_COLOR])) (new ColorXml(XlsxCreator::FILL_FG_COLOR))->render($xml, $model[XlsxCreator::FILL_FG_COLOR]);
		if (isset($model[XlsxCreator::FILL_BG_COLOR])) (new ColorXml(XlsxCreator::FILL_BG_COLOR))->render($xml, $model[XlsxCreator::FILL_BG_COLOR]);

		$xml->endElement();
	}
}