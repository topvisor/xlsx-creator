<?php

namespace XlsxCreator\Xml\Styles\Fill;

use XlsxCreator\Xml\BaseXml;
use XlsxCreator\Xml\Styles\ColorXml;
use XMLWriter;

class PatternFillXml extends BaseXml{
	function render(XMLWriter $xml, array $model = null){
		if (!$model) return;

		$xml->startElement('patternFill');

		$xml->writeAttribute('patternType', $model['pattern']);

		if ($model['fgColor'] ?? false) (new ColorXml('fgColor'))->render($xml, $model['fgColor']);
		if ($model['bgColor'] ?? false) (new ColorXml('bgColor'))->render($xml, $model['bgColor']);

		$xml->endElement();
	}
}