<?php

namespace Decaseal\XlsxCreator\Xml\Style\Fill;

use Decaseal\XlsxCreator\Xml\BaseXml;
use Decaseal\XlsxCreator\Xml\Style\ColorXml;
use XMLWriter;

class PatternFillXml extends BaseXml{
	function render(XMLWriter $xml, $model = null){
		if(is_null($model)) return;

		$xml->startElement('patternFill');

		$xml->writeAttribute('patternType', $model['pattern']);
		if (isset($model['fgColor'])) (new ColorXml('fgColor'))->render($xml, $model['fgColor']);
		if (isset($model['bgColor'])) (new ColorXml('bgColor'))->render($xml, $model['bgColor']);

		$xml->endElement();
	}
}