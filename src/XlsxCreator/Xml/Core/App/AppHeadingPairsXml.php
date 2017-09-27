<?php

namespace XlsxCreator\Xml\Core\App;

use XlsxCreator\Xml\BaseXml;
use XMLWriter;

class AppHeadingPairsXml extends BaseXml{
	function render(XMLWriter $xml, array $model = null){
		if (is_null($model)) return;

		$xml->startElement('HeadingPairs');
		$xml->startElement('vt:vector');

		$xml->writeAttribute('size', '2');
		$xml->writeAttribute('baseType', 'variant');

		$xml->startElement('vt:variant');
		$xml->writeElement('vt:lpstr', 'Worksheets');
		$xml->endElement();

		$xml->startElement('vt:variant');
		$xml->writeElement('vt:i4', (string) count($model));
		$xml->endElement();

		$xml->endElement();
		$xml->endElement();
	}
}