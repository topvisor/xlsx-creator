<?php

namespace Decaseal\XlsxCreator\Xml\Core\App;

use Decaseal\XlsxCreator\Xml\BaseXml;
use XMLWriter;

class AppTitlesOfPartsXml extends BaseXml{
	function render(XMLWriter $xml, array $model = null){
		if (is_null($model)) return;

		$xml->startElement('TitlesOfParts');
		$xml->startElement('vt:vector');

		$xml->writeAttribute('size', (string) count($model));
		$xml->writeAttribute('baseType', 'lpstr');

		foreach ($model as $worksheet) $xml->writeElement('vt:lpstr', $worksheet['name']);

		$xml->endElement();
		$xml->endElement();
	}
}