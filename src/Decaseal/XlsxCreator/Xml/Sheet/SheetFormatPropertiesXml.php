<?php

namespace Decaseal\XlsxCreator\Xml\Sheet;

use Decaseal\XlsxCreator\Xml\BaseXml;
use XMLWriter;

class SheetFormatPropertiesXml extends BaseXml{
	function render(XMLWriter $xml, array $model = null){
		if (!$model) return;

		$xml->startElement('sheetFormatPr');

		if ($model['defaultRowHeight'] ?? false) $xml->writeAttribute('defaultRowHeight', $model['defaultRowHeight']);
		if ($model['outlineLevelRow'] ?? false) $xml->writeAttribute('outlineLevelRow', $model['outlineLevelRow']);
		if ($model['outlineLevelCol'] ?? false) $xml->writeAttribute('outlineLevelCol', $model['outlineLevelCol']);
		if ($model['dyDescent'] ?? false) $xml->writeAttribute('x14ac:dyDescent', $model['dyDescent']);

		$xml->endElement();
	}
}