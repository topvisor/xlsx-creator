<?php

namespace Decaseal\XlsxCreator\Xml\Sheet;

use Decaseal\XlsxCreator\Xml\BaseXml;
use XMLWriter;

class SheetFormatPropertiesXml extends BaseXml{
	function render(XMLWriter $xml, array $model = null){
		if (!$model) return;

		$xml->startElement('sheetFormatPr');

		if (isset($model['defaultRowHeight'])) $xml->writeAttribute('defaultRowHeight', $model['defaultRowHeight']);
		if (isset($model['outlineLevelRow'])) $xml->writeAttribute('outlineLevelRow', $model['outlineLevelRow']);
		if (isset($model['outlineLevelCol'])) $xml->writeAttribute('outlineLevelCol', $model['outlineLevelCol']);
		if (isset($model['dyDescent'])) $xml->writeAttribute('x14ac:dyDescent', $model['dyDescent']);

		$xml->endElement();
	}
}