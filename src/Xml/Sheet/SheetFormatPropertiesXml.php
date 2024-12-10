<?php

namespace Topvisor\XlsxCreator\Xml\Sheet;

use Topvisor\XlsxCreator\Xml\BaseXml;
use XMLWriter;

class SheetFormatPropertiesXml extends BaseXml {
	public function render(XMLWriter $xml, ?array $model = null) {
		if (!$model) return;

		$xml->startElement('sheetFormatPr');

		if (isset($model['defaultRowHeight'])) $xml->writeAttribute('defaultRowHeight', $model['defaultRowHeight']);
		if (isset($model['outlineLevelRow'])) $xml->writeAttribute('outlineLevelRow', $model['outlineLevelRow']);
		if (isset($model['outlineLevelCol'])) $xml->writeAttribute('outlineLevelCol', $model['outlineLevelCol']);
		if (isset($model['dyDescent'])) $xml->writeAttribute('x14ac:dyDescent', $model['dyDescent']);

		$xml->endElement();
	}
}
