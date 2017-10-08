<?php

namespace XlsxCreator\Xml\Sheet;

use XlsxCreator\Xml\BaseXml;
use XlsxCreator\Xml\Styles\ColorXml;
use XMLWriter;

class SheetPropertiesXml extends BaseXml{

	function render(XMLWriter $xml, array $model = null){
		if (!$model || !$model['tabColor'] && (!$model['pageSetup'] || !$model['pageSetup']['fitToPage'])) return;

		$xml->startElement('sheetPr');

		if ($model['tabColor'] ?? false) (new ColorXml('tabColor'))->render($xml, $model['tabColor']);

		(new PageSetupPropertiesXml())->render($xml, $model['pageSetup'] ?? null);

		$xml->endElement();
	}
}