<?php

namespace Decaseal\XlsxCreator\Xml\Sheet;

use Decaseal\XlsxCreator\Xml\BaseXml;
use Decaseal\XlsxCreator\Xml\Styles\ColorXml;
use XMLWriter;

class SheetPropertiesXml extends BaseXml{

	function render(XMLWriter $xml, array $model = null){
		if (!$model || !$model['tabColor'] && (!$model['pageSetup'] || !$model['pageSetup']['fitToPage'])) return;

		$xml->startElement('sheetPr');

		(new ColorXml('tabColor'))->render($xml, $model['tabColor']);
		(new PageSetupPropertiesXml())->render($xml, $model['pageSetup']);

		$xml->endElement();
	}
}