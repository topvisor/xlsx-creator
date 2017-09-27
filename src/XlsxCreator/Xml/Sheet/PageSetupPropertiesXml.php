<?php

namespace XlsxCreator\Xml\Sheet;

use XlsxCreator\Xml\BaseXml;
use XMLWriter;

class PageSetupPropertiesXml extends BaseXml{
	function render(XMLWriter $xml, array $model = null){
		if (!$model || !$model['fitToPage']) return;

		$xml->startElement('pageSetUpPr');
		$xml->writeAttribute('fitToPage', 1);
		$xml->endElement();
	}
}