<?php

namespace Decaseal\XlsxCreator\Xml\Sheet;

use Decaseal\XlsxCreator\Xml\BaseXml;
use Decaseal\XlsxCreator\Xml\Styles\ColorXml;
use XMLWriter;

class SheetPropertiesXml extends BaseXml{

	function render(XMLWriter $xml, array $model = null){
		if (!$model) return;

		$xml->startElement('sheetPr');

		(new ColorXml('tabColor'))->render($xml, $model);

		$xml->endElement();
	}
}