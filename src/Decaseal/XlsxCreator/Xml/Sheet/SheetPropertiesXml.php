<?php

namespace Decaseal\XlsxCreator\Xml\Sheet;

use Decaseal\XlsxCreator\Xml\BaseXml;
use Decaseal\XlsxCreator\Xml\Styles\ColorXml;
use XMLWriter;

class SheetPropertiesXml extends BaseXml{

	function render(XMLWriter $xml, $model = null){
		if (!$model) return;

		$xml->startElement('sheetPr');

		(new ColorXml('tabColor'))->render($model);

		$xml->endElement();
	}
}