<?php

namespace Decaseal\XlsxCreator\Xml\Sheet;

use Decaseal\XlsxCreator\Xml\BaseXml;
use XMLWriter;

class RowXml extends BaseXml {
	function render(XMLWriter $xml, array $model = null){
		if (!$model) return;

		$xml->startElement('row');

	}
}