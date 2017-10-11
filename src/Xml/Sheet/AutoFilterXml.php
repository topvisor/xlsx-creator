<?php

namespace Topvisor\XlsxCreator\Xml\Sheet;

use Topvisor\XlsxCreator\Xml\BaseXml;
use XMLWriter;

class AutoFilterXml extends BaseXml{
	function render(XMLWriter $xml, array $model = null){
		if (!$model || !($model[0] ?? false)) return;

		$xml->startElement('autoFilter');
		$xml->writeAttribute('ref', $model[0]);
		$xml->endElement();
	}
}