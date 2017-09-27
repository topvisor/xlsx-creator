<?php

namespace XlsxCreator\Xml\Sheet;


use XlsxCreator\Xml\BaseXml;
use XMLWriter;

class HyperlinkXml extends BaseXml{
	function render(XMLWriter $xml, array $model = null){
		if (!$model) return;

		$xml->startElement('hyperlink');

		$xml->writeAttribute('ref', $model['address']);
		$xml->writeAttribute('r:id', $model['rId']);

		$xml->endElement();
	}
}