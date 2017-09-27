<?php

namespace XlsxCreator\Xml\Styles;

use XlsxCreator\Xml\BaseXml;
use XMLWriter;

class NumFmtXml extends BaseXml{
	function render(XMLWriter $xml, array $model = null){
		if (!$model) return;

		$xml->startElement('numFmt');

		$xml->writeAttribute('id', $model['id']);
		$xml->writeAttribute('formatCode', $model['formatCode']);

		$xml->endElement();
	}
}