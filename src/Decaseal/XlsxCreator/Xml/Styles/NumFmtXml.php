<?php

namespace Decaseal\XlsxCreator\Xml\Styles;

use Decaseal\XlsxCreator\Xml\BaseXml;
use XMLWriter;

class NumFmtXml extends BaseXml{
	function render(XMLWriter $xml, $model = null){
		if (!$model) return;

		$xml->startElement('numFmt');

		$xml->writeAttribute('id', $model['id']);
		$xml->writeAttribute('formatCode', $model['formatCode']);

		$xml->endElement();
	}
}