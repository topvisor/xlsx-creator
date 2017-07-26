<?php

namespace Decaseal\XlsxCreator\Xml\Style;

use Decaseal\XlsxCreator\Xml\BaseXml;
use XMLWriter;

class NumFmtXml extends BaseXml{
	private const TAG = 'numFmt';

	function render(XMLWriter $xml, $model = null){
		if(is_null($model)) return;

		$xml->startElement(NumFmtXml::TAG);
		$xml->writeAttribute('numFmtId', $model['id']);
		$xml->writeAttribute('formatCode', $model['formatCode']);
		$xml->endElement();
	}
}