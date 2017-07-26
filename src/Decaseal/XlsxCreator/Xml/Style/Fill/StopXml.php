<?php

namespace Decaseal\XlsxCreator\Xml\Style\Fill;

use Decaseal\XlsxCreator\Xml\BaseXml;
use Decaseal\XlsxCreator\Xml\Style\ColorXml;
use XMLWriter;

class StopXml extends BaseXml{
	function render(XMLWriter $xml, $model = null){
		if(is_null($model)) return;

		$xml->startElement('stop');

		$xml->writeAttribute('position', $model['position']);
		(new ColorXml())->render($xml, $model['color']);

		$xml->endElement();
	}
}