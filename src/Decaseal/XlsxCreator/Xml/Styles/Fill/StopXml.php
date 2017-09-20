<?php

namespace Decaseal\XlsxCreator\Xml\Styles\Fill;

use Decaseal\XlsxCreator\Xml\BaseXml;
use Decaseal\XlsxCreator\Xml\Styles\ColorXml;
use XMLWriter;

class StopXml extends BaseXml{
	function render(XMLWriter $xml, array $model = null){
		if (!$model) return;

		$xml->startElement('stop');

		$xml->writeAttribute('position', $model['position']);
		(new ColorXml())->render($xml, $model['color']);

		$xml->endElement();
	}
}