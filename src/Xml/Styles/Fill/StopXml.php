<?php

namespace Topvisor\XlsxCreator\Xml\Styles\Fill;

use Topvisor\XlsxCreator\Xml\BaseXml;
use Topvisor\XlsxCreator\Xml\Styles\ColorXml;
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