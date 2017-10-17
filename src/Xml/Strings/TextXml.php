<?php

namespace Topvisor\XlsxCreator\Xml\Strings;

use Topvisor\XlsxCreator\Xml\BaseXml;
use XMLWriter;

class TextXml extends BaseXml {
	function render(XMLWriter $xml, array $model = null){
		if (!$model) return;

		$xml->startElement('t');
		$xml->writeAttribute('xml:space', 'preserve');
		$xml->text($model['value']);
		$xml->endElement();
	}
}