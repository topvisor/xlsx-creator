<?php

namespace Topvisor\XlsxCreator\Xml\Strings;

use Topvisor\XlsxCreator\Xml\BaseXml;
use XMLWriter;

class TextXml extends BaseXml {
	function render(XMLWriter $xml, array $model = null){
		if (!$model) return;

		$xml->startElement('t');
		$xml->writeAttribute('xml:space', 'preserve');
		$xml->writeElement('v', preg_replace('/_x\d{4}_/', '_x005F$0', $model['value']));
		$xml->endElement();
	}
}