<?php

namespace Topvisor\XlsxCreator\Xml\Strings;

use Topvisor\XlsxCreator\Xml\BaseXml;
use XMLWriter;

class TextXml extends BaseXml {
	public function render(XMLWriter $xml, ?array $model = null) {
		if (!$model) return;

		$xml->startElement('t');
		$xml->writeAttribute('xml:space', 'preserve');
		$xml->text($this->prepareText($model['value']));
		$xml->endElement();
	}
}
