<?php

namespace Topvisor\XlsxCreator\Xml\Styles\Font;

use Topvisor\XlsxCreator\Xml\BaseXml;
use XMLWriter;

class UnderlineXml extends BaseXml {
	public function render(XMLWriter $xml, ?array $model = null) {
		if (!$model || !$model[0]) return;

		$xml->startElement('u');

		switch ($model[0]) {
			case 'double': $xml->writeAttribute('val', 'double');

			break;
			case 'singleAccounting': $xml->writeAttribute('val', 'singleAccounting');

			break;
			case 'doubleAccounting': $xml->writeAttribute('val', 'doubleAccounting');

			break;
		}

		$xml->endElement();
	}
}
