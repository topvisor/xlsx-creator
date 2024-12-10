<?php

namespace Topvisor\XlsxCreator\Xml\Sheet;

use Topvisor\XlsxCreator\Xml\BaseXml;
use XMLWriter;

class PageSetupPropertiesXml extends BaseXml {
	public function render(XMLWriter $xml, ?array $model = null) {
		if (!$model || !$model['fitToPage']) return;

		$xml->startElement('pageSetUpPr');
		$xml->writeAttribute('fitToPage', 1);
		$xml->endElement();
	}
}
