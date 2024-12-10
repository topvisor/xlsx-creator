<?php

namespace Topvisor\XlsxCreator\Xml\Sheet;

use Topvisor\XlsxCreator\Xml\BaseXml;
use XMLWriter;

class MergeXml extends BaseXml {
	public function render(XMLWriter $xml, ?array $model = null) {
		if (!$model || !($model[0] ?? false)) return;

		$xml->startElement('mergeCell');
		$xml->writeAttribute('ref', $model[0]);
		$xml->endElement();
	}
}
