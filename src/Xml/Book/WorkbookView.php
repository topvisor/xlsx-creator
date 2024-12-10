<?php

namespace Topvisor\XlsxCreator\Xml\Book;

use Topvisor\XlsxCreator\Xml\BaseXml;
use XMLWriter;

class WorkbookView extends BaseXml {
	public function render(XMLWriter $xml, ?array $model = null) {
		if (!$model) return;

		$xml->startElement('workbookView');

		$xml->writeAttribute('xWindow', $model['xWindow'] ?? 0);
		$xml->writeAttribute('yWindow', $model['yWindow'] ?? 0);
		$xml->writeAttribute('windowWidth', $model['windowWidth'] ?? 12000);
		$xml->writeAttribute('windowHeight', $model['windowHeight'] ?? 24000);

		if ($model['firstSheet'] ?? false) $xml->writeAttribute('firstSheet', $model['firstSheet']);
		if ($model['activeTab'] ?? false) $xml->writeAttribute('activeTab', $model['activeTab']);

		if ($model['visibility'] ?? false && $model['visibility'] !== 'visible') $xml->writeAttribute('visibility', $model['visibility']);

		$xml->endElement();
	}
}
