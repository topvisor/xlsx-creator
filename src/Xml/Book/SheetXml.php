<?php

namespace Topvisor\XlsxCreator\Xml\Book;

use Topvisor\XlsxCreator\Xml\BaseXml;
use XMLWriter;

class SheetXml extends BaseXml {
	public function render(XMLWriter $xml, ?array $model = null) {
		if (!$model) return;

		$xml->startElement('sheet');

		$xml->writeAttribute('sheetId', $model['id']);
		$xml->writeAttribute('name', $model['name']);
		$xml->writeAttribute('r:id', $model['rId']);

		$xml->endElement();
	}
}
