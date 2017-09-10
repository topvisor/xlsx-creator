<?php

namespace Decaseal\XlsxCreator\Xml\Book\Workbook;

use Decaseal\XlsxCreator\Xml\BaseXml;
use XMLWriter;

class SheetXml extends BaseXml{
	function render(XMLWriter $xml, $model = null){
		if (!$model) return;

		$xml->startElement('sheet');

		$xml->writeAttribute('sheetId', $model->getId());
		$xml->writeAttribute('name', $model->getName());
		$xml->writeAttribute('r:id', $model->getRId());

		$xml->endElement();
	}
}