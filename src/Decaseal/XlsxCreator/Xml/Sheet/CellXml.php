<?php

namespace Decaseal\XlsxCreator\Xml\Sheet;

use Decaseal\XlsxCreator\Cell;
use Decaseal\XlsxCreator\Xml\BaseXml;
use XMLWriter;

class CellXml extends BaseXml{
	function render(XMLWriter $xml, array $model = null){
		if (!$model || $model['type'] === Cell::TYPE_NULL && !$model['styleId']) return;

		$xml->startElement('c');

		$xml->writeAttribute('r', $model['address']);
		if($model['styleId']) $xml->writeAttribute('s', $model['styleId']);

		switch ($model['type']) {
			case Cell::TYPE_NUMBER:
				$xml->writeElement('v', $model['value']);
				break;

			case Cell::TYPE_BOOL:
				$xml->writeAttribute('t', 'b');
				$xml->writeElement('v', $model['value'] ? '1' : '0');
				break;

			case Cell::TYPE_ERROR:
				$xml->writeAttribute('t', 'e');
				$xml->writeElement('v', $model['value']);
				break;

			case Cell::TYPE_STRING:
				$xml->writeAttribute('t', 'str');
				$xml->writeElement('v', $model['value']);
				break;

			case Cell::TYPE_DATE:
				$xml->writeElement('v', 25569 + ($model['value']->getTimestamp() / (24 * 3600)));
				break;

			case Cell::TYPE_HYPERLINK:
				$xml->writeAttribute('t', 'str');
				$xml->writeElement('v', $model['text']);
				break;
		}

		$xml->endElement();
	}
}