<?php

namespace Topvisor\XlsxCreator\Xml\Sheet;

use Topvisor\XlsxCreator\Structures\Values\Value;
use Topvisor\XlsxCreator\Xml\BaseXml;
use XMLWriter;

class CellXml extends BaseXml{
	function render(XMLWriter $xml, array $model = null){
		if (!$model || $model['type'] === Value::TYPE_NULL && !$model['styleId']) return;

		$xml->startElement('c');

		$xml->writeAttribute('r', $model['address']);
		if($model['styleId']) $xml->writeAttribute('s', $model['styleId']);

		switch ($model['type']) {
			case Value::TYPE_NUMBER:
				$xml->writeElement('v', $model['value']);
				break;

			case Value::TYPE_BOOL:
				$xml->writeAttribute('t', 'b');
				$xml->writeElement('v', $model['value'] ? '1' : '0');
				break;

			case Value::TYPE_ERROR:
				$xml->writeAttribute('t', 'e');
				$xml->writeElement('v', $model['value']);
				break;

			case Value::TYPE_STRING:
				$xml->writeAttribute('t', 'str');
				$xml->writeAttribute('xml:space', 'preserve');
				$xml->writeElement('v', preg_replace('/_x\d{4}_/', '_x005F$0', $model['value']));
				break;

			case Value::TYPE_DATE:
				$xml->writeElement('v', 25569 + ($model['value']->getTimestamp() / (24 * 3600)));
				break;

			case Value::TYPE_HYPERLINK:
				if (isset($model['value']['ssId'])) {
					$xml->writeAttribute('t', 's');
					$xml->writeElement('v', $model['value']['ssId']);
				} else {
					$xml->writeAttribute('t', 'str');
					$xml->writeElement('v', $model['value']['text']);
				}
				break;

			case Value::TYPE_FORMULA:
				$xml->writeElement('f', $model['value']);
				break;

			case Value::TYPE_SHARED_STRING:
				$xml->writeAttribute('t', 's');
				$xml->writeElement('v', $model['value']);
				break;
		}

		$xml->endElement();
	}
}