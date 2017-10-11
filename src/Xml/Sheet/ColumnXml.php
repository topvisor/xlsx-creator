<?php

namespace Topvisor\XlsxCreator\Xml\Sheet;

use Topvisor\XlsxCreator\Xml\BaseXml;
use XMLWriter;

class ColumnXml extends BaseXml {
	function render(XMLWriter $xml, array $model = null){
		if (!$model) return;

		$xml->startElement('col');

		$xml->writeAttribute('min', $model['min']);
		$xml->writeAttribute('max', $model['max']);

		if ($model['width'] ?? false) $xml->writeAttribute('width', $model['width']);
		if ($model['styleId'] ?? false) $xml->writeAttribute('style', $model['styleId']);
		if ($model['hidden'] ?? false) $xml->writeAttribute('hidden', 1);
		if ($model['bestFit'] ?? false) $xml->writeAttribute('bestFit', 1);
		if ($model['outlineLevel'] ?? false) $xml->writeAttribute('outlineLevel', $model['outlineLevel']);
		if ($model['collapsed'] ?? false) $xml->writeAttribute('collapsed', 1);

		$xml->writeAttribute('customWidth', 1);

		$xml->endElement();
	}
}