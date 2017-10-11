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

		if ($model['width']) $xml->writeAttribute('width', $model['width']);
		if ($model['styleId']) $xml->writeAttribute('style', $model['styleId']);
		if ($model['hidden']) $xml->writeAttribute('hidden', 1);
		if ($model['bestFit']) $xml->writeAttribute('bestFit', 1);
		if ($model['outlineLevel']) $xml->writeAttribute('outlineLevel', $model['outlineLevel']);
		if ($model['collapsed']) $xml->writeAttribute('collapsed', 1);

		$xml->writeAttribute('customWidth', 1);

		$xml->endElement();
	}
}