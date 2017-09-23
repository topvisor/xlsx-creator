<?php

namespace Decaseal\XlsxCreator\Xml\Sheet;

use Decaseal\XlsxCreator\Row;
use Decaseal\XlsxCreator\Xml\BaseXml;
use Decaseal\XlsxCreator\Xml\Styles\StylesXml;
use XMLWriter;

class RowXml extends BaseXml {
	function render(XMLWriter $xml, array $model = null){
		if (!$model) return;

		$xml->startElement('row');

		$xml->writeAttribute('r', $model['number']);

		if ($model['height']) {
			$xml->writeAttribute('ht', $model['height']);
			$xml->writeAttribute('customHeight', 1);
		}

		if ($model['hidden']) $xml->writeAttribute('hidden', 1);
		if ($model['min'] && $model['max'] && $model['min'] <= $model['max']) $xml->writeAttribute('spans', "$model[min]:$model[max]");

		if ($model['styleId']) {
			$xml->writeAttribute('s', $model['styleId']);
			$xml->writeAttribute('customFormat', 1);
		}

		$xml->writeAttribute('x14ac:dyDescent', '0.25');

		if ($model['outlineLevel']) $xml->writeAttribute('outlineLevel', $model['outlineLevel']);
		if ($model['collapsed']) $xml->writeAttribute('collapsed', 1);

		$cellXml = new CellXml();
		foreach ($model['cells'] as $cellModel) $cellXml->render($xml, $cellModel);

		$xml->endElement();
	}
}