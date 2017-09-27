<?php

namespace XlsxCreator\Xml\Styles\Border;

use XlsxCreator\Xml\BaseXml;
use XMLWriter;

class BorderXml extends BaseXml{
	function render(XMLWriter $xml, array $model = null){
		$xml->startElement('border');

		$model = $model ?? [];
		$defaultColor = $model['color'] ?? null;

		if (($model['diagonal'] ?? false) && ($model['diagonal']['style'] ?? false)) {
			if ($model['diagonal']['up'] ?? false) $xml->writeAttribute('diagonalUp', 1);
			if ($model['diagonal']['down'] ?? false) $xml->writeAttribute('diagonalDown', 1);
		}

		(new EdgeXml('left', $defaultColor))->render($xml, $model['left'] ?? null);
		(new EdgeXml('right', $defaultColor))->render($xml, $model['right'] ?? null);
		(new EdgeXml('top', $defaultColor))->render($xml, $model['top'] ?? null);
		(new EdgeXml('bottom', $defaultColor))->render($xml, $model['bottom'] ?? null);
		(new EdgeXml('diagonal', $defaultColor))->render($xml, $model['diagonal'] ?? null);

		$xml->endElement();
	}
}