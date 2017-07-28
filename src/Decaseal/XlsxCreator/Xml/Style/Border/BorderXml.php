<?php

namespace Decaseal\XlsxCreator\Xml\Style\Border;

use Decaseal\XlsxCreator\Xml\BaseXml;
use XMLWriter;

class BorderXml extends BaseXml{
	function render(XMLWriter $xml, $model = null){
		if (is_null($model)) return;

		$xml->startElement('border');

		if (isset($model['diagonal'])) {
			if ($model['diagonal']['up'] ?? false) $xml->writeAttribute('diagonalUp', 1);
			if ($model['diagonal']['down'] ?? false) $xml->writeAttribute('diagonalDown', 1);
		}

		$color = $model['color'] ?? null;
		foreach (['left', 'right', 'top', 'bottom', 'diagonal'] as $edge)
			if (isset($model[$edge])) (new EdgeXml($edge, $color))->render($xml, $model[$edge]);

		$xml->endElement();
	}
}