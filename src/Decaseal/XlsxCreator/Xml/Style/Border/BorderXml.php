<?php

namespace Decaseal\XlsxCreator\Xml\Style\Border;

use Decaseal\XlsxCreator\XlsxCreator;
use Decaseal\XlsxCreator\Xml\BaseXml;
use XMLWriter;

class BorderXml extends BaseXml{
	const TAG = 'border';
	const DIAGONAL_UP = 'diagonalUp';
	const DIAGONAL_DOWN = 'diagonalDown';

	function render(XMLWriter $xml, $model = null){
		if (is_null($model)) return;

		$xml->startElement(BorderXml::TAG);

		if (isset($model[XlsxCreator::BORDER_DIAGONAL])) {
			if ($model[XlsxCreator::BORDER_DIAGONAL][XlsxCreator::BORDER_DIAGONAL_UP] ?? false)
				$xml->writeAttribute(XlsxCreator::BORDER_DIAGONAL_UP, 1);

			if ($model[XlsxCreator::BORDER_DIAGONAL][XlsxCreator::BORDER_DIAGONAL_DOWN] ?? false)
				$xml->writeAttribute(XlsxCreator::BORDER_DIAGONAL_DOWN, 1);
		}

		$color = $model[XlsxCreator::BORDER_COLOR] ?? null;
		$edges = [
			XlsxCreator::BORDER_LEFT,
			XlsxCreator::BORDER_RIGHT,
			XlsxCreator::BORDER_TOP,
			XlsxCreator::BORDER_BOTTOM,
			XlsxCreator::BORDER_DIAGONAL
		];

		foreach ($edges as $edge) (new EdgeXml($edge, $color))->render($xml, $model[$edge] ?? null);

		$xml->endElement();
	}
}