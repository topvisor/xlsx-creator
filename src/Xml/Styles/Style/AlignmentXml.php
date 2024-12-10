<?php

namespace Topvisor\XlsxCreator\Xml\Styles\Style;

use Topvisor\XlsxCreator\Xml\BaseXml;
use XMLWriter;

class AlignmentXml extends BaseXml {
	public function render(XMLWriter $xml, ?array $model = null) {
		if (!$model) return;

		$xml->startElement('alignment');

		if ($model['horizontal'] ?? false) $xml->writeAttribute('horizontal', $model['horizontal']);
		if ($model['vertical'] ?? false) $xml->writeAttribute('vertical', $model['vertical']);
		if ($model['wrapText'] ?? false) $xml->writeAttribute('wrapText', 1);
		if ($model['shrinkToFit'] ?? false) $xml->writeAttribute('shrinkToFit', 1);
		if ($model['indent'] ?? false) $xml->writeAttribute('indent', (int) $model['indent']);

		if ($model['textRotation'] ?? false) $xml->writeAttribute(
			'textRotation',
			$model['textRotation'] === 'vertical' ? 255 : (int) $model['textRotation']
		);

		switch ($model['readingOrder'] ?? false) {
			case 'leftToRight': $readingOrder = 1;

			break;
			case 'rightToLeft': $readingOrder = 2;

			break;
			default: $readingOrder = false;

			break;
		}

		if ($readingOrder) $xml->writeAttribute('readingOrder', $readingOrder);

		$xml->endElement();
	}
}
