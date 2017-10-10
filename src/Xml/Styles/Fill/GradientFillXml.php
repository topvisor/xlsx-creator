<?php

namespace Topvisor\XlsxCreator\Xml\Styles\Fill;

use Topvisor\XlsxCreator\Xml\BaseXml;
use XMLWriter;

class GradientFillXml extends BaseXml{
	function render(XMLWriter $xml, array $model = null){
		if (!$model) return;

		$xml->startElement('gradientFill');

		switch ($model['gradient']) {
			case 'angle':
				$xml->writeAttribute('degree', $model['degree']);
				break;

			case 'path':
				$xml->writeAttribute('type', 'path');

				if ($model['left'] ?? false) {
					$xml->writeAttribute('left', $model['left']);
					if (!($model['right'] ?? false)) $xml->writeAttribute('right', $model['left']);
				}

				if ($model['right'] ?? false) $xml->writeAttribute('right', $model['right']);

				if ($model['top'] ?? false) {
					$xml->writeAttribute('top', $model['top']);
					if (!($model['bottom'] ?? false)) $xml->writeAttribute('bottom', $model['top']);
				}

				if ($model['bottom'] ?? false) $xml->writeAttribute('bottom', $model['bottom']);

				break;
		}

		if ($model['stops'] ?? false) {
			$stopXml = new StopXml();
			foreach ($model['stops'] as $stopModel) $stopXml->render($xml, $stopModel);
		}

		$xml->endElement();
	}
}