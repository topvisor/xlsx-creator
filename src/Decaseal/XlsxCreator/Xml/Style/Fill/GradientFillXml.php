<?php

namespace Decaseal\XlsxCreator\Xml\Style\Fill;

use Decaseal\XlsxCreator\XlsxCreator;
use Decaseal\XlsxCreator\Xml\BaseXml;
use XMLWriter;

class GradientFillXml extends BaseXml{
	const TAG = 'gradientFill';
	const TYPE = 'type';

	function render(XMLWriter $xml, $model = null){
		if(is_null($model)) return;

		$xml->startElement(GradientFillXml::TAG);

		switch ($model[XlsxCreator::FILL_GRADIENT]) {
			case XlsxCreator::FILL_GRADIENT_ANGLE:
				$xml->writeAttribute(XlsxCreator::FILL_DEGREE, $model[XlsxCreator::FILL_DEGREE]);
				break;

			case XlsxCreator::FILL_GRADIENT_PATH:
				$xml->writeAttribute(GradientFillXml::TYPE, XlsxCreator::FILL_GRADIENT_PATH);

				$center = $model[XlsxCreator::FILL_CENTER] ?? [];

				if ($center[XlsxCreator::FILL_CENTER_LEFT]) {
					$xml->writeAttribute(XlsxCreator::FILL_CENTER_LEFT, $center[XlsxCreator::FILL_CENTER_LEFT]);
					if(!$center[XlsxCreator::FILL_CENTER_RIGHT]) $xml->writeAttribute(XlsxCreator::FILL_CENTER_RIGHT, $center[XlsxCreator::FILL_CENTER_LEFT]);
				}

				if ($center[XlsxCreator::FILL_CENTER_TOP]) {
					$xml->writeAttribute(XlsxCreator::FILL_CENTER_TOP, $center[XlsxCreator::FILL_CENTER_TOP]);
					if(!$center[XlsxCreator::FILL_CENTER_BOTTOM]) $xml->writeAttribute(XlsxCreator::FILL_CENTER_BOTTOM, $center[XlsxCreator::FILL_CENTER_TOP]);
				}

				if($center[XlsxCreator::FILL_CENTER_RIGHT]) $xml->writeAttribute(XlsxCreator::FILL_CENTER_RIGHT, $center[XlsxCreator::FILL_CENTER_RIGHT]);

				if($center[XlsxCreator::FILL_CENTER_BOTTOM]) $xml->writeAttribute(XlsxCreator::FILL_CENTER_BOTTOM, $center[XlsxCreator::FILL_CENTER_BOTTOM]);

				break;
		}

		$stopXml = new StopXml();
		foreach ($model[XlsxCreator::FILL_STOPS] as $stopModel) $stopXml->render($xml, $stopModel);

		$xml->endElement();
	}
}