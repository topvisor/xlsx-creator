<?php

namespace Decaseal\XlsxCreator\Xml\Style\Fill;

use Decaseal\XlsxCreator\XlsxCreator;
use Decaseal\XlsxCreator\Xml\BaseXml;
use XMLWriter;

class FillXml extends BaseXml {
	const TAG = 'fill';

	function render(XMLWriter $xml, $model = null){
		if(is_null($model)) return;
		if($model[XlsxCreator::FILL_TYPE] != XlsxCreator::FILL_PATTERN && $model[XlsxCreator::FILL_TYPE] != XlsxCreator::FILL_GRADIENT) return;

		$xml->startElement(FillXml::TAG);

		switch ($model[XlsxCreator::FILL_TYPE]) {
			case XlsxCreator::FILL_PATTERN: (new PatternFillXml())->render($xml, $model); break;
			case XlsxCreator::FILL_GRADIENT: (new PatternFillXml())->render($xml, $model); break;
		}

		$xml->endElement();
	}
}