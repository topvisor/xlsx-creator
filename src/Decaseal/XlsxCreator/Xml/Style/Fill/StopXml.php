<?php

namespace Decaseal\XlsxCreator\Xml\Style\Fill;

use Decaseal\XlsxCreator\XlsxCreator;
use Decaseal\XlsxCreator\Xml\BaseXml;
use Decaseal\XlsxCreator\Xml\Style\ColorXml;
use XMLWriter;

class StopXml extends BaseXml{
	const TAG = 'stop';

	function render(XMLWriter $xml, $model = null){
		if(is_null($model)) return;

		$xml->startElement(StopXml::TAG);

		$xml->writeAttribute(XlsxCreator::FILL_STOP_POSITION, $model[XlsxCreator::FILL_STOP_POSITION]);
		(new ColorXml())->render($xml, $model[XlsxCreator::FILL_STOP_COLOR]);

		$xml->endElement();
	}
}